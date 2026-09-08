import os
import sys
import time
import threading
from threading import Lock, Event  # CORREÇÃO BUG #1: Sincronização de threads para race condition
import subprocess
import tempfile
import pandas as pd
import json
import re
import uuid
import html as html_lib
import logging  # CORREÇÃO CLEAN CODE: Logging estruturado em vez de print()
from typing import Optional, Dict, List  # CORREÇÃO CLEAN CODE: Type hints

# Configurar logging estruturado
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('r9bot.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# CONSTANTES DE CONFIGURAÇÃO (CORREÇÃO CLEAN CODE: Magic numbers/strings centralizados)
MAX_TEMPLATE_SIZE_MB = 10  # Máximo 10MB para arquivo de template
DEFAULT_SEND_INTERVAL = 2  # Intervalo padrão entre emails (segundos)
EMAIL_REGEX = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'  # Regex robusta para validação de email
RGB_PATTERN = r'rgba?\s*\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)'  # CORREÇÃO BUG #5: Regex mais robusta para RGB

# Importação da biblioteca COM para comunicação segura com Outlook em threads separadas
try:
    import pythoncom
except ImportError:
    pass

from bs4 import BeautifulSoup

# IMPORTANTE: os imports do WebEngine precisam vir ANTES de qualquer QApplication
# ser criado, senão o Qt lança erro em tempo de execução.
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebChannel import QWebChannel

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QFileDialog,
    QMessageBox, QAbstractItemView, QProgressBar, QHeaderView, QFrame, QInputDialog, QMenu
)
from PyQt6.QtCore import Qt, pyqtSignal, pyqtSlot, QObject, QUrl, QTimer
from PyQt6.QtGui import QKeySequence, QShortcut, QFont, QColor

try:
    # Usado para "inlinar" o CSS no HTML final (necessário para o Outlook
    # respeitar a formatação, já que ele ignora <style> no <head> na maioria
    # dos casos). Instale com: pip install premailer
    from premailer import transform as inline_css
    PREMAILER_DISPONIVEL = True
except ImportError:
    PREMAILER_DISPONIVEL = False

    def inline_css(html, base_url=None):
        return html


def get_resource_path(relative_path):
    """ Retorna o caminho absoluto, compatível com o PyInstaller (_MEIPASS) e desenvolvimento local """
    try:
        # Quando empacotado pelo PyInstaller, os assets ficam na pasta temporária _MEIPASS
        base_path = sys._MEIPASS
    except AttributeError:
        # Quando rodando diretamente via Python, usa o diretório do script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def montar_variaveis_substituicao(nome: str, contato: Dict, escapar: bool = False) -> Dict[str, str]:
    """Monta o dicionário de substituição ({{nome}}, {{var1}}, {{var2}}, {{varN}}...) a
    partir de um contato — funciona pra QUALQUER quantidade de variáveis (Var1..VarN),
    não fica preso a exatamente três. Se escapar=True, aplica html.escape em cada valor
    (uso no corpo HTML); senão mantém texto puro (uso no Assunto, que não é HTML)."""
    def val(v):
        v = v or ""
        return html_lib.escape(v) if escapar else v

    resultado = {}
    nome_v = val(nome)
    resultado["{{nome}}"] = nome_v
    resultado["{{Nome}}"] = nome_v
    resultado["{{NOME}}"] = nome_v

    for chave, valor in contato.items():
        m = re.match(r'^Var(\d+)$', chave)
        if not m:
            continue
        i = m.group(1)
        v = val(valor)
        resultado[f"{{{{var{i}}}}}"] = v
        resultado[f"{{{{Var{i}}}}}"] = v
        resultado[f"{{{{VAR{i}}}}}"] = v
    return resultado


def substituir_variaveis(texto: str, mapa: Dict[str, str]) -> str:
    """Troca cada {{...}} em `texto` pelo valor correspondente em `mapa`. Tolerante a
    espaços dentro das chaves (ex: "{{ nome }}" funciona igual a "{{nome}}") — muita
    gente digita assim naturalmente (é como funciona em outros sistemas de template tipo
    Jinja2/Handlebars), e sem essa tolerância a substituição falhava silenciosamente,
    deixando o "{{ nome }}" literal no e-mail/assunto em vez do valor de verdade."""
    def replacer(match):
        bruto = match.group(0)
        normalizado = '{{' + re.sub(r'\s+', '', bruto[2:-2]) + '}}'
        return mapa.get(normalizado, bruto)
    return re.sub(r"\{\{[^}]+\}\}", replacer, texto)


def formatar_duracao(segundos: int) -> str:
    """Formata uma duração em segundos como '2min 30s' ou '45s', pra mostrar o tempo
    restante estimado durante uma campanha."""
    segundos = max(0, int(segundos))
    minutos, s = divmod(segundos, 60)
    if minutos > 0:
        return f"{minutos}min {s}s"
    return f"{s}s"


def ler_arquivo_template(caminho: str) -> Optional[str]:
    """
    OTIMIZAÇÃO: Lê arquivo template com fallback de encoding.
    Elimina duplicação de código (estava em 2 lugares).
    
    Args:
        caminho: Caminho do arquivo template HTML
        
    Returns:
        Conteúdo do arquivo ou None se falhar
    """
    try:
        # Tentar UTF-8 primeiro (encoding mais moderno)
        with open(caminho, "r", encoding="utf-8") as f:
            conteudo = f.read()
            # VALIDAÇÃO BUG #7: Verificar tamanho antes de processar
            if len(conteudo) > MAX_TEMPLATE_SIZE_MB * 1024 * 1024:
                logger.error(f"Template excede tamanho máximo de {MAX_TEMPLATE_SIZE_MB}MB")
                return None
            return conteudo
    except UnicodeDecodeError:
        # Fallback para Latin-1 se UTF-8 falhar
        try:
            with open(caminho, "r", encoding="latin-1") as f:
                conteudo = f.read()
                if len(conteudo) > MAX_TEMPLATE_SIZE_MB * 1024 * 1024:
                    logger.error(f"Template excede tamanho máximo de {MAX_TEMPLATE_SIZE_MB}MB")
                    return None
                return conteudo
        except Exception as e:
            logger.error(f"Erro ao ler template {caminho}: {type(e).__name__}: {e}")
            return None


def preparar_imagens_inline(html: str):
    """Procura <img src="data:image/...;base64,..."> no HTML (imagens embutidas pelo
    editor via upload local) e troca cada uma por um `cid:` único, devolvendo também a
    lista de arquivos temporários a anexar em cada e-mail.

    Por quê: Outlook (motor do Word) não é confiável renderizando data URIs direto no
    HTMLBody. O jeito nativo e robusto de embutir imagem num e-mail do Outlook via COM é
    anexá-la como anexo "inline" com um Content-ID e referenciá-la como `cid:` no HTML —
    é assim que o próprio Outlook faz quando você cola uma imagem num e-mail.

    Returns:
        (html_com_cid, lista_de_anexos) onde lista_de_anexos é uma lista de dicts
        {"caminho": <arquivo temporário>, "content_id": <cid usado no html>}.
        Os arquivos temporários devem ser apagados pelo chamador após o uso.
    """
    import base64 as b64lib

    anexos = []
    padrao = re.compile(r'src="data:image/([a-zA-Z0-9.+-]+);base64,([^"]+)"')

    def substituir(match):
        mime_sub, dados_b64 = match.group(1), match.group(2)
        ext = "jpg" if mime_sub == "jpeg" else re.sub(r'[^a-z0-9]', '', mime_sub.lower()) or "png"
        try:
            dados_binarios = b64lib.b64decode(dados_b64)
        except Exception as e:
            logger.warning(f"Falha ao decodificar imagem embutida: {e}")
            return match.group(0)  # mantém como estava se não conseguir decodificar

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}")
        tmp.write(dados_binarios)
        tmp.close()

        content_id = f"img{len(anexos)}_{uuid.uuid4().hex[:8]}@r9bot"
        anexos.append({"caminho": tmp.name, "content_id": content_id})
        return f'src="cid:{content_id}"'

    html_convertido = padrao.sub(substituir, html)
    return html_convertido, anexos


def anexar_imagens_inline(mail, anexos) -> None:
    """Anexa ao MailItem do Outlook cada imagem preparada por preparar_imagens_inline,
    marcando-a com o Content-ID correto (propriedade MAPI PR_ATTACH_CONTENT_ID) pra
    que o `cid:` no HTMLBody seja resolvido corretamente pelo Outlook."""
    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    for item in anexos:
        try:
            anexo = mail.Attachments.Add(item["caminho"])
            anexo.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, item["content_id"])
        except Exception as e:
            logger.warning(f"Falha ao anexar imagem inline {item['caminho']}: {e}")


def limpar_arquivos_temporarios(anexos) -> None:
    """Remove os arquivos temporários criados por preparar_imagens_inline depois que a
    campanha (ou o teste) termina de usá-los."""
    for item in anexos:
        try:
            os.unlink(item["caminho"])
        except Exception:
            pass


def carregar_stylesheet_global() -> str:
    """
    DESIGN SYSTEM: Carrega o QSS global corporativo.
    Centraliza toda a identidade visual em um único arquivo.
    
    Returns:
        String com o QSS completo ou vazio se arquivo não encontrado
    """
    qss_path = get_resource_path("app_style.qss")
    try:
        with open(qss_path, "r", encoding="utf-8") as f:
            qss = f.read()
            logger.info(f"QSS global carregado com sucesso: {qss_path}")
            return qss
    except FileNotFoundError:
        logger.warning(f"Arquivo QSS não encontrado: {qss_path}")
        logger.warning("Aplicação continuará com estilos padrão. Certifique-se que app_style.qss está no mesmo diretório de app_qt.py")
        return ""
    except Exception as e:
        logger.error(f"Erro ao carregar QSS: {type(e).__name__}: {e}")
        return ""


class WorkerSignals(QObject):
    """ Define os sinais utilizados pela Thread de Envio para atualizar a UI de forma segura (Thread-Safe) """
    status_updated = pyqtSignal(str, str, int, int)
    progress_updated = pyqtSignal(int, int)
    finished = pyqtSignal(int, int)
    row_updated = pyqtSignal(int, str)  # Para atualizar a tabela de forma segura
    error_emitted = pyqtSignal(str, str) # Para exibir popups de erro fora da thread
    teste_enviado = pyqtSignal(bool, str)  # (sucesso, mensagem) — resultado do envio de teste


class TemplateBridge(QObject):
    """
    CORREÇÃO CLEAN CODE: Docstring melhorada
    Ponte entre o JavaScript (rodando dentro do QWebEngineView) e o Python.
    Gerencia salvamento e carregamento de templates HTML com CSS inline.
    """
    template_salvo = pyqtSignal(str)  # emite o caminho do arquivo salvo

    def __init__(self, pasta_templates: str) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """
        Args:
            pasta_templates: Caminho da pasta onde templates serão salvos
        """
        super().__init__()
        self.pasta_templates = pasta_templates
        os.makedirs(self.pasta_templates, exist_ok=True)
        logger.info(f"TemplateBridge inicializado em: {self.pasta_templates}")

    @pyqtSlot(result=str)
    def escolher_e_codificar_imagem(self) -> str:
        """Abre um diálogo pra escolher uma imagem local e devolve ela como data URI
        base64 (data:image/png;base64,...). Isso permite inserir imagens no template
        sem precisar hospedar em nenhum serviço externo (Firebase, Imgur, etc.) — a
        imagem fica embutida no próprio arquivo do template. Retorna string vazia se
        o usuário cancelar ou se algo der errado."""
        try:
            caminho, _ = QFileDialog.getOpenFileName(
                None, "Escolha uma imagem",
                "", "Imagens (*.png *.jpg *.jpeg *.gif *.webp)"
            )
            if not caminho:
                return ""

            ext = os.path.splitext(caminho)[1].lower().lstrip(".")
            mime_map = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif", "webp": "webp"}
            mime = mime_map.get(ext)
            if not mime:
                logger.warning(f"Extensão de imagem não suportada: {ext}")
                return ""

            tamanho = os.path.getsize(caminho)
            LIMITE_BYTES = 3 * 1024 * 1024  # 3MB — imagem maior que isso deixa o e-mail pesado
            if tamanho > LIMITE_BYTES:
                logger.warning(f"Imagem descartada por exceder {LIMITE_BYTES} bytes: {tamanho}")
                return "TAMANHO_EXCEDIDO"

            import base64
            with open(caminho, "rb") as f:
                dados_b64 = base64.b64encode(f.read()).decode("ascii")

            return f"data:image/{mime};base64,{dados_b64}"
        except Exception as e:
            logger.error(f"Falha ao codificar imagem: {type(e).__name__}: {e}")
            return ""

    @pyqtSlot(str, str, str, result=str)
    def salvar_template(self, html: str, css: str, nome_arquivo: str) -> str:  # CORREÇÃO CLEAN CODE: Type hints
        """
        Salva um template HTML com CSS inline.
        
        Args:
            html: Código HTML do template
            css: Código CSS do template
            nome_arquivo: Nome do arquivo (será sanitizado)
            
        Returns:
            Caminho do arquivo salvo
            
        Raises:
            ValueError: Se não conseguir sanitizar o nome do arquivo
        """
        try:
            nome_arquivo = nome_arquivo.strip() or "template_sem_nome"
            nome_arquivo = "".join(c for c in nome_arquivo if c.isalnum() or c in ("_", "-", " ")).strip()
            nome_arquivo = nome_arquivo.replace(" ", "_") or "template_sem_nome"

            placeholders = {}
            def protect_vars(texto):
                import re
                def repl(m):
                    key = f"__KEEP_VAR_{len(placeholders)}__"
                    placeholders[key] = m.group(0)
                    return key
                return re.sub(r"\{\{[^}]+\}\}", repl, texto)

            html_protegido = protect_vars(html)
            css_protegido = protect_vars(css)

            html_com_estilo = f'<html><head><meta charset="utf-8"><style>{css_protegido}</style></head><body>{html_protegido}</body></html>'

            try:
                html_final = inline_css(html_com_estilo, base_url=None)
            except Exception as e:
                # CORREÇÃO CLEAN CODE: Usar logging em vez de print()
                logger.warning(f"inline_css falhou, usando HTML sem inline: {type(e).__name__}: {e}")
                html_final = html_com_estilo

            for key, val in placeholders.items():
                html_final = html_final.replace(key, val)

            caminho = os.path.join(self.pasta_templates, f"{nome_arquivo}.html")

            if os.path.exists(caminho):
                resp = QMessageBox.question(
                    None, "Sobrescrever template?",
                    f"Já existe um template salvo com esse nome:\n\n{nome_arquivo}.html\n\n"
                    "Deseja substituir o arquivo existente?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No
                )
                if resp != QMessageBox.StandardButton.Yes:
                    logger.info(f"Usuário cancelou sobrescrever template existente: {caminho}")
                    return ""  # sentinela: JS não mostra "salvo com sucesso" quando vazio

            with open(caminho, "w", encoding="utf-8") as f:
                f.write(html_final)

            self.template_salvo.emit(caminho)
            logger.info(f"Template salvo com sucesso: {caminho}")

            return caminho
        except Exception as e:
            # CORREÇÃO CLEAN CODE: Usar logging em vez de print()
            logger.error(f"Erro ao salvar template: {type(e).__name__}: {e}")
            raise

    @pyqtSlot(str, result=str)
    def carregar_template(self, caminho: str) -> str:  # CORREÇÃO CLEAN CODE: Type hints
        """Carrega um template HTML do disco."""
        if caminho and os.path.exists(caminho):
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    conteudo = f.read()
                    logger.info(f"Template carregado: {caminho}")
                    return conteudo
            except Exception as e:
                logger.error(f"Erro ao carregar template {caminho}: {e}")
                return ""
        return ""

    @pyqtSlot(str, result=str)
    def parse_html_to_blocks(self, html):
        import re
        import json
        from bs4 import BeautifulSoup

        soup = BeautifulSoup(html, 'html.parser')
        body = soup.body or soup

        def clean_text(text):
            return re.sub(r'[\s\u00a0]+', ' ', text or '').strip()

        def get_inline_styles(el):
            styles = {}
            if getattr(el, 'get', None) and el.get('style'):
                for item in el['style'].split(';'):
                    if ':' in item:
                        k, v = item.split(':', 1)
                        v = re.sub(r'!\s*important\s*$', '', v.strip(), flags=re.IGNORECASE).strip()
                        styles[k.strip().lower()] = v
            return styles

        def get_effective_style(el, prop, default='', max_depth=4):
            """Sobe a árvore de elementos pais até achar um valor pra essa propriedade CSS.
            Em HTML de e-mail a cor de fundo/fonte de um título costuma estar no <td>/<div>
            pai (a "caixa" do cabeçalho), não no próprio <h1>/<h2>/texto — por isso o parser
            antigo perdia essas cores quando elas não estavam no elemento exato. Também cai
            pro atributo antigo bgcolor="" quando a propriedade pedida é background-color,
            já que muito HTML de e-mail (compatibilidade Outlook) usa esse atributo em vez
            de style inline."""
            node = el
            depth = 0
            while node is not None and depth <= max_depth and getattr(node, 'name', None):
                val = get_inline_styles(node).get(prop)
                if val and val.strip().lower() not in ('transparent', 'inherit', 'initial', ''):
                    return val.strip()
                if prop == 'background-color' and node.get('bgcolor'):
                    return node.get('bgcolor').strip()
                node = node.parent
                depth += 1
            return default

        def get_effective_align(el, default='left', max_depth=3):
            """Mesma ideia do get_effective_style, mas pro alinhamento — que em e-mail muitas
            vezes vem do atributo HTML align="center" (não de CSS text-align), sobretudo em
            <td align=\"center\">. O parser antigo só olhava text-align inline no elemento
            exato e por isso perdia alinhamento centralizado feito via atributo HTML (era
            essa a causa do rodapé saindo alinhado à esquerda mesmo devendo ficar centralizado)."""
            node = el
            depth = 0
            while node is not None and depth <= max_depth and getattr(node, 'name', None):
                styles = get_inline_styles(node)
                if styles.get('text-align'):
                    return styles['text-align'].strip()
                if node.get('align'):
                    return node.get('align').strip()
                node = node.parent
                depth += 1
            return default

        def get_content_align(el, default='center'):
            """Como get_effective_align, mas olha pra BAIXO (filhos) em vez de subir. Isso e
            necessario pro rodape: o <div> com background-color (onde a checagem de rodape
            acontece) normalmente nao tem text-align proprio -- quem tem e o <div> de texto
            logo dentro dele. get_effective_align nunca acharia isso, pois so sobe a arvore."""
            styles = get_inline_styles(el)
            if styles.get('text-align'):
                return styles['text-align'].strip()
            if el.get('align'):
                return el.get('align').strip()
            inner = el.find(style=re.compile(r'text-align\s*:', re.I))
            if inner:
                inner_align = get_inline_styles(inner).get('text-align')
                if inner_align:
                    return inner_align.strip()
            return default

        def is_dark_color(color):
            """Estima se uma cor é escura, lendo hex (#rgb/#rrggbb) ou rgb()/rgba(), em vez
            de depender de uma lista fechada de códigos hex específicos (que quebra assim
            que a cor real do template diverge por 1 tom)."""
            if not color:
                return False
            color = color.strip().lower()
            try:
                if color.startswith('#'):
                    hexs = color.lstrip('#')
                    if len(hexs) == 3:
                        hexs = ''.join(c * 2 for c in hexs)
                    if len(hexs) != 6:
                        return False
                    r, g, b = int(hexs[0:2], 16), int(hexs[2:4], 16), int(hexs[4:6], 16)
                elif color.startswith('rgb'):
                    # CORREÇÃO BUG #5: Usar regex robusta em vez de findall genérico
                    # O padrão antigo aceitava "1.2.3" como número válido
                    match = re.search(RGB_PATTERN, color)
                    if not match:
                        return False
                    try:
                        r, g, b = float(match.group(1)), float(match.group(2)), float(match.group(3))
                        # Validar que RGB está no intervalo 0-255
                        if not (0 <= r <= 255 and 0 <= g <= 255 and 0 <= b <= 255):
                            return False
                    except (ValueError, AttributeError):
                        return False
                else:
                    return False
                luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
                return luminance < 0.6
            except Exception as e:
                logger.debug(f"Erro ao processar cor '{color}': {e}")  # CORREÇÃO CLEAN CODE: Debug logging
                return False

        ALLOWED_INLINE = {'a', 'b', 'strong', 'em', 'i', 'u', 'span'}

        def extract_rich_lines(el):
            """Extrai as linhas de um elemento preservando tags inline (a, b, strong, em, i, u, span)
            e quebrando apenas em <br>, em vez de get_text(separator=...), que quebra em qualquer
            fronteira de tag (ex: em volta de um <a>) e descarta atributos como href."""
            from bs4 import NavigableString

            def is_full_wrapper_span(node, root):
                """True quando um <span> (ou similar) só existe pra segurar a cor/fonte padrão
                que o editor de origem aplicou ao parágrafo inteiro — cobre TODO o texto do
                bloco (root), não uma seleção específica. Nesse caso a cor do span costuma ser
                só o "reset" do editor de origem (ex: cinza-azulado padrão de framework), não
                uma escolha proposital de destacar uma palavra — preservá-la sobrescreveria a
                cor do bloco de destino sem necessidade."""
                if node.name != 'span':
                    return False
                root_text = clean_text(root.get_text())
                node_text = clean_text(node.get_text())
                return bool(root_text) and root_text == node_text

            def render(node):
                if isinstance(node, NavigableString):
                    return str(node)
                name = (node.name or '').lower()
                if name == 'br':
                    return '\x00BR\x00'
                if name in ALLOWED_INLINE:
                    if is_full_wrapper_span(node, el):
                        return ''.join(render(c) for c in node.children)
                    attrs = ''
                    if name == 'a' and node.get('href'):
                        attrs += f' href="{node.get("href")}"'
                        target = node.get('target')
                        if target:
                            attrs += f' target="{target}"'
                    style = node.get('style')
                    if style:
                        attrs += f' style="{style}"'
                    inner = ''.join(render(c) for c in node.children)
                    return f'<{name}{attrs}>{inner}</{name}>'
                # tags de bloco ou desconhecidas: mantém só o conteúdo (unwrap)
                return ''.join(render(c) for c in node.children)

            raw = ''.join(render(c) for c in el.children)
            raw_parts = raw.split('\x00BR\x00')

            # Um <br> simples mantém a linha dentro do MESMO parágrafo (junta com <br>
            # na hora de renderizar, sem espaçamento extra). Um <br><br> (duplo) sinaliza
            # quebra de parágrafo de verdade e gera um novo item na lista de linhas,
            # que o editor renderiza como um <p> separado (com espaço entre parágrafos).
            paragraphs = []
            current = []
            for part in raw_parts:
                cleaned = re.sub(r'[ \t\u00a0]+', ' ', part).strip()
                if cleaned == '':
                    if current:
                        paragraphs.append('<br>'.join(current))
                        current = []
                else:
                    current.append(cleaned)
            if current:
                paragraphs.append('<br>'.join(current))
            return paragraphs

        def is_footer_container(el, text):
            lower = text.lower()
            # © e "todos os direitos" são sinais fortes de rodapé por si só.
            if '©' in text or 'todos os direitos' in lower:
                return True
            # "desconsiderar o email" sozinho NÃO basta — esse aviso costuma aparecer
            # em um parágrafo normal (fundo branco), não dentro da caixa de rodapé.
            # Só conta como rodapé se também tiver um fundo escuro/acinzentado típico de rodapé.
            # Em vez de comparar contra uma lista fechada de hex (que erra assim que o tom
            # muda um pouco), estima se a cor efetiva (inclusive herdada do pai) é escura,
            # ou reconhece os tons de cinza claro de rodapé mais comuns.
            bg = get_effective_style(el, 'background-color', '').lower()
            light_footer_grays = ('#f5f5f5', '#f1f5f9', '#f8f9fa', '#eeeeee')
            if is_dark_color(bg) or bg in light_footer_grays:
                if len(text) < 250 and ('2026' in text or 'direitos' in lower or 'desconsiderar' in lower):
                    return True
            return False

        def parse_container(el):
            if not el or not getattr(el, 'name', None):
                return []

            tag = el.name.lower()
            inline = get_inline_styles(el)
            text = clean_text(el.get_text())

            if tag in ('script', 'style', 'head', 'meta', 'link', 'title'):
                return []

            if tag == 'img':
                src = el.get('src', '')
                parent_a = el.find_parent('a')
                # É logo/header (não banner de conteúdo) quando não existe nenhum título ou
                # parágrafo com texto ANTES dela no documento — ou seja, é a primeira coisa
                # visual do e-mail, antes de qualquer título/texto.
                previous_texts = [
                    clean_text(sib.get_text())
                    for sib in el.find_all_previous(['h1', 'h2', 'h3', 'p'])
                ]
                is_first_visual = not any(previous_texts)
                block_type = 'header-image' if is_first_visual else 'banner-image'
                props = {
                    'imgUrl': src or 'https://via.placeholder.com/600x200',
                    'alt': el.get('alt', 'Banner'),
                    'link': parent_a.get('href', '') if parent_a else '',
                    'width': str(el.get('width', '600'))
                }
                if block_type == 'header-image':
                    # Cor de fundo em volta da logo geralmente vem do <td>/<div> que envolve a
                    # imagem, não da tag <img> em si.
                    props['backgroundColor'] = get_effective_style(el, 'background-color', '#ffffff')
                    props['fontFamily'] = get_effective_style(el, 'font-family', 'Segoe UI, Arial, sans-serif')
                return [{'type': block_type, 'props': props}]

            # Identifica botões reais de CTA isolados
            if tag == 'a' and ('button' in ' '.join(el.get('class', [])) or 'display: inline-block' in el.get('style', '').lower() or 'background-color' in inline):
                if len(text) < 50:
                    return [{
                        'type': 'button',
                        'props': {
                            'label': text or 'Clique Aqui',
                            'url': el.get('href', '#'),
                            'bgColor': inline.get('background-color', '#4f46e5'),
                            'textColor': inline.get('color', '#ffffff'),
                            'fontFamily': get_effective_style(el, 'font-family', 'Segoe UI, Arial, sans-serif'),
                            'fontSize': inline.get('font-size', '14px'),
                            'align': get_effective_align(el, 'center')
                        }
                    }]

            if is_footer_container(el, text):
                lines = extract_rich_lines(el)
                return [{
                    'type': 'footer',
                    'props': {
                        'lines': lines or [text],
                        'textColor': get_effective_style(el, 'color', '#64748b'),
                        'bgColor': get_effective_style(el, 'background-color', '#1e1b4b'),
                        'align': get_content_align(el, 'center'),
                        'fontFamily': get_effective_style(el, 'font-family', 'Segoe UI, Arial, sans-serif'),
                        'fontSize': inline.get('font-size', '12px')
                    }
                }]

            if tag in ('h1', 'h2', 'h3'):
                # A cor de fundo do "título com faixa colorida" (ex: a barra azul com
                # "SUA MATRÍCULA ESTÁ QUASE COMPLETA") normalmente está no <td>/<div> pai que
                # envolve o heading, não no <h1>/<h2>/<h3> em si — por isso sobe a árvore.
                bg = get_effective_style(el, 'background-color', '')
                b_type = 'header-text' if bg and bg not in ('#ffffff', '#fff', 'white', 'rgb(255, 255, 255)') else 'title'
                font_family = get_effective_style(el, 'font-family', 'Segoe UI, Arial, sans-serif')
                # <h1>/<h2>/<h3> são negrito por padrão no navegador mesmo sem CSS explícito —
                # então se não houver font-weight explícito, assume negrito (é o que o
                # HTML original realmente mostrava). Só desliga se o CSS disser normal/400-.
                font_weight_raw = get_effective_style(el, 'font-weight', '')
                is_bold = font_weight_raw.strip().lower() not in ('normal', '400', '300', '200', '100') if font_weight_raw else True
                return [{
                    'type': b_type,
                    'props': {
                        'text': text,
                        'bold': is_bold,
                        'align': get_effective_align(el, 'center' if b_type == 'header-text' else 'left'),
                        'fontSize': inline.get('font-size', '18px'),
                        'fontFamily': font_family,
                        'textColor': inline.get('color', get_effective_style(el, 'color', '#ffffff' if b_type == 'header-text' else '#1e1b4b')),
                        'bgColor': bg if bg else '#ffffff',
                        'backgroundColor': bg if bg else '#ffffff'
                    }
                }]

            # E-mails corporativos quase sempre são montados com <table><tr><td> em vez de
            # <div> (é o jeito "email-safe" que funciona no Outlook), então table/tr/td
            # também precisam entrar como containers a percorrer recursivamente — antes
            # disso, tudo que estivesse dentro de uma tabela caía direto no fallback de
            # texto genérico, perdendo cor de fundo, fonte e alinhamento do <td> real.
            children_tags = ['p', 'h1', 'h2', 'h3', 'h4', 'img', 'a', 'div', 'table', 'tbody', 'tr', 'td', 'th']
            direct_children = [c for c in el.find_all(recursive=False) if getattr(c, 'name', None) in children_tags]

            if direct_children and tag in ('div', 'td', 'th', 'section', 'article', 'card', 'table', 'tr', 'tbody'):
                res = []
                bg = inline.get('background-color', '') or (el.get('bgcolor', '').strip() if el.get('bgcolor') else '')
                
                if bg and bg not in ('#ffffff', '#fff', 'white', 'rgb(255, 255, 255)', ''):
                    heading_el = el.find(['h1', 'h2', 'h3'])
                    # Bug real encontrado: el.find() busca o heading em QUALQUER profundidade
                    # dentro de "el". Se "el" for um wrapper GRANDE que engloba tanto a faixa
                    # do logo quanto a faixa do título (cada uma com sua própria cor), esse
                    # atalho pegava a cor do wrapper externo (ex: azul do logo) em vez da cor
                    # própria e mais específica da faixa onde o título realmente está (ex: azul
                    # -marinho mais escuro). Confirma que "bg" é de fato a cor MAIS PRÓXIMA do
                    # heading antes de usar esse atalho — senão deixa a recursão normal achar o
                    # container certo (mais interno) quando chegar nele.
                    closest_bg_to_heading = get_effective_style(heading_el, 'background-color', '') if heading_el else ''
                    if heading_el and closest_bg_to_heading.lower() == bg.lower():
                        font_weight_raw = get_effective_style(heading_el, 'font-weight', '')
                        is_bold = font_weight_raw.strip().lower() not in ('normal', '400', '300', '200', '100') if font_weight_raw else True
                        res.append({
                            'type': 'header-text',
                            'props': {
                                'text': clean_text(heading_el.get_text()),
                                'bold': is_bold,
                                'textColor': get_inline_styles(heading_el).get('color', inline.get('color', '#ffffff')),
                                'bgColor': bg,
                                'backgroundColor': bg,
                                'fontFamily': get_effective_style(heading_el, 'font-family', 'Segoe UI, Arial, sans-serif'),
                                'fontSize': get_inline_styles(heading_el).get('font-size', '18px'),
                                'align': get_effective_align(heading_el, 'center')
                            }
                        })
                        for child in direct_children:
                            if child != heading_el:
                                res.extend(parse_container(child))
                        return res

                for child in direct_children:
                    res.extend(parse_container(child))
                if res:
                    return res

            if text:
                if tag == 'a':
                    return []

                lines = extract_rich_lines(el)
                if inline.get('font-style', '').lower() == 'italic':
                    lines = [f'<i>{line}</i>' for line in lines]
                return [{
                    'type': 'text',
                    'props': {
                        'lines': lines,
                        'align': get_effective_align(el, 'left'),
                        'backgroundColor': get_effective_style(el, 'background-color', '#ffffff'),
                        'fontFamily': get_effective_style(el, 'font-family', 'Segoe UI, Arial, sans-serif'),
                        'fontSize': inline.get('font-size', '14px'),
                        'textColor': get_effective_style(el, 'color', '#555555')
                    }
                }]

            return []

        blocks = []
        card_container = body.find(class_=re.compile('card', re.I)) or body

        for child in card_container.find_all(recursive=False):
            if getattr(child, 'name', None):
                parsed = parse_container(child)
                for b in parsed:
                    if b['type'] == 'footer' and blocks and blocks[-1]['type'] == 'footer':
                        blocks[-1]['props']['lines'].extend(b['props']['lines'])
                    else:
                        blocks.append(b)

        if not blocks:
            blocks.append({'type': 'text', 'props': {'lines': [clean_text(body.get_text()) or "Conteúdo importado"]}})

        return json.dumps({'blocks': blocks}, ensure_ascii=False)


class R9BotQtApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("R9Bot - Mail Sender Professional")
        self.resize(980, 720)
        
        # DESIGN SYSTEM: Aplicar QSS Global Corporativo
        # Remove o setStyleSheet inline (~40 linhas) e centraliza em app_style.qss
        # Benefícios: Manutenção fácil, consistência visual, alterações rápidas
        qss_global = carregar_stylesheet_global()
        if qss_global:
            self.setStyleSheet(qss_global)
            logger.info("Design System Global aplicado com sucesso")
        else:
            logger.warning("Usando estilos padrão do sistema")

        self.executando = False
        self.ultimo_log_path = None
        self.lista_anexos: List[str] = []
        self.num_variaveis = 3  # padrão; pode ser ajustado pelo usuário (➕/➖ Variável) e é persistido
        self._lock_executando = Lock()  # CORREÇÃO BUG #1: Lock para race condition em self.executando
        self._lock_log = Lock()
        self._linhas_log_atual: List = []  # linhas do log da campanha em andamento (parcial, até agora)
        self.thread_envio = None
        
        # Conexão Segura de Sinais (Thread-Safety)
        self.signals = WorkerSignals()
        self.signals.status_updated.connect(self.atualizar_status_ui)
        self.signals.progress_updated.connect(self.atualizar_progresso)
        self.signals.finished.connect(self.fim_disparo)
        self.signals.row_updated.connect(self.atualizar_status_tabela)
        self.signals.error_emitted.connect(self.mostrar_erro_thread)
        self.signals.teste_enviado.connect(self.resultado_envio_teste)

        # Usar diretório físico persistente para os templates salvos e não o temporário do MEIPASS
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.pasta_templates = os.path.join(base_dir, "templates_salvos")
        self.pasta_dados = os.path.join(base_dir, "dados_app")
        self.pasta_logs = os.path.join(base_dir, "logs_envio")
        os.makedirs(self.pasta_dados, exist_ok=True)
        os.makedirs(self.pasta_logs, exist_ok=True)
        self.arquivo_contatos = os.path.join(self.pasta_dados, "contatos.json")
        self.arquivo_config = os.path.join(self.pasta_dados, "config.json")

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.tab_dashboard = QWidget()
        self.tab_contatos = QWidget()
        self.tab_editor = QWidget()
        self.tab_historico = QWidget()

        self.tabs.addTab(self.tab_dashboard, "📊 Dashboard de Disparo")
        self.tabs.addTab(self.tab_contatos, "👥 Gerenciador de Contatos")
        self.tabs.addTab(self.tab_editor, "🎨 Editor de Templates")
        self.tabs.addTab(self.tab_historico, "🕓 Histórico de Campanhas")

        # Dicas de fluxo — a ordem das abas na tela não muda (mantendo o que quem já usa
        # o app está acostumado), mas o tooltip ajuda quem está vendo isso pela primeira vez
        # a entender por onde começar.
        self.tabs.setTabToolTip(0, "Passo 3: configure assunto, anexos e dispare a campanha aqui (depois de montar o template e a lista de contatos)")
        self.tabs.setTabToolTip(1, "Passo 2: cole ou digite a lista de contatos e variáveis aqui")
        self.tabs.setTabToolTip(2, "Passo 1: monte o e-mail (ou importe um HTML pronto) aqui")
        self.tabs.setTabToolTip(3, "Consulte campanhas já enviadas anteriormente")

        self.criar_aba_dashboard()
        self.criar_aba_contatos()
        self.criar_aba_editor()
        self.criar_aba_historico()

        # Restaura o que foi salvo da última vez que o app foi usado (config + contatos),
        # senão o usuário tem que preencher tudo de novo e recolar a lista toda vez que
        # abre o programa.
        self.carregar_config()
        self.carregar_contatos_salvos()
        self.mostrar_onboarding_se_necessario()

        # Auto-salva a tabela de contatos sempre que algo muda (colar do Excel, editar
        # célula, etc.), com um pequeno atraso pra não salvar a cada tecla digitada.
        self._timer_autosave_contatos = QTimer(self)
        self._timer_autosave_contatos.setSingleShot(True)
        self._timer_autosave_contatos.timeout.connect(self.salvar_contatos_em_disco)
        self.tabela.itemChanged.connect(lambda _item: self._timer_autosave_contatos.start(1500))

    def closeEvent(self, event):
        """Salva config e contatos automaticamente ao fechar o app. Se houver uma
        campanha em andamento, avisa o usuário antes de fechar de verdade (fechar no
        meio do envio mata a thread sem aviso e sem log) e salva um log parcial do que
        já foi enviado até esse momento."""
        if self.executando:
            resp = QMessageBox.question(
                self, "Campanha em andamento",
                "Uma campanha de envio ainda está em execução!\n\n"
                "Se você fechar agora, o disparo será interrompido no meio e os contatos "
                "restantes NÃO receberão o e-mail. Um log parcial com o que já foi "
                "enviado até agora será salvo antes de fechar.\n\n"
                "Deseja realmente fechar o programa?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if resp != QMessageBox.StandardButton.Yes:
                event.ignore()
                return

            with self._lock_executando:
                self.executando = False
            with self._lock_log:
                caminho_parcial = self.salvar_log_envio(list(self._linhas_log_atual))
            if caminho_parcial:
                logger.info(f"Log parcial salvo antes de fechar durante campanha: {caminho_parcial}")

        try:
            self.salvar_config()
            self.salvar_contatos_em_disco()
        except Exception as e:
            logger.warning(f"Falha ao salvar dados ao fechar: {e}")
        super().closeEvent(event)

    def mostrar_onboarding_se_necessario(self) -> None:
        """Mostra uma explicação rápida da ordem recomendada de uso, só na primeira vez
        que o app é aberto nessa máquina (controlado por um arquivo marcador em
        dados_app/). Sem isso, um usuário novo vê 4 abas lado a lado sem nenhuma pista
        de por onde começar."""
        marcador = os.path.join(self.pasta_dados, "onboarding_visto.txt")
        if os.path.exists(marcador):
            return
        QMessageBox.information(
            self, "Bem-vindo ao R9Bot Mail Sender",
            "Pra mandar sua primeira campanha, essa é a ordem recomendada:\n\n"
            "1️⃣  Editor de Templates — monte o e-mail (ou clique em \"Importar HTML\" pra usar um já pronto)\n\n"
            "2️⃣  Gerenciador de Contatos — cole a lista de destinatários e variáveis (Ctrl+V do Excel)\n\n"
            "3️⃣  Dashboard de Disparo — escolha o template salvo, preencha o assunto e dispare\n\n"
            "Dica: use o botão \"🧪 Enviar Teste\" no Dashboard pra conferir como o e-mail chega antes de disparar pra lista toda."
        )
        try:
            with open(marcador, "w", encoding="utf-8") as f:
                f.write("ok")
        except Exception as e:
            logger.warning(f"Falha ao gravar marcador de onboarding: {e}")

    def salvar_config(self):
        try:
            config = {
                "assunto": self.txt_assunto.text(),
                "template_path": self.txt_template.text(),
                "anexos": self.lista_anexos,
                "intervalo": self.txt_intervalo.text(),
                "cc": self.txt_cc.text(),
                "bcc": self.txt_bcc.text(),
                "num_variaveis": self.num_variaveis,
            }
            with open(self.arquivo_config, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"Falha ao salvar config: {e}")

    def carregar_config(self):
        if not os.path.exists(self.arquivo_config):
            return
        try:
            with open(self.arquivo_config, "r", encoding="utf-8") as f:
                config = json.load(f)
            if config.get("assunto"):
                self.txt_assunto.setText(config["assunto"])
            if config.get("template_path") and os.path.exists(config["template_path"]):
                self.txt_template.setText(config["template_path"])
            anexos_salvos = config.get("anexos") or []
            # Compatibilidade com o formato antigo (um único "anexo_path" em vez de lista)
            if not anexos_salvos and config.get("anexo_path"):
                anexos_salvos = [config["anexo_path"]]
            self.lista_anexos = [a for a in anexos_salvos if os.path.exists(a)]
            self._atualizar_texto_anexos()
            if config.get("intervalo"):
                self.txt_intervalo.setText(config["intervalo"])
            if config.get("cc"):
                self.txt_cc.setText(config["cc"])
            if config.get("bcc"):
                self.txt_bcc.setText(config["bcc"])
            # A tabela já foi criada com 3 variáveis por padrão — se o usuário tinha
            # configurado um número diferente da última vez, ajusta as colunas AGORA,
            # antes de carregar os contatos salvos (que dependem dessa estrutura).
            num_salvo = config.get("num_variaveis")
            if num_salvo and num_salvo != self.num_variaveis:
                self.ajustar_num_variaveis_para(int(num_salvo))
        except Exception as e:
            logger.warning(f"Falha ao carregar config: {e}")

    def salvar_contatos_em_disco(self):
        try:
            contatos = self.obter_lista_contatos()
            with open(self.arquivo_contatos, "w", encoding="utf-8") as f:
                json.dump(contatos, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.warning(f"Falha ao salvar contatos: {e}")

    def carregar_contatos_salvos(self):
        if not os.path.exists(self.arquivo_contatos):
            return
        try:
            with open(self.arquivo_contatos, "r", encoding="utf-8") as f:
                contatos = json.load(f)
            if not contatos:
                return
            self.tabela.blockSignals(True)
            self.tabela.setRowCount(max(len(contatos) + 10, 25))
            for r_idx, c in enumerate(contatos):
                self.tabela.setItem(r_idx, 0, QTableWidgetItem(c.get("Nome", "")))
                self.tabela.setItem(r_idx, 1, QTableWidgetItem(c.get("Email", "")))
                for i in range(1, self.num_variaveis + 1):
                    self.tabela.setItem(r_idx, self.col_var(i), QTableWidgetItem(c.get(f"Var{i}", "")))
                self.tabela.setItem(r_idx, self.col_status(), QTableWidgetItem(c.get("Status", "")))
            self.tabela.blockSignals(False)
            self.atualizar_contador_contatos()
        except Exception as e:
            logger.warning(f"Falha ao carregar contatos salvos: {e}")

    def criar_aba_dashboard(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(12)

        header_layout = QVBoxLayout()
        lbl_titulo = QLabel("R9BOT - MAIL SENDER")
        lbl_titulo.setFont(QFont("Segoe UI", 22, QFont.Weight.Bold))
        lbl_titulo.setProperty("class", "h1")
        
        lbl_sub = QLabel("Automação profissional de campanhas via Microsoft Outlook")
        lbl_sub.setProperty("class", "caption")
        
        header_layout.addWidget(lbl_titulo)
        header_layout.addWidget(lbl_sub)
        main_layout.addLayout(header_layout)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(line)

        self.form_container = QWidget()
        form_layout = QVBoxLayout(self.form_container)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(12)

        form_layout.addWidget(QLabel("<b>Assunto do E-mail:</b>"))
        lbl_assunto = form_layout.itemAt(form_layout.count() - 1).widget()
        if lbl_assunto:
            lbl_assunto.setProperty("class", "label")
        
        self.txt_assunto = QLineEdit("Última oportunidade")
        self.txt_assunto.setToolTip("Você pode usar {{nome}}, {{var1}}, {{var2}}, etc. aqui também, igual no corpo do e-mail (a quantidade de variáveis é a que estiver configurada na aba Contatos).")
        self.txt_assunto.editingFinished.connect(self.salvar_config)
        form_layout.addWidget(self.txt_assunto)

        h_layout_cc_bcc = QHBoxLayout()
        v_cc = QVBoxLayout()
        lbl_cc = QLabel("<b>CC (opcional):</b>")
        lbl_cc.setProperty("class", "label")
        v_cc.addWidget(lbl_cc)
        self.txt_cc = QLineEdit()
        self.txt_cc.setPlaceholderText("email1@exemplo.com; email2@exemplo.com")
        self.txt_cc.editingFinished.connect(self.salvar_config)
        v_cc.addWidget(self.txt_cc)
        h_layout_cc_bcc.addLayout(v_cc)

        v_bcc = QVBoxLayout()
        lbl_bcc = QLabel("<b>CCO / BCC (opcional):</b>")
        lbl_bcc.setProperty("class", "label")
        v_bcc.addWidget(lbl_bcc)
        self.txt_bcc = QLineEdit()
        self.txt_bcc.setPlaceholderText("email1@exemplo.com; email2@exemplo.com")
        self.txt_bcc.editingFinished.connect(self.salvar_config)
        v_bcc.addWidget(self.txt_bcc)
        h_layout_cc_bcc.addLayout(v_bcc)

        form_layout.addLayout(h_layout_cc_bcc)

        form_layout.addWidget(QLabel("<b>Template HTML:</b>"))
        lbl_template = form_layout.itemAt(form_layout.count() - 1).widget()
        if lbl_template:
            lbl_template.setProperty("class", "label")
        
        h_layout_template = QHBoxLayout()
        self.txt_template = QLineEdit()
        self.txt_template.setPlaceholderText("Selecione o arquivo .html...")
        
        self.btn_sel_template = QPushButton("Examinar...")
        self.btn_sel_template.setFixedWidth(100)
        self.btn_sel_template.setProperty("class", "secondary")
        self.btn_sel_template.clicked.connect(self.selecionar_template)

        self.btn_templates_salvos = QPushButton("📋 Salvos")
        self.btn_templates_salvos.setFixedWidth(90)
        self.btn_templates_salvos.setProperty("class", "secondary")
        self.btn_templates_salvos.clicked.connect(self.abrir_lista_templates_salvos)

        h_layout_template.addWidget(self.txt_template)
        h_layout_template.addWidget(self.btn_sel_template)
        h_layout_template.addWidget(self.btn_templates_salvos)
        form_layout.addLayout(h_layout_template)

        form_layout.addWidget(QLabel("<b>Anexos (Opcional):</b>"))
        lbl_anexo = form_layout.itemAt(form_layout.count() - 1).widget()
        if lbl_anexo:
            lbl_anexo.setProperty("class", "label")
        
        h_layout_anexo = QHBoxLayout()
        self.txt_anexo = QLineEdit()
        self.txt_anexo.setReadOnly(True)
        self.txt_anexo.setPlaceholderText("Nenhum anexo selecionado...")
        
        self.btn_sel_anexo = QPushButton("Adicionar...")
        self.btn_sel_anexo.setFixedWidth(100)
        self.btn_sel_anexo.setProperty("class", "secondary")
        self.btn_sel_anexo.clicked.connect(self.selecionar_anexo)

        self.btn_limpar_anexos = QPushButton("Limpar")
        self.btn_limpar_anexos.setFixedWidth(80)
        self.btn_limpar_anexos.setProperty("class", "secondary")
        self.btn_limpar_anexos.clicked.connect(self.limpar_anexos)
        
        h_layout_anexo.addWidget(self.txt_anexo)
        h_layout_anexo.addWidget(self.btn_sel_anexo)
        h_layout_anexo.addWidget(self.btn_limpar_anexos)
        form_layout.addLayout(h_layout_anexo)

        form_layout.addWidget(QLabel("<b>Intervalo entre disparos (Segundos):</b>"))
        lbl_intervalo = form_layout.itemAt(form_layout.count() - 1).widget()
        if lbl_intervalo:
            lbl_intervalo.setProperty("class", "label")
        
        self.txt_intervalo = QLineEdit("2")
        self.txt_intervalo.setFixedWidth(100)
        self.txt_intervalo.editingFinished.connect(self.salvar_config)
        form_layout.addWidget(self.txt_intervalo)

        main_layout.addWidget(self.form_container)

        self.lbl_status = QLabel("Status: Pronto para iniciar a campanha.")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_status.setProperty("class", "alert-info")
        self.lbl_status.setMaximumHeight(40)
        main_layout.addWidget(self.lbl_status)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        botoes_layout = QHBoxLayout()
        botoes_layout.setSpacing(12)

        self.btn_limpar_status = QPushButton("🔄 Resetar Status")
        self.btn_limpar_status.setProperty("class", "ghost")
        self.btn_limpar_status.clicked.connect(self.limpar_status)

        self.btn_previa = QPushButton("👁️ Pré-visualizar HTML")
        self.btn_previa.setProperty("class", "tertiary")
        self.btn_previa.clicked.connect(self.previsualizar_html)

        self.btn_teste = QPushButton("🧪 Enviar Teste")
        self.btn_teste.setProperty("class", "tertiary")
        self.btn_teste.clicked.connect(self.enviar_email_teste)

        botoes_layout.addWidget(self.btn_limpar_status)
        botoes_layout.addWidget(self.btn_previa)
        botoes_layout.addWidget(self.btn_teste)
        main_layout.addLayout(botoes_layout)

        acao_principal_layout = QHBoxLayout()
        acao_principal_layout.setSpacing(12)

        self.btn_enviar = QPushButton("🚀 INICIAR DISPARO")
        self.btn_enviar.setProperty("class", "primary")
        self.btn_enviar.clicked.connect(self.executar_campanha)

        self.btn_pausar = QPushButton("⏸ PAUSAR / PARAR")
        self.btn_pausar.setProperty("class", "danger")
        self.btn_pausar.clicked.connect(self.pausar_campanha)

        acao_principal_layout.addWidget(self.btn_enviar)
        acao_principal_layout.addWidget(self.btn_pausar)
        main_layout.addLayout(acao_principal_layout)

        self.tab_dashboard.setLayout(main_layout)

    def criar_aba_contatos(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        top_layout = QHBoxLayout()
        lbl_titulo = QLabel("Gerenciador de Contatos e Variáveis")
        lbl_titulo.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        lbl_titulo.setProperty("class", "h2")
        top_layout.addWidget(lbl_titulo)

        top_layout.addStretch()

        self.lbl_contador = QLabel("0 contatos carregados")
        self.lbl_contador.setProperty("class", "status")
        top_layout.addWidget(self.lbl_contador)
        
        layout.addLayout(top_layout)

        self.lbl_dica_variaveis = QLabel()
        self.lbl_dica_variaveis.setProperty("class", "caption")
        layout.addWidget(self.lbl_dica_variaveis)

        h_layout_busca = QHBoxLayout()
        self.txt_busca_contato = QLineEdit()
        self.txt_busca_contato.setPlaceholderText("🔍 Buscar por nome ou e-mail...")
        self.txt_busca_contato.textChanged.connect(self.filtrar_contatos)
        h_layout_busca.addWidget(self.txt_busca_contato)

        self.btn_remover_variavel = QPushButton("➖ Variável")
        self.btn_remover_variavel.setProperty("class", "secondary")
        self.btn_remover_variavel.setToolTip("Remove a última coluna de variável (Var" + str(self.num_variaveis) + ")")
        self.btn_remover_variavel.clicked.connect(self.remover_coluna_variavel)
        h_layout_busca.addWidget(self.btn_remover_variavel)

        self.btn_adicionar_variavel = QPushButton("➕ Variável")
        self.btn_adicionar_variavel.setProperty("class", "secondary")
        self.btn_adicionar_variavel.setToolTip("Adiciona uma nova coluna de variável (ex: Var4, Var5...)")
        self.btn_adicionar_variavel.clicked.connect(self.adicionar_coluna_variavel)
        h_layout_busca.addWidget(self.btn_adicionar_variavel)

        self.btn_exportar_contatos = QPushButton("⬇️ Exportar Contatos")
        self.btn_exportar_contatos.setProperty("class", "secondary")
        self.btn_exportar_contatos.clicked.connect(self.exportar_contatos)
        h_layout_busca.addWidget(self.btn_exportar_contatos)
        layout.addLayout(h_layout_busca)

        self.tabela = QTableWidget(25, 2 + self.num_variaveis + 1)
        self.atualizar_cabecalho_tabela()
        self.tabela.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tabela.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        
        header = self.tabela.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        
        layout.addWidget(self.tabela)

        self.tabela.itemChanged.connect(self.atualizar_contador_contatos)

        self.shortcut_paste = QShortcut(QKeySequence.StandardKey.Paste, self.tabela)
        self.shortcut_paste.activated.connect(self.colar_do_excel)

        self.shortcut_delete = QShortcut(QKeySequence.StandardKey.Delete, self.tabela)
        self.shortcut_delete.activated.connect(self.apagar_selecionado)

        h_layout_acoes = QHBoxLayout()
        h_layout_acoes.setSpacing(12)

        self.btn_colar = QPushButton("📋 Colar do Excel")
        self.btn_colar.setProperty("class", "secondary")
        self.btn_colar.setToolTip("Cole aqui o que copiou do Excel (mesmo efeito de dar Ctrl+V com a tabela selecionada)")
        self.btn_colar.clicked.connect(self.colar_do_excel)
        h_layout_acoes.addWidget(self.btn_colar)

        self.btn_importar = QPushButton("📥 Importar Planilha")
        self.btn_importar.setProperty("class", "secondary")
        self.btn_importar.clicked.connect(self.importar_excel)
        h_layout_acoes.addWidget(self.btn_importar)

        h_layout_acoes.addStretch()

        self.btn_limpar_tabela = QPushButton("🗑️ Limpar Tabela")
        self.btn_limpar_tabela.setProperty("class", "danger")
        self.btn_limpar_tabela.clicked.connect(self.limpar_tabela)
        h_layout_acoes.addWidget(self.btn_limpar_tabela)

        layout.addLayout(h_layout_acoes)
        self.tab_contatos.setLayout(layout)

    def filtrar_contatos(self, texto: str) -> None:
        """Esconde as linhas da tabela que não têm o texto buscado no Nome ou E-mail.
        Não mexe nos dados, só na visibilidade — útil pra achar um contato específico
        numa lista grande sem precisar rolar manualmente."""
        termo = texto.strip().lower()
        for row in range(self.tabela.rowCount()):
            if not termo:
                self.tabela.setRowHidden(row, False)
                continue
            nome_item = self.tabela.item(row, 0)
            email_item = self.tabela.item(row, 1)
            nome = nome_item.text().lower() if nome_item else ""
            email = email_item.text().lower() if email_item else ""
            self.tabela.setRowHidden(row, termo not in nome and termo not in email)

    def exportar_contatos(self) -> None:
        """Exporta a lista de contatos atual pra um arquivo CSV ou Excel — útil pra
        conferir a lista fora do app ou levar os dados pra outro lugar."""
        contatos = self.obter_lista_contatos()
        if not contatos:
            QMessageBox.information(self, "Exportar Contatos", "Não há contatos pra exportar.")
            return

        caminho, filtro = QFileDialog.getSaveFileName(
            self, "Exportar Contatos", "contatos.xlsx",
            "Planilha Excel (*.xlsx);;CSV (*.csv)"
        )
        if not caminho:
            return

        try:
            colunas_var = [f"Var{i}" for i in range(1, self.num_variaveis + 1)]
            df = pd.DataFrame([
                {"Nome": c["Nome"], "Email": c["Email"], **{v: c.get(v, "") for v in colunas_var}, "Status": c["Status"]}
                for c in contatos
            ])

            if caminho.lower().endswith(".csv"):
                df.to_csv(caminho, index=False, sep=";", encoding="utf-8-sig")
            else:
                if not caminho.lower().endswith(".xlsx"):
                    caminho += ".xlsx"
                df.to_excel(caminho, index=False)

            QMessageBox.information(self, "Exportar Contatos", f"Contatos exportados com sucesso para:\n{caminho}")
        except Exception as e:
            logger.error(f"Falha ao exportar contatos: {type(e).__name__}: {e}")
            QMessageBox.critical(self, "Erro", f"Não foi possível exportar os contatos:\n{e}")

    def criar_aba_historico(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        top_layout = QHBoxLayout()
        lbl_titulo = QLabel("Histórico de Campanhas")
        lbl_titulo.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        lbl_titulo.setProperty("class", "h2")
        top_layout.addWidget(lbl_titulo)
        top_layout.addStretch()

        self.btn_atualizar_historico = QPushButton("🔄 Atualizar")
        self.btn_atualizar_historico.setProperty("class", "secondary")
        self.btn_atualizar_historico.clicked.connect(self.atualizar_historico_campanhas)
        top_layout.addWidget(self.btn_atualizar_historico)

        self.btn_abrir_pasta_logs = QPushButton("📂 Abrir Pasta de Logs")
        self.btn_abrir_pasta_logs.setProperty("class", "secondary")
        self.btn_abrir_pasta_logs.clicked.connect(self.abrir_pasta_logs)
        top_layout.addWidget(self.btn_abrir_pasta_logs)

        layout.addLayout(top_layout)

        lbl_sub = QLabel("Cada campanha (ou teste parcial interrompido) gera um arquivo CSV aqui. Clique duas vezes numa linha pra abrir o arquivo.")
        lbl_sub.setProperty("class", "caption")
        layout.addWidget(lbl_sub)

        self.tabela_historico = QTableWidget(0, 5)
        self.tabela_historico.setHorizontalHeaderLabels(["Data/Hora", "Arquivo", "Enviados", "Erros/Inválidos", "Total"])
        self.tabela_historico.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabela_historico.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        header = self.tabela_historico.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.tabela_historico.doubleClicked.connect(self.abrir_log_selecionado)
        layout.addWidget(self.tabela_historico)

        self.tab_historico.setLayout(layout)
        self.atualizar_historico_campanhas()

    def atualizar_historico_campanhas(self) -> None:
        """Lê todos os CSVs de logs_envio/ e monta a lista de campanhas anteriores,
        com um resumo rápido (quantos enviados, quantos com erro) de cada uma."""
        try:
            arquivos = [f for f in os.listdir(self.pasta_logs) if f.lower().endswith(".csv")]
        except Exception as e:
            logger.warning(f"Falha ao listar histórico de campanhas: {e}")
            arquivos = []

        arquivos.sort(key=lambda f: os.path.getmtime(os.path.join(self.pasta_logs, f)), reverse=True)

        self.tabela_historico.setRowCount(len(arquivos))
        for row, nome_arquivo in enumerate(arquivos):
            caminho = os.path.join(self.pasta_logs, nome_arquivo)
            data_mod = time.strftime("%d/%m/%Y %H:%M", time.localtime(os.path.getmtime(caminho)))
            enviados, erros, total = 0, 0, 0
            try:
                import csv
                with open(caminho, "r", encoding="utf-8-sig", newline="") as f:
                    reader = csv.reader(f, delimiter=";")
                    next(reader, None)  # pula o cabeçalho
                    for linha in reader:
                        if len(linha) < 4:
                            continue
                        total += 1
                        if linha[3] == "ENVIADO":
                            enviados += 1
                        else:
                            erros += 1
            except Exception as e:
                logger.warning(f"Falha ao ler log {nome_arquivo}: {e}")

            self.tabela_historico.setItem(row, 0, QTableWidgetItem(data_mod))
            self.tabela_historico.setItem(row, 1, QTableWidgetItem(nome_arquivo))
            self.tabela_historico.setItem(row, 2, QTableWidgetItem(str(enviados)))
            self.tabela_historico.setItem(row, 3, QTableWidgetItem(str(erros)))
            self.tabela_historico.setItem(row, 4, QTableWidgetItem(str(total)))
            self.tabela_historico.item(row, 1).setData(Qt.ItemDataRole.UserRole, caminho)

    def abrir_log_selecionado(self) -> None:
        row = self.tabela_historico.currentRow()
        if row < 0:
            return
        item = self.tabela_historico.item(row, 1)
        if not item:
            return
        caminho = item.data(Qt.ItemDataRole.UserRole)
        if caminho and os.path.exists(caminho):
            try:
                subprocess.run(['start', '', caminho], shell=True)
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Não foi possível abrir o arquivo:\n{e}")

    def abrir_pasta_logs(self) -> None:
        try:
            subprocess.run(['explorer', self.pasta_logs])
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível abrir a pasta:\n{e}")

    def criar_aba_editor(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        if not PREMAILER_DISPONIVEL:
            aviso = QLabel(
                "⚠ Pacote 'premailer' não encontrado — o CSS será salvo sem inline "
                "(pode não aparecer corretamente no Outlook). Rode: pip install premailer"
            )
            aviso.setProperty("class", "alert-warning")
            layout.addWidget(aviso)

        self.web_editor = QWebEngineView()
        self.channel = QWebChannel()
        self.template_bridge = TemplateBridge(self.pasta_templates)
        self.template_bridge.template_salvo.connect(self.template_recebido)
        self.channel.registerObject("bridge", self.template_bridge)
        self.web_editor.page().setWebChannel(self.channel)

        # O botão "Exportar HTML" no editor usa a técnica padrão de navegador (Blob URL +
        # clique num <a download>) pra disparar o download. Isso funciona sozinho num
        # navegador normal, mas dentro de um QWebEngineView (Chromium embutido) o Qt NÃO
        # trata downloads por padrão — sem conectar downloadRequested, o clique não faz
        # nada (nem abre diálogo, nem salva arquivo). Esse handler resolve isso.
        self.web_editor.page().profile().downloadRequested.connect(self.exportar_html_download)

        index_path = get_resource_path(os.path.join("assets", "editor", "index.html"))

        if os.path.exists(index_path):
            self.web_editor.setUrl(QUrl.fromLocalFile(index_path))
        else:
            self.web_editor.setHtml(f"<h3>Erro: Arquivo assets/editor/index.html não encontrado no caminho {index_path}!</h3>")

        layout.addWidget(self.web_editor)
        self.tab_editor.setLayout(layout)

    def exportar_html_download(self, download):
        """Chamado quando o QWebEngineView pede pra salvar um arquivo (ex: clique em
        'Exportar HTML' no editor, que gera um Blob e simula um <a download>). Sem isso,
        o Qt ignora silenciosamente o pedido de download."""
        sugestao = download.suggestedFileName() or "template_email.html"
        caminho, _ = QFileDialog.getSaveFileName(
            self, "Exportar HTML", sugestao, "Arquivos HTML (*.html)"
        )
        if caminho:
            pasta = os.path.dirname(caminho)
            nome = os.path.basename(caminho)
            download.setDownloadDirectory(pasta)
            download.setDownloadFileName(nome)
            download.accept()
        else:
            download.cancel()

    def template_recebido(self, caminho):
        self.txt_template.setText(caminho)
        self.tabs.setCurrentWidget(self.tab_dashboard)
        QMessageBox.information(
            self, "Template salvo",
            f"Template salvo em:\n{caminho}\n\nJá foi selecionado como template do disparo."
        )

    def atualizar_contador_contatos(self):
        contatos = self.obter_lista_contatos()
        total = len(contatos)
        invalidos = sum(1 for c in contatos if c["Email"] and not re.match(EMAIL_REGEX, c["Email"]))
        if invalidos:
            self.lbl_contador.setText(f"{total} contato(s) pronto(s) — ⚠ {invalidos} com e-mail inválido")
            self.lbl_contador.setStyleSheet("color: #c0392b; font-weight: bold;")
        else:
            self.lbl_contador.setText(f"{total} contato(s) pronto(s)")
            self.lbl_contador.setStyleSheet("")
        self.marcar_emails_invalidos(contatos)

    def marcar_emails_invalidos(self, contatos=None):
        """Pinta de vermelho claro a célula de e-mail que não bate com o formato
        esperado, pra avisar o usuário antes de disparar a campanha em vez de só
        descobrir o problema no meio do envio."""
        if contatos is None:
            contatos = self.obter_lista_contatos()
        linhas_com_conteudo = {c["row"] for c in contatos}
        self.tabela.blockSignals(True)
        try:
            for c in contatos:
                item = self.tabela.item(c["row"], 1)
                if not item:
                    continue
                if c["Email"] and not re.match(EMAIL_REGEX, c["Email"]):
                    item.setBackground(QColor("#f8d7da"))
                else:
                    item.setBackground(QColor("#ffffff"))
            # Limpa a cor de linhas que não têm mais conteúdo (ex: usuário apagou tudo)
            for row in range(self.tabela.rowCount()):
                if row not in linhas_com_conteudo:
                    item = self.tabela.item(row, 1)
                    if item:
                        item.setBackground(QColor("#ffffff"))
        finally:
            self.tabela.blockSignals(False)

    def colar_do_excel(self):
        try:
            df = pd.read_clipboard(header=None)
            df.dropna(how='all', inplace=True)
            df = df.fillna("")

            current_row = self.tabela.currentRow()
            current_col = self.tabela.currentColumn()
            if current_row < 0: current_row = 0
            if current_col < 0: current_col = 0

            self.tabela.blockSignals(True)
            for r_idx, row in df.iterrows():
                target_row = current_row + r_idx
                if target_row >= self.tabela.rowCount():
                    self.tabela.setRowCount(target_row + 5)

                for c_idx, val in enumerate(row):
                    target_col = current_col + c_idx
                    if target_col < self.col_status():
                        self.tabela.setItem(target_row, target_col, QTableWidgetItem(str(val)))
            self.tabela.blockSignals(False)
            self.atualizar_contador_contatos()
        except Exception as e:
            self.tabela.blockSignals(False)
            QMessageBox.critical(self, "Erro", f"Não foi possível colar os dados:\n{e}")

    def apagar_selecionado(self):
        for item in self.tabela.selectedItems():
            if item.column() < self.col_status():
                item.setText("")
        self.atualizar_contador_contatos()

    def limpar_tabela(self):
        confirm = QMessageBox.question(self, "Confirmar Limpeza", "Deseja realmente apagar todos os dados da tabela?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm == QMessageBox.StandardButton.Yes:
            self.tabela.clearContents()
            self.tabela.setRowCount(25)
            self.atualizar_contador_contatos()

    def limpar_status(self):
        for row in range(self.tabela.rowCount()):
            item = self.tabela.item(row, self.col_status())
            if item:
                item.setText("")
        self.lbl_status.setText("Status resetados. Pronto para novo envio.")
        self.lbl_status.setStyleSheet("font-size: 10pt; font-weight: bold; background-color: #eef2f7; color: #004687; padding: 12px; border-radius: 6px;")
        self.progress_bar.setValue(0)

    def importar_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecione a Planilha", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                df = pd.read_excel(file_path, header=0)
                df.dropna(how='all', inplace=True)
                df = df.fillna("")

                self.tabela.blockSignals(True)
                self.tabela.setRowCount(len(df) + 10)
                for r_idx, row in df.iterrows():
                    for c_idx in range(min(self.col_status(), len(df.columns))):
                        self.tabela.setItem(r_idx, c_idx, QTableWidgetItem(str(row.iloc[c_idx])))
                self.tabela.blockSignals(False)
                self.atualizar_contador_contatos()
                QMessageBox.information(self, "Sucesso", "Planilha importada com sucesso!")
            except Exception as e:
                self.tabela.blockSignals(False)
                QMessageBox.critical(self, "Erro", f"Erro ao importar arquivo:\n{e}")

    def col_var(self, i: int) -> int:
        """Índice da coluna da variável i (1-based: Var1, Var2, ...)."""
        return 1 + i

    def col_status(self) -> int:
        """Índice da coluna Status — sempre a última, depende de quantas variáveis existem."""
        return 2 + self.num_variaveis

    def atualizar_cabecalho_tabela(self) -> None:
        """(Re)monta os títulos e tooltips das colunas com base em self.num_variaveis.
        Chamado sempre que uma variável é adicionada/removida ou ao carregar a config."""
        labels = ["Nome", "Email"] + [f"Var{i}" for i in range(1, self.num_variaveis + 1)] + ["Status"]
        self.tabela.setHorizontalHeaderLabels(labels)
        self.tabela.horizontalHeaderItem(0).setToolTip("Substitui {{nome}} no template")
        self.tabela.horizontalHeaderItem(1).setToolTip("Endereço de e-mail do destinatário")
        for i in range(1, self.num_variaveis + 1):
            self.tabela.horizontalHeaderItem(self.col_var(i)).setToolTip(f"Substitui {{{{var{i}}}}} no template")
        self.tabela.horizontalHeaderItem(self.col_status()).setToolTip("Preenchido automaticamente durante o disparo")

        nomes_vars = "/".join(f"Var{i}" for i in range(1, self.num_variaveis + 1))
        placeholders_vars = ", ".join(f"{{{{var{i}}}}}" for i in range(1, self.num_variaveis + 1))
        self.lbl_dica_variaveis.setText(
            "Dica: Cole dados diretamente do Excel usando <b>Ctrl+V</b>. Use a tecla <b>Delete</b> para limpar células. "
            f"As colunas <b>{nomes_vars}</b> substituem <b>{placeholders_vars}</b> no template."
        )
        if hasattr(self, "btn_remover_variavel"):
            self.btn_remover_variavel.setToolTip(f"Remove a última coluna de variável (Var{self.num_variaveis})")
            self.btn_remover_variavel.setEnabled(self.num_variaveis > 1)

    def adicionar_coluna_variavel(self) -> None:
        """Adiciona uma nova coluna de variável (Var4, Var5, ...) à tabela de contatos,
        logo antes da coluna Status."""
        LIMITE_MAX_VARIAVEIS = 15
        if self.num_variaveis >= LIMITE_MAX_VARIAVEIS:
            QMessageBox.information(self, "Limite atingido", f"Você já tem {LIMITE_MAX_VARIAVEIS} variáveis — deve ser o suficiente!")
            return
        idx_status = self.col_status()
        self.tabela.insertColumn(idx_status)  # a nova var entra no lugar do Status, empurrando-o pra frente
        self.num_variaveis += 1
        self.atualizar_cabecalho_tabela()
        self.salvar_config()

    def remover_coluna_variavel(self) -> None:
        """Remove a última coluna de variável, avisando antes se ela tiver dados
        preenchidos (pra não apagar algo sem querer)."""
        if self.num_variaveis <= 1:
            QMessageBox.information(self, "Não é possível remover", "É preciso manter pelo menos uma variável.")
            return

        idx_ultima_var = self.col_var(self.num_variaveis)
        tem_dados = any(
            self.tabela.item(row, idx_ultima_var) and self.tabela.item(row, idx_ultima_var).text().strip()
            for row in range(self.tabela.rowCount())
        )
        if tem_dados:
            resp = QMessageBox.question(
                self, "Remover variável",
                f"A coluna Var{self.num_variaveis} tem dados preenchidos. Removê-la vai apagar esses "
                "valores permanentemente. Deseja continuar?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if resp != QMessageBox.StandardButton.Yes:
                return

        self.tabela.removeColumn(idx_ultima_var)
        self.num_variaveis -= 1
        self.atualizar_cabecalho_tabela()
        self.salvar_config()
        self.salvar_contatos_em_disco()

    def ajustar_num_variaveis_para(self, total_desejado: int) -> None:
        """Ajusta silenciosamente (sem confirmação) a tabela pra ter exatamente
        `total_desejado` colunas de variável. Usado ao restaurar a configuração salva
        no início do app, antes de carregar os contatos salvos — nesse momento o
        usuário já tinha escolhido essa quantidade antes, não precisa confirmar de novo."""
        total_desejado = max(1, min(total_desejado, 15))
        while self.num_variaveis < total_desejado:
            self.tabela.insertColumn(self.col_status())
            self.num_variaveis += 1
        while self.num_variaveis > total_desejado:
            self.tabela.removeColumn(self.col_var(self.num_variaveis))
            self.num_variaveis -= 1
        self.atualizar_cabecalho_tabela()

    def obter_lista_contatos(self) -> List[Dict]:  # CORREÇÃO CLEAN CODE: Type hints
        """
        Extrai lista de contatos da tabela.
        
        Returns:
            Lista de dicionários com informações do contato
        """
        dados = []
        for row in range(self.tabela.rowCount()):
            nome_item = self.tabela.item(row, 0)
            email_item = self.tabela.item(row, 1)
            
            nome = nome_item.text().strip() if nome_item else ""
            email = email_item.text().strip() if email_item else ""

            if nome or email:
                registro = {
                    "row": row,
                    "Nome": nome,
                    "Email": email,
                    "Status": self.tabela.item(row, self.col_status()).text().strip() if self.tabela.item(row, self.col_status()) else ""
                }
                for i in range(1, self.num_variaveis + 1):
                    item = self.tabela.item(row, self.col_var(i))
                    registro[f"Var{i}"] = item.text().strip() if item else ""
                dados.append(registro)
        return dados

    def selecionar_template(self) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Abre diálogo para selecionar arquivo de template HTML."""
        file_path, _ = QFileDialog.getOpenFileName(self, "Escolha o Template HTML", "", "Arquivos HTML (*.html *.htm)")
        if file_path:
            self.txt_template.setText(file_path)
            self.salvar_config()

    def abrir_lista_templates_salvos(self) -> None:
        """Mostra um menu com os templates HTML já salvos em templates_salvos/, pra
        reabrir um deles sem precisar navegar manualmente pelas pastas toda vez."""
        try:
            arquivos = [
                f for f in os.listdir(self.pasta_templates)
                if f.lower().endswith((".html", ".htm"))
            ]
        except Exception as e:
            logger.warning(f"Falha ao listar templates salvos: {e}")
            arquivos = []

        if not arquivos:
            QMessageBox.information(self, "Templates Salvos", "Nenhum template salvo ainda em:\n" + self.pasta_templates)
            return

        # Mais recentes primeiro
        arquivos.sort(key=lambda f: os.path.getmtime(os.path.join(self.pasta_templates, f)), reverse=True)

        menu = QMenu(self)
        for nome_arquivo in arquivos:
            caminho_completo = os.path.join(self.pasta_templates, nome_arquivo)
            data_mod = time.strftime("%d/%m/%Y %H:%M", time.localtime(os.path.getmtime(caminho_completo)))
            acao = menu.addAction(f"{nome_arquivo}   ({data_mod})")
            acao.triggered.connect(lambda checked=False, c=caminho_completo: self._selecionar_template_salvo(c))
        menu.exec(self.btn_templates_salvos.mapToGlobal(self.btn_templates_salvos.rect().bottomLeft()))

    def _selecionar_template_salvo(self, caminho: str) -> None:
        self.txt_template.setText(caminho)
        self.salvar_config()

    def selecionar_anexo(self) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Abre diálogo para selecionar um ou mais arquivos de anexo. Os novos arquivos
        são somados aos já escolhidos (sem duplicar), pra dar pra montar uma lista aos
        poucos em vez de só substituir o anexo anterior."""
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Escolha o(s) Anexo(s)", "", "Todos os Arquivos (*.*)")
        if file_paths:
            for fp in file_paths:
                if fp not in self.lista_anexos:
                    self.lista_anexos.append(fp)
            self._atualizar_texto_anexos()
            self.salvar_config()

    def limpar_anexos(self) -> None:
        self.lista_anexos = []
        self._atualizar_texto_anexos()
        self.salvar_config()

    def _atualizar_texto_anexos(self) -> None:
        if not self.lista_anexos:
            self.txt_anexo.setText("")
            return
        nomes = [os.path.basename(p) for p in self.lista_anexos]
        if len(nomes) <= 3:
            self.txt_anexo.setText(", ".join(nomes))
        else:
            self.txt_anexo.setText(f"{len(nomes)} arquivos: " + ", ".join(nomes[:3]) + ", ...")

    def previsualizar_html(self):
        path_template = self.txt_template.text().strip()
        if not path_template or not os.path.exists(path_template):
            QMessageBox.warning(self, "Atenção", "Selecione um arquivo de template HTML válido antes de visualizar!")
            return

        # OTIMIZAÇÃO #4: Usar função reutilizável em vez de duplicar código de leitura
        corpo_base = ler_arquivo_template(path_template)
        if corpo_base is None:
            QMessageBox.warning(self, "Atenção", "Erro ao ler o template. Verifique o encoding ou tamanho do arquivo.")
            return

        contatos = self.obter_lista_contatos()
        contato_escolhido = None
        if contatos:
            # Se o usuário tinha selecionado uma linha específica na aba Contatos, usa
            # os dados DAQUELA linha na prévia em vez de sempre pegar o primeiro contato
            # — assim dá pra conferir como o e-mail fica pra um destinatário específico.
            linha_selecionada = self.tabela.currentRow()
            for c in contatos:
                if c["row"] == linha_selecionada:
                    contato_escolhido = c
                    break
            if contato_escolhido is None:
                contato_escolhido = contatos[0]

        if contato_escolhido:
            nome = contato_escolhido["Nome"] or "Aluno (Exemplo)"
        else:
            nome = "Aluno (Exemplo)"
            contato_escolhido = {f"Var{i}": "" for i in range(1, self.num_variaveis + 1)}

        # Escapa os valores antes de inserir no HTML (mesmo tratamento aplicado no envio
        # de verdade) — assim a prévia mostra fielmente o que vai ser enviado, inclusive
        # quando um contato tem "&", "<" ou ">" no nome/variáveis. Funciona pra quantas
        # variáveis existirem (Var1..VarN), não só as três primeiras.
        variaveis = montar_variaveis_substituicao(nome, contato_escolhido, escapar=True)
        corpo_final = substituir_variaveis(corpo_base, variaveis)

        temp_path = os.path.join(tempfile.gettempdir(), "preview_r9bot_temp.html")
        try:
            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(corpo_final)
            subprocess.run(['start', temp_path], shell=True)
        except Exception as e:
            logger.error(f"Erro ao abrir pré-visualização: {e}")
            QMessageBox.critical(self, "Erro", f"Erro ao abrir pré-visualização: {e}")


    def alternar_bloqueio_interface(self, bloquear: bool) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Ativa ou desativa os controles da interface durante o envio."""
        self.form_container.setEnabled(not bloquear)
        self.btn_colar.setEnabled(not bloquear)
        self.btn_importar.setEnabled(not bloquear)
        self.btn_limpar_tabela.setEnabled(not bloquear)
        self.btn_limpar_status.setEnabled(not bloquear)
        self.btn_enviar.setEnabled(not bloquear)
        self.btn_teste.setEnabled(not bloquear)

    def executar_campanha(self):
        if self.executando:
            QMessageBox.warning(self, "Aviso", "Já existe uma campanha em execução.")
            return

        path_template = self.txt_template.text().strip()
        if not path_template or not os.path.exists(path_template):
            QMessageBox.warning(self, "Atenção", "Selecione um template HTML válido.")
            return

        assunto = self.txt_assunto.text().strip()
        if not assunto:
            QMessageBox.warning(self, "Atenção", "Preencha o Assunto do E-mail antes de disparar a campanha.")
            return

        contatos = self.obter_lista_contatos()
        if not contatos:
            QMessageBox.warning(self, "Atenção", "Nenhum contato cadastrado na aba 'Contatos'.")
            return

        # Detecta e-mails duplicados na lista antes de disparar, pra não mandar 2x pro
        # mesmo destinatário sem o usuário perceber. Se houver duplicados, mantém só a
        # primeira ocorrência de cada e-mail e avisa quais foram ignorados.
        vistos = set()
        contatos_unicos = []
        duplicados = []
        for c in contatos:
            chave = c["Email"].strip().lower()
            if chave and chave in vistos:
                duplicados.append(c["Email"])
                continue
            vistos.add(chave)
            contatos_unicos.append(c)

        if duplicados:
            amostra = "\n".join(f"• {e}" for e in duplicados[:10])
            a_mais = f"\n... e mais {len(duplicados) - 10}" if len(duplicados) > 10 else ""
            resp = QMessageBox.question(
                self, "E-mails duplicados encontrados",
                f"{len(duplicados)} e-mail(s) duplicado(s) foram encontrados na lista:\n\n{amostra}{a_mais}\n\n"
                "Deseja continuar enviando só uma vez para cada e-mail (recomendado)?\n"
                "Escolha 'Não' para cancelar e revisar a lista manualmente.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if resp != QMessageBox.StandardButton.Yes:
                return
            contatos = contatos_unicos

        confirm = QMessageBox.question(self, "Confirmar Disparo", f"Deseja iniciar o disparo para {len(contatos)} contatos via Outlook?",
                                       QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if confirm != QMessageBox.StandardButton.Yes:
            return

        # OTIMIZAÇÃO #4: Usar função reutilizável em vez de duplicar código de leitura
        corpo_base = ler_arquivo_template(path_template)
        if corpo_base is None:
            QMessageBox.critical(self, "Erro Template", f"Não foi possível ler o template.\nVerifique o encoding ou o tamanho do arquivo.")
            return

        # Imagens que o usuário enviou pelo editor via upload local ficam embutidas no
        # HTML como data URI (base64). O Outlook não é confiável renderizando isso
        # direto no HTMLBody, então convertemos pra anexos inline com Content-ID aqui,
        # UMA vez antes do loop de envio (o mesmo arquivo temporário é reaproveitado
        # em todos os e-mails da campanha).
        corpo_base, anexos_inline = preparar_imagens_inline(corpo_base)

        anexos = list(self.lista_anexos)
        cc = self.txt_cc.text().strip()
        bcc = self.txt_bcc.text().strip()
        
        # CLEAN CODE: Usar valor constante como fallback e validar entrada
        try:
            intervalo = int(self.txt_intervalo.text().strip() or DEFAULT_SEND_INTERVAL)
            if intervalo < 0:
                intervalo = DEFAULT_SEND_INTERVAL
        except ValueError:
            logger.warning(f"Intervalo inválido, usando padrão {DEFAULT_SEND_INTERVAL}s")
            intervalo = DEFAULT_SEND_INTERVAL

        self.executando = True
        self.progress_bar.setValue(0)
        self.alternar_bloqueio_interface(True)
        
        self.thread_envio = threading.Thread(target=self.processar_envio_background, args=(contatos, corpo_base, assunto, anexos, intervalo, anexos_inline, cc, bcc))
        self.thread_envio.daemon = True  # CORREÇÃO: Permitir que app feche mesmo com thread ativa
        self.thread_envio.start()

    def enviar_email_teste(self):
        """Envia UM único e-mail de teste (pro próprio usuário verificar antes de disparar
        pra lista toda), usando os dados do primeiro contato da tabela pra preencher as
        variáveis {{nome}}, {{var1}}, etc. Não altera status nem toca no log da campanha."""
        if self.executando:
            QMessageBox.warning(self, "Aviso", "Aguarde a campanha atual terminar antes de enviar um teste.")
            return

        path_template = self.txt_template.text().strip()
        if not path_template or not os.path.exists(path_template):
            QMessageBox.warning(self, "Atenção", "Selecione um template HTML válido.")
            return

        assunto = self.txt_assunto.text().strip()
        if not assunto:
            QMessageBox.warning(self, "Atenção", "Preencha o Assunto do E-mail antes de enviar um teste.")
            return

        destino, ok = QInputDialog.getText(
            self, "Enviar Teste", "Enviar e-mail de teste para qual endereço?"
        )
        if not ok or not destino.strip():
            return
        destino = destino.strip()
        if not re.match(EMAIL_REGEX, destino):
            QMessageBox.warning(self, "Atenção", "Esse endereço de e-mail não parece válido.")
            return

        corpo_base = ler_arquivo_template(path_template)
        if corpo_base is None:
            QMessageBox.critical(self, "Erro Template", "Não foi possível ler o template.\nVerifique o encoding ou o tamanho do arquivo.")
            return

        corpo_base, anexos_inline = preparar_imagens_inline(corpo_base)

        # Usa os dados do primeiro contato cadastrado (se houver) só pra dar um preview
        # realista das variáveis; se a lista estiver vazia, envia com os placeholders em branco.
        contatos = self.obter_lista_contatos()
        primeiro = contatos[0] if contatos else {"Nome": "", **{f"Var{i}": "" for i in range(1, self.num_variaveis + 1)}}

        anexos = list(self.lista_anexos)

        self.btn_teste.setEnabled(False)
        self.btn_teste.setText("Enviando teste...")

        thread = threading.Thread(
            target=self._enviar_teste_background,
            args=(corpo_base, assunto, anexos, destino, primeiro, anexos_inline)
        )
        thread.daemon = True
        thread.start()

    def _enviar_teste_background(self, corpo_base: str, assunto: str, anexos: Optional[List[str]], destino: str, contato_referencia: Dict, anexos_inline: Optional[List[Dict]] = None, cc: str = "", bcc: str = "") -> None:
        """Roda em thread separada pra não travar a UI enquanto o Outlook processa o
        envio de teste (a automação COM pode levar um instante)."""
        import win32com.client as win32
        anexos_inline = anexos_inline or []
        anexos = anexos or []

        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except NameError:
            pass
        except Exception as e:
            logger.warning(f"Erro ao inicializar COM (teste): {e}")

        try:
            try:
                outlook = win32.Dispatch("Outlook.Application")
            except Exception as e:
                self.signals.teste_enviado.emit(False, f"Não foi possível conectar ao Outlook local.\nErro: {e}")
                return

            nome_ref = contato_referencia.get("Nome", "") or ""
            variaveis_html = montar_variaveis_substituicao(nome_ref, contato_referencia, escapar=True)
            variaveis_assunto = montar_variaveis_substituicao(nome_ref, contato_referencia, escapar=False)

            corpo_final = substituir_variaveis(corpo_base, variaveis_html)
            assunto_final = substituir_variaveis(assunto, variaveis_assunto)

            mail = None
            try:
                mail = outlook.CreateItem(0)
                mail.To = destino
                mail.Subject = f"[TESTE] {assunto_final}"
                if anexos_inline:
                    anexar_imagens_inline(mail, anexos_inline)
                mail.HTMLBody = corpo_final
                for caminho_anexo in anexos:
                    if caminho_anexo and os.path.exists(caminho_anexo):
                        mail.Attachments.Add(caminho_anexo)
                mail.Send()
                self.signals.teste_enviado.emit(True, destino)
            except Exception as e:
                self.signals.teste_enviado.emit(False, f"{type(e).__name__}: {e}")
            finally:
                if mail is not None:
                    mail = None
        finally:
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            if anexos_inline:
                limpar_arquivos_temporarios(anexos_inline)

    def resultado_envio_teste(self, sucesso: bool, mensagem: str) -> None:
        """Chamado de volta na thread principal com o resultado do envio de teste."""
        self.btn_teste.setEnabled(True)
        self.btn_teste.setText("🧪 Enviar Teste")
        if sucesso:
            QMessageBox.information(self, "Teste Enviado", f"E-mail de teste enviado com sucesso para:\n{mensagem}")
        else:
            QMessageBox.critical(self, "Falha no Teste", f"Não foi possível enviar o e-mail de teste.\n\n{mensagem}")


    def processar_envio_background(self, contatos: List[Dict], corpo_base: str, assunto: str, anexos: Optional[List[str]], intervalo: int, anexos_inline: Optional[List[Dict]] = None, cc: str = "", bcc: str = "") -> None:
        """
        CORREÇÃO: Type hints adicionados (CLEAN CODE)
        Processa o envio de emails em thread separada de forma thread-safe.
        """
        import win32com.client as win32
        anexos_inline = anexos_inline or []
        anexos = anexos or []
        
        # CORREÇÃO BUG #3: Inicialização COM mais robusta com tratamento de erro
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            logger.info("COM inicializado com sucesso na thread de envio")
        except NameError:
            # pythoncom não disponível (OK em algumas instalações)
            logger.info("pythoncom não disponível, continuando sem inicialização COM")
        except Exception as e:
            logger.warning(f"Erro ao inicializar COM: {e}, continuando mesmo assim")

        try:
            try:
                outlook = win32.Dispatch("Outlook.Application")
                logger.info("Conexão com Outlook estabelecida")
            except Exception as e:
                logger.error(f"Não foi possível conectar ao Outlook: {e}")
                self.signals.error_emitted.emit("Erro Outlook", f"Não foi possível conectar ao Outlook local.\nErro: {e}")
                with self._lock_executando:  # CORREÇÃO BUG #1: Usar lock
                    self.executando = False
                self.signals.finished.emit(0, 0)
                return

            enviados = 0
            erros = 0
            processados = 0
            pendentes = [c for c in contatos if c["Status"] != "ENVIADO"]
            total_enviar = len(pendentes)
            # Lista compartilhada (protegida por lock) em vez de variável local: assim,
            # se o app for fechado no meio da campanha, o closeEvent consegue salvar um
            # log parcial de tudo que já foi enviado até aquele momento.
            with self._lock_log:
                self._linhas_log_atual = []
            MAX_TENTATIVAS = 3  # 1 tentativa original + 2 retentativas em falha transitória

            for idx, c in enumerate(pendentes):
                # CORREÇÃO BUG #1: Usar lock para acessar self.executando de forma segura
                with self._lock_executando:
                    should_stop = not self.executando
                
                if should_stop:
                    logger.info("Envio pausado pelo usuário")
                    break

                row_idx = c["row"]
                nome = c["Nome"]
                email_dest = c["Email"]
                
                processados += 1
                novo_status = ""
                detalhe_erro = ""
                
                # CORREÇÃO BUG #4: Validação de email ROBUSTA com regex em vez de checagem fraca
                if re.match(EMAIL_REGEX, email_dest):
                    try:
                        # CORREÇÃO BUG #6: Substituição SEGURA de variáveis com dicionário
                        # Evita injection quando valores contêm "{{...}}"
                        # Funciona pra Var1..VarN (quantas variáveis o usuário tiver
                        # configurado), não só as três primeiras.
                        #
                        # Os valores são escapados (html.escape) antes de entrar no HTML do
                        # corpo — sem isso, um contato com "&", "<" ou ">" no nome/variável
                        # (ex: "Marcos & Filhos Ltda") quebra a estrutura do HTML, o que é
                        # especialmente arriscado no motor do Word do Outlook.
                        variaveis_html = montar_variaveis_substituicao(nome, c, escapar=True)
                        # Para o Assunto, usa os valores SEM escape (é texto puro, não HTML —
                        # escapar aqui faria aparecer "&amp;" literalmente no assunto).
                        variaveis_assunto = montar_variaveis_substituicao(nome, c, escapar=False)
                        
                        # Substituir todas as variáveis em uma única passagem usando regex
                        # (agora também tolera espaços dentro das chaves, ex: "{{ nome }}")
                        corpo_final = substituir_variaveis(corpo_base, variaveis_html)
                        # Permite personalizar o Assunto também, ex: "{{nome}}, sua matrícula
                        # está quase completa" — antes o assunto era sempre um texto fixo.
                        assunto_final = substituir_variaveis(assunto, variaveis_assunto)

                        # Retry em falha transitória: erros de COM (Outlook ocupado, rede
                        # instável, etc.) às vezes se resolvem sozinhos numa nova tentativa
                        # alguns segundos depois. Sem isso, qualquer soluço momentâneo do
                        # Outlook marcava o contato como ERRO definitivo sem tentar de novo.
                        ultimo_erro = None
                        enviado_com_sucesso = False
                        for tentativa in range(1, MAX_TENTATIVAS + 1):
                            # CORREÇÃO BUG #2: Garantir cleanup de objeto mail
                            mail = None
                            try:
                                mail = outlook.CreateItem(0)
                                mail.To = email_dest
                                if cc:
                                    mail.CC = cc
                                if bcc:
                                    mail.BCC = bcc
                                mail.Subject = assunto_final
                                if anexos_inline:
                                    anexar_imagens_inline(mail, anexos_inline)
                                mail.HTMLBody = corpo_final
                                for caminho_anexo in anexos:
                                    if caminho_anexo and os.path.exists(caminho_anexo):
                                        mail.Attachments.Add(caminho_anexo)
                                mail.Send()

                                enviado_com_sucesso = True
                                logger.debug(f"Email enviado para {email_dest} (tentativa {tentativa})")
                                break
                            except Exception as e:
                                ultimo_erro = e
                                logger.warning(
                                    f"Falha ao enviar para {email_dest} na tentativa {tentativa}/{MAX_TENTATIVAS}: "
                                    f"{type(e).__name__}: {e}"
                                )
                                if tentativa < MAX_TENTATIVAS:
                                    time.sleep(2)
                            finally:
                                # CORREÇÃO BUG #2: Liberar referência COM explicitamente
                                if mail is not None:
                                    try:
                                        mail = None  # Permite garbage collection
                                    except Exception:
                                        pass

                        if enviado_com_sucesso:
                            novo_status = "ENVIADO"
                            enviados += 1
                        else:
                            novo_status = "ERRO"
                            detalhe_erro = f"{type(ultimo_erro).__name__}: {ultimo_erro}" if ultimo_erro else "Falha desconhecida"
                            erros += 1
                            logger.error(f"Erro ao enviar para {email_dest} após {MAX_TENTATIVAS} tentativas: {detalhe_erro}")
                    except Exception as e:
                        novo_status = "ERRO"
                        detalhe_erro = f"{type(e).__name__}: {e}"
                        erros += 1
                        logger.error(f"Erro ao enviar para {email_dest}: {type(e).__name__}: {e}")
                else:
                    novo_status = "E-mail inválido"
                    detalhe_erro = "Formato de e-mail inválido"
                    erros += 1
                    logger.warning(f"Email inválido rejeitado: {email_dest}")

                with self._lock_log:
                    self._linhas_log_atual.append((
                        time.strftime("%Y-%m-%d %H:%M:%S"), nome, email_dest, novo_status, detalhe_erro
                    ))

                self.signals.row_updated.emit(row_idx, novo_status)

                restantes = total_enviar - processados
                eta_segundos = restantes * intervalo
                eta_txt = formatar_duracao(eta_segundos) if restantes > 0 else "concluindo..."

                status_txt = (
                    f"Enviando... ({processados}/{total_enviar}) | Sucesso: {enviados} | "
                    f"Erros: {erros} | Tempo restante estimado: {eta_txt}"
                )
                percentual = int((processados / total_enviar) * 100) if total_enviar > 0 else 0
                
                self.signals.status_updated.emit(status_txt, "#d1ecf1", enviados, erros)
                self.signals.progress_updated.emit(percentual, total_enviar)

                time.sleep(intervalo)

            # Salva um CSV com o resultado detalhado da campanha (quem recebeu, quem deu
            # erro e por quê) — sem isso, todo esse histórico desaparecia ao fechar o app.
            with self._lock_log:
                caminho_log = self.salvar_log_envio(list(self._linhas_log_atual))
            self.ultimo_log_path = caminho_log

            # CORREÇÃO BUG #1: Usar lock ao finalizar
            with self._lock_executando:
                self.executando = False
            
            self.signals.finished.emit(enviados, erros)
            logger.info(f"Campanha finalizada: {enviados} enviados, {erros} erros. Log: {caminho_log}")
            
        finally:
            # CORREÇÃO BUG #3: Finalizações COM apenas se foi inicializado com sucesso
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    logger.info("COM finalizado com sucesso")
                except Exception as e:
                    logger.warning(f"Erro ao finalizar COM: {e}")
            # Remove os arquivos temporários das imagens embutidas (upload local) — não
            # são mais necessários depois que a campanha termina de enviar.
            if anexos_inline:
                limpar_arquivos_temporarios(anexos_inline)

    def salvar_log_envio(self, linhas_log) -> Optional[str]:
        """Grava um CSV com o resultado detalhado da campanha (timestamp, nome, email,
        status, detalhe do erro) em logs_envio/. Retorna o caminho do arquivo, ou None
        se não houver nada pra salvar ou se a gravação falhar."""
        if not linhas_log:
            return None
        try:
            import csv
            nome_arquivo = f"campanha_{time.strftime('%Y%m%d_%H%M%S')}.csv"
            caminho = os.path.join(self.pasta_logs, nome_arquivo)
            with open(caminho, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["Data/Hora", "Nome", "Email", "Status", "Detalhe do Erro"])
                writer.writerows(linhas_log)
            return caminho
        except Exception as e:
            logger.warning(f"Falha ao salvar log de envio: {e}")
            return None


    def mostrar_erro_thread(self, titulo: str, mensagem: str) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Exibe diálogo de erro emitido de thread."""
        QMessageBox.critical(self, titulo, mensagem)

    def atualizar_status_tabela(self, row_idx: int, status: str) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Atualiza coluna de status de uma linha da tabela."""
        self.tabela.setItem(row_idx, self.col_status(), QTableWidgetItem(status))

    def atualizar_status_ui(self, text: str, color: str, enviados: int, erros: int) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Atualiza label de status com cor e contador de emails."""
        self.lbl_status.setText(text)
        self.lbl_status.setStyleSheet(f"font-size: 10pt; font-weight: bold; background-color: {color}; color: #0c5460; padding: 12px; border-radius: 6px;")

    def atualizar_progresso(self, val: int, total: int) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Atualiza barra de progresso."""
        self.progress_bar.setValue(val)

    def fim_disparo(self, enviados: int, erros: int) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Finaliza a campanha de envio e exibe resultado."""
        self.alternar_bloqueio_interface(False)
        self.lbl_status.setText(f"Campanha Concluída! Total enviados: {enviados} | Erros: {erros}")
        self.lbl_status.setStyleSheet("font-size: 10pt; font-weight: bold; background-color: #d4edda; color: #155724; padding: 12px; border-radius: 6px;")
        self.progress_bar.setValue(100)
        log_msg = ""
        if getattr(self, "ultimo_log_path", None):
            log_msg = f"\n\nLog detalhado salvo em:\n{self.ultimo_log_path}"
        QMessageBox.information(self, "Fim do Disparo", f"Campanha finalizada com sucesso!\n\nE-mails enviados: {enviados}\nErros/Inválidos: {erros}{log_msg}")
        self.atualizar_historico_campanhas()

    def pausar_campanha(self) -> None:  # CORREÇÃO CLEAN CODE: Type hints
        """Para a campanha de envio em progresso."""
        if self.executando:
            # CORREÇÃO BUG #1: Usar lock para acessar e modificar self.executando de forma segura
            with self._lock_executando:
                self.executando = False
            
            self.alternar_bloqueio_interface(False)
            self.lbl_status.setText("Disparo pausado pelo usuário.")
            self.lbl_status.setStyleSheet("font-size: 10pt; font-weight: bold; background-color: #fff3cd; color: #856404; padding: 12px; border-radius: 6px;")
            logger.info("Campanha pausada pelo usuário")  # CORREÇÃO CLEAN CODE: Log em vez de deixar silencioso
            QMessageBox.warning(self, "Pausado", "O processo de envio foi interrompido com segurança!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = R9BotQtApp()
    window.show()
    app.exec()