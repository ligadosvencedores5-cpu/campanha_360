# ════════════════════════════════════════════════════════════════════════
# CAMPANHA 360 IA — MeusAgentesIA.com.br  (arquivo único)
#
# Do briefing ao relatório: estratégia, funil de vendas, posts com arte
# pronta, mensagens e links de WhatsApp, acompanhamento de cada contato
# e relatório em texto e Excel.
#
# COMO PUBLICAR
#   Coloque SÓ este arquivo num repositório do GitHub e publique no
#   Streamlit Cloud. Não precisa de requirements.txt: usa apenas o que já
#   vem com o Streamlit (requests, pandas, Pillow).
#   Cada pessoa entra com o próprio nome e a própria chave Groq.
# ════════════════════════════════════════════════════════════════════════
import base64
import io
import json
import re
import urllib.parse
import zipfile
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

st.set_page_config(page_title="Campanha 360 IA", page_icon="🚀", layout="wide")

MODELO   = "openai/gpt-oss-120b"
GROQ_URL = "https://api.groq.com/openai/v1/chat/completions"
SITE     = "https://meusagentesia.com.br"

STATUS = ["Novo", "Enviado", "Respondeu", "Interessado", "Comprou", "Não quer"]

REGRA_IA = (
    "Você é um estrategista de marketing digital e copywriter sênior brasileiro, especialista em "
    "funis de vendas, redes sociais e WhatsApp. Escreva em português do Brasil, com linguagem clara, "
    "persuasiva e ética: nada de promessas impossíveis, falsa escassez ou informação inventada. "
    "Use apenas os dados do briefing; quando algo não foi informado, sugira de forma genérica e marque "
    "como [AJUSTAR]. Não invente preços, números de clientes, depoimentos ou estatísticas."
)

# ══════════════════════════════════════════════════════════════════════════
# ESTILO
# ══════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
#MainMenu, footer, header, [data-testid="stToolbar"], [data-testid="stDecoration"] { visibility:hidden; }
.stApp { background:#F8F9FA; }
.block-container { padding-top:1.2rem; }
h1, h2, h3, h4, p, li, label, span, div { color:#1A1A2E; }
.cab { background:linear-gradient(135deg,#16A34A,#0EA5E9); border-radius:16px; padding:18px 22px; margin-bottom:10px; }
.cab h1 { margin:0; font-size:1.6em; }
.cab p  { margin:4px 0 0 0; }
.cab, .cab * { color:#FFFFFF !important; }
.card { background:#fff; border:1px solid #E5E7EB; border-radius:14px; padding:16px 18px; margin-bottom:12px; }
.passo { display:inline-block; background:#DCFCE7; color:#166534 !important; border-radius:20px;
         padding:2px 12px; font-size:0.8em; font-weight:700; margin-bottom:6px; }
.metrica { background:#fff; border:1px solid #E5E7EB; border-radius:14px; padding:14px; text-align:center; }
.metrica b { font-size:1.7em; display:block; color:#16A34A !important; }
.aviso { background:#FFFBEB; border:1px solid #FDE68A; border-radius:12px; padding:12px 16px;
         font-size:0.9em; line-height:1.6; }
.aviso, .aviso * { color:#78350F !important; }
.stButton > button[kind="primary"] { background:#16A34A !important; border-color:#16A34A !important; color:#fff !important; }
.stButton > button[kind="primary"] *, [data-testid="stFormSubmitButton"] button, [data-testid="stFormSubmitButton"] button * { color:#fff !important; }
[data-testid="stFormSubmitButton"] button { background:#16A34A !important; border-color:#16A34A !important; }
[data-testid="stFormSubmitButton"] button:hover { background:#15803D !important; }
[data-baseweb="tag"], [data-testid="stMultiSelect"] [data-tag] { background:#16A34A !important; }
[data-baseweb="tag"] *, [data-testid="stMultiSelect"] [data-tag], [data-testid="stMultiSelect"] [data-tag] * { color:#fff !important; }
[data-baseweb="tab-highlight"], .stTabs .react-aria-SelectionIndicator { background:#16A34A !important; }
.stTabs [data-testid="stTab"][aria-selected="true"] { border-bottom-color:#16A34A !important; }
.stTabs [aria-selected="true"] p { color:#16A34A !important; }
.stButton > button[kind="primary"]:hover { background:#15803D !important; border-color:#15803D !important; }
.stButton > button, [data-testid="stDownloadButton"] > button { border-radius:10px; font-weight:600; }
[data-testid="stTextInputRootElement"]:focus-within, [data-baseweb="textarea"]:focus-within,
[data-baseweb="select"] > div:focus-within { border-color:#16A34A !important; }
.stTabs [data-baseweb="tab"] { font-weight:600; }
.rodape { text-align:center; color:#9CA3AF !important; font-size:0.78em; margin-top:30px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════
# ESTADO
# ══════════════════════════════════════════════════════════════════════════
PADRAO = {
    "c_nome_user": "", "c_chave": "",
    "briefing": {}, "estrategia": "", "funil": "",
    "posts": [], "mensagens": [], "contatos": [],
    "marca_cor1": "#16A34A", "marca_cor2": "#0F172A", "marca_logo": "",
    "proxima_msg": {}, "analise_final": "",
}
for _k, _v in PADRAO.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v.copy() if isinstance(_v, (dict, list)) else _v
S = st.session_state


# ══════════════════════════════════════════════════════════════════════════
# IA (Groq via HTTP — sem biblioteca extra)
# ══════════════════════════════════════════════════════════════════════════
class ErroIA(Exception):
    pass


def ia(prompt: str, json_mode: bool = False, temperatura: float = 0.7) -> str:
    msgs = [{"role": "system", "content": REGRA_IA}, {"role": "user", "content": prompt}]
    corpo = {"model": MODELO, "messages": msgs, "temperature": temperatura}
    if json_mode:
        corpo["response_format"] = {"type": "json_object"}
    try:
        r = requests.post(GROQ_URL, json=corpo, timeout=180,
                          headers={"Authorization": f"Bearer {S.c_chave}", "Content-Type": "application/json"})
    except requests.RequestException as e:
        raise ErroIA(f"Sem conexão com a IA agora. Tente de novo. ({e.__class__.__name__})")
    if r.status_code == 401:
        raise ErroIA("🔑 Sua chave Groq não foi aceita. Clique em Sair e entre com uma chave válida.")
    if r.status_code == 429:
        raise ErroIA("⏳ Limite de uso da sua chave Groq atingido. Espere um pouco e tente de novo.")
    if r.status_code == 400 and json_mode:
        return ia(prompt + "\n\nResponda SOMENTE com JSON válido.", json_mode=False, temperatura=temperatura)
    if r.status_code != 200:
        raise ErroIA(f"A IA não respondeu (código {r.status_code}). Tente de novo em instantes.")
    try:
        txt = r.json()["choices"][0]["message"]["content"] or ""
    except (ValueError, KeyError, IndexError):
        raise ErroIA("A IA devolveu uma resposta inesperada. Tente de novo.")
    return re.sub(r"<think>.*?</think>", "", txt, flags=re.S).strip()


def ia_json(prompt: str) -> dict:
    txt = ia(prompt, json_mode=True)
    try:
        return json.loads(txt)
    except ValueError:
        m = re.search(r"\{.*\}", txt, flags=re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except ValueError:
                pass
    raise ErroIA("A IA respondeu fora do formato esperado. Clique de novo para gerar.")


def rodar(fn, *a, **kw):
    """Executa uma chamada de IA mostrando erro amigável."""
    try:
        return fn(*a, **kw)
    except ErroIA as e:
        st.error(str(e))
        return None


def briefing_txt() -> str:
    b = S.briefing
    rot = {
        "negocio": "Negócio / marca", "produto": "Produto ou serviço", "preco": "Preço / condição",
        "publico": "Público-alvo", "diferencial": "Diferencial", "objetivo": "Objetivo da campanha",
        "duracao": "Duração", "canais": "Canais", "tom": "Tom de voz", "oferta": "Oferta / bônus",
        "link": "Link de compra/contato", "cidade": "Cidade / região",
    }
    return "\n".join(f"- {rot.get(k, k)}: {v}" for k, v in b.items() if str(v).strip())


def tem_briefing() -> bool:
    return bool(S.briefing.get("produto") and S.briefing.get("publico"))


# ══════════════════════════════════════════════════════════════════════════
# ARTE DOS POSTS (Pillow)
# ══════════════════════════════════════════════════════════════════════════
_FONTES_NEGRITO = [
    "DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", "Arial Bold.ttf", "arialbd.ttf",
]
_FONTES_NORMAL = [
    "DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "Arial.ttf", "arial.ttf",
]


@st.cache_resource
def fonte(tamanho: int, negrito: bool = True):
    for caminho in (_FONTES_NEGRITO if negrito else _FONTES_NORMAL):
        try:
            return ImageFont.truetype(caminho, tamanho)
        except OSError:
            continue
    try:
        return ImageFont.load_default(size=tamanho)
    except TypeError:
        return ImageFont.load_default()


def sem_emoji(txt: str) -> str:
    txt = re.sub(r"[^\w\s.,;:!?¿¡\-–—'\"()%/$@#&+ºª°*]", "", str(txt or ""), flags=re.UNICODE)
    return re.sub(r"\s{2,}", " ", txt).strip()


def hex_rgb(h: str):
    h = (h or "#16A34A").lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def quebrar(draw, texto, fnt, largura_max):
    linhas, atual = [], ""
    for palavra in texto.split():
        teste = f"{atual} {palavra}".strip()
        if draw.textlength(teste, font=fnt) <= largura_max:
            atual = teste
        else:
            if atual:
                linhas.append(atual)
            atual = palavra
    if atual:
        linhas.append(atual)
    return linhas


def texto_ajustado(draw, texto, largura, altura_max, maior, menor, negrito=True):
    """Escolhe o maior tamanho de fonte em que o texto cabe na área."""
    for t in range(maior, menor - 1, -4):
        f = fonte(t, negrito)
        linhas = quebrar(draw, texto, f, largura)
        alt = len(linhas) * int(t * 1.22)
        if alt <= altura_max:
            return f, linhas, int(t * 1.22)
    f = fonte(menor, negrito)
    return f, quebrar(draw, texto, f, largura), int(menor * 1.22)


def contraste(rgb):
    r, g, b = rgb
    return (20, 20, 30) if (0.299 * r + 0.587 * g + 0.114 * b) > 165 else (255, 255, 255)


def gerar_arte(post: dict, modelo: str, marca: str) -> bytes:
    return _arte(post.get("titulo_arte", ""), post.get("subtitulo", ""), post.get("cta", ""),
                 modelo, marca, S.marca_cor1, S.marca_cor2, S.marca_logo)


@st.cache_data(show_spinner=False, max_entries=200)
def _arte(titulo, sub, cta, modelo, marca, cor1, cor2, logo_b64) -> bytes:
    W = H = 1080
    c1, c2 = hex_rgb(cor1), hex_rgb(cor2)
    titulo = sem_emoji(titulo) or "Sua oferta aqui"
    sub    = sem_emoji(sub)
    cta    = sem_emoji(cta) or "Saiba mais"
    marca  = sem_emoji(marca)

    img = Image.new("RGB", (W, H), c2)
    d = ImageDraw.Draw(img)

    if modelo == "Gradiente":
        for y in range(H):
            t = y / H
            cor = tuple(int(c1[i] * (1 - t) + c2[i] * t) for i in range(3))
            d.line([(0, y), (W, y)], fill=cor)
        cor_txt = (255, 255, 255)
        area = (90, 200, W - 180, 520)
    elif modelo == "Faixa":
        d.rectangle([0, 0, W, H], fill=(255, 255, 255))
        d.rectangle([0, 0, W, 250], fill=c1)
        d.rectangle([0, H - 110, W, H], fill=c2)
        cor_txt = c2
        area = (90, 320, W - 180, 330)
    else:  # Minimal
        d.rectangle([0, 0, W, H], fill=(250, 250, 247))
        d.rectangle([60, 60, W - 60, H - 60], outline=c1, width=10)
        cor_txt = c2
        area = (130, 220, W - 260, 480)

    x0, y0, larg, alt = area
    f_t, linhas_t, lh = texto_ajustado(d, titulo, larg, alt, 112, 52)
    y = y0
    for ln in linhas_t:
        d.text((x0, y), ln, font=f_t, fill=cor_txt)
        y += lh

    limite_botao = H - (270 if modelo == "Minimal" else 230)   # topo máximo do botão
    if sub:
        y += 24
        disp = limite_botao - 40 - y                                 # espaço que sobra para o subtítulo
        if disp >= 30:
            f_s, linhas_s, lh_s = texto_ajustado(d, sub, larg, min(170, disp), 46, 26, negrito=False)
            cabem = max(1, disp // lh_s)
            if len(linhas_s) > cabem:                                # corta com reticências se não couber
                linhas_s = linhas_s[:cabem]
                linhas_s[-1] = linhas_s[-1].rstrip(".,; ") + "…"
            cor_sub = cor_txt if modelo != "Faixa" else (70, 70, 80)
            for ln in linhas_s:
                d.text((x0, y), ln, font=f_s, fill=cor_sub)
                y += lh_s

    # botão de chamada
    f_c = fonte(40)
    tw = d.textlength(cta, font=f_c)
    bx0 = x0
    by0 = min(y + 40, limite_botao)
    cor_botao = c1 if modelo != "Gradiente" else (255, 255, 255)
    cor_txt_botao = contraste(cor_botao)
    d.rounded_rectangle([bx0, by0, bx0 + tw + 80, by0 + 86], radius=43, fill=cor_botao)
    d.text((bx0 + 40, by0 + 20), cta, font=f_c, fill=cor_txt_botao)

    # marca no rodapé
    if marca:
        if modelo == "Faixa":       # nome da marca na faixa de cima
            f_m, _, _ = texto_ajustado(d, marca, W - 380, 70, 60, 34)
            d.text((90, 125 - f_m.size // 2 if hasattr(f_m, "size") else 95), marca, font=f_m, fill=contraste(c1))
        else:
            f_m = fonte(34)
            cor_m = (255, 255, 255) if modelo == "Gradiente" else c2
            ym = H - 95 if modelo == "Gradiente" else H - 140
            xm = 90 if modelo == "Gradiente" else 130
            d.text((xm, ym), marca, font=f_m, fill=cor_m)

    # logo (opcional) no canto superior direito
    if logo_b64:
        try:
            logo = Image.open(io.BytesIO(base64.b64decode(logo_b64))).convert("RGBA")
            logo.thumbnail((180, 180))
            img.paste(logo, (W - logo.width - 80, 80), logo)
        except Exception:
            pass

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════
# WHATSAPP (links de clique para conversar)
# ══════════════════════════════════════════════════════════════════════════
def numero_wa(tel: str) -> str:
    bruto = str(tel or "").strip()
    n = re.sub(r"\D", "", bruto)
    internacional = bruto.startswith("+") or n.startswith("00")
    if n.startswith("00"):
        n = n[2:]
    if not internacional and len(n) in (10, 11):   # DDD + número brasileiro sem o 55
        n = "55" + n
    return n


def personalizar(texto: str, contato: dict) -> str:
    nome = str(contato.get("Nome", "") or "").strip().split(" ")[0] or "tudo bem"
    return (texto or "").replace("{nome}", nome).replace("{link}", S.briefing.get("link", "") or "")


def link_wa(tel: str, texto: str) -> str:
    n = numero_wa(tel)
    base = f"https://wa.me/{n}" if n else "https://wa.me/"
    return base + "?text=" + urllib.parse.quote(texto or "")


# ══════════════════════════════════════════════════════════════════════════
# EXPORTAÇÕES (texto, Excel sem biblioteca extra, projeto .json)
# ══════════════════════════════════════════════════════════════════════════
def _xml(t) -> str:
    t = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", str(t))[:32000]
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def xlsx_bytes(abas: list) -> bytes:
    """abas = [(nome, [[cabeçalho], [linha], ...]), ...]"""
    folhas = []
    for _, linhas in abas:
        xml, larg = [], {}
        for i, linha in enumerate(linhas, 1):
            cels = []
            for j, v in enumerate(linha, 1):
                ref = f"{_col(j)}{i}"
                if isinstance(v, bool) or v is None:
                    v = "" if v is None else ("Sim" if v else "Não")
                if isinstance(v, (int, float)):
                    est = 1 if i == 1 else (2 if isinstance(v, float) else 4)
                    cels.append(f'<c r="{ref}" s="{est}"><v>{v}</v></c>')
                    tam = len(str(v))
                else:
                    sv = str(v)
                    est = 1 if i == 1 else (3 if (len(sv) > 60 or "\n" in sv) else 4)
                    cels.append(f'<c r="{ref}" s="{est}" t="inlineStr"><is><t xml:space="preserve">{_xml(sv)}</t></is></c>')
                    tam = max((len(x) for x in sv.splitlines()), default=0)
                larg[j] = max(larg.get(j, 10), min(tam + 2, 70))
            xml.append(f'<row r="{i}">{"".join(cels)}</row>')
        cols = "".join(f'<col min="{j}" max="{j}" width="{w}" customWidth="1"/>' for j, w in sorted(larg.items()))
        congelar = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
                    'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>') if len(linhas) > 1 else ""
        folhas.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                      + congelar + (f"<cols>{cols}</cols>" if cols else "")
                      + f'<sheetData>{"".join(xml)}</sheetData></worksheet>')
    n = len(folhas)
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{k}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for k in range(1, n + 1))
          + "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{ns_rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
    usados, nomes = set(), []
    for nome, _ in abas:
        base = re.sub(r"[\[\]:*?/\\]", " ", nome).strip()[:28] or "Dados"
        final, c = base, 2
        while final.lower() in usados:
            final, c = f"{base[:25]} {c}", c + 1
        usados.add(final.lower())
        nomes.append(final)
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{ns_rel}"><sheets>'
          + "".join(f'<sheet name="{_xml(nomes[k - 1])}" sheetId="{k}" r:id="rId{k}"/>' for k in range(1, n + 1))
          + "</sheets></workbook>")
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join(f'<Relationship Id="rId{k}" Type="{ns_rel}/worksheet" Target="worksheets/sheet{k}.xml"/>' for k in range(1, n + 1))
               + f'<Relationship Id="rId{n + 1}" Type="{ns_rel}/styles" Target="styles.xml"/></Relationships>')
    estilos = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
               '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>'
               '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
               '<fill><patternFill patternType="solid"><fgColor rgb="FF16A34A"/><bgColor indexed="64"/></patternFill></fill></fills>'
               '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
               '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
               '<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
               '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
               '<xf numFmtId="4" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="top"/></xf>'
               '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
               '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs>'
               '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", estilos)
        for k, x in enumerate(folhas, 1):
            z.writestr(f"xl/worksheets/sheet{k}.xml", x)
    return buf.getvalue()


def limpar_md(t: str) -> str:
    t = re.sub(r"\*\*(.+?)\*\*", r"\1", str(t or ""))
    t = re.sub(r"(?m)^\s{0,3}#{1,6}\s+", "", t)
    return t.strip()


def contagem_status() -> dict:
    c = {s: 0 for s in STATUS}
    for ct in S.contatos:
        c[ct.get("Status") or "Novo"] = c.get(ct.get("Status") or "Novo", 0) + 1
    return c


def relatorio_txt() -> str:
    b = S.briefing
    L = ["=" * 64, "🚀 CAMPANHA 360 IA — RELATÓRIO DA CAMPANHA",
         f"Responsável: {S.c_nome_user}", f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", "=" * 64, ""]
    L += ["■ BRIEFING", "-" * 64, briefing_txt() or "(não preenchido)", ""]
    if S.estrategia:
        L += ["■ ESTRATÉGIA", "-" * 64, limpar_md(S.estrategia), ""]
    if S.funil:
        L += ["■ FUNIL DE VENDAS", "-" * 64, limpar_md(S.funil), ""]
    if S.posts:
        L += [f"■ POSTS ({len(S.posts)})", "-" * 64]
        for i, p in enumerate(S.posts, 1):
            L += [f"{i}) {p.get('tema', '')} — {p.get('formato', '')}",
                  f"   Texto da arte: {p.get('titulo_arte', '')}",
                  f"   Subtítulo: {p.get('subtitulo', '')}",
                  f"   Chamada: {p.get('cta', '')}",
                  "   Legenda:", "\n".join("      " + x for x in str(p.get('legenda', '')).splitlines()),
                  f"   Hashtags: {p.get('hashtags', '')}", ""]
    if S.mensagens:
        L += [f"■ MENSAGENS DE WHATSAPP ({len(S.mensagens)})", "-" * 64]
        for i, m in enumerate(S.mensagens, 1):
            L += [f"{i}) {m.get('etapa', '')} — {m.get('quando', '')}",
                  "\n".join("   " + x for x in str(m.get('texto', '')).splitlines()), ""]
    if S.contatos:
        c = contagem_status()
        L += [f"■ CONTATOS ({len(S.contatos)})", "-" * 64,
              "Resumo: " + " | ".join(f"{k}: {v}" for k, v in c.items()), ""]
        for i, ct in enumerate(S.contatos, 1):
            L.append(f"{i}) {ct.get('Nome', '')} — {ct.get('Telefone', '')} — {ct.get('Status', '')}"
                     + (f" — {ct.get('Observação', '')}" if ct.get('Observação') else ""))
        L.append("")
    if S.analise_final:
        L += ["■ ANÁLISE E PRÓXIMOS PASSOS", "-" * 64, limpar_md(S.analise_final), ""]
    return "\n".join(L)


def relatorio_xlsx() -> bytes:
    rot = {"negocio": "Negócio / marca", "produto": "Produto ou serviço", "preco": "Preço / condição",
           "publico": "Público-alvo", "diferencial": "Diferencial", "objetivo": "Objetivo",
           "duracao": "Duração", "canais": "Canais", "tom": "Tom de voz", "oferta": "Oferta / bônus",
           "link": "Link", "cidade": "Cidade / região"}
    abas = [("Briefing", [["Campo", "Valor"]] + [[rot.get(k, k), v] for k, v in S.briefing.items() if str(v).strip()])]
    if S.estrategia:
        abas.append(("Estratégia", [["Estratégia"]] + [[l] for l in limpar_md(S.estrategia).splitlines() if l.strip()]))
    if S.funil:
        abas.append(("Funil", [["Funil de vendas"]] + [[l] for l in limpar_md(S.funil).splitlines() if l.strip()]))
    if S.posts:
        abas.append(("Posts", [["#", "Tema", "Formato", "Texto da arte", "Subtítulo", "Chamada", "Legenda", "Hashtags"]]
                     + [[i, p.get("tema", ""), p.get("formato", ""), p.get("titulo_arte", ""), p.get("subtitulo", ""),
                         p.get("cta", ""), p.get("legenda", ""), p.get("hashtags", "")] for i, p in enumerate(S.posts, 1)]))
    if S.mensagens:
        abas.append(("Mensagens WhatsApp", [["#", "Etapa", "Quando enviar", "Mensagem"]]
                     + [[i, m.get("etapa", ""), m.get("quando", ""), m.get("texto", "")] for i, m in enumerate(S.mensagens, 1)]))
    if S.contatos:
        abas.append(("Contatos", [["#", "Nome", "Telefone", "Status", "Observação"]]
                     + [[i, c.get("Nome", ""), str(c.get("Telefone", "")), c.get("Status", ""), c.get("Observação", "")]
                        for i, c in enumerate(S.contatos, 1)]))
        cs = contagem_status()
        total = max(len(S.contatos), 1)
        abas.append(("Resultados", [["Status", "Quantidade", "% do total"]]
                     + [[k, v, round(v / total * 100, 1)] for k, v in cs.items()]))
    return xlsx_bytes(abas)


def projeto_json() -> bytes:
    chaves = ["briefing", "estrategia", "funil", "posts", "mensagens", "contatos",
              "marca_cor1", "marca_cor2", "marca_logo", "analise_final"]
    dados = {k: S[k] for k in chaves}
    dados["_app"] = "campanha_360"
    dados["_salvo_em"] = datetime.now().isoformat(timespec="seconds")
    return json.dumps(dados, ensure_ascii=False, indent=2).encode("utf-8")


# ══════════════════════════════════════════════════════════════════════════
# TELA DE ENTRADA
# ══════════════════════════════════════════════════════════════════════════
if not S.c_chave:
    _, meio, _ = st.columns([1, 2, 1])
    with meio:
        st.markdown("""
        <div style="text-align:center;padding:20px 0 6px 0;">
          <div style="font-size:3.4em;">🚀</div>
          <h1 style="margin:0;">Campanha 360 IA</h1>
          <p style="color:#6C757D !important;">Estratégia, funil, posts, WhatsApp e relatório — em um só lugar</p>
        </div>""", unsafe_allow_html=True)
        nome = st.text_input("👤 Seu nome", placeholder="Como quer ser chamado?", key="c_inp_nome")
        chave = st.text_input("🔑 Sua chave Groq API", type="password", placeholder="gsk_...", key="c_inp_chave")
        if st.button("✨ ENTRAR", type="primary", width="stretch"):
            if len(nome.strip()) < 2:
                st.warning("Digite um nome com pelo menos 2 letras.")
            elif not chave.strip().startswith("gsk_"):
                st.warning("A chave Groq começa com **gsk_**. Confira e cole novamente.")
            else:
                S.c_nome_user, S.c_chave = nome.strip(), chave.strip()
                st.rerun()
        st.markdown(f"""<div class="card" style="font-size:0.9em;line-height:1.7;">
        🔑 <b>Não tem chave?</b> Crie grátis em <a href="https://console.groq.com/keys" target="_blank">console.groq.com/keys</a>.<br>
        🔒 Sua chave fica só nesta sessão do navegador: não é salva nem enviada para nós.<br>
        💾 Dica: use <b>Salvar projeto</b> para guardar sua campanha e continuar outro dia.<br>
        🤖 Mais agentes em <a href="{SITE}" target="_blank">MeusAgentesIA.com.br</a>.
        </div>""", unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════
# CABEÇALHO + PROJETO
# ══════════════════════════════════════════════════════════════════════════
c_cab, c_proj = st.columns([3, 2])
with c_cab:
    nome_neg = S.briefing.get("negocio") or "sua campanha"
    st.markdown(f"""<div class="cab"><h1>🚀 Campanha 360 IA</h1>
    <p>Olá, {S.c_nome_user}! Montando: <b>{nome_neg}</b></p></div>""", unsafe_allow_html=True)
with c_proj:
    with st.expander("💾 Projeto: salvar, abrir ou sair", expanded=False):
        st.download_button("💾 Salvar projeto (.json)", data=projeto_json(),
                           file_name=f"campanha_{re.sub(r'[^a-z0-9]+', '_', nome_neg.lower())[:30]}.json",
                           mime="application/json", width="stretch")
        arq = st.file_uploader("📂 Abrir projeto salvo", type=["json"], key="c_upload_proj")
        if arq is not None and st.button("Carregar este projeto", width="stretch"):
            try:
                dados = json.loads(arq.getvalue().decode("utf-8"))
                for k in ["briefing", "estrategia", "funil", "posts", "mensagens", "contatos",
                          "marca_cor1", "marca_cor2", "marca_logo", "analise_final"]:
                    if k in dados:
                        S[k] = dados[k]
                st.success("Projeto carregado!")
                st.rerun()
            except Exception:
                st.error("Esse arquivo não é um projeto válido da Campanha 360.")
        if st.button("🧹 Começar campanha nova", width="stretch"):
            for k, v in PADRAO.items():
                if k not in ("c_nome_user", "c_chave"):
                    S[k] = v.copy() if isinstance(v, (dict, list)) else v
            st.rerun()
        if st.button("🚪 Sair", width="stretch"):
            for k in list(S.keys()):
                del S[k]
            st.rerun()

abas = st.tabs(["1️⃣ Briefing", "2️⃣ Estratégia", "3️⃣ Funil", "4️⃣ Posts",
                "5️⃣ WhatsApp", "6️⃣ Acompanhamento", "7️⃣ Relatório"])


def precisa_briefing():
    st.info("👈 Preencha primeiro a aba **1️⃣ Briefing** (pelo menos produto e público).")


# ══════════════════════════════════════════════════════════════════════════
# 1. BRIEFING
# ══════════════════════════════════════════════════════════════════════════
with abas[0]:
    st.markdown('<span class="passo">PASSO 1</span>', unsafe_allow_html=True)
    st.markdown("### 📝 Conte sobre o que você vai vender")
    b = S.briefing
    with st.form("f_briefing"):
        c1, c2 = st.columns(2)
        with c1:
            negocio = st.text_input("Nome do negócio / marca", value=b.get("negocio", ""), placeholder="Ex: Doce Sabor Confeitaria")
            produto = st.text_area("Produto ou serviço *", value=b.get("produto", ""), height=90,
                                   placeholder="Ex: Bolos de pote artesanais, 6 sabores")
            preco   = st.text_input("Preço / condição", value=b.get("preco", ""), placeholder="Ex: R$ 12 cada ou 5 por R$ 50")
            oferta  = st.text_input("Oferta ou bônus da campanha", value=b.get("oferta", ""), placeholder="Ex: frete grátis no bairro até domingo")
            link    = st.text_input("Link de compra ou contato", value=b.get("link", ""), placeholder="Ex: instagram.com/docesabor ou site")
        with c2:
            publico = st.text_area("Público-alvo *", value=b.get("publico", ""), height=90,
                                   placeholder="Ex: mulheres 25-45 anos, que trabalham fora e gostam de um mimo")
            diferencial = st.text_input("Diferencial", value=b.get("diferencial", ""), placeholder="Ex: receitas da avó, sem conservantes")
            objetivos = ["Vender mais", "Captar contatos (leads)", "Lançar produto novo", "Lotar evento/agenda",
                         "Reativar clientes antigos", "Fortalecer a marca"]
            objetivo = st.selectbox("Objetivo da campanha", objetivos,
                                    index=objetivos.index(b["objetivo"]) if b.get("objetivo") in objetivos else 0)
            duracoes = ["7 dias", "15 dias", "30 dias"]
            duracao = st.radio("Duração", duracoes, horizontal=True,
                               index=duracoes.index(b["duracao"]) if b.get("duracao") in duracoes else 0)
            canais_op = ["Instagram", "WhatsApp", "Facebook", "TikTok", "Google Meu Negócio", "E-mail"]
            canais = st.multiselect("Canais", canais_op,
                                    default=[x for x in (b.get("canais") or "Instagram, WhatsApp").split(", ") if x in canais_op])
            tons = ["Próximo e amigável", "Profissional e confiável", "Divertido e descontraído", "Inspirador", "Direto e urgente"]
            tom = st.selectbox("Tom de voz", tons, index=tons.index(b["tom"]) if b.get("tom") in tons else 0)
            cidade = st.text_input("Cidade / região (se for local)", value=b.get("cidade", ""))
        if st.form_submit_button("💾 Salvar briefing", type="primary"):
            S.briefing = {"negocio": negocio, "produto": produto, "preco": preco, "publico": publico,
                          "diferencial": diferencial, "objetivo": objetivo, "duracao": duracao,
                          "canais": ", ".join(canais), "tom": tom, "oferta": oferta, "link": link, "cidade": cidade}
            if not (produto.strip() and publico.strip()):
                st.warning("Preencha pelo menos **Produto** e **Público-alvo** (marcados com *).")
            else:
                st.success("✅ Briefing salvo! Agora vá para **2️⃣ Estratégia**.")

    st.markdown("#### 🎨 Identidade visual (para as artes dos posts)")
    k1, k2, k3 = st.columns([1, 1, 2])
    with k1:
        S.marca_cor1 = st.color_picker("Cor principal", S.marca_cor1)
    with k2:
        S.marca_cor2 = st.color_picker("Cor de apoio (escura)", S.marca_cor2)
    with k3:
        logo = st.file_uploader("Logo (PNG, opcional)", type=["png", "jpg", "jpeg"], key="c_logo")
        if logo is not None:
            S.marca_logo = base64.b64encode(logo.getvalue()).decode()
        if S.marca_logo:
            st.caption("✅ Logo carregada — vai aparecer no canto das artes.")


# ══════════════════════════════════════════════════════════════════════════
# 2. ESTRATÉGIA
# ══════════════════════════════════════════════════════════════════════════
with abas[1]:
    st.markdown('<span class="passo">PASSO 2</span>', unsafe_allow_html=True)
    st.markdown("### 🧭 Estratégia da campanha")
    if not tem_briefing():
        precisa_briefing()
    else:
        if st.button("🧠 Gerar estratégia completa", type="primary", key="b_estr"):
            with st.spinner("A IA está montando sua estratégia…"):
                res = rodar(ia, f"""Crie a ESTRATÉGIA de uma campanha de marketing com base neste briefing:
{briefing_txt()}

Formato (use títulos com ## e listas):
## 🎯 Resumo da campanha (3 linhas)
## 👤 Persona (nome fictício, rotina, dores, desejos, onde está online)
## 💥 Promessa principal — 3 opções de headline
## 🧱 Objeções e como responder (tabela: Objeção | Resposta)
## 🗣️ Tom de voz (o que fazer e o que evitar)
## 📅 Calendário da campanha ({S.briefing.get('duracao', '7 dias')}) — tabela: Dia | Canal | Conteúdo | Objetivo
## 📊 Metas e indicadores (o que medir e uma meta realista para cada)""")
            if res:
                S.estrategia = res
        if S.estrategia:
            st.markdown(f'<div class="card">', unsafe_allow_html=True)
            st.markdown(S.estrategia)
            st.markdown("</div>", unsafe_allow_html=True)
            with st.expander("✏️ Editar texto da estratégia"):
                novo = st.text_area("Estratégia", S.estrategia, height=300, key="ed_estr", label_visibility="collapsed")
                if st.button("Salvar edição", key="sv_estr"):
                    S.estrategia = novo
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# 3. FUNIL
# ══════════════════════════════════════════════════════════════════════════
with abas[2]:
    st.markdown('<span class="passo">PASSO 3</span>', unsafe_allow_html=True)
    st.markdown("### 🔻 Funil de vendas")
    if not tem_briefing():
        precisa_briefing()
    else:
        if st.button("🔻 Montar funil completo", type="primary", key="b_funil"):
            with st.spinner("Montando topo, meio, fundo e pós-venda…"):
                res = rodar(ia, f"""Monte um FUNIL DE VENDAS completo para este briefing:
{briefing_txt()}
{('Estratégia já definida (resumo):' + S.estrategia[:1500]) if S.estrategia else ''}

Formato (títulos com ##):
## 🔝 Topo — atrair (objetivo, 5 ideias de conteúdo, formatos)
## 🧲 Meio — captar (isca digital sugerida + TEXTO PRONTO da página de captura: headline, subheadline, 4 bullets de benefício, chamada do botão)
## 💰 Fundo — vender (estrutura da oferta, garantia, bônus, gatilhos ÉTICOS, script de fechamento no WhatsApp)
## 🤝 Pós-venda (agradecimento, pedido de depoimento, oferta de recompra/indicação)
## 📈 Números para acompanhar em cada etapa""")
            if res:
                S.funil = res
        if S.funil:
            st.markdown(S.funil)
            with st.expander("✏️ Editar texto do funil"):
                novo = st.text_area("Funil", S.funil, height=300, key="ed_funil", label_visibility="collapsed")
                if st.button("Salvar edição", key="sv_funil"):
                    S.funil = novo
                    st.rerun()


# ══════════════════════════════════════════════════════════════════════════
# 4. POSTS
# ══════════════════════════════════════════════════════════════════════════
with abas[3]:
    st.markdown('<span class="passo">PASSO 4</span>', unsafe_allow_html=True)
    st.markdown("### 🖼️ Posts com arte pronta")
    if not tem_briefing():
        precisa_briefing()
    else:
        p1, p2, p3 = st.columns([1, 1, 1])
        with p1:
            qtd = st.selectbox("Quantos posts", [3, 6, 9, 12], index=1, key="qtd_posts")
        with p2:
            modelo_arte = st.selectbox("Modelo da arte", ["Gradiente", "Faixa", "Minimal"], key="modelo_arte")
        with p3:
            st.write("")
            gerar = st.button("✨ Criar posts", type="primary", key="b_posts", width="stretch")
        if gerar:
            with st.spinner("Escrevendo os posts…"):
                dados = rodar(ia_json, f"""Crie {qtd} posts para redes sociais para esta campanha:
{briefing_txt()}
Distribua entre as etapas do funil (atrair, educar, provar, vender, urgência ética).
Responda SOMENTE em JSON neste formato:
{{"posts":[{{"tema":"...", "etapa":"atrair|educar|provar|vender|urgência",
"formato":"feed|carrossel|reels|stories",
"titulo_arte":"frase curta e forte para a arte, máximo 8 palavras, sem emoji",
"subtitulo":"complemento da arte, máximo 14 palavras, sem emoji",
"cta":"chamada curta do botão, máximo 4 palavras",
"legenda":"legenda completa e envolvente, com quebras de linha e emojis",
"hashtags":"8 a 12 hashtags relevantes"}}]}}""")
            if dados and isinstance(dados.get("posts"), list):
                S.posts = [p for p in dados["posts"] if isinstance(p, dict)][:qtd]
                st.success(f"✅ {len(S.posts)} posts criados!")

        if S.posts:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                for i, p in enumerate(S.posts, 1):
                    z.writestr(f"post_{i:02d}.png", gerar_arte(p, modelo_arte, S.briefing.get("negocio", "")))
                    z.writestr(f"post_{i:02d}_legenda.txt",
                               f"{p.get('legenda', '')}\n\n{p.get('hashtags', '')}")
            st.download_button("📦 Baixar todas as artes + legendas (.zip)", data=zip_buf.getvalue(),
                               file_name="posts_campanha.zip", mime="application/zip")
            for i, p in enumerate(S.posts):
                with st.container(border=True):
                    a, b_ = st.columns([1, 2])
                    with a:
                        png = gerar_arte(p, modelo_arte, S.briefing.get("negocio", ""))
                        st.image(png)
                        st.download_button("⬇️ Baixar arte", data=png, file_name=f"post_{i + 1:02d}.png",
                                           mime="image/png", key=f"dl_arte_{i}", width="stretch")
                    with b_:
                        st.markdown(f"**{i + 1}. {p.get('tema', '')}**  ·  `{p.get('formato', '')}`  ·  _{p.get('etapa', '')}_")
                        p["titulo_arte"] = st.text_input("Texto da arte", p.get("titulo_arte", ""), key=f"pt_{i}")
                        p["subtitulo"] = st.text_input("Subtítulo", p.get("subtitulo", ""), key=f"ps_{i}")
                        p["cta"] = st.text_input("Chamada do botão", p.get("cta", ""), key=f"pc_{i}")
                        p["legenda"] = st.text_area("Legenda", p.get("legenda", ""), height=150, key=f"pl_{i}")
                        p["hashtags"] = st.text_input("Hashtags", p.get("hashtags", ""), key=f"ph_{i}")
                        st.caption("Edite os textos acima e a arte se atualiza sozinha.")


# ══════════════════════════════════════════════════════════════════════════
# 5. WHATSAPP
# ══════════════════════════════════════════════════════════════════════════
with abas[4]:
    st.markdown('<span class="passo">PASSO 5</span>', unsafe_allow_html=True)
    st.markdown("### 💬 Mensagens e envio pelo WhatsApp")
    st.markdown("""<div class="aviso">
    ✅ <b>Envie só para quem autorizou receber</b> (clientes, quem pediu informação, quem se cadastrou).
    Mandar para números desconhecidos ou em grande quantidade faz o WhatsApp <b>bloquear o seu número</b>.<br>
    📲 Aqui cada contato ganha um link que abre o WhatsApp com a mensagem já escrita — você confere e toca em enviar.
    </div>""", unsafe_allow_html=True)
    st.write("")
    if not tem_briefing():
        precisa_briefing()
    else:
        # ── mensagens
        st.markdown("#### ✍️ Sequência de mensagens")
        m1, m2 = st.columns([3, 1])
        with m1:
            sair = st.checkbox("Incluir no final: “Se não quiser receber mais mensagens, responda SAIR.”", value=True)
        with m2:
            gerar_m = st.button("✨ Criar mensagens", type="primary", key="b_msgs", width="stretch")
        if gerar_m:
            with st.spinner("Escrevendo a sequência…"):
                dados = rodar(ia_json, f"""Crie uma SEQUÊNCIA de mensagens de WhatsApp para esta campanha:
{briefing_txt()}
Etapas: 1) primeiro contato/abertura, 2) conteúdo de valor, 3) apresentação da oferta,
4) quebra de objeção, 5) lembrete/últimas horas (sem falsa escassez), 6) agradecimento pós-compra,
7) reativação para quem não respondeu.
Regras: mensagens curtas (até 600 caracteres), naturais, com no máximo 2 emojis, use {{nome}} para o
primeiro nome do contato e {{link}} onde entrar o link. Nada de CAIXA ALTA exagerada.
Responda SOMENTE em JSON: {{"mensagens":[{{"etapa":"...","quando":"ex: Dia 1 de manhã","texto":"..."}}]}}""")
            if dados and isinstance(dados.get("mensagens"), list):
                S.mensagens = [m for m in dados["mensagens"] if isinstance(m, dict)]
                if sair:
                    for m in S.mensagens:
                        if "SAIR" not in m.get("texto", ""):
                            m["texto"] = m.get("texto", "").rstrip() + "\n\nSe não quiser receber mais mensagens, responda SAIR."
                st.success(f"✅ {len(S.mensagens)} mensagens criadas!")

        if S.mensagens:
            for i, m in enumerate(S.mensagens):
                with st.expander(f"{i + 1}. {m.get('etapa', '')} — {m.get('quando', '')}"):
                    m["texto"] = st.text_area("Mensagem", m.get("texto", ""), height=140, key=f"mt_{i}")

        # ── contatos
        st.markdown("#### 👥 Lista de contatos")
        with st.expander("➕ Adicionar contatos (colar lista ou subir planilha)", expanded=not S.contatos):
            st.caption("Uma pessoa por linha, no formato **Nome, Telefone** (com DDD). Ex.: `Maria Souza, 11 98888-7777`")
            colado = st.text_area("Colar lista", height=120, key="c_colar", label_visibility="collapsed",
                                  placeholder="Maria Souza, 11 98888-7777\nJoão Lima, (21) 97777-6666")
            planilha = st.file_uploader("…ou subir CSV/Excel com colunas Nome e Telefone", type=["csv", "xlsx"], key="c_planilha")
            if st.button("Adicionar à lista", key="b_add_cont"):
                novos = []
                for linha in (colado or "").splitlines():
                    partes = [x.strip() for x in re.split(r"[,;\t]", linha) if x.strip()]
                    if not partes:
                        continue
                    tel = next((x for x in partes if len(re.sub(r"\D", "", x)) >= 8), "")
                    nome_c = next((x for x in partes if x != tel), "")
                    if tel:
                        novos.append({"Nome": nome_c, "Telefone": tel, "Status": "Novo", "Observação": ""})
                if planilha is not None:
                    try:
                        df_p = pd.read_csv(planilha, sep=None, engine="python") if planilha.name.endswith(".csv") \
                            else pd.read_excel(planilha)
                        cols = {c.lower().strip(): c for c in df_p.columns}
                        cn = cols.get("nome") or list(df_p.columns)[0]
                        ct = cols.get("telefone") or cols.get("celular") or cols.get("whatsapp") or list(df_p.columns)[-1]
                        for _, r in df_p.iterrows():
                            if str(r[ct]).strip() and str(r[ct]) != "nan":
                                novos.append({"Nome": str(r[cn]) if str(r[cn]) != "nan" else "",
                                              "Telefone": str(r[ct]).replace(".0", ""), "Status": "Novo", "Observação": ""})
                    except Exception:
                        st.error("Não consegui ler a planilha. Para Excel, use CSV se der erro.")
                existentes = {numero_wa(c["Telefone"]) for c in S.contatos}
                add = []
                for c in novos:
                    num = numero_wa(c["Telefone"])
                    if num and num not in existentes:
                        existentes.add(num)
                        add.append(c)
                S.contatos.extend(add)
                st.success(f"✅ {len(add)} contato(s) adicionado(s)" + (f" — {len(novos) - len(add)} repetido(s) ignorado(s)." if len(novos) > len(add) else "."))
                st.rerun()

        if S.contatos:
            opcoes_msg = [f"{i + 1}. {m.get('etapa', '')}" for i, m in enumerate(S.mensagens)] or ["(crie as mensagens acima)"]
            escolha = st.selectbox("Qual mensagem enviar agora?", opcoes_msg, key="c_msg_escolhida")
            idx_m = opcoes_msg.index(escolha) if S.mensagens else -1
            texto_base = S.mensagens[idx_m]["texto"] if idx_m >= 0 else ""

            filtro = st.multiselect("Mostrar contatos com status", STATUS, default=STATUS, key="c_filtro")
            df = pd.DataFrame(S.contatos)
            for col in ["Nome", "Telefone", "Status", "Observação"]:
                if col not in df.columns:
                    df[col] = ""
            df["Telefone"] = df["Telefone"].astype(str)
            df["WhatsApp"] = [link_wa(r["Telefone"], personalizar(texto_base, r)) for _, r in df.iterrows()]
            visiveis = df[df["Status"].isin(filtro)] if filtro else df
            st.caption("Toque em **Abrir WhatsApp**, confira e envie. Depois mude o **Status** na tabela "
                       "(Enviado, Respondeu, Comprou…). As mudanças são salvas na hora.")
            # a key só muda quando as LINHAS mudam (filtro, contatos adicionados/apagados):
            # assim nenhuma edição se perde e uma edição antiga nunca cai na linha errada
            _assin = json.dumps([[int(i) for i in visiveis.index], visiveis["Telefone"].astype(str).tolist(),
                                 len(S.contatos), filtro], ensure_ascii=False)
            editado = st.data_editor(
                visiveis[["Nome", "Telefone", "WhatsApp", "Status", "Observação"]],
                key=f"c_editor_{abs(hash(_assin)) % 10**10}", width="stretch", hide_index=True, num_rows="fixed",
                column_config={
                    "WhatsApp": st.column_config.LinkColumn("WhatsApp", display_text="📲 Abrir WhatsApp", disabled=True),
                    "Status": st.column_config.SelectboxColumn("Status", options=STATUS, required=True),
                    "Telefone": st.column_config.TextColumn("Telefone"),
                },
            )
            # grava as alterações de volta na lista completa
            mudou = False
            for idx, linha in editado.iterrows():
                novo = {"Nome": linha["Nome"] if pd.notna(linha["Nome"]) else "",
                        "Telefone": linha["Telefone"] if pd.notna(linha["Telefone"]) else "",
                        "Status": linha["Status"] or "Novo",
                        "Observação": linha["Observação"] if pd.notna(linha["Observação"]) else ""}
                if any(S.contatos[idx].get(k) != v for k, v in novo.items()):
                    S.contatos[idx].update(novo)
                    mudou = True
            if mudou:
                st.rerun()   # redesenha a tabela já com a mudança (e com o filtro aplicado)

            c_a, c_b = st.columns(2)
            with c_a:
                if st.button("✔️ Marcar todos os 'Novo' visíveis como Enviado", key="b_marcar_env"):
                    for idx in visiveis.index:
                        if S.contatos[idx].get("Status") == "Novo":
                            S.contatos[idx]["Status"] = "Enviado"
                    st.rerun()
            with c_b:
                if st.button("🗑️ Apagar contatos 'Não quer'", key="b_apagar_nq"):
                    S.contatos = [c for c in S.contatos if c.get("Status") != "Não quer"]
                    st.rerun()

        st.markdown("#### 🔗 Link para bio, status e anúncios")
        st.caption("Qualquer pessoa que tocar nesse link abre uma conversa com você, com a mensagem já escrita.")
        l1, l2 = st.columns([1, 2])
        with l1:
            meu_num = st.text_input("Seu WhatsApp (com DDD)", key="c_meu_num", placeholder="11 98888-7777")
        with l2:
            frase = st.text_input("Mensagem que a pessoa vai enviar", key="c_frase",
                                  value=f"Olá! Vi a campanha e quero saber mais sobre {S.briefing.get('produto', '')[:60]}")
        if meu_num:
            st.code(link_wa(meu_num, frase), language=None)


# ══════════════════════════════════════════════════════════════════════════
# 6. ACOMPANHAMENTO
# ══════════════════════════════════════════════════════════════════════════
with abas[5]:
    st.markdown('<span class="passo">PASSO 6</span>', unsafe_allow_html=True)
    st.markdown("### 📊 Quem recebeu, respondeu e comprou")
    if not S.contatos:
        st.info("Adicione contatos na aba **5️⃣ WhatsApp** para acompanhar aqui.")
    else:
        cs = contagem_status()
        total = len(S.contatos)
        enviados = total - cs["Novo"]
        cols = st.columns(6)
        for col, (rot, val) in zip(cols, [("Contatos", total), ("Enviados", enviados), ("Responderam", cs["Respondeu"] + cs["Interessado"] + cs["Comprou"]),
                                          ("Interessados", cs["Interessado"]), ("Compraram", cs["Comprou"]), ("Não querem", cs["Não quer"])]):
            col.markdown(f'<div class="metrica"><b>{val}</b>{rot}</div>', unsafe_allow_html=True)
        st.write("")
        if enviados:
            taxa_resp = (cs["Respondeu"] + cs["Interessado"] + cs["Comprou"]) / enviados * 100
            taxa_conv = cs["Comprou"] / enviados * 100
            st.markdown(f"**Taxa de resposta:** {taxa_resp:.1f}%  ·  **Conversão em venda:** {taxa_conv:.1f}% (sobre enviados)")
        graf = pd.DataFrame({"Contatos": [cs[s] for s in STATUS]},
                            index=[f"{i + 1}. {s}" for i, s in enumerate(STATUS)])
        st.bar_chart(graf, color="#16A34A", height=260, horizontal=True)

        st.markdown("#### 🧠 Próxima mensagem sugerida para cada grupo")
        grupo = st.selectbox("Grupo", ["Enviado (não respondeu)", "Respondeu", "Interessado", "Comprou"], key="c_grupo")
        if st.button("✨ Sugerir próxima mensagem", type="primary", key="b_prox"):
            with st.spinner("Pensando na melhor abordagem…"):
                res = rodar(ia, f"""Campanha:
{briefing_txt()}
Grupo de contatos: {grupo}.
Escreva 2 opções de PRÓXIMA mensagem de WhatsApp para esse grupo (curtas, naturais, éticas, use {{nome}}),
e diga em 1 linha o melhor momento para enviar e o objetivo da mensagem.""")
            if res:
                S.proxima_msg[grupo] = res
        if S.proxima_msg.get(grupo):
            st.markdown(f'<div class="card">', unsafe_allow_html=True)
            st.markdown(S.proxima_msg[grupo])
            st.markdown("</div>", unsafe_allow_html=True)
            if st.button("➕ Adicionar à sequência de mensagens", key="b_add_prox"):
                S.mensagens.append({"etapa": f"Follow-up: {grupo}", "quando": "conforme sugestão",
                                    "texto": limpar_md(S.proxima_msg[grupo])})
                st.success("Adicionada na aba 5️⃣ WhatsApp.")


# ══════════════════════════════════════════════════════════════════════════
# 7. RELATÓRIO
# ══════════════════════════════════════════════════════════════════════════
with abas[6]:
    st.markdown('<span class="passo">PASSO 7</span>', unsafe_allow_html=True)
    st.markdown("### 📑 Relatório da campanha")
    feito = {"Briefing": tem_briefing(), "Estratégia": bool(S.estrategia), "Funil": bool(S.funil),
             "Posts": bool(S.posts), "Mensagens": bool(S.mensagens), "Contatos": bool(S.contatos)}
    st.markdown("  ·  ".join(f"{'✅' if v else '⬜'} {k}" for k, v in feito.items()))
    if S.contatos and st.button("🧠 Analisar resultados e sugerir próximos passos", type="primary", key="b_analise"):
        cs = contagem_status()
        with st.spinner("Analisando…"):
            res = rodar(ia, f"""Campanha:
{briefing_txt()}
Resultados até agora (contatos por status): {json.dumps(cs, ensure_ascii=False)}. Total: {len(S.contatos)}.
Posts criados: {len(S.posts)}. Mensagens na sequência: {len(S.mensagens)}.
Faça: 1) leitura honesta dos números, 2) 3 pontos que funcionaram ou não, 3) 5 ações práticas para a próxima semana.
Não invente números além dos informados.""")
        if res:
            S.analise_final = res
    if S.analise_final:
        st.markdown(S.analise_final)

    st.markdown("#### 📥 Baixar")
    d1, d2, d3 = st.columns(3)
    agora = datetime.now().strftime("%Y%m%d_%H%M")
    with d1:
        st.download_button("📄 Relatório (.txt)", data=relatorio_txt().encode("utf-8"),
                           file_name=f"campanha_{agora}.txt", mime="text/plain", width="stretch")
    with d2:
        st.download_button("📊 Relatório (Excel)", data=relatorio_xlsx(), file_name=f"campanha_{agora}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           width="stretch")
    with d3:
        st.download_button("💾 Projeto completo (.json)", data=projeto_json(), file_name=f"projeto_campanha_{agora}.json",
                           mime="application/json", width="stretch")
    with st.expander("👀 Ver o relatório em texto"):
        st.text(relatorio_txt())

st.markdown(f'<div class="rodape">Os textos são gerados por IA: revise antes de publicar ou enviar.<br>'
            f'© 2026 <a href="{SITE}" target="_blank" style="color:#9CA3AF;">MeusAgentesIA.com.br</a></div>',
            unsafe_allow_html=True)
