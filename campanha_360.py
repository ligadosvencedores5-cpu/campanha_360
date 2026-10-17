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
from PIL import Image, ImageDraw, ImageFilter, ImageFont

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
    "marca_cor1": "#16A34A", "marca_cor2": "#0F172A", "marca_logo": "", "fotos": [],
    "proxima_msg": {}, "analise_final": "", "ver": 0,
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
        "link": "Link de compra/contato", "cidade": "Cidade / região", "perfil": "Instagram",
    }
    return "\n".join(f"- {rot.get(k, k)}: {v}" for k, v in b.items() if str(v).strip())


def tem_briefing() -> bool:
    return bool(S.briefing.get("produto") and S.briefing.get("publico"))


# ══════════════════════════════════════════════════════════════════════════
# ARTE DOS POSTS (Pillow) — 6 modelos, 3 formatos, fotos, selo de oferta
# ══════════════════════════════════════════════════════════════════════════
MODELOS_ARTE = ["Impacto", "Foto destaque", "Foto + painel", "Oferta", "Elegante", "Citação"]
FORMATOS_ARTE = {
    "Quadrado — feed (1080×1080)": (1080, 1080),
    "Retrato — feed (1080×1350)": (1080, 1350),
    "Stories / Reels (1080×1920)": (1080, 1920),
}

# Fontes profissionais (Google Fonts, licença livre). Baixadas uma vez e guardadas;
# se não houver internet, usa as fontes do sistema.
_FONTES_WEB = {
    "titulo": "ofl/poppins/Poppins-ExtraBold.ttf",
    "negrito": "ofl/poppins/Poppins-Bold.ttf",
    "medio": "ofl/poppins/Poppins-Medium.ttf",
    "texto": "ofl/poppins/Poppins-Regular.ttf",
    "serif": "ofl/dmserifdisplay/DMSerifDisplay-Regular.ttf",
}
_FONTES_SISTEMA = {
    "titulo": ["DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "arialbd.ttf"],
    "negrito": ["DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", "arialbd.ttf"],
    "medio": ["DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "arial.ttf"],
    "texto": ["DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "arial.ttf"],
    "serif": ["DejaVuSerif-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf", "georgiab.ttf"],
}


@st.cache_resource(show_spinner=False)
def _arquivo_fonte(estilo: str):
    import os
    import tempfile
    pasta = os.path.join(tempfile.gettempdir(), "campanha360_fontes")
    os.makedirs(pasta, exist_ok=True)
    caminho = os.path.join(pasta, estilo + ".ttf")
    if not os.path.exists(caminho) or os.path.getsize(caminho) < 10000:
        try:
            r = requests.get("https://raw.githubusercontent.com/google/fonts/main/" + _FONTES_WEB[estilo], timeout=15)
            if r.status_code == 200 and len(r.content) > 10000:
                with open(caminho, "wb") as f:
                    f.write(r.content)
        except requests.RequestException:
            pass
    if os.path.exists(caminho) and os.path.getsize(caminho) > 10000:
        return caminho
    return None


@st.cache_resource(show_spinner=False)
def fonte(tamanho: int, estilo: str = "titulo"):
    tentativas = ([_arquivo_fonte(estilo)] if _arquivo_fonte(estilo) else []) + _FONTES_SISTEMA[estilo]
    for caminho in tentativas:
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


def mistura(c1, c2, t: float):
    return tuple(int(c1[i] * (1 - t) + c2[i] * t) for i in range(3))


def luminancia(rgb) -> float:
    return 0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]


def contraste(rgb):
    return (24, 24, 32) if luminancia(rgb) > 165 else (255, 255, 255)


def destaque(c1, c2):
    """Cor de destaque para selos: amarelo vibrante, ou a cor principal clareada se ela já for amarela."""
    amarelo = (250, 204, 21)
    if abs(luminancia(c1) - luminancia(amarelo)) < 40 and abs(c1[2] - amarelo[2]) < 80:
        return (255, 255, 255)
    return amarelo


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


def texto_ajustado(draw, texto, largura, altura_max, maior, menor, estilo="titulo", entrelinha=1.12):
    """Maior fonte em que o texto cabe na área (largura × altura)."""
    for t in range(maior, menor - 1, -3):
        f = fonte(t, estilo)
        linhas = quebrar(draw, texto, f, largura)
        lh = int(t * entrelinha)
        if len(linhas) * lh <= altura_max and all(draw.textlength(l, font=f) <= largura for l in linhas):
            return f, linhas, lh
    f = fonte(menor, estilo)
    lh = int(menor * entrelinha)
    linhas = quebrar(draw, texto, f, largura)
    cabem = max(1, altura_max // lh)
    if len(linhas) > cabem:
        linhas = linhas[:cabem]
        linhas[-1] = linhas[-1].rstrip(".,; ") + "…"
    return f, linhas, lh


def escrever(img, xy, linhas, fnt, lh, cor, sombra=False, alinhar="esquerda", largura=0):
    """Escreve linhas (com sombra suave opcional). Retorna o y final."""
    x, y = xy
    if sombra:
        camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
        dc = ImageDraw.Draw(camada)
        yy = y
        for ln in linhas:
            xx = x if alinhar == "esquerda" else x + (largura - dc.textlength(ln, font=fnt)) / 2
            dc.text((xx + 3, yy + 5), ln, font=fnt, fill=(0, 0, 0, 150))
            yy += lh
        camada = camada.filter(ImageFilter.GaussianBlur(7))
        img.alpha_composite(camada)
    d = ImageDraw.Draw(img)
    for ln in linhas:
        xx = x if alinhar == "esquerda" else x + (largura - d.textlength(ln, font=fnt)) / 2
        d.text((xx, y), ln, font=fnt, fill=cor)
        y += lh
    return y


def gradiente(w, h, c1, c2, diagonal=True):
    """Degradê suave (diagonal ou vertical)."""
    base = Image.new("RGB", (2, 2))
    base.putpixel((0, 0), c1)
    base.putpixel((1, 0), mistura(c1, c2, 0.5) if diagonal else c1)
    base.putpixel((0, 1), mistura(c1, c2, 0.5) if diagonal else c2)
    base.putpixel((1, 1), c2)
    return base.resize((w, h), Image.BICUBIC).convert("RGBA")


def circulos_decorativos(img, cor, alfa=34):
    w, h = img.size
    camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(camada)
    d.ellipse([w * 0.55, -h * 0.18, w * 1.25, h * 0.38], fill=cor + (alfa,))
    d.ellipse([-w * 0.28, h * 0.62, w * 0.42, h * 1.18], fill=cor + (alfa - 10,))
    d.ellipse([w * 0.78, h * 0.55, w * 0.98, h * 0.55 + w * 0.2], outline=cor + (alfa + 30,), width=6)
    img.alpha_composite(camada)


def pontilhado(img, x0, y0, colunas, linhas, passo, cor, alfa=90, raio=5):
    camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(camada)
    for i in range(colunas):
        for j in range(linhas):
            cx, cy = x0 + i * passo, y0 + j * passo
            d.ellipse([cx - raio, cy - raio, cx + raio, cy + raio], fill=cor + (alfa,))
    img.alpha_composite(camada)


def listras_diagonais(img, cor, alfa=18, espaco=46, largura=14):
    w, h = img.size
    camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(camada)
    for x in range(-h, w + h, espaco):
        d.line([(x, 0), (x + h, h)], fill=cor + (alfa,), width=largura)
    img.alpha_composite(camada)


def selo_estrela(img, cx, cy, raio, texto, cor_fundo, cor_texto, pontas=16):
    """Selo tipo 'explosão' (ex.: -20%, FRETE GRÁTIS)."""
    import math
    camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(camada)
    pts = []
    for k in range(pontas * 2):
        r = raio if k % 2 == 0 else raio * 0.86
        ang = math.pi * k / pontas - math.pi / 2
        pts.append((cx + r * math.cos(ang), cy + r * math.sin(ang)))
    sombra = Image.new("RGBA", img.size, (0, 0, 0, 0))
    ImageDraw.Draw(sombra).polygon([(x + 6, y + 10) for x, y in pts], fill=(0, 0, 0, 90))
    img.alpha_composite(sombra.filter(ImageFilter.GaussianBlur(10)))
    d.polygon(pts, fill=cor_fundo + (255,))
    img.alpha_composite(camada)
    d = ImageDraw.Draw(img)
    f, linhas, lh = texto_ajustado(d, texto.upper(), raio * 1.3, raio * 1.1, int(raio * 0.62), 22, "titulo", 1.0)
    y = cy - len(linhas) * lh / 2 - lh * 0.08
    for ln in linhas:
        d.text((cx - d.textlength(ln, font=f) / 2, y), ln, font=f, fill=cor_texto)
        y += lh


def alt_botao(tamanho: int) -> int:
    return int(tamanho * 2.1)


def botao(img, x, y, texto, cor_fundo, cor_texto, tamanho=40, centro=False, largura_area=0, seta=True):
    """Botão arredondado com sombra; a seta é desenhada (não depende da fonte)."""
    d = ImageDraw.Draw(img)
    f = fonte(tamanho, "negrito")
    tw = d.textlength(texto, font=f)
    alt = alt_botao(tamanho)
    larg_seta = int(tamanho * 1.1) if seta else 0
    larg = tw + tamanho * 2 + larg_seta
    if centro:
        x = x + (largura_area - larg) / 2
    sombra = Image.new("RGBA", img.size, (0, 0, 0, 0))
    ImageDraw.Draw(sombra).rounded_rectangle([x + 4, y + 8, x + larg + 4, y + alt + 8], radius=alt // 2, fill=(0, 0, 0, 80))
    img.alpha_composite(sombra.filter(ImageFilter.GaussianBlur(9)))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([x, y, x + larg, y + alt], radius=alt // 2, fill=cor_fundo)
    bb = d.textbbox((0, 0), texto, font=f)
    ty = y + (alt - (bb[3] - bb[1])) / 2 - bb[1]
    d.text((x + tamanho, ty), texto, font=f, fill=cor_texto)
    if seta:
        sx = x + tamanho + tw + tamanho * 0.45
        sy = y + alt / 2
        e = max(3, tamanho // 9)
        d.line([(sx, sy), (sx + tamanho * 0.62, sy)], fill=cor_texto, width=e)
        d.line([(sx + tamanho * 0.36, sy - tamanho * 0.26), (sx + tamanho * 0.64, sy)], fill=cor_texto, width=e)
        d.line([(sx + tamanho * 0.36, sy + tamanho * 0.26), (sx + tamanho * 0.64, sy)], fill=cor_texto, width=e)
    return y + alt


def etiqueta(img, x, y, texto, cor_fundo_rgba, cor_texto, tamanho=28):
    """Etiqueta arredondada (ex.: NOVIDADE) com fundo translúcido."""
    d = ImageDraw.Draw(img)
    f = fonte(tamanho, "negrito")
    t = texto.upper()
    tw = d.textlength(t, font=f)
    alt = int(tamanho * 1.9)
    camada = Image.new("RGBA", img.size, (0, 0, 0, 0))
    ImageDraw.Draw(camada).rounded_rectangle([x, y, x + tw + tamanho * 1.6, y + alt], radius=alt // 2, fill=cor_fundo_rgba)
    img.alpha_composite(camada)
    d = ImageDraw.Draw(img)
    bb = d.textbbox((0, 0), t, font=f)
    d.text((x + tamanho * 0.8, y + (alt - (bb[3] - bb[1])) / 2 - bb[1]), t, font=f, fill=cor_texto)
    return y + alt


def bloco_texto(img, x, y_top, y_max, largura, titulo, sub, cor_t, cor_s, est_t="titulo", t_max=110, t_min=44,
                s_max=42, s_min=28, est_s="medio", alinhar="esquerda", sombra=False, barra=None,
                barra_antes=False, barra_curta=False, ornamento=None, de_baixo=False, centro_v=False):
    """Título + subtítulo dentro de [y_top, y_max] sem nunca passar do limite.
    de_baixo=True alinha o bloco pela parte de baixo da área (bom para fotos)."""
    d = ImageDraw.Draw(img)
    disp = max(80, y_max - y_top)
    extra = 54 if (barra or ornamento) else 0
    alt_t = int((disp - extra) * (0.60 if sub else 1.0))
    f_t, lt, lh = texto_ajustado(d, titulo, largura, alt_t, t_max, t_min, est_t, 1.1 if est_t != "serif" else 1.12)
    h_t = len(lt) * lh
    resto = disp - extra - h_t - 22
    if sub and resto >= s_min * 1.3:
        f_s, ls, lhs = texto_ajustado(d, sub, largura, resto, s_max, s_min, est_s, 1.32)
    else:
        f_s, ls, lhs = None, [], 0
    total = h_t + extra + (22 + len(ls) * lhs if ls else 0)
    y = (y_max - total) if de_baixo else (y_top + max(0, (disp - total) / 2) if centro_v else y_top)
    larg_al = largura if alinhar == "centro" else 0
    if barra and barra_antes:
        ImageDraw.Draw(img).rounded_rectangle([x, y, x + 120, y + 12], radius=6, fill=barra)
        y += 34
    y = escrever(img, (x, y), lt, f_t, lh, cor_t, sombra=sombra, alinhar=alinhar, largura=larg_al)
    if barra and not barra_antes:
        w_b = 80 if barra_curta else 140
        ImageDraw.Draw(img).rounded_rectangle([x, y + 24, x + w_b, y + 36], radius=6, fill=barra)
        y += 54
    if ornamento:
        cx = x + largura / 2
        dd = ImageDraw.Draw(img)
        dd.line([(cx - 70, y + 20), (cx + 70, y + 20)], fill=ornamento, width=3)
        dd.ellipse([cx - 7, y + 13, cx + 7, y + 27], fill=ornamento)
        y += 44
    if ls:
        y = escrever(img, (x, y + 22), ls, f_s, lhs, cor_s, alinhar=alinhar, largura=larg_al)
    return y


def foto_cobrir(foto_b64: str, w: int, h: int):
    from PIL import ImageOps
    foto = Image.open(io.BytesIO(base64.b64decode(foto_b64)))
    foto = ImageOps.exif_transpose(foto).convert("RGB")
    return ImageOps.fit(foto, (w, h), Image.LANCZOS, centering=(0.5, 0.45)).convert("RGBA")


def colar_logo(img, logo_b64, x, y, max_lado):
    if not logo_b64:
        return 0
    try:
        logo = Image.open(io.BytesIO(base64.b64decode(logo_b64))).convert("RGBA")
        logo.thumbnail((max_lado, max_lado))
        img.alpha_composite(logo, (int(x), int(y)))
        return logo.width
    except Exception:
        return 0


def rodape_marca(img, x, y, marca, perfil, cor, logo_b64="", tamanho=30, centro=False):
    d = ImageDraw.Draw(img)
    partes = [p for p in (marca, perfil) if p]
    if not partes:
        return
    texto = "   •   ".join(partes)
    f = fonte(tamanho, "medio")
    if centro:
        x = (img.size[0] - d.textlength(texto, font=f)) / 2
    d.text((x, y), texto, font=f, fill=cor)


@st.cache_data(show_spinner=False, max_entries=300)
def _arte(titulo, sub, cta, selo, modelo, formato, marca, perfil, cor1, cor2, logo_b64, foto_b64, rotulo) -> bytes:
    W, H = FORMATOS_ARTE.get(formato, (1080, 1080))
    c1, c2 = hex_rgb(cor1), hex_rgb(cor2)
    acc = destaque(c1, c2)
    titulo = sem_emoji(titulo) or "Sua oferta aqui"
    sub, cta, selo = sem_emoji(sub), sem_emoji(cta) or "Saiba mais", sem_emoji(selo)
    marca, perfil, rotulo = sem_emoji(marca), sem_emoji(perfil), sem_emoji(rotulo)
    if modelo in ("Foto destaque", "Foto + painel") and not foto_b64:
        modelo = "Impacto"                       # sem foto, usa o modelo sem foto
    M = 90                                       # margem lateral
    esc = H / 1080                               # escala vertical para retrato/stories
    branco, quase_preto = (255, 255, 255), (24, 24, 32)

    # ── 1. IMPACTO ─────────────────────────────────────────────────────────
    if modelo == "Impacto":
        img = gradiente(W, H, c1, mistura(c1, c2, 0.85))
        circulos_decorativos(img, branco, 30)
        pontilhado(img, W - 300, int(H * 0.70), 6, 4, 40, branco, 70)
        colar_logo(img, logo_b64, M, 60, 100)
        y = int(H * 0.16)
        if rotulo:
            y = etiqueta(img, M, y, rotulo, (255, 255, 255, 50), branco) + 28
        if selo:
            selo_estrela(img, W - 185, int(H * 0.15) + 40, 130, selo, acc, quase_preto)
        y_max = H - int(110 * esc) - alt_botao(40) - 46
        larg_t = W - 2 * M - (250 if selo and H <= 1350 else 30)
        y = bloco_texto(img, M, max(y, int(H * 0.25)), y_max, larg_t, titulo, sub, branco, (238, 242, 248),
                        t_max=118, sombra=True, barra=acc, centro_v=H > 1080)
        botao(img, M, y + 40, cta, branco, c1 if luminancia(c1) < 200 else quase_preto, 40)
        rodape_marca(img, M, H - int(95 * esc), marca, perfil, branco)

    # ── 2. FOTO DESTAQUE ───────────────────────────────────────────────────
    elif modelo == "Foto destaque":
        img = foto_cobrir(foto_b64, W, H)
        sombra = Image.new("RGBA", (1, 256))
        for yy in range(256):
            t = max(0.0, (yy / 255 - 0.28) / 0.72)
            sombra.putpixel((0, yy), mistura((0, 0, 0), c2, 0.35) + (int(240 * t ** 1.1),))
        img.alpha_composite(sombra.resize((W, H)))
        topo = Image.new("RGBA", (1, 256))
        for yy in range(256):
            topo.putpixel((0, yy), (0, 0, 0, int(120 * max(0, 1 - yy / 60))))
        img.alpha_composite(topo.resize((W, H)))
        lx = colar_logo(img, logo_b64, M, 60, 90)
        if marca:
            ImageDraw.Draw(img).text((M + (lx + 20 if lx else 0), 70), marca, font=fonte(34, "negrito"), fill=branco)
        if selo:
            selo_estrela(img, W - 175, 185, 118, selo, acc, quase_preto)
        rodape = int(100 * esc)
        y_max = H - rodape - alt_botao(38) - 40
        y_top = int(H * (0.50 if H <= 1080 else 0.55))
        y = bloco_texto(img, M, y_top, y_max, W - 2 * M, titulo, sub, branco, (232, 232, 238),
                        t_max=104, sombra=True, barra=acc, barra_antes=True, de_baixo=True)
        botao(img, M, y + 34, cta, c1, contraste(c1), 38)
        if perfil:
            ImageDraw.Draw(img).text((M, H - int(78 * esc)), perfil, font=fonte(30, "medio"), fill=(230, 230, 235))

    # ── 3. FOTO + PAINEL ───────────────────────────────────────────────────
    elif modelo == "Foto + painel":
        img = Image.new("RGBA", (W, H), c1 + (255,))
        h_foto = int(H * (0.47 if H <= 1080 else (0.55 if H <= 1350 else 0.60)))
        img.alpha_composite(foto_cobrir(foto_b64, W, h_foto + 60), (0, 0))
        sombra = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        ImageDraw.Draw(sombra).rounded_rectangle([0, h_foto - 12, W, H], radius=60, fill=(0, 0, 0, 110))
        img.alpha_composite(sombra.filter(ImageFilter.GaussianBlur(18)))
        painel = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        ImageDraw.Draw(painel).rounded_rectangle([0, h_foto, W, H + 60], radius=60, fill=c1 + (255,))
        img.alpha_composite(painel)
        cor_t = contraste(c1)
        pontilhado(img, W - 250, H - int(230 * esc), 5, 3, 36, cor_t, 55)
        colar_logo(img, logo_b64, M, 50, 100)
        if selo:
            selo_estrela(img, W - 180, h_foto - 10, 118, selo, acc, quase_preto)
        y_max = H - int(95 * esc) - alt_botao(36) - 34
        y = bloco_texto(img, M, h_foto + int(60 * esc), y_max, W - 2 * M - (150 if selo else 0), titulo, sub,
                        cor_t, mistura(cor_t, c1, 0.15), t_max=92 if H <= 1350 else 104, est_s="texto",
                        centro_v=H > 1350)
        botao(img, M, y + 30, cta, c2 if luminancia(c2) < 150 else quase_preto, branco, 36)
        rodape_marca(img, M, H - int(78 * esc), marca, perfil, mistura(cor_t, c1, 0.2), tamanho=28)

    # ── 4. OFERTA ──────────────────────────────────────────────────────────
    elif modelo == "Oferta":
        fundo = c2 if luminancia(c2) < 140 else mistura(c2, (0, 0, 0), 0.7)
        img = gradiente(W, H, mistura(fundo, c1, 0.25), fundo)
        listras_diagonais(img, branco, 10)
        circulos_decorativos(img, c1, 60)
        colar_logo(img, logo_b64, M, 60, 90)
        d = ImageDraw.Draw(img)
        if marca:
            f_m = fonte(32, "negrito")
            d.text(((W - d.textlength(marca.upper(), font=f_m)) / 2, int(70 * esc)), marca.upper(), font=f_m, fill=(220, 220, 230))
        raio = int(min(W, H) * (0.17 if H <= 1080 else 0.20))
        cy = (int(150 * esc) + raio + 10) if H <= 1080 else int(H * 0.30)
        selo_estrela(img, W // 2, cy, raio, selo or "OFERTA", acc, quase_preto, pontas=20)
        y_max = H - int(100 * esc) - alt_botao(42) - 40
        y = bloco_texto(img, M, cy + raio + int(40 * esc), y_max, W - 2 * M, titulo.upper(), sub, branco,
                        (220, 225, 235), t_max=96 if H > 1080 else 88, sombra=True, alinhar="centro", centro_v=H > 1080)
        botao(img, M, y + 36, cta, c1, contraste(c1), 42, centro=True, largura_area=W - 2 * M)
        rodape_marca(img, 0, H - int(78 * esc), "", perfil, (200, 200, 210), tamanho=28, centro=True)

    # ── 5. ELEGANTE ────────────────────────────────────────────────────────
    elif modelo == "Elegante":
        creme = (248, 244, 238)
        img = Image.new("RGBA", (W, H), creme + (255,))
        d = ImageDraw.Draw(img)
        d.rectangle([40, 40, W - 40, H - 40], outline=c1, width=3)
        d.rectangle([56, 56, W - 56, H - 56], outline=mistura(c1, creme, 0.55), width=1)
        y = int(110 * esc) if H <= 1350 else int(H * 0.12)
        if foto_b64:
            lado = int(min(W, H) * (0.27 if H <= 1080 else 0.36))
            foto = foto_cobrir(foto_b64, lado, lado)
            mascara = Image.new("L", (lado, lado), 0)
            ImageDraw.Draw(mascara).ellipse([0, 0, lado, lado], fill=255)
            borda = Image.new("RGBA", (lado + 16, lado + 16), (0, 0, 0, 0))
            ImageDraw.Draw(borda).ellipse([0, 0, lado + 16, lado + 16], fill=c1 + (255,))
            img.alpha_composite(borda, ((W - lado - 16) // 2, y - 8))
            img.paste(foto, ((W - lado) // 2, y), mascara)
            y += lado + 44
        elif logo_b64:
            colar_logo(img, logo_b64, (W - 130) / 2, y, 130)
            y += 160
        d = ImageDraw.Draw(img)
        eyebrow = (rotulo or marca or "").upper()
        if eyebrow:
            f_e = fonte(26, "negrito")
            espacado = " ".join(eyebrow)
            d.text(((W - d.textlength(espacado, font=f_e)) / 2, y), espacado, font=f_e, fill=c1)
            y += 58
        rodape = int(120 * esc)
        y_max = H - rodape - alt_botao(34) - 40
        y = bloco_texto(img, M + 30, y, y_max, W - 2 * M - 60, titulo, sub,
                        c2 if luminancia(c2) < 160 else quase_preto, (90, 90, 100), est_t="serif", t_max=100,
                        est_s="texto", alinhar="centro", ornamento=c1, centro_v=H > 1080)
        botao(img, M, y + 34, cta, c1, contraste(c1), 34, centro=True, largura_area=W - 2 * M, seta=False)
        rodape_marca(img, 0, H - int(105 * esc), marca if eyebrow != (marca or "").upper() else "", perfil,
                     (120, 120, 130), tamanho=28, centro=True)

    # ── 6. CITAÇÃO / DEPOIMENTO ────────────────────────────────────────────
    else:
        img = gradiente(W, H, mistura(c1, (255, 255, 255), 0.85), mistura(c1, (255, 255, 255), 0.65))
        circulos_decorativos(img, c1, 40)
        cx0, cy0, cx1, cy1 = 70, int(140 * esc), W - 70, H - int(150 * esc)
        if H > 1350:                     # stories: cartão centralizado, não tão alto
            cy0, cy1 = int(H * 0.22), int(H * 0.76)
        sombra = Image.new("RGBA", img.size, (0, 0, 0, 0))
        ImageDraw.Draw(sombra).rounded_rectangle([cx0 + 6, cy0 + 16, cx1 + 6, cy1 + 16], radius=44, fill=(0, 0, 0, 70))
        img.alpha_composite(sombra.filter(ImageFilter.GaussianBlur(22)))
        d = ImageDraw.Draw(img)
        d.rounded_rectangle([cx0, cy0, cx1, cy1], radius=44, fill=(255, 255, 255))
        d.text((cx0 + 50, cy0 - 60), "“", font=fonte(240, "serif"), fill=c1)
        y_max = cy1 - 60 - alt_botao(34) - 34
        y = bloco_texto(img, cx0 + 70, cy0 + 170, y_max, cx1 - cx0 - 140, titulo, sub, quase_preto, (100, 100, 110),
                        est_t="serif", t_max=92 if H > 1080 else 78, t_min=36, s_max=36, barra=c1, barra_curta=True,
                        centro_v=True)
        botao(img, cx0 + 70, y + 30, cta, c1, contraste(c1), 34)
        colar_logo(img, logo_b64, cx1 - 140, cy1 - 140, 90)
        rodape_marca(img, 0, H - int(100 * esc), marca, perfil, mistura(c1, (0, 0, 0), 0.4), tamanho=30, centro=True)

    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG", optimize=True)
    return buf.getvalue()


def modelo_do_post(i: int, p: dict, modelo_geral: str) -> str:
    """Modelo escolhido no post; senão o geral; no 'Variado', alterna os modelos e
    usa 'Citação' para posts de prova (depoimento) e 'Oferta' só quando há selo."""
    if p.get("modelo") in MODELOS_ARTE:
        return p["modelo"]
    if modelo_geral != "🎲 Variado":
        return modelo_geral
    etapa = str(p.get("etapa", "")).lower()
    if "prov" in etapa:
        return "Citação"
    tem_foto = bool(S.get("fotos"))
    ordem = (["Foto destaque", "Impacto", "Foto + painel", "Oferta", "Elegante"] if tem_foto
             else ["Impacto", "Oferta", "Elegante"])
    escolha = ordem[i % len(ordem)]
    if escolha == "Oferta" and not str(p.get("selo", "")).strip():
        escolha = "Impacto" if not tem_foto else "Foto + painel"
    return escolha


def foto_do_post(i: int, p: dict) -> str:
    fotos = S.get("fotos") or []
    if not fotos:
        return ""
    idx = p.get("foto_idx")
    if idx == -1:
        return ""
    if isinstance(idx, int) and 0 <= idx < len(fotos):
        return fotos[idx]
    return fotos[i % len(fotos)]


def gerar_arte(post: dict, modelo: str, marca: str, formato: str = "", i: int = 0) -> bytes:
    formato = formato or list(FORMATOS_ARTE)[0]
    return _arte(post.get("titulo_arte", ""), post.get("subtitulo", ""), post.get("cta", ""), post.get("selo", ""),
                 modelo, formato, marca, S.briefing.get("perfil", ""), S.marca_cor1, S.marca_cor2, S.marca_logo,
                 foto_do_post(i, post), post.get("etiqueta", ""))


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
              "marca_cor1", "marca_cor2", "marca_logo", "fotos", "analise_final"]
    dados = {k: S[k] for k in chaves}
    dados["_app"] = "campanha_360"
    dados["_salvo_em"] = datetime.now().isoformat(timespec="seconds")
    return json.dumps(dados, ensure_ascii=False, indent=2).encode("utf-8")


# ══════════════════════════════════════════════════════════════════════════
# GUIA PARA INICIANTES (aparece no rodapé — inclusive na tela de entrada)
# ══════════════════════════════════════════════════════════════════════════
GUIA = [
    ("🧰 O que você precisa antes de começar", """
- **Um computador ou celular com internet.** No computador a experiência é mais confortável.
- **Uma chave Groq** (é grátis — veja o próximo item). É ela que faz a inteligência artificial funcionar.
- **WhatsApp instalado** no celular, ou o **WhatsApp Web** aberto no computador (web.whatsapp.com).
- **Uma lista de contatos que autorizaram receber suas mensagens** (clientes, pessoas que pediram informação).
- **Uns 20 minutos** para montar a primeira campanha do começo ao fim.
"""),
    ("🔑 Como criar sua chave Groq (grátis, 2 minutos)", """
1. Acesse **console.groq.com/keys** (no computador ou no celular).
2. Clique em **Sign in / Login** e entre com sua conta Google (é o jeito mais fácil).
3. Clique no botão **Create API Key**.
4. Dê um nome qualquer (ex.: *campanha*) e confirme.
5. Vai aparecer um código que começa com **gsk_**. Clique em **Copy** para copiar.
6. Volte aqui, cole no campo **Sua chave Groq API** e clique em **ENTRAR**.

⚠️ **Guarde a chave num lugar seguro** (bloco de notas, por exemplo). O site da Groq só mostra o código completo uma vez.
Se perder, é só criar outra. **Não compartilhe sua chave com ninguém.**
"""),
    ("1️⃣ Aba Briefing — conte o que você vende", """
É a parte **mais importante**: tudo o que a IA cria depois sai daqui. Quanto mais detalhes, melhor o resultado.
- **Produto ou serviço** e **Público-alvo** são obrigatórios (têm *).
- Em **Público-alvo**, descreva quem compra: idade, profissão, o que essa pessoa quer, onde mora.
  Ex.: *“Mães de 28 a 40 anos, que trabalham fora e querem praticidade no lanche dos filhos”*.
- Em **Diferencial**, diga por que comprar de você e não do concorrente.
- **Oferta ou bônus**: uma vantagem com prazo (frete grátis, brinde, desconto) aumenta muito as vendas.
- **Link de compra ou contato**: seu Instagram, site ou link do WhatsApp. Ele entra automaticamente nas mensagens.
- Clique em **💾 Salvar briefing**.
- Embaixo, escolha as **cores da sua marca** e, se tiver, envie a **logo** (de preferência PNG com fundo transparente).
"""),
    ("2️⃣ Aba Estratégia — o plano da campanha", """
Clique em **🧠 Gerar estratégia completa** e espere alguns segundos.
Você recebe: o perfil do seu cliente ideal, 3 frases de impacto, as dúvidas que impedem a compra (e como responder),
o calendário dia a dia e as metas.
- Leia com calma: é o seu **roteiro**. Se algo não combinar com o seu negócio, abra **✏️ Editar texto** e ajuste.
- Partes marcadas como **[AJUSTAR]** são informações que você não preencheu — troque pelo dado real.
"""),
    ("3️⃣ Aba Funil — o caminho até a venda", """
O funil é o caminho que a pessoa faz: **conhece você → se interessa → compra → volta a comprar**.
Clique em **🔻 Montar funil completo**. Você recebe:
- **Topo:** ideias de conteúdo para atrair gente nova.
- **Meio:** uma isca (ex.: e-book, cupom) e o **texto pronto** de uma página de cadastro.
- **Fundo:** como apresentar a oferta e um roteiro para fechar a venda no WhatsApp.
- **Pós-venda:** como agradecer, pedir depoimento e vender de novo.
"""),
    ("4️⃣ Aba Posts — artes prontas para publicar", """
1. **Envie fotos do seu produto** em **📸 Fotos do produto** (de 1 a 6). Fotos deixam as artes muito mais bonitas —
   tire com boa luz, de preferência perto de uma janela.
2. Escolha **quantos posts**, o **modelo** e o **formato**:
   - **🎲 Variado**: um modelo diferente para cada post (recomendado).
   - **Impacto** (cores fortes), **Foto destaque** (foto de fundo), **Foto + painel**, **Oferta** (selo grande de promoção),
     **Elegante** (fino, com foto redonda) e **Citação** (para depoimentos).
   - **Quadrado** e **Retrato** são para o feed; **Stories / Reels** é o formato em pé, de tela cheia.
3. Clique em **✨ Criar posts**.
4. Cada post vem com **arte, legenda e hashtags**. Mude qualquer texto e a arte **se atualiza sozinha**.
   Em cada post dá para trocar o **modelo** e a **foto**, e editar o **selo** (ex.: *-20%*, *FRETE GRÁTIS*) e a **etiqueta** do topo.
   Deixe o selo vazio se não tiver promoção.
5. Clique em **⬇️ Baixar arte** (um por um) ou em **📦 Baixar todas as artes + legendas (.zip)**.

**Para publicar no Instagram/Facebook:**
- No computador: abra o arquivo .zip (clique duas vezes ou botão direito → *Extrair tudo*).
  Você terá as imagens (.png) e as legendas (.txt).
- Mande as imagens para o seu celular (WhatsApp para você mesmo, e-mail ou Google Drive).
- No Instagram, toque em **+**, escolha a imagem e **cole a legenda** do arquivo .txt correspondente.
- Siga o **calendário** da aba Estratégia para saber o que postar em cada dia.
"""),
    ("5️⃣ Aba WhatsApp — mensagens e envio", """
**Passo A — criar as mensagens:** clique em **✨ Criar mensagens**. São 7 mensagens, uma para cada momento da campanha.
Onde aparecer **{nome}**, o app coloca o primeiro nome de cada contato; **{link}** vira o link do seu briefing.

**Passo B — colocar seus contatos:** abra **➕ Adicionar contatos** e cole um por linha, assim:
`Maria Souza, 11 98888-7777`
Sempre com **DDD**. Pode também subir uma planilha **.csv** com as colunas **Nome** e **Telefone**.

**Passo C — enviar:**
1. Escolha em **Qual mensagem enviar agora?** (comece pela 1, Abertura).
2. Na tabela, clique em **📲 Abrir WhatsApp** do primeiro contato.
3. O WhatsApp abre com a conversa e a mensagem **já escrita**. Confira e toque em **enviar**.
4. Volte para o app e mude o **Status** daquele contato para **Enviado**
   (clique duas vezes na célula Status e escolha). Ou use **✔️ Marcar todos os 'Novo' visíveis como Enviado** no final.
5. Repita para o próximo contato.

💡 **No computador**, deixe o **WhatsApp Web** aberto e conectado antes: o link abre direto nele.
💡 **No celular**, o link abre o aplicativo do WhatsApp; depois de enviar, volte ao navegador.

**Link para bio/anúncios:** no fim da aba, digite **seu** número e copie o link gerado.
Coloque na bio do Instagram, no status e nos anúncios: quem clicar abre uma conversa com você com a mensagem pronta.
"""),
    ("6️⃣ Aba Acompanhamento — quem respondeu e comprou", """
Sempre que alguém responder, se interessar ou comprar, volte na tabela da aba **5️⃣ WhatsApp** e **atualize o Status**:
- **Enviado** → você mandou a mensagem.
- **Respondeu** → a pessoa respondeu.
- **Interessado** → perguntou preço, pediu detalhes.
- **Comprou** → fechou! 🎉
- **Não quer** → pediu para não receber mais. **Respeite e não mande mais nada** (pode apagar com 🗑️).

Na aba **6️⃣ Acompanhamento** você vê os números e a taxa de resposta e de venda.
Escolha um **grupo** (ex.: *Enviado (não respondeu)*) e clique em **✨ Sugerir próxima mensagem**:
a IA escreve o próximo passo para aquele grupo. Gostou? Clique em **➕ Adicionar à sequência**.
"""),
    ("7️⃣ Aba Relatório — resultado e próximos passos", """
- Clique em **🧠 Analisar resultados** para a IA ler seus números e sugerir 5 ações para a próxima semana.
- **📄 Relatório (.txt)**: abre no bloco de notas ou no celular.
- **📊 Relatório (Excel)**: abre no Excel, Google Planilhas ou LibreOffice — uma aba para cada parte da campanha.
- **💾 Projeto completo (.json)**: é o arquivo para **continuar a campanha outro dia** (veja abaixo).
"""),
    ("💾 Como salvar e continuar outro dia (IMPORTANTE)", """
O app **não guarda nada sozinho**: se você fechar a página ou ficar muito tempo parado, **perde o que fez**.
Por isso, sempre que terminar (ou fizer uma parte grande):
1. Abra **💾 Projeto: salvar, abrir, nova campanha ou sair** (no topo, à direita).
2. Clique em **💾 Salvar projeto (.json)** — o arquivo vai para a pasta *Downloads*.

**Para continuar depois:** entre de novo com sua chave, abra **💾 Projeto**, clique em **📂 Abrir projeto salvo**,
escolha o arquivo .json e clique em **Carregar este projeto**. Tudo volta: textos, posts, mensagens e contatos.

💡 Dá para ter **um arquivo por campanha ou por cliente**. Dê nomes fáceis, ex.: *campanha_natal_docesabor.json*.
"""),
    ("🆕 Como começar uma campanha nova (outro produto ou outro cliente)", """
1. **Salve a campanha atual** primeiro: **💾 Projeto → 💾 Salvar projeto (.json)**.
2. No mesmo painel, clique em **🆕 Começar campanha nova** e confirme em **✅ Sim, nova campanha**.
3. Tudo fica em branco (briefing, posts, mensagens, contatos, fotos e cores). Sua chave continua — não precisa entrar de novo.
4. Comece pela aba **1️⃣ Briefing**.

**Quer usar uma campanha antiga como base?** Abra o projeto antigo (**📂 Abrir projeto salvo**), mude o que precisar
no briefing e gere de novo só as partes que quiser. Depois **salve com outro nome** para não misturar.
"""),
    ("🛡️ Regras de ouro para não ter o WhatsApp bloqueado", """
- Envie **só para quem conhece você** ou pediu informação. Nunca compre listas de números.
- Comece devagar: **no máximo 20 a 30 mensagens por dia** nos primeiros dias, e aumente aos poucos.
- **Personalize** (o app já coloca o nome) e evite mandar o mesmo texto idêntico para centenas de pessoas.
- Mantenha a frase **“responda SAIR”** e **respeite** quem pedir para sair.
- Se muita gente bloquear ou denunciar seu número, o WhatsApp pode **restringir ou banir** sua conta.
- Prefira usar o **WhatsApp Business** (é grátis): fica mais profissional e tem respostas rápidas e catálogo.
"""),
    ("🗓️ Rotina sugerida para uma campanha de 7 dias", """
| Dia | O que fazer |
|---|---|
| Antes | Briefing → Estratégia → Funil → Posts → Mensagens. **Salve o projeto.** |
| Dia 1 | Post de atração + mensagem **1 (Abertura)** para os contatos |
| Dia 2 | Post de conteúdo útil + mensagem **2 (Valor)** |
| Dia 3 | Post de prova (depoimento/bastidor) + mensagem **3 (Oferta)** |
| Dia 4 | Responder dúvidas + mensagem **4 (Objeção)** para quem não respondeu |
| Dia 5 | Post de oferta + atualizar os Status |
| Dia 6 | Mensagem **5 (Lembrete)** — só para quem demonstrou interesse |
| Dia 7 | Mensagem **6 (Pós-compra)** para quem comprou · Relatório + análise da IA |
"""),
    ("❓ Problemas comuns e como resolver", """
- **“Sua chave Groq não foi aceita”** → a chave foi copiada errada ou apagada. Crie uma nova em console.groq.com/keys, clique em **Sair** e entre de novo.
- **“Limite de uso atingido”** → a conta grátis da Groq tem limite por minuto/dia. Espere alguns minutos e tente de novo.
- **“A IA respondeu fora do formato”** → acontece às vezes. Só clique no botão de novo.
- **Sumiu tudo!** → a página ficou fechada ou parada muito tempo. Abra o seu **projeto salvo (.json)**. Por isso é importante salvar sempre.
- **A planilha não foi lida** → no Excel, vá em *Arquivo → Salvar como → CSV* e suba o arquivo .csv.
- **O link do WhatsApp abriu o número errado** → confira se o telefone tem **DDD** (ex.: 11 98888-7777). Números de outros países: comece com **+** (ex.: +351 912 345 678).
- **O WhatsApp Web não abre a conversa** → entre em web.whatsapp.com antes e deixe conectado; depois clique no link de novo.
- **A arte ficou com o texto pequeno** → o texto está longo. Encurte o **Texto da arte** (até 7 palavras) e o **Subtítulo**.
- **Escolhi “Foto destaque” e não apareceu foto** → envie as fotos em **📸 Fotos do produto**; sem foto, o app usa o modelo Impacto.
- **Emojis não aparecem na arte** → é normal: a arte usa só letras para ficar nítida. Os emojis continuam na legenda.
"""),
    ("📖 Dicionário rápido (palavras do marketing)", """
- **Briefing:** resumo do que você vende e para quem.
- **Persona:** um retrato do seu cliente ideal (idade, rotina, desejos).
- **Funil de vendas:** as etapas do “não te conheço” até o “comprei e indico”.
- **Lead:** pessoa interessada que deixou o contato.
- **Isca digital:** algo grátis que você oferece em troca do contato (e-book, cupom, aula).
- **CTA (chamada para ação):** o convite do botão — *“Peça agora”*, *“Quero saber mais”*.
- **Objeção:** o motivo que faz a pessoa hesitar (*“está caro”*, *“não tenho tempo”*).
- **Taxa de resposta:** de cada 100 mensagens enviadas, quantas pessoas responderam.
- **Conversão:** de cada 100 contatos, quantos compraram.
- **Hashtag:** palavra com # que ajuda as pessoas a encontrarem seu post.
"""),
]


def guia_txt() -> str:
    linhas = ["=" * 64, "📘 CAMPANHA 360 IA — GUIA COMPLETO PARA INICIANTES", "=" * 64, ""]
    for titulo, corpo in GUIA:
        linhas += [titulo.upper(), "-" * 64, limpar_md(corpo).replace("`", ""), ""]
    linhas.append(f"Mais ferramentas: {SITE}")
    return "\n".join(linhas)


def mostrar_guia():
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("## 📘 Como usar — guia completo para iniciantes")
        st.caption("Nunca fez campanha? Sem problema. Abra cada item na ordem e siga o passo a passo.")
        for titulo, corpo in GUIA:
            with st.expander(titulo):
                st.markdown(corpo)
        st.download_button("📥 Baixar este guia (.txt)", data=guia_txt().encode("utf-8"),
                           file_name="guia_campanha_360.txt", mime="text/plain", key="dl_guia")


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
    mostrar_guia()
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
    with st.expander("💾 Projeto: salvar, abrir, nova campanha ou sair", expanded=False):
        st.download_button("💾 Salvar projeto (.json)", data=projeto_json(),
                           file_name=f"campanha_{re.sub(r'[^a-z0-9]+', '_', nome_neg.lower())[:30]}.json",
                           mime="application/json", width="stretch")
        arq = st.file_uploader("📂 Abrir projeto salvo", type=["json"], key=f"c_upload_proj_{S.ver}")
        if arq is not None and st.button("Carregar este projeto", width="stretch"):
            try:
                dados = json.loads(arq.getvalue().decode("utf-8"))
                for k in ["briefing", "estrategia", "funil", "posts", "mensagens", "contatos",
                          "marca_cor1", "marca_cor2", "marca_logo", "fotos", "analise_final"]:
                    if k in dados:
                        S[k] = dados[k]
                S.proxima_msg = {}
                S.ver += 1
                st.success("Projeto carregado!")
                st.rerun()
            except Exception:
                st.error("Esse arquivo não é um projeto válido da Campanha 360.")
        st.divider()
        if not S.get("_confirma_nova"):
            if st.button("🆕 Começar campanha nova", width="stretch", type="primary"):
                S["_confirma_nova"] = True
                st.rerun()
        else:
            st.warning("A campanha atual será apagada da tela. **Salve o projeto antes** se quiser guardar "
                       "(botão 💾 acima). Continuar?")
            n1, n2 = st.columns(2)
            with n1:
                if st.button("✅ Sim, nova campanha", width="stretch", key="b_nova_sim"):
                    ver = S.ver
                    for k, v in PADRAO.items():
                        if k not in ("c_nome_user", "c_chave"):
                            S[k] = v.copy() if isinstance(v, (dict, list)) else v
                    S.ver = ver + 1
                    S["_confirma_nova"] = False
                    st.rerun()
            with n2:
                if st.button("✖ Cancelar", width="stretch", key="b_nova_nao"):
                    S["_confirma_nova"] = False
                    st.rerun()
        st.divider()
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
            perfil = st.text_input("@ do Instagram (aparece nas artes)", value=b.get("perfil", ""), placeholder="@docesabor")
        if st.form_submit_button("💾 Salvar briefing", type="primary"):
            S.briefing = {"negocio": negocio, "produto": produto, "preco": preco, "publico": publico,
                          "diferencial": diferencial, "objetivo": objetivo, "duracao": duracao,
                          "canais": ", ".join(canais), "tom": tom, "oferta": oferta, "link": link, "cidade": cidade,
                          "perfil": ("@" + perfil.strip().lstrip("@")) if perfil.strip() else ""}
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
        logo = st.file_uploader("Logo (PNG, opcional)", type=["png", "jpg", "jpeg"], key=f"c_logo_{S.ver}")
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
                S.ver += 1
        if S.estrategia:
            st.markdown(f'<div class="card">', unsafe_allow_html=True)
            st.markdown(S.estrategia)
            st.markdown("</div>", unsafe_allow_html=True)
            with st.expander("✏️ Editar texto da estratégia"):
                novo = st.text_area("Estratégia", S.estrategia, height=300, key=f"ed_estr_{S.ver}", label_visibility="collapsed")
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
                S.ver += 1
        if S.funil:
            st.markdown(S.funil)
            with st.expander("✏️ Editar texto do funil"):
                novo = st.text_area("Funil", S.funil, height=300, key=f"ed_funil_{S.ver}", label_visibility="collapsed")
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
        with st.expander("📸 Fotos do produto (deixam as artes muito mais bonitas)", expanded=not S.fotos):
            st.caption("Envie de 1 a 6 fotos do seu produto, loja ou serviço. Os modelos **Foto destaque**, "
                       "**Foto + painel** e **Elegante** usam essas fotos. Dica: fotos bem iluminadas, na vertical ou quadradas.")
            novas = st.file_uploader("Fotos", type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True,
                                     key=f"c_fotos_{S.ver}", label_visibility="collapsed")
            f1, f2 = st.columns(2)
            with f1:
                if novas and st.button("➕ Usar estas fotos", key="b_add_fotos", width="stretch"):
                    for arq in novas[:6]:
                        try:
                            from PIL import ImageOps
                            im = ImageOps.exif_transpose(Image.open(arq)).convert("RGB")
                            im.thumbnail((1600, 1600))
                            buf = io.BytesIO()
                            im.save(buf, format="JPEG", quality=86)
                            S.fotos.append(base64.b64encode(buf.getvalue()).decode())
                        except Exception:
                            st.warning(f"Não consegui abrir {arq.name}.")
                    S.fotos = S.fotos[-6:]
                    st.rerun()
            with f2:
                if S.fotos and st.button("🗑️ Remover todas as fotos", key="b_rm_fotos", width="stretch"):
                    S.fotos = []
                    st.rerun()
            if S.fotos:
                cols_f = st.columns(6)
                for k, fb in enumerate(S.fotos):
                    cols_f[k].image(base64.b64decode(fb), caption=f"Foto {k + 1}")

        p1, p2, p3, p4 = st.columns([1, 1.3, 1.6, 1.2])
        with p1:
            qtd = st.selectbox("Quantos posts", [3, 6, 9, 12], index=1, key="qtd_posts")
        with p2:
            modelo_arte = st.selectbox("Modelo da arte", ["🎲 Variado"] + MODELOS_ARTE, key="modelo_arte",
                                       help="Variado escolhe um modelo diferente para cada post, de acordo com o objetivo dele.")
        with p3:
            formato_arte = st.selectbox("Formato", list(FORMATOS_ARTE), key="formato_arte")
        with p4:
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
"etiqueta":"1 ou 2 palavras para a etiqueta no topo da arte, ex.: NOVIDADE, DICA, SÓ HOJE, LANÇAMENTO",
"titulo_arte":"frase curta e forte para a arte, máximo 7 palavras, sem emoji",
"subtitulo":"complemento da arte, máximo 12 palavras, sem emoji",
"selo":"selo de destaque com no máximo 3 palavras ou um número (ex.: -20%, FRETE GRÁTIS, 5 POR R$50, NOVO). Use apenas dados do briefing; deixe vazio se não houver oferta real",
"cta":"chamada curta do botão, máximo 3 palavras",
"legenda":"legenda completa e envolvente, com quebras de linha e emojis",
"hashtags":"8 a 12 hashtags relevantes"}}]}}""")
            if dados and isinstance(dados.get("posts"), list):
                S.posts = [p for p in dados["posts"] if isinstance(p, dict)][:qtd]
                S.ver += 1
                st.success(f"✅ {len(S.posts)} posts criados!")

        if S.posts:
            marca_nome = S.briefing.get("negocio", "")
            artes = [gerar_arte(p, modelo_do_post(i, p, modelo_arte), marca_nome, formato_arte, i)
                     for i, p in enumerate(S.posts)]
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                for i, (p, png) in enumerate(zip(S.posts, artes), 1):
                    z.writestr(f"post_{i:02d}.png", png)
                    z.writestr(f"post_{i:02d}_legenda.txt", f"{p.get('legenda', '')}\n\n{p.get('hashtags', '')}")
            st.download_button("📦 Baixar todas as artes + legendas (.zip)", data=zip_buf.getvalue(),
                               file_name="posts_campanha.zip", mime="application/zip")
            opcoes_foto = ["Automática"] + [f"Foto {k + 1}" for k in range(len(S.fotos))] + ["Sem foto"]
            for i, p in enumerate(S.posts):
                with st.container(border=True):
                    a, b_ = st.columns([1, 1.6])
                    with a:
                        st.image(artes[i])
                        st.download_button("⬇️ Baixar arte", data=artes[i], file_name=f"post_{i + 1:02d}.png",
                                           mime="image/png", key=f"dl_arte_{i}", width="stretch")
                    with b_:
                        st.markdown(f"**{i + 1}. {p.get('tema', '')}**  ·  `{p.get('formato', '')}`  ·  _{p.get('etapa', '')}_")
                        m1, m2 = st.columns(2)
                        with m1:
                            atual = modelo_do_post(i, p, modelo_arte)
                            escolhido = st.selectbox("Modelo desta arte", MODELOS_ARTE,
                                                     index=MODELOS_ARTE.index(atual), key=f"pm_{S.ver}_{i}")
                            if escolhido != atual:
                                p["modelo"] = escolhido
                                st.rerun()
                        with m2:
                            idx_atual = p.get("foto_idx")
                            pos = 0 if idx_atual is None else (len(opcoes_foto) - 1 if idx_atual == -1 else idx_atual + 1)
                            pos = min(pos, len(opcoes_foto) - 1)
                            esc_f = st.selectbox("Foto", opcoes_foto, index=pos, key=f"pf_{S.ver}_{i}",
                                                 disabled=not S.fotos, help="Envie fotos no topo desta aba.")
                            novo_idx = None if esc_f == "Automática" else (-1 if esc_f == "Sem foto" else opcoes_foto.index(esc_f) - 1)
                            if novo_idx != idx_atual and S.fotos:
                                p["foto_idx"] = novo_idx
                                st.rerun()
                        p["titulo_arte"] = st.text_input("Texto da arte", p.get("titulo_arte", ""), key=f"pt_{S.ver}_{i}")
                        p["subtitulo"] = st.text_input("Subtítulo", p.get("subtitulo", ""), key=f"ps_{S.ver}_{i}")
                        c_a, c_b, c_c = st.columns(3)
                        with c_a:
                            p["selo"] = st.text_input("Selo (ex.: -20%)", p.get("selo", ""), key=f"psl_{S.ver}_{i}")
                        with c_b:
                            p["etiqueta"] = st.text_input("Etiqueta do topo", p.get("etiqueta", ""), key=f"pe_{S.ver}_{i}")
                        with c_c:
                            p["cta"] = st.text_input("Chamada do botão", p.get("cta", ""), key=f"pc_{S.ver}_{i}")
                        p["legenda"] = st.text_area("Legenda", p.get("legenda", ""), height=130, key=f"pl_{S.ver}_{i}")
                        p["hashtags"] = st.text_input("Hashtags", p.get("hashtags", ""), key=f"ph_{S.ver}_{i}")
                        st.caption("Edite qualquer campo e a arte se atualiza sozinha. Deixe o selo vazio para não mostrar.")


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
                S.ver += 1
                if sair:
                    for m in S.mensagens:
                        if "SAIR" not in m.get("texto", ""):
                            m["texto"] = m.get("texto", "").rstrip() + "\n\nSe não quiser receber mais mensagens, responda SAIR."
                st.success(f"✅ {len(S.mensagens)} mensagens criadas!")

        if S.mensagens:
            for i, m in enumerate(S.mensagens):
                with st.expander(f"{i + 1}. {m.get('etapa', '')} — {m.get('quando', '')}"):
                    m["texto"] = st.text_area("Mensagem", m.get("texto", ""), height=140, key=f"mt_{S.ver}_{i}")

        # ── contatos
        st.markdown("#### 👥 Lista de contatos")
        with st.expander("➕ Adicionar contatos (colar lista ou subir planilha)", expanded=not S.contatos):
            st.caption("Uma pessoa por linha, no formato **Nome, Telefone** (com DDD). Ex.: `Maria Souza, 11 98888-7777`")
            colado = st.text_area("Colar lista", height=120, key=f"c_colar_{S.ver}", label_visibility="collapsed",
                                  placeholder="Maria Souza, 11 98888-7777\nJoão Lima, (21) 97777-6666")
            planilha = st.file_uploader("…ou subir CSV/Excel com colunas Nome e Telefone", type=["csv", "xlsx"], key=f"c_planilha_{S.ver}")
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
            frase = st.text_input("Mensagem que a pessoa vai enviar", key=f"c_frase_{S.ver}",
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

mostrar_guia()

st.markdown(f'<div class="rodape">Os textos são gerados por IA: revise antes de publicar ou enviar.<br>'
            f'© 2026 <a href="{SITE}" target="_blank" style="color:#9CA3AF;">MeusAgentesIA.com.br</a></div>',
            unsafe_allow_html=True)
