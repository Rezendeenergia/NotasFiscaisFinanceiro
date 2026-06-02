import streamlit as st
import pdfplumber
import zipfile
import io
import re
import unicodedata
from datetime import datetime

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
:root {
    --laranja: #F7931E; --laranja-vivo: #FF8C00; --laranja-suave: #FFF4E8;
    --laranja-borda: #FFD39A; --cinza-claro: #F4F4F8; --cinza-borda: #E0E0EB;
    --branco: #FFFFFF; --texto: #1C1C2E; --texto-leve: #6B6B8A;
    --sucesso: #1DB954; --erro: #E53E3E; --aviso: #F59E0B; --radius: 12px;
    --shadow: 0 4px 24px rgba(247,147,30,0.10);
    --dup-bg: #FFF8F0; --dup-borda: #FBBF24;
}
html, body, [class*="css"] { font-family: 'Sora', sans-serif !important; }
.main .block-container { padding: 2rem 3rem 3rem 3rem; max-width: 1100px; }
.hero-header {
    background: linear-gradient(135deg, #1A1A2E 0%, #2D2D44 60%, #3D2A10 100%);
    border-radius: 18px; padding: 2.2rem 2.5rem; margin-bottom: 2rem;
    border: 1px solid #3A3A5C; position: relative; overflow: hidden;
}
.hero-header::before { content:''; position:absolute; top:-60px; right:-60px; width:220px; height:220px;
    background:radial-gradient(circle,rgba(247,147,30,.18) 0%,transparent 70%); border-radius:50%; }
.hero-header::after  { content:''; position:absolute; bottom:-40px; left:-40px; width:160px; height:160px;
    background:radial-gradient(circle,rgba(255,140,0,.10) 0%,transparent 70%); border-radius:50%; }
.hero-title { color:#FFF; font-size:1.9rem; font-weight:700; margin:0 0 .4rem 0; letter-spacing:-.5px; }
.hero-title span { color:var(--laranja); }
.hero-subtitle { color:#A0A0C0; font-size:.95rem; margin:0; line-height:1.6; }
.hero-badges { display:flex; gap:.5rem; margin-top:1rem; flex-wrap:wrap; }
.badge { background:rgba(247,147,30,.15); border:1px solid rgba(247,147,30,.35);
    color:var(--laranja); font-size:.78rem; font-weight:600; padding:.25rem .7rem; border-radius:20px; }
.section-label { display:flex; align-items:center; gap:.5rem; font-size:.8rem; font-weight:700;
    color:var(--laranja-vivo); letter-spacing:1.2px; text-transform:uppercase; margin-bottom:.6rem; }
.section-title { font-size:1.15rem; font-weight:600; color:var(--texto); margin:0 0 1rem 0; }
.metrics-grid { display:grid; grid-template-columns:repeat(5,1fr); gap:1rem; margin:1.2rem 0; }
.metric-card { background:var(--branco); border:1px solid var(--cinza-borda); border-radius:var(--radius);
    padding:1.2rem 1.4rem; text-align:center; transition:all .25s; position:relative; overflow:hidden; }
.metric-card::before { content:''; position:absolute; top:0; left:0; right:0; height:3px;
    background:linear-gradient(90deg,var(--laranja),var(--laranja-vivo)); border-radius:var(--radius) var(--radius) 0 0; }
.metric-card.aviso::before { background:linear-gradient(90deg,var(--aviso),#F97316); }
.metric-card:hover { box-shadow:var(--shadow); transform:translateY(-2px); }
.metric-value { font-size:1.8rem; font-weight:700; color:var(--texto); line-height:1;
    margin-bottom:.3rem; font-family:'JetBrains Mono',monospace; }
.metric-value.laranja { color:var(--laranja-vivo); }
.metric-value.sucesso { color:var(--sucesso); }
.metric-value.erro    { color:var(--erro); }
.metric-value.aviso   { color:var(--aviso); }
.metric-label { font-size:.78rem; font-weight:500; color:var(--texto-leve); letter-spacing:.5px; text-transform:uppercase; }
.dup-banner {
    background: linear-gradient(135deg, #FFF8E1 0%, #FFF3CD 100%);
    border: 2px solid var(--dup-borda);
    border-left: 5px solid #F59E0B;
    border-radius: var(--radius);
    padding: 1.2rem 1.5rem;
    margin: 1rem 0;
}
.dup-banner-title {
    font-size: 1rem; font-weight: 700; color: #92400E;
    display: flex; align-items: center; gap: .5rem; margin-bottom: .8rem;
}
.dup-pair {
    background: white;
    border: 1px solid var(--dup-borda);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin-bottom: .8rem;
}
.dup-pair-header { font-size:.8rem; font-weight:700; color:#92400E; text-transform:uppercase;
    letter-spacing:.8px; margin-bottom:.6rem; }
.dup-fields { display:grid; grid-template-columns:repeat(4,1fr); gap:.5rem; }
.dup-field { background:#FFF8F0; border-radius:6px; padding:.4rem .7rem; }
.dup-field-label { font-size:.7rem; color:var(--texto-leve); font-weight:600; text-transform:uppercase; }
.dup-field-value { font-size:.85rem; color:var(--texto); font-weight:600; font-family:'JetBrains Mono',monospace; }
.stButton > button {
    background:linear-gradient(135deg,var(--laranja) 0%,var(--laranja-vivo) 100%) !important;
    color:white !important; border:none !important; border-radius:var(--radius) !important;
    font-family:'Sora',sans-serif !important; font-weight:600 !important; font-size:.95rem !important;
    padding:.7rem 1.5rem !important; box-shadow:0 4px 16px rgba(247,147,30,.30) !important; transition:all .2s !important; }
.stButton > button:hover { box-shadow:0 6px 24px rgba(247,147,30,.45) !important; transform:translateY(-1px) !important; }
.stDownloadButton > button {
    background:linear-gradient(135deg,#1A1A2E 0%,#2D2D44 100%) !important;
    color:white !important; border:1px solid var(--laranja) !important; border-radius:var(--radius) !important;
    font-family:'Sora',sans-serif !important; font-weight:600 !important; font-size:.95rem !important;
    padding:.7rem 1.5rem !important; box-shadow:0 4px 16px rgba(0,0,0,.15) !important; transition:all .2s !important; }
.stDownloadButton > button:hover { box-shadow:0 6px 24px rgba(247,147,30,.25) !important; transform:translateY(-1px) !important; }
[data-testid="stFileUploader"] { background:var(--laranja-suave) !important;
    border:2px dashed var(--laranja-borda) !important; border-radius:var(--radius) !important; padding:.5rem !important; }
[data-testid="stFileUploader"]:hover { border-color:var(--laranja-vivo) !important; }
.stProgress > div > div > div > div { background:linear-gradient(90deg,var(--laranja),var(--laranja-vivo)) !important; border-radius:4px !important; }
[data-testid="stDataFrame"] { border:1px solid var(--cinza-borda) !important; border-radius:var(--radius) !important; overflow:hidden !important; }
hr { border-color:var(--cinza-borda) !important; margin:1.5rem 0 !important; }
.stCaption { color:var(--texto-leve) !important; font-size:.82rem !important; text-align:center !important; }
::-webkit-scrollbar { width:6px; height:6px; }
::-webkit-scrollbar-track { background:var(--cinza-claro); border-radius:3px; }
::-webkit-scrollbar-thumb { background:var(--laranja-borda); border-radius:3px; }
::-webkit-scrollbar-thumb:hover { background:var(--laranja); }
@media (max-width:768px) {
    .main .block-container { padding:1rem 1.2rem; }
    .metrics-grid { grid-template-columns:repeat(2,1fr); }
    .hero-title { font-size:1.4rem; }
    .dup-fields { grid-template-columns:repeat(2,1fr); }
}
</style>
"""


# =============================================================================
#  UTILITARIOS
# =============================================================================

def ascii_normalizar(txt):
    """Remove acentos e converte para ASCII puro (facilita regex)."""
    return unicodedata.normalize('NFKD', txt).encode('ascii', 'ignore').decode('ascii')


def normalizar_data(raw):
    """Converte qualquer formato de data para DD-MM-YYYY."""
    raw = raw.strip()
    m = re.match(r'^(\d{4})[/-](\d{2})[/-](\d{2})$', raw)
    if m:
        return "{}-{}-{}".format(m.group(3), m.group(2), m.group(1))
    m = re.match(r'^(\d{2})[/-](\d{2})[/-](\d{4})$', raw)
    if m:
        return "{}-{}-{}".format(m.group(1), m.group(2), m.group(3))
    m = re.match(r'^(\d{2})[/-](\d{2})[/-](\d{2})$', raw)
    if m:
        ano = int(m.group(3))
        return "{}-{}-{}".format(m.group(1), m.group(2), 2000 + ano if ano <= 50 else 1900 + ano)
    return None


def candidato_valido(c):
    """Verifica se o candidato a emitente parece um nome de empresa valido."""
    c = c.strip()
    if len(c) < 3:
        return False
    if re.match(r'^[\d.\-/\s]+$', c):
        return False
    lixo = ['AVENIDA', 'RUA ', 'CEP', 'CNPJ', 'CPF', 'FONE', 'BAIRRO', 'MUNICIPIO',
            'DOCUMENTO', 'AUXILIAR', 'ELETRONICA', 'ENTRADA', 'SAIDA', 'FOLHA',
            'NATUREZA', 'PROTOCOLO', 'INSCRICAO', 'VENDA', 'RECEBEMOS', 'RECEBI',
            'DANFE', 'ABAIXO', 'SERIE', 'ENDERECO', 'E-MAIL', 'EMAIL',
            'TV ', 'QD.', 'QUADRA', 'LOTE', 'TRAVESSA', 'ALAMEDA']
    cu = ascii_normalizar(c).upper()
    return not any(cu.startswith(l) for l in lixo)


def extrair_numero_nf(texto):
    """
    Extrai o numero da NF/NFS-e corretamente.
    Suporta: NF-e DANFE, Omie, DANFSe Santarem, NFSe prefeituras.
    """
    # 1. DANFSe Santarem
    m = re.search(r'NumerodaNFS-e\s+\S+\s+\S+\s*\n(\d+)\s+', texto, re.IGNORECASE)
    if m:
        try: return str(int(m.group(1)))
        except ValueError: pass

    # 2. NFSe Belem/prefeituras
    m = re.search(r'Numero\s*/\s*Serie[^\n]*\n[^\n]*?(\d{2,})\s*/\s*[A-Z]', texto, re.IGNORECASE)
    if m:
        try: return str(int(m.group(1)))
        except ValueError: pass

    # 3. NFSe generica
    m = re.search(r'Numero\s+da\s+NFS-?e\s*\n(\d+)', texto, re.IGNORECASE)
    if m:
        try: return str(int(m.group(1)))
        except ValueError: pass

    # 4. DANFE NF-e
    m = re.search(
        r'(?:^|\n)\s*N[o.]?\s*\.?\s*([\d. ]+)\s*\n\s*(?:Serie|SERIE|DATA|Folha|FOLHA)',
        texto, re.IGNORECASE | re.MULTILINE
    )
    if m:
        raw = m.group(1).replace('.', '').replace(' ', '')
        try: return str(int(raw))
        except ValueError: pass

    # 5. Goianesia NFS-e
    m = re.search(r'(?:^|\n)\s*No\s+(\d+)\s*\n\s*(?:PAGINA|NF-e\s+Emitida)', texto, re.IGNORECASE | re.MULTILINE)
    if m:
        try: return str(int(m.group(1)))
        except ValueError: pass

    # 6. Fallback
    m = re.search(r'N[o.]+\s*([\d.]{1,12})(?:\s|$)', texto, re.IGNORECASE | re.MULTILINE)
    if m:
        raw = m.group(1).replace('.', '').replace(' ', '')
        if len(raw) <= 9:
            try: return str(int(raw))
            except ValueError: pass

    return None


def limpar_nome(nome):
    """Remove caracteres invalidos e normaliza para maiusculas."""
    nome = re.sub(r'[<>:"/\\|?*]', '', nome)
    return ' '.join(nome.split()).upper()


def montar_nome(numero_nf, emitente):
    """Monta o nome final: 'NF 202001 - EMITENTE.PDF'"""
    return "NF {} - {}.PDF".format(numero_nf, emitente)


# =============================================================================
#  EXTRACAO DE DADOS DA NOTA FISCAL
# =============================================================================

def _ascii(txt):
    """Helper interno: remove acentos e converte para ASCII."""
    return unicodedata.normalize('NFKD', txt).encode('ascii', 'ignore').decode('ascii')


def _extrair_emitente_por_bbox(pdf_bytes):
    """
    Extrai o nome do emitente usando posicao (bbox) com x_tolerance=1.
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[0]
            words = page.extract_words(x_tolerance=1, y_tolerance=3)

        blocos = [
            (['EMITENTE', 'NFS'], ['TOMADOR'], ['NOME', 'EMPRESARIAL']),
            (['EMITENTE', 'PRESTADOR'], ['TOMADOR'], ['NOME', 'EMPRESARIAL']),
            (['IDENTIFICACAO', 'EMITENTE'], ['DESTINATARIO'], ['NOME', 'EMPRESARIAL']),
            (['IDENTIFICACAO', 'EMITENTE'], ['DESTINATARIO'], ['RAZAO', 'SOCIAL']),
        ]

        def achar_y_subsequencia(words_list, palavras, y_min=0, y_max=9999):
            linhas = {}
            for w in words_list:
                if not (y_min <= w['top'] <= y_max):
                    continue
                y_key = round(w['top'], 0)
                linhas.setdefault(y_key, []).append(_ascii(w['text']).upper())
            for y_key in sorted(linhas.keys()):
                tokens = ' '.join(linhas[y_key])
                if all(p in tokens for p in palavras):
                    return float(y_key)
            return None

        def palavras_na_linha(words_list, y_ref, y_offset_min=3, y_offset_max=22, x_max=420):
            resultado = []
            for w in words_list:
                if y_ref + y_offset_min <= w['top'] <= y_ref + y_offset_max and w['x0'] < x_max:
                    t = _ascii(w['text'])
                    if '@' in t:
                        break
                    if re.match(r'^\d{2}\.\d{3}', t):
                        break
                    if re.match(r'^[\d./@\-()+]+$', t):
                        continue
                    resultado.append(t)
            return resultado

        for palavras_inicio, palavras_fim, palavras_label in blocos:
            y_inicio = achar_y_subsequencia(words, palavras_inicio)
            if y_inicio is None:
                continue
            y_fim = achar_y_subsequencia(words, palavras_fim, y_min=y_inicio + 5) or 9999
            y_label = achar_y_subsequencia(words, palavras_label, y_min=y_inicio, y_max=y_fim)
            if y_label is None:
                continue
            nome_words = palavras_na_linha(words, y_label, x_max=420)
            if nome_words:
                nome = ' '.join(nome_words).strip()
                if nome and len(nome) >= 3:
                    return nome

    except Exception:
        pass
    return None


def _extrair_tomador_por_bbox(pdf_bytes):
    """
    Extrai o nome do tomador/destinatario usando posicao (bbox).
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            page = pdf.pages[0]
            words = page.extract_words(x_tolerance=1, y_tolerance=3)

        def agrupar_linhas(words_list, y_min=0, y_max=9999):
            linhas = {}
            for w in words_list:
                if not (y_min <= w['top'] <= y_max):
                    continue
                y_key = round(w['top'], 0)
                linhas.setdefault(y_key, []).append(w)
            return dict(sorted(linhas.items()))

        def achar_y_bloco(linhas_dict, palavras):
            for y_key, ws in linhas_dict.items():
                tokens = ' '.join(_ascii(w['text']).upper() for w in ws)
                if all(p in tokens for p in palavras):
                    return float(y_key)
            return None

        def nome_apos_label(linhas_dict, y_label, x_label_max=90, y_fim=9999):
            y_key = round(y_label, 0)
            if y_key not in linhas_dict:
                return []
            resultado = []
            for w in linhas_dict[y_key]:
                if w['x0'] > x_label_max:
                    t = _ascii(w['text'])
                    if '@' in t: break
                    if re.match(r'^[\d./@\-()+:]+$', t): continue
                    resultado.append(t)
            return resultado

        def nome_linha_seguinte(linhas_dict, y_label, y_fim=9999, x_max=450):
            ys = sorted(k for k in linhas_dict if k > y_label + 2 and k < y_fim)
            if not ys:
                return []
            y_next = ys[0]
            resultado = []
            for w in linhas_dict[y_next]:
                if w['x0'] < x_max:
                    t = _ascii(w['text'])
                    if '@' in t: break
                    if re.match(r'^\d{2}\.\d{3}', t): break
                    if re.match(r'^[\d./@\-()+:]+$', t): continue
                    resultado.append(t)
            return resultado

        linhas = agrupar_linhas(words)

        blocos = [
            (['TOMADOR'], ['NOME', 'RAZAO']),
            (['DESTINATARIO'], ['NOME', 'RAZAO']),
            (['DESTINATARIO'], ['RAZAO', 'SOCIAL']),
            (['TOMADOR'], ['NOME', 'EMPRESARIAL']),
        ]

        for palavras_bloco, palavras_label in blocos:
            y_bloco = achar_y_bloco(linhas, palavras_bloco)
            if y_bloco is None:
                continue
            y_label = achar_y_bloco(
                {k: v for k, v in linhas.items() if k > y_bloco + 2},
                palavras_label
            )
            if y_label is None:
                continue
            nome_words = nome_apos_label(linhas, y_label, x_label_max=80)
            if not nome_words:
                nome_words = nome_linha_seguinte(linhas, y_label)
            if nome_words:
                nome = ' '.join(nome_words).strip()
                if nome and len(nome) >= 3 and candidato_valido(nome):
                    return nome

    except Exception:
        pass
    return None


def extrair_info_nota_fiscal(pdf_bytes):
    """
    Extrai numero, emitente, tomador, tipo, data e valor de um PDF de NF.
    Retorna: (data, emitente, tipo_nf, valor, numero_nf, tomador)
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto_raw = ""
            for page in pdf.pages[:2]:
                t = page.extract_text()
                if t:
                    texto_raw += t + "\n"

        if not texto_raw.strip():
            return None, None, None, None, None, None

        texto = ascii_normalizar(texto_raw)
        texto_up = texto.upper()

        data = None
        emitente = None
        tomador = None
        tipo_nf = None
        valor = None
        numero_nf = None

        # ── Tipo ──────────────────────────────────────────────────────────
        if "NFS-E" in texto_up or "NOTA FISCAL DE SERVICO" in texto_up or "NOTA FISCAL DE SERVI" in texto_up or "DANFSE" in texto_up:
            tipo_nf = "NFS-e"
        elif "DANFE" in texto_up or "DOCUMENTO AUXILIAR DA NOTA FISCAL" in texto_up:
            tipo_nf = "NF-e"
        elif "NOTA FISCAL" in texto_up:
            tipo_nf = "NF"
        elif "CUPOM FISCAL" in texto_up or "CF-E" in texto_up:
            tipo_nf = "CF-e"
        else:
            tipo_nf = "Documento"

        # ── Numero da NF ──────────────────────────────────────────────────
        numero_nf = extrair_numero_nf(texto)

        # ── Data de Emissao ───────────────────────────────────────────────
        for p in [
            r'EMISSAO:\s*(\d{2}[/\-]\d{2}[/\-]\d{4})',
            r'DATA\s+DA\s+EMISSAO\s+(\d{2}[/\-]\d{2}[/\-]\d{4})',
            r'NF-e\s+Emitida\s+em:\s*(\d{2}[/\-]\d{2}[/\-]\d{4})',
            r'Data\s+emissao\s+(\d{2}[/\-]\d{2}[/\-]\d{4})',
            r'EMISSAO\s*[:\s]+(\d{2}[/\-]\d{2}[/\-]\d{4})',
            r'Data\s+de\s+Competencia\s*[:\s]+(\d{2}[/\-]\d{2}[/\-]\d{4})',
            r'(\d{4}-\d{2}-\d{2})',
            r'(\d{2}[/\-]\d{2}[/\-]\d{4})',
        ]:
            m = re.search(p, texto, re.IGNORECASE)
            if m:
                data = normalizar_data(m.group(1))
                if data:
                    break

        # ── Emitente ─────────────────────────────────────────────────────
        emitente_bbox = _extrair_emitente_por_bbox(pdf_bytes)
        if emitente_bbox and candidato_valido(emitente_bbox):
            emitente = emitente_bbox

        if not emitente:
            m = re.search(r'TOMADOR\s+DE\s+SERVICOS.*?Nome/Razao:\s*([^\n]+)', texto, re.IGNORECASE | re.DOTALL)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        if not emitente:
            m = re.search(r'RECEBEMOS\s+DE\s+(.+?)\s+OS\s+PRODUTOS', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        if not emitente:
            m = re.search(r'RECEBI\(EMOS\)\s+DE\s+(.+?),\s+OS\s+PRODUTOS', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        if not emitente:
            m = re.search(
                r'IDENTIFICACAO\s+DO\s+EMITENTE[^\n]*\n([A-Z][A-Z0-9 &.,/\-]{2,70}?)\s+(?:Eletro|DANFE|0\s+-|1\s+-)',
                texto, re.IGNORECASE
            )
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        if not emitente and "DANFSE" in texto_up or "DANFSEV" in texto_up:
            c = _extrair_emitente_por_bbox(pdf_bytes)
            if c and candidato_valido(c):
                emitente = c
            else:
                m = re.search(
                    r'EMITENTEDANFS-e.*?Nome/NomeEmpresarial[^\n]*\n([^\n]{3,80})',
                    texto, re.IGNORECASE | re.DOTALL
                )
                if m:
                    c = re.split(r'\s+\S+@\S+', m.group(1))[0].strip()
                    if candidato_valido(c):
                        emitente = c

        if not emitente:
            m = re.search(
                r'EMITENTE\s+PRESTADOR\s+DO\s+SERVICO.*?Nome\s*/\s*Nome\s+Empresarial[^\n]*\n([^\n]{3,80})',
                texto, re.IGNORECASE | re.DOTALL
            )
            if m:
                c = re.split(r'\s+\S+@\S+', m.group(1))[0].strip()
                if candidato_valido(c):
                    emitente = c

        if not emitente:
            m = re.search(r'Prestador\s+de\s+Servicos\s*\n\s*([A-Z][^\n\r]{2,79})', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        if not emitente:
            m = re.search(r'Emitente\s*[:\s]+([^\n\r]{3,80})', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        if not emitente:
            m = re.search(r'Razao\s+Social\s*[:\s]+([^\n\r]{3,80})', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        if not emitente:
            m = re.search(
                r'\n([A-Z][A-Z0-9 &.,/\-]{2,60}(?:S\.A|LTDA|EIRELI|DISTRIBUIDORA|INDUSTRIA|CONSTRUTORA|COMERCIO|ENERGIA)\.?)\n',
                texto
            )
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        # ── Valor Total ───────────────────────────────────────────────────
        for p in [
            r'VALOR\s+TOTAL:\s*R.\s*([\d.,]+)',
            r'VALOR\s+TOTAL\s+DA\s+NOTA\s+([\d.,]+)',
            r'Valor\s+Total\s+da\s+Nota\s*[:\s]*R?.\s*([\d.,]+)',
            r'Valor\s+Total\s*[:\s]*R?.\s*([\d.,]+)',
            r'TOTAL\s+GERAL\s*[:\s]*R?.\s*([\d.,]+)',
            r'=\)\s*Valor\s+liquido\s+R\$\s*([\d.,]+)',
            r'Valor\s+da\s+nota\s+R\$\s*([\d.,]+)',
            r'Valor\s+dos\s+Servicos\s*[:\s]*R?.\s*([\d.,]+)',
        ]:
            m = re.search(p, texto, re.IGNORECASE)
            if m:
                v = m.group(1).strip()
                if re.search(r'\d{1,3}(?:\.\d{3})+,\d{2}$', v):
                    v = v.replace('.', '').replace(',', '.')
                elif re.search(r'\d{1,3}(?:,\d{3})+\.\d{2}$', v):
                    v = v.replace(',', '')
                else:
                    v = v.replace(',', '.')
                try:
                    valor = float(v)
                    break
                except ValueError:
                    continue

        # ── Tomador ───────────────────────────────────────────────────────
        tomador_bbox = _extrair_tomador_por_bbox(pdf_bytes)
        if tomador_bbox and candidato_valido(tomador_bbox):
            tomador = tomador_bbox

        if not tomador:
            for p in [
                r'Tomador\s+de\s+Servicos\s*\n\s*([A-Z][^\n\r]{2,79})',
                r'TOMADOR\s+DE\s+SERVICOS[^\n]*\n[^\n]*\n([A-Z][^\n\r]{2,79})',
                r'DESTINATARIO[^\n]*\n([A-Z][^\n\r]{2,79})',
            ]:
                m = re.search(p, texto, re.IGNORECASE)
                if m and candidato_valido(m.group(1)):
                    tomador = m.group(1).strip()
                    break

        if not tomador:
            m = re.search(r'Nome/Razao[^\n]*\n([A-Z][^\n\r]{2,79})', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                tomador = m.group(1).strip()

        return data, emitente, tipo_nf, valor, numero_nf, tomador

    except Exception as e:
        st.error("Erro ao processar PDF: {}".format(str(e)))
        return None, None, None, None, None, None


# =============================================================================
#  DETECÇÃO DE DUPLICATAS
# =============================================================================

def chave_duplicata(r):
    """
    Gera uma chave de comparação para detectar duplicatas reais.
    Usa: numero_nf + fornecedor_utilizado + data + valor (arredondado).
    """
    num = (r.get('numero_nf') or '').strip().upper()
    fornecedor = (r.get('fornecedor_usado') or '').strip().upper()
    data = (r.get('data') or '').strip()
    # Arredonda valor para 2 casas para evitar falsos negativos por float
    valor = r.get('valor')
    valor_key = "{:.4f}".format(valor) if valor is not None else 'N/A'
    return (num, fornecedor, data, valor_key)


def detectar_duplicatas(resultados):
    """
    Analisa a lista de resultados e retorna:
      - grupos_dup: lista de listas de índices que são duplicatas entre si
      - indices_dup: set de todos os índices que fazem parte de algum grupo duplicado
    Um documento é considerado duplicata quando NF + fornecedor + data + valor coincidem.
    Ignora documentos onde algum campo-chave está ausente (N/A).
    """
    from collections import defaultdict
    chaves = {}
    grupos_por_chave = defaultdict(list)

    for i, r in enumerate(resultados):
        chave = chave_duplicata(r)
        # Só detecta duplicata se tiver pelo menos número e fornecedor
        if chave[0] == '' or chave[0] == 'N/A' or chave[1] == '' or chave[1] == 'N/A':
            continue
        grupos_por_chave[chave].append(i)

    grupos_dup = [indices for indices in grupos_por_chave.values() if len(indices) > 1]
    indices_dup = set(i for grupo in grupos_dup for i in grupo)
    return grupos_dup, indices_dup


# =============================================================================
#  PROCESSAMENTO
# =============================================================================

def processar_pdfs(uploaded_files, modo='emitente'):
    resultados = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uf in enumerate(uploaded_files):
        status_text.text("Processando {}...".format(uf.name))
        pdf_bytes = uf.read()

        data, emitente, tipo_nf, valor, numero_nf, tomador = extrair_info_nota_fiscal(pdf_bytes)

        razao_escolhida = tomador if modo == 'tomador' else emitente

        if numero_nf and razao_escolhida:
            razao_limpa = limpar_nome(razao_escolhida)
            novo_nome = montar_nome(numero_nf, razao_limpa)
            status = '✅ Sucesso'
        else:
            razao_limpa = limpar_nome(razao_escolhida) if razao_escolhida else 'N/A'
            novo_nome = '-'
            status = '⚠️ Informações não encontradas'

        emitente_limpo = limpar_nome(emitente) if emitente else 'N/A'
        tomador_limpo  = limpar_nome(tomador)  if tomador  else 'N/A'

        resultados.append({
            'original':       uf.name,
            'novo_nome':      novo_nome,
            'status':         status,
            'tipo':           tipo_nf or 'Desconhecido',
            'numero_nf':      numero_nf or 'N/A',
            'data':           data or 'N/A',
            'emitente':       emitente_limpo,
            'tomador':        tomador_limpo,
            'fornecedor_usado': razao_limpa if razao_limpa != 'N/A' else (emitente_limpo if emitente_limpo != 'N/A' else tomador_limpo),
            'valor':          valor,
            '_bytes':         pdf_bytes,
            'duplicata':      False,  # será preenchido depois
            'grupo_dup':      None,
        })
        progress_bar.progress((idx + 1) / len(uploaded_files))

    status_text.text("✅ Processamento concluído!")

    # ── Detecção de duplicatas ─────────────────────────────────────────
    grupos_dup, indices_dup = detectar_duplicatas(resultados)
    for grupo_idx, grupo in enumerate(grupos_dup):
        for i in grupo:
            resultados[i]['duplicata'] = True
            resultados[i]['grupo_dup'] = grupo_idx
            resultados[i]['status'] = '⚠️ Possível duplicata'

    # ── Resolve nomes duplicados (contador) apenas para nomes de arquivo ──
    # NÃO aplica contador em notas identificadas como duplicata real —
    # elas ficam com o mesmo nome para evidenciar a duplicidade.
    nomes_counter = {}
    for i in range(len(resultados)):
        nome_orig = resultados[i]['novo_nome']
        if nome_orig == '-':
            continue
        if resultados[i]['duplicata']:
            # Mantém o mesmo nome; o usuário decide o que fazer
            continue
        if nome_orig not in nomes_counter:
            nomes_counter[nome_orig] = 1
        else:
            nomes_counter[nome_orig] += 1
            base, ext = (nome_orig.rsplit('.', 1) if '.' in nome_orig else (nome_orig, ''))
            resultados[i]['novo_nome'] = "{} ({}).{}".format(base, nomes_counter[nome_orig], ext) if ext else "{} ({})".format(base, nomes_counter[nome_orig])

    return resultados, grupos_dup


def montar_zip(pares):
    """Recebe lista de (nome, bytes) e retorna bytes de um ZIP."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for nome, conteudo in pares:
            zf.writestr(nome, conteudo)
    return buf.getvalue()


# =============================================================================
#  METRICAS
# =============================================================================

def render_metricas(resultados, grupos_dup):
    total = len(resultados)
    sucesso = sum(1 for r in resultados if '✅' in r['status'])
    erros = sum(1 for r in resultados if '⚠️ Informações' in r['status'])
    duplicatas = sum(1 for r in resultados if r['duplicata'])
    unicas = total - duplicatas
    valor_total = sum(r.get('valor') or 0 for r in resultados if not r['duplicata'])
    # Para valor, conta apenas 1 de cada grupo de duplicatas
    for grupo in grupos_dup:
        r = resultados[grupo[0]]
        valor_total += r.get('valor') or 0
    valor_fmt = "R$ {:,.2f}".format(valor_total).replace(',', '_').replace('.', ',').replace('_', '.')

    aviso_class = ' aviso' if duplicatas > 0 else ''

    st.markdown("""
    <div class="metrics-grid">
        <div class="metric-card"><div class="metric-value laranja">{total}</div><div class="metric-label">Total Processadas</div></div>
        <div class="metric-card"><div class="metric-value sucesso">{sucesso}</div><div class="metric-label">Sucesso</div></div>
        <div class="metric-card"><div class="metric-value">{unicas}</div><div class="metric-label">Notas Únicas</div></div>
        <div class="metric-card{aviso_class}"><div class="metric-value aviso">{duplicatas}</div><div class="metric-label">Possíveis Duplicatas</div></div>
        <div class="metric-card"><div class="metric-value laranja" style="font-size:1.2rem">{valor}</div><div class="metric-label">Valor Total (únicas)</div></div>
    </div>
    """.format(
        total=total, sucesso=sucesso, unicas=unicas,
        duplicatas=duplicatas, valor=valor_fmt, aviso_class=aviso_class
    ), unsafe_allow_html=True)


# =============================================================================
#  PAINEL DE DUPLICATAS
# =============================================================================

def render_painel_duplicatas(resultados, grupos_dup):
    """
    Exibe o painel de análise de duplicatas com opções de ação para cada grupo.
    Retorna: dict { grupo_idx -> acao } onde acao in ['ignorar', 'manter_primeiro', 'manter_segundo', 'baixar_ambos']
    """
    if not grupos_dup:
        return {}

    st.markdown("""
    <div class="dup-banner">
        <div class="dup-banner-title">⚠️ Possíveis Notas Fiscais Duplicadas Detectadas</div>
        <p style="margin:0; color:#78350F; font-size:.9rem;">
            Foram encontrados <strong>{qtd} grupo(s)</strong> de notas com mesmo número, fornecedor, data e valor.
            Revise abaixo e escolha como proceder para cada grupo antes de baixar os arquivos.
        </p>
    </div>
    """.format(qtd=len(grupos_dup)), unsafe_allow_html=True)

    acoes = {}

    for grupo_idx, grupo in enumerate(grupos_dup):
        r0 = resultados[grupo[0]]
        r1 = resultados[grupo[1]] if len(grupo) > 1 else None

        valor_fmt = "R$ {:,.2f}".format(r0['valor']).replace(',', '_').replace('.', ',').replace('_', '.') if r0.get('valor') else 'N/A'

        st.markdown("""
        <div class="dup-pair">
            <div class="dup-pair-header">🔍 Grupo {} — {} arquivo(s) idêntico(s)</div>
            <div class="dup-fields">
                <div class="dup-field">
                    <div class="dup-field-label">Número NF</div>
                    <div class="dup-field-value">{numero}</div>
                </div>
                <div class="dup-field">
                    <div class="dup-field-label">Fornecedor</div>
                    <div class="dup-field-value">{fornecedor}</div>
                </div>
                <div class="dup-field">
                    <div class="dup-field-label">Data Emissão</div>
                    <div class="dup-field-value">{data}</div>
                </div>
                <div class="dup-field">
                    <div class="dup-field-label">Valor Total</div>
                    <div class="dup-field-value">{valor}</div>
                </div>
            </div>
        </div>
        """.format(
            grupo_idx=grupo_idx + 1,
            qtd=len(grupo),
            numero=r0.get('numero_nf', 'N/A'),
            fornecedor=r0.get('fornecedor_usado', 'N/A')[:35] + ('...' if len(r0.get('fornecedor_usado','')) > 35 else ''),
            data=r0.get('data', 'N/A'),
            valor=valor_fmt,
        ), unsafe_allow_html=True)

        # Arquivos envolvidos
        for i, idx in enumerate(grupo):
            r = resultados[idx]
            st.caption("📄 Arquivo {}: `{}`".format(i + 1, r['original']))

        col1, col2, col3, col4 = st.columns(4)
        chave_sessao = 'dup_acao_{}'.format(grupo_idx)

        with col1:
            if st.button("✅ Ignorar duplicidade", key="ign_{}".format(grupo_idx), use_container_width=True):
                st.session_state[chave_sessao] = 'ignorar'
        with col2:
            label_manter1 = "📄 Manter apenas 1º" if len(grupo) == 2 else "📄 Manter 1º arquivo"
            if st.button(label_manter1, key="man1_{}".format(grupo_idx), use_container_width=True):
                st.session_state[chave_sessao] = 'manter_primeiro'
        with col3:
            if len(grupo) >= 2:
                if st.button("📄 Manter apenas 2º", key="man2_{}".format(grupo_idx), use_container_width=True):
                    st.session_state[chave_sessao] = 'manter_segundo'
            else:
                st.empty()
        with col4:
            if st.button("📥 Baixar os dois", key="amb_{}".format(grupo_idx), use_container_width=True):
                st.session_state[chave_sessao] = 'baixar_ambos'

        acao_atual = st.session_state.get(chave_sessao, None)
        if acao_atual:
            labels = {
                'ignorar':         '✅ Ambos incluídos no download com nomes distintos',
                'manter_primeiro': '📄 Apenas o 1º arquivo será incluído no download',
                'manter_segundo':  '📄 Apenas o 2º arquivo será incluído no download',
                'baixar_ambos':    '📥 Ambos serão incluídos com nomes distintos (com sufixo)',
            }
            st.info(labels.get(acao_atual, ''))
            acoes[grupo_idx] = acao_atual
        else:
            st.warning("⏳ Aguardando sua decisão para este grupo...")

        st.divider()

    return acoes


# =============================================================================
#  CONSTRUÇÃO DO ZIP FINAL
# =============================================================================

def construir_arquivos_download(resultados, grupos_dup, acoes):
    """
    Com base nas ações escolhidas pelo usuário para cada grupo de duplicatas,
    retorna a lista final de (nome, bytes) para download.
    """
    # Índices a excluir do download
    excluir = set()
    # Índices que devem ter sufixo numérico (baixar ambos / ignorar)
    sufixar = {}  # idx -> sufixo string

    for grupo_idx, grupo in enumerate(grupos_dup):
        acao = acoes.get(grupo_idx, 'baixar_ambos')  # padrão: inclui ambos

        if acao == 'manter_primeiro':
            for idx in grupo[1:]:
                excluir.add(idx)
        elif acao == 'manter_segundo':
            excluir.add(grupo[0])
            for idx in grupo[2:]:
                excluir.add(idx)
        elif acao in ('ignorar', 'baixar_ambos'):
            # Inclui todos, mas com sufixo para diferenciar
            for sufixo_n, idx in enumerate(grupo, start=1):
                sufixar[idx] = '_DUP{}'.format(sufixo_n)

    pares = []
    nomes_usados = {}
    for i, r in enumerate(resultados):
        if i in excluir:
            continue
        nome = r['novo_nome'] if r['novo_nome'] != '-' else r['original']
        if i in sufixar:
            base, ext = (nome.rsplit('.', 1) if '.' in nome else (nome, ''))
            nome = "{}{}.{}".format(base, sufixar[i], ext) if ext else "{}{}".format(base, sufixar[i])
        # Evita colisão residual
        if nome in nomes_usados:
            nomes_usados[nome] += 1
            base, ext = (nome.rsplit('.', 1) if '.' in nome else (nome, ''))
            nome = "{} ({}).{}".format(base, nomes_usados[nome], ext) if ext else "{} ({})".format(base, nomes_usados[nome])
        else:
            nomes_usados[nome] = 1
        pares.append((nome, r['_bytes']))
    return pares


# =============================================================================
#  INTERFACE PRINCIPAL
# =============================================================================

def main():
    st.set_page_config(page_title="Notas Fiscais | Rezende Energia", page_icon="🧾", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    st.markdown("""
    <div class="hero-header">
        <div class="hero-title">🧾 Renomeador de <span>Notas Fiscais</span></div>
        <div class="hero-subtitle">
            Processa notas fiscais em PDF e as renomeia automaticamente no padrão<br>
            <strong style="color:#FFD39A;">NF [Numero] &mdash; Razão Social</strong>
            &mdash; com detecção inteligente de <strong style="color:#F7931E;">duplicatas reais</strong>.
        </div>
        <div class="hero-badges">
            <span class="badge">📄 NF-e DANFE</span>
            <span class="badge">⚡ OMIE</span>
            <span class="badge">🛠️ NFS-e</span>
            <span class="badge">🧾 NF SIMPLES</span>
            <span class="badge">🏷️ CF-e / SAT</span>
            <span class="badge">🔍 DETECÇÃO DE DUPLICATAS</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Upload ────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="section-label"><span>01</span> Upload</div>
    <div class="section-title">Envie os PDFs das notas fiscais (pode selecionar vários de uma vez)</div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Arraste ou clique para selecionar os PDFs",
        type=['pdf'],
        accept_multiple_files=True,
        help="Selecione um ou mais PDFs de notas fiscais"
    )

    if not uploaded_files:
        return

    qtd = len(uploaded_files)
    st.success("✅ {} arquivo(s) carregado(s)".format(qtd))

    # ── Modo ──────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="section-label"><span>02</span> Modo de Renomeação</div>
    <div class="section-title">Escolha qual razão social usar no nome do arquivo</div>
    """, unsafe_allow_html=True)

    col1, _ = st.columns(2)
    with col1:
        modo = st.radio(
            "Usar como nome:",
            options=["emitente", "tomador"],
            format_func=lambda x: "🏭 Emitente / Prestador (quem emitiu a NF)" if x == "emitente" else "🏢 Tomador / Destinatário (quem recebeu)",
            index=0,
            label_visibility="collapsed"
        )

    if not st.button("🚀 Processar Notas Fiscais", type="primary", use_container_width=True):
        return

    # ── Processamento ─────────────────────────────────────────────────────
    with st.spinner("Processando notas fiscais..."):
        resultados, grupos_dup = processar_pdfs(uploaded_files, modo=modo)

    st.divider()

    # ── Resultados ────────────────────────────────────────────────────────
    st.markdown("""
    <div class="section-label"><span>03</span> Resultados</div>
    <div class="section-title">Resumo do Processamento</div>
    """, unsafe_allow_html=True)

    render_metricas(resultados, grupos_dup)

    # Tabela — marca duplicatas visualmente
    df_display = []
    for r in resultados:
        linha = {k: v for k, v in r.items() if not k.startswith('_') and k not in ('duplicata', 'grupo_dup', 'fornecedor_usado')}
        if r['duplicata']:
            linha['status'] = '🔴 Possível duplicata'
        df_display.append(linha)

    st.dataframe(df_display, use_container_width=True, hide_index=True,
        column_config={
            'original':  st.column_config.TextColumn('Nome Original'),
            'novo_nome': st.column_config.TextColumn('Novo Nome'),
            'status':    st.column_config.TextColumn('Status'),
            'tipo':      st.column_config.TextColumn('Tipo'),
            'numero_nf': st.column_config.TextColumn('Número NF'),
            'data':      st.column_config.TextColumn('Data de Emissão'),
            'emitente':  st.column_config.TextColumn('Emitente / Prestador'),
            'tomador':   st.column_config.TextColumn('Tomador / Destinatário'),
            'valor':     st.column_config.NumberColumn('Valor Total', format="R$ %.2f"),
        })

    st.divider()

    # ── Painel de duplicatas ──────────────────────────────────────────────
    acoes = {}
    if grupos_dup:
        st.markdown("""
        <div class="section-label"><span>04</span> Gestão de Duplicatas</div>
        <div class="section-title">Revise e decida o que fazer com cada grupo duplicado</div>
        """, unsafe_allow_html=True)
        acoes = render_painel_duplicatas(resultados, grupos_dup)

    # ── Download ──────────────────────────────────────────────────────────
    secao_dl = "05" if grupos_dup else "04"
    st.markdown("""
    <div class="section-label"><span>{s}</span> Download</div>
    <div class="section-title">Baixar Arquivos Renomeados</div>
    """.format(s=secao_dl), unsafe_allow_html=True)

    # Se há duplicatas sem decisão, avisa mas não bloqueia
    grupos_sem_decisao = [i for i in range(len(grupos_dup)) if i not in acoes]
    if grupos_sem_decisao:
        st.warning("⚠️ {} grupo(s) de duplicata ainda sem decisão. O download incluirá ambos os arquivos com sufixo _DUP1 / _DUP2 por padrão.".format(len(grupos_sem_decisao)))

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if qtd == 1 and not grupos_dup:
        r = resultados[0]
        nome_dl = r['novo_nome'] if r['novo_nome'] != '-' else r['original']
        st.download_button(label="📥 Baixar PDF Renomeado", data=r['_bytes'],
            file_name=nome_dl, mime="application/pdf", use_container_width=True)
    else:
        pares = construir_arquivos_download(resultados, grupos_dup, acoes)
        if pares:
            zip_bytes = montar_zip(pares)
            st.download_button(
                label="📥 Baixar ZIP com Notas Fiscais Renomeadas ({} arquivo(s))".format(len(pares)),
                data=zip_bytes,
                file_name="NOTAS_FISCAIS_RENOMEADAS_{}.zip".format(timestamp),
                mime="application/zip",
                use_container_width=True
            )
            st.success("✅ Pronto! Clique no botão acima para baixar os arquivos renomeados.")
        else:
            st.info("Nenhum arquivo selecionado para download após as decisões de duplicata.")

    # ── Expander de ajuda ─────────────────────────────────────────────────
    with st.expander("ℹ️ Formato, Exemplos e Critérios de Duplicata"):
        st.markdown("""
        ### Formato do Nome
        ```
        NF [NUMERO] - RAZAO SOCIAL DO EMITENTE.PDF
        ```

        ### Critérios de Duplicata Real
        O sistema considera dois arquivos como **possíveis duplicatas** quando todos os campos abaixo coincidem:

        | Campo | Descrição |
        |---|---|
        | **Número da NF** | Mesmo número extraído do PDF |
        | **Fornecedor** | Mesma razão social usada na nomenclatura |
        | **Data de Emissão** | Mesma data (normalizada para DD-MM-AAAA) |
        | **Valor Total** | Mesmo valor (comparado com 2 casas decimais) |

        > **Atenção:** Arquivos com número ou fornecedor não identificados **não** são marcados como duplicata.

        ### Opções ao detectar duplicata
        - **Ignorar duplicidade** — inclui os dois no ZIP com sufixo `_DUP1` / `_DUP2`
        - **Manter apenas 1º** — exclui o(s) demais do download
        - **Manter apenas 2º** — exclui o 1º do download
        - **Baixar os dois** — inclui ambos com sufixo para diferenciação

        ### Formatos suportados
        - NF-e DANFE (padrão SEFAZ) | NF-e **Omie** | NFS-e | CF-e/SAT
        """)

    st.divider()
    st.caption("Desenvolvido para Rezende Energia · Processamento de Notas Fiscais")


if __name__ == "__main__":
    main()
