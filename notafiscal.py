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
    --sucesso: #1DB954; --erro: #E53E3E; --radius: 12px;
    --shadow: 0 4px 24px rgba(247,147,30,0.10);
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
.metrics-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:1rem; margin:1.2rem 0; }
.metric-card { background:var(--branco); border:1px solid var(--cinza-borda); border-radius:var(--radius);
    padding:1.2rem 1.4rem; text-align:center; transition:all .25s; position:relative; overflow:hidden; }
.metric-card::before { content:''; position:absolute; top:0; left:0; right:0; height:3px;
    background:linear-gradient(90deg,var(--laranja),var(--laranja-vivo)); border-radius:var(--radius) var(--radius) 0 0; }
.metric-card:hover { box-shadow:var(--shadow); transform:translateY(-2px); }
.metric-value { font-size:1.8rem; font-weight:700; color:var(--texto); line-height:1;
    margin-bottom:.3rem; font-family:'JetBrains Mono',monospace; }
.metric-value.laranja { color:var(--laranja-vivo); }
.metric-value.sucesso { color:var(--sucesso); }
.metric-value.erro    { color:var(--erro); }
.metric-label { font-size:.78rem; font-weight:500; color:var(--texto-leve); letter-spacing:.5px; text-transform:uppercase; }
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
            'DANFE', 'ABAIXO', 'SERIE']
    cu = ascii_normalizar(c).upper()
    return not any(cu.startswith(l) for l in lixo)


def extrair_numero_nf(texto):
    """
    Extrai o numero da NF/NFS-e corretamente.
    Suporta: NF-e DANFE, Omie, DANFSe Santarem, NFSe prefeituras.
    """
    # 1. DANFSe Santarem: "NumerodaNFS-e Competencia DataHora\n30 09/02/2026..."
    m = re.search(r'NumerodaNFS-e\s+\S+\s+\S+\s*\n(\d+)\s+', texto, re.IGNORECASE)
    if m:
        try: return str(int(m.group(1)))
        except ValueError: pass

    # 2. NFSe Belem/prefeituras: "Numero / Serie\n09/12/2025 11:55:31 12/2025 9313 / E"
    m = re.search(r'Numero\s*/\s*Serie[^\n]*\n[^\n]*?(\d{2,})\s*/\s*[A-Z]', texto, re.IGNORECASE)
    if m:
        try: return str(int(m.group(1)))
        except ValueError: pass

    # 3. NFSe generica: "Numero da NFS-e\n30"
    m = re.search(r'Numero\s+da\s+NFS-?e\s*\n(\d+)', texto, re.IGNORECASE)
    if m:
        try: return str(int(m.group(1)))
        except ValueError: pass

    # 4. DANFE NF-e: linha "No 202.001" ou "No. 000.202.001" seguida de "Serie" ou "DATA"
    m = re.search(
        r'(?:^|\n)\s*N[o.]?\s*\.?\s*([\d. ]+)\s*\n\s*(?:Serie|SERIE|DATA|Folha|FOLHA)',
        texto, re.IGNORECASE | re.MULTILINE
    )
    if m:
        raw = m.group(1).replace('.', '').replace(' ', '')
        try: return str(int(raw))
        except ValueError: pass

    # 5. Fallback: No/N. numero curto (max 9 digitos, evita chave de acesso)
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

def extrair_info_nota_fiscal(pdf_bytes):
    """
    Extrai numero, emitente, tipo, data e valor de um PDF de NF.
    Retorna: (data, emitente, tipo_nf, valor, numero_nf)
    """
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            texto_raw = ""
            for page in pdf.pages[:2]:
                t = page.extract_text()
                if t:
                    texto_raw += t + "\n"

        if not texto_raw.strip():
            return None, None, None, None, None

        # Trabalhar com texto ASCII (sem acentos) para facilitar regex
        texto = ascii_normalizar(texto_raw)
        texto_up = texto.upper()

        data = None
        emitente = None
        tipo_nf = None
        valor = None
        numero_nf = None

        # ── Tipo ──────────────────────────────────────────────────────────
        if "DANFE" in texto_up or "DOCUMENTO AUXILIAR DA NOTA FISCAL" in texto_up:
            tipo_nf = "NF-e"
        elif "NOTA FISCAL DE SERVI" in texto_up or "NFS-E" in texto_up:
            tipo_nf = "NFS-e"
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

        # ── Emitente (fornecedor que emitiu a NF) ─────────────────────────
        # 1. Omie: "RECEBEMOS DE [EMPRESA] OS PRODUTOS"
        if not emitente:
            m = re.search(r'RECEBEMOS\s+DE\s+(.+?)\s+OS\s+PRODUTOS', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        # 2. Padrao SEFAZ: "RECEBI(EMOS) DE [EMPRESA], OS PRODUTOS"
        if not emitente:
            m = re.search(r'RECEBI\(EMOS\)\s+DE\s+(.+?),\s+OS\s+PRODUTOS', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        # 3. Omie: linha apos "IDENTIFICACAO DO EMITENTE"
        if not emitente:
            m = re.search(
                r'IDENTIFICACAO\s+DO\s+EMITENTE[^\n]*\n([A-Z][A-Z0-9 &.,/\-]{2,70}?)\s+(?:Eletro|DANFE|0\s+-|1\s+-)',
                texto, re.IGNORECASE
            )
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        # 4. DANFSe Santarem: "EMITENTEDANFS-e ... Nome/NomeEmpresarial\nNOME email"
        if not emitente:
            m = re.search(
                r'EMITENTEDANFS-e.*?Nome/NomeEmpresarial[^\n]*\n([^\n]{3,80})',
                texto, re.IGNORECASE | re.DOTALL
            )
            if m:
                c = re.split(r'\s+\S+@\S+', m.group(1))[0].strip()
                if candidato_valido(c):
                    emitente = c

        # 5. NFSe Belem/outras prefeituras: "EMITENTE PRESTADOR DO SERVICO ... Nome / Nome Empresarial\nNOME email"
        if not emitente:
            m = re.search(
                r'EMITENTE\s+PRESTADOR\s+DO\s+SERVICO.*?Nome\s*/\s*Nome\s+Empresarial[^\n]*\n([^\n]{3,80})',
                texto, re.IGNORECASE | re.DOTALL
            )
            if m:
                c = re.split(r'\s+\S+@\S+', m.group(1))[0].strip()
                if candidato_valido(c):
                    emitente = c

        # 6. NFS-e generica: Prestador de Servicos
        if not emitente:
            m = re.search(r'Prestador\s+de\s+Servicos\s*\n\s*([A-Z][^\n\r]{2,79})', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        # 7. Label "Emitente:"
        if not emitente:
            m = re.search(r'Emitente\s*[:\s]+([^\n\r]{3,80})', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        # 8. Razao Social generica (primeira ocorrencia)
        if not emitente:
            m = re.search(r'Razao\s+Social\s*[:\s]+([^\n\r]{3,80})', texto, re.IGNORECASE)
            if m and candidato_valido(m.group(1)):
                emitente = m.group(1).strip()

        # 9. Fallback: linha com sufixo empresarial
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

        return data, emitente, tipo_nf, valor, numero_nf

    except Exception as e:
        st.error("Erro ao processar PDF: {}".format(str(e)))
        return None, None, None, None, None


# =============================================================================
#  PROCESSAMENTO
# =============================================================================

def processar_pdfs(uploaded_files):
    resultados = []
    arquivos_zip = []
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, uf in enumerate(uploaded_files):
        status_text.text("Processando {}...".format(uf.name))
        pdf_bytes = uf.read()

        data, emitente, tipo_nf, valor, numero_nf = extrair_info_nota_fiscal(pdf_bytes)

        if numero_nf and emitente:
            emitente_limpo = limpar_nome(emitente)
            novo_nome = montar_nome(numero_nf, emitente_limpo)
            status = '✅ Sucesso'
        else:
            emitente_limpo = limpar_nome(emitente) if emitente else 'N/A'
            novo_nome = '-'
            status = '⚠️ Informações não encontradas'

        resultados.append({
            'original':  uf.name,
            'novo_nome': novo_nome,
            'status':    status,
            'tipo':      tipo_nf or 'Desconhecido',
            'numero_nf': numero_nf or 'N/A',
            'data':      data or 'N/A',
            'emitente':  emitente_limpo,
            'valor':     valor,
            '_bytes':    pdf_bytes,
        })
        arquivos_zip.append((novo_nome if novo_nome != '-' else uf.name, pdf_bytes))
        progress_bar.progress((idx + 1) / len(uploaded_files))

    status_text.text("✅ Processamento concluído!")

    if len(uploaded_files) == 1:
        return None, resultados

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for nome, conteudo in arquivos_zip:
            zf.writestr(nome, conteudo)
    return buf.getvalue(), resultados


# =============================================================================
#  METRICAS
# =============================================================================

def render_metricas(resultados):
    total = len(resultados)
    sucesso = sum(1 for r in resultados if '✅' in r['status'])
    erros = total - sucesso
    valor_total = sum(r.get('valor') or 0 for r in resultados)
    valor_fmt = "R$ {:,.2f}".format(valor_total).replace(',', '_').replace('.', ',').replace('_', '.')
    st.markdown("""
    <div class="metrics-grid">
        <div class="metric-card"><div class="metric-value laranja">{total}</div><div class="metric-label">Total de Arquivos</div></div>
        <div class="metric-card"><div class="metric-value sucesso">{sucesso}</div><div class="metric-label">Processados com Sucesso</div></div>
        <div class="metric-card"><div class="metric-value erro">{erros}</div><div class="metric-label">Erros / Avisos</div></div>
        <div class="metric-card"><div class="metric-value laranja" style="font-size:1.3rem">{valor}</div><div class="metric-label">Valor Total</div></div>
    </div>
    """.format(total=total, sucesso=sucesso, erros=erros, valor=valor_fmt), unsafe_allow_html=True)


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
            Processa notas fiscais em PDF e as renomeia automaticamente no padrao<br>
            <strong style="color:#FFD39A;">NF [Numero] &mdash; Razao Social do Emitente (Fornecedor)</strong>
            &mdash; sempre em <strong style="color:#F7931E;">MAIUSCULAS</strong>.
        </div>
        <div class="hero-badges">
            <span class="badge">📄 NF-e DANFE</span>
            <span class="badge">⚡ OMIE</span>
            <span class="badge">🛠️ NFS-e</span>
            <span class="badge">🧾 NF SIMPLES</span>
            <span class="badge">🏷️ CF-e / SAT</span>
            <span class="badge">🔠 SEMPRE MAIUSCULAS</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="section-label"><span>01</span> Upload</div>
    <div class="section-title">Envie os PDFs das notas fiscais (pode selecionar varios de uma vez)</div>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Arraste ou clique para selecionar os PDFs",
        type=['pdf'],
        accept_multiple_files=True,
        help="Selecione um ou mais PDFs de notas fiscais"
    )

    if uploaded_files:
        qtd = len(uploaded_files)
        st.success("✅ {} arquivo(s) carregado(s)".format(qtd))

        if st.button("🚀 Processar Notas Fiscais", type="primary", use_container_width=True):
            with st.spinner("Processando notas fiscais..."):
                zip_output, resultados = processar_pdfs(uploaded_files)

            if resultados:
                st.divider()
                st.markdown("""
                <div class="section-label"><span>02</span> Resultados</div>
                <div class="section-title">Resumo do Processamento</div>
                """, unsafe_allow_html=True)

                render_metricas(resultados)

                df_display = [{k: v for k, v in r.items() if not k.startswith('_')} for r in resultados]
                st.dataframe(df_display, use_container_width=True, hide_index=True,
                    column_config={
                        'original':  st.column_config.TextColumn('Nome Original'),
                        'novo_nome': st.column_config.TextColumn('Novo Nome'),
                        'status':    st.column_config.TextColumn('Status'),
                        'tipo':      st.column_config.TextColumn('Tipo'),
                        'numero_nf': st.column_config.TextColumn('Numero NF'),
                        'data':      st.column_config.TextColumn('Data de Emissao'),
                        'emitente':  st.column_config.TextColumn('Emitente'),
                        'valor':     st.column_config.NumberColumn('Valor Total', format="R$ %.2f"),
                    })

                st.divider()
                st.markdown("""
                <div class="section-label"><span>03</span> Download</div>
                <div class="section-title">Baixar Arquivos Renomeados</div>
                """, unsafe_allow_html=True)

                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

                if qtd == 1:
                    r = resultados[0]
                    nome_dl = r['novo_nome'] if r['novo_nome'] != '-' else r['original']
                    st.download_button(label="📥 Baixar PDF Renomeado", data=r['_bytes'],
                        file_name=nome_dl, mime="application/pdf", use_container_width=True)
                else:
                    st.download_button(label="📥 Baixar ZIP com Notas Fiscais Renomeadas",
                        data=zip_output,
                        file_name="NOTAS_FISCAIS_RENOMEADAS_{}.zip".format(timestamp),
                        mime="application/zip", use_container_width=True)

                st.success("✅ Pronto! Clique no botao acima para baixar os arquivos renomeados.")

    with st.expander("ℹ️ Formato, Exemplos e Requisitos"):
        st.markdown("""
        ### Formato do Nome
        ```
        NF [NUMERO] - RAZAO SOCIAL DO EMITENTE.PDF
        ```
        ### Exemplos
        | Antes | Depois |
        |---|---|
        | `NF_202001_-_PMZ.pdf` | `NF 202001 - PMZ DISTRIBUIDORA SA.PDF` |
        | `152603...193.pdf` | `NF 6 - BERIT PECAS AUTOMOTIVAS LTDA ME.PDF` |

        ### Como usar
        1. Selecione um ou mais PDFs
        2. Clique em **Processar Notas Fiscais**
        3. PDF unico: baixa o PDF ja renomeado | Varios: baixa ZIP

        ### Formatos suportados
        - NF-e DANFE (padrao SEFAZ) | NF-e **Omie** | NFS-e | CF-e/SAT
        """)

    st.divider()
    st.caption("Desenvolvido para Rezende Energia · Processamento de Notas Fiscais")


if __name__ == "__main__":
    main()
