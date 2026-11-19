import streamlit as st
import pandas as pd
import re, io, os, base64, warnings, openpyxl, pdfplumber
from datetime import datetime, timedelta
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from fpdf import FPDF

warnings.filterwarnings("ignore")
st.set_page_config(page_title="Escalas ORV CP - by Luís Maia", page_icon="🚄", layout="wide")

FERIADOS = {
    "2026-01-01","2026-04-03","2026-04-05","2026-04-25","2026-05-01","2026-06-04","2026-06-10","2026-07-04","2026-08-15","2026-10-05","2026-11-01","2026-12-01","2026-12-08","2026-12-25",
    "2027-01-01","2027-03-26","2027-03-28","2027-04-25","2027-05-01","2027-05-27","2027-06-10","2027-07-04","2027-08-15","2027-10-05","2027-11-01","2027-12-01","2027-12-08","2027-12-25",
    "2028-01-01","2028-04-14","2028-04-16","2028-04-25","2028-05-01","2028-06-10","2028-06-15","2028-07-04","2028-08-15","2028-10-05","2028-11-01","2028-12-01","2028-12-08","2028-12-25",
    "2029-01-01","2029-03-30","2029-04-01","2029-04-25","2029-05-01","2029-05-31","2029-06-10","2029-07-04","2029-08-15","2029-10-05","2029-11-01","2029-12-01","2029-12-08","2029-12-25"
}

def hora_para_min(h_str):
    m = re.match(r'^(\d{1,2}):(\d{2})', str(h_str).strip())
    return int(m.group(1)) * 60 + int(m.group(2)) if m else None

def min_para_hora(mins):
    mins = int(round(mins)) % (24 * 60)
    return f"{mins // 60:02d}:{mins % 60:02d}"

def fmt_hm(mins, com_sinal=False):
    s = "-" if mins < 0 else ("+" if com_sinal and mins > 0 else "")
    m = abs(int(round(mins)))
    return f"{s}{m//60:02d}h{m%60:02d}m"

def extrair_marchas_livro_ip_cp(pdf_obj):
    """Lê tabelas verticais dos Livros Horários Técnicos IP/CP (Estações à esquerda, Comboios nas colunas)."""
    encontrados = {}
    for pag in pdf_obj.pages:
        tabelas = pag.extract_tables() or []
        for tab in tabelas:
            if not tab or len(tab) < 4:
                continue
            # Procura a linha de cabeçalho que contém os números dos comboios
            idx_cab = -1
            cols_cb = {}
            for r_i, row in enumerate(tab[:6]):
                for c_i, cell in enumerate(row):
                    txt = str(cell or "").strip()
                    m = re.match(r'^(\d{3,5})(?:/\d)?$', txt)
                    if m:
                        cols_cb[c_i] = m.group(1)
                if len(cols_cb) >= 2:
                    idx_cab = r_i
                    break
            if idx_cab == -1:
                continue

            # Varre cada coluna de comboio para apanhar Origem/Partida e Destino/Chegada
            for c_i, num_cb in cols_cb.items():
                pontos = []
                for row in tab[idx_cab + 1:]:
                    if len(row) <= c_i:
                        continue
                    est = re.sub(r'\s+', ' ', str(row or "")).strip()
                    val = str(row[c_i] or "").strip()
                    horas = re.findall(r'\b([0-2]?\d:[0-5]\d)\b', val)
                    if horas and est and not est.isdigit():
                        pontos.append((est, horas, horas[-1]))
                if len(pontos) >= 2 and num_cb not in encontrados:
                    encontrados[num_cb] = {
                        "Origem": pontos[:12],
                        "Partida": pontos,
                        "Destino": pontos[-1][:12],
                        "Chegada": pontos[-1]
                    }
    return encontrados

@st.cache_data(show_spinner=False)
def carregar_base_comboios():
    base = {}
    if os.path.exists("comboios.csv"):
        try:
            df = pd.read_csv("comboios.csv", sep=";", dtype=str).fillna("")
            for _, row in df.iterrows():
                cb = str(row["Comboio"]).strip()
                if cb:
                    base[cb] = {
                        "Origem": row["Origem"].strip(),
                        "Partida": row["Partida"].strip(),
                        "Destino": row["Destino"].strip(),
                        "Chegada": row["Chegada"].strip()
                    }
        except Exception:
            pass

    if os.path.isdir("horarios"):
        for fich in os.listdir("horarios"):
            if fich.lower().endswith(".pdf"):
                try:
                    with pdfplumber.open(os.path.join("horarios", fich)) as pdf:
                        base.update(extrair_marchas_livro_ip_cp(pdf))
                except Exception:
                    pass
    return base

def extrair_horarios_de_pdf_cp(uploaded_pdf):
    with pdfplumber.open(io.BytesIO(uploaded_pdf.read())) as pdf:
        return extrair_marchas_livro_ip_cp(pdf)


def processar_linhas_brutas(linhas):
    dados, os_atual, dias_atual = [], None, ""
    col_os, col_dias, col_lapr, col_hapr, col_serv, col_lret, col_hret, col_pnt, col_reps, col_repf, col_obs = 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
    lixo = ("EM VIGOR", "ALTERAÇÃO", "CATEGORIA", "DEPÓSITO", "Planeamento", "Versão electrónica", "F.500.02", "Revisão e Venda", "Local", "Hora")
    for r in linhas:
        r = [re.sub(r'\s+', ' ', str(c)).strip() if c is not None else "" for c in r]
        if not any(r) or any(p in " ".join(r) for p in lixo): continue
        if any("Dias da Semana" in c for c in r):
            for idx, txt in enumerate(r):
                if txt == "OS": col_os = idx
                elif "Dias da Semana" in txt: col_dias = idx
                elif "Designação" in txt: col_serv = idx
                elif "PNT" in txt: col_pnt = idx
            col_lapr, col_hapr = col_dias + 1, col_dias + 2
            col_lret, col_hret = max(0, col_pnt - 2), max(0, col_pnt - 1)
            col_reps, col_repf, col_obs = col_pnt + 1, col_pnt + 2, col_pnt + 3
            continue
        if len(r) <= col_pnt: continue
        c_os = r[col_os]
        if c_os.replace(".0", "").isdigit() and 1 <= int(float(c_os)) <= 365:
            os_atual = int(float(c_os))
        c_dias = r[col_dias]
        if c_dias and "Dias" not in c_dias: dias_atual = c_dias
        if os_atual:
            serv = r[col_serv] if len(r) > col_serv else ""
            lapr = r[col_lapr] if len(r) > col_lapr else ""
            hapr = r[col_hapr] if len(r) > col_hapr else ""
            lret = r[col_lret] if len(r) > col_lret else ""
            hret = r[col_hret] if len(r) > col_hret else ""
            pnt = r[col_pnt] if len(r) > col_pnt else ""
            reps = r[col_reps] if len(r) > col_reps else ""
            repf = r[col_repf] if len(r) > col_repf else ""
            obs = r[col_obs] if len(r) > col_obs else ""
            if c_dias or serv or pnt:
                dados.append({
                    "OS": os_atual, "Dias": c_dias, "Dias_Efet": dias_atual,
                    "LocalApr": lapr, "HoraApr": hapr, "Serv": serv,
                    "LocalRet": lret, "HoraRet": hret, "PNT": pnt,
                    "RepSede": reps.lower() == "x", "RepFora": repf.lower() == "x", "Obs": obs
                })
    return dados

DEPOSITOS_CP = {
    "COB": "COIMBRA B",
    "PCA": "PORTO CAMPANHÃ",
    "ENT": "ENTRONCAMENTO",
    "LSA": "LISBOA SANTA APOLÓNIA",
    "LRO": "LISBOA ROSSIO",
    "FAR": "FARO",
    "REG": "RÉGUA",
    
}

def identificar_deposito(texto, sede_sigla):
    t_up = texto.upper()
    for sigla, nome_cid in DEPOSITOS_CP.items():
        if nome_cid in t_up:
            return nome_cid
    return DEPOSITOS_CP.get(sede_sigla, sede_sigla)

def ler_ficheiro_carregado(uploaded_file):
    nome, bytes_data = uploaded_file.name.lower(), uploaded_file.read()
    meta = {"Escala": "ORV CP", "Deposito": "NACIONAL", "Sede": "COB"}
    if nome.endswith(".pdf"):
        todas_linhas = []
        with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
            txt1 = pdf.pages[0].extract_text() or ""
            m_esc = re.search(r'\b([A-Z]{3}R\d{2})\b', txt1)
            if m_esc:
                meta["Escala"] = m_esc.group(1)
                meta["Sede"] = m_esc.group(1)[:3]
            meta["Deposito"] = identificar_deposito(txt1, meta["Sede"])
            for pagina in pdf.pages:
                for tabela in pagina.extract_tables(): todas_linhas.extend(tabela)
        return processar_linhas_brutas(todas_linhas), meta
    wb = openpyxl.load_workbook(io.BytesIO(bytes_data), data_only=True)
    for sheet in wb.sheetnames:
        linhas = [["" if c is None else str(c).strip() for c in r] for r in wb[sheet].iter_rows(values_only=True)]
        txt_topo = " ".join(" ".join(l) for l in linhas[:4])
        m_esc = re.search(r'\b([A-Z]{3}R\d{2})\b', txt_topo)
        if m_esc:
            meta["Escala"] = m_esc.group(1)
            meta["Sede"] = m_esc.group(1)[:3]
        meta["Deposito"] = identificar_deposito(txt_topo, meta["Sede"])
        if any(any("Dias da Semana" in c for c in l) for l in linhas):
            return processar_linhas_brutas(linhas), meta
    return [], meta

def obter_total_os(dados):
    return max([x["OS"] for x in dados], default=120)

def contar_repousos_escala(dados, max_os):
    return sum(len(set(x["Dias_Efet"] for x in dados if x["OS"] == os_n and x["RepFora"])) for os_n in range(1, max_os + 1))

def expandir(regra):
    pares = set()
    for b in str(regra).split():
        m = re.match(r'^([0-9\.]+)([NVFS]*)$', b)
        if not m: continue
        dt, tt = m.groups()
        tipos = list(tt) if tt else ['N','V','F','S']
        if '..' in dt:
            partes = dt.split('..')
            dias = range(int(partes[0]), int(partes[1]) + 1)
        else:
            dias = [int(dt)]
        for d in dias:
            for t in tipos: pares.add((d, t))
    return pares

def calcular_semanas_detalhado(dados, data_ini_str, max_os):
    regras_os = {os_n: [] for os_n in range(1, max_os + 1)}
    for os_n in range(1, max_os + 1):
        por_regra = {}
        for x in [k for k in dados if k["OS"] == os_n]:
            reg = x["Dias_Efet"]
            if not reg: continue
            por_regra.setdefault(reg, 0)
            m_pnt = re.match(r'^(\d{1,2}):(\d{2})', str(x["PNT"]).strip())
            if m_pnt: por_regra[reg] += int(m_pnt.group(1))*60 + int(m_pnt.group(2))
        for reg, mins in por_regra.items(): regras_os[os_n].append((expandir(reg), mins))
    base = datetime.strptime(data_ini_str, "%Y-%m-%d").date()
    lista_semanas = []
    for sem in range(max_os):
        mins_sem = 0
        for d in range(7):
            dia_abs = sem * 7 + d
            dt = base + timedelta(days=dia_abs)
            de = 1 if dt.weekday() == 6 else dt.weekday() + 2
            tp = 'F' if dt.strftime("%Y-%m-%d") in FERIADOS else ('V' if (dt+timedelta(days=1)).strftime("%Y-%m-%d") in FERIADOS else ('S' if (dt-timedelta(days=1)).strftime("%Y-%m-%d") in FERIADOS else 'N'))
            for pv, mins in regras_os[(dia_abs % max_os) + 1]:
                if (de, tp) in pv:
                    mins_sem += mins
                    break
        lista_semanas.append(mins_sem)
    return lista_semanas

def calcular_indicadores_dia(dados, max_os):
    regras_os = {os_n: [] for os_n in range(1, max_os + 1)}
    for os_n in range(1, max_os + 1):
        for x in [k for k in dados if k["OS"] == os_n]:
            reg = x["Dias_Efet"]
            if not reg: continue
            serv_txt = x["Serv"].upper()
            estado = "DESCANSO" if "DESCANSO" in serv_txt else ("SUPRA" if "SUPRA" in serv_txt else "SERVICOS")
            regras_os[os_n].append((expandir(reg), estado))
    contagem = {d: {"SERVICOS": 0, "SUPRA": 0, "DESCANSO": 0} for d in range(1, 8)}
    for d in range(1, 8):
        for os_n in range(1, max_os + 1):
            estado_encontrado = "DESCANSO"
            for pv, est in regras_os[os_n]:
                if (d, 'N') in pv:
                    estado_encontrado = est
                    break
            contagem[d][estado_encontrado] += 1
    return contagem

def comboios(txt):
    limpo = re.sub(r'\(.*?\)', '', str(txt))
    limpo = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', '', limpo)
    return set(re.findall(r'\b\d{3,5}\b', limpo))

def limpar_latin1(txt):
    return str(txt).replace("•", "-").replace("→", "->").encode("latin-1", "replace").decode("latin-1")

def auditar_linha_editada(serv_txt, base_cb, sede="COB"):
    cbs = re.findall(r'\b\d{3,5}\b', re.sub(r'\(.*?\)', '', serv_txt))
    cbs_conhecidos = [c for c in cbs if c in base_cb]
    cbs_desconhecidos = [c for c in cbs if c not in base_cb]
    alertas = []
    if cbs_desconhecidos:
        alertas.append(f"⚠️ Sem horário no comboios.csv: {', '.join(cbs_desconhecidos)}")

    loc_apr, hora_apr, loc_ret, hora_ret, pnt_str = sede, "", sede, "", ""
    if cbs_conhecidos:
        primeiro, ultimo = base_cb[cbs_conhecidos[0]], base_cb[cbs_conhecidos[-1]]
        loc_apr, loc_ret = primeiro["Origem"], ultimo["Destino"]
        m_p, m_c = hora_para_min(primeiro["Partida"]), hora_para_min(ultimo["Chegada"])
        if m_p is not None and m_c is not None:
            ini_j = (m_p - 10) % (24 * 60)
            fim_j = (m_c + 10) % (24 * 60)
            hora_apr, hora_ret = min_para_hora(ini_j), min_para_hora(fim_j)
            dur = (fim_j - ini_j) if fim_j >= ini_j else ((24 * 60 - ini_j) + fim_j)
            pnt_str = f"{dur // 60:02d}:{dur % 60:02d}"
            if dur < 6 * 60:
                alertas.append(f"🔴 PNT abaixo do mínimo (06h00): {fmt_hm(dur)}")
            elif dur > 9 * 60:
                alertas.append(f"🔴 PNT acima do máximo (09h00): {fmt_hm(dur)}")

    if "(TR)" in serv_txt.upper():
        partes = re.split(r'\(TR\)', serv_txt, flags=re.IGNORECASE)
        if len(partes) >= 2:
            c_antes = re.findall(r'\b\d{3,5}\b', partes[0])
            c_depois = re.findall(r'\b\d{3,5}\b', partes[1])
            if c_antes and c_depois and c_antes[-1] in base_cb and c_depois[0] in base_cb:
                h_cheg_antes = hora_para_min(base_cb[c_antes[-1]]["Chegada"])
                h_part_depois = hora_para_min(base_cb[c_depois[0]]["Partida"])
                if h_cheg_antes is not None and h_part_depois is not None:
                    tr_mins = (h_part_depois - h_cheg_antes) if h_part_depois >= h_cheg_antes else ((24 * 60 - h_cheg_antes) + h_part_depois)
                    if tr_mins < 45:
                        alertas.append(f"🔴 Falta de (TR): {tr_mins}m (Mínimo obrigatório: 45m)")
                    else:
                        alertas.append(f"🟢 (TR) Válida: {tr_mins}m (>= 45m)")
    return loc_apr, hora_apr, loc_ret, hora_ret, pnt_str, alertas

class PDFComAssinatura(FPDF):
    def __init__(self, titulo_escala="COBR20", *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.titulo_escala = titulo_escala

    def header(self):
        self.set_font("Helvetica", "B", 38)
        self.set_text_color(242, 244, 247)
        with self.rotation(30, x=148, y=105):
            self.text(85, 115, limpar_latin1("by Luís Maia"))
        self.set_text_color(20, 20, 20)

    def footer(self):
        self.set_y(-11)
        self.set_font("Helvetica", "I", 8.5)
        self.set_text_color(110, 110, 110)
        self.cell(0, 6, limpar_latin1(f"Escala {self.titulo_escala}  |  by Luís Maia  |  Página {self.page_no()}/{{nb}}"), align="C")

def gerar_pdf_nativo(lista_os_alt, lista_cb_mov, lista_rep_alt, balanco_rows, dist_rows, sem_rows, titulo_esc="COBR20", so_os=False):
    pdf = PDFComAssinatura(titulo_escala=titulo_esc, orientation="L", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()

    def titulo_sec(txt):
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(31, 78, 120)
        pdf.cell(0, 8, limpar_latin1(txt), new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(20, 20, 20)

    def cabecalho_tabela(cols, largs):
        pdf.set_fill_color(31, 78, 120)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 8.5)
        for c, w in zip(cols, largs):
            pdf.cell(w, 7, limpar_latin1(c), border=1, align="C", fill=True)
        pdf.ln()
        pdf.set_text_color(20, 20, 20)

    titulo_sec(f"Resumo de OS Alteradas - Escala {titulo_esc}" if so_os else f"Comparação Completa - Escala {titulo_esc}")
    pdf.ln(1)
    titulo_sec("1. Alterações por Ordem de Serviço (OS)")
    l_os = (20, 46, 105, 105)
    cabecalho_tabela(["OS", "Tipo de Alteração", "Escala Anterior", "Escala Nova (Negrito)"], l_os)

    for os_str, tp, ta, tn, cor_rgb in lista_os_alt:
        linhas_max = max(ta.count("\n") + 1, tn.count("\n") + 1, 1)
        alt_linha = max(7, linhas_max * 4.5)
        if pdf.get_y() + alt_linha > 192:
            pdf.add_page()
            cabecalho_tabela(["OS", "Tipo de Alteração", "Escala Anterior", "Escala Nova (Negrito)"], l_os)
        y_ini = pdf.get_y()
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_os[0], alt_linha, limpar_latin1(os_str), border=1, align="C", fill=True)
        pdf.cell(l_os[1], alt_linha, limpar_latin1(tp), border=1, align="C", fill=True)
        x_ta = pdf.get_x()
        pdf.set_font("Helvetica", "", 7.5)
        pdf.multi_cell(l_os[2], alt_linha / linhas_max, limpar_latin1(ta), border=1, fill=True)
        pdf.set_xy(x_ta + l_os[2], y_ini)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.multi_cell(l_os[3], alt_linha / linhas_max, limpar_latin1(tn), border=1, fill=True)
        pdf.set_y(y_ini + alt_linha)

    if so_os:
        return bytes(pdf.output())

    pdf.add_page()
    titulo_sec("2. Mobilidade de Comboios entre OS")
    l_cb = (24, 46, 50, 50, 53, 53)
    cabecalho_tabela(["Comboio", "Movimento", "Saiu das OS", "Entrou nas OS", "OS Antes", "OS Agora"], l_cb)
    for cb, mv, saiu, entrou, oa, on, cor_rgb in lista_cb_mov:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8)
        for idx_c, val in enumerate((cb, mv, saiu, entrou, oa, on)):
            pdf.cell(l_cb[idx_c], 6.5, limpar_latin1(val), border=1, align=("C" if idx_c < 2 else "L"), fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("3. Alterações nos Repousos Fora da Sede (REP F)")
    l_rep = (22, 54, 100, 100)
    cabecalho_tabela(["OS", "Situação Repouso", "Repousos Antes", "Repousos Agora"], l_rep)
    for os_s, st_r, ra, rn, cor_rgb in lista_rep_alt:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8)
        for idx_c, val in enumerate((os_s, st_r, ra, rn)):
            pdf.cell(l_rep[idx_c], 6.5, limpar_latin1(val), border=1, align=("C" if idx_c < 2 else "L"), fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("4. Balanço Final: Ganhos vs Perdas")
    l_bal = (115, 45, 116)
    cabecalho_tabela(["Indicador", "Antes -> Agora", "Impacto / Saldo"], l_bal)
    for ind, qtd, obs, cor_rgb in balanco_rows:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(l_bal[0], 7, limpar_latin1(ind), border=1, fill=True)
        pdf.cell(l_bal[1], 7, limpar_latin1(qtd), border=1, align="C", fill=True)
        pdf.cell(l_bal[2], 7, limpar_latin1(obs), border=1, fill=True)
        pdf.ln()

    pdf.add_page()
    titulo_sec("5. Distribuição Diária da Escala (Serviços, SUPRA e Descanso)")
    l_dist = (48, 52, 36, 52, 44, 44)
    cabecalho_tabela(["Dia da Semana", "Serviços (Ant -> Nov)", "Dif.", "SUPRA (Ant -> Nov)", "Descanso", "Total OS"], l_dist)
    for d_row in dist_rows:
        pdf.set_fill_color(*d_row[6])
        pdf.set_font("Helvetica", "B", 8.5)
        for idx_c in range(6):
            pdf.cell(l_dist[idx_c], 7, limpar_latin1(d_row[idx_c]), border=1, align=("L" if idx_c == 0 else "C"), fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("6. Cômputo a cada 12 Semanas")
    l_sem = (45, 55, 58, 58, 60)
    cabecalho_tabela(["Bloco", "Semanas", "Horas Anterior", "Horas Nova", "Diferença (12 Sem.)"], l_sem)
    for s_row in sem_rows:
        pdf.set_fill_color(255, 235, 156)
        pdf.set_font("Helvetica", "B", 8.5)
        for idx_c in range(5):
            pdf.cell(l_sem[idx_c], 7, limpar_latin1(s_row[idx_c]), border=1, align="C", fill=True)
        pdf.ln()

    return bytes(pdf.output())

def gerar_pdf_impacto_simulacao(linhas_comp, resumo_infracoes, titulo_esc="COBR20"):
    pdf = PDFComAssinatura(titulo_escala=titulo_esc, orientation="L", unit="mm", format="A4")
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(31, 78, 120)
    pdf.cell(0, 8, limpar_latin1(f"Relatório de Impacto de Simulação - Escala {titulo_esc}"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 7, limpar_latin1("1. Auditoria de Regras (Repouso Sede 14h | Fora 9h | TR 45m | PNT 6h-9h)"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(20, 20, 20)
    for inf in resumo_infracoes:
        pdf.cell(0, 6, limpar_latin1(f"- {inf}"), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(31, 78, 120)
    pdf.cell(0, 7, limpar_latin1("2. Comparação das OS Editadas (Original vs. Simulação)"), new_x="LMARGIN", new_y="NEXT")
    l_sim = (20, 128, 128)
    pdf.set_fill_color(31, 78, 120)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8.5)
    for c, w in zip(["OS", "Original", "Editada (Simulação)"], l_sim):
        pdf.cell(w, 7, limpar_latin1(c), border=1, align="C", fill=True)
    pdf.ln()
    pdf.set_text_color(20, 20, 20)
    for os_s, orig_t, edit_t in linhas_comp:
        pdf.set_fill_color(198, 239, 206)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_sim[0], 10, limpar_latin1(os_s), border=1, align="C", fill=True)
        pdf.set_font("Helvetica", "", 7.5)
        pdf.cell(l_sim[1], 10, limpar_latin1(orig_t[:85]), border=1, fill=True)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.cell(l_sim[2], 10, limpar_latin1(edit_t[:85]), border=1, fill=True)
        pdf.ln()
    return bytes(pdf.output())

def converter_pdf_para_excel_validado(uploaded_pdf):
    bytes_data = uploaded_pdf.read()
    todas_linhas = []
    data_vigor, num_alt, cod_esc = "22/11/2026", "ALTERAÇÃO", "ORV_CP"
    with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
        txt_p1 = pdf.pages[0].extract_text() or ""
        m_dt = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', txt_p1)
        if m_dt: data_vigor = m_dt.group(1)
        m_al = re.search(r'(\d+ª\s+ALTERAÇÃO)', txt_p1)
        if m_al: num_alt = m_al.group(1)
        m_es = re.search(r'\b([A-Z]{3}R\d{2})\b', txt_p1)
        if m_es: cod_esc = m_es.group(1)
        for pagina in pdf.pages:
            for tabela in pagina.extract_tables(): todas_linhas.extend(tabela)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = cod_esc
    verde_cp = PatternFill("solid", fgColor="1B4F3E")
    fonte_branca = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fonte_bold = Font(name="Calibri", size=9, bold=True, color="000000")
    fonte_norm = Font(name="Calibri", size=9, bold=False, color="000000")
    borda_fina = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    esq = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.append(["EM VIGOR DESDE", "CATEGORIA", "", f"{num_alt} EM VIGOR A PARTIR DE:", "", "ESCALA", "DEPÓSITO DE REVISÃO", "", "", "", ""])
    ws.append(["31/05/2020", "Operador de Revisão e Venda", "", data_vigor, "", cod_esc, "NACIONAL CP", "", "", "", ""])
    ws.append([""] * 11)
    ws.append(["OS", "Dias da Semana", "Apresentação", "", "Designação dos Serviços", "Ret. Prevista", "", "PNT", "REP", "", "Obs"])
    ws.append(["", "", "Local", "Hora", "", "Local", "Hora", "", "S", "F", ""])

    for r_idx in (1, 4, 5):
        for c_idx in range(1, 12):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.fill, cell.font, cell.alignment, cell.border = verde_cp, fonte_branca, centro, borda_fina
    for c_idx in range(1, 12):
        cell = ws.cell(row=2, column=c_idx)
        cell.font, cell.alignment, cell.border = fonte_bold, centro, borda_fina

    for rng in ("C1:E1", "C2:E2", "G1:K1", "G2:K2", "A4:A5", "B4:B5", "C4:D4", "E4:E5", "F4:G4", "H4:H5", "I4:J4", "K4:K5"):
        ws.merge_cells(rng)

    lixo = ("EM VIGOR", "ALTERAÇÃO", "CATEGORIA", "DEPÓSITO", "Planeamento", "Versão electrónica", "F.500.02", "Revisão e Venda", "Local", "Hora", "Dias da Semana")
    os_encontradas, os_atual, inicio_os_row = set(), None, None

    for r in todas_linhas:
        r = [re.sub(r'\s+', ' ', str(c)).strip() if c is not None else "" for c in r]
        if not any(r) or any(p in " ".join(r) for p in lixo) or len(r) < 10: continue
        linha_11 = (r + [""] * 11)[:11]
        c_os = linha_11[0]
        if c_os.isdigit() and 1 <= int(c_os) <= 365:
            if os_atual and inicio_os_row and ws.max_row > inicio_os_row:
                ws.merge_cells(start_row=inicio_os_row, start_column=1, end_row=ws.max_row, end_column=1)
            os_atual = int(c_os)
def aplicar_visual_cp(img_path="alfa-pendular-wext3_856.jpg"):
    bg_css = ""
    try:
        with open(img_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        bg_css = f"background-image: linear-gradient(120deg, rgba(0, 55, 30, 0.80), rgba(15, 23, 42, 0.70)), url('data:image/jpeg;base64,{b64}');"
    except Exception:
        bg_css = "background: linear-gradient(120deg, #006837, #0F172A);"

    st.markdown(f"""
    <style>
        .hero-cp {{
            {bg_css}
            background-size: cover;
            background-position: center 65%;
            padding: 28px 24px;
            border-radius: 16px;
            color: white;
            margin-bottom: 18px;
            box-shadow: 0 8px 20px rgba(0,0,0,0.15);
        }}
        .hero-cp h1 {{ color: #FFFFFF !important; margin: 0; font-size: 1.9rem; }}
        .hero-cp p {{ color: #E2E8F0 !important; margin: 6px 0 0 0; font-size: 0.95rem; }}
        div[data-testid="stMetric"] {{
            background-color: #FFFFFF;
            border-left: 5px solid #006837;
            padding: 12px 14px;
            border-radius: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.05);
        }}
    </style>
    <div class="hero-cp">
        <h1>Escalas ORV CP</h1>
        <p>Gestão, Comparação e Simulação Nacional de Escalas &bull; <b>by Luís Maia</b></p>
    </div>
    """, unsafe_allow_html=True)

def verificar_acesso():
    if st.session_state.get("autenticado"):
        return True
    st.markdown("### 🔒 Acesso Reservado")
    senha_digitada = st.text_input("Insere a palavra-passe de acesso:", type="password")
    if st.button("Entrar", type="primary"):
        senha_correta = st.secrets.get("SENHA_APP", "COBR20")
        if senha_digitada == senha_correta:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("❌ Palavra-passe incorreta.")
    return False

def processar_comparacao_mobile(d_ant, d_nov, data_ini_str, titulo_esc="COBR20"):
    max_os = max(obter_total_os(d_ant), obter_total_os(d_nov))
    sems_ant = calcular_semanas_detalhado(d_ant, data_ini_str, max_os)
    sems_nov = calcular_semanas_detalhado(d_nov, data_ini_str, max_os)
    c_ant, c_nov, os_alt, os_gan, os_sup, rep_alt = {}, {}, 0, 0, 0, []
    lista_os_alt, lista_cb_mov, lista_rep_alt = [], [], []

    for os_n in range(1, max_os + 1):
        sa, sn = [x for x in d_ant if x["OS"] == os_n], [x for x in d_nov if x["OS"] == os_n]
        ta = "\n".join(f"- [{x['Dias']}] {x['Serv']} ({x['PNT'][:5]}){' [REP FORA: '+x['LocalRet']+']' if x['RepFora'] else ''}" for x in sa if x["Dias"] or x["Serv"])
        tn = "\n".join(f"- [{x['Dias']}] {x['Serv']} ({x['PNT'][:5]}){' [REP FORA: '+x['LocalRet']+']' if x['RepFora'] else ''}" for x in sn if x["Dias"] or x["Serv"])
        rep_a = sorted(set(f"{x['Dias_Efet']} ({x['LocalRet']})" for x in sa if x["RepFora"]))
        rep_n = sorted(set(f"{x['Dias_Efet']} ({x['LocalRet']})" for x in sn if x["RepFora"]))
        if rep_a != rep_n: rep_alt.append((os_n, rep_a, rep_n))
        for x in sa:
            for cb in comboios(x["Serv"]): c_ant.setdefault(cb, set()).add(os_n)
        for x in sn:
            for cb in comboios(x["Serv"]): c_nov.setdefault(cb, set()).add(os_n)
        if ta != tn:
            os_alt += 1
            if "SUPRA" not in ta and "SUPRA" in tn: tp, cor, os_sup = "PASSOU A SUPRA", (255, 235, 156), os_sup + 1
            elif "SUPRA" in ta and "SUPRA" not in tn: tp, cor, os_gan = "GANHOU SERVIÇO", (198, 239, 206), os_gan + 1
            else: tp, cor = "AJUSTE HORÁRIO / COMBOIO", (217, 225, 242)
            lista_os_alt.append((f"OS {os_n:03d}", tp, ta, tn, cor))

    cb_nov, cb_sup, cb_trf = 0, 0, 0
    for cb in sorted(set(c_ant) | set(c_nov), key=int):
        oa, on = sorted(c_ant.get(cb, set())), sorted(c_nov.get(cb, set()))
        if oa != on:
            saiu = ", ".join(f"OS {x}" for x in sorted(set(oa)-set(on))) or "-"
            entrou = ", ".join(f"OS {x}" for x in sorted(set(on)-set(oa))) or "-"
            if not oa: mv, cor, cb_nov = "NOVO NA ESCALA", (198, 239, 206), cb_nov + 1
            elif not on: mv, cor, cb_sup = "SUPRIMIDO", (255, 199, 206), cb_sup + 1
            else: mv, cor, cb_trf = "TRANSFERIDO", (255, 235, 156), cb_trf + 1
            lista_cb_mov.append((cb, mv, saiu, entrou, ", ".join(f"OS {x}" for x in oa) or "-", ", ".join(f"OS {x}" for x in on) or "-", cor))

    for os_n, ra, rn in rep_alt:
        if len(rn) > len(ra): st_r, cor = "GANHOU REPOUSO", (198, 239, 206)
        elif len(rn) < len(ra): st_r, cor = "PERDEU REPOUSO", (255, 199, 206)
        else: st_r, cor = "ALTEROU REP FORA", (217, 225, 242)
        lista_rep_alt.append((f"OS {os_n:03d}", st_r, ", ".join(ra) if ra else "Sem rep. fora", ", ".join(rn) if rn else "Sem rep. fora", cor))

    m_ant, m_nov = sum(sems_ant) / float(max_os), sum(sems_nov) / float(max_os)
    rep_tot_ant, rep_tot_nov = contar_repousos_escala(d_ant, max_os), contar_repousos_escala(d_nov, max_os)
    dif_rep, dm, scb, sos = rep_tot_nov - rep_tot_ant, m_nov - m_ant, len(c_nov) - len(c_ant), os_gan - os_sup
    txt_rep = f"GANHOU {dif_rep} repouso(s) (+{dif_rep})" if dif_rep > 0 else (f"PERDEU {abs(dif_rep)} repouso(s) ({dif_rep})" if dif_rep < 0 else "Manteve repousos (0)")
    cor_rep = (198, 239, 206) if dif_rep > 0 else ((255, 199, 206) if dif_rep < 0 else (217, 225, 242))

    balanco_rows = [
        ("Total de Repousos Fora da Sede (REP F)", f"{rep_tot_ant} -> {rep_tot_nov}", txt_rep, cor_rep),
        ("Carga Horária Média Semanal", f"{fmt_hm(m_ant)} -> {fmt_hm(m_nov)}", f"Diferença semanal: {fmt_hm(dm, True)}", (217, 225, 242)),
        ("Comboios distintos na Escala", f"{len(c_ant)} -> {len(c_nov)}", f"Novos: +{cb_nov} | Suprimidos: -{cb_sup} | Transf.: {cb_trf}", (245, 245, 245)),
        ("SALDO LÍQUIDO DE COMBOIOS", f"{scb:+d}", "Ganho/Perda líquida de circulações", (198, 239, 206) if scb>0 else ((255, 199, 206) if scb<0 else (217, 225, 242))),
        ("Total de OS alteradas", f"{os_alt} / {max_os}", f"Ganharam serviço: +{os_gan} | Passaram a SUPRA: -{os_sup}", (217, 225, 242)),
        ("SALDO LÍQUIDO DE OS COM SERVIÇO", f"{sos:+d}", "Balanço entre OS ganhas e SUPRA", (198, 239, 206) if sos>0 else ((255, 199, 206) if sos<0 else (217, 225, 242))),
    ]

    ind_ant, ind_nov = calcular_indicadores_dia(d_ant, max_os), calcular_indicadores_dia(d_nov, max_os)
    nomes_dias = ("Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado")
    dist_rows = []
    for d_idx in range(1, 8):
        s_a, s_n = ind_ant[d_idx]["SERVICOS"], ind_nov[d_idx]["SERVICOS"]
        u_a, u_n = ind_ant[d_idx]["SUPRA"], ind_nov[d_idx]["SUPRA"]
        desc_n, dif_s = ind_nov[d_idx]["DESCANSO"], s_n - s_a
        cor_l = (198, 239, 206) if dif_s > 0 else ((255, 199, 206) if dif_s < 0 else (248, 248, 248))
        dist_rows.append((nomes_dias[d_idx-1], f"{s_a} -> {s_n}", f"{dif_s:+d}", f"{u_a} -> {u_n}", str(desc_n), str(s_n+u_n+desc_n), cor_l))

    num_blocos = max(1, max_os // 12)
    sem_rows = []
    for b_idx in range(num_blocos):
        ini_s, fim_s = b_idx * 12, (b_idx + 1) * 12
        sa12, sn12 = sum(sems_ant[ini_s:fim_s]), sum(sems_nov[ini_s:fim_s])
        sem_rows.append((f"BLOCO {b_idx+1}", f"Sem {ini_s+1} a {fim_s}", fmt_hm(sa12), fmt_hm(sn12), fmt_hm(sn12 - sa12, True)))

    pdf_completo = gerar_pdf_nativo(lista_os_alt, lista_cb_mov, lista_rep_alt, balanco_rows, dist_rows, sem_rows, titulo_esc=titulo_esc, so_os=False)
    pdf_so_os = gerar_pdf_nativo(lista_os_alt, lista_cb_mov, lista_rep_alt, balanco_rows, dist_rows, sem_rows, titulo_esc=titulo_esc, so_os=True)
    resumo = {"rep_ant": rep_tot_ant, "rep_nov": rep_tot_nov, "dif_rep": dif_rep, "scb": scb, "sos": sos, "h_ant": fmt_hm(m_ant), "h_nov": fmt_hm(m_nov)}
    return pdf_completo, pdf_so_os, resumo

aplicar_visual_cp("alfa-pendular-wext3_856.jpg")
if not verificar_acesso():
    st.stop()

if "base_comboios" not in st.session_state:
    st.session_state["base_comboios"] = carregar_base_comboios()

aba_comp, aba_edit, aba_conv = st.tabs(["📊 Comparar Escalas", "✏️ Editor & Simulador de OS", "🔄 Converter PDF → Excel"])

with aba_comp:
    arq_ant = st.file_uploader("1. Escala Anterior (Excel ou PDF)", type=["xlsx", "pdf"], key="ant")
    arq_nov = st.file_uploader("2. Escala Nova (Excel ou PDF)", type=["xlsx", "pdf"], key="nov")
    data_ini = st.date_input("Início da Escala (Dom)", value=datetime(2026, 11, 22))

    if arq_nov:
        d_preview, meta_nov = ler_ficheiro_carregado(arq_nov)
        arq_nov.seek(0)
        with st.expander(f"👁️ Visualizar Escala: {meta_nov['Deposito']} ({meta_nov['Escala']} — {obter_total_os(d_preview)} OS)", expanded=True):
            filtro_os = st.text_input("Filtrar por OS ou Comboio:", placeholder="Ex: 117 ou 4670")
            df_prev = pd.DataFrame(d_preview)[["OS", "Dias", "LocalApr", "HoraApr", "Serv", "LocalRet", "HoraRet", "PNT", "RepFora"]]
            if filtro_os.strip():
                df_prev = df_prev[df_prev["OS"].astype(str).str.contains(filtro_os.strip()) | df_prev["Serv"].str.contains(filtro_os.strip(), case=False)]
            st.dataframe(df_prev, use_container_width=True, hide_index=True)

    if arq_ant and arq_nov:
        if st.button("🔍 Comparar Escalas", type="primary", use_container_width=True):
            with st.spinner("A processar escalas..."):
                d_a, _ = ler_ficheiro_carregado(arq_ant)
                d_n, meta_n = ler_ficheiro_carregado(arq_nov)
                p_comp, p_os, res = processar_comparacao_mobile(d_a, d_n, data_ini.strftime("%Y-%m-%d"), titulo_esc=meta_n["Escala"])
                st.session_state["mob_res"] = (p_comp, p_os, res, meta_n["Escala"])

    if "mob_res" in st.session_state:
        p_comp, p_os, res, cod_esc_res = st.session_state["mob_res"]
        st.success(f"Comparação concluída ({cod_esc_res})!")
        c1, c2 = st.columns(2)
        c1.metric("Repousos Fora", f"{res['rep_ant']} → {res['rep_nov']}", f"{res['dif_rep']:+d}")
        c2.metric("Saldo Comboios", f"{res['scb']:+d}")
        c3, c4 = st.columns(2)
        c3.metric("Saldo OS Serviço", f"{res['sos']:+d}")
        c4.metric("Média Horas/Sem", res["h_nov"], f"Antes: {res['h_ant']}")

        st.download_button("📕 Descarregar PDF Completo", data=p_comp, file_name=f"Comparacao_{cod_esc_res}.pdf", mime="application/pdf", use_container_width=True)
        st.download_button("📋 Descarregar PDF Só OS Alteradas", data=p_os, file_name=f"OS_Alteradas_{cod_esc_res}.pdf", mime="application/pdf", use_container_width=True)
with aba_edit:
    st.markdown("Edita as **OS pretendidas** usando os números dos comboios (ex: `722 - (TR) - 4674`). O sistema preenche horários via `comboios.csv` e audita **Repouso Sede (14h)**, **Fora (9h)**, **TR (45m)** e **PNT (6h-9h)**.")
    with st.expander("🚆 Gerir Horários de Comboios (comboios.csv)", expanded=False):
        c_g1, c_g2, c_g3, c_g4, c_g5 = st.columns(5)
        n_cb = c_g1.text_input("Comboio", placeholder="722")
        n_or = c_g2.text_input("Origem", placeholder="AVO")
        n_pa = c_g3.text_input("Partida", placeholder="15:56")
        n_de = c_g4.text_input("Destino", placeholder="COB")
        n_ch = c_g5.text_input("Chegada", placeholder="16:30")
        if st.button("➕ Guardar Comboio na Sessão"):
            if n_cb.strip() and n_pa.strip() and n_ch.strip():
                st.session_state["base_comboios"][n_cb.strip()] = {"Origem": n_or.strip() or "EST", "Partida": n_pa.strip(), "Destino": n_de.strip() or "EST", "Chegada": n_ch.strip()}
                st.success(f"Comboio {n_cb} adicionado!")
        pdf_hor = st.file_uploader("Ou extrair comboios de um PDF de Horários CP:", type=["pdf"], key="pdf_hor")
               if pdf_hor and st.button("📥 Extrair Horários do PDF"):
            ext = extrair_horarios_de_pdf_cp(pdf_hor)
            st.session_state["base_comboios"].update(ext)
            st.success(f"Extraídos {len(ext)} comboios do PDF!")

    arq_sim = st.file_uploader("Carrega a Escala Base para Editar (.xlsx ou .pdf)", type=["xlsx", "pdf"], key="sim")
    if arq_sim:

        tot_os_sim = obter_total_os(d_sim)
        os_sel = st.multiselect("Seleciona as OS que queres editar:", options=list(range(1, tot_os_sim + 1)), default=[min(117, tot_os_sim)])
        if os_sel:
            # Podes encontrar a lógica completa de filtragem, edição de tabela, auditoria de infrações e geração de PDF no repositório do ficheiro app_mobile.py
            pass
with aba_conv:
    st.markdown("Converte qualquer **PDF oficial ORV da CP** para **Excel (`.xlsx`)** mantendo a estrutura exata das 11 colunas e validação automática.")
    pdf_conv = st.file_uploader("Carrega o PDF da Escala CP", type=["pdf"], key="conv")
    if pdf_conv:
        if st.button("⚙️ Converter PDF → Excel (.xlsx)", type="primary", use_container_width=True):
            with st.spinner("A extrair e validar as OS..."):
                xlsx_bytes, lidas, tot_esp, em_falta = converter_pdf_para_excel_validado(pdf_conv)
                st.session_state["conv_res"] = (xlsx_bytes, lidas, tot_esp, em_falta, pdf_conv.name.replace(".pdf", ".xlsx"))

    if "conv_res" in st.session_state:
        xlsx_bytes, lidas, tot_esp, em_falta, nome_xlsx = st.session_state["conv_res"]
        if lidas == tot_esp and not em_falta:
            st.success(f"✅ Validação 100% OK: Todas as {tot_esp} OS foram convertidas com estrutura oficial CP!")
        else:
            st.warning(f"⚠️ Lidas {lidas}/{tot_esp} OS. Em falta: {em_falta}")
        st.download_button("📗 Descarregar Excel Oficial (.xlsx)", data=xlsx_bytes, file_name=nome_xlsx, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

