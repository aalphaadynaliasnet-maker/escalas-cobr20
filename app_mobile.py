import streamlit as st
import re, io, base64, warnings, openpyxl, pdfplumber
from datetime import datetime, timedelta
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from fpdf import FPDF

warnings.filterwarnings("ignore")
st.set_page_config(page_title="Escalas COBR20 - by Luís Maia", page_icon="alfa-pendular-wext3_856.jpg", layout="wide")

FERIADOS = {
    "2026-01-01","2026-04-03","2026-04-05","2026-04-25","2026-05-01","2026-06-04","2026-06-10","2026-07-04","2026-08-15","2026-10-05","2026-11-01","2026-12-01","2026-12-08","2026-12-25",
    "2027-01-01","2027-03-26","2027-03-28","2027-04-25","2027-05-01","2027-05-27","2027-06-10","2027-07-04","2027-08-15","2027-10-05","2027-11-01","2027-12-01","2027-12-08","2027-12-25",
    "2028-01-01","2028-04-14","2028-04-16","2028-04-25","2028-05-01","2028-06-10","2028-06-15","2028-07-04","2028-08-15","2028-10-05","2028-11-01","2028-12-01","2028-12-08","2028-12-25",
    "2029-01-01","2029-03-30","2029-04-01","2029-04-25","2029-05-01","2029-05-31","2029-06-10","2029-07-04","2029-08-15","2029-10-05","2029-11-01","2029-12-01","2029-12-08","2029-12-25"
}

def processar_linhas_brutas(linhas):
    dados, os_atual, dias_atual = [], None, ""
    col_os, col_dias, col_serv, col_lret, col_pnt, col_repf = 0, 1, 4, 6, 8, 10
    lixo = ("EM VIGOR", "ALTERAÇÃO", "CATEGORIA", "DEPÓSITO", "COIMBRA", "Planeamento", "Versão electrónica", "F.500.02", "Revisão e Venda", "Local", "Hora")
    for r in linhas:
        r = [re.sub(r'\s+', ' ', str(c)).strip() if c is not None else "" for c in r]
        if not any(r) or any(p in " ".join(r) for p in lixo): continue
        if any("Dias da Semana" in c for c in r):
            for idx, txt in enumerate(r):
                if txt == "OS": col_os = idx
                elif "Dias da Semana" in txt: col_dias = idx
                elif "Designação" in txt: col_serv = idx
                elif "PNT" in txt: col_pnt = idx
            col_lret, col_repf = max(0, col_pnt - 2), col_pnt + 2
            continue
        if len(r) <= col_pnt: continue
        c_os = r.__getitem__(col_os)
        if c_os.replace(".0", "").isdigit() and 1 <= int(float(c_os)) <= 120:
            os_atual = int(float(c_os))
        c_dias = r.__getitem__(col_dias)
        if c_dias and "Dias" not in c_dias: dias_atual = c_dias
        if os_atual:
            serv = r.__getitem__(col_serv) if len(r) > col_serv else ""
            loc_ret = r.__getitem__(col_lret) if len(r) > col_lret else ""
            pnt = r.__getitem__(col_pnt) if len(r) > col_pnt else ""
            rep_f_col = r.__getitem__(col_repf) if len(r) > col_repf else ""
            if c_dias or serv or pnt:
                dados.append({"OS": os_atual, "Dias": c_dias, "Dias_Efet": dias_atual, "Serv": serv, "LocalRet": loc_ret, "PNT": pnt, "RepFora": rep_f_col.lower() == "x"})
    return dados

def ler_ficheiro_carregado(uploaded_file):
    nome, bytes_data = uploaded_file.name.lower(), uploaded_file.read()
    if nome.endswith(".pdf"):
        todas_linhas = []
        with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
            for pagina in pdf.pages:
                for tabela in pagina.extract_tables(): todas_linhas.extend(tabela)
        return processar_linhas_brutas(todas_linhas)
    wb = openpyxl.load_workbook(io.BytesIO(bytes_data), data_only=True)
    for sheet in wb.sheetnames:
        linhas = [["" if c is None else str(c).strip() for c in r] for r in wb[sheet].iter_rows(values_only=True)]
        if any(any("Dias da Semana" in c for c in l) for l in linhas): return processar_linhas_brutas(linhas)
    return []

def contar_repousos_escala(dados):
    return sum(len(set(x["Dias_Efet"] for x in dados if x["OS"] == os_n and x["RepFora"])) for os_n in range(1, 121))

def expandir(regra):
    pares = set()
    for b in str(regra).split():
        m = re.match(r'^([0-9\.]+)([NVFS]*)$', b)
        if not m: continue
        dt, tt = m.groups()
        tipos = list(tt) if tt else ['N','V','F','S']
        dias = range(int(dt.split('..').__getitem__(0)), int(dt.split('..').__getitem__(1)) + 1) if '..' in dt else [int(dt)]
        for d in dias:
            for t in tipos: pares.add((d, t))
    return pares

def calcular_semanas_detalhado(dados, data_ini_str):
    regras_os = {os_n: [] for os_n in range(1, 121)}
    for os_n in range(1, 121):
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
    for sem in range(120):
        mins_sem = 0
        for d in range(7):
            dia_abs = sem * 7 + d
            dt = base + timedelta(days=dia_abs)
            de = 1 if dt.weekday() == 6 else dt.weekday() + 2
            tp = 'F' if dt.strftime("%Y-%m-%d") in FERIADOS else ('V' if (dt+timedelta(days=1)).strftime("%Y-%m-%d") in FERIADOS else ('S' if (dt-timedelta(days=1)).strftime("%Y-%m-%d") in FERIADOS else 'N'))
            for pv, mins in regras_os[(dia_abs % 120) + 1]:
                if (de, tp) in pv:
                    mins_sem += mins
                    break
        lista_semanas.append(mins_sem)
    return lista_semanas

def calcular_indicadores_dia(dados):
    regras_os = {os_n: [] for os_n in range(1, 121)}
    for os_n in range(1, 121):
        for x in [k for k in dados if k["OS"] == os_n]:
            reg = x["Dias_Efet"]
            if not reg: continue
            serv_txt = x["Serv"].upper()
            estado = "DESCANSO" if "DESCANSO" in serv_txt else ("SUPRA" if "SUPRA" in serv_txt else "SERVICOS")
            regras_os[os_n].append((expandir(reg), estado))
    contagem = {d: {"SERVICOS": 0, "SUPRA": 0, "DESCANSO": 0} for d in range(1, 8)}
    for d in range(1, 8):
        for os_n in range(1, 121):
            estado_encontrado = "DESCANSO"
            for pv, est in regras_os[os_n]:
                if (d, 'N') in pv:
                    estado_encontrado = est
                    break
            contagem[d][estado_encontrado] += 1
    return contagem

def fmt_hm(mins, com_sinal=False):
    s = "-" if mins < 0 else ("+" if com_sinal and mins > 0 else "")
    m = abs(int(round(mins)))
    return f"{s}{m//60:02d}h{m%60:02d}m"

def comboios(txt):
    limpo = re.sub(r'\(.*?\)', '', str(txt))
    limpo = re.sub(r'\b\d{2}/\d{2}/\d{4}\b', '', limpo)
    return set(re.findall(r'\b\d{3,5}\b', limpo))

def limpar_latin1(txt):
    return str(txt).replace("•", "-").replace("→", "->").encode("latin-1", "replace").decode("latin-1")

class PDFComAssinatura(FPDF):
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
        self.cell(0, 6, limpar_latin1(f"Escalas COBR20  |  by Luís Maia  |  Página {self.page_no()}/{{nb}}"), align="C")
def gerar_pdf_nativo(lista_os_alt, lista_cb_mov, lista_rep_alt, balanco_rows, dist_rows, sem_rows, so_os=False):
    pdf = PDFComAssinatura(orientation="L", unit="mm", format="A4")
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

    titulo_sec("Resumo de OS Alteradas - Escala COBR20" if so_os else "Comparação Completa de Escalas COBR20")
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
        pdf.cell(l_os.__getitem__(0), alt_linha, limpar_latin1(os_str), border=1, align="C", fill=True)
        pdf.cell(l_os.__getitem__(1), alt_linha, limpar_latin1(tp), border=1, align="C", fill=True)
        x_ta = pdf.get_x()
        pdf.set_font("Helvetica", "", 7.5)
        pdf.multi_cell(l_os.__getitem__(2), alt_linha / linhas_max, limpar_latin1(ta), border=1, fill=True)
        pdf.set_xy(x_ta + l_os.__getitem__(2), y_ini)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.multi_cell(l_os.__getitem__(3), alt_linha / linhas_max, limpar_latin1(tn), border=1, fill=True)
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
        pdf.cell(l_cb.__getitem__(0), 6.5, limpar_latin1(cb), border=1, align="C", fill=True)
        pdf.cell(l_cb.__getitem__(1), 6.5, limpar_latin1(mv), border=1, align="C", fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(l_cb.__getitem__(2), 6.5, limpar_latin1(saiu), border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_cb.__getitem__(3), 6.5, limpar_latin1(entrou), border=1, fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(l_cb.__getitem__(4), 6.5, limpar_latin1(oa), border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_cb.__getitem__(5), 6.5, limpar_latin1(on), border=1, fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("3. Alterações nos Repousos Fora da Sede (REP F)")
    l_rep = (22, 54, 100, 100)
    cabecalho_tabela(["OS", "Situação Repouso", "Repousos Antes", "Repousos Agora"], l_rep)
    for os_s, st_r, ra, rn, cor_rgb in lista_rep_alt:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_rep.__getitem__(0), 6.5, limpar_latin1(os_s), border=1, align="C", fill=True)
        pdf.cell(l_rep.__getitem__(1), 6.5, limpar_latin1(st_r), border=1, align="C", fill=True)
        pdf.set_font("Helvetica", "", 8)
        pdf.cell(l_rep.__getitem__(2), 6.5, limpar_latin1(ra), border=1, fill=True)
        pdf.set_font("Helvetica", "B", 8)
        pdf.cell(l_rep.__getitem__(3), 6.5, limpar_latin1(rn), border=1, fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("4. Balanço Final: Ganhos vs Perdas")
    l_bal = (115, 45, 116)
    cabecalho_tabela(["Indicador", "Antes -> Agora", "Impacto / Saldo"], l_bal)
    for ind, qtd, obs, cor_rgb in balanco_rows:
        pdf.set_fill_color(*cor_rgb)
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.cell(l_bal.__getitem__(0), 7, limpar_latin1(ind), border=1, fill=True)
        pdf.cell(l_bal.__getitem__(1), 7, limpar_latin1(qtd), border=1, align="C", fill=True)
        pdf.cell(l_bal.__getitem__(2), 7, limpar_latin1(obs), border=1, fill=True)
        pdf.ln()

    pdf.add_page()
    titulo_sec("5. Distribuição Diária da Escala (Serviços, SUPRA e Descanso)")
    l_dist = (48, 52, 36, 52, 44, 44)
    cabecalho_tabela(["Dia da Semana", "Serviços (Ant -> Nov)", "Dif.", "SUPRA (Ant -> Nov)", "Descanso", "Total OS"], l_dist)
    for d_row in dist_rows:
        pdf.set_fill_color(*d_row.__getitem__(6))
        pdf.set_font("Helvetica", "B", 8.5)
        for idx_c in range(6):
            pdf.cell(l_dist.__getitem__(idx_c), 7, limpar_latin1(d_row.__getitem__(idx_c)), border=1, align=("L" if idx_c==0 else "C"), fill=True)
        pdf.ln()

    pdf.ln(5)
    titulo_sec("6. Cômputo a cada 12 Semanas (Rotação 120 Semanas)")
    l_sem = (45, 55, 58, 58, 60)
    cabecalho_tabela(["Bloco", "Semanas", "Horas Anterior", "Horas Nova", "Diferença (12 Sem.)"], l_sem)
    for s_row in sem_rows:
        pdf.set_fill_color(255, 235, 156)
        pdf.set_font("Helvetica", "B", 8.5)
        for idx_c in range(5):
            pdf.cell(l_sem.__getitem__(idx_c), 7, limpar_latin1(s_row.__getitem__(idx_c)), border=1, align="C", fill=True)
        pdf.ln()

    return bytes(pdf.output())
def converter_pdf_para_excel_validado(uploaded_pdf):
    bytes_data = uploaded_pdf.read()
    todas_linhas = []
    data_vigor, num_alt = "22/11/2026", "62ª ALTERAÇÃO"
    with pdfplumber.open(io.BytesIO(bytes_data)) as pdf:
        txt_p1 = pdf.pages.__getitem__(0).extract_text() or ""
        m_dt = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', txt_p1)
        if m_dt: data_vigor = m_dt.group(1)
        m_al = re.search(r'(\d+ª\s+ALTERAÇÃO)', txt_p1)
        if m_al: num_alt = m_al.group(1)
        for pagina in pdf.pages:
            for tabela in pagina.extract_tables(): todas_linhas.extend(tabela)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "COBR20"
    verde_cp = PatternFill("solid", fgColor="1B4F3E")
    fonte_branca = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fonte_bold = Font(name="Calibri", size=9, bold=True, color="000000")
    fonte_norm = Font(name="Calibri", size=9, bold=False, color="000000")
    borda_fina = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    esq = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws.append(["EM VIGOR DESDE", "CATEGORIA", "", f"{num_alt} EM VIGOR A PARTIR DE:", "", "ESCALA", "DEPÓSITO DE REVISÃO", "", "", "", ""])
    ws.append(["31/05/2020", "Operador de Revisão e Venda", "", data_vigor, "", "COBR20", "COIMBRA B", "", "", "", ""])
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

    lixo = ("EM VIGOR", "ALTERAÇÃO", "CATEGORIA", "DEPÓSITO", "COIMBRA", "Planeamento", "Versão electrónica", "F.500.02", "Revisão e Venda", "Local", "Hora", "Dias da Semana")
    os_encontradas, os_atual, inicio_os_row = set(), None, None

    for r in todas_linhas:
        r = [re.sub(r'\s+', ' ', str(c)).strip() if c is not None else "" for c in r]
        if not any(r) or any(p in " ".join(r) for p in lixo) or len(r) < 10: continue
        linha_11 = (r + [""] * 11)[:11]
        c_os = linha_11.__getitem__(0)
        if c_os.isdigit() and 1 <= int(c_os) <= 120:
            if os_atual and inicio_os_row and ws.max_row > inicio_os_row:
                ws.merge_cells(start_row=inicio_os_row, start_column=1, end_row=ws.max_row, end_column=1)
            os_atual = int(c_os)
            os_encontradas.add(os_atual)
            inicio_os_row = ws.max_row + 1
            linha_11 = os_atual
        else:
            linha_11 = ""
        if os_atual:
            ws.append(linha_11)
            curr_r = ws.max_row
            for c_idx in range(1, 12):
                cell = ws.cell(row=curr_r, column=c_idx)
                cell.font = fonte_bold if c_idx == 1 else fonte_norm
                cell.alignment = esq if c_idx == 5 else centro
                cell.border = borda_fina

    if os_atual and inicio_os_row and ws.max_row > inicio_os_row:
        ws.merge_cells(start_row=inicio_os_row, start_column=1, end_row=ws.max_row, end_column=1)

    for col_letra, larg in {"A": 6, "B": 16, "C": 8, "D": 8, "E": 58, "F": 8, "G": 8, "H": 8, "I": 5, "J": 5, "K": 6}.items():
        ws.column_dimensions[col_letra].width = larg

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), len(os_encontradas), sorted(set(range(1, 121)) - os_encontradas)

def aplicar_visual_cp(img_path="alfa-pendular-wext3_856.jpg"):
    bg_css = ""
    try:
