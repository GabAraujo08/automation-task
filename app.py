import streamlit as st
import holidays
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import date, timedelta
import io

# ─── LÓGICA DE DATAS (igual ao script anterior) ──────────────────────────────
FERIADOS_BR = holidays.country_holidays("BR")

def eh_dia_util(d):
    return d.weekday() < 5 and d not in FERIADOS_BR

def proximo_dia_util(d):
    while not eh_dia_util(d):
        d += timedelta(days=1)
    return d

def adicionar_dias_uteis(inicio, n):
    atual, cont = inicio, 0
    while cont < n:
        atual += timedelta(days=1)
        if eh_dia_util(atual):
            cont += 1
    return atual

def proxima_segunda_ate_dia_20(ref):
    d = ref
    while d.weekday() != 0:
        d += timedelta(days=1)
    if d.day > 20:
        mes = d.month % 12 + 1
        ano = d.year + (1 if d.month == 12 else 0)
        d = date(ano, mes, 1)
        while d.weekday() != 0:
            d += timedelta(days=1)
    return d

def fmt(d):
    return d.strftime("%d/%m")

def fmt_periodo(ini, fim):
    return fmt(ini) if ini == fim else f"{fmt(ini)} a {fmt(fim)}"

# ─── GERAÇÃO DO EXCEL EM MEMÓRIA ─────────────────────────────────────────────
def gerar_excel(data_rc, data_alinhamento):
    etapas = []

    etapas.append(("Recebimento de RC", data_rc, data_rc, "—"))
    etapas.append(("Alinhamento de Perfil", data_alinhamento, data_alinhamento, 1))

    ini = proximo_dia_util(data_alinhamento + timedelta(days=1))
    fim = adicionar_dias_uteis(ini, 6)
    etapas.append(("Divulgação da Vaga", ini, fim, 7))

    ini = proximo_dia_util(fim + timedelta(days=1))
    etapas.append(("Triagem dos Inscritos", ini, ini, 1))

    ini = proximo_dia_util(ini + timedelta(days=1))
    etapas.append(("Mapeamento dos Candidatos", ini, ini, 1))

    ini = proximo_dia_util(ini + timedelta(days=1))
    fim = adicionar_dias_uteis(ini, 1)
    etapas.append(("Entrevistas RH", ini, fim, 2))

    ini = proximo_dia_util(fim + timedelta(days=1))
    fim = adicionar_dias_uteis(ini, 1)
    etapas.append(("Apresentação Relatório Gestor", ini, fim, 2))

    ini = proximo_dia_util(fim + timedelta(days=1))
    fim = adicionar_dias_uteis(ini, 1)
    etapas.append(("Entrevista Gestor + RH", ini, fim, 2))

    ini = proximo_dia_util(fim + timedelta(days=1))
    fim = adicionar_dias_uteis(ini, 14)
    etapas.append(("Proposta + Processo de Admissão", ini, fim, "10 / 15"))

    ini_prev = proxima_segunda_ate_dia_20(proximo_dia_util(fim + timedelta(days=1)))
    fim_prev = adicionar_dias_uteis(ini_prev, 4)
    etapas.append(("Previsão de Início", ini_prev, fim_prev, "—"))

    # Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cronograma"

    h_fill  = PatternFill("solid", fgColor="1F4E79")
    h_font  = Font(bold=True, color="FFFFFF", size=11)
    alt     = PatternFill("solid", fgColor="D6E4F0")
    normal  = PatternFill("solid", fgColor="FFFFFF")
    div_fill= PatternFill("solid", fgColor="2E75B6")
    div_font= Font(bold=True, color="FFFFFF", size=10)
    center  = Alignment(horizontal="center", vertical="center")
    left    = Alignment(horizontal="left",   vertical="center")
    borda   = Border(*[Side(style="thin")] * 0,
                     left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"),  bottom=Side(style="thin"))

    for col, (h, w) in enumerate(zip(["Atividades","Período","Dias Úteis"],[35,22,12]), 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = h_font, h_fill, center, borda
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[1].height = 22

    for i, (ativ, ini, fim, du) in enumerate(etapas, 2):
        fill = div_fill if ativ == "Divulgação da Vaga" else (alt if i%2==0 else normal)
        font = div_font if ativ == "Divulgação da Vaga" else Font(size=10)
        for col, val in enumerate([ativ, fmt_periodo(ini, fim), str(du)], 1):
            c = ws.cell(row=i, column=col, value=val)
            c.fill, c.font, c.border = fill, font, borda
            c.alignment = left if col == 1 else center
        ws.row_dimensions[i].height = 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, etapas

# ─── INTERFACE STREAMLIT ──────────────────────────────────────────────────────
st.set_page_config(page_title="Cronograma R&S", page_icon="📋", layout="centered")

st.title("📋 Gerador de Cronograma — R&S")
st.markdown("Preencha as duas primeiras datas e baixe a planilha pronta.")

col1, col2 = st.columns(2)
with col1:
    data_rc = st.date_input("📥 Recebimento do RC", value=date.today(), format="DD/MM/YYYY")
with col2:
    data_ali = st.date_input("🎯 Alinhamento de Perfil", value=date.today(), format="DD/MM/YYYY")

if st.button("⚡ Gerar Cronograma", use_container_width=True, type="primary"):
    buf, etapas = gerar_excel(data_rc, data_ali)

    # Preview na tela
    st.success("Cronograma gerado com sucesso!")
    st.markdown("### 📅 Preview")
    for ativ, ini, fim, du in etapas:
        periodo = fmt_periodo(ini, fim)
        st.markdown(f"**{ativ}** — `{periodo}` — {du} dias úteis")

    # Botão de download
    nome = f"cronograma_{data_rc.strftime('%d%m%Y')}.xlsx"
    st.download_button(
        label="⬇️ Baixar planilha Excel",
        data=buf,
        file_name=nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
