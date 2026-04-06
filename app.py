import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="Ponto de Situação do Inquérito", page_icon="📊", layout="wide")

ILHAS = {1:"Santo Antão",2:"São Vicente",3:"São Nicolau",4:"Sal",5:"Boavista",6:"Maio",7:"Santiago",8:"Fogo",9:"Brava"}
CONCELHOS = { 11: "Ribeira Grande", 12: "Paul", 13: "Porto Novo",21: "São Vicente",31: "Ribeira Brava", 32: "Tarrafal de São Nicolau",41: "Sal",51: "Boavista",61: "Maio",71: "Tarrafal", 72: "Santa Catarina", 73: "Santa Cruz",74: "Praia", 75: "S. Domingos", 76: "S. Miguel",77: "S. Salvador do Mundo", 78: "S. Lourenço dos Órgãos", 79: "Ribeira Grande de Santiago",81: "Mosteiro", 82: "São Filipe",83: "Santa Catarina do Fogo",91: "Brava"}

st.markdown('<h1 style="color:#1a3c6e;border-bottom:3px solid #1a3c6e">📊 Ponto de Situação — Monitorização do Inquérito no Terreno</h1>', unsafe_allow_html=True)

uploaded = st.file_uploader("📂 Carregar ficheiro de extração (Excel .xlsx)", type=["xlsx"])
if uploaded is None:
    st.info("👆 Carregue o ficheiro de extração para visualizar o ponto de situação.")
    st.stop()

@st.cache_data
def load_data(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), sheet_name="Alojamento")
    df["agente"] = df["USER_CREATE"].astype(str).replace("None","Desconhecido").str.split("@").str[0].str.replace("."," ",regex=False).str.title()
    df["ilha_nome"] = df["cod_ilha"].map(ILHAS).fillna(df["cod_ilha"].astype(str))
    df["concelho_nome"] = df["cod_concelho"].map(CONCELHOS).fillna(df["cod_concelho"].astype(str))
    df["ponto_valido"]   = (df["AA0200"]==1).astype(int)
    df["ponto_invalido"] = (df["AA0200"]!=1).astype(int)
    df["res_habitual"]   = (df["AA0302"]==1).astype(int)
    df["secundaria"]     = (df["AA0302"]==2).astype(int)
    df["vazio"]          = (df["AA0302"]==3).astype(int)
    df["outros_fins"]    = (df["AA0302"]==4).astype(int)
    df["inacessivel"]    = (df["AA0302"]==5).astype(int)
    df["outra_situacao"] = (df["AA0302"]==6).astype(int)
    df["recusa"]         = (df["AA0401"]==7).astype(int)
    df["agreg_inquiridos"] = (df["AA0605"] == 1).astype(int)
    return df

df = load_data(uploaded.read())

col_f1, col_f2 = st.columns(2)
with col_f1:
    ilha_sel = st.selectbox("🏝️ Selecionar Ilha", ["Todas"] + sorted(df["ilha_nome"].unique()))
df_filt = df if ilha_sel=="Todas" else df[df["ilha_nome"]==ilha_sel]
with col_f2:
    concelho_sel = st.selectbox("🏙️ Selecionar Concelho", ["Todos"] + sorted(df_filt["concelho_nome"].unique()))
if concelho_sel!="Todos":
    df_filt = df_filt[df_filt["concelho_nome"]==concelho_sel]

st.divider()

total_aloj    = len(df_filt)
total_validos = int(df_filt["ponto_valido"].sum())
total_inv     = int(df_filt["ponto_invalido"].sum())
total_agentes = df_filt["agente"].nunique()


k1,k2,k3,k4 = st.columns(4)
k1.metric("🏠 Total Alojamentos", f"{total_aloj:,}")
k2.metric("✅ Pontos Válidos",    f"{total_validos:,}")
k3.metric("❌ Pontos Inválidos",  f"{total_inv:,}")
k4.metric("👤 Agentes Ativos",   f"{total_agentes}")


st.divider()
st.subheader("📋 Quadro de Pontos por Agente de Inquirição")

resumo = df_filt.groupby(["ilha_nome","concelho_nome","agente"]).agg(
    Total=("REFERENCIA","count"),
    Pontos_Validos=("ponto_valido","sum"),
    Pontos_Invalidos=("ponto_invalido","sum"),
    Res_Habitual=("res_habitual","sum"),
    Secundaria_Sazonal=("secundaria","sum"),
    Vazio=("vazio","sum"),
    Outros_Fins=("outros_fins","sum"),
    Inacessivel=("inacessivel","sum"),
    Outra_Situacao=("outra_situacao","sum"),
    Recusa=("recusa","sum"),
    agreg_inquiridos=("agreg_inquiridos","sum"),
).reset_index()

linha_total = pd.DataFrame([{"ilha_nome":"TOTAL","concelho_nome":"","agente":"",
    "Total":resumo["Total"].sum(),"Pontos_Validos":resumo["Pontos_Validos"].sum(),
    "Pontos_Invalidos":resumo["Pontos_Invalidos"].sum(),"Res_Habitual":resumo["Res_Habitual"].sum(),
    "Secundaria_Sazonal":resumo["Secundaria_Sazonal"].sum(),"Vazio":resumo["Vazio"].sum(),
    "Outros_Fins":resumo["Outros_Fins"].sum(),"Inacessivel":resumo["Inacessivel"].sum(),
    "Outra_Situacao":resumo["Outra_Situacao"].sum(),"Recusa":resumo["Recusa"].sum(),"agreg_inquiridos":resumo["agreg_inquiridos"].sum()}])

resumo_display = pd.concat([resumo, linha_total], ignore_index=True)
resumo_display.columns = ["Ilha","Concelho","Agente","Total","Pontos Válidos","Pontos Inválidos","Residência Habitual","Secundária/Sazonal","Vazio","Ocupados Outros Fins","Alojamento Inacessível","Outra Situação","Recusa","Agregados Inquiridos"]

def highlight_total(row):
    if row["Ilha"]=="TOTAL":
        return ["background-color:#1a3c6e;color:white;font-weight:bold"]*len(row)
    return [""]*len(row)

st.dataframe(resumo_display.style.apply(highlight_total,axis=1), use_container_width=True, hide_index=True, height=400)

st.divider()
st.subheader("⬇️ Exportar Dados")

col_d1, col_d2 = st.columns(2)

# Download do quadro resumo (já filtrado por ilha/concelho)
@st.cache_data
def to_excel_resumo(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Resumo")
    return output.getvalue()

# Download dos dados brutos filtrados
@st.cache_data
def to_excel_bruto(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
    return output.getvalue()

# Nome do ficheiro dinâmico conforme filtro
if ilha_sel == "Todas" and concelho_sel == "Todos":
    nome_ficheiro = "ponto_situacao_geral"
elif concelho_sel != "Todos":
    nome_ficheiro = f"ponto_situacao_{concelho_sel.replace(' ', '_')}"
else:
    nome_ficheiro = f"ponto_situacao_{ilha_sel.replace(' ', '_')}"

with col_d1:
    st.download_button(
        "📊 Exportar Quadro Resumo (Excel)",
        data=to_excel_resumo(resumo_display),
        file_name=f"{nome_ficheiro}_resumo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
with col_d2:
    # Resumo agregado por ilha/concelho (sem descer ao nível do agente)
    resumo_territorio = df_filt.groupby(["ilha_nome", "concelho_nome"]).agg(
        Total=("REFERENCIA", "count"),
        Pontos_Validos=("ponto_valido", "sum"),
        Pontos_Invalidos=("ponto_invalido", "sum"),
        Res_Habitual=("res_habitual", "sum"),
        Secundaria_Sazonal=("secundaria", "sum"),
        Vazio=("vazio", "sum"),
        Outros_Fins=("outros_fins", "sum"),
        Inacessivel=("inacessivel", "sum"),
        Outra_Situacao=("outra_situacao", "sum"),
        Recusa=("recusa", "sum"),
        Agregados_Inquiridos=("agreg_inquiridos", "sum"),
    ).reset_index()

    # Linha de total
    linha_total_terr = pd.DataFrame([{
        "ilha_nome": "TOTAL", "concelho_nome": "",
        "Total": resumo_territorio["Total"].sum(),
        "Pontos_Validos": resumo_territorio["Pontos_Validos"].sum(),
        "Pontos_Invalidos": resumo_territorio["Pontos_Invalidos"].sum(),
        "Res_Habitual": resumo_territorio["Res_Habitual"].sum(),
        "Secundaria_Sazonal": resumo_territorio["Secundaria_Sazonal"].sum(),
        "Vazio": resumo_territorio["Vazio"].sum(),
        "Outros_Fins": resumo_territorio["Outros_Fins"].sum(),
        "Inacessivel": resumo_territorio["Inacessivel"].sum(),
        "Outra_Situacao": resumo_territorio["Outra_Situacao"].sum(),
        "Recusa": resumo_territorio["Recusa"].sum(),
        "Agregados_Inquiridos": resumo_territorio["Agregados_Inquiridos"].sum(),
    }])

    resumo_territorio = pd.concat([resumo_territorio, linha_total_terr], ignore_index=True)
    resumo_territorio.columns = [
        "Ilha", "Concelho", "Total", "Pontos Válidos", "Pontos Inválidos",
        "Residência Habitual", "Secundária/Sazonal", "Vazio",
        "Ocupados Outros Fins", "Alojamento Inacessível",
        "Outra Situação", "Recusa", "Agregados Inquiridos"
    ]

    st.download_button(
        "📋 Exportar Resumo por Ilha/Concelho (Excel)",
        data=to_excel_resumo(resumo_territorio),
        file_name=f"{nome_ficheiro}_por_territorio.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    

    st.subheader("📊 Análise Visual")
tab1, tab2 = st.tabs(["Progresso por Agente","Distribuição de Alojamentos"])

with tab1:
    resumo_sorted = resumo.sort_values("Pontos_Validos", ascending=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=resumo_sorted["Pontos_Validos"],y=resumo_sorted["agente"],orientation="h",name="Válidos",marker_color="#1a3c6e",text=resumo_sorted["Pontos_Validos"],textposition="auto"))
    fig.add_trace(go.Bar(x=resumo_sorted["Pontos_Invalidos"],y=resumo_sorted["agente"],orientation="h",name="Inválidos",marker_color="#e57373",text=resumo_sorted["Pontos_Invalidos"],textposition="auto"))
    fig.update_layout(barmode="stack",title="Pontos por Agente (Válidos vs Inválidos)",xaxis_title="Nº de Alojamentos",yaxis_title="Agente",height=400)
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    labels=["Res. Habitual","Secundária/Sazonal","Vazio","Outros Fins","Inacessível","Outra Situação","Recusa","Agregados Inquiridos"]
    values=[int(df_filt["res_habitual"].sum()),int(df_filt["secundaria"].sum()),int(df_filt["vazio"].sum()),int(df_filt["outros_fins"].sum()),int(df_filt["inacessivel"].sum()),int(df_filt["outra_situacao"].sum()),int(df_filt["recusa"].sum()),int(df_filt["agreg_inquiridos"].sum())]
    fig2 = go.Figure(go.Pie(labels=labels,values=values,marker=dict(colors=["#1a3c6e","#2e86ab","#a8dadc","#f4a261","#e76f51","#e9c46a","#e57373"]),hole=0.4,textinfo="label+percent"))
    fig2.update_layout(title="Distribuição por Tipo de Alojamento",height=400)
    st.plotly_chart(fig2, use_container_width=True)



st.divider()
st.caption("Sistema de Monitorização do Inquérito — INE Cabo Verde")