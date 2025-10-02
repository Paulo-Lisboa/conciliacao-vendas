import streamlit as st
import pandas as pd
from io import BytesIO

# ============================
# Função para ler arquivos flexíveis
# ============================
def ler_arquivo_flexivel(file, tipo="operadora"):
    if file is None:
        return None
    
    if tipo == "operadora":
        melhor_df = None
        melhor_cols = -1
        for skip in range(0, 8):
            try:
                df = pd.read_excel(file, dtype=str, skiprows=skip, engine="openpyxl")
                # Número de colunas não-nulas
                non_null_cols = df.notna().sum().sum()
                if non_null_cols > melhor_cols:
                    melhor_cols = non_null_cols
                    melhor_df = df
            except Exception:
                continue
        return melhor_df
    else:
        # PDV: leitura direta
        return pd.read_excel(file, dtype=str, engine="openpyxl")

# ============================
# Função para tratar o PDV
# (replicar campos de vendas em casos de múltiplos pagamentos)
# ============================
def tratar_pdv(df_pdv):
    df_pdv = df_pdv.copy()

    # Lista de colunas que precisam ser replicadas
    colunas_para_preencher = [
        "numero_venda", "IdUnico", "codfilial", "terminal", "nome terminal",
        "operador", "status cupom", "Cupom Fiscal", "Serie", "TipoNotaFiscal",
        "DocEmitido", "ChaveXML", "CodStatusMigrate", "DescrStatusMigrate",
        "IdTransação"
    ]

    if "DataPgto" in df_pdv.columns and "HoraPgto" in df_pdv.columns:
        # Agrupar por DataPgto + HoraPgto
        grupos = df_pdv.groupby(["DataPgto", "HoraPgto"])
        for (data, hora), grupo in grupos:
            for col in colunas_para_preencher:
                if col in grupo.columns:
                    valores_validos = grupo[col].dropna().unique()
                    if len(valores_validos) == 1:  # só replica se houver 1 valor válido
                        valor = valores_validos[0]
                        mask = (df_pdv["DataPgto"] == data) & (df_pdv["HoraPgto"] == hora)
                        df_pdv.loc[mask, col] = df_pdv.loc[mask, col].fillna(valor)

    return df_pdv

# ============================
# Função de Conciliação
# ============================
def conciliar(df_op, df_pdv, chave_op, chave_pdv):
    df_op = df_op.copy()
    df_pdv = df_pdv.copy()

    # Renomear colunas temporárias para chave
    df_op = df_op.rename(columns={chave_op: "chave"})
    df_pdv = df_pdv.rename(columns={chave_pdv: "chave"})

    # Limpeza de espaços/brancos
    df_op["chave"] = df_op["chave"].astype(str).str.strip()
    df_pdv["chave"] = df_pdv["chave"].astype(str).str.strip()

    # Merge externo
    df_final = pd.merge(
        df_pdv, df_op, on="chave", how="outer", indicator=True, suffixes=("_PDV", "_OPERADORA")
    )

    # Status da Conciliação
    status_map = {
        "both": "Presente nos dois",
        "left_only": "Presente no PDV",
        "right_only": "Presente na Operadora"
    }
    df_final["Status Conciliação"] = df_final["_merge"].map(status_map)
    df_final = df_final.drop(columns=["_merge"])

    return df_final

# ============================
# Streamlit App
# ============================
st.set_page_config(page_title="Conciliação de Vendas", layout="wide")

st.title("📊 Conciliação de Vendas - PDV x Operadora")

# Upload dos arquivos
file_op = st.file_uploader("Upload Arquivo da Operadora", type=["xlsx"])
file_pdv = st.file_uploader("Upload Arquivo do PDV", type=["xlsx"])

if file_op and file_pdv:
    df_op = ler_arquivo_flexivel(file_op, tipo="operadora")
    df_pdv = ler_arquivo_flexivel(file_pdv, tipo="pdv")

    # Aplicar regra de preenchimento no PDV
    df_pdv = tratar_pdv(df_pdv)

    if df_op is not None and df_pdv is not None:
        st.success("✅ Arquivos carregados com sucesso!")

        # Escolha de colunas para conciliação
        chave_op = st.selectbox("Selecione a coluna da Operadora (chave de conciliação)", df_op.columns)
        chave_pdv = st.selectbox("Selecione a coluna do PDV (chave de conciliação)", df_pdv.columns)

        if st.button("Executar Conciliação"):
            df_final = conciliar(df_op, df_pdv, chave_op, chave_pdv)

            st.subheader("Resultado da Conciliação")
            st.dataframe(df_final, use_container_width=True)

            # Botão de download
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Conciliação")
            st.download_button(
                label="📥 Baixar Excel Consolidado",
                data=buffer.getvalue(),
                file_name="consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
