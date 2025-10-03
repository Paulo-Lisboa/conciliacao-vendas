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
                non_null_cols = df.notna().sum().sum()
                if non_null_cols > melhor_cols:
                    melhor_cols = non_null_cols
                    melhor_df = df
            except Exception:
                continue
        return melhor_df
    else:
        return pd.read_excel(file, dtype=str, engine="openpyxl")

# ============================
# Função para tratar o PDV
# ============================
def tratar_pdv(df_pdv):
    df_pdv = df_pdv.copy()
    colunas_para_preencher = [
        "numero_venda", "IdUnico", "codfilial", "terminal", "nome terminal",
        "operador", "status cupom", "Cupom Fiscal", "Serie", "TipoNotaFiscal",
        "DocEmitido", "ChaveXML", "CodStatusMigrate", "DescrStatusMigrate",
        "IdTransação"
    ]

    if "DataPgto" in df_pdv.columns and "HoraPgto" in df_pdv.columns:
        grupos = df_pdv.groupby(["DataPgto", "HoraPgto"])
        for (data, hora), grupo in grupos:
            for col in colunas_para_preencher:
                if col in grupo.columns:
                    valores_validos = grupo[col].dropna().unique()
                    if len(valores_validos) == 1:
                        valor = valores_validos[0]
                        mask = (df_pdv["DataPgto"] == data) & (df_pdv["HoraPgto"] == hora)
                        df_pdv.loc[mask, col] = df_pdv.loc[mask, col].fillna(valor)

    return df_pdv

# ============================
# Função de Conciliação (multi-colunas com ordenação final)
# ============================
def conciliar(df_op, df_pdv, cols_op, cols_pdv):
    df_op = df_op.copy()
    df_pdv = df_pdv.copy()

    # Criar chave concatenando colunas escolhidas
    df_op["chave"] = df_op[cols_op].astype(str).agg(" | ".join, axis=1)
    df_pdv["chave"] = df_pdv[cols_pdv].astype(str).agg(" | ".join, axis=1)

    df_op["chave"] = df_op["chave"].str.strip()
    df_pdv["chave"] = df_pdv["chave"].str.strip()

    df_final = pd.merge(
        df_pdv, df_op, on="chave", how="outer", indicator=True, suffixes=("_PDV", "_OPERADORA")
    )

    # Mapeamento do status
    status_map = {
        "both": "Presente nos dois",
        "right_only": "Presente na Operadora",
        "left_only": "Presente no PDV"
    }
    df_final["Status Conciliação"] = df_final["_merge"].map(status_map)
    df_final = df_final.drop(columns=["_merge"])

    # Reordenar: primeiro matches, depois operadora, depois PDV
    ordem = {
        "Presente nos dois": 0,
        "Presente na Operadora": 1,
        "Presente no PDV": 2
    }
    df_final["ordem"] = df_final["Status Conciliação"].map(ordem)
    df_final = df_final.sort_values(by="ordem").drop(columns=["ordem"]).reset_index(drop=True)

    return df_final

# ============================
# Streamlit App
# ============================
st.set_page_config(page_title="Conciliação de Vendas", layout="wide")

st.title("📊 Conciliação de Vendas - PDV x Operadora")

file_op = st.file_uploader("Upload Arquivo da Operadora", type=["xlsx"])
file_pdv = st.file_uploader("Upload Arquivo do PDV", type=["xlsx"])

if file_op and file_pdv:
    df_op = ler_arquivo_flexivel(file_op, tipo="operadora")
    df_pdv = ler_arquivo_flexivel(file_pdv, tipo="pdv")
    df_pdv = tratar_pdv(df_pdv)

    if df_op is not None and df_pdv is not None:
        st.success("✅ Arquivos carregados com sucesso!")

        # Multiselect para múltiplas colunas
        cols_op = st.multiselect("Selecione colunas da Operadora (chave de conciliação)", df_op.columns)
        cols_pdv = st.multiselect("Selecione colunas do PDV (chave de conciliação)", df_pdv.columns)

        if st.button("Executar Conciliação") and cols_op and cols_pdv:
            df_final = conciliar(df_op, df_pdv, cols_op, cols_pdv)

            st.subheader("Resultado da Conciliação")
            st.dataframe(df_final, use_container_width=True)

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df_final.to_excel(writer, index=False, sheet_name="Conciliação")
            st.download_button(
                label="📥 Baixar Excel Consolidado",
                data=buffer.getvalue(),
                file_name="consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
