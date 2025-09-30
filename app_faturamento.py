import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import os
import tempfile
import pdfplumber
import tabula
from rapidfuzz import fuzz, process

st.set_page_config(layout="wide", page_title="Conferência de Faturamento")

# ---------------------------
# Utils
# ---------------------------
def normalize_code(s: str):
    if pd.isna(s): return None
    s = str(s).strip()
    s = re.sub(r"[^\d\.\-]", "", s)  # mantém apenas dígitos, ponto, hífen
    return s

def parse_pdf_bytes_tabula(pdf_bytes: bytes, file_name="arquivo.pdf"):
    """ Extrai tabelas de PDF usando tabula-py (requer Java). """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name
    try:
        dfs = tabula.read_pdf(tmp_path, pages="all", multiple_tables=True, lattice=True)
        if not dfs:
            return pd.DataFrame()
        df = pd.concat(dfs, ignore_index=True)
        df["__source_file"] = file_name
        return df
    except Exception as e:
        st.error(f"Erro extraindo tabelas do PDF {file_name}: {e}")
        return pd.DataFrame()
    finally:
        os.remove(tmp_path)

def parse_excel_file_bytes_to_df(excel_bytes: bytes):
    try:
        xls = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=None)
        dfs = []
        for name, df in xls.items():
            df["__sheet_name"] = name
            dfs.append(df)
        return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
    except Exception as e:
        st.error(f"Erro lendo Excel: {e}")
        return pd.DataFrame()

def aggregate_codes_from_df(df, code_col_candidates=None, qty_col_candidates=None):
    if df is None or df.empty:
        return pd.DataFrame(columns=["codigo","descricao","quant_total"])
    dfc = df.copy()

    # detectar colunas
    code_col = None
    qty_col = None
    desc_col = None

    for cand in (code_col_candidates or ["NC","CÓDIGO","CODIGO","codigo","cod","procedimento"]):
        for c in dfc.columns:
            if cand.lower() in str(c).lower():
                code_col = c; break
        if code_col: break

    for cand in (qty_col_candidates or ["QUANT","QTD","QUANTIDADE","quantidade","qtde"]):
        for c in dfc.columns:
            if cand.lower() in str(c).lower():
                qty_col = c; break
        if qty_col: break

    for cand in ["DESCRIÇÃO","DESCRICAO","descricao","exame","procedimento","item","nome"]:
        for c in dfc.columns:
            if cand.lower() in str(c).lower():
                desc_col = c; break
        if desc_col: break

    rows = []
    for _, r in dfc.iterrows():
        codigo = normalize_code(r.get(code_col)) if code_col else None
        qty = None
        if qty_col and not pd.isna(r.get(qty_col)):
            try:
                qty = int(r.get(qty_col))
            except:
                pass
        desc = r.get(desc_col) if desc_col else None
        rows.append({"codigo": codigo, "descricao": desc, "quant_total": qty})

    out = pd.DataFrame(rows)
    if out.empty:
        return out

    out = out.groupby("codigo", dropna=False).agg({
        "descricao": lambda x: "; ".join(pd.unique(x.dropna().astype(str))) if x.notna().any() else None,
        "quant_total": lambda x: int(np.nansum([v for v in x if pd.notna(v)])) if any(pd.notna(x)) else None
    }).reset_index()
    return out

def to_excel_bytes(df_dict):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        writer.save()
    return out.getvalue()

def fuzzy_match_list(base_list, compare_list, cutoff=85):
    """Faz fuzzy matching entre duas listas de strings"""
    matches = []
    for item in base_list:
        best = process.extractOne(item, compare_list, scorer=fuzz.token_sort_ratio)
        if best and best[1] >= cutoff:
            matches.append({"item": item, "match": best[0], "score": best[1]})
        else:
            matches.append({"item": item, "match": None, "score": None})
    return pd.DataFrame(matches)

# ---------------------------
# Streamlit UI
# ---------------------------
st.title("📊 Conferência de Faturamento")

st.markdown("""
Faça upload da **planilha do sistema MV** (Excel) como referência principal, e depois envie os demais arquivos (PDF/Excel) contendo:
- Relatório de Produção (códigos e quantidades)
- Relatório Nominal (paciente + exame)
- Relatório de Faturamento / Locação
""")

with st.sidebar:
    st.header("Configurações")
    fuzzy_cutoff = st.slider("Fuzzy cutoff para nomes", 60, 100, 85)

# Uploads
mv_file = st.file_uploader("📂 Upload - Planilha MV (Excel)", type=["xlsx","xls"])
other_files = st.file_uploader("📂 Upload - Outros arquivos (PDF/Excel)", type=["pdf","xlsx","xls"], accept_multiple_files=True)

if not mv_file:
    st.warning("Envie a planilha MV para começar.")
    st.stop()

# Processa MV
try:
    mv_df = pd.read_excel(mv_file)
    st.write("Pré-visualização MV:")
    st.dataframe(mv_df.head(10))
except Exception as e:
    st.error(f"Erro ao abrir MV: {e}")
    st.stop()

# detectar colunas MV
cand_code = next((c for c in mv_df.columns if any(k in c.lower() for k in ["codigo","nc","proced","cod"])), None)
cand_qty  = next((c for c in mv_df.columns if any(k in c.lower() for k in ["quant","qtd","realiz"])), None)
cand_desc = next((c for c in mv_df.columns if any(k in c.lower() for k in ["descr","exame","proced"])), None)
cand_name = next((c for c in mv_df.columns if any(k in c.lower() for k in ["nome","paciente"])), None)

mv_work = mv_df.copy()
mv_work["codigo_norm"] = mv_work[cand_code].astype(str).apply(normalize_code) if cand_code else None
mv_work["quant_mv"] = pd.to_numeric(mv_work[cand_qty], errors="coerce").fillna(0).astype(int) if cand_qty else 0
if cand_desc: mv_work["desc_mv"] = mv_work[cand_desc].astype(str)
if cand_name: mv_work["nome_mv"] = mv_work[cand_name].astype(str)

agg_mv = mv_work.groupby("codigo_norm", dropna=False).agg({
    "desc_mv": lambda x: "; ".join(pd.unique(x.dropna().astype(str))) if x.notna().any() else None,
    "quant_mv": "sum"
}).reset_index().rename(columns={"codigo_norm":"codigo"})

# ---------------------------
# Processa demais arquivos
# ---------------------------
all_extracted = []
nominal_dfs = []

for f in other_files:
    st.write(f"➡️ Processando: {f.name}")
    b = f.read()
    if f.name.lower().endswith(".pdf"):
        df_pdf = parse_pdf_bytes_tabula(b, f.name)
        if not df_pdf.empty:
            st.write("Preview PDF extraído:")
            st.dataframe(df_pdf.head(10))
            # Detecta se é nominal
            if any("paciente" in str(c).lower() for c in df_pdf.columns):
                nominal_dfs.append(df_pdf)
            else:
                agg = aggregate_codes_from_df(df_pdf)
                if not agg.empty:
                    agg["fonte"] = f.name
                    all_extracted.append(agg)
    elif f.name.lower().endswith((".xlsx",".xls")):
        df_excel = parse_excel_file_bytes_to_df(b)
        st.write("Preview Excel:")
        st.dataframe(df_excel.head(10))
        if any("paciente" in str(c).lower() for c in df_excel.columns):
            nominal_dfs.append(df_excel)
        else:
            agg = aggregate_codes_from_df(df_excel)
            if not agg.empty:
                agg["fonte"] = f.name
                all_extracted.append(agg)

# ---------------------------
# Comparação MV x Produção
# ---------------------------
if all_extracted:
    combined = pd.concat(all_extracted, ignore_index=True).fillna({"codigo":None})
    summary = combined.groupby("codigo", dropna=False).agg({
        "descricao": lambda x: "; ".join(pd.unique(x.dropna().astype(str))) if x.notna().any() else None,
        "quant_total": "sum"
    }).reset_index()

    result = pd.merge(agg_mv, summary, how="outer", on="codigo", suffixes=("_mv","_fontes"))
    result["quant_mv"] = result["quant_mv"].fillna(0).astype(int)
    result["quant_total"] = result["quant_total"].fillna(0).astype(int)
    result["diff_quant"] = result["quant_mv"] - result["quant_total"]
    result["status"] = result["diff_quant"].apply(lambda d: "OK" if d==0 else "DIVERGENTE")

    st.subheader("📊 Comparação MV x Produção")
    st.dataframe(result.sort_values("diff_quant", ascending=False).head(200))
else:
    st.info("Nenhum relatório de produção encontrado.")

# ---------------------------
# Comparação Nominal (Pacientes)
# ---------------------------
if nominal_dfs and cand_name:
    st.subheader("🧑‍⚕️ Conferência Nominal de Pacientes")

    nominal = pd.concat(nominal_dfs, ignore_index=True)
    col_paciente = next((c for c in nominal.columns if "paciente" in str(c).lower()), None)
    lista_mv = mv_work["nome_mv"].dropna().unique().tolist()
    lista_nominal = nominal[col_paciente].dropna().unique().tolist()

    st.write(f"Total nomes MV: {len(lista_mv)} | Total nomes Relatório Nominal: {len(lista_nominal)}")

    df_matches = fuzzy_match_list(lista_mv, lista_nominal, cutoff=fuzzy_cutoff)
    st.write("🔎 Matching de pacientes (MV → Nominal):")
    st.dataframe(df_matches.head(50))

    ausentes = df_matches[df_matches["match"].isna()]
    if not ausentes.empty:
        st.warning("Pacientes no MV que não aparecem no Relatório Nominal:")
        st.dataframe(ausentes)

else:
    st.info("Nenhum relatório nominal detectado.")

# ---------------------------
# Exportar
# ---------------------------
st.subheader("💾 Exportar relatório")
export_name = st.text_input("Nome do arquivo", value="relatorio_conferencia")
if st.button("Gerar Excel"):
    dfs_export = {"mv_agregado": agg_mv}
    if all_extracted:
        dfs_export["comparacao"] = result
        dfs_export["fontes_agregadas"] = summary
    if nominal_dfs and cand_name:
        dfs_export["nominal_matching"] = df_matches
    excel_bytes = to_excel_bytes(dfs_export)
    st.download_button("Download .xlsx", data=excel_bytes,
                       file_name=f"{export_name}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
