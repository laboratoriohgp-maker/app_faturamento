import streamlit as st
import pandas as pd
import unicodedata
import os
from datetime import datetime
import matplotlib.pyplot as plt
import plotly.express as px
from fpdf import FPDF

st.set_page_config(layout="wide", page_title="App Faturamento - Avançado", page_icon="🔬")

# ---------------------- Utilitários ----------------------

def normalize_text(s):
    if pd.isna(s): return ""
    if not isinstance(s, str): s = str(s)
    s = s.strip().upper()
    s = unicodedata.normalize('NFKD', s)
    s = ''.join([c for c in s if not unicodedata.combining(c)])
    return s

def token_set_ratio(a, b):
    if not a and not b: return 100
    sa = set(a.split())
    sb = set(b.split())
    if not sa or not sb: return 0
    return int(len(sa.intersection(sb)) / len(sa.union(sb)) * 100)

# canonicalize exam using synonyms mapping (dict)
def canonical_exam(name, synonyms_map=None):
    if not name: return name
    n = normalize_text(name)
    if synonyms_map:
        if n in synonyms_map:
            return synonyms_map[n]
        # try exact token match
        for k,v in synonyms_map.items():
            if token_set_ratio(k, n) > 90:
                return v
    return n

# detect datetime col heuristics
def detect_datetime_col(df):
    for c in df.columns:
        low = c.lower()
        if "data" in low and "hora" in low:
            return c
    for c in df.columns:
        low = c.lower()
        if any(w in low for w in ["data","hora","timestamp","ts","date","time"]):
            return c
    return None

# find col
def find_col(df, keywords):
    cols = list(df.columns)
    low = [c.lower() for c in cols]
    for kw in keywords:
        for i, lc in enumerate(low):
            if kw in lc:
                return cols[i]
    return None

# ---------------------- Sidebar / config ----------------------
st.sidebar.title("Configuração")
uploaded_synonyms = st.sidebar.file_uploader("(Opcional) Upload synonyms.csv (colunas: synonym,canonical)", type=["csv"]) 
use_rapidfuzz = st.sidebar.checkbox("Usar rapidfuzz (se instalado) para matching mais veloz", value=False)

# matching weights
st.sidebar.markdown("**Pesos para matching**")
weight_patient = st.sidebar.slider("Peso nome paciente", 0.0, 1.0, 0.6)
weight_exam = st.sidebar.slider("Peso exame", 0.0, 1.0, 0.4)

# filters defaults
tolerance = st.sidebar.number_input("Tolerância (minutos) p/ atraso", min_value=0, value=130)
fuzzy_threshold = st.sidebar.slider("Threshold token-set p/ pareamento (0-100)", 30, 100, 60)

# date filter
st.sidebar.markdown("---")
start_date = st.sidebar.date_input("Data início (opcional)", value=None)
end_date = st.sidebar.date_input("Data fim (opcional)", value=None)

# ---------------------- Main UI ----------------------
st.title("🔬 App Faturamento — MV x Laboratório (Avançado)")
st.markdown("Use as abas abaixo.

**Importante:** carregue as planilhas MV e Laboratório.
Você pode também enviar um arquivo de sinônimos para melhorar o matching.")

# Uploads
col1, col2 = st.columns(2)
with col1:
    uploaded_mv = st.file_uploader("Planilha MV (solicitações)", type=["xlsx","xls","csv"], key='mv')
with col2:
    uploaded_lab = st.file_uploader("Planilha Laboratório (realizações)", type=["xlsx","xls","csv"], key='lab')

# load synonyms
syn_map = None
if uploaded_synonyms is not None:
    try:
        df_syn = pd.read_csv(uploaded_synonyms)
        # expect columns 'synonym' and 'canonical'
        syn_map = {normalize_text(r['synonym']): normalize_text(r['canonical']) for _, r in df_syn.iterrows()}
        st.sidebar.success(f"Loaded {len(syn_map)} synonyms")
    except Exception as e:
        st.sidebar.error(f"Erro lendo synonyms: {e}")

# Require uploads
if not (uploaded_mv and uploaded_lab):
    st.info("Carregue as duas planilhas para começar a análise.")
    st.stop()

# ---------------------- Read files ----------------------
try:
    mv = pd.read_excel(uploaded_mv) if str(uploaded_mv).lower().endswith(('xlsx','xls')) else pd.read_csv(uploaded_mv)
    lab = pd.read_excel(uploaded_lab) if str(uploaded_lab).lower().endswith(('xlsx','xls')) else pd.read_csv(uploaded_lab)
except Exception as e:
    st.error(f"Erro lendo arquivos: {e}")
    st.stop()

# auto-detect columns
mv_patient_col = find_col(mv, ['paciente','nome'])
mv_exam_col    = find_col(mv, ['exame','proced','procedimento'])
mv_id_col      = find_col(mv, ['atendimento','id_atend','numero_atendimento','num_atend'])
mv_time_col    = detect_datetime_col(mv)

lab_patient_col = find_col(lab, ['paciente','nome'])
lab_exam_col    = find_col(lab, ['exame','proced','procedimento'])
lab_id_col      = find_col(lab, ['atendimento','id_atend','numero_atendimento'])
lab_time_col    = detect_datetime_col(lab)

if mv_patient_col is None or mv_exam_col is None:
    st.error("Não foi possível identificar colunas essenciais na planilha MV (paciente/exame). Renomeie ou verifique o arquivo.")
    st.stop()

# normalize and canonicalize
mv['paciente_norm'] = mv[mv_patient_col].apply(normalize_text)
mv['exame_norm'] = mv[mv_exam_col].apply(lambda x: canonical_exam(x, syn_map))
mv['atendimento_id'] = mv[mv_id_col].astype(str) if mv_id_col else ''
mv['ts_mv'] = pd.to_datetime(mv[mv_time_col], errors='coerce', dayfirst=True) if mv_time_col else pd.NaT

lab['paciente_norm'] = lab[lab_patient_col].apply(normalize_text) if lab_patient_col else ''
lab['exame_norm'] = lab[lab_exam_col].apply(lambda x: canonical_exam(x, syn_map)) if lab_exam_col else ''
lab['atendimento_id'] = lab[lab_id_col].astype(str) if lab_id_col else ''
lab['ts_lab'] = pd.to_datetime(lab[lab_time_col], errors='coerce', dayfirst=True) if lab_time_col else pd.NaT

# optional date filter
if start_date and end_date and start_date > end_date:
    st.error("Data início é maior que data fim.")

if start_date:
    mv = mv[mv['ts_mv'].dt.date >= start_date] if 'ts_mv' in mv else mv
    lab = lab[lab['ts_lab'].dt.date >= start_date] if 'ts_lab' in lab else lab
if end_date:
    mv = mv[mv['ts_mv'].dt.date <= end_date] if 'ts_mv' in mv else mv
    lab = lab[lab['ts_lab'].dt.date <= end_date] if 'ts_lab' in lab else lab

# aggregates
agg_mv = mv.groupby(['atendimento_id','paciente_norm','exame_norm']).size().reset_index(name='qtd_mv')
agg_lab = lab.groupby(['atendimento_id','paciente_norm','exame_norm']).size().reset_index(name='qtd_lab')

# UI - tabs
tab1, tab2, tab3, tab4 = st.tabs(["Dashboard","Comparação","Auditoria Manual","Export & PDF"])

# ---------------------- Dashboard ----------------------
with tab1:
    st.header("📈 Dashboard")
    col1, col2, col3 = st.columns(3)
    total_solicitados = len(mv)
    total_realizados = len(lab)
    not_realizados_abs = total_solicitados - total_realizados
    pct_not_realizados = not_realizados_abs / max(1, total_solicitados) * 100

    # time matching quick heuristic (for KPIs)
    # group lab by first token for speed
    lab_first = {}
    for _, r in lab.iterrows():
        key = r['exame_norm'].split()[0] if isinstance(r['exame_norm'], str) and r['exame_norm'].split() else ''
        lab_first.setdefault(key, []).append(r)

    # compute time deltas for sample
    time_deltas = []
    for _, r in mv.iterrows():
        key = r['exame_norm'].split()[0] if isinstance(r['exame_norm'], str) and r['exame_norm'].split() else ''
        candidates = lab_first.get(key, [])
        candidates = [c for c in candidates if token_set_ratio(c['paciente_norm'], r['paciente_norm']) >= fuzzy_threshold]
        if not candidates:
            continue
        cand_with_ts = [c for c in candidates if pd.notna(c['ts_lab']) and pd.notna(r['ts_mv'])]
        if not cand_with_ts:
            continue
        best = min(cand_with_ts, key=lambda x: abs((x['ts_lab'] - r['ts_mv']).total_seconds()))
        delta = (best['ts_lab'] - r['ts_mv']).total_seconds() / 60.0
        time_deltas.append(delta)

    avg_delta = sum(time_deltas)/len(time_deltas) if time_deltas else None
    atrasos_count = len([d for d in time_deltas if d > tolerance])
    pct_atraso = atrasos_count / max(1, len(time_deltas)) * 100 if time_deltas else 0

    col1.metric("Solicitados", total_solicitados)
    col2.metric("Realizados", total_realizados)
    col3.metric("% Não realizados", f"{pct_not_realizados:.2f}% ({not_realizados_abs})")

    st.markdown("---")
    st.subheader("Distribuição de solicitações por exame (Top 20)")
    top_exams = agg_mv.groupby('exame_norm').agg(solicitados=('qtd_mv','sum')).reset_index().sort_values('solicitados', ascending=False).head(20)
    fig = px.bar(top_exams, x='solicitados', y='exame_norm', orientation='h', height=500, labels={'exame_norm':'Exame','solicitados':'Solicitações'})
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Tempo de resposta — distribuição")
    if time_deltas:
        fig2 = px.histogram(pd.DataFrame({'delta_min': time_deltas}), x='delta_min', nbins=50, title='Distribuição do tempo (min)')
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Sem pares temporais suficientes para plotar distribuição de tempos")

# ---------------------- Comparação ----------------------
with tab2:
    st.header("🔍 Comparação detalhada MV x Laboratório")

    # allow exam filter here too
    exams = sorted(agg_mv['exame_norm'].unique())
    sel_exams = st.multiselect("Filtrar exames (opcional)", exams, default=None, key='comp_exams')
    mv_f = mv.copy()
    lab_f = lab.copy()
    if sel_exams:
        mv_f = mv_f[mv_f['exame_norm'].isin(sel_exams)]
        lab_f = lab_f[lab_f['exame_norm'].isin(sel_exams)]

    # Build fast index of lab aggregates by first token
    lab_by_first = {}
    for _, r in agg_lab.iterrows():
        first = r['exame_norm'].split()[0] if r['exame_norm'] else ''
        lab_by_first.setdefault(first, []).append(r)

    matches = []
    for _, row in agg_mv.iterrows():
        exame = row['exame_norm']
        exact = agg_lab[(agg_lab['exame_norm']==exame) & (agg_lab['paciente_norm']==row['paciente_norm'])]
        if not exact.empty:
            matches.append({'paciente_mv': row['paciente_norm'],'exame_mv': exame, 'qtd_mv': int(row['qtd_mv']), 'qtd_lab': int(exact['qtd_lab'].sum()), 'score':100})
            continue
        first = exame.split()[0] if exame else ''
        candidates = lab_by_first.get(first, [])
        if not candidates:
            candidates = agg_lab.sample(n=min(80, len(agg_lab)), random_state=1).to_dict('records')
        best_score=-1; best_q=0
        for c in candidates:
            sp = token_set_ratio(row['paciente_norm'], c['paciente_norm'])
            se = token_set_ratio(exame, c['exame_norm'])
            score = int(sp*weight_patient + se*weight_exam)
            if score>best_score:
                best_score=score; best_q=int(c['qtd_lab'])
        matches.append({'paciente_mv': row['paciente_norm'],'exame_mv': exame,'qtd_mv': int(row['qtd_mv']),'qtd_lab': best_q,'score':best_score})

    comp_df = pd.DataFrame(matches)
    comp_df['diff'] = comp_df['qtd_mv'] - comp_df['qtd_lab']
    comp_df['status'] = comp_df['diff'].apply(lambda d: 'OK' if d==0 else ('FALTAM' if d>0 else 'EXCESSO'))

    st.subheader("Resumo agregado")
    c1, c2 = st.columns(2)
    c1.write(comp_df['status'].value_counts())
    c2.write(comp_df[['qtd_mv','qtd_lab']].sum())

    st.subheader("Top divergências")
    top_div = comp_df[comp_df['diff']>0].sort_values('diff', ascending=False).head(20)
    if not top_div.empty:
        fig3 = px.bar(top_div, x='diff', y='exame_mv', orientation='h', title='Top divergências (faltam no Lab)')
        st.plotly_chart(fig3, use_container_width=True)
    else:
        st.info("Nenhuma divergência positiva detectada no filtro atual")

    st.subheader("Tabela de comparação (editar/filtrar)")
    st.dataframe(comp_df.sort_values(['status','score'], ascending=[True,False]).head(500))

# ---------------------- Auditoria Manual ----------------------
with tab3:
    st.header("🧪 Auditoria manual — revisar e corrigir pares")
    st.write("Nesta aba você pode revisar correspondências com baixa similaridade e marcar decisões. As alterações serão salvas no log de auditoria.")

    # pick candidates with low score
    review_candidates = comp_df[comp_df['score'] < max(60, fuzzy_threshold)].copy()
    if review_candidates.empty:
        st.info("Não há candidatos com pontuação baixa para revisar — ajuste filtros ou threshold.")
    else:
        st.write(f"{len(review_candidates)} registros para revisão (score baixo)")
        # allow editing via data_editor
        edited = st.data_editor(review_candidates, num_rows="dynamic")
        # decision column: user can change 'qtd_lab' or 'status' to correct
        if st.button("Salvar decisões de auditoria"):
            # append to audit log file
            log_fname = 'audit_log.csv'
            edited['review_ts'] = datetime.now()
            if os.path.exists(log_fname):
                pd.concat([pd.read_csv(log_fname), edited]).to_csv(log_fname, index=False)
            else:
                edited.to_csv(log_fname, index=False)
            st.success(f"Decisões salvas em {log_fname}")

# ---------------------- Export & PDF ----------------------
with tab4:
    st.header("📦 Export & PDF")
    st.write("Gere arquivo Excel consolidado e um PDF resumo para auditoria.")

    if st.button("Gerar Excel consolidado"):
        out_name = 'comparativo_faturamento_avancado.xlsx'
        with pd.ExcelWriter(out_name, engine='openpyxl') as writer:
            mv.to_excel(writer, sheet_name='mv_raw', index=False)
            lab.to_excel(writer, sheet_name='lab_raw', index=False)
            agg_mv.to_excel(writer, sheet_name='agg_mv', index=False)
            agg_lab.to_excel(writer, sheet_name='agg_lab', index=False)
            comp_df.to_excel(writer, sheet_name='comparacao', index=False)
            time_df.to_excel(writer, sheet_name='time_matches', index=False)
            # audit log
            if os.path.exists('audit_log.csv'):
                pd.read_csv('audit_log.csv').to_excel(writer, sheet_name='audit_log', index=False)
        with open(out_name, 'rb') as f:
            st.download_button('Download Excel consolidado', f, file_name=out_name)

    if st.button("Gerar PDF resumo para auditoria"):
        pdf_name = 'relatorio_auditoria.pdf'
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 8, 'Relatório de Auditoria - Comparativo MV x Laboratório', ln=True)
        pdf.ln(4)
        pdf.set_font('Arial', '', 10)
        pdf.cell(0, 6, f'Data geração: {datetime.now()}', ln=True)
        pdf.ln(6)
        # key indicators
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0,6,'Indicadores principais', ln=True)
        pdf.set_font('Arial', '', 10)
        pdf.cell(0,6,f'Solicitados: {total_solicitados}', ln=True)
        pdf.cell(0,6,f'Realizados: {total_realizados}', ln=True)
        pdf.cell(0,6,f'Não realizados (abs): {not_realizados_abs}', ln=True)
        pdf.cell(0,6,f'% Não realizados: {pct_not_realizados:.2f}%', ln=True)
        pdf.ln(6)
        # include top 10 divergences table text
        pdf.set_font('Arial', 'B', 12)
        pdf.cell(0,6,'Top divergências (amostra)', ln=True)
        pdf.set_font('Arial', '', 9)
        top10 = comp_df.sort_values('diff', ascending=False).head(10)
        for _, r in top10.iterrows():
            pdf.multi_cell(0,5,f"{r['exame_mv'][:80]} -- Sol: {r['qtd_mv']} Lab: {r['qtd_lab']} Dif: {r['diff']}")
        pdf.output(pdf_name)
        with open(pdf_name, 'rb') as f:
            st.download_button('Download PDF da auditoria', f, file_name=pdf_name)

st.sidebar.markdown('---')
st.sidebar.write('Versão: Avançada — entregue pelo assistente')

