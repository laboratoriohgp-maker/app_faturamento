import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import os
from datetime import datetime, timedelta
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
from fpdf import FPDF
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(layout="wide", page_title="App Faturamento Pro", page_icon="🏥")

# ---------------------- Utilitários ----------------------

def normalize_text(s):
    """Normaliza texto removendo acentos e padronizando"""
    if pd.isna(s): return ""
    if not isinstance(s, str): s = str(s)
    s = s.strip().upper()
    s = unicodedata.normalize('NFKD', s)
    s = ''.join([c for c in s if not unicodedata.combining(c)])
    return s

def token_set_ratio(a, b):
    """Calcula similaridade entre textos usando tokens"""
    if not a and not b: return 100
    sa = set(a.split())
    sb = set(b.split())
    if not sa or not sb: return 0
    return int(len(sa.intersection(sb)) / len(sa.union(sb)) * 100)

def levenshtein_similarity(s1, s2):
    """Calcula similaridade usando distância de Levenshtein"""
    if s1 == s2: return 100
    if not s1 or not s2: return 0
    len1, len2 = len(s1), len(s2)
    if len1 < len2:
        s1, s2, len1, len2 = s2, s1, len2, len1
    
    current = range(len2 + 1)
    for i in range(1, len1 + 1):
        previous, current = current, [i] + [0] * len2
        for j in range(1, len2 + 1):
            add, delete, change = previous[j] + 1, current[j - 1] + 1, previous[j - 1]
            if s1[i - 1] != s2[j - 1]:
                change += 1
            current[j] = min(add, delete, change)
    
    distance = current[len2]
    max_len = max(len1, len2)
    return int((1 - distance / max_len) * 100)

def canonical_exam(name, synonyms_map=None):
    """Canonicaliza nome do exame usando mapa de sinônimos"""
    if not name: return name
    n = normalize_text(name)
    if synonyms_map:
        if n in synonyms_map:
            return synonyms_map[n]
        for k, v in synonyms_map.items():
            if token_set_ratio(k, n) > 90:
                return v
    return n

def detect_datetime_col(df):
    """Detecta coluna de data/hora automaticamente"""
    for c in df.columns:
        low = c.lower()
        if "data" in low and "hora" in low:
            return c
    for c in df.columns:
        low = c.lower()
        if any(w in low for w in ["data", "hora", "timestamp", "ts", "date", "time"]):
            return c
    return None

def find_col(df, keywords):
    """Encontra coluna por palavras-chave"""
    cols = list(df.columns)
    low = [c.lower() for c in cols]
    for kw in keywords:
        for i, lc in enumerate(low):
            if kw in lc:
                return cols[i]
    return None

def detect_duplicates(df, key_cols):
    """Detecta duplicatas em dataframe"""
    if not key_cols or not all(c in df.columns for c in key_cols):
        return pd.DataFrame()
    dups = df[df.duplicated(subset=key_cols, keep=False)]
    return dups.sort_values(by=key_cols)

def calculate_financial_impact(comp_df, value_col='valor_unit'):
    """Calcula impacto financeiro das divergências"""
    if value_col not in comp_df.columns:
        return None
    
    comp_df['impacto_financeiro'] = comp_df['diff'] * comp_df[value_col]
    summary = {
        'total_faltam': comp_df[comp_df['diff'] > 0]['impacto_financeiro'].sum(),
        'total_excesso': abs(comp_df[comp_df['diff'] < 0]['impacto_financeiro'].sum()),
        'divergencia_liquida': comp_df['impacto_financeiro'].sum()
    }
    return summary

# ---------------------- Sidebar / Configuração ----------------------
st.sidebar.title("⚙️ Configuração")

uploaded_synonyms = st.sidebar.file_uploader(
    "📋 (Opcional) Upload synonyms.csv", 
    type=["csv"],
    help="Arquivo CSV com colunas: synonym,canonical"
)

st.sidebar.markdown("### 🎯 Parâmetros de Matching")
weight_patient = st.sidebar.slider("Peso nome paciente", 0.0, 1.0, 0.5, 0.05)
weight_exam = st.sidebar.slider("Peso exame", 0.0, 1.0, 0.3, 0.05)
weight_id = st.sidebar.slider("Peso ID atendimento", 0.0, 1.0, 0.2, 0.05)

fuzzy_threshold = st.sidebar.slider("Threshold matching (0-100)", 30, 100, 65, 5)
tolerance = st.sidebar.number_input("Tolerância tempo (min)", min_value=0, value=120, step=10)

st.sidebar.markdown("### 📅 Filtros de Data")
start_date = st.sidebar.date_input("Data início", value=None)
end_date = st.sidebar.date_input("Data fim", value=None)

st.sidebar.markdown("### 🔍 Opções Avançadas")
detect_dups = st.sidebar.checkbox("Detectar duplicatas", value=True)
show_low_score = st.sidebar.checkbox("Alertar baixa similaridade", value=True)
use_multi_algorithm = st.sidebar.checkbox("Usar múltiplos algoritmos de matching", value=True)

# ---------------------- Main UI ----------------------
st.title("🏥 Sistema de Análise de Faturamento - MV x Laboratório")
st.markdown("""
**Sistema profissional de análise e reconciliação de faturamento**  
Carregue as planilhas MV (solicitações) e Laboratório (realizações) para começar a análise.
""")

# Uploads
col1, col2 = st.columns(2)
with col1:
    uploaded_mv = st.file_uploader(
        "📊 Planilha MV (Solicitações)", 
        type=["xlsx", "xls", "csv"], 
        key='mv',
        help="Arquivo com as solicitações do sistema MV"
    )
with col2:
    uploaded_lab = st.file_uploader(
        "🧪 Planilha Laboratório (Realizações)", 
        type=["xlsx", "xls", "csv"], 
        key='lab',
        help="Arquivo com os exames realizados pelo laboratório"
    )

# Load synonyms
syn_map = None
if uploaded_synonyms is not None:
    try:
        df_syn = pd.read_csv(uploaded_synonyms)
        if 'synonym' in df_syn.columns and 'canonical' in df_syn.columns:
            syn_map = {
                normalize_text(r['synonym']): normalize_text(r['canonical']) 
                for _, r in df_syn.iterrows()
            }
            st.sidebar.success(f"✅ {len(syn_map)} sinônimos carregados")
        else:
            st.sidebar.error("❌ Arquivo deve ter colunas 'synonym' e 'canonical'")
    except Exception as e:
        st.sidebar.error(f"❌ Erro lendo sinônimos: {e}")

# Require uploads
if not (uploaded_mv and uploaded_lab):
    st.info("👆 Carregue as duas planilhas nas abas acima para começar a análise.")
    st.stop()

# ---------------------- Read Files ----------------------
try:
    # Read MV
    if uploaded_mv.name.lower().endswith(('.xlsx', '.xls')):
        mv = pd.read_excel(uploaded_mv)
    else:
        mv = pd.read_csv(uploaded_mv)
    
    # Read Lab
    if uploaded_lab.name.lower().endswith(('.xlsx', '.xls')):
        lab = pd.read_excel(uploaded_lab)
    else:
        lab = pd.read_csv(uploaded_lab)
    
    st.success(f"✅ Arquivos carregados: MV ({len(mv)} linhas), Lab ({len(lab)} linhas)")
    
except Exception as e:
    st.error(f"❌ Erro lendo arquivos: {e}")
    st.stop()

# ---------------------- Column Mapping ----------------------
st.info("🔍 Identificando colunas nas planilhas...")

# Mapeamento de colunas MV (baseado na imagem fornecida)
mv_col_mapping = {
    'patient': 'NOME DO PACIENTE',
    'exam': 'EXAME COLETADO',
    'id': 'ATEND',
    'datetime_solicit': 'DATA PEDIDO',
    'datetime_coleta': 'DATA COLETA',
    'solicitante': 'SOLICITANTE',
    'local': 'LOCAL'
}

# Mapeamento de colunas Laboratório (baseado na imagem fornecida)
lab_col_mapping = {
    'patient': 'Nome',
    'exam': 'Cadastro',
    'id': 'Cód. Doente',
    'material': 'Material',
    'valor': 'Valor',
    'datetime': 'Tempo Cadastro/Itagem'
}

# Verificar se as colunas esperadas existem, senão tentar auto-detectar
def get_column(df, preferred_name, fallback_keywords):
    """Tenta pegar coluna preferencial, senão usa fallback"""
    if preferred_name in df.columns:
        return preferred_name
    return find_col(df, fallback_keywords)

# Colunas MV
mv_patient_col = get_column(mv, mv_col_mapping['patient'], ['paciente', 'nome', 'patient'])
mv_exam_col = get_column(mv, mv_col_mapping['exam'], ['exame', 'proced', 'coletado', 'exam'])
mv_id_col = get_column(mv, mv_col_mapping['id'], ['atend', 'atendimento', 'id'])
mv_time_col = get_column(mv, mv_col_mapping.get('datetime_solicit'), ['data', 'pedido', 'solicit'])
mv_coleta_col = get_column(mv, mv_col_mapping.get('datetime_coleta'), ['coleta', 'data coleta'])
mv_solicitante_col = get_column(mv, mv_col_mapping.get('solicitante'), ['solicitante', 'medico', 'doctor'])
mv_local_col = get_column(mv, mv_col_mapping.get('local'), ['local', 'setor', 'unidade'])
mv_value_col = None  # MV não parece ter valor
mv_plan_col = None  # MV não parece ter convênio

# Colunas Lab
lab_patient_col = get_column(lab, lab_col_mapping['patient'], ['nome', 'paciente', 'patient'])
lab_exam_col = get_column(lab, lab_col_mapping['exam'], ['cadastro', 'exame', 'exam'])
lab_id_col = get_column(lab, lab_col_mapping['id'], ['cod', 'doente', 'codigo', 'id'])
lab_time_col = get_column(lab, lab_col_mapping['datetime'], ['tempo', 'cadastro', 'itagem', 'data'])
lab_material_col = get_column(lab, lab_col_mapping.get('material'), ['material', 'amostra'])
lab_value_col = get_column(lab, lab_col_mapping.get('valor'), ['valor', 'preco'])

# Validação e feedback
col_check1, col_check2 = st.columns(2)

with col_check1:
    st.markdown("**📋 Colunas MV Identificadas:**")
    mv_status = {
        'Paciente': '✅' if mv_patient_col else '❌',
        'Exame': '✅' if mv_exam_col else '❌',
        'ID Atendimento': '✅' if mv_id_col else '⚠️',
        'Data/Hora': '✅' if mv_time_col else '⚠️',
        'Solicitante': '✅' if mv_solicitante_col else '⚠️',
        'Local': '✅' if mv_local_col else '⚠️'
    }
    for key, status in mv_status.items():
        st.text(f"{status} {key}")

with col_check2:
    st.markdown("**🧪 Colunas LAB Identificadas:**")
    lab_status = {
        'Paciente': '✅' if lab_patient_col else '❌',
        'Exame': '✅' if lab_exam_col else '❌',
        'Cód. Doente': '✅' if lab_id_col else '⚠️',
        'Data/Hora': '✅' if lab_time_col else '⚠️',
        'Material': '✅' if lab_material_col else '⚠️',
        'Valor': '✅' if lab_value_col else '⚠️'
    }
    for key, status in lab_status.items():
        st.text(f"{status} {key}")

# Validação crítica
if mv_patient_col is None or mv_exam_col is None:
    st.error("❌ Colunas essenciais não encontradas na planilha MV")
    with st.expander("🔍 Ver colunas disponíveis no MV"):
        st.write(list(mv.columns))
    st.stop()

if lab_patient_col is None or lab_exam_col is None:
    st.error("❌ Colunas essenciais não encontradas na planilha LAB")
    with st.expander("🔍 Ver colunas disponíveis no LAB"):
        st.write(list(lab.columns))
    st.stop()

st.success("✅ Todas as colunas essenciais foram identificadas!")

# ---------------------- Data Processing ----------------------

# Normalize and canonicalize MV
mv['paciente_norm'] = mv[mv_patient_col].apply(normalize_text)
mv['exame_norm'] = mv[mv_exam_col].apply(lambda x: canonical_exam(x, syn_map))
mv['atendimento_id'] = mv[mv_id_col].astype(str).str.strip() if mv_id_col else ''

# Parse datetime - MV pode ter data e hora em colunas separadas
if mv_time_col:
    # Se tem "DATA PEDIDO", tentar converter
    mv['ts_mv'] = pd.to_datetime(mv[mv_time_col], errors='coerce', dayfirst=True)
elif mv_coleta_col:
    # Se não tem pedido, usar coleta
    mv['ts_mv'] = pd.to_datetime(mv[mv_coleta_col], errors='coerce', dayfirst=True)
else:
    mv['ts_mv'] = pd.NaT

# Adicionar colunas extras do MV se existirem
if mv_solicitante_col:
    mv['solicitante'] = mv[mv_solicitante_col].astype(str)
if mv_local_col:
    mv['local'] = mv[mv_local_col].astype(str)

# Normalize and canonicalize Lab
lab['paciente_norm'] = lab[lab_patient_col].apply(normalize_text)
lab['exame_norm'] = lab[lab_exam_col].apply(lambda x: canonical_exam(x, syn_map))
lab['atendimento_id'] = lab[lab_id_col].astype(str).str.strip() if lab_id_col else ''

# Parse datetime Lab - formato pode ser "dd/mm/aaaa hh:mm:ss"
if lab_time_col:
    lab['ts_lab'] = pd.to_datetime(lab[lab_time_col], errors='coerce', dayfirst=True)
else:
    lab['ts_lab'] = pd.NaT

# Parse valores se existirem
if lab_value_col:
    # Limpar e converter valores (pode vir como "R$ 123,45")
    lab['valor_lab'] = lab[lab_value_col].astype(str).str.replace('R')

# Date filtering
if start_date and end_date and start_date > end_date:
    st.error("❌ Data início é maior que data fim")
    st.stop()

mv_original_len = len(mv)
lab_original_len = len(lab)

if start_date:
    if 'ts_mv' in mv.columns:
        mv = mv[mv['ts_mv'].dt.date >= start_date]
    if 'ts_lab' in lab.columns:
        lab = lab[lab['ts_lab'].dt.date >= start_date]

if end_date:
    if 'ts_mv' in mv.columns:
        mv = mv[mv['ts_mv'].dt.date <= end_date]
    if 'ts_lab' in lab.columns:
        lab = lab[lab['ts_lab'].dt.date <= end_date]

if start_date or end_date:
    st.info(f"📅 Filtro aplicado: MV {mv_original_len}→{len(mv)}, Lab {lab_original_len}→{len(lab)}")

# Detect duplicates
duplicates_mv = pd.DataFrame()
duplicates_lab = pd.DataFrame()

if detect_dups:
    dup_cols = ['paciente_norm', 'exame_norm', 'atendimento_id']
    duplicates_mv = detect_duplicates(mv, dup_cols)
    duplicates_lab = detect_duplicates(lab, dup_cols)

# Aggregates - verificar se colunas existem antes de agregar
agg_mv_dict = {'ts_mv': 'first'} if 'ts_mv' in mv.columns else {}
if mv_solicitante_col and 'solicitante' in mv.columns:
    agg_mv_dict['solicitante'] = 'first'
if mv_local_col and 'local' in mv.columns:
    agg_mv_dict['local'] = 'first'

if agg_mv_dict:
    agg_mv = mv.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).agg(agg_mv_dict).reset_index()
else:
    agg_mv = mv.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).size().reset_index(name='_temp')
    agg_mv = agg_mv.drop('_temp', axis=1)

agg_mv['qtd_mv'] = mv.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).size().values

# Aggregates Lab
agg_lab_dict = {'ts_lab': 'first'} if 'ts_lab' in lab.columns else {}
if lab_value_col and 'valor_lab' in lab.columns:
    agg_lab_dict['valor_lab'] = 'sum'
if lab_material_col and 'material' in lab.columns:
    agg_lab_dict['material'] = 'first'

if agg_lab_dict:
    agg_lab = lab.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).agg(agg_lab_dict).reset_index()
else:
    agg_lab = lab.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).size().reset_index(name='_temp')
    agg_lab = agg_lab.drop('_temp', axis=1)

agg_lab['qtd_lab'] = lab.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).size().values

# ---------------------- Matching Logic ----------------------

def advanced_match(row_mv, agg_lab_data, use_multi=True):
    """Matching avançado com múltiplos algoritmos"""
    best_match = None
    best_score = -1
    
    for _, row_lab in agg_lab_data.iterrows():
        # Score de ID
        score_id = 100 if row_mv['atendimento_id'] and row_mv['atendimento_id'] == row_lab['atendimento_id'] else 0
        
        # Score de paciente
        score_patient = token_set_ratio(row_mv['paciente_norm'], row_lab['paciente_norm'])
        if use_multi:
            score_patient_lev = levenshtein_similarity(row_mv['paciente_norm'], row_lab['paciente_norm'])
            score_patient = (score_patient + score_patient_lev) / 2
        
        # Score de exame
        score_exam = token_set_ratio(row_mv['exame_norm'], row_lab['exame_norm'])
        if use_multi:
            score_exam_lev = levenshtein_similarity(row_mv['exame_norm'], row_lab['exame_norm'])
            score_exam = (score_exam + score_exam_lev) / 2
        
        # Score final ponderado
        total_score = (
            score_patient * weight_patient +
            score_exam * weight_exam +
            score_id * weight_id
        )
        
        if total_score > best_score:
            best_score = total_score
            best_match = row_lab
    
    return best_match, best_score

# Build matches
st.info("🔄 Processando matching entre planilhas...")
matches = []
time_matches = []

for idx, row in agg_mv.iterrows():
    # Try exact match first
    exact = agg_lab[
        (agg_lab['exame_norm'] == row['exame_norm']) & 
        (agg_lab['paciente_norm'] == row['paciente_norm'])
    ]
    
    if not exact.empty:
        match_data = {
            'paciente_mv': row['paciente_norm'],
            'exame_mv': row['exame_norm'],
            'atendimento_mv': row['atendimento_id'],
            'qtd_mv': int(row['qtd_mv']),
            'qtd_lab': int(exact['qtd_lab'].sum()),
            'score': 100,
            'match_type': 'EXACT'
        }
        # Não adicionar valor_mv pois MV não tem essa coluna
        if lab_value_col and 'valor_lab' in exact.columns:
            match_data['valor_lab'] = exact['valor_lab'].sum()
        
        matches.append(match_data)
        
        # Time matching
        if pd.notna(row['ts_mv']) and pd.notna(exact.iloc[0]['ts_lab']):
            delta_min = (exact.iloc[0]['ts_lab'] - row['ts_mv']).total_seconds() / 60
            time_matches.append({
                'paciente': row['paciente_norm'],
                'exame': row['exame_norm'],
                'ts_mv': row['ts_mv'],
                'ts_lab': exact.iloc[0]['ts_lab'],
                'delta_min': delta_min,
                'status': 'DENTRO_SLA' if delta_min <= tolerance else 'ATRASO'
            })
    else:
        # Advanced fuzzy match
        best_match, best_score = advanced_match(row, agg_lab, use_multi_algorithm)
        
        match_data = {
            'paciente_mv': row['paciente_norm'],
            'exame_mv': row['exame_norm'],
            'atendimento_mv': row['atendimento_id'],
            'qtd_mv': int(row['qtd_mv']),
            'qtd_lab': int(best_match['qtd_lab']) if best_match is not None else 0,
            'score': int(best_score),
            'match_type': 'FUZZY' if best_score >= fuzzy_threshold else 'NO_MATCH'
        }
        
        # Não adicionar valor_mv
        if best_match is not None and lab_value_col and 'valor_lab' in best_match:
            match_data['valor_lab'] = best_match['valor_lab']
        
        matches.append(match_data)

comp_df = pd.DataFrame(matches)
comp_df['diff'] = comp_df['qtd_mv'] - comp_df['qtd_lab']
comp_df['status'] = comp_df.apply(
    lambda r: 'OK' if r['diff'] == 0 else ('FALTAM' if r['diff'] > 0 else 'EXCESSO'),
    axis=1
)

# Calculate financial impact if values exist (só Lab tem valores)
financial_summary = None
if 'valor_lab' in comp_df.columns:
    # Calcular o valor médio por exame para estimar MV
    valor_medio_exame = lab.groupby('exame_norm')['valor_lab'].mean().to_dict()
    
    # Estimar valor MV baseado nos valores do Lab
    comp_df['valor_mv_estimado'] = comp_df['exame_mv'].map(valor_medio_exame).fillna(0) * comp_df['qtd_mv']
    comp_df['valor_lab'] = comp_df['valor_lab'].fillna(0)
    comp_df['diff_valor'] = comp_df['valor_mv_estimado'] - comp_df['valor_lab']
    
    financial_summary = {
        'total_mv': comp_df['valor_mv_estimado'].sum(),
        'total_lab': comp_df['valor_lab'].sum(),
        'divergencia': comp_df['diff_valor'].sum(),
        'faltam': comp_df[comp_df['diff_valor'] > 0]['diff_valor'].sum(),
        'excesso': abs(comp_df[comp_df['diff_valor'] < 0]['diff_valor'].sum())
    }

time_df = pd.DataFrame(time_matches) if time_matches else pd.DataFrame()

# ---------------------- Tabs ----------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Dashboard", 
    "🔍 Comparação", 
    "⏱️ Análise Temporal", 
    "🧪 Auditoria",
    "💰 Financeiro",
    "📦 Export"
])

# ---------------------- TAB 1: Dashboard ----------------------
with tab1:
    st.header("📊 Dashboard Executivo")
    
    # KPIs principais
    col1, col2, col3, col4 = st.columns(4)
    
    total_solicitados = len(mv)
    total_realizados = len(lab)
    taxa_realizacao = (total_realizados / total_solicitados * 100) if total_solicitados > 0 else 0
    
    col1.metric("📋 Solicitados", f"{total_solicitados:,}", delta=None)
    col2.metric("✅ Realizados", f"{total_realizados:,}", delta=f"{taxa_realizacao:.1f}%")
    
    divergencias = len(comp_df[comp_df['status'] != 'OK'])
    col3.metric("⚠️ Divergências", f"{divergencias:,}", delta=f"{divergencias/len(comp_df)*100:.1f}%")
    
    if not time_df.empty:
        atrasos = len(time_df[time_df['status'] == 'ATRASO'])
        col4.metric("🕐 Atrasos", f"{atrasos:,}", delta=f"{atrasos/len(time_df)*100:.1f}%")
    else:
        col4.metric("🕐 Atrasos", "N/A")
    
    st.markdown("---")
    
    # Duplicatas
    if detect_dups and (not duplicates_mv.empty or not duplicates_lab.empty):
        st.warning("⚠️ **Duplicatas Detectadas!**")
        coldup1, coldup2 = st.columns(2)
        coldup1.metric("Duplicatas MV", len(duplicates_mv))
        coldup2.metric("Duplicatas LAB", len(duplicates_lab))
    
    # Gráficos
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("📈 Distribuição por Status")
        status_counts = comp_df['status'].value_counts()
        fig_status = px.pie(
            values=status_counts.values, 
            names=status_counts.index,
            title="Status das Solicitações",
            color_discrete_map={'OK':'green', 'FALTAM':'red', 'EXCESSO':'orange'}
        )
        st.plotly_chart(fig_status, use_container_width=True)
    
    with col_right:
        st.subheader("🎯 Qualidade do Matching")
        match_quality = comp_df['match_type'].value_counts()
        fig_match = px.bar(
            x=match_quality.index, 
            y=match_quality.values,
            title="Tipo de Correspondência",
            labels={'x':'Tipo', 'y':'Quantidade'},
            color=match_quality.index
        )
        st.plotly_chart(fig_match, use_container_width=True)
    
    # Top exames
    st.subheader("🔬 Top 15 Exames Solicitados")
    top_exams = comp_df.groupby('exame_mv').agg({
        'qtd_mv': 'sum',
        'qtd_lab': 'sum',
        'diff': 'sum'
    }).reset_index().sort_values('qtd_mv', ascending=False).head(15)
    
    fig_top = go.Figure()
    fig_top.add_trace(go.Bar(name='Solicitados', x=top_exams['exame_mv'], y=top_exams['qtd_mv']))
    fig_top.add_trace(go.Bar(name='Realizados', x=top_exams['exame_mv'], y=top_exams['qtd_lab']))
    fig_top.update_layout(barmode='group', height=400)
    st.plotly_chart(fig_top, use_container_width=True)
    
    # Alertas
    if show_low_score:
        low_scores = comp_df[comp_df['score'] < fuzzy_threshold]
        if not low_scores.empty:
            st.warning(f"⚠️ **{len(low_scores)} registros com baixa similaridade** (< {fuzzy_threshold}%)")
            with st.expander("Ver registros problemáticos"):
                st.dataframe(low_scores[['paciente_mv', 'exame_mv', 'score', 'status']].head(20))

# ---------------------- TAB 2: Comparação ----------------------
with tab2:
    st.header("🔍 Comparação Detalhada MV x Laboratório")
    
    # Filters
    col_f1, col_f2, col_f3 = st.columns(3)
    
    with col_f1:
        status_filter = st.multiselect(
            "Filtrar por Status",
            options=comp_df['status'].unique(),
            default=None
        )
    
    with col_f2:
        match_filter = st.multiselect(
            "Filtrar por Tipo de Match",
            options=comp_df['match_type'].unique(),
            default=None
        )
    
    with col_f3:
        score_min = st.slider("Score mínimo", 0, 100, 0)
    
    # Apply filters
    filtered_comp = comp_df.copy()
    if status_filter:
        filtered_comp = filtered_comp[filtered_comp['status'].isin(status_filter)]
    if match_filter:
        filtered_comp = filtered_comp[filtered_comp['match_type'].isin(match_filter)]
    if score_min > 0:
        filtered_comp = filtered_comp[filtered_comp['score'] >= score_min]
    
    st.info(f"📊 Mostrando {len(filtered_comp)} de {len(comp_df)} registros")
    
    # Summary
    col_s1, col_s2, col_s3 = st.columns(3)
    col_s1.metric("OK", len(filtered_comp[filtered_comp['status'] == 'OK']))
    col_s2.metric("Faltam", len(filtered_comp[filtered_comp['status'] == 'FALTAM']))
    col_s3.metric("Excesso", len(filtered_comp[filtered_comp['status'] == 'EXCESSO']))
    
    # Top divergências
    st.subheader("🔴 Maiores Divergências")
    top_div = filtered_comp[filtered_comp['diff'] != 0].sort_values('diff', key=abs, ascending=False).head(20)
    
    if not top_div.empty:
        fig_div = px.bar(
            top_div,
            x='diff',
            y='exame_mv',
            orientation='h',
            title='Top 20 Divergências',
            color='status',
            color_discrete_map={'FALTAM':'red', 'EXCESSO':'orange'}
        )
        st.plotly_chart(fig_div, use_container_width=True)
    else:
        st.success("✅ Nenhuma divergência encontrada!")
    
    # Data table
    st.subheader("📋 Tabela Completa")
    st.dataframe(
        filtered_comp.sort_values(['status', 'score'], ascending=[True, False]),
        use_container_width=True,
        height=400
    )

# ---------------------- TAB 3: Análise Temporal ----------------------
with tab3:
    st.header("⏱️ Análise de Tempo e SLA")
    
    if time_df.empty:
        st.warning("⚠️ Sem dados temporais suficientes para análise")
    else:
        # KPIs temporais
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        
        avg_delta = time_df['delta_min'].mean()
        median_delta = time_df['delta_min'].median()
        atrasos = len(time_df[time_df['status'] == 'ATRASO'])
        taxa_sla = (1 - atrasos / len(time_df)) * 100
        
        col_t1.metric("⏱️ Tempo Médio", f"{avg_delta:.1f} min")
        col_t2.metric("📊 Mediana", f"{median_delta:.1f} min")
        col_t3.metric("⚠️ Atrasos", f"{atrasos}", delta=f"-{atrasos/len(time_df)*100:.1f}%")
        col_t4.metric("✅ Taxa SLA", f"{taxa_sla:.1f}%")
        
        # Distribuição temporal
        st.subheader("📈 Distribuição de Tempo de Resposta")
        fig_time_hist = px.histogram(
            time_df,
            x='delta_min',
            nbins=50,
            title='Distribuição do Tempo (minutos)',
            labels={'delta_min': 'Tempo (min)', 'count': 'Frequência'},
            color_discrete_sequence=['#636EFA']
        )
        fig_time_hist.add_vline(x=tolerance, line_dash="dash", line_color="red", 
                                annotation_text=f"SLA: {tolerance} min")
        st.plotly_chart(fig_time_hist, use_container_width=True)
        
        # Evolução temporal
        if 'ts_mv' in time_df.columns:
            time_df['data'] = pd.to_datetime(time_df['ts_mv']).dt.date
            daily_avg = time_df.groupby('data')['delta_min'].agg(['mean', 'count']).reset_index()
            
            fig_evolution = go.Figure()
            fig_evolution.add_trace(go.Scatter(
                x=daily_avg['data'],
                y=daily_avg['mean'],
                mode='lines+markers',
                name='Tempo Médio',
                line=dict(color='blue')
            ))
            fig_evolution.add_hline(y=tolerance, line_dash="dash", line_color="red",
                                   annotation_text="SLA")
            fig_evolution.update_layout(
                title='Evolução do Tempo Médio de Resposta',
                xaxis_title='Data',
                yaxis_title='Tempo Médio (min)',
                height=400
            )
            st.plotly_chart(fig_evolution, use_container_width=True)
        
        # Detalhamento de atrasos
        st.subheader("🔴 Exames com Maior Atraso")
        atrasos_df = time_df[time_df['status'] == 'ATRASO'].sort_values('delta_min', ascending=False).head(20)
        
        if not atrasos_df.empty:
            st.dataframe(
                atrasos_df[['paciente', 'exame', 'delta_min', 'ts_mv', 'ts_lab']],
                use_container_width=True
            )
        else:
            st.success("✅ Nenhum atraso identificado!")

# ---------------------- TAB 4: Auditoria ----------------------
with tab4:
    st.header("🧪 Auditoria e Revisão Manual")
    
    st.markdown("""
    **Objetivo:** Revisar registros com baixa similaridade e fazer correções manuais.
    """)
    
    # Filtros para auditoria
    audit_threshold = st.slider("Mostrar registros com score abaixo de:", 0, 100, fuzzy_threshold)
    
    review_candidates = comp_df[comp_df['score'] < audit_threshold].copy()
    
    if review_candidates.empty:
        st.success("✅ Nenhum registro requer revisão manual com os critérios atuais!")
    else:
        st.warning(f"⚠️ **{len(review_candidates)} registros** requerem revisão")
        
        # Estatísticas
        col_a1, col_a2, col_a3 = st.columns(3)
        col_a1.metric("Baixa Similaridade", len(review_candidates))
        col_a2.metric("Score Médio", f"{review_candidates['score'].mean():.1f}")
        col_a3.metric("Impacto", f"{review_candidates['diff'].sum()}")
        
        st.subheader("📝 Registros para Revisão")
        
        # Editor interativo
        edited_df = st.data_editor(
            review_candidates[['paciente_mv', 'exame_mv', 'atendimento_mv', 
                              'qtd_mv', 'qtd_lab', 'score', 'status', 'match_type']],
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            column_config={
                "score": st.column_config.NumberColumn("Score", format="%.0f"),
                "qtd_mv": st.column_config.NumberColumn("Qtd MV"),
                "qtd_lab": st.column_config.NumberColumn("Qtd Lab (editável)", disabled=False),
            }
        )
        
        # Salvar decisões
        col_save1, col_save2 = st.columns([1, 3])
        
        with col_save1:
            if st.button("💾 Salvar Decisões", type="primary"):
                log_fname = 'audit_log.csv'
                edited_df['review_ts'] = datetime.now()
                edited_df['reviewer'] = 'user'
                
                if os.path.exists(log_fname):
                    existing = pd.read_csv(log_fname)
                    combined = pd.concat([existing, edited_df], ignore_index=True)
                    combined.to_csv(log_fname, index=False)
                else:
                    edited_df.to_csv(log_fname, index=False)
                
                st.success(f"✅ Decisões salvas em {log_fname}")
        
        with col_save2:
            if os.path.exists('audit_log.csv'):
                audit_log = pd.read_csv('audit_log.csv')
                st.info(f"📋 Log de auditoria contém {len(audit_log)} revisões")
        
        # Sugestões automáticas
        st.subheader("💡 Sugestões de Reconciliação")
        
        suggestions = []
        for _, row in review_candidates.head(10).iterrows():
            if row['diff'] > 0:
                suggestions.append({
                    'Paciente': row['paciente_mv'][:30],
                    'Exame': row['exame_mv'][:40],
                    'Ação Sugerida': f"Investigar {row['diff']} solicitações não realizadas",
                    'Prioridade': 'ALTA' if row['diff'] > 5 else 'MÉDIA'
                })
            elif row['diff'] < 0:
                suggestions.append({
                    'Paciente': row['paciente_mv'][:30],
                    'Exame': row['exame_mv'][:40],
                    'Ação Sugerida': f"Verificar {abs(row['diff'])} exames em excesso",
                    'Prioridade': 'MÉDIA'
                })
        
        if suggestions:
            st.dataframe(pd.DataFrame(suggestions), use_container_width=True)

# ---------------------- TAB 5: Financeiro ----------------------
with tab5:
    st.header("💰 Análise Financeira")
    
    if financial_summary is None:
        st.warning("⚠️ Dados de valores não encontrados nas planilhas")
        st.info("Para análise financeira, certifique-se de que há colunas de valores nas planilhas")
    else:
        # KPIs financeiros
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        col_f1.metric("💵 Total MV", f"R$ {financial_summary['total_mv']:,.2f}")
        col_f2.metric("💵 Total Lab", f"R$ {financial_summary['total_lab']:,.2f}")
        col_f3.metric("⚠️ Divergência", f"R$ {financial_summary['divergencia']:,.2f}")
        col_f4.metric("📊 % Divergência", 
                     f"{abs(financial_summary['divergencia'])/max(financial_summary['total_mv'],1)*100:.2f}%")
        
        st.info("ℹ️ **Nota:** Valores MV foram estimados com base na média dos valores do Laboratório por tipo de exame, já que a planilha MV não contém valores.")
        
        st.markdown("---")
        
        # Breakdown financeiro
        col_left, col_right = st.columns(2)
        
        with col_left:
            st.subheader("📉 Análise de Divergências")
            div_data = pd.DataFrame({
                'Categoria': ['A Faturar (Faltam)', 'Faturado em Excesso'],
                'Valor': [financial_summary['faltam'], financial_summary['excesso']]
            })
            fig_fin = px.bar(
                div_data,
                x='Categoria',
                y='Valor',
                title='Impacto Financeiro das Divergências',
                color='Categoria',
                color_discrete_map={'A Faturar (Faltam)': 'red', 'Faturado em Excesso': 'orange'}
            )
            st.plotly_chart(fig_fin, use_container_width=True)
        
        with col_right:
            st.subheader("🎯 Reconciliação")
            reconciliation = pd.DataFrame({
                'Item': ['Total Solicitado (MV)', 'Total Realizado (Lab)', 'Diferença'],
                'Valor': [
                    financial_summary['total_mv'],
                    financial_summary['total_lab'],
                    financial_summary['divergencia']
                ]
            })
            st.dataframe(reconciliation.style.format({'Valor': 'R$ {:,.2f}'}), 
                        use_container_width=True)
        
        # Top divergências financeiras
        st.subheader("💸 Maiores Impactos Financeiros")
        
        if 'diff_valor' in comp_df.columns:
            top_fin = comp_df[comp_df['diff_valor'] != 0].sort_values(
                'diff_valor', 
                key=abs, 
                ascending=False
            ).head(15)
            
            if not top_fin.empty:
                fig_top_fin = px.bar(
                    top_fin,
                    x='diff_valor',
                    y='exame_mv',
                    orientation='h',
                    title='Top 15 Divergências Financeiras',
                    labels={'diff_valor': 'Diferença (R$)', 'exame_mv': 'Exame'},
                    color='diff_valor',
                    color_continuous_scale=['red', 'yellow', 'green']
                )
                st.plotly_chart(fig_top_fin, use_container_width=True)
                
                # Tabela detalhada
                st.dataframe(
                    top_fin[['exame_mv', 'qtd_mv', 'qtd_lab', 'valor_mv_estimado', 'valor_lab', 'diff_valor']]
                    .rename(columns={'valor_mv_estimado': 'valor_mv_est'})
                    .style.format({
                        'valor_mv_est': 'R$ {:,.2f}',
                        'valor_lab': 'R$ {:,.2f}',
                        'diff_valor': 'R$ {:,.2f}'
                    }),
                    use_container_width=True
                )
        
        # Análise por convênio (se disponível) - Removido pois MV não tem
        # if mv_plan_col and 'convenio_norm' in mv.columns:
            st.subheader("🏥 Análise por Convênio/Plano")
            
            convenio_analysis = mv.groupby('convenio_norm').agg({
                'valor_mv': 'sum' if 'valor_mv' in mv.columns else 'count'
            }).reset_index().sort_values('valor_mv', ascending=False).head(10)
            
            fig_conv = px.pie(
                convenio_analysis,
                values='valor_mv',
                names='convenio_norm',
                title='Distribuição por Convênio'
            )
            st.plotly_chart(fig_conv, use_container_width=True)

# ---------------------- TAB 6: Export ----------------------
with tab6:
    st.header("📦 Exportação de Dados")
    
    st.markdown("""
    Gere relatórios consolidados em Excel e PDF para auditoria e apresentação.
    """)
    
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        st.subheader("📊 Excel Consolidado")
        
        if st.button("🔄 Gerar Excel Completo", type="primary"):
            with st.spinner("Gerando arquivo Excel..."):
                out_name = f'relatorio_faturamento_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
                
                try:
                    with pd.ExcelWriter(out_name, engine='openpyxl') as writer:
                        # Dados brutos
                        mv.to_excel(writer, sheet_name='MV_Raw', index=False)
                        lab.to_excel(writer, sheet_name='Lab_Raw', index=False)
                        
                        # Agregados
                        agg_mv.to_excel(writer, sheet_name='MV_Agregado', index=False)
                        agg_lab.to_excel(writer, sheet_name='Lab_Agregado', index=False)
                        
                        # Comparação
                        comp_df.to_excel(writer, sheet_name='Comparacao', index=False)
                        
                        # Dados temporais
                        if not time_df.empty:
                            time_df.to_excel(writer, sheet_name='Analise_Temporal', index=False)
                        
                        # Divergências
                        divergencias_df = comp_df[comp_df['status'] != 'OK']
                        divergencias_df.to_excel(writer, sheet_name='Divergencias', index=False)
                        
                        # Duplicatas
                        if not duplicates_mv.empty:
                            duplicates_mv.to_excel(writer, sheet_name='Duplicatas_MV', index=False)
                        if not duplicates_lab.empty:
                            duplicates_lab.to_excel(writer, sheet_name='Duplicatas_Lab', index=False)
                        
                        # Log de auditoria
                        if os.path.exists('audit_log.csv'):
                            pd.read_csv('audit_log.csv').to_excel(writer, sheet_name='Audit_Log', index=False)
                        
                        # Resumo executivo
                        summary_data = {
                            'Métrica': [
                                'Total Solicitados',
                                'Total Realizados',
                                'Taxa de Realização (%)',
                                'Divergências',
                                'Score Médio',
                                'Duplicatas MV',
                                'Duplicatas Lab'
                            ],
                            'Valor': [
                                total_solicitados,
                                total_realizados,
                                taxa_realizacao,
                                divergencias,
                                comp_df['score'].mean(),
                                len(duplicates_mv),
                                len(duplicates_lab)
                            ]
                        }
                        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Resumo_Executivo', index=False)
                    
                    with open(out_name, 'rb') as f:
                        st.download_button(
                            '⬇️ Download Excel',
                            f,
                            file_name=out_name,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    
                    st.success(f"✅ Arquivo gerado: {out_name}")
                    
                except Exception as e:
                    st.error(f"❌ Erro ao gerar Excel: {e}")
    
    with col_exp2:
        st.subheader("📄 Relatório PDF")
        
        if st.button("🔄 Gerar PDF de Auditoria", type="primary"):
            with st.spinner("Gerando PDF..."):
                try:
                    pdf_name = f'auditoria_faturamento_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
                    pdf = FPDF()
                    pdf.set_auto_page_break(auto=True, margin=15)
                    pdf.add_page()
                    
                    # Título
                    pdf.set_font('Arial', 'B', 16)
                    pdf.cell(0, 10, 'Relatório de Auditoria de Faturamento', ln=True, align='C')
                    pdf.ln(5)
                    
                    # Data
                    pdf.set_font('Arial', '', 10)
                    pdf.cell(0, 6, f'Data de Geração: {datetime.now().strftime("%d/%m/%Y %H:%M")}', ln=True)
                    pdf.ln(10)
                    
                    # Indicadores Principais
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 8, 'Indicadores Principais', ln=True)
                    pdf.ln(3)
                    
                    pdf.set_font('Arial', '', 10)
                    metrics = [
                        f'Total de Solicitações (MV): {total_solicitados:,}',
                        f'Total de Realizações (Lab): {total_realizados:,}',
                        f'Taxa de Realização: {taxa_realizacao:.2f}%',
                        f'Divergências Identificadas: {divergencias:,}',
                        f'Score Médio de Matching: {comp_df["score"].mean():.1f}',
                        f'Registros OK: {len(comp_df[comp_df["status"]=="OK"]):,}',
                        f'Faltam no Lab: {len(comp_df[comp_df["status"]=="FALTAM"]):,}',
                        f'Excesso no Lab: {len(comp_df[comp_df["status"]=="EXCESSO"]):,}'
                    ]
                    
                    for metric in metrics:
                        pdf.cell(0, 6, metric, ln=True)
                    
                    pdf.ln(10)
                    
                    # Análise Temporal
                    if not time_df.empty:
                        pdf.set_font('Arial', 'B', 14)
                        pdf.cell(0, 8, 'Análise de Tempo e SLA', ln=True)
                        pdf.ln(3)
                        
                        pdf.set_font('Arial', '', 10)
                        time_metrics = [
                            f'Tempo Médio de Resposta: {avg_delta:.1f} minutos',
                            f'Tempo Mediano: {median_delta:.1f} minutos',
                            f'Total de Atrasos: {atrasos:,}',
                            f'Taxa de Cumprimento SLA: {taxa_sla:.2f}%',
                            f'Tolerância Configurada: {tolerance} minutos'
                        ]
                        
                        for metric in time_metrics:
                            pdf.cell(0, 6, metric, ln=True)
                        
                        pdf.ln(10)
                    
                    # Análise Financeira
                    if financial_summary:
                        pdf.set_font('Arial', 'B', 14)
                        pdf.cell(0, 8, 'Análise Financeira', ln=True)
                        pdf.ln(3)
                        
                        pdf.set_font('Arial', '', 10)
                        fin_metrics = [
                            f'Total Valor MV: R$ {financial_summary["total_mv"]:,.2f}',
                            f'Total Valor Lab: R$ {financial_summary["total_lab"]:,.2f}',
                            f'Divergência Total: R$ {financial_summary["divergencia"]:,.2f}',
                            f'Valores a Faturar: R$ {financial_summary["faltam"]:,.2f}',
                            f'Valores em Excesso: R$ {financial_summary["excesso"]:,.2f}'
                        ]
                        
                        for metric in fin_metrics:
                            pdf.cell(0, 6, metric, ln=True)
                        
                        pdf.ln(10)
                    
                    # Top Divergências
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 8, 'Top 10 Divergências', ln=True)
                    pdf.ln(3)
                    
                    pdf.set_font('Arial', '', 9)
                    top_div = comp_df[comp_df['diff'] != 0].sort_values('diff', key=abs, ascending=False).head(10)
                    
                    for idx, row in top_div.iterrows():
                        exame_short = row['exame_mv'][:60] if len(row['exame_mv']) > 60 else row['exame_mv']
                        pdf.multi_cell(0, 5, 
                            f"{exame_short}\n   Sol: {row['qtd_mv']} | Real: {row['qtd_lab']} | Dif: {row['diff']} | Score: {row['score']}\n"
                        )
                    
                    # Recomendações
                    pdf.add_page()
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 8, 'Recomendações', ln=True)
                    pdf.ln(3)
                    
                    pdf.set_font('Arial', '', 10)
                    recommendations = [
                        '1. Revisar registros com score abaixo de 60%',
                        '2. Investigar exames com alta divergência de quantidade',
                        '3. Verificar duplicatas identificadas no sistema',
                        '4. Analisar atrasos no cumprimento do SLA',
                        '5. Validar valores financeiros das divergências',
                        '6. Atualizar mapa de sinônimos de exames',
                        '7. Treinar equipe sobre padrões de nomenclatura'
                    ]
                    
                    for rec in recommendations:
                        pdf.multi_cell(0, 6, rec)
                        pdf.ln(2)
                    
                    # Rodapé
                    pdf.ln(10)
                    pdf.set_font('Arial', 'I', 8)
                    pdf.cell(0, 5, 'Relatório gerado automaticamente pelo Sistema de Faturamento', ln=True, align='C')
                    
                    pdf.output(pdf_name)
                    
                    with open(pdf_name, 'rb') as f:
                        st.download_button(
                            '⬇️ Download PDF',
                            f,
                            file_name=pdf_name,
                            mime='application/pdf'
                        )
                    
                    st.success(f"✅ PDF gerado: {pdf_name}")
                    
                except Exception as e:
                    st.error(f"❌ Erro ao gerar PDF: {e}")
    
    # Exportação rápida de divergências
    st.markdown("---")
    st.subheader("⚡ Exportação Rápida")
    
    col_quick1, col_quick2 = st.columns(2)
    
    with col_quick1:
        if st.button("📋 Exportar apenas Divergências (CSV)"):
            div_export = comp_df[comp_df['status'] != 'OK']
            csv = div_export.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                "⬇️ Download Divergências CSV",
                csv,
                f"divergencias_{datetime.now().strftime('%Y%m%d')}.csv",
                "text/csv"
            )
    
    with col_quick2:
        if st.button("🕐 Exportar Análise Temporal (CSV)"):
            if not time_df.empty:
                csv_time = time_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "⬇️ Download Temporal CSV",
                    csv_time,
                    f"analise_temporal_{datetime.now().strftime('%Y%m%d')}.csv",
                    "text/csv"
                )
            else:
                st.warning("Sem dados temporais disponíveis")

# ---------------------- Footer ----------------------
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p><strong>Sistema de Análise de Faturamento Pro</strong></p>
    <p>Versão 2.0 - Melhorado com múltiplos algoritmos e análise financeira</p>
    <p>💡 <em>Dica: Use o menu lateral para ajustar parâmetros de matching</em></p>
</div>
""", unsafe_allow_html=True), ('').str.replace('.', '').str.replace(',', '.').str.strip()
lab['valor_lab'] = pd.to_numeric(lab['valor_lab'], errors='coerce')

# Adicionar material se existir
if lab_material_col:
    lab['material'] = lab[lab_material_col].astype(str)

# Date filtering
if start_date and end_date and start_date > end_date:
    st.error("❌ Data início é maior que data fim")
    st.stop()

mv_original_len = len(mv)
lab_original_len = len(lab)

if start_date:
    if 'ts_mv' in mv.columns:
        mv = mv[mv['ts_mv'].dt.date >= start_date]
    if 'ts_lab' in lab.columns:
        lab = lab[lab['ts_lab'].dt.date >= start_date]

if end_date:
    if 'ts_mv' in mv.columns:
        mv = mv[mv['ts_mv'].dt.date <= end_date]
    if 'ts_lab' in lab.columns:
        lab = lab[lab['ts_lab'].dt.date <= end_date]

if start_date or end_date:
    st.info(f"📅 Filtro aplicado: MV {mv_original_len}→{len(mv)}, Lab {lab_original_len}→{len(lab)}")

# Detect duplicates
duplicates_mv = pd.DataFrame()
duplicates_lab = pd.DataFrame()

if detect_dups:
    dup_cols = ['paciente_norm', 'exame_norm', 'atendimento_id']
    duplicates_mv = detect_duplicates(mv, dup_cols)
    duplicates_lab = detect_duplicates(lab, dup_cols)

# Aggregates
agg_mv = mv.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).agg({
    'ts_mv': 'first',
    **({'valor_mv': 'sum'} if mv_value_col else {})
}).reset_index()
agg_mv['qtd_mv'] = mv.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).size().values

agg_lab = lab.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).agg({
    'ts_lab': 'first',
    **({'valor_lab': 'sum'} if lab_value_col else {})
}).reset_index()
agg_lab['qtd_lab'] = lab.groupby(['atendimento_id', 'paciente_norm', 'exame_norm']).size().values

# ---------------------- Matching Logic ----------------------

def advanced_match(row_mv, agg_lab_data, use_multi=True):
    """Matching avançado com múltiplos algoritmos"""
    best_match = None
    best_score = -1
    
    for _, row_lab in agg_lab_data.iterrows():
        # Score de ID
        score_id = 100 if row_mv['atendimento_id'] and row_mv['atendimento_id'] == row_lab['atendimento_id'] else 0
        
        # Score de paciente
        score_patient = token_set_ratio(row_mv['paciente_norm'], row_lab['paciente_norm'])
        if use_multi:
            score_patient_lev = levenshtein_similarity(row_mv['paciente_norm'], row_lab['paciente_norm'])
            score_patient = (score_patient + score_patient_lev) / 2
        
        # Score de exame
        score_exam = token_set_ratio(row_mv['exame_norm'], row_lab['exame_norm'])
        if use_multi:
            score_exam_lev = levenshtein_similarity(row_mv['exame_norm'], row_lab['exame_norm'])
            score_exam = (score_exam + score_exam_lev) / 2
        
        # Score final ponderado
        total_score = (
            score_patient * weight_patient +
            score_exam * weight_exam +
            score_id * weight_id
        )
        
        if total_score > best_score:
            best_score = total_score
            best_match = row_lab
    
    return best_match, best_score

# Build matches
st.info("🔄 Processando matching entre planilhas...")
matches = []
time_matches = []

for idx, row in agg_mv.iterrows():
    # Try exact match first
    exact = agg_lab[
        (agg_lab['exame_norm'] == row['exame_norm']) & 
        (agg_lab['paciente_norm'] == row['paciente_norm'])
    ]
    
    if not exact.empty:
        match_data = {
            'paciente_mv': row['paciente_norm'],
            'exame_mv': row['exame_norm'],
            'atendimento_mv': row['atendimento_id'],
            'qtd_mv': int(row['qtd_mv']),
            'qtd_lab': int(exact['qtd_lab'].sum()),
            'score': 100,
            'match_type': 'EXACT'
        }
        if mv_value_col and 'valor_mv' in row:
            match_data['valor_mv'] = row['valor_mv']
        if lab_value_col and 'valor_lab' in exact.columns:
            match_data['valor_lab'] = exact['valor_lab'].sum()
        
        matches.append(match_data)
        
        # Time matching
        if pd.notna(row['ts_mv']) and pd.notna(exact.iloc[0]['ts_lab']):
            delta_min = (exact.iloc[0]['ts_lab'] - row['ts_mv']).total_seconds() / 60
            time_matches.append({
                'paciente': row['paciente_norm'],
                'exame': row['exame_norm'],
                'ts_mv': row['ts_mv'],
                'ts_lab': exact.iloc[0]['ts_lab'],
                'delta_min': delta_min,
                'status': 'DENTRO_SLA' if delta_min <= tolerance else 'ATRASO'
            })
    else:
        # Advanced fuzzy match
        best_match, best_score = advanced_match(row, agg_lab, use_multi_algorithm)
        
        match_data = {
            'paciente_mv': row['paciente_norm'],
            'exame_mv': row['exame_norm'],
            'atendimento_mv': row['atendimento_id'],
            'qtd_mv': int(row['qtd_mv']),
            'qtd_lab': int(best_match['qtd_lab']) if best_match is not None else 0,
            'score': int(best_score),
            'match_type': 'FUZZY' if best_score >= fuzzy_threshold else 'NO_MATCH'
        }
        
        if mv_value_col and 'valor_mv' in row:
            match_data['valor_mv'] = row['valor_mv']
        if best_match is not None and lab_value_col and 'valor_lab' in best_match:
            match_data['valor_lab'] = best_match['valor_lab']
        
        matches.append(match_data)

comp_df = pd.DataFrame(matches)
comp_df['diff'] = comp_df['qtd_mv'] - comp_df['qtd_lab']
comp_df['status'] = comp_df.apply(
    lambda r: 'OK' if r['diff'] == 0 else ('FALTAM' if r['diff'] > 0 else 'EXCESSO'),
    axis=1
)

# Calculate financial impact if values exist
financial_summary = None
if 'valor_mv' in comp_df.columns and 'valor_lab' in comp_df.columns:
    comp_df['diff_valor'] = comp_df['valor_mv'].fillna(0) - comp_df['valor_lab'].fillna(0)
    financial_summary = {
        'total_mv': comp_df['valor_mv'].sum(),
        'total_lab': comp_df['valor_lab'].sum(),
        'divergencia': comp_df['diff_valor'].sum(),
        'faltam': comp_df[comp_df['diff_valor'] > 0]['diff_valor'].sum(),
        'excesso': abs(comp_df[comp_df['diff_valor'] < 0]['diff_valor'].sum())
    }

time_df = pd.DataFrame(time_matches) if time_matches else pd.DataFrame()

# ---------------------- Tabs ----------------------
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Dashboard", 
    "🔍 Comparação", 
    "⏱️ Análise Temporal", 
    "🧪 Auditoria",
    "💰 Financeiro",
    "📦 Export"
])

# ---------------------- TAB 1: Dashboard ----------------------
with tab1:
    st.header("📊 Dashboard Executivo")
    
    # KPIs principais
    col1, col2, col3, col4 = st.columns(4)
    
    total_solicitados = len(mv)
    total_realizados = len(lab)
    taxa_realizacao = (total_realizados / total_solicitados * 100) if total_solicitados > 0 else 0
    
    col1.metric("📋 Solicitados", f"{total_solicitados:,}", delta=None)
    col2.metric("✅ Realizados", f"{total_realizados:,}", delta=f"{taxa_realizacao:.1f}%")
    
    divergencias = len(comp_df[comp_df['status'] != 'OK'])
    col3.metric("⚠️ Divergências", f"{divergencias:,}", delta=f"{divergencias/len(comp_df)*100:.1f}%")
    
    if not time_df.empty:
        atrasos = len(time_df[time_df['status'] == 'ATRASO'])
        col4.metric("🕐 Atrasos", f"{atrasos:,}", delta=f"{atrasos/len(time_df)*100:.1f}%")
    else:
        col4.metric("🕐 Atrasos", "N/A")
    
    st.markdown("---")
    
    # Duplicatas
    if detect_dups and (not duplicates_mv.empty or not duplicates_lab.empty):
        st.warning("⚠️ **Duplicatas Detectadas!**")
        coldup1, coldup2 = st.columns(2)
        coldup1.metric("Duplicatas MV", len(duplicates_mv))
        coldup2.metric("Duplicatas LAB", len(duplicates_lab))
    
    # Gráficos
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.subheader("📈 Distribuição por Status")
        status_counts = comp_df['status'].value_counts()
        fig_status = px.pie(
            values=status_counts.values, 
            names=status_counts.index,
            title="Status das Solicitações",
            color_discrete_map={'OK':'green', 'FALTAM':'red', 'EXCESSO':'orange'}
        )
        st.plotly_chart(fig_status, use_container_width=True)
    
    with col_right:
        st.subheader("🎯 Qualidade do Matching")
        match_quality = comp_df['match_type'].value_counts()
        fig_match = px.bar(
            x=match_quality.index, 
            y=match_quality.values,
            title="Tipo de Correspondência",
            labels={'x':'Tipo', 'y':'Quantidade'},
            color=match_quality.index
        )
        st.plotly_chart(fig_match, use_container_width=True)
    
    # Top exames
    st.subheader("🔬 Top 15 Exames Solicitados")
    top_exams = comp_df.groupby('exame_mv').agg({
        'qtd_mv': 'sum',
        'qtd_lab': 'sum',
        'diff': 'sum'
    }).reset_index().sort_values('qtd_mv', ascending=False).head(15)
    
    fig_top = go.Figure()
    fig_top.add_trace(go.Bar(name='Solicitados', x=top_exams['exame_mv'], y=top_exams['qtd_mv']))
    fig_top.add_trace(go.Bar(name='Realizados', x=top_exams['exame_mv'], y=top_exams['qtd_lab']))
    fig_top.update_layout(barmode='group', height=400)
    st.plotly_chart(fig_top, use_container_width=True)
    
    # Alertas
    if show_low_score:
        low_scores = comp_df[comp_df['score'] < fuzzy_threshold]
        if not low_scores.empty:
            st.warning(f"⚠️ **{len(low_scores)} registros com baixa similaridade** (< {fuzzy_threshold}%)")
            with st.expander("Ver registros problemáticos"):
                st.dataframe(low_scores[['paciente_mv', 'exame_mv', 'score', 'status']].head(20))

# ---------------------- TAB 2: Comparação ----------------------
with tab2:
    st.header("🔍 Comparação Detalhada MV x Laboratório")
    
    # Filters
    col_f1, col_f2, col_f3 = st.columns(3)
    
    with col_f1:
        status_filter = st.multiselect(
            "Filtrar por Status",
            options=comp_df['status'].unique(),
            default=None
        )
    
    with col_f2:
        match_filter = st.multiselect(
            "Filtrar por Tipo de Match",
            options=comp_df['match_type'].unique(),
            default=None
        )
    
    with col_f3:
        score_min = st.slider("Score mínimo", 0, 100, 0)
    
    # Apply filters
    filtered_comp = comp_df.copy()
    if status_filter:
        filtered_comp = filtered_comp[filtered_comp['status'].isin(status_filter)]
    if match_filter:
        filtered_comp = filtered_comp[filtered_comp['match_type'].isin(match_filter)]
    if score_min > 0:
        filtered_comp = filtered_comp[filtered_comp['score'] >= score_min]
    
    st.info(f"📊 Mostrando {len(filtered_comp)} de {len(comp_df)} registros")
    
    # Summary
    col_s1, col_s2, col_s3 = st.columns(3)
    col_s1.metric("OK", len(filtered_comp[filtered_comp['status'] == 'OK']))
    col_s2.metric("Faltam", len(filtered_comp[filtered_comp['status'] == 'FALTAM']))
    col_s3.metric("Excesso", len(filtered_comp[filtered_comp['status'] == 'EXCESSO']))
    
    # Top divergências
    st.subheader("🔴 Maiores Divergências")
    top_div = filtered_comp[filtered_comp['diff'] != 0].sort_values('diff', key=abs, ascending=False).head(20)
    
    if not top_div.empty:
        fig_div = px.bar(
            top_div,
            x='diff',
            y='exame_mv',
            orientation='h',
            title='Top 20 Divergências',
            color='status',
            color_discrete_map={'FALTAM':'red', 'EXCESSO':'orange'}
        )
        st.plotly_chart(fig_div, use_container_width=True)
    else:
        st.success("✅ Nenhuma divergência encontrada!")
    
    # Data table
    st.subheader("📋 Tabela Completa")
    st.dataframe(
        filtered_comp.sort_values(['status', 'score'], ascending=[True, False]),
        use_container_width=True,
        height=400
    )

# ---------------------- TAB 3: Análise Temporal ----------------------
with tab3:
    st.header("⏱️ Análise de Tempo e SLA")
    
    if time_df.empty:
        st.warning("⚠️ Sem dados temporais suficientes para análise")
    else:
        # KPIs temporais
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        
        avg_delta = time_df['delta_min'].mean()
        median_delta = time_df['delta_min'].median()
        atrasos = len(time_df[time_df['status'] == 'ATRASO'])
        taxa_sla = (1 - atrasos / len(time_df)) * 100
        
        col_t1.metric("⏱️ Tempo Médio", f"{avg_delta:.1f} min")
        col_t2.metric("📊 Mediana", f"{median_delta:.1f} min")
        col_t3.metric("⚠️ Atrasos", f"{atrasos}", delta=f"-{atrasos/len(time_df)*100:.1f}%")
        col_t4.metric("✅ Taxa SLA", f"{taxa_sla:.1f}%")
        
        # Distribuição temporal
        st.subheader("📈 Distribuição de Tempo de Resposta")
        fig_time_hist = px.histogram(
            time_df,
            x='delta_min',
            nbins=50,
            title='Distribuição do Tempo (minutos)',
            labels={'delta_min': 'Tempo (min)', 'count': 'Frequência'},
            color_discrete_sequence=['#636EFA']
        )
        fig_time_hist.add_vline(x=tolerance, line_dash="dash", line_color="red", 
                                annotation_text=f"SLA: {tolerance} min")
        st.plotly_chart(fig_time_hist, use_container_width=True)
        
        # Evolução temporal
        if 'ts_mv' in time_df.columns:
            time_df['data'] = pd.to_datetime(time_df['ts_mv']).dt.date
            daily_avg = time_df.groupby('data')['delta_min'].agg(['mean', 'count']).reset_index()
            
            fig_evolution = go.Figure()
            fig_evolution.add_trace(go.Scatter(
                x=daily_avg['data'],
                y=daily_avg['mean'],
                mode='lines+markers',
                name='Tempo Médio',
                line=dict(color='blue')
            ))
            fig_evolution.add_hline(y=tolerance, line_dash="dash", line_color="red",
                                   annotation_text="SLA")
            fig_evolution.update_layout(
                title='Evolução do Tempo Médio de Resposta',
                xaxis_title='Data',
                yaxis_title='Tempo Médio (min)',
                height=400
            )
            st.plotly_chart(fig_evolution, use_container_width=True)
        
        # Detalhamento de atrasos
        st.subheader("🔴 Exames com Maior Atraso")
        atrasos_df = time_df[time_df['status'] == 'ATRASO'].sort_values('delta_min', ascending=False).head(20)
        
        if not atrasos_df.empty:
            st.dataframe(
                atrasos_df[['paciente', 'exame', 'delta_min', 'ts_mv', 'ts_lab']],
                use_container_width=True
            )
        else:
            st.success("✅ Nenhum atraso identificado!")

# ---------------------- TAB 4: Auditoria ----------------------
with tab4:
    st.header("🧪 Auditoria e Revisão Manual")
    
    st.markdown("""
    **Objetivo:** Revisar registros com baixa similaridade e fazer correções manuais.
    """)
    
    # Filtros para auditoria
    audit_threshold = st.slider("Mostrar registros com score abaixo de:", 0, 100, fuzzy_threshold)
    
    review_candidates = comp_df[comp_df['score'] < audit_threshold].copy()
    
    if review_candidates.empty:
        st.success("✅ Nenhum registro requer revisão manual com os critérios atuais!")
    else:
        st.warning(f"⚠️ **{len(review_candidates)} registros** requerem revisão")
        
        # Estatísticas
        col_a1, col_a2, col_a3 = st.columns(3)
        col_a1.metric("Baixa Similaridade", len(review_candidates))
        col_a2.metric("Score Médio", f"{review_candidates['score'].mean():.1f}")
        col_a3.metric("Impacto", f"{review_candidates['diff'].sum()}")
        
        st.subheader("📝 Registros para Revisão")
        
        # Editor interativo
        edited_df = st.data_editor(
            review_candidates[['paciente_mv', 'exame_mv', 'atendimento_mv', 
                              'qtd_mv', 'qtd_lab', 'score', 'status', 'match_type']],
            num_rows="dynamic",
            use_container_width=True,
            height=400,
            column_config={
                "score": st.column_config.NumberColumn("Score", format="%.0f"),
                "qtd_mv": st.column_config.NumberColumn("Qtd MV"),
                "qtd_lab": st.column_config.NumberColumn("Qtd Lab (editável)", disabled=False),
            }
        )
        
        # Salvar decisões
        col_save1, col_save2 = st.columns([1, 3])
        
        with col_save1:
            if st.button("💾 Salvar Decisões", type="primary"):
                log_fname = 'audit_log.csv'
                edited_df['review_ts'] = datetime.now()
                edited_df['reviewer'] = 'user'
                
                if os.path.exists(log_fname):
                    existing = pd.read_csv(log_fname)
                    combined = pd.concat([existing, edited_df], ignore_index=True)
                    combined.to_csv(log_fname, index=False)
                else:
                    edited_df.to_csv(log_fname, index=False)
                
                st.success(f"✅ Decisões salvas em {log_fname}")
        
        with col_save2:
            if os.path.exists('audit_log.csv'):
                audit_log = pd.read_csv('audit_log.csv')
                st.info(f"📋 Log de auditoria contém {len(audit_log)} revisões")
        
        # Sugestões automáticas
        st.subheader("💡 Sugestões de Reconciliação")
        
        suggestions = []
        for _, row in review_candidates.head(10).iterrows():
            if row['diff'] > 0:
                suggestions.append({
                    'Paciente': row['paciente_mv'][:30],
                    'Exame': row['exame_mv'][:40],
                    'Ação Sugerida': f"Investigar {row['diff']} solicitações não realizadas",
                    'Prioridade': 'ALTA' if row['diff'] > 5 else 'MÉDIA'
                })
            elif row['diff'] < 0:
                suggestions.append({
                    'Paciente': row['paciente_mv'][:30],
                    'Exame': row['exame_mv'][:40],
                    'Ação Sugerida': f"Verificar {abs(row['diff'])} exames em excesso",
                    'Prioridade': 'MÉDIA'
                })
        
        if suggestions:
            st.dataframe(pd.DataFrame(suggestions), use_container_width=True)

# ---------------------- TAB 5: Financeiro ----------------------
with tab5:
    st.header("💰 Análise Financeira")
    
    if financial_summary is None:
        st.warning("⚠️ Dados de valores não encontrados nas planilhas")
        st.info("Para análise financeira, certifique-se de que há colunas de valores nas planilhas")
    else:
        # KPIs financeiros
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        col_f1.metric("💵 Total MV", f"R$ {financial_summary['total_mv']:,.2f}")
        col_f2.metric("💵 Total Lab", f"R$ {financial_summary['total_lab']:,.2f}")
        col_f3.metric("⚠️ Divergência", f"R$ {financial_summary['divergencia']:,.2f}")
        col_f4.metric("📊 % Divergência", 
                     f"{abs(financial_summary['divergencia'])/financial_summary['total_mv']*100:.2f}%")
        
        st.markdown("---")
        
        # Breakdown financeiro
        col_left, col_right = st.columns(2)
        
        with col_left:
            st.subheader("📉 Análise de Divergências")
            div_data = pd.DataFrame({
                'Categoria': ['A Faturar (Faltam)', 'Faturado em Excesso'],
                'Valor': [financial_summary['faltam'], financial_summary['excesso']]
            })
            fig_fin = px.bar(
                div_data,
                x='Categoria',
                y='Valor',
                title='Impacto Financeiro das Divergências',
                color='Categoria',
                color_discrete_map={'A Faturar (Faltam)': 'red', 'Faturado em Excesso': 'orange'}
            )
            st.plotly_chart(fig_fin, use_container_width=True)
        
        with col_right:
            st.subheader("🎯 Reconciliação")
            reconciliation = pd.DataFrame({
                'Item': ['Total Solicitado (MV)', 'Total Realizado (Lab)', 'Diferença'],
                'Valor': [
                    financial_summary['total_mv'],
                    financial_summary['total_lab'],
                    financial_summary['divergencia']
                ]
            })
            st.dataframe(reconciliation.style.format({'Valor': 'R$ {:,.2f}'}), 
                        use_container_width=True)
        
        # Top divergências financeiras
        st.subheader("💸 Maiores Impactos Financeiros")
        
        if 'diff_valor' in comp_df.columns:
            top_fin = comp_df[comp_df['diff_valor'] != 0].sort_values(
                'diff_valor', 
                key=abs, 
                ascending=False
            ).head(15)
            
            if not top_fin.empty:
                fig_top_fin = px.bar(
                    top_fin,
                    x='diff_valor',
                    y='exame_mv',
                    orientation='h',
                    title='Top 15 Divergências Financeiras',
                    labels={'diff_valor': 'Diferença (R$)', 'exame_mv': 'Exame'},
                    color='diff_valor',
                    color_continuous_scale=['red', 'yellow', 'green']
                )
                st.plotly_chart(fig_top_fin, use_container_width=True)
                
                # Tabela detalhada
                st.dataframe(
                    top_fin[['exame_mv', 'qtd_mv', 'qtd_lab', 'valor_mv', 'valor_lab', 'diff_valor']]
                    .style.format({
                        'valor_mv': 'R$ {:,.2f}',
                        'valor_lab': 'R$ {:,.2f}',
                        'diff_valor': 'R$ {:,.2f}'
                    }),
                    use_container_width=True
                )
        
        # Análise por convênio (se disponível)
        if mv_plan_col and 'convenio_norm' in mv.columns:
            st.subheader("🏥 Análise por Convênio/Plano")
            
            convenio_analysis = mv.groupby('convenio_norm').agg({
                'valor_mv': 'sum' if 'valor_mv' in mv.columns else 'count'
            }).reset_index().sort_values('valor_mv', ascending=False).head(10)
            
            fig_conv = px.pie(
                convenio_analysis,
                values='valor_mv',
                names='convenio_norm',
                title='Distribuição por Convênio'
            )
            st.plotly_chart(fig_conv, use_container_width=True)

# ---------------------- TAB 6: Export ----------------------
with tab6:
    st.header("📦 Exportação de Dados")
    
    st.markdown("""
    Gere relatórios consolidados em Excel e PDF para auditoria e apresentação.
    """)
    
    col_exp1, col_exp2 = st.columns(2)
    
    with col_exp1:
        st.subheader("📊 Excel Consolidado")
        
        if st.button("🔄 Gerar Excel Completo", type="primary"):
            with st.spinner("Gerando arquivo Excel..."):
                out_name = f'relatorio_faturamento_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
                
                try:
                    with pd.ExcelWriter(out_name, engine='openpyxl') as writer:
                        # Dados brutos
                        mv.to_excel(writer, sheet_name='MV_Raw', index=False)
                        lab.to_excel(writer, sheet_name='Lab_Raw', index=False)
                        
                        # Agregados
                        agg_mv.to_excel(writer, sheet_name='MV_Agregado', index=False)
                        agg_lab.to_excel(writer, sheet_name='Lab_Agregado', index=False)
                        
                        # Comparação
                        comp_df.to_excel(writer, sheet_name='Comparacao', index=False)
                        
                        # Dados temporais
                        if not time_df.empty:
                            time_df.to_excel(writer, sheet_name='Analise_Temporal', index=False)
                        
                        # Divergências
                        divergencias_df = comp_df[comp_df['status'] != 'OK']
                        divergencias_df.to_excel(writer, sheet_name='Divergencias', index=False)
                        
                        # Duplicatas
                        if not duplicates_mv.empty:
                            duplicates_mv.to_excel(writer, sheet_name='Duplicatas_MV', index=False)
                        if not duplicates_lab.empty:
                            duplicates_lab.to_excel(writer, sheet_name='Duplicatas_Lab', index=False)
                        
                        # Log de auditoria
                        if os.path.exists('audit_log.csv'):
                            pd.read_csv('audit_log.csv').to_excel(writer, sheet_name='Audit_Log', index=False)
                        
                        # Resumo executivo
                        summary_data = {
                            'Métrica': [
                                'Total Solicitados',
                                'Total Realizados',
                                'Taxa de Realização (%)',
                                'Divergências',
                                'Score Médio',
                                'Duplicatas MV',
                                'Duplicatas Lab'
                            ],
                            'Valor': [
                                total_solicitados,
                                total_realizados,
                                taxa_realizacao,
                                divergencias,
                                comp_df['score'].mean(),
                                len(duplicates_mv),
                                len(duplicates_lab)
                            ]
                        }
                        pd.DataFrame(summary_data).to_excel(writer, sheet_name='Resumo_Executivo', index=False)
                    
                    with open(out_name, 'rb') as f:
                        st.download_button(
                            '⬇️ Download Excel',
                            f,
                            file_name=out_name,
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    
                    st.success(f"✅ Arquivo gerado: {out_name}")
                    
                except Exception as e:
                    st.error(f"❌ Erro ao gerar Excel: {e}")
    
    with col_exp2:
        st.subheader("📄 Relatório PDF")
        
        if st.button("🔄 Gerar PDF de Auditoria", type="primary"):
            with st.spinner("Gerando PDF..."):
                try:
                    pdf_name = f'auditoria_faturamento_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
                    pdf = FPDF()
                    pdf.set_auto_page_break(auto=True, margin=15)
                    pdf.add_page()
                    
                    # Título
                    pdf.set_font('Arial', 'B', 16)
                    pdf.cell(0, 10, 'Relatório de Auditoria de Faturamento', ln=True, align='C')
                    pdf.ln(5)
                    
                    # Data
                    pdf.set_font('Arial', '', 10)
                    pdf.cell(0, 6, f'Data de Geração: {datetime.now().strftime("%d/%m/%Y %H:%M")}', ln=True)
                    pdf.ln(10)
                    
                    # Indicadores Principais
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 8, 'Indicadores Principais', ln=True)
                    pdf.ln(3)
                    
                    pdf.set_font('Arial', '', 10)
                    metrics = [
                        f'Total de Solicitações (MV): {total_solicitados:,}',
                        f'Total de Realizações (Lab): {total_realizados:,}',
                        f'Taxa de Realização: {taxa_realizacao:.2f}%',
                        f'Divergências Identificadas: {divergencias:,}',
                        f'Score Médio de Matching: {comp_df["score"].mean():.1f}',
                        f'Registros OK: {len(comp_df[comp_df["status"]=="OK"]):,}',
                        f'Faltam no Lab: {len(comp_df[comp_df["status"]=="FALTAM"]):,}',
                        f'Excesso no Lab: {len(comp_df[comp_df["status"]=="EXCESSO"]):,}'
                    ]
                    
                    for metric in metrics:
                        pdf.cell(0, 6, metric, ln=True)
                    
                    pdf.ln(10)
                    
                    # Análise Temporal
                    if not time_df.empty:
                        pdf.set_font('Arial', 'B', 14)
                        pdf.cell(0, 8, 'Análise de Tempo e SLA', ln=True)
                        pdf.ln(3)
                        
                        pdf.set_font('Arial', '', 10)
                        time_metrics = [
                            f'Tempo Médio de Resposta: {avg_delta:.1f} minutos',
                            f'Tempo Mediano: {median_delta:.1f} minutos',
                            f'Total de Atrasos: {atrasos:,}',
                            f'Taxa de Cumprimento SLA: {taxa_sla:.2f}%',
                            f'Tolerância Configurada: {tolerance} minutos'
                        ]
                        
                        for metric in time_metrics:
                            pdf.cell(0, 6, metric, ln=True)
                        
                        pdf.ln(10)
                    
                    # Análise Financeira
                    if financial_summary:
                        pdf.set_font('Arial', 'B', 14)
                        pdf.cell(0, 8, 'Análise Financeira', ln=True)
                        pdf.ln(3)
                        
                        pdf.set_font('Arial', '', 10)
                        fin_metrics = [
                            f'Total Valor MV: R$ {financial_summary["total_mv"]:,.2f}',
                            f'Total Valor Lab: R$ {financial_summary["total_lab"]:,.2f}',
                            f'Divergência Total: R$ {financial_summary["divergencia"]:,.2f}',
                            f'Valores a Faturar: R$ {financial_summary["faltam"]:,.2f}',
                            f'Valores em Excesso: R$ {financial_summary["excesso"]:,.2f}'
                        ]
                        
                        for metric in fin_metrics:
                            pdf.cell(0, 6, metric, ln=True)
                        
                        pdf.ln(10)
                    
                    # Top Divergências
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 8, 'Top 10 Divergências', ln=True)
                    pdf.ln(3)
                    
                    pdf.set_font('Arial', '', 9)
                    top_div = comp_df[comp_df['diff'] != 0].sort_values('diff', key=abs, ascending=False).head(10)
                    
                    for idx, row in top_div.iterrows():
                        exame_short = row['exame_mv'][:60] if len(row['exame_mv']) > 60 else row['exame_mv']
                        pdf.multi_cell(0, 5, 
                            f"{exame_short}\n   Sol: {row['qtd_mv']} | Real: {row['qtd_lab']} | Dif: {row['diff']} | Score: {row['score']}\n"
                        )
                    
                    # Recomendações
                    pdf.add_page()
                    pdf.set_font('Arial', 'B', 14)
                    pdf.cell(0, 8, 'Recomendações', ln=True)
                    pdf.ln(3)
                    
                    pdf.set_font('Arial', '', 10)
                    recommendations = [
                        '1. Revisar registros com score abaixo de 60%',
                        '2. Investigar exames com alta divergência de quantidade',
                        '3. Verificar duplicatas identificadas no sistema',
                        '4. Analisar atrasos no cumprimento do SLA',
                        '5. Validar valores financeiros das divergências',
                        '6. Atualizar mapa de sinônimos de exames',
                        '7. Treinar equipe sobre padrões de nomenclatura'
                    ]
                    
                    for rec in recommendations:
                        pdf.multi_cell(0, 6, rec)
                        pdf.ln(2)
                    
                    # Rodapé
                    pdf.ln(10)
                    pdf.set_font('Arial', 'I', 8)
                    pdf.cell(0, 5, 'Relatório gerado automaticamente pelo Sistema de Faturamento', ln=True, align='C')
                    
                    pdf.output(pdf_name)
                    
                    with open(pdf_name, 'rb') as f:
                        st.download_button(
                            '⬇️ Download PDF',
                            f,
                            file_name=pdf_name,
                            mime='application/pdf'
                        )
                    
                    st.success(f"✅ PDF gerado: {pdf_name}")
                    
                except Exception as e:
                    st.error(f"❌ Erro ao gerar PDF: {e}")
    
    # Exportação rápida de divergências
    st.markdown("---")
    st.subheader("⚡ Exportação Rápida")
    
    col_quick1, col_quick2 = st.columns(2)
    
    with col_quick1:
        if st.button("📋 Exportar apenas Divergências (CSV)"):
            div_export = comp_df[comp_df['status'] != 'OK']
            csv = div_export.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                "⬇️ Download Divergências CSV",
                csv,
                f"divergencias_{datetime.now().strftime('%Y%m%d')}.csv",
                "text/csv"
            )
    
    with col_quick2:
        if st.button("🕐 Exportar Análise Temporal (CSV)"):
            if not time_df.empty:
                csv_time = time_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    "⬇️ Download Temporal CSV",
                    csv_time,
                    f"analise_temporal_{datetime.now().strftime('%Y%m%d')}.csv",
                    "text/csv"
                )
            else:
                st.warning("Sem dados temporais disponíveis")

# ---------------------- Footer ----------------------
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; padding: 20px;'>
    <p><strong>Sistema de Análise de Faturamento Pro</strong></p>
    <p>Versão 2.0 - Melhorado com múltiplos algoritmos e análise financeira</p>
    <p>💡 <em>Dica: Use o menu lateral para ajustar parâmetros de matching</em></p>
</div>
""", unsafe_allow_html=True)