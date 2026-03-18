import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import re, unicodedata, io

st.set_page_config(page_title="Gestão de OMs – Einstein | Grupo GPS", page_icon="🔧", layout="wide")

SHEET_URL = st.secrets.get("SHEET_URL", "") if hasattr(st, "secrets") else ""

COLORS = ["#3b82f6","#06b6d4","#10b981","#8b5cf6","#f59e0b","#ef4444","#ec4899",
          "#84cc16","#14b8a6","#f97316","#6366f1","#a855f7","#0ea5e9","#22c55e","#eab308","#64748b"]

PT = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
          font=dict(family="Inter,sans-serif", color="#e8eaf0", size=12),
          margin=dict(l=10, r=10, t=30, b=10))

GC = "rgba(255,255,255,0.06)"  # gridcolor shorthand

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
html,[class*="css"]{font-family:'Inter',sans-serif}
.stApp{background:#0d1117;color:#e8eaf0}
section[data-testid="stSidebar"]{background:#161b24;border-right:1px solid rgba(255,255,255,.07)}
.kpi{background:#161b24;border:1px solid rgba(255,255,255,.08);border-radius:10px;
     padding:16px 20px;position:relative;overflow:hidden;margin-bottom:10px}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--a,#3b82f6)}
.kpi-l{font-size:10px;color:#8b92a5;letter-spacing:.5px;text-transform:uppercase;margin-bottom:5px}
.kpi-v{font-size:26px;font-weight:600;color:#e8eaf0;letter-spacing:-1px}
.kpi-s{font-size:10px;color:#8b92a5;margin-top:3px}
.alert{background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.3);
       border-radius:8px;padding:10px 16px;color:#fca5a5;font-size:13px;margin:8px 0}
.sec{font-size:11px;font-weight:500;color:#8b92a5;letter-spacing:1px;text-transform:uppercase;
     padding-bottom:8px;border-bottom:1px solid rgba(255,255,255,.07);margin:1.2rem 0 .8rem}
.esp-bar-bg{height:7px;border-radius:4px;background:#1c2333;overflow:hidden;display:flex;margin:3px 0}
#MainMenu,footer{visibility:hidden}
</style>""", unsafe_allow_html=True)

# ── HELPERS ────────────────────────────────────────────────────────────────────
def norm(s):
    s = unicodedata.normalize('NFD', str(s).strip().upper())
    return re.sub(r'\s+', ' ', ''.join(c for c in s if unicodedata.category(c) != 'Mn'))

def find_col(cols, *hints):
    nc = {norm(c): c for c in cols}
    for h in hints:
        if norm(h) in nc: return nc[norm(h)]
    for h in hints:
        nh = norm(h)
        for k, v in nc.items():
            if nh in k: return v
    return None

def extract_om(v):
    if pd.isna(v): return ''
    m = re.search(r'\d{5,}', str(v))
    return f"OM {m.group()}" if m else str(v).strip()

def extract_mantenedor(val):
    if not val or str(val).strip() == '': return ''
    s = str(val).strip()
    nomes = []
    for m in re.finditer(r'(?:utilizador|usu[aá]rio)\s*:\s*(.+?)(?=Data\s*:|$)', s, re.IGNORECASE):
        nome = m.group(1)
        for pat in [r'Data\s*:.*', r'\d{2}/\d{2}/\d{4}.*', r'\d{2}:\d{2}.*', r'[Ss]ervi[cç]o.*']:
            nome = re.sub(pat, '', nome, flags=re.IGNORECASE)
        nome = nome.strip().rstrip(',.;:').strip()
        if nome and len(nome) > 1: nomes.append(nome)
    if nomes: return ' / '.join(dict.fromkeys(nomes))
    for pat in [r'^(utilizador|usu[aá]rio|tecnico)\s*:\s*', r'Data\s*:.*', r'\d{2}/\d{2}/\d{4}.*', r'\d{2}:\d{2}.*', r'[Ss]ervi[cç]o.*']:
        s = re.sub(pat, '', s, flags=re.IGNORECASE)
    return s.strip().rstrip(',.;:').strip()

def is_closed(s):
    n = norm(str(s or ''))
    return any(k in n for k in ['FECHA','CONCLU','EXECUT','FINALIZ','ENCERR','RESOLV','APONTAMENTO'])

def classify_status(s):
    n = norm(str(s or ''))
    if is_closed(s): return 'fechada'
    if any(k in n for k in ['ANDAMENTO','EXECUCAO','INICIADA']): return 'andamento'
    return 'pendente'

def to_csv_url(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not m: return url
    gid = (re.search(r'gid=(\d+)', url) or type('', (), {'group': lambda s,x:'0'})()).group(1)
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv&gid={gid}"

def lead_time(row):
    if pd.isna(row['data_abertura']) or pd.isna(row['data_fechamento']): return None
    diff = (row['data_fechamento'] - row['data_abertura']).days
    return max(0, diff)

def median_val(arr):
    if not len(arr): return 0
    s = sorted(arr)
    m = len(s) // 2
    return s[m] if len(s) % 2 else round((s[m-1]+s[m])/2, 1)

def percentile_val(arr, p):
    if not len(arr): return 0
    s = sorted(arr)
    i = max(0, int(len(s)*p/100)-1)
    return s[i]

def delta_html(cur, prev, label):
    if not prev: return f"<span style='color:#8b92a5'>sem dados {label}</span>"
    diff = cur - prev
    pct = abs(round(diff/prev*100)) if prev else 0
    c = "#10b981" if diff >= 0 else "#ef4444"
    arrow = "▲" if diff >= 0 else "▼"
    return f"<span style='color:{c}'>{arrow} {abs(diff)} OMs vs {label}</span> <span style='color:#8b92a5;font-size:10px'>({pct}%)</span>"

@st.cache_data(ttl=120)
def load(url):
    df = pd.read_csv(to_csv_url(url))
    c = df.columns.tolist()
    mp = {
        'om':   find_col(c,'NUMERO DE ORDEM','NUMERO DA ORDEM','OM','ORDEM','OS','NUMERO'),
        'st':   find_col(c,'STATUS','SITUACAO','ESTADO'),
        'uni':  find_col(c,'UNIDADE','UNID','LOCAL','HOSPITAL'),
        'man':  find_col(c,'SERVICO EXECUTADO','SERVICO EXEC','TECNICO DE MANUTENCAO RESPONSAVEL','TECNICO DE MANUTENCAO','MANTENEDOR','TECNICO','RESPONSAVEL'),
        'da':   find_col(c,'DATA DE INICIO DA PROGRAMACAO','DATA DE INICIO DE PROGRAMACAO','DATA INICIO DA PROGRAMACAO','DATA INICIO PROGRAMACAO','DATA ABERTURA','DATA DE ABERTURA','DATA INICIO'),
        'df':   find_col(c,'DATA FINAL DO SERVICO','DATA FINAL DE SERVICO','DATA FINAL SERVICO','DATA FECHAMENTO','DATA DE FECHAMENTO','DATA FINAL','DATA CONCLUSAO'),
    }
    out = pd.DataFrame()
    out['om']         = df[mp['om']].apply(extract_om)          if mp['om']  else ''
    out['status']     = df[mp['st']].astype(str).str.strip()    if mp['st']  else ''
    out['unidade']    = df[mp['uni']].astype(str).str.strip()   if mp['uni'] else ''
    out['mantenedor'] = df[mp['man']].apply(extract_mantenedor) if mp['man'] else ''
    out['data_abertura']   = pd.to_datetime(df[mp['da']], dayfirst=True, errors='coerce') if mp['da'] else pd.NaT
    out['data_fechamento'] = pd.to_datetime(df[mp['df']], dayfirst=True, errors='coerce') if mp['df'] else pd.NaT
    mask = out['om'].str.strip().ne('') | out['status'].str.strip().ne('') | out['unidade'].str.strip().ne('')
    out = out[mask].copy()
    out['lead_time'] = out.apply(lead_time, axis=1)
    return out

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔧 Dashboard Einstein")
    st.markdown("<div style='font-size:12px;color:#8b92a5;margin-bottom:1rem'>Grupo GPS · Análise Operacional</div>", unsafe_allow_html=True)

    with st.expander("⚙️ Google Sheets", expanded=not SHEET_URL):
        inp = st.text_input("URL da planilha", value=SHEET_URL,
                            placeholder="https://docs.google.com/spreadsheets/d/...")
        if inp: st.session_state['url'] = inp
        st.markdown("<div style='font-size:11px;color:#8b92a5;line-height:1.7;margin-top:6px'>1. Abra o Google Sheets<br>2. Compartilhar → Qualquer pessoa com o link<br>3. Cole a URL acima</div>", unsafe_allow_html=True)

    url = st.session_state.get('url', SHEET_URL)
    if not url:
        st.warning("Configure a URL do Google Sheets acima.")
        st.stop()

    if st.button("🔄 Atualizar agora", use_container_width=True):
        st.cache_data.clear()

    try:
        df_all = load(url)
    except Exception as e:
        st.error(f"Erro: {e}")
        st.markdown("<div style='font-size:11px;color:#8b92a5'>Verifique se a planilha está compartilhada como Qualquer pessoa com o link.</div>", unsafe_allow_html=True)
        st.stop()

    if df_all.empty:
        st.warning("Planilha sem dados."); st.stop()

    st.markdown(f"<div style='font-size:11px;color:#10b981;margin-bottom:2px'>✓ {len(df_all)} registros</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:11px;color:#8b92a5;margin-bottom:1rem'>⏱ Atualiza a cada 2 min · {datetime.now().strftime('%d/%m %H:%M')}</div>", unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Filtros rápidos")
    today = datetime.today().date()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Hoje", use_container_width=True): st.session_state['qf'] = 'today'
        if st.button("7 dias", use_container_width=True): st.session_state['qf'] = '7d'
        if st.button("Mês atual", use_container_width=True): st.session_state['qf'] = 'month'
    with col2:
        if st.button("Ontem", use_container_width=True): st.session_state['qf'] = 'yesterday'
        if st.button("30 dias", use_container_width=True): st.session_state['qf'] = '30d'
        if st.button("Mês anterior", use_container_width=True): st.session_state['qf'] = 'lastmonth'

    # Apply quick filter to date range
    qf = st.session_state.get('qf', None)
    datas = df_all['data_fechamento'].dropna()
    default_min = datas.min().date() if not datas.empty else today - timedelta(30)
    default_max = datas.max().date() if not datas.empty else today

    if qf == 'today':      df_min, df_max = today, today
    elif qf == 'yesterday': df_min, df_max = today-timedelta(1), today-timedelta(1)
    elif qf == '7d':       df_min, df_max = today-timedelta(6), today
    elif qf == '30d':      df_min, df_max = today-timedelta(29), today
    elif qf == 'month':    df_min, df_max = today.replace(day=1), today
    elif qf == 'lastmonth':
        pm_end = today.replace(day=1) - timedelta(1)
        df_min, df_max = pm_end.replace(day=1), pm_end
    else: df_min, df_max = default_min, default_max

    st.divider()
    st.markdown("#### Filtros")
    # Agrupa todos os status fechados em "Fechadas" no filtro
    status_raw = sorted(df_all['status'].dropna().unique())
    status_nao_fechados = [s for s in status_raw if not is_closed(s)]
    status_opts_filtro = ['Fechadas'] + sorted(status_nao_fechados)
    status_sel_labels = st.multiselect("Status", status_opts_filtro, default=['Fechadas'])
    # Expande "Fechadas" de volta para os status reais
    status_sel = []
    if 'Fechadas' in status_sel_labels:
        status_sel += [s for s in status_raw if is_closed(s)]
    status_sel += [s for s in status_sel_labels if s != 'Fechadas']

    per = st.date_input("Período (fechamento)", value=(df_min, df_max),
                        min_value=default_min, max_value=default_max)
    d_from, d_to = (per[0], per[1]) if len(per) == 2 else (df_min, df_max)

    uni_opts = ["Todas"] + sorted(df_all['unidade'].dropna().unique())
    uni_sel  = st.selectbox("Unidade", uni_opts)

    man_opts = ["Todos"] + sorted(df_all['mantenedor'].replace('', pd.NA).dropna().unique())
    man_sel  = st.selectbox("Mantenedor", man_opts)

    if st.button("✕ Limpar filtros", use_container_width=True):
        st.session_state['qf'] = None
        st.rerun()

    st.divider()
    st.markdown("#### Compartilhar")
    st.markdown("<div style='font-size:11px;color:#8b92a5;line-height:1.7'>🌐 Mande o link do Streamlit.<br>Gestores só visualizam.<br><br>🖨️ Ctrl+P → Salvar como PDF</div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    df_exp = df_all[df_all['status'].apply(is_closed)].copy()
    if not df_exp.empty:
        te = df_exp[['om','status','unidade','mantenedor','data_abertura','data_fechamento','lead_time']].copy()
        te.columns = ['OM','Status','Unidade','Mantenedor','Data Abertura','Data Fechamento','Lead Time (dias)']
        te['Data Abertura']   = te['Data Abertura'].dt.strftime('%d/%m/%Y').fillna('—')
        te['Data Fechamento'] = te['Data Fechamento'].dt.strftime('%d/%m/%Y').fillna('—')
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w: te.to_excel(w, index=False)
        buf.seek(0)
        st.download_button("📥 Exportar Excel", buf, f"OMs_{datetime.today().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:10px;color:#4b5563;text-align:center'>Gerente: Vanessa Medina<br>Analista: Heverton Sales</div>", unsafe_allow_html=True)

# ── FILTER DATA ────────────────────────────────────────────────────────────────
df = df_all.copy()
if status_sel: df = df[df['status'].isin(status_sel)]
if uni_sel != "Todas": df = df[df['unidade'] == uni_sel]
if man_sel != "Todos": df = df[df['mantenedor'] == man_sel]
if not df['data_fechamento'].isna().all():
    df = df[(df['data_fechamento'].dt.date >= d_from) & (df['data_fechamento'].dt.date <= d_to)]

closed = df[df['status'].apply(is_closed)].copy()
if closed.empty: closed = df.copy()

# Dates for comparisons
week_start  = today - timedelta(days=today.weekday())
month_start = today.replace(day=1)
pw_start    = week_start - timedelta(weeks=1)
pm_end      = month_start - timedelta(days=1)
pm_start    = pm_end.replace(day=1)

# ── HEADER ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='display:flex;align-items:center;justify-content:space-between;
            padding-bottom:1rem;border-bottom:1px solid rgba(255,255,255,.07);margin-bottom:1.25rem'>
  <div>
    <span style='background:#3b82f6;color:#fff;font-size:11px;font-weight:600;
                 letter-spacing:1.5px;padding:4px 10px;border-radius:4px;margin-right:10px'>GPS</span>
    <span style='font-size:15px;font-weight:500'>Gestão de OMs — Contrato Einstein</span>
  </div>
  <div style='font-size:12px;color:#8b92a5'>Grupo GPS · Análise Operacional</div>
</div>""", unsafe_allow_html=True)

# ── TABS ───────────────────────────────────────────────────────────────────────
tab_ger, tab_op, tab_prod, tab_det = st.tabs(["📊 Visão Gerencial", "⚙️ Visão Operacional", "🏆 Produtividade", "🔍 Detalhamento"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: GERENCIAL
# ══════════════════════════════════════════════════════════════════════════════
with tab_ger:
    n       = len(closed)
    n_uni   = closed['unidade'].nunique()
    n_man   = closed['mantenedor'].replace('', pd.NA).dropna().nunique()
    lts     = closed['lead_time'].dropna().tolist()
    lt_med  = round(sum(lts)/len(lts), 1) if lts else None

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    for col, lbl, val, sub, acc in [
        (c1,"Total OMs Fechadas", f"{n:,}".replace(",","."), "no período", "#3b82f6"),
        (c2,"Unidades Atendidas", n_uni, "de 19 previstas", "#06b6d4"),
        (c3,"Mantenedores Ativos", n_man, "no período", "#10b981"),
        (c4,"Média / Unidade", f"{n/n_uni:.1f}" if n_uni else "—", "OMs/unidade", "#8b5cf6"),
        (c5,"Média / Mantenedor", f"{n/n_man:.1f}" if n_man else "—", "OMs/mantenedor", "#f59e0b"),
        (c6,"Lead Time Médio", f"{lt_med}d" if lt_med is not None else "—", "dias p/ fechar", "#10b981"),
    ]:
        with col:
            st.markdown(f'<div class="kpi" style="--a:{acc}"><div class="kpi-l">{lbl}</div>'
                        f'<div class="kpi-v">{val}</div><div class="kpi-s">{sub}</div></div>',
                        unsafe_allow_html=True)

    # Evolução diária + semanal
    col_a, col_b = st.columns([3,2])
    with col_a:
        daily = closed[closed['data_fechamento'].notna()].groupby(closed['data_fechamento'].dt.date).size().reset_index(name='n')
        daily.columns = ['data','n']
        if not daily.empty:
            d_max = daily['n'].max()
            fig = px.line(daily, x='data', y='n', markers=True, labels={'data':'','n':'OMs'},
                          color_discrete_sequence=['#3b82f6'])
            fig.update_traces(line_width=2, marker_size=4, fill='tozeroy', fillcolor='rgba(59,130,246,.07)')
            fig.update_layout(**PT, height=220, title="Evolução Diária de Fechamentos")
            fig.update_yaxes(range=[0, d_max*1.25])
            st.plotly_chart(fig, use_container_width=True)

    with col_b:
        closed2 = closed[closed['data_fechamento'].notna()].copy()
        closed2['semana'] = closed2['data_fechamento'].dt.to_period('W').dt.start_time
        weekly = closed2.groupby('semana').size().reset_index(name='n')
        if not weekly.empty:
            w_max = weekly['n'].max()
            colors_w = ['rgba(16,185,129,.5)'] * len(weekly)
            colors_w[-1] = 'rgba(16,185,129,.9)'
            fig2 = px.bar(weekly, x='semana', y='n', labels={'semana':'','n':'OMs'})
            fig2.update_traces(marker_color=colors_w, marker_line_width=0)
            fig2.update_layout(**PT, height=220, title="Produção Semanal", showlegend=False)
            fig2.update_yaxes(range=[0, w_max*1.25])
            fig2.update_xaxes(tickformat='%d/%m', tickangle=-30)
            st.plotly_chart(fig2, use_container_width=True)

    # Pareto
    st.markdown('<div class="sec">Top 10 Unidades — Pareto</div>', unsafe_allow_html=True)
    by_uni = closed.groupby('unidade').size().reset_index(name='n').sort_values('n', ascending=False).head(10)
    total_uni = by_uni['n'].sum()
    by_uni['pct'] = (by_uni['n'] / total_uni * 100).round(1)
    by_uni['acc'] = by_uni['pct'].cumsum().round(1)

    col_p1, col_p2 = st.columns([1, 1])
    with col_p1:
        # Horizontal bars
        fig_h = px.bar(by_uni.iloc[::-1], x='n', y='unidade', orientation='h', text='n',
                       color='n', color_continuous_scale=[[0,'#1e3a5f'],[1,'#3b82f6']],
                       labels={'unidade':'','n':'OMs'})
        fig_h.update_traces(textposition='outside', textfont_size=11, marker_line_width=0)
        fig_h.update_layout(**PT, showlegend=False, coloraxis_showscale=False, height=320)
        fig_h.update_yaxes(gridcolor='rgba(0,0,0,0)', tickfont_size=10)
        fig_h.update_xaxes(gridcolor=GC)
        st.plotly_chart(fig_h, use_container_width=True)

    with col_p2:
        # Pareto chart
        fig_p = go.Figure()
        fig_p.add_bar(x=by_uni['unidade'], y=by_uni['n'],
                      marker_color=[COLORS[i % len(COLORS)] for i in range(len(by_uni))],
                      name='OMs', yaxis='y')
        fig_p.add_scatter(x=by_uni['unidade'], y=by_uni['acc'],
                          mode='lines+markers', name='% Acumulado',
                          line=dict(color='#f59e0b', width=2),
                          marker=dict(size=5, color='#f59e0b'), yaxis='y2')
        fig_p.add_hline(y=80, line_dash='dash', line_color='rgba(245,158,11,.4)', yref='y2',
                        annotation_text='80%', annotation_position='right')
        fig_p.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter,sans-serif", color="#e8eaf0", size=12),
            margin=dict(l=10, r=10, t=30, b=10),
            height=320, barmode='overlay',
            yaxis2=dict(overlaying='y', side='right', range=[0,110],
                        ticksuffix='%', gridcolor='rgba(0,0,0,0)', title='% Acumulado'),
            legend=dict(orientation='h', y=1.1, font_size=10))
        fig_p.update_xaxes(gridcolor=GC, tickangle=-35, tickfont_size=9)
        fig_p.update_yaxes(gridcolor=GC, title='OMs')
        st.plotly_chart(fig_p, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2: OPERACIONAL
# ══════════════════════════════════════════════════════════════════════════════
with tab_op:
    all_df = df.copy()
    has_dt = not all_df['data_fechamento'].isna().all()
    has_da = not all_df['data_abertura'].isna().all()

    fech_hj  = len(closed[closed['data_fechamento'].dt.date == today]) if has_dt else 0
    ab_hj    = len(all_df[all_df['data_abertura'].dt.date == today]) if has_da else 0
    andamento = len(all_df[all_df['status'].apply(lambda s: classify_status(s) == 'andamento')])
    pendentes = len(all_df[all_df['status'].apply(lambda s: classify_status(s) == 'pendente')])
    total_all = len(all_df) or 1
    semana    = len(closed[closed['data_fechamento'].dt.date >= week_start]) if has_dt else 0
    mes       = len(closed[closed['data_fechamento'].dt.date >= month_start]) if has_dt else 0
    sem_ant   = len(closed[closed['data_fechamento'].dt.date.between(pw_start, week_start-timedelta(1))]) if has_dt else 0
    mes_ant   = len(closed[closed['data_fechamento'].dt.date.between(pm_start, pm_end)]) if has_dt else 0

    c1,c2,c3,c4 = st.columns(4)
    for col, lbl, val, sub, acc in [
        (c1,"✅ Fechadas Hoje", fech_hj, today.strftime('%d/%m/%Y'), "#10b981"),
        (c2,"📅 Abertas Hoje", ab_hj, today.strftime('%d/%m/%Y'), "#3b82f6"),
        (c3,"🔄 Em Andamento", andamento, f"{andamento/total_all*100:.0f}% do total", "#f59e0b"),
        (c4,"⏳ Pendentes", pendentes, f"{pendentes/total_all*100:.0f}% do total", "#ef4444"),
    ]:
        with col:
            color = acc
            st.markdown(f'<div class="kpi" style="--a:{acc}"><div class="kpi-l">{lbl}</div>'
                        f'<div class="kpi-v" style="color:{color}">{val}</div>'
                        f'<div class="kpi-s">{sub}</div></div>', unsafe_allow_html=True)

    # Alerta divergência
    if has_dt and not closed.empty:
        freq = closed['data_fechamento'].dt.date.value_counts()
        if not freq.empty:
            esp = freq.idxmax()
            fora = closed[closed['data_fechamento'].dt.date != esp]
            if not fora.empty:
                st.markdown(f'<div class="alert">⚠ <strong>{len(fora)} OM(s)</strong> com data de fechamento '
                            f'fora do padrão esperado ({esp.strftime("%d/%m/%Y")}). '
                            f'Verifique na aba Detalhamento.</div>', unsafe_allow_html=True)

    # Volume semana e mês com delta
    cv1, cv2 = st.columns(2)
    with cv1:
        st.markdown(f'<div class="kpi" style="--a:#10b981"><div class="kpi-l">Volume na Semana</div>'
                    f'<div class="kpi-v">{semana}</div>'
                    f'<div class="kpi-s">{delta_html(semana,sem_ant,"semana anterior")}</div></div>',
                    unsafe_allow_html=True)
    with cv2:
        st.markdown(f'<div class="kpi" style="--a:#3b82f6"><div class="kpi-l">Volume no Mês</div>'
                    f'<div class="kpi-v">{mes}</div>'
                    f'<div class="kpi-s">{delta_html(mes,mes_ant,"mês anterior")}</div></div>',
                    unsafe_allow_html=True)

    # Status breakdown + donut
    st.markdown('<div class="sec">Distribuição por Status</div>', unsafe_allow_html=True)
    # Status breakdown — agrupa todos fechados em "Fechadas"
    st_map = {}
    for _, row in all_df.iterrows():
        k = 'Fechadas' if is_closed(row['status']) else (row['status'] or '—')
        st_map[k] = st_map.get(k, 0) + 1
    st_sorted = sorted(st_map.items(), key=lambda x: (x[0]!='Fechadas', -x[1]))
    chips = ''.join(
        f"<span style='display:inline-flex;align-items:center;gap:5px;"
        f"background:{color_map.get('fechada' if s=='Fechadas' else classify_status(s),'rgba(139,146,165,.15)')};"
        f"border:1px solid {text_map.get('fechada' if s=='Fechadas' else classify_status(s),'#8b92a5')}33;"
        f"color:{text_map.get('fechada' if s=='Fechadas' else classify_status(s),'#8b92a5')};"
        f"font-size:11px;padding:3px 9px;border-radius:20px;margin:2px'>"
        f"<strong>{n}</strong> {s}</span>"
        for s, n in st_sorted
    )
    st.markdown(f"<div style='display:flex;flex-wrap:wrap;margin-bottom:12px'>{chips}</div>", unsafe_allow_html=True)

    cs1, cs2 = st.columns([2,1])
    with cs1:
        # Mensal
        if has_dt:
            cm = closed.copy()
            cm['mes'] = closed['data_fechamento'].dt.to_period('M').dt.start_time
            monthly = cm.groupby('mes').size().reset_index(name='n')
            if not monthly.empty:
                m_max = monthly['n'].max()
                colors_m = ['rgba(245,158,11,.5)']*len(monthly); colors_m[-1]='rgba(245,158,11,.9)'
                fig_m = px.bar(monthly, x='mes', y='n', labels={'mes':'','n':'OMs'})
                fig_m.update_traces(marker_color=colors_m, marker_line_width=0, text=monthly['n'], textposition='outside')
                fig_m.update_layout(**PT, height=220, title="Produção Mensal", showlegend=False)
                fig_m.update_yaxes(range=[0,m_max*1.3])
                fig_m.update_xaxes(tickformat='%b/%Y', tickangle=-30)
                st.plotly_chart(fig_m, use_container_width=True)
    with cs2:
        fig_d = go.Figure(go.Pie(
            labels=['Fechadas','Em Andamento','Pendentes'],
            values=[len(closed), andamento, pendentes],
            hole=0.62, marker=dict(colors=['#10b981','#f59e0b','#ef4444'], line=dict(color='#161b24',width=2)),
            textinfo='percent', textfont_size=11))
        fig_d.update_layout(**PT, height=220, showlegend=True,
            legend=dict(font_size=10, orientation='h', x=0, y=-0.1))
        st.plotly_chart(fig_d, use_container_width=True)

    # Especiais
    fl = df[df['unidade'].str.upper().str.contains('FARIA LIMA', na=False)]
    kl = df[df['unidade'].str.upper().str.contains('KLABIN', na=False)]
    if not fl.empty or not kl.empty:
        st.markdown('<div class="sec">⭐ Faria Lima & Klabin <span style="background:rgba(245,158,11,.15);color:#fbbf24;font-size:10px;padding:2px 8px;border-radius:10px;margin-left:8px">Liderança Local</span></div>', unsafe_allow_html=True)
        ce1, ce2, ce3 = st.columns([1,1,2])
        for col, data, label, acc in [(ce1,fl,'Faria Lima','#f59e0b'),(ce2,kl,'Klabin','#8b5cf6')]:
            with col:
                tot = len(data); fech = data['status'].apply(is_closed).sum(); pend = tot-fech
                pf = int(fech/tot*100) if tot else 0; pp = 100-pf
                st.markdown(
                    f'<div class="kpi" style="--a:{acc}">'
                    f'<div class="kpi-l">📍 {label}</div>'
                    f'<div class="kpi-v" style="color:{acc};font-size:22px">{tot}</div>'
                    f'<div class="kpi-s"><span style="color:#10b981">{fech} fech.</span> · <span style="color:#ef4444">{pend} pend.</span></div>'
                    f'<div class="esp-bar-bg" style="margin-top:8px"><div style="width:{pf}%;background:#10b981;border-radius:4px 0 0 4px"></div><div style="width:{pp}%;background:#ef4444;border-radius:0 4px 4px 0"></div></div>'
                    f'<div style="display:flex;gap:8px;font-size:10px;color:#8b92a5;margin-top:2px"><span>● Fech {pf}%</span><span>● Pend {pp}%</span></div>'
                    f'</div>', unsafe_allow_html=True)
        with ce3:
            norm_s = lambda s: 'Fechadas' if is_closed(s) else s
            all_sts = sorted(set(fl['status'].apply(norm_s).tolist() + kl['status'].apply(norm_s).tolist()),
                             key=lambda x: (x != 'Fechadas', x))
            fig_esp = go.Figure()
            fig_esp.add_bar(name='Faria Lima', x=all_sts, y=[fl[fl['status'].apply(norm_s)==s].shape[0] for s in all_sts], marker_color='rgba(245,158,11,.8)', text=[fl[fl['status'].apply(norm_s)==s].shape[0] for s in all_sts], textposition='outside')
            fig_esp.add_bar(name='Klabin', x=all_sts, y=[kl[kl['status'].apply(norm_s)==s].shape[0] for s in all_sts], marker_color='rgba(139,92,246,.8)', text=[kl[kl['status'].apply(norm_s)==s].shape[0] for s in all_sts], textposition='outside')
            fig_esp.update_layout(**PT, barmode='group', height=220, showlegend=True,
                legend=dict(font_size=10, orientation='h', x=0, y=1.1))
            fig_esp.update_yaxes(gridcolor=GC)
            fig_esp.update_xaxes(tickfont_size=11)
            st.plotly_chart(fig_esp, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3: PRODUTIVIDADE
# ══════════════════════════════════════════════════════════════════════════════
with tab_prod:
    lts = closed['lead_time'].dropna().tolist()
    lt_med  = round(sum(lts)/len(lts), 1) if lts else None
    lt_med2 = median_val(lts)
    lt_p90  = percentile_val(lts, 90)
    mesmo_dia = len(closed[closed['lead_time'] == 0])

    c1,c2,c3,c4 = st.columns(4)
    for col, lbl, val, sub, acc in [
        (c1,"Lead Time Médio", f"{lt_med}d" if lt_med else "—", "dias (média)", "#3b82f6"),
        (c2,"Lead Time Mediana", f"{lt_med2}d", "dias (mediana)", "#10b981"),
        (c3,"Fechadas Mesmo Dia", mesmo_dia, f"{mesmo_dia/len(closed)*100:.0f}% do total" if len(closed) else "—", "#f59e0b"),
        (c4,"P90 Lead Time", f"{lt_p90}d", "90% fecham em até X dias", "#8b5cf6"),
    ]:
        with col:
            st.markdown(f'<div class="kpi" style="--a:{acc}"><div class="kpi-l">{lbl}</div>'
                        f'<div class="kpi-v" style="color:{acc}">{val}</div>'
                        f'<div class="kpi-s">{sub}</div></div>', unsafe_allow_html=True)

    cp1, cp2 = st.columns([1,1])
    with cp1:
        st.markdown('<div class="sec">Ranking de Mantenedores</div>', unsafe_allow_html=True)
        by_man = (closed[closed['mantenedor'].str.strip().ne('')]
                  .groupby('mantenedor').agg(n=('om','count'), lt=('lead_time','mean'))
                  .reset_index().sort_values('n', ascending=False))
        by_man['lt'] = by_man['lt'].round(1)
        fig_man = px.bar(by_man.head(15), x='n', y='mantenedor', orientation='h', text='n',
                         color='n', color_continuous_scale=[[0,'#1e3a5f'],[1,'#3b82f6']],
                         labels={'mantenedor':'','n':'OMs'})
        fig_man.update_traces(textposition='outside', textfont_size=11, marker_line_width=0)
        fig_man.update_layout(**PT, showlegend=False, coloraxis_showscale=False,
                              height=max(280, len(by_man.head(15))*28))
        fig_man.update_yaxes(gridcolor='rgba(0,0,0,0)', tickfont_size=11)
        fig_man.update_xaxes(gridcolor=GC)
        st.plotly_chart(fig_man, use_container_width=True)

        # Lead time por mantenedor
        if not by_man.empty:
            st.markdown("**Lead time médio por mantenedor (dias)**")
            fig_lt_man = px.bar(by_man.head(12).sort_values('lt'), x='lt', y='mantenedor',
                                orientation='h', text='lt', labels={'mantenedor':'','lt':'dias'},
                                color='lt', color_continuous_scale=[[0,'#10b981'],[0.5,'#f59e0b'],[1,'#ef4444']])
            fig_lt_man.update_traces(textposition='outside', textfont_size=10, marker_line_width=0)
            fig_lt_man.update_layout(**PT, showlegend=False, coloraxis_showscale=False,
                                     height=max(220, len(by_man.head(12))*24))
            fig_lt_man.update_yaxes(gridcolor='rgba(0,0,0,0)', tickfont_size=10)
            fig_lt_man.update_xaxes(gridcolor=GC)
            st.plotly_chart(fig_lt_man, use_container_width=True)

    with cp2:
        st.markdown('<div class="sec">Lead Time por Unidade (dias)</div>', unsafe_allow_html=True)
        by_uni_lt = (closed[closed['lead_time'].notna()]
                     .groupby('unidade')['lead_time'].mean().round(1)
                     .reset_index().sort_values('lead_time', ascending=False).head(12))
        fig_lt = px.bar(by_uni_lt.iloc[::-1], x='lead_time', y='unidade', orientation='h',
                        text='lead_time', labels={'unidade':'','lead_time':'dias'},
                        color='lead_time', color_continuous_scale=[[0,'#10b981'],[0.5,'#f59e0b'],[1,'#ef4444']])
        fig_lt.update_traces(textposition='outside', textfont_size=10, marker_line_width=0)
        fig_lt.update_layout(**PT, showlegend=False, coloraxis_showscale=False,
                             height=max(280, len(by_uni_lt)*28))
        fig_lt.update_yaxes(gridcolor='rgba(0,0,0,0)', tickfont_size=10)
        fig_lt.update_xaxes(gridcolor=GC)
        st.plotly_chart(fig_lt, use_container_width=True)

        # Lead time evolução diária
        lt_day = (closed[closed['data_fechamento'].notna() & closed['lead_time'].notna()]
                  .groupby(closed['data_fechamento'].dt.date)['lead_time'].mean().round(1).reset_index())
        lt_day.columns = ['data','lt']
        if not lt_day.empty:
            fig_lt_d = px.line(lt_day, x='data', y='lt', markers=True,
                               labels={'data':'','lt':'dias'}, color_discrete_sequence=['#8b5cf6'])
            fig_lt_d.update_traces(line_width=2, marker_size=4, fill='tozeroy',
                                   fillcolor='rgba(139,92,246,.07)')
            fig_lt_d.update_layout(**PT, height=200, title="Lead Time Médio Diário")
            st.plotly_chart(fig_lt_d, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4: DETALHAMENTO
# ══════════════════════════════════════════════════════════════════════════════
with tab_det:
    # Expected close date for flags
    exp_date = None
    if not closed.empty and not closed['data_fechamento'].isna().all():
        exp_date = closed['data_fechamento'].dt.date.value_counts().idxmax()

    srch = st.text_input("", placeholder="🔍  Buscar por OM, unidade, mantenedor, status...", label_visibility="collapsed")
    col_cb1, col_cb2 = st.columns([2,1])
    with col_cb2:
        only_alert = st.checkbox("Somente com inconsistência de data")

    tbl = df.copy()
    tbl['status_display'] = tbl['status'].apply(lambda s: 'Fechadas' if is_closed(s) else s)
    if srch:
        m = (tbl['om'].str.contains(srch, case=False, na=False) |
             tbl['unidade'].str.contains(srch, case=False, na=False) |
             tbl['mantenedor'].str.contains(srch, case=False, na=False) |
             tbl['status'].str.contains(srch, case=False, na=False))
        tbl = tbl[m]

    if exp_date:
        tbl = tbl.copy()
        tbl['alerta'] = tbl.apply(lambda r: '⚠ Fora do padrão'
                                  if is_closed(r['status']) and pd.notna(r['data_fechamento'])
                                  and r['data_fechamento'].date() != exp_date else '', axis=1)
        if only_alert:
            tbl = tbl[tbl['alerta'] == '⚠ Fora do padrão']

    tbl_show = tbl[['om','status_display','unidade','mantenedor','data_abertura','data_fechamento','lead_time'] +
                   (['alerta'] if 'alerta' in tbl.columns else [])].copy()
    tbl_show.columns = ['OM','Status','Unidade','Mantenedor','Abertura','Fechamento','Lead Time (d)'] + \
                       (['⚠'] if 'alerta' in tbl.columns else [])
    tbl_show['Abertura']    = tbl_show['Abertura'].dt.strftime('%d/%m/%Y').fillna('—')
    tbl_show['Fechamento']  = tbl_show['Fechamento'].dt.strftime('%d/%m/%Y').fillna('—')
    tbl_show['Lead Time (d)'] = tbl_show['Lead Time (d)'].apply(lambda x: f"{int(x)}d" if pd.notna(x) else '—')

    st.dataframe(tbl_show, use_container_width=True, hide_index=True, height=480)
    st.caption(f"{len(tbl_show)} registros · atualiza a cada 2 min")

    # Export filtered
    buf2 = io.BytesIO()
    tbl_exp2 = tbl_show.copy()
    with pd.ExcelWriter(buf2, engine='openpyxl') as w: tbl_exp2.to_excel(w, index=False)
    buf2.seek(0)
    st.download_button("📥 Exportar tabela filtrada", buf2,
                       f"OMs_Filtrado_{datetime.today().strftime('%Y%m%d')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── FOOTER ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='margin-top:2rem;padding-top:1rem;border-top:1px solid rgba(255,255,255,.07);
            display:flex;justify-content:space-between;font-size:12px;color:#4b5563'>
  <span>Grupo GPS · Contrato Einstein · Dashboard Operacional</span>
  <span>Gerente: <strong style='color:#8b92a5'>Vanessa Medina</strong>
        &nbsp;·&nbsp; Analista: <strong style='color:#8b92a5'>Heverton Sales</strong></span>
</div>""", unsafe_allow_html=True)
