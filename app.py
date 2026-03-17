import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import re, unicodedata, io

st.set_page_config(page_title="Dashboard OMs – Einstein | Grupo GPS", page_icon="🔧", layout="wide")

# URL do Google Sheets — lida do secrets.toml ou configurada na sidebar
SHEET_URL = st.secrets.get("SHEET_URL", "") if hasattr(st, "secrets") else ""

COLORS = ["#3b82f6","#06b6d4","#10b981","#8b5cf6","#f59e0b","#ef4444","#ec4899",
          "#84cc16","#14b8a6","#f97316","#6366f1","#a855f7","#0ea5e9","#22c55e",
          "#eab308","#64748b","#e11d48","#0891b2","#16a34a"]

PT = dict(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
          font=dict(family="Inter,sans-serif", color="#e8eaf0", size=12),
          margin=dict(l=10, r=10, t=30, b=10),
          xaxis=dict(gridcolor="rgba(255,255,255,0.06)"),
          yaxis=dict(gridcolor="rgba(255,255,255,0.06)"))

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
html,[class*="css"]{font-family:'Inter',sans-serif}
.stApp{background:#0d1117;color:#e8eaf0}
section[data-testid="stSidebar"]{background:#161b24;border-right:1px solid rgba(255,255,255,.07)}
.kpi{background:#161b24;border:1px solid rgba(255,255,255,.08);border-radius:10px;
     padding:16px 20px;position:relative;overflow:hidden;margin-bottom:12px}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:var(--a,#3b82f6)}
.kpi-l{font-size:11px;color:#8b92a5;letter-spacing:.5px;text-transform:uppercase;margin-bottom:6px}
.kpi-v{font-size:28px;font-weight:600;color:#e8eaf0;letter-spacing:-1px}
.kpi-s{font-size:11px;color:#8b92a5;margin-top:4px}
.alert{background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.3);
       border-radius:8px;padding:10px 16px;color:#fca5a5;font-size:13px;margin-bottom:12px}
.sec{font-size:11px;font-weight:500;color:#8b92a5;letter-spacing:1px;text-transform:uppercase;
     padding-bottom:8px;border-bottom:1px solid rgba(255,255,255,.07);margin:1.5rem 0 1rem}
#MainMenu,footer{visibility:hidden}
</style>""", unsafe_allow_html=True)

# ── HELPERS ───────────────────────────────────────────────────────────────────
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

def is_closed(s):
    s = norm(str(s or ''))
    return any(k in s for k in ['FECHA','CONCLU','EXECUT','FINALIZ','ENCERR','RESOLV'])

def to_csv_url(url):
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    if not m: return url
    gid = (re.search(r'gid=(\d+)', url) or type('', (), {'group': lambda s,x: '0'})()).group(1)
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv&gid={gid}"

@st.cache_data(ttl=120)
def load(url):
    df = pd.read_csv(to_csv_url(url))
    c = df.columns.tolist()
    mp = {
        'om':   find_col(c,'NUMERO DE ORDEM','NUMERO DA ORDEM','OM','ORDEM','OS','NUMERO'),
        'st':   find_col(c,'STATUS','SITUACAO','ESTADO'),
        'uni':  find_col(c,'UNIDADE','UNID','LOCAL','HOSPITAL'),
        'man':  find_col(c,'TECNICO DE MANUTENCAO RESPONSAVEL','TECNICO DE MANUTENCAO','MANTENEDOR','TECNICO','RESPONSAVEL'),
        'da':   find_col(c,'DATA DE INICIO DA PROGRAMACAO','DATA DE INICIO DE PROGRAMACAO','DATA INICIO PROGRAMACAO','DATA ABERTURA','DATA DE ABERTURA','DATA INICIO'),
        'df':   find_col(c,'DATA FINAL DO SERVICO','DATA FINAL DE SERVICO','DATA FINAL SERVICO','DATA FECHAMENTO','DATA DE FECHAMENTO','DATA FINAL','DATA CONCLUSAO'),
    }
    out = pd.DataFrame()
    out['om']  = df[mp['om']].apply(extract_om) if mp['om'] else ''
    out['status']     = df[mp['st']].astype(str).str.strip()  if mp['st']  else ''
    out['unidade']    = df[mp['uni']].astype(str).str.strip() if mp['uni'] else ''
    out['mantenedor'] = df[mp['man']].astype(str).str.strip() if mp['man'] else ''
    out['data_abertura']   = pd.to_datetime(df[mp['da']], dayfirst=True, errors='coerce') if mp['da'] else pd.NaT
    out['data_fechamento'] = pd.to_datetime(df[mp['df']], dayfirst=True, errors='coerce') if mp['df'] else pd.NaT
    return out[out['om'].str.strip().ne('') | out['status'].str.strip().ne('') | out['unidade'].str.strip().ne('')]

# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🔧 Dashboard Einstein")
    st.markdown("<div style='font-size:12px;color:#8b92a5;margin-bottom:1rem'>Grupo GPS · Análise Operacional</div>", unsafe_allow_html=True)

    with st.expander("⚙️ Google Sheets", expanded=not SHEET_URL):
        inp = st.text_input("URL da planilha", value=SHEET_URL,
                            placeholder="https://docs.google.com/spreadsheets/d/...")
        if inp: st.session_state['url'] = inp
        st.markdown("""<div style='font-size:11px;color:#8b92a5;line-height:1.7;margin-top:8px'>
            1. Abra o Google Sheets<br>
            2. <em>Arquivo → Compartilhar → Qualquer pessoa com o link</em><br>
            3. Cole a URL acima</div>""", unsafe_allow_html=True)

    url = st.session_state.get('url', SHEET_URL)
    if not url:
        st.warning("Configure a URL do Google Sheets acima.")
        st.stop()

    if st.button("🔄 Atualizar agora", use_container_width=True):
        st.cache_data.clear()

    try:
        df_all = load(url)
    except Exception as e:
        st.error(f"Erro ao carregar: {e}")
        st.markdown("<div style='font-size:11px;color:#8b92a5'>Verifique se a planilha está compartilhada como <em>Qualquer pessoa com o link pode ver</em>.</div>", unsafe_allow_html=True)
        st.stop()

    if df_all.empty:
        st.warning("Planilha sem dados reconhecidos.")
        st.stop()

    st.markdown(f"<div style='font-size:11px;color:#10b981;margin-bottom:1rem'>✓ {len(df_all)} registros · {datetime.now().strftime('%d/%m %H:%M')}</div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:11px;color:#8b92a5;margin-bottom:1rem'>⏱ Atualiza automaticamente a cada 2 min</div>", unsafe_allow_html=True)

    st.divider()
    st.markdown("#### Filtros")

    status_opts = sorted(df_all['status'].dropna().unique())
    closed_def  = [s for s in status_opts if is_closed(s)]
    status_sel  = st.multiselect("Status", status_opts, default=closed_def or list(status_opts))

    datas = df_all['data_fechamento'].dropna()
    if not datas.empty:
        mn, mx = datas.min().date(), datas.max().date()
        per = st.date_input("Período (fechamento)", value=(mn, mx), min_value=mn, max_value=mx)
        d_from, d_to = (per[0], per[1]) if len(per)==2 else (mn, mx)
    else:
        d_from = d_to = None

    uni_opts = ["Todas"] + sorted(df_all['unidade'].dropna().unique())
    uni_sel  = st.selectbox("Unidade", uni_opts)

    man_opts = ["Todos"] + sorted(df_all['mantenedor'].dropna().unique())
    man_sel  = st.selectbox("Mantenedor", man_opts)

    st.divider()
    st.markdown("#### Compartilhar")
    st.markdown("""<div style='font-size:11px;color:#8b92a5;line-height:1.7'>
        🌐 <strong style='color:#e8eaf0'>Link público</strong><br>
        Após publicar no Streamlit Cloud você terá um link fixo — mande por WhatsApp ou e-mail.
        Gestores abrem no celular de qualquer lugar, sem instalar nada, sem editar nada.<br><br>
        🖨️ <strong style='color:#e8eaf0'>PDF:</strong> Ctrl+P → Salvar como PDF
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Exportar Excel
    df_exp = df_all[df_all['status'].apply(is_closed)].copy()
    if not df_exp.empty:
        te = df_exp[['om','status','unidade','mantenedor','data_abertura','data_fechamento']].copy()
        te.columns = ['OM','Status','Unidade','Mantenedor','Data Abertura','Data Fechamento']
        te['Data Abertura']   = te['Data Abertura'].dt.strftime('%d/%m/%Y').fillna('—')
        te['Data Fechamento'] = te['Data Fechamento'].dt.strftime('%d/%m/%Y').fillna('—')
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            te.to_excel(w, index=False, sheet_name='OMs Fechadas')
        buf.seek(0)
        st.download_button("📥 Exportar Excel", buf,
                           f"OMs_Einstein_{datetime.today().strftime('%Y%m%d')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:10px;color:#4b5563;text-align:center'>Gerente: Vanessa Medina<br>Analista: Heverton Sales</div>", unsafe_allow_html=True)

# ── FILTRAR ───────────────────────────────────────────────────────────────────
df = df_all.copy()
if status_sel:  df = df[df['status'].isin(status_sel)]
if uni_sel != "Todas": df = df[df['unidade'] == uni_sel]
if man_sel != "Todos": df = df[df['mantenedor'] == man_sel]
if d_from and d_to and not df['data_fechamento'].isna().all():
    df = df[(df['data_fechamento'].dt.date >= d_from) & (df['data_fechamento'].dt.date <= d_to)]

base = df[df['status'].apply(is_closed)]
if base.empty: base = df

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='display:flex;align-items:center;justify-content:space-between;
            padding-bottom:1rem;border-bottom:1px solid rgba(255,255,255,.07);margin-bottom:1.5rem'>
  <div>
    <span style='background:#3b82f6;color:#fff;font-size:11px;font-weight:600;
                 letter-spacing:1.5px;padding:4px 10px;border-radius:4px;
                 text-transform:uppercase;margin-right:12px'>GPS</span>
    <span style='font-size:16px;font-weight:500'>Dashboard de OMs Fechadas — Contrato Einstein</span>
  </div>
  <div style='font-size:12px;color:#8b92a5'>Grupo GPS · Análise Operacional</div>
</div>""", unsafe_allow_html=True)

# ── KPIs ──────────────────────────────────────────────────────────────────────
today        = datetime.today().date()
week_start   = today - timedelta(days=today.weekday())
month_start  = today.replace(day=1)
pw_start     = week_start - timedelta(weeks=1)
pm_end       = month_start - timedelta(days=1)
pm_start     = pm_end.replace(day=1)
has_dt       = not base['data_fechamento'].isna().all()

def cnt(d1, d2): return len(base[base['data_fechamento'].dt.date.between(d1,d2)]) if has_dt else '—'

total    = len(base)
n_uni    = base['unidade'].nunique()
n_man    = base['mantenedor'].nunique()
vol_dia  = cnt(today, today)
vol_sem  = cnt(week_start, today)
vol_mes  = cnt(month_start, today)
vol_sem_ant = cnt(pw_start, week_start-timedelta(1))
vol_mes_ant = cnt(pm_start, pm_end)

def delta(cur, prev):
    if not isinstance(cur,int) or prev==0: return ""
    pct=(cur-prev)/prev*100
    c="#10b981" if pct>=0 else "#ef4444"
    return f"<span style='color:{c};font-size:11px'>{'▲' if pct>=0 else '▼'} {abs(pct):.0f}%</span>"

kpis=[
    ("Total OMs Fechadas",  f"{total:,}".replace(",","."), "no período",           "#3b82f6",""),
    ("Unidades Atendidas",  str(n_uni),                    "de 19 previstas",      "#06b6d4",""),
    ("Mantenedores Ativos", str(n_man),                    "no período",           "#10b981",""),
    ("Média / Unidade",     f"{total/n_uni:.1f}" if n_uni else "—","OMs/unidade", "#8b5cf6",""),
    ("Média / Mantenedor",  f"{total/n_man:.1f}" if n_man else "—","OMs/mantenedor","#f59e0b",""),
    ("Volume Hoje",         str(vol_dia),                  today.strftime('%d/%m/%Y'),"#3b82f6",""),
    ("Volume na Semana",    str(vol_sem),                  "semana atual",         "#06b6d4", delta(vol_sem,vol_sem_ant)),
    ("Volume no Mês",       str(vol_mes),                  today.strftime('%B/%Y'),"#10b981", delta(vol_mes,vol_mes_ant)),
]

for i, row in enumerate([kpis[i:i+4] for i in range(0,8,4)]):
    for col, (lbl,val,sub,acc,tr) in zip(st.columns(4), row):
        with col:
            st.markdown(f'<div class="kpi" style="--a:{acc}"><div class="kpi-l">{lbl}</div>'
                        f'<div class="kpi-v">{val}</div><div class="kpi-s">{sub} {tr}</div></div>',
                        unsafe_allow_html=True)

# ── ALERTA DIVERGÊNCIA ────────────────────────────────────────────────────────
if has_dt:
    freq = base['data_fechamento'].dt.date.value_counts()
    if not freq.empty:
        esp = freq.idxmax()
        fora = base[(base['data_fechamento'].dt.date != esp) & base['status'].apply(is_closed)]
        if not fora.empty:
            st.markdown(f'<div class="alert">⚠ <strong>{len(fora)} OM(s)</strong> com data de fechamento '
                        f'fora do padrão ({esp.strftime("%d/%m/%Y")}). Verifique na tabela abaixo.</div>',
                        unsafe_allow_html=True)

# ── UNIDADE ───────────────────────────────────────────────────────────────────
st.markdown('<div class="sec">Distribuição por Unidade</div>', unsafe_allow_html=True)
by_uni = base.groupby('unidade').size().reset_index(name='n').sort_values('n',ascending=False)

c1,c2 = st.columns([3,2])
with c1:
    f=px.bar(by_uni,x='unidade',y='n',text='n',color='unidade',
             color_discrete_sequence=COLORS,labels={'unidade':'','n':'OMs'})
    f.update_traces(textposition='outside',textfont_size=11,marker_line_width=0)
    f.update_layout(**PT,showlegend=False,height=300); f.update_xaxes(tickangle=-35,tickfont_size=10)
    st.plotly_chart(f,use_container_width=True)

with c2:
    f2=px.pie(by_uni.head(10),values='n',names='unidade',hole=0.6,color_discrete_sequence=COLORS)
    f2.update_traces(textinfo='percent',textfont_size=10)
    f2.update_layout(**PT,height=300,legend=dict(font_size=10,orientation='v',x=1,y=0.5))
    st.plotly_chart(f2,use_container_width=True)

# ── MANTENEDORES ──────────────────────────────────────────────────────────────
st.markdown('<div class="sec">Ranking de Mantenedores</div>', unsafe_allow_html=True)
by_man = base.groupby('mantenedor').size().reset_index(name='n').sort_values('n')
c3,c4 = st.columns([3,2])

with c3:
    f3=px.bar(by_man,x='n',y='mantenedor',orientation='h',text='n',
              color='n',color_continuous_scale=[[0,'#1e3a5f'],[1,'#3b82f6']],
              labels={'mantenedor':'','n':'OMs'})
    f3.update_traces(textposition='outside',textfont_size=11,marker_line_width=0)
    f3.update_layout(**PT,showlegend=False,coloraxis_showscale=False,
                     height=max(280,len(by_man)*28))
    f3.update_yaxes(tickfont_size=11)
    st.plotly_chart(f3,use_container_width=True)

with c4:
    st.markdown("**Top 5 Unidades**")
    top5 = by_uni.head(5); mx=top5['n'].max() or 1
    for i,r in enumerate(top5.itertuples()):
        st.markdown(f"""
        <div style='display:flex;align-items:center;gap:10px;padding:7px 0;
                    border-bottom:1px solid rgba(255,255,255,.06)'>
          <span style='font-size:11px;color:#8b92a5;min-width:22px'>#{i+1}</span>
          <div style='flex:1'>
            <div style='font-size:13px;color:#e8eaf0;margin-bottom:3px'>{r.unidade}</div>
            <div style='height:5px;border-radius:3px;background:#1c2333;overflow:hidden'>
              <div style='width:{int(r.n/mx*100)}%;height:100%;background:{COLORS[i]};border-radius:3px'></div>
            </div>
          </div>
          <span style='font-weight:600;color:{COLORS[i]};min-width:32px;text-align:right'>{r.n}</span>
          <span style='font-size:11px;color:#8b92a5;min-width:40px;text-align:right'>{r.n/len(base)*100:.1f}%</span>
        </div>""", unsafe_allow_html=True)

# ── TEMPORAL ──────────────────────────────────────────────────────────────────
st.markdown('<div class="sec">Evolução Temporal</div>', unsafe_allow_html=True)
if has_dt:
    t1,t2,t3,t4 = st.tabs(["📅 Diário","📆 Semanal","🗓️ Mensal","📊 Comparativo"])
    with t1:
        d=base.groupby(base['data_fechamento'].dt.date).size().reset_index(name='n')
        d.columns=['data','n']
        fd=px.line(d,x='data',y='n',markers=True,color_discrete_sequence=['#3b82f6'],labels={'data':'','n':'OMs'})
        fd.update_traces(line_width=2,marker_size=5,fill='tozeroy',fillcolor='rgba(59,130,246,.07)')
        fd.update_layout(**PT,height=250); st.plotly_chart(fd,use_container_width=True)
    with t2:
        ws=base.copy(); ws['s']=base['data_fechamento'].dt.to_period('W').dt.start_time
        ws=ws.groupby('s').size().reset_index(name='n')
        fs=px.bar(ws,x='s',y='n',text='n',color_discrete_sequence=['#10b981'],labels={'s':'','n':'OMs'})
        fs.update_traces(textposition='outside',marker_line_width=0)
        fs.update_layout(**PT,height=250); st.plotly_chart(fs,use_container_width=True)
    with t3:
        mm=base.copy(); mm['m']=base['data_fechamento'].dt.to_period('M').dt.start_time
        mm=mm.groupby('m').size().reset_index(name='n')
        fm=px.bar(mm,x='m',y='n',text='n',color_discrete_sequence=['#f59e0b'],labels={'m':'','n':'OMs'})
        fm.update_traces(textposition='outside',marker_line_width=0)
        fm.update_layout(**PT,height=250); st.plotly_chart(fm,use_container_width=True)
    with t4:
        periodos={"Esta semana":(week_start,today),"Semana passada":(pw_start,week_start-timedelta(1)),
                  "Este mês":(month_start,today),"Mês passado":(pm_start,pm_end)}
        cp=pd.DataFrame([{"Período":k,"OMs":cnt(d1,d2)} for k,(d1,d2) in periodos.items()])
        fc=px.bar(cp,x='Período',y='OMs',text='OMs',color='Período',color_discrete_sequence=COLORS)
        fc.update_traces(textposition='outside',marker_line_width=0)
        fc.update_layout(**PT,showlegend=False,height=250); st.plotly_chart(fc,use_container_width=True)
else:
    st.info("Coluna de data de fechamento não detectada — gráficos temporais indisponíveis.")

# ── TABELA ────────────────────────────────────────────────────────────────────
st.markdown('<div class="sec">Detalhamento de OMs</div>', unsafe_allow_html=True)
srch = st.text_input("", placeholder="🔍  Buscar por OM, unidade ou mantenedor...", label_visibility="collapsed")
tbl = base.copy()
if srch:
    m=(tbl['om'].str.contains(srch,case=False,na=False)|
       tbl['unidade'].str.contains(srch,case=False,na=False)|
       tbl['mantenedor'].str.contains(srch,case=False,na=False)|
       tbl['status'].str.contains(srch,case=False,na=False))
    tbl=tbl[m]

if has_dt and not tbl['data_fechamento'].isna().all():
    esp2=tbl['data_fechamento'].dt.date.value_counts().idxmax()
    tbl=tbl.copy()
    tbl['⚠']=tbl['data_fechamento'].dt.date.apply(lambda d: '⚠ Fora do padrão' if pd.notna(d) and d!=esp2 else '')

show_cols=['om','status','unidade','mantenedor','data_abertura','data_fechamento']+(['⚠'] if '⚠' in tbl.columns else [])
tbl2=tbl[show_cols].copy()
tbl2.columns=['OM','Status','Unidade','Mantenedor','Data Abertura','Data Fechamento']+(['⚠'] if '⚠' in tbl.columns else [])
tbl2['Data Abertura']  =tbl2['Data Abertura'].dt.strftime('%d/%m/%Y').fillna('—')
tbl2['Data Fechamento']=tbl2['Data Fechamento'].dt.strftime('%d/%m/%Y').fillna('—')
st.dataframe(tbl2,use_container_width=True,hide_index=True,height=420)
st.caption(f"{len(tbl2)} registros · atualiza a cada 2 min")

# ── RODAPÉ ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='margin-top:2rem;padding-top:1rem;border-top:1px solid rgba(255,255,255,.07);
            display:flex;justify-content:space-between;font-size:12px;color:#4b5563'>
  <span>Grupo GPS · Contrato Einstein · Dashboard Operacional</span>
  <span>Gerente: <strong style='color:#8b92a5'>Vanessa Medina</strong>
        &nbsp;·&nbsp; Analista: <strong style='color:#8b92a5'>Heverton Sales</strong></span>
</div>""", unsafe_allow_html=True)
