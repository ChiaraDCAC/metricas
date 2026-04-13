"""
DCAC Careers - Actualizador de Dashboards
==========================================
Uso:
  1. Poner este archivo en la misma carpeta que el Excel
  2. Configurar las variables de la seccion CONFIGURACION
  3. Doble click en el archivo o correr: python actualizar_dashboard.py

Requiere: pip install pandas openpyxl gitpython
"""

import pandas as pd
import os
import sys
import subprocess
from datetime import datetime

# ============================================================
# CONFIGURACION - EDITAR ESTOS VALORES
# ============================================================

EXCEL_FILE     = "Busquedas_Activas_RRHH.xlsx"   # Nombre del archivo Excel
FECHA_INICIO   = "2026-03-20"                      # Fecha inicio del periodo
FECHA_FIN      = datetime.today().strftime("%Y-%m-%d")  # Hoy automatico
REPO_PATH   = r"C:\Users\Admin\Desktop\metricas-repo"
GITHUB_USER = "ChiaraDCAC"
GITHUB_REPO = "metricas"

# ============================================================


def cargar_datos(excel_file, fecha_inicio, fecha_fin):
    print(f"📂 Leyendo {excel_file}...")
    base = pd.read_excel(excel_file, sheet_name='BASE')
    apps = pd.read_excel(excel_file, sheet_name='Busquedas ACTIVAS RRHH')
    visits = pd.read_excel(excel_file, sheet_name='VISITS')

    apps['Timestamp'] = pd.to_datetime(apps['Timestamp'])
    visits['timestamp'] = pd.to_datetime(visits['timestamp'])

    a = apps[(apps['Timestamp'] >= fecha_inicio) & (apps['Timestamp'] <= fecha_fin + " 23:59:59")]
    v = visits[(visits['timestamp'] >= fecha_inicio) & (visits['timestamp'] <= fecha_fin + " 23:59:59")]

    print(f"✅ Datos cargados: {len(v)} eventos · {len(a)} postulaciones")
    return base, a, v


def calcular_metricas(base, a, v):
    print("🔢 Calculando métricas...")

    # VISITS
    pv = v[v['event_name'] == 'page_view']
    vj = v[v['event_name'] == 'view_job']
    ac = v[v['event_name'] == 'apply_click']

    total_eventos   = len(v)
    trafico         = len(pv)
    fichas_vistas   = len(vj)
    apply_clicks    = len(ac)
    visitantes_uniq = pv['visitor_id'].nunique()

    nuevos      = len(pv[pv['is_returning'] == 'NO'])
    recurrentes = len(pv[pv['is_returning'] == 'SI'])

    mobile  = len(pv[pv['device'] == 'mobile'])
    desktop = len(pv[pv['device'] == 'desktop'])

    # Por dia
    dias = pd.date_range(fecha_inicio, fecha_fin)
    daily = []
    for d in dias:
        ds = d.strftime('%Y-%m-%d')
        dd = pv[pv['timestamp'].dt.date == d.date()]
        dv = vj[vj['timestamp'].dt.date == d.date()]
        da = ac[ac['timestamp'].dt.date == d.date()]
        daily.append({
            'label': d.strftime('%d %b'),
            'uniq': dd['visitor_id'].nunique(),
            'pv': len(dd),
            'vj': len(dv),
            'ac': len(da)
        })

    media_diaria = round(sum(d['uniq'] for d in daily) / len(daily), 1)

    # Conversion por posicion
    vj_by_job = vj.groupby('job_id').size().reset_index(name='vistas')
    ac_exp = ac.copy()
    ac_exp['job_id'] = ac_exp['job_id'].str.split(',')
    ac_exp = ac_exp.explode('job_id')
    ac_exp['job_id'] = ac_exp['job_id'].str.strip()
    ac_by_job = ac_exp.groupby('job_id').size().reset_index(name='clicks')
    conv = vj_by_job.merge(ac_by_job, on='job_id', how='left').fillna(0)
    conv['conv'] = (conv['clicks'] / conv['vistas'] * 100).round(1)
    conv = conv.merge(base[['id', 'posicion']], left_on='job_id', right_on='id', how='left')
    conv = conv.sort_values('vistas', ascending=False).head(10)

    # APPS
    total_apps = len(a)
    a2 = a.copy()
    a2['ID'] = a2['ID Posiciones'].str.split(', ')
    a2 = a2.explode('ID')
    a2['ID'] = a2['ID'].str.strip()
    merged = a2.merge(base[['id', 'posicion', 'Estado', 'Gerencia']], left_on='ID', right_on='id', how='left')
    by_pos = merged.groupby('posicion').size().sort_values(ascending=False).head(16)
    by_ger = merged.groupby('Gerencia').size().sort_values(ascending=False)

    # BASE
    estado   = base['Estado'].value_counts().to_dict()
    gerencia = base['Gerencia'].value_counts().to_dict()
    prioridad = base['Prioridad Empresa'].value_counts().to_dict()
    activas_sin_apps = len(set(base[base['Estado']=='Activa']['posicion']) - set(merged['posicion'].dropna()))

    return {
        'trafico': trafico, 'fichas_vistas': fichas_vistas,
        'apply_clicks': apply_clicks, 'total_apps': total_apps,
        'visitantes_uniq': visitantes_uniq, 'media_diaria': media_diaria,
        'nuevos': nuevos, 'recurrentes': recurrentes,
        'mobile': mobile, 'desktop': desktop,
        'daily': daily, 'conv': conv,
        'by_pos': by_pos, 'by_ger': by_ger,
        'estado': estado, 'gerencia': gerencia,
        'prioridad': prioridad, 'activas_sin_apps': activas_sin_apps,
        'total_busquedas': len(base['id'].unique()),
    }


def generar_html_product(m, fecha_fin, logo_b64, chartjs):
    labels_daily  = str([d['label'] for d in m['daily']])
    data_uniq     = str([d['uniq'] for d in m['daily']])
    data_pv       = str([d['pv'] for d in m['daily']])
    data_vj       = str([d['vj'] for d in m['daily']])
    data_ac       = str([d['ac'] for d in m['daily']])
    media         = m['media_diaria']

    conv_labels = str([r['posicion'] for _, r in m['conv'].iterrows()])
    conv_vals   = str([r['conv'] for _, r in m['conv'].iterrows()])
    conv_vistas = str([int(r['vistas']) for _, r in m['conv'].iterrows()])

    pos_labels = str(list(m['by_pos'].index))
    pos_vals   = str(list(m['by_pos'].values))
    total_apps = m['total_apps']

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DCAC Careers · Product</title>
<script>{chartjs}</script>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500&display=swap');
:root{{--bg:#f4f7fa;--white:#fff;--ink:#0d1a24;--muted:#7a93aa;--border:#d8e4ee;--blue:#3879a3;--blue-pale:#b5d4f4;--accent:#0c447c;--topbar:#3879a3}}
body.dark{{--bg:#0f1923;--white:#162232;--ink:#e8f0f7;--muted:#7a93aa;--border:#1e3248;--blue:#5ba3d0;--blue-pale:#1e3a52;--accent:#85c0e8;--topbar:#0a1520}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:var(--bg);color:var(--ink);font-family:'DM Sans',sans-serif;font-size:13px}}
.topbar{{background:var(--topbar);display:flex;align-items:center;justify-content:space-between;padding:0 32px;height:52px}}
.topbar-right{{display:flex;align-items:center;gap:16px}}
.topbar-meta{{font-family:'DM Mono',monospace;font-size:10px;color:rgba(255,255,255,.4)}}
.toggle-btn{{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;padding:5px 14px;border-radius:20px;cursor:pointer;font-family:'DM Mono',monospace;font-size:10px;letter-spacing:1px}}
.toggle-btn:hover{{background:rgba(255,255,255,.25)}}
.wrap{{max-width:1100px;margin:0 auto;padding:28px 28px 56px}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--border);border:1px solid var(--border);border-radius:10px;overflow:hidden;margin-bottom:24px}}
.kpi{{background:var(--white);padding:16px 18px}}
.kpi-lbl{{font-family:'DM Mono',monospace;font-size:9px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:5px}}
.kpi-val{{font-size:30px;font-weight:300;line-height:1;color:var(--ink)}}
.kpi-val.blue{{color:var(--blue)}}
.kpi-sub{{font-size:10px;color:var(--muted);margin-top:3px}}
.sec{{font-family:'DM Mono',monospace;font-size:9px;text-transform:uppercase;letter-spacing:1.8px;color:var(--muted);margin:20px 0 12px;display:flex;align-items:center;gap:10px}}
.sec::after{{content:'';flex:1;height:1px;background:var(--border)}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}}
.card{{background:var(--white);border:1px solid var(--border);border-radius:10px;padding:18px}}
.ct{{font-family:'DM Mono',monospace;font-size:9px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:12px}}
.chart-wrap{{position:relative;width:100%}}
</style>
</head>
<body>
<div class="topbar">
  <img src="data:image/svg+xml;base64,{logo_b64}" style="height:30px;object-fit:contain">
  <div class="topbar-right">
    <div class="topbar-meta">Portal de Búsquedas · Equipo de Producto · {fecha_fin}</div>
    <button class="toggle-btn" onclick="toggleMode()" id="modeBtn">MODO OSCURO</button>
  </div>
</div>
<div class="wrap">
  <div class="kpis">
    <div class="kpi"><div class="kpi-lbl">Tráfico total</div><div class="kpi-val">{m['trafico']}</div><div class="kpi-sub">page_view · visitas al portal</div></div>
    <div class="kpi"><div class="kpi-lbl">Fichas vistas</div><div class="kpi-val blue">{m['fichas_vistas']}</div><div class="kpi-sub">view_job</div></div>
    <div class="kpi"><div class="kpi-lbl">Clicks aplicar</div><div class="kpi-val blue">{m['apply_clicks']}</div><div class="kpi-sub">apply_click</div></div>
    <div class="kpi"><div class="kpi-lbl">Aplicantes</div><div class="kpi-val blue">{m['total_apps']}</div><div class="kpi-sub">formulario completado</div></div>
  </div>
  <div class="sec">Tráfico diario</div>
  <div class="g2">
    <div class="card"><div class="ct">Visitantes únicos por día · media {media}</div><div class="chart-wrap" style="height:200px"><canvas id="c1"></canvas></div></div>
    <div class="card"><div class="ct">Eventos por tipo y día</div><div class="chart-wrap" style="height:200px"><canvas id="c2"></canvas></div></div>
  </div>
  <div class="sec">Conversión & aplicantes</div>
  <div class="g2">
    <div class="card"><div class="ct">Tasa de conversión por posición (fichas vistas → apply click)</div><div class="chart-wrap" style="height:300px"><canvas id="c3"></canvas></div></div>
    <div class="card"><div class="ct">Top posiciones por aplicantes · total {total_apps}</div><div class="chart-wrap" style="height:300px"><canvas id="c4"></canvas></div></div>
  </div>
  <div class="sec">Comportamiento</div>
  <div class="g2">
    <div class="card"><div class="ct">Dispositivo · sobre page_view</div><div class="chart-wrap" style="height:180px"><canvas id="c5"></canvas></div></div>
    <div class="card"><div class="ct">Nuevos vs. recurrentes · sobre page_view</div><div class="chart-wrap" style="height:180px"><canvas id="c6"></canvas></div></div>
  </div>
</div>
<script>
const C={{blue:'#3879a3',pale:'#b5d4f4',accent:'#0c447c',muted:'#7a93aa',grid:'#eef2f7'}};
const labels={labels_daily},uniq={data_uniq},media={media};
new Chart(document.getElementById('c1'),{{type:'bar',data:{{labels,datasets:[{{data:uniq,backgroundColor:uniq.map((v,i)=>i===0?C.accent:C.pale),borderRadius:4,borderSkipped:false,order:1}},{{type:'line',label:'Media ('+media+')',data:uniq.map(()=>media),borderColor:C.accent,borderWidth:2,borderDash:[5,4],pointRadius:0,fill:false,order:0}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:true,position:'bottom',labels:{{font:{{size:10}},boxWidth:10,padding:10,color:C.muted}}}}}},scales:{{x:{{grid:{{display:false}},ticks:{{font:{{size:11}},color:C.muted}}}},y:{{grid:{{color:C.grid}},ticks:{{font:{{size:10}},color:C.muted}}}}}}}}}});
new Chart(document.getElementById('c2'),{{type:'bar',data:{{labels,datasets:[{{label:'view_job',data:{data_vj},backgroundColor:C.blue,borderRadius:3,borderSkipped:false}},{{label:'page_view',data:{data_pv},backgroundColor:C.pale,borderRadius:3,borderSkipped:false}},{{label:'apply_click',data:{data_ac},backgroundColor:C.accent,borderRadius:3,borderSkipped:false}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:true,position:'bottom',labels:{{font:{{size:10}},boxWidth:10,padding:10,color:C.muted}}}}}},scales:{{x:{{grid:{{display:false}},ticks:{{font:{{size:11}},color:C.muted}}}},y:{{grid:{{color:C.grid}},ticks:{{font:{{size:10}},color:C.muted}}}}}}}}}});
const convLabels={conv_labels},convVals={conv_vals},convVistas={conv_vistas};
new Chart(document.getElementById('c3'),{{type:'bar',data:{{labels:convLabels,datasets:[{{data:convVals,backgroundColor:convVals.map(v=>v>=20?C.accent:C.pale),borderRadius:3,borderSkipped:false}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>` ${{ctx.raw}}% · ${{convVistas[ctx.dataIndex]}} vistas`}}}}}},scales:{{x:{{grid:{{color:C.grid}},ticks:{{callback:v=>v+'%',font:{{size:10}},color:C.muted}},max:35}},y:{{grid:{{display:false}},ticks:{{font:{{size:11}},color:C.muted}}}}}}}}}});
const posLabels={pos_labels},posVals={pos_vals};
new Chart(document.getElementById('c4'),{{type:'bar',data:{{labels:posLabels,datasets:[{{data:posVals,backgroundColor:posVals.map((v,i)=>i===0?C.accent:C.pale),borderRadius:3,borderSkipped:false}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>` ${{ctx.raw}} apps · ${{(ctx.raw/{total_apps}*100).toFixed(1)}}%`}}}}}},scales:{{x:{{grid:{{color:C.grid}},ticks:{{font:{{size:10}},color:C.muted}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}},color:C.muted}}}}}}}}}});
new Chart(document.getElementById('c5'),{{type:'doughnut',data:{{labels:['Mobile {m['mobile']}','Desktop {m['desktop']}'],datasets:[{{data:[{m['mobile']},{m['desktop']}],backgroundColor:[C.blue,C.pale],borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}},boxWidth:10,padding:12,color:C.muted}}}}}}}}}});
new Chart(document.getElementById('c6'),{{type:'doughnut',data:{{labels:['Nuevos {m['nuevos']}','Recurrentes {m['recurrentes']}'],datasets:[{{data:[{m['nuevos']},{m['recurrentes']}],backgroundColor:[C.blue,C.accent],borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}},boxWidth:10,padding:12,color:C.muted}}}}}}}}}});
function toggleMode(){{
  const dark=document.body.classList.toggle('dark');
  document.getElementById('modeBtn').textContent=dark?'MODO DÍA':'MODO OSCURO';
  const blue=dark?'#5ba3d0':'#3879a3',pale=dark?'#1e3a52':'#b5d4f4',accent=dark?'#85c0e8':'#0c447c',grid=dark?'#1e3248':'#eef2f7',muted='#7a93aa';
  ['c1','c2','c3','c4','c5','c6'].forEach(id=>{{
    const ch=Chart.getChart(id);if(!ch)return;
    if(ch.options.scales?.x?.grid)ch.options.scales.x.grid.color=grid;
    if(ch.options.scales?.y?.grid)ch.options.scales.y.grid.color=grid;
    if(ch.options.scales?.x?.ticks)ch.options.scales.x.ticks.color=muted;
    if(ch.options.scales?.y?.ticks)ch.options.scales.y.ticks.color=muted;
    if(ch.options.plugins?.legend?.labels)ch.options.plugins.legend.labels.color=muted;
    ch.data.datasets.forEach(ds=>{{
      if(ds.type==='line'){{ds.borderColor=accent;return;}}
      if(Array.isArray(ds.backgroundColor))ds.backgroundColor=ds.backgroundColor.map(c=>c==='#3879a3'||c==='#5ba3d0'?blue:c==='#b5d4f4'||c==='#1e3a52'?pale:c==='#0c447c'||c==='#85c0e8'?accent:c);
    }});
    ch.update();
  }});
}}
</script>
</body>
</html>"""


def generar_html_rrhh(m, fecha_fin, logo_b64, chartjs):
    ger_labels = str(list(m['by_ger'].index))
    ger_vals   = str(list(m['by_ger'].values))
    pos_labels = str(list(m['by_pos'].index))
    pos_vals   = str(list(m['by_pos'].values))
    total_apps = m['total_apps']
    activas    = m['estado'].get('Activa', 0)
    cerradas   = m['estado'].get('Cerrada', 0)
    inactivas  = m['estado'].get('Inactiva', 0)
    total_b    = m['total_busquedas']
    prio_si    = m['prioridad'].get('Prioridad empresa', 0)
    prio_no    = m['prioridad'].get('No prioridad Empresa', 0)
    ger_b      = m['gerencia']

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DCAC Careers · RRHH</title>
<script>{chartjs}</script>
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500&display=swap');
:root{{--bg:#f4f7fa;--white:#fff;--ink:#0d1a24;--muted:#7a93aa;--border:#d8e4ee;--blue:#3879a3;--blue-pale:#b5d4f4;--accent:#0c447c;--topbar:#3879a3}}
body.dark{{--bg:#0f1923;--white:#162232;--ink:#e8f0f7;--muted:#7a93aa;--border:#1e3248;--blue:#5ba3d0;--blue-pale:#1e3a52;--accent:#85c0e8;--topbar:#0a1520}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:var(--bg);color:var(--ink);font-family:'DM Sans',sans-serif;font-size:13px}}
.topbar{{background:var(--topbar);display:flex;align-items:center;justify-content:space-between;padding:0 32px;height:52px}}
.topbar-right{{display:flex;align-items:center;gap:16px}}
.topbar-meta{{font-family:'DM Mono',monospace;font-size:10px;color:rgba(255,255,255,.4)}}
.toggle-btn{{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;padding:5px 14px;border-radius:20px;cursor:pointer;font-family:'DM Mono',monospace;font-size:10px;letter-spacing:1px}}
.toggle-btn:hover{{background:rgba(255,255,255,.25)}}
.wrap{{max-width:1100px;margin:0 auto;padding:28px 28px 56px}}
.kpis{{display:grid;grid-template-columns:repeat(4,1fr);gap:1px;background:var(--border);border:1px solid var(--border);border-radius:10px;overflow:hidden;margin-bottom:24px}}
.kpi{{background:var(--white);padding:16px 18px}}
.kpi-lbl{{font-family:'DM Mono',monospace;font-size:9px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:5px}}
.kpi-val{{font-size:30px;font-weight:300;line-height:1;color:var(--ink)}}
.kpi-val.blue{{color:var(--blue)}}.kpi-val.accent{{color:var(--accent)}}
.kpi-sub{{font-size:10px;color:var(--muted);margin-top:3px}}
.sec{{font-family:'DM Mono',monospace;font-size:9px;text-transform:uppercase;letter-spacing:1.8px;color:var(--muted);margin:20px 0 12px;display:flex;align-items:center;gap:10px}}
.sec::after{{content:'';flex:1;height:1px;background:var(--border)}}
.g2{{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}}
.g3{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px;margin-bottom:14px}}
.card{{background:var(--white);border:1px solid var(--border);border-radius:10px;padding:18px}}
.ct{{font-family:'DM Mono',monospace;font-size:9px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);margin-bottom:12px}}
.chart-wrap{{position:relative;width:100%}}
</style>
</head>
<body>
<div class="topbar">
  <img src="data:image/svg+xml;base64,{logo_b64}" style="height:30px;object-fit:contain">
  <div class="topbar-right">
    <div class="topbar-meta">Portal de Búsquedas · RRHH · {fecha_fin}</div>
    <button class="toggle-btn" onclick="toggleMode()" id="modeBtn">MODO OSCURO</button>
  </div>
</div>
<div class="wrap">
  <div class="kpis">
    <div class="kpi"><div class="kpi-lbl">Búsquedas activas</div><div class="kpi-val">{activas}</div><div class="kpi-sub">de {total_b} registradas</div></div>
    <div class="kpi"><div class="kpi-lbl">Cerradas</div><div class="kpi-val blue">{cerradas}</div><div class="kpi-sub">{round(cerradas/total_b*100)}% del total</div></div>
    <div class="kpi"><div class="kpi-lbl">Postulaciones</div><div class="kpi-val blue">{total_apps}</div><div class="kpi-sub">20 mar – {fecha_fin}</div></div>
    <div class="kpi"><div class="kpi-lbl">Sin postulaciones</div><div class="kpi-val accent">{m['activas_sin_apps']}</div><div class="kpi-sub">búsquedas activas</div></div>
  </div>
  <div class="sec">Estado de búsquedas</div>
  <div class="g3">
    <div class="card"><div class="ct">Estado general</div><div class="chart-wrap" style="height:200px"><canvas id="c1"></canvas></div></div>
    <div class="card"><div class="ct">Búsquedas por gerencia</div><div class="chart-wrap" style="height:200px"><canvas id="c2"></canvas></div></div>
    <div class="card"><div class="ct">Prioridad empresa</div><div class="chart-wrap" style="height:200px"><canvas id="c3"></canvas></div></div>
  </div>
  <div class="sec">Postulaciones · total {total_apps}</div>
  <div class="g2">
    <div class="card"><div class="ct">Top posiciones por postulaciones</div><div class="chart-wrap" style="height:420px"><canvas id="c4"></canvas></div></div>
    <div class="card"><div class="ct">Postulaciones por gerencia</div><div class="chart-wrap" style="height:420px"><canvas id="c5"></canvas></div></div>
  </div>
</div>
<script>
const C={{blue:'#3879a3',pale:'#b5d4f4',accent:'#0c447c',muted:'#7a93aa',grid:'#eef2f7'}};
new Chart(document.getElementById('c1'),{{type:'doughnut',data:{{labels:['Activa {activas}','Cerrada {cerradas}','Inactiva {inactivas}'],datasets:[{{data:[{activas},{cerradas},{inactivas}],backgroundColor:[C.blue,C.pale,C.muted],borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}},boxWidth:10,padding:12,color:C.muted}}}}}}}}}});
new Chart(document.getElementById('c2'),{{type:'bar',data:{{labels:{str(list(ger_b.keys()))},datasets:[{{data:{str(list(ger_b.values()))},backgroundColor:{str([('#0c447c' if i==0 else '#3879a3' if i<3 else '#b5d4f4') for i in range(len(ger_b))])},borderRadius:3,borderSkipped:false}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}}}},scales:{{x:{{grid:{{color:C.grid}},ticks:{{font:{{size:10}},color:C.muted}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}},color:C.muted}}}}}}}}}});
new Chart(document.getElementById('c3'),{{type:'doughnut',data:{{labels:['No prioridad {prio_no}','Prioridad {prio_si}'],datasets:[{{data:[{prio_no},{prio_si}],backgroundColor:[C.pale,C.accent],borderWidth:0}}]}},options:{{responsive:true,maintainAspectRatio:false,cutout:'65%',plugins:{{legend:{{position:'bottom',labels:{{font:{{size:11}},boxWidth:10,padding:12,color:C.muted}}}}}}}}}});
new Chart(document.getElementById('c4'),{{type:'bar',data:{{labels:{pos_labels},datasets:[{{data:{pos_vals},backgroundColor:{pos_labels}.map((v,i)=>i===0?C.accent:C.pale),borderRadius:3,borderSkipped:false}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>` ${{ctx.raw}} postulaciones · ${{(ctx.raw/{total_apps}*100).toFixed(1)}}%`}}}}}},scales:{{x:{{grid:{{color:C.grid}},ticks:{{font:{{size:10}},color:C.muted}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}},color:C.muted}}}}}}}}}});
new Chart(document.getElementById('c5'),{{type:'bar',data:{{labels:{ger_labels},datasets:[{{data:{ger_vals},backgroundColor:{ger_labels}.map((v,i)=>i===0?C.accent:i<3?C.blue:C.pale),borderRadius:3,borderSkipped:false}}]}},options:{{indexAxis:'y',responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>` ${{ctx.raw}} postulaciones · ${{(ctx.raw/{total_apps}*100).toFixed(1)}}%`}}}}}},scales:{{x:{{grid:{{color:C.grid}},ticks:{{font:{{size:10}},color:C.muted}}}},y:{{grid:{{display:false}},ticks:{{font:{{size:10}},color:C.muted}}}}}}}}}});
function toggleMode(){{
  const dark=document.body.classList.toggle('dark');
  document.getElementById('modeBtn').textContent=dark?'MODO DÍA':'MODO OSCURO';
  const blue=dark?'#5ba3d0':'#3879a3',pale=dark?'#1e3a52':'#b5d4f4',accent=dark?'#85c0e8':'#0c447c',grid=dark?'#1e3248':'#eef2f7',muted='#7a93aa';
  ['c1','c2','c3','c4','c5'].forEach(id=>{{
    const ch=Chart.getChart(id);if(!ch)return;
    if(ch.options.scales?.x?.grid)ch.options.scales.x.grid.color=grid;
    if(ch.options.scales?.y?.grid)ch.options.scales.y.grid.color=grid;
    if(ch.options.scales?.x?.ticks)ch.options.scales.x.ticks.color=muted;
    if(ch.options.scales?.y?.ticks)ch.options.scales.y.ticks.color=muted;
    if(ch.options.plugins?.legend?.labels)ch.options.plugins.legend.labels.color=muted;
    ch.data.datasets.forEach(ds=>{{
      if(Array.isArray(ds.backgroundColor))ds.backgroundColor=ds.backgroundColor.map(c=>c==='#3879a3'||c==='#5ba3d0'?blue:c==='#b5d4f4'||c==='#1e3a52'?pale:c==='#0c447c'||c==='#85c0e8'?accent:c);
    }});
    ch.update();
  }});
}}
</script>
</body>
</html>"""


def subir_github(repo_path, mensaje):
    print("🚀 Subiendo a GitHub...")
    try:
        subprocess.run(["git", "add", "."], cwd=repo_path, check=True)
        subprocess.run(["git", "commit", "-m", mensaje], cwd=repo_path, check=True)
        subprocess.run(["git", "push"], cwd=repo_path, check=True)
        print("✅ Subido correctamente")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error al subir: {e}")
        print("   Verificá que el repositorio esté configurado correctamente")


def main():
    global fecha_inicio, fecha_fin

    # Verificar Excel
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ No se encontró el archivo: {EXCEL_FILE}")
        print("   Asegurate de que el script esté en la misma carpeta que el Excel")
        input("Presioná Enter para salir...")
        sys.exit(1)

    # Leer logo
    logo_path = os.path.join(os.path.dirname(__file__), "Logo_blanco.svg")
    if os.path.exists(logo_path):
        import base64
        with open(logo_path, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
    else:
        logo_b64 = ""
        print("⚠️  Logo no encontrado — el topbar no tendrá imagen")

    # Leer Chart.js embebido
    chartjs_path = os.path.join(os.path.dirname(__file__), "chart.umd.min.js")
    if not os.path.exists(chartjs_path):
        print("❌ Falta chart.umd.min.js — descargalo de:")
        print("   https://cdn.jsdelivr.net/npm/chart.js/dist/chart.umd.min.js")
        input("Presioná Enter para salir...")
        sys.exit(1)

    with open(chartjs_path, "r", encoding="utf-8") as f:
        chartjs = f.read()

    # Calcular métricas
    base, a, v = cargar_datos(EXCEL_FILE, FECHA_INICIO, FECHA_FIN)
    m = calcular_metricas(base, a, v)

    fecha_display = datetime.today().strftime("%d %b %Y")

    # Generar HTMLs
    print("🎨 Generando dashboards...")
    html_product = generar_html_product(m, fecha_display, logo_b64, chartjs)
    html_rrhh    = generar_html_rrhh(m, fecha_display, logo_b64, chartjs)

    # Guardar en repo
    out_product = os.path.join(REPO_PATH, "dashboard_product.html")
    out_rrhh    = os.path.join(REPO_PATH, "dashboard_rrhh.html")

    with open(out_product, "w", encoding="utf-8") as f:
        f.write(html_product)
    with open(out_rrhh, "w", encoding="utf-8") as f:
        f.write(html_rrhh)

    print(f"✅ dashboard_product.html guardado")
    print(f"✅ dashboard_rrhh.html guardado")

    # Subir
    fecha_commit = datetime.today().strftime("%Y-%m-%d")
    subir_github(REPO_PATH, f"actualizar dashboards {fecha_commit}")

    print()
    print("🎉 Listo. Links:")
    print(f"   https://{GITHUB_USER}.github.io/{GITHUB_REPO}/dashboard_product.html")
    print(f"   https://{GITHUB_USER}.github.io/{GITHUB_REPO}/dashboard_rrhh.html")
    input("\nPresioná Enter para cerrar...")


if __name__ == "__main__":
    main()
