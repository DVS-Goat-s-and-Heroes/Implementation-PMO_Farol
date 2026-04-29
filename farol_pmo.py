"""
Farol PMO - Traffic Light Report Generator
Reads an Implementation Plan .xlsx and generates an HTML traffic light report.
Usage: python3 farol_pmo.py <path_to_excel_file>
"""

import sys
import os
import pandas as pd
from datetime import datetime, date

# ─── Stage weights (from email brief) ────────────────────────────────────────
STAGE_WEIGHTS = [
    {"label": "Initiation",                   "weight": 0.05, "min": 201, "max": 299},
    {"label": "Design Mapeamento operacional","weight": 0.30, "min": 401, "max": 499},
    {"label": "Contrato",                     "weight": 0.10, "min": 301, "max": 399},
    {"label": "Systems (execução)",           "weight": 0.30, "min": 503, "max": 503},
    {"label": "Cadastros",                    "weight": 0.10, "min": 502, "max": 502},
    {"label": "Go Live",                      "weight": 0.10, "min": 601, "max": 699},
    {"label": "Evaluation",                   "weight": 0.05, "min": 701, "max": 799},
]

STATUS_SCORE = {
    "finished":     1.0,
    "on track":     0.5,
    "not started":  0.0,
    "postponed":    0.3,
    "delayed":      0.3,
    "overdue":      0.1,
    "on hold":      0.0,
}

# Lê a Planilha retorna o DataFrame
def load_data(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath, sheet_name="Implementation Plan", header=None)

    # Lemos a Tabela e nomeamos as colunas do DataFrame
    data = df.iloc[6:, [1, 2, 3, 4, 5, 6, 7, 8, 9]].copy()
    data.columns = ["task_id", "subtask_id", "task", "responsible",
                    "start_date", "end_date", "days", "status", "remarks"]
    
    # Lemos apenas onde 'task' está preenchido
    data = data[data["task"].notna() & (data["task"].astype(str).str.strip() != "")]
    data = data[~data["task"].astype(str).str.contains("Task ID|Sub Task ID", na=False)] # ????????
    data["task_id"]    = data["task_id"].astype(str).str.strip().replace("nan", "")
    data["subtask_id"] = data["subtask_id"].astype(str).str.strip().replace("nan", "")
    data["status_raw"] = data["status"].astype(str).str.strip()
    data["status_norm"]= data["status_raw"].str.lower().str.strip()
    data["start_date"] = pd.to_datetime(data["start_date"], errors="coerce")
    data["end_date"]   = pd.to_datetime(data["end_date"],   errors="coerce")
    # Retorna o DataFrame lido

    return data.reset_index(drop=True)

# ─── NOVA LÓGICA: Extração de Metadados do Cabeçalho ─────────────────────────
# ─── NOVA LÓGICA: Extração de Metadados do Cabeçalho ─────────────────────────
def get_project_metadata(filepath: str):
    """Lê as primeiras 10 linhas do Excel para buscar o Responsável e a Última Atualização"""
    try:
        # Lê apenas o topo da planilha para ficar rápido
        df = pd.read_excel(filepath, sheet_name="Implementation Plan", header=None, nrows=10)
        latest_update = "Não informado"
        responsible = "Não informado"
        
        for r_idx, row in df.iterrows():
            # Interrompe a busca na linha 6 para NÃO invadir o cabeçalho da tabela de tarefas
            if r_idx > 5:
                break
                
            row_list = row.tolist() # Passa as colunas para a lista
            for i, cell in enumerate(row_list):
                cell_str = str(cell).strip().lower()
                
                # Procura a célula "Latest update:"
                if cell_str.startswith("latest update"):
                    # 1º TENTATIVA: Verifica se a data ficou presa na mesma célula (ex: "Latest update: 15/04/2026")
                    if ":" in str(cell):
                        parts = str(cell).split(":", 1)
                        if len(parts) > 1 and parts[1].strip() != "":
                            val = parts[1].strip()
                            try:
                                latest_update = pd.to_datetime(val).strftime("%d/%m/%Y")
                            except:
                                latest_update = val
                            continue # Achou aqui, vai pro próximo loop
                            
                    # 2º TENTATIVA: Olha no máximo as próximas células ao lado para achar a data
                    for val in row_list[i+1:i+6]:
                        if pd.notna(val) and str(val).strip() != "":
                            try:
                                latest_update = pd.to_datetime(val).strftime("%d/%m/%Y")
                            except:
                                latest_update = str(val).strip()
                            break
                            
                # Procura a célula "Responsible:"
                elif cell_str.startswith("responsible"):
                    # 1º TENTATIVA: Verifica se o nome ficou preso na mesma célula (ex: "Responsible: Luiza")
                    if ":" in str(cell):
                        parts = str(cell).split(":", 1)
                        if len(parts) > 1 and parts[1].strip() != "":
                            responsible = parts[1].strip()
                            continue # Achou aqui, vai pro próximo loop

                    # 2º TENTATIVA: Olha as próximas células, mas ignora textos de status como "overdue"
                    for val in row_list[i+1:i+8]:
                        val_str = str(val).strip().lower()
                        if pd.notna(val) and val_str != "" and "overdue" not in val_str:
                            responsible = str(val).strip()
                            break
                            
        return latest_update, responsible
    except Exception:
        return "Não informado", "Não informado"

# ─── NOVA LÓGICA: Extração de Datas Específicas ──────────────────────────────
def get_milestone_date(data: pd.DataFrame, task_id: str) -> str:
    """Busca a data final (end_date) de uma tarefa específica pelo ID"""
    row = data[data["task_id"] == str(task_id)] #Vai buscar a task ID inserida na hora da função
    if not row.empty:
        return fmt_date(row["end_date"].iloc[0]) #Chama a função para a formatação de data de XX.XX.XX para XX/XX/XX
    return "Não definido"

# Função de
def task_num(row) -> float | None:
    try:
        tid = row["task_id"] # TASK ID
        sid = row["subtask_id"] # SUBTASK ID

        if tid:
            print("\n\n TID: ", tid, "\n\n")
            return float(tid) # Retorna a task_id
        elif sid:
            print("\n\n SID: ", sid, "\n\n")
            return float(sid.split(".")[0]) # Retorna a task_id dessa subtask (o número inteiro do index)
        
    except (ValueError, TypeError):
        return "Noness"

def status_color(status_norm: str, end_date=None) -> str:
    today = datetime.today()
    if "finished" in status_norm:
        return "blue"
    if "on track" in status_norm:
        if end_date is not None and pd.notna(end_date):
            delta = (end_date - today).days
            if delta < 0:
                return "red"
            if delta <= 2:
                return "yellow"
        return "green"
    if "overdue" in status_norm:
        return "red"
    if "delayed" in status_norm or "postponed" in status_norm:
        return "yellow"
    if "not started" in status_norm or "on hold" in status_norm:
        return "gray"
    
    if end_date is not None and pd.notna(end_date):
        delta = (end_date - today).days
        if delta < 0:
            return "red"
        if delta <= 2:
            return "yellow"
    return "gray"

def pct_for_tasks(rows: pd.DataFrame) -> float:
    if rows.empty:
        return 0.0
    scores = rows["status_norm"].map(
        lambda s: next((v for k, v in STATUS_SCORE.items() if k in s), 0.0)
    )
    return round(scores.mean() * 100, 1)

def stage_dominant_color(rows: pd.DataFrame) -> str:
    if rows.empty:
        return "gray"
    colors = rows.apply(lambda r: status_color(r["status_norm"], r["end_date"]), axis=1)
    priority = ["red", "yellow", "green", "blue", "gray"]
    for c in priority:
        if (colors == c).any():
            return c
    return "gray"

# Construindo os estágios (tasks)
def build_stages(data: pd.DataFrame):
    data = data.copy()
    data["_tnum"] = None

    # Adicionamos os números das tasks
    for idx, row in data.iterrows():
        if idx == 0:
            continue

        data.loc[idx, "_tnum"] = task_num(row)
        print("VALOR NA LINHA: ", data.loc[idx, "_tnum"])
        input("")

    #data.apply(task_num, axis=1) # Cria coluna de task number com os números das tasks
    
    for idx, row in data.iterrows():
        print("\n\n LINHA DATA:", row)
        input("")

    stages = []
    
    # Percorre o dicionário STAGE_WEIGHTS (label; weight; min; max)
    for conf in STAGE_WEIGHTS:
        mask = data["_tnum"].notna() & (data["_tnum"] >= conf["min"]) & (data["_tnum"] <= conf["max"]) # Dos valores da coluna de task number, batemos com as 'conf' min e max                                                                                     
        subset = data[mask].copy() # Passa o subset filtrado com base no peso das Tasks verificado

        # for idx, row in subset.iterrows():
            # print("\n\n LINHA DO SUBSET:", row)
            # input("")

        tasks_out = []
        
        # Percorre as linhas do sub(Data)set filtrado para essa iteração do STAGE_WEIGHTS
        for _, row in subset.iterrows(): # _ recupera a quantidade de linhas, row recupera os valores da linha percorrida

            print("\n\nDEBUG  LINHA DO SUBSET\n", row, "\nDEBUG  LINHA DO SUBSET\n")

            tid = row["task_id"] # recupera o task_id da linha
            sid = row["subtask_id"] # recupera o subtask_id da linha

            if tid or sid:
                is_subtask = bool(sid) # Verifica se a linha lida do subset é uma subtask
                current_id = sid if is_subtask else tid # Coloca o número da TASK / SUBTASK
                
                tasks_out.append({ # Adiciona um dicionário à lista de outputs das Tasks (e Substasks também)
                    "id": current_id, 
                    "is_subtask": is_subtask,
                    "name": str(row["task"]),
                    "responsible": str(row["responsible"]) if row["responsible"] else "",
                    "start": row["start_date"], 
                    "end":   row["end_date"],
                    "status": row["status_raw"],
                    "status_norm": row["status_norm"],
                    "color": status_color(row["status_norm"], row["end_date"]),
                    "remarks": str(row["remarks"]) if pd.notna(row["remarks"]) else "",
                })

        stages.append({ # Adiciona um dicionário à lista de estágios
            "label": conf["label"],
            "weight": conf["weight"],
            "pct": pct_for_tasks(subset),
            "color": stage_dominant_color(subset),
            "tasks": tasks_out
        })

    return stages

def overall_pct(stages) -> float:
    return round(sum(s["pct"] * s["weight"] for s in stages), 1) #Calcula, dentro de cada estágio, a label "pct" dela * a label "weight", retornando no cálculo uma só casa decimal 

def get_overall_color(stages) -> str: 
    colors = [s["color"] for s in stages] #Filtra a label "color" dentro de stages
    for c in ["red", "yellow", "green", "blue"]:
        if c in colors: #Retorna a cor correspondente a etapa do projeto
            return c
    return "gray" 

def fmt_date(d) -> str:
    if d is None or (isinstance(d, float) and pd.isna(d)): # Verifica se o data representa essas condições
        return "—"
    try:
        if pd.isna(d):
            return "—"
        return pd.Timestamp(d).strftime("%d/%m/%Y") #Retorna o dado inputado na formatação de data
    except Exception:
        return "—"

COLOR_HEX = {
    "green":  "#22c55e",
    "yellow": "#eab308",
    "red":    "#ef4444",
    "blue":   "#3b82f6",
    "gray":   "#94a3b8",
}
COLOR_LABEL = {
    "green":  "On Track",
    "yellow": "Atenção",
    "red":    "Atrasado",
    "blue":   "Concluído",
    "gray":   "Não Iniciado",
}

def render_html(stages, filepath: str, meta: dict) -> str:
    overall = overall_pct(stages) #Porcentagem total do projeto
    overall_color = get_overall_color(stages) #A cor de andamento total do projeto
    today_str = datetime.today().strftime("%d/%m/%Y") #Data atual na formatação XX/XX/XX
    filename = os.path.basename(filepath) #nome do arquivo .html

    def dot(color, size=14):
        hex_ = COLOR_HEX.get(color, "#94a3b8")
        return f'<span class="dot" style="background:{hex_};width:{size}px;height:{size}px;box-shadow:0 0 6px {hex_}88"></span>' #Cria os pontos de cor do farol

    def build_farol(active_color):
        colors = ["gray", "red", "yellow", "green", "blue"]
        lights = ""
        for c in colors:
            hex_ = COLOR_HEX.get(c, "#94a3b8") #Seta a cor padrão para cinza
            style = f'background:{hex_}; opacity:1; box-shadow:0 0 8px {hex_}' if c == active_color else f'background:{hex_};' #Função para buscar a cor real da task
            lights += f'<div class="farol-luz" style="{style}"></div>' #Soma na variável as luzes do farol
        return f'<div class="farol">{lights}</div>' #Retorna o farol com a cor

    def build_macro_farol(active_color): #Constrói o farol do projeto total
        items = ""
        for c, label in [('blue', 'Concluído'), ('green', 'On track'), ('red', 'Atrasado'), ('gray', 'Não iniciado')]:
            is_active = "active" if c == active_color else ""
            hex_ = COLOR_HEX.get(c, "#94a3b8")
            shadow = f"box-shadow: 0 0 8px {hex_};" if is_active else ""
            items += f'<div class="mf-item {is_active}"><div class="mf-light" style="background:{hex_}; {shadow}"></div><span>{label}</span></div>'
        return f'<div class="macro-farol">{items}</div>' #Retorna o andamento do farol do projeto total

    all_tbody_rows = ""

    for i, s in enumerate(stages):
        hex_ = COLOR_HEX[s["color"]] #Filtra na label "color" dentro de stages

        task_rows = ""
        for t in s["tasks"]: #Para item dentro do dict "tasks" de stages
            
            
            tc = COLOR_HEX.get(t["color"], "#94a3b8")
            indent = "padding-left:28px" if t["is_subtask"] else "" #Cria um espaço se for uma subtask
            prefix = "└ " if t["is_subtask"] else "" #Se for uma subtask, cria um "└" antes do nome dela
            sr = "subtask-row" if t["is_subtask"] else "" #Cria uma subtask row 
            task_rows += ( 
                f'<tr class="task-row {sr}">' #Formatação da "linha" para modelo de task
                f'<td style="{indent}">{dot(t["color"], 10)} {prefix}<code>{t["id"]}</code></td>' #Pega a cor da task
                f'<td style="color: var(--text)">{t["name"]}</td>' #Nome da task
                f'<td class="resp-col">{t["responsible"]}</td>' #Responsável
                f'<td class="date-col">{fmt_date(t["start"])}</td>' #Data inicial
                f'<td class="date-col">{fmt_date(t["end"])}</td>'#Data final
                f'<td><span class="status-badge" style="background:{tc}22;color:{tc};border:1px solid {tc}44">{t["status"]}</span></td>' #Status da task
                f'<td class="remarks-col">{t["remarks"]}</td>' #Comentários
                f'</tr>'
            )

        all_tbody_rows += (
            f'<tr class="stage-row" onclick="toggleDetail(\'detail-{i}\')">'
            f'<td class="stage-name"><span class="expand-icon">▶</span>'
            f'<strong style="color: var(--text)">{s["label"]}</strong></td>'
            f'<td class="weight-col">{int(s["weight"]*100)}%</td>'
            f'<td><div class="pbar-wrap"><div class="pbar-fill" style="width:{s["pct"]}%;background:{hex_}"></div></div>'
            f'<span class="pct-label">{s["pct"]}%</span></td>'
            f'<td class="status-col">{build_farol(s["color"])}'
            f'<span style="color:{hex_};font-weight:600;font-size:12px;margin-left:8px">{COLOR_LABEL[s["color"]]}</span></td>'
            f'</tr>'
            f'<tr id="detail-{i}" class="detail-section" style="display:none">'
            f'<td colspan="4" style="padding:0"><div class="detail-inner">'
            f'<table class="inner-table"><thead><tr>'
            f'<th>ID</th><th>Task / Subtask</th><th>Responsável</th>'
            f'<th>Início</th><th>Fim</th><th>Status</th><th>Comentários</th>'
            f'</tr></thead><tbody>{task_rows}</tbody></table>'
            f'</div></td></tr>' #Todas as rows de estágios juntos na macro
        )

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Farol PMO – Implementation Report</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

  /* ── TEMA ESCURO (Padrão) ── */
  :root {{
    --bg:         #0d1117;
    --surface:    #161b22;
    --border:     #30363d;
    --text:       #e6edf3;
    --muted:      #8b949e;
    --accent:     #58a6ff;
    --element-bg: #1c2128;
    --farol-bg:   #11151a;
    --track-bg:   #21262d;
  }}

  /* ── TEMA CLARO ── */
  body.light-theme {{
    --bg:         #f6f8fa;
    --surface:    #ffffff;
    --border:     #d0d7de;
    --text:       #1f2328;
    --muted:      #656d76;
    --accent:     #0969da;
    --element-bg: #f3f4f6;
    --farol-bg:   #e5e7eb;
    --track-bg:   #e5e7eb;
  }}

  body {{
    background: var(--bg);
    color: var(--text);
    font-family: 'IBM Plex Sans', sans-serif;
    font-size: 14px;
    min-height: 100vh;
    padding: 32px 24px;
    transition: background-color 0.3s, color 0.3s;
  }}

  /* ── Header Centralizado ── */
  .header {{
    display: flex; flex-direction: column; align-items: center; text-align: center;
    margin-bottom: 32px; padding-bottom: 32px; border-bottom: 1px solid var(--border);
  }}
  .header h1 {{ font-size: 24px; font-weight: 700; letter-spacing: -0.5px; color: var(--text); }}
  .header .sub {{ color: var(--muted); font-size: 12px; margin-top: 6px; font-family: 'IBM Plex Mono', monospace; }}

  .theme-btn {{
    padding: 6px 16px; border-radius: 6px; cursor: pointer;
    background: var(--surface); color: var(--text); border: 1px solid var(--border);
    font-family: inherit; font-size: 12px; font-weight: 600; margin-top: 16px;
    transition: all 0.2s;
  }}
  .theme-btn:hover {{ background: var(--element-bg); }}

  /* ── Meta Dados do Projeto (Datas e Responsável) ── */
  .project-info {{
    display: flex; flex-wrap: wrap; gap: 40px; justify-content: center;
    margin-top: 24px; padding: 16px 32px; background: var(--element-bg);
    border-radius: 12px; border: 1px solid var(--border);
  }}
  .info-item {{
    display: flex; flex-direction: column; align-items: center;
  }}
  .info-item strong {{
    color: var(--muted); font-size: 10px; text-transform: uppercase;
    letter-spacing: 0.08em; margin-bottom: 6px;
  }}
  .info-value {{
    font-family: 'IBM Plex Mono', monospace; font-size: 14px;
    color: var(--accent); font-weight: 600;
  }}

  .header-metrics {{
    display: flex; align-items: center; justify-content: center; gap: 48px;
    margin-top: 24px; background: var(--surface); padding: 20px 40px;
    border-radius: 16px; border: 1px solid var(--border);
  }}

  /* ── Overall gauge ── */
  .gauge-wrap {{ text-align: center; }}
  .gauge-ring {{
    width: 90px; height: 90px; border-radius: 50%;
    background: conic-gradient(var(--accent) calc({overall}% * 3.6deg), var(--track-bg) 0deg);
    display: flex; align-items: center; justify-content: center;
    position: relative; margin: 0 auto;
  }}
  .gauge-ring::after {{ content: ''; position: absolute; width: 70px; height: 70px; border-radius: 50%; background: var(--surface); }}
  .gauge-pct {{ position: relative; z-index: 1; font-size: 18px; font-weight: 700; color: var(--accent); font-family: 'IBM Plex Mono', monospace; }}
  .gauge-label {{ font-size: 11px; color: var(--muted); margin-top: 8px; text-transform: uppercase; letter-spacing: 0.08em; }}

  /* ── Macro Farol ── */
  .macro-farol-wrap {{ text-align: center; }}
  .macro-farol {{
    display: flex; background: var(--farol-bg); padding: 8px 16px; border-radius: 12px; gap: 16px;
    border: 1px solid var(--border); box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
  }}
  .mf-item {{ display: flex; align-items: center; gap: 8px; opacity: 0.3; transition: opacity 0.3s; }}
  .mf-item.active {{ opacity: 1; }}
  .mf-item span {{ font-size: 12px; font-weight: 600; color: var(--muted); }}
  .mf-item.active span {{ color: var(--text); }}
  .mf-light {{ width: 14px; height: 14px; border-radius: 50%; }}

  /* ── Main table ── */
  .card {{ background: var(--surface); border: 1px solid var(--border); border-radius: 10px; overflow: hidden; margin-bottom: 20px; }}
  table {{ width: 100%; border-collapse: collapse; }}
  thead tr {{ background: var(--element-bg); }}
  thead th {{ padding: 10px 16px; text-align: left; font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--muted); font-weight: 600; border-bottom: 1px solid var(--border); }}
  .stage-row {{ cursor: pointer; transition: background 0.15s; border-bottom: 1px solid var(--border); }}
  .stage-row:hover {{ background: var(--element-bg); }}
  .stage-row td {{ padding: 14px 16px; vertical-align: middle; }}

  .stage-name {{ width: 260px; }}
  .stage-name strong {{ font-size: 13px; }}
  .weight-col {{ width: 60px; text-align: center; font-family: 'IBM Plex Mono', monospace; color: var(--muted); }}
  .status-col {{ width: 220px; display: flex; align-items: center; gap: 6px; padding-top: 17px; }}

  .expand-icon {{ font-size: 9px; margin-right: 8px; color: var(--muted); transition: transform 0.2s; display: inline-block; }}
  .expanded .expand-icon {{ transform: rotate(90deg); }}

  /* ── Progress bar ── */
  .pbar-wrap {{ height: 6px; border-radius: 3px; background: var(--track-bg); overflow: hidden; margin-bottom: 4px; width: 200px; }}
  .pbar-fill {{ height: 100%; border-radius: 3px; transition: width 0.4s ease; }}
  .pct-label {{ font-size: 11px; font-family: 'IBM Plex Mono', monospace; color: var(--muted); }}

  /* ── Farol (Semáforo) Etapas ── */
  .dot {{ display: inline-block; border-radius: 50%; flex-shrink: 0; vertical-align: middle; }}
  .farol {{
    display: inline-flex; background: var(--farol-bg); padding: 4px; border-radius: 12px; gap: 4px;
    border: 1px solid var(--border); box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
  }}
  .farol-luz {{ width: 12px; height: 12px; border-radius: 50%; opacity: 0.15; transition: all 0.3s ease; }}

  /* ── Detail section ── */
  .detail-section td {{ padding: 0 !important; }}
  .detail-inner {{ border-top: 1px solid var(--border); background: var(--bg); padding: 0; }}
  .inner-table thead th {{ background: var(--element-bg); font-size: 10px; padding: 8px 14px; border-bottom: 1px solid var(--border); }}
  .task-row td {{ padding: 8px 14px; font-size: 12px; border-bottom: 1px solid var(--border); vertical-align: middle; }}
  .subtask-row {{ background: var(--surface); }}
  .task-row:last-child td {{ border-bottom: none; }}
  .resp-col {{ color: var(--muted); font-size: 11px; max-width: 180px; }}
  .date-col {{ font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--muted); white-space: nowrap; }}
  .remarks-col {{ max-width: 250px; font-size: 11px; color: var(--muted); word-wrap: break-word; }}
  .status-badge {{ display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 10px; font-weight: 600; white-space: nowrap; }}
  code {{ background: var(--element-bg); border-radius: 4px; padding: 1px 5px; font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--accent); border: 1px solid var(--border); }}

  /* ── Footer ── */
  .footer {{ margin-top: 32px; padding-top: 16px; border-top: 1px solid var(--border); font-size: 11px; color: var(--muted); font-family: 'IBM Plex Mono', monospace; display: flex; justify-content: space-between; }}
</style>
</head>
<body>

<div class="header">
  <h1> Farol PMO — Implementation Report</h1>
  <div class="sub">Arquivo: {filename} &nbsp;·&nbsp; Gerado em: {today_str}</div>
  
  <button onclick="toggleTheme()" class="theme-btn"> Alternar Modo Claro/Escuro</button>

  <div class="project-info">
    <div class="info-item"><strong>Responsável</strong><span class="info-value">{meta['responsible']}</span></div>
    <div class="info-item"><strong>Última Atualização</strong><span class="info-value">{meta['latest_update']}</span></div>
    <div class="info-item"><strong>Go-Live (Wave 1)</strong><span class="info-value">{meta['go_live']}</span></div>
    <div class="info-item"><strong>Project Closure</strong><span class="info-value">{meta['closure']}</span></div>
  </div>

  <div class="header-metrics">
    <div class="gauge-wrap">
      <div class="gauge-ring"><span class="gauge-pct">{overall}%</span></div>
      <div class="gauge-label">Conclusão Geral</div>
    </div>
    
    <div class="macro-farol-wrap">
    
      {build_macro_farol(overall_color)}
      <div class="gauge-label">Status do Projeto</div>
    </div>
  </div>
</div>

<div class="card">
  <table>
    <thead>
      <tr>
        <th>Etapa</th>
        <th>Peso</th>
        <th>Progresso</th>
        <th>Status</th>
      </tr>
    </thead>
    <tbody>
      {''.join([all_tbody_rows])}
    </tbody>
  </table>
</div>

<div class="footer">
  <span>Farol PMO Generator · {today_str}</span>
  <span>Pesos: Initiation 5% · Contrato 10% · Cadastros 10% · Go Live 10% · Design 30% · Systems 30% · Evaluation 5%</span>
</div>

<script>
function toggleDetail(id) {{
  const el = document.getElementById(id);
  const row = el.previousElementSibling;
  const visible = el.style.display !== 'none';
  el.style.display = visible ? 'none' : 'table-row';
  row.classList.toggle('expanded', !visible);
}}

function toggleTheme() {{
  document.body.classList.toggle('light-theme');
}}
</script>
</body>
</html>"""
    return html

def main():
    
    # Define o path
    filepath = "Implementation Plan and Timeline.xlsx"
    if not os.path.exists(filepath):
        print(f"[ERRO] Arquivo não encontrado: {filepath}")
        sys.exit(1)

    print(f"[INFO] Lendo arquivo: {filepath}")
    
    # 1. Carrega os dados das tarefas
    data = load_data(filepath) # Lê a Planilha e carrega em 'data'
    stages = build_stages(data) # Configura os estágios
    
    # 2. Busca os metadados do cabeçalho do Excel e as datas dos milestones
    latest_update, responsible = get_project_metadata(filepath) 
    go_live_date = get_milestone_date(data, "602") #Pega a data do GoLive na formatação XX/XX/XX
    closure_date = get_milestone_date(data, "706") #Pega a data de Closure Date na formatação XX/XX/XX
    
    # 3. Consolida num dicionário para passar ao HTML
    meta = {
        "latest_update": latest_update,
        "responsible": responsible,
        "go_live": go_live_date,
        "closure": closure_date
    }
    
    # 4. Renderiza enviando as tarefas e os metadados
    html = render_html(stages, filepath, meta) #Volta o .html feito 

    base = os.path.splitext(os.path.basename(filepath))[0]
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), base + "_farol.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"[OK]   Relatório gerado: {out_path}")
    print(f"\n{'─'*50}")
    print(f"  CONCLUSÃO GERAL: {overall_pct(stages)}%")
    print(f"{'─'*50}")
    for s in stages:
        bar = "█" * int(s["pct"] / 5) + "░" * (20 - int(s["pct"] / 5))
        print(f"  {s['label']:<35} [{bar}] {s['pct']:>5.1f}%  ({int(s['weight']*100)}%)  {COLOR_LABEL[s['color']]}")
    print(f"{'─'*50}\n")

if __name__ == "__main__":
    main()