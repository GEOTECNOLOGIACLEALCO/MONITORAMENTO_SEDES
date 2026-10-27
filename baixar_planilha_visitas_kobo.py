# -*- coding: utf-8 -*-
import os
import requests
import pandas as pd
from datetime import datetime

print("=" * 65)
print("  CLEALCO - EXPORTADOR DE VISTORIAS DAS SEDES (KOBOTOOLBOX -> EXCEL)")
print("=" * 65)

KOBO_TOKEN = "21419ca95379f5336eec89bd4decb2a420628d5b"
ASSET_UID = "a6xVWc5GTZEWFzSceegC4G"
KOBO_SERVER = "kf.kobotoolbox.org"

base_dir = os.path.dirname(os.path.abspath(__file__))
out_xlsx = os.path.join(base_dir, "Visitas_Monitoramento_Sedes_Kobo.xlsx")

# 1. Carregar Sedes para base de referencia
sede_shp = os.path.join(base_dir, "Sede.shp")
sedes_dict = {}
if os.path.exists(sede_shp):
    try:
        import geopandas as gpd
        gdf_s = gpd.read_file(sede_shp, encoding="utf-8")
        for idx, row in gdf_s.iterrows():
            n_sede = str(row.get("NUM_SEDE", idx + 1)).strip()
            sedes_dict[n_sede] = {
                "NUM_SEDE": n_sede,
                "NOME_SEDE": str(row.get("NOME_SEDE", "")).strip(),
                "FA": str(row.get("FA", "")).strip(),
                "FAZENDA": str(row.get("FAZENDA", "")).strip(),
                "PROPRIETARIO": str(row.get("proprietar", "")).strip(),
                "REGIAO": str(row.get("regiao", "")).strip(),
                "LATITUDE": round(row.geometry.y, 6) if row.geometry else "",
                "LONGITUDE": round(row.geometry.x, 6) if row.geometry else "",
                "STATUS_MONITORAMENTO": "Pendente",
                "STATUS_TEMATICO": "Vermelho (Sem Vistoria)",
                "DIAS_DESDE_ULTIMA_VISTORIA": "",
                "TOTAL_VISTORIAS": 0,
                "ULTIMA_DATA": "",
                "ULTIMA_PRESSAO": "",
                "ULTIMAS_BANDEIRAS": "",
                "ULTIMO_LARVA_PUPA": "",
                "ULTIMO_INSETICIDA": "",
                "ULTIMOS_CAUSADORES": "",
                "ULTIMO_AVALIADOR": ""
            }
    except Exception as e:
        print(f"[AVISO] Falha ao ler shapefile de sedes: {e}")

# 2. Consultar KoboToolbox API
url = f"https://{KOBO_SERVER}/api/v2/assets/{ASSET_UID}/data.json"
headers = {"Authorization": f"Token {KOBO_TOKEN}"}

print("-> Conectando ao KoboToolbox na nuvem...")
results = []
try:
    resp = requests.get(url, headers=headers, timeout=15)
    if resp.status_code == 200:
        results = resp.json().get("results", [])
        print(f"-> Total de vistorias recebidas da nuvem Kobo: {len(results)}")
    else:
        print(f"[AVISO] Resposta do Kobo: {resp.status_code}")
except Exception as e:
    print(f"[AVISO] Nao foi possivel conectar ao Kobo online: {e}")

rows = []
for r in results:
    num_s = str(r.get("num_sede", "")).strip()
    nome_s = str(r.get("nome_sede", "")).strip()
    fa = str(r.get("fa", "")).strip()
    data_v = str(r.get("data_vistoria", "")).strip()
    hora_v = str(r.get("hora_vistoria", "")).strip()
    matricula = str(r.get("matricula", "")).strip()
    pressao = str(r.get("pressao", "")).strip()
    bandeiras = r.get("bandeiras", "")
    larva_pupa = str(r.get("larva_pupa", "")).strip()
    inseticida = str(r.get("inseticida", "")).strip()
    causador = str(r.get("possiveis_causadores", "")).strip()
    if causador == "Outro" and r.get("causador_outro"):
        causador = f"Outro: {r.get('causador_outro')}"
    obs = str(r.get("observacoes", "")).strip()

    lat, lon = None, None
    gps_str = r.get("localizacao_gps")
    if gps_str:
        parts = str(gps_str).strip().split()
        if len(parts) >= 2:
            try:
                lat = float(parts[0])
                lon = float(parts[1])
            except Exception:
                pass

    rows.append({
        "ORIGEM": "KoboToolbox (Nuvem)",
        "DATA_VISTORIA": data_v,
        "HORA": hora_v,
        "NUM_SEDE": num_s,
        "NOME_SEDE": nome_s,
        "FA": fa,
        "MATRICULA_AVALIADOR": matricula,
        "PRESSAO": pressao,
        "NUM_BANDEIRAS": bandeiras,
        "LARVA_PUPA": larva_pupa,
        "INSETICIDA": inseticida,
        "POSSIVEIS_CAUSADORES": causador,
        "OBSERVACOES": obs,
        "LATITUDE": lat or r.get("latitude", ""),
        "LONGITUDE": lon or r.get("longitude", ""),
        "DATA_ENVIO": r.get("_submission_time", "")
    })

    if num_s in sedes_dict:
        sedes_dict[num_s]["STATUS_MONITORAMENTO"] = "Visitado"
        sedes_dict[num_s]["TOTAL_VISTORIAS"] += 1
        sedes_dict[num_s]["ULTIMA_DATA"] = f"{data_v} {hora_v}".strip()
        sedes_dict[num_s]["ULTIMA_PRESSAO"] = pressao
        sedes_dict[num_s]["ULTIMAS_BANDEIRAS"] = bandeiras
        sedes_dict[num_s]["ULTIMO_LARVA_PUPA"] = larva_pupa
        sedes_dict[num_s]["ULTIMO_INSETICIDA"] = inseticida
        sedes_dict[num_s]["ULTIMOS_CAUSADORES"] = causador
        sedes_dict[num_s]["ULTIMO_AVALIADOR"] = matricula

        try:
            d_str = data_v.split()[0]
            if "/" in d_str:
                dt_v = datetime.strptime(d_str, "%d/%m/%Y")
            else:
                dt_v = datetime.strptime(d_str, "%Y-%m-%d")
            diff_d = max(0, (datetime.now().date() - dt_v.date()).days)
            sedes_dict[num_s]["DIAS_DESDE_ULTIMA_VISTORIA"] = diff_d
            if diff_d <= 7:
                sedes_dict[num_s]["STATUS_TEMATICO"] = "Verde (<= 7 dias)"
            elif diff_d <= 15:
                sedes_dict[num_s]["STATUS_TEMATICO"] = "Amarelo (8 a 15 dias)"
            else:
                sedes_dict[num_s]["STATUS_TEMATICO"] = "Vermelho (>= 16 dias)"
        except Exception:
            pass

df_visitas = pd.DataFrame(rows)
df_sedes = pd.DataFrame(list(sedes_dict.values()))

with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
    if not df_visitas.empty:
        df_visitas.to_excel(writer, index=False, sheet_name="Vistorias_Kobo")
    else:
        df_empty = pd.DataFrame(columns=[
            "ORIGEM", "DATA_VISTORIA", "HORA", "NUM_SEDE", "NOME_SEDE", "FA",
            "MATRICULA_AVALIADOR", "PRESSAO", "NUM_BANDEIRAS", "LARVA_PUPA",
            "INSETICIDA", "POSSIVEIS_CAUSADORES", "OBSERVACOES", "LATITUDE", "LONGITUDE"
        ])
        df_empty.to_excel(writer, index=False, sheet_name="Vistorias_Kobo")

    if not df_sedes.empty:
        df_sedes.to_excel(writer, index=False, sheet_name="Status_Sedes")

print(f"[SUCESSO] Planilha Excel exportada: {out_xlsx}")
print("=" * 65)
