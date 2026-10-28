# -*- coding: utf-8 -*-
import os
import sys
import math
import requests
import pandas as pd
from datetime import datetime

# Forçar UTF-8 no terminal Windows
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

print("=" * 70)
print("  CLEALCO - EXPORTADOR DE VISTORIAS DAS SEDES (KOBOTOOLBOX -> EXCEL)")
print("=" * 70)

KOBO_TOKEN = "21419ca95379f5336eec89bd4decb2a420628d5b"
ASSET_UID = "a6xVWc5GTZEWFzSceegC4G"
KOBO_SERVER = "kf.kobotoolbox.org"

script_dir = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(script_dir).lower() == "scripts":
    root_dir = os.path.dirname(script_dir)
else:
    root_dir = script_dir

out_xlsx_dir = os.path.join(root_dir, "relatorios e formularios")
os.makedirs(out_xlsx_dir, exist_ok=True)
out_xlsx = os.path.join(out_xlsx_dir, "Visitas_Monitoramento_Sedes_Kobo.xlsx")

def calcular_distancia_metros(lat1, lon1, lat2, lon2):
    try:
        R = 6371000  # Raio da Terra em metros
        dlat = math.radians(float(lat2) - float(lat1))
        dlon = math.radians(float(lon2) - float(lon1))
        a = math.sin(dlat / 2)**2 + math.cos(math.radians(float(lat1))) * math.cos(math.radians(float(lat2))) * math.sin(dlon / 2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        return round(R * c, 1)
    except Exception:
        return ""

# 1. Carregar Sedes para base de referência e coordenadas oficiais
cand_sede = [
    os.path.join(root_dir, "banco de dados geograficos", "Sede.shp"),
    os.path.join(root_dir, "SHP", "Sede.shp"),
    os.path.join(root_dir, "Sede.shp")
]
sede_shp = next((p for p in cand_sede if os.path.exists(p)), None)
sedes_dict = {}

if sede_shp and os.path.exists(sede_shp):
    try:
        import geopandas as gpd
        gdf_s = gpd.read_file(sede_shp, encoding="utf-8").to_crs(epsg=4326)
        for idx, row in gdf_s.iterrows():
            n_sede = str(row.get("NUM_SEDE", idx + 1)).strip()
            lat_pt = round(row.geometry.y, 6) if row.geometry else ""
            lon_pt = round(row.geometry.x, 6) if row.geometry else ""
            loc_str = f"{lat_pt}, {lon_pt}" if lat_pt and lon_pt else ""
            link_maps = f"https://www.google.com/maps?q={lat_pt},{lon_pt}" if lat_pt and lon_pt else ""

            sedes_dict[n_sede] = {
                "NUM_SEDE": n_sede,
                "NOME_SEDE": str(row.get("NOME_SEDE", "")).strip(),
                "FA": str(row.get("FA", "")).strip(),
                "FAZENDA": str(row.get("FAZENDA", "")).strip(),
                "PROPRIETARIO": str(row.get("proprietar", "")).strip(),
                "REGIAO": str(row.get("regiao", "")).strip(),
                "LATITUDE_SEDE": lat_pt,
                "LONGITUDE_SEDE": lon_pt,
                "LOCALIZACAO_SEDE": loc_str,
                "LINK_GOOGLE_MAPS": link_maps,
                "STATUS_MONITORAMENTO": "Pendente",
                "STATUS_TEMATICO": "Vermelho (Sem Vistoria)",
                "DIAS_DESDE_ULTIMA_VISTORIA": "",
                "TOTAL_VISTORIAS": 0,
                "ULTIMA_DATA": "",
                "ULTIMA_HORA": "",
                "ULTIMA_PRESSAO": "",
                "ULTIMAS_BANDEIRAS": "",
                "ULTIMO_LARVA_PUPA": "",
                "ULTIMO_INSETICIDA": "",
                "ULTIMOS_CAUSADORES": "",
                "ULTIMO_AVALIADOR": ""
            }
        print(f"-> Base geográfica de referência: {len(sedes_dict)} sedes carregadas com coordenadas.")
    except Exception as e:
        print(f"[AVISO] Falha ao ler shapefile de sedes: {e}")

# 2. Consultar KoboToolbox API
url = f"https://{KOBO_SERVER}/api/v2/assets/{ASSET_UID}/data.json"
headers = {"Authorization": f"Token {KOBO_TOKEN}"}

print("-> Conectando ao KoboToolbox na nuvem...")
results = []
try:
    resp = requests.get(url, headers=headers, timeout=20)
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
    fazenda = str(r.get("fazenda", "")).strip()
    proprietario = str(r.get("proprietario", "")).strip()
    regiao = str(r.get("regiao", "")).strip()
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

    # Coordenadas do GPS de onde o avaliador lançou a vistoria
    lat_gps, lon_gps = None, None
    gps_str = r.get("localizacao_gps")
    if gps_str:
        parts = str(gps_str).strip().split()
        if len(parts) >= 2:
            try:
                lat_gps = float(parts[0])
                lon_gps = float(parts[1])
            except Exception:
                pass
    if lat_gps is None and r.get("latitude"):
        try:
            lat_gps = float(r.get("latitude"))
        except Exception:
            pass
    if lon_gps is None and r.get("longitude"):
        try:
            lon_gps = float(r.get("longitude"))
        except Exception:
            pass

    # Coordenadas oficiais da sede fixa
    sede_ref = sedes_dict.get(num_s, {})
    lat_sede = sede_ref.get("LATITUDE_SEDE", "")
    lon_sede = sede_ref.get("LONGITUDE_SEDE", "")
    loc_sede = sede_ref.get("LOCALIZACAO_SEDE", "")
    link_maps = sede_ref.get("LINK_GOOGLE_MAPS", "")

    # Se a sede não estava no shapefile mas tem coords na submissão
    if not lat_sede and r.get("lat_sede"):
        lat_sede = r.get("lat_sede")
        lon_sede = r.get("lon_sede")
        loc_sede = f"{lat_sede}, {lon_sede}"
        link_maps = f"https://www.google.com/maps?q={lat_sede},{lon_sede}"

    # Calcular distância em metros entre o GPS de lançamento e a sede
    dist_metros = ""
    if lat_gps is not None and lon_gps is not None and lat_sede and lon_sede:
        dist_metros = calcular_distancia_metros(lat_gps, lon_gps, lat_sede, lon_sede)

    rows.append({
        "ORIGEM": "KoboToolbox (Nuvem)",
        "DATA_VISTORIA": data_v,
        "HORA": hora_v,
        "NUM_SEDE": num_s,
        "NOME_SEDE": nome_s or sede_ref.get("NOME_SEDE", f"Sede {num_s}"),
        "FA": fa or sede_ref.get("FA", ""),
        "FAZENDA": fazenda or sede_ref.get("FAZENDA", ""),
        "PROPRIETARIO": proprietario or sede_ref.get("PROPRIETARIO", ""),
        "REGIAO": regiao or sede_ref.get("REGIAO", ""),
        "LATITUDE_SEDE": lat_sede,
        "LONGITUDE_SEDE": lon_sede,
        "LOCALIZACAO_SEDE": loc_sede,
        "LINK_GOOGLE_MAPS_SEDE": link_maps,
        "MATRICULA_AVALIADOR": matricula,
        "PRESSAO": pressao,
        "NUM_BANDEIRAS": bandeiras,
        "LARVA_PUPA": larva_pupa,
        "INSETICIDA": inseticida,
        "POSSIVEIS_CAUSADORES": causador,
        "OBSERVACOES": obs,
        "LATITUDE_GPS_LANCAMENTO": lat_gps if lat_gps is not None else "",
        "LONGITUDE_GPS_LANCAMENTO": lon_gps if lon_gps is not None else "",
        "DISTANCIA_SEDE_METROS": dist_metros,
        "DATA_ENVIO": r.get("_submission_time", "")
    })

    if num_s in sedes_dict:
        sedes_dict[num_s]["STATUS_MONITORAMENTO"] = "Visitado"
        sedes_dict[num_s]["TOTAL_VISTORIAS"] += 1
        sedes_dict[num_s]["ULTIMA_DATA"] = data_v
        sedes_dict[num_s]["ULTIMA_HORA"] = hora_v
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
            "FAZENDA", "PROPRIETARIO", "REGIAO", "LATITUDE_SEDE", "LONGITUDE_SEDE",
            "LOCALIZACAO_SEDE", "LINK_GOOGLE_MAPS_SEDE", "MATRICULA_AVALIADOR",
            "PRESSAO", "NUM_BANDEIRAS", "LARVA_PUPA", "INSETICIDA", "POSSIVEIS_CAUSADORES",
            "OBSERVACOES", "LATITUDE_GPS_LANCAMENTO", "LONGITUDE_GPS_LANCAMENTO",
            "DISTANCIA_SEDE_METROS", "DATA_ENVIO"
        ])
        df_empty.to_excel(writer, index=False, sheet_name="Vistorias_Kobo")

    if not df_sedes.empty:
        df_sedes.to_excel(writer, index=False, sheet_name="Status_Sedes")

print(f"[SUCESSO] Planilha Excel exportada com localização da sede: {out_xlsx}")
print("=" * 70)
