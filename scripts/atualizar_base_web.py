# -*- coding: utf-8 -*-
import os
import sys
import json
import math
import requests
from datetime import datetime
import geopandas as gpd
import pandas as pd

# Forçar stdout UTF-8 para evitar erros de console no Windows
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

print("=" * 70)
print("  CLEALCO - ATUALIZADOR DA BASE WEB DE SEDES (KOBOTOOLBOX & GIS)")
print("=" * 70)

KOBO_TOKEN = "21419ca95379f5336eec89bd4decb2a420628d5b"
ASSET_UID = "a6xVWc5GTZEWFzSceegC4G"
KOBO_SERVER = "kf.kobotoolbox.org"

script_dir = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(script_dir).lower() == "scripts":
    root_dir = os.path.dirname(script_dir)
else:
    root_dir = script_dir

base_dir = root_dir
out_js = os.path.join(root_dir, "dados_mapa.js")
out_vistorias_json = os.path.join(root_dir, "vistorias_kobo.json")
out_xlsx_dir = os.path.join(root_dir, "relatorios e formularios")
os.makedirs(out_xlsx_dir, exist_ok=True)
out_xlsx = os.path.join(out_xlsx_dir, "Visitas_Monitoramento_Sedes_Kobo.xlsx")

def calcular_distancia_metros(lat1, lon1, lat2, lon2):
    try:
        R = 6371000
        dlat = math.radians(float(lat2) - float(lat1))
        dlon = math.radians(float(lon2) - float(lon1))
        a = math.sin(dlat / 2)**2 + math.cos(math.radians(float(lat1))) * math.cos(math.radians(float(lat2))) * math.sin(dlon / 2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
        return round(R * c, 1)
    except Exception:
        return ""

# -------------------------------------------------------------
# 1. CARREGAR SHAPEFILE DE SEDES (BASE DE REFERÊNCIA ESPACIAL)
# -------------------------------------------------------------
cand_sede = [
    os.path.join(root_dir, "banco de dados geograficos", "Sede.shp"),
    os.path.join(root_dir, "SHP", "Sede.shp"),
    os.path.join(root_dir, "Sede.shp")
]
sede_shp = next((p for p in cand_sede if os.path.exists(p)), None)

if not sede_shp:
    print(f"[ERRO] Arquivo de sedes nao encontrado: {cand_sede}")
    exit(1)

print(f"-> Lendo shapefile de sedes: {os.path.basename(sede_shp)}...")
pts = gpd.read_file(sede_shp, encoding="utf-8").to_crs(epsg=4326)

sedes_ref_map = {}
for idx, row in pts.iterrows():
    num_s = str(row.get("NUM_SEDE", idx + 1)).strip()
    lat_val = round(row.geometry.y, 6) if row.geometry else ""
    lon_val = round(row.geometry.x, 6) if row.geometry else ""
    sedes_ref_map[num_s] = {
        "num_sede": num_s,
        "nome_sede": str(row.get("NOME_SEDE", "")).strip(),
        "fa": str(row.get("FA", "")).strip(),
        "fazenda": str(row.get("FAZENDA", "")).strip(),
        "proprietario": str(row.get("proprietar", "")).strip(),
        "regiao": str(row.get("regiao", "")).strip(),
        "lat_sede": lat_val,
        "lon_sede": lon_val,
        "localizacao_sede": f"{lat_val}, {lon_val}" if lat_val and lon_val else "",
        "link_maps": f"https://www.google.com/maps?q={lat_val},{lon_val}" if lat_val and lon_val else ""
    }

print(f"-> {len(sedes_ref_map)} sedes de referência carregadas com localização exata.")

# -------------------------------------------------------------
# 2. OBTER VISTORIAS DO KOBOTOOLBOX
# -------------------------------------------------------------
url_kobo = f"https://{KOBO_SERVER}/api/v2/assets/{ASSET_UID}/data.json"
headers_kobo = {"Authorization": f"Token {KOBO_TOKEN}"}

excel_datas_manuais = {}
if os.path.exists(out_xlsx):
    try:
        df_ex = pd.read_excel(out_xlsx, sheet_name="Vistorias_Kobo")
        for _, rx in df_ex.iterrows():
            ns = str(rx.get("NUM_SEDE", "")).strip()
            dv = rx.get("DATA_VISTORIA")
            if pd.notna(dv):
                if hasattr(dv, "strftime"):
                    excel_datas_manuais[ns] = dv.strftime("%d/%m/%Y")
                else:
                    d_s = str(dv).split()[0]
                    if "-" in d_s:
                        p = d_s.split("-")
                        if len(p) == 3 and len(p[0]) == 4:
                            excel_datas_manuais[ns] = f"{p[2]}/{p[1]}/{p[0]}"
                        else:
                            excel_datas_manuais[ns] = d_s
                    else:
                        excel_datas_manuais[ns] = d_s
        print(f"-> [EXCEL] {len(excel_datas_manuais)} datas preservadas do Excel local.")
    except Exception as e:
        pass

todas_vistorias = []
visitas_registradas = {} # num_sede -> ultima_vistoria

print("-> Conectando a nuvem KoboToolbox (Monitoramento de Sedes)...")
kobo_online = False
try:
    resp = requests.get(url_kobo, headers=headers_kobo, timeout=20)
    if resp.status_code == 200:
        data_json = resp.json()
        kobo_results = data_json.get("results", [])
        print(f"-> [KOBOTOOLBOX] {len(kobo_results)} vistoria(s) baixada(s) com sucesso!")
        kobo_online = True

        for r in kobo_results:
            num_s = str(r.get("num_sede", "")).strip()
            s_ref = sedes_ref_map.get(num_s, {})

            nome_s = str(r.get("nome_sede", "")).strip() or s_ref.get("nome_sede", f"Sede {num_s}")
            fa = str(r.get("fa", "")).strip() or s_ref.get("fa", "")
            fazenda = str(r.get("fazenda", "")).strip() or s_ref.get("fazenda", "")
            proprietario = str(r.get("proprietario", "")).strip() or s_ref.get("proprietario", "")
            regiao = str(r.get("regiao", "")).strip() or s_ref.get("regiao", "")
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

            # GPS onde o ponto foi lançado
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

            # Localização oficial da sede fixa
            lat_s = s_ref.get("lat_sede", "")
            lon_s = s_ref.get("lon_sede", "")
            loc_s = s_ref.get("localizacao_sede", "")
            link_m = s_ref.get("link_maps", "")

            dist_m = ""
            if lat_gps is not None and lon_gps is not None and lat_s and lon_s:
                dist_m = calcular_distancia_metros(lat_gps, lon_gps, lat_s, lon_s)

            # Verificar se o usuário editou a data manualmente no Excel
            if num_s in excel_datas_manuais:
                data_v = excel_datas_manuais[num_s]

            v_item = {
                "num_sede": num_s,
                "nome_sede": nome_s,
                "fa": fa,
                "fazenda": fazenda,
                "proprietario": proprietario,
                "regiao": regiao,
                "data_vistoria": data_v,
                "hora_vistoria": hora_v,
                "matricula": matricula,
                "pressao": pressao,
                "bandeiras": bandeiras,
                "larva_pupa": larva_pupa,
                "inseticida": inseticida,
                "possiveis_causadores": causador,
                "observacoes": obs,
                "lat": lat_gps,
                "lon": lon_gps,
                "latitude_gps_lancamento": lat_gps if lat_gps is not None else "",
                "longitude_gps_lancamento": lon_gps if lon_gps is not None else "",
                "lat_sede": lat_s,
                "lon_sede": lon_s,
                "localizacao_sede": loc_s,
                "link_google_maps_sede": link_m,
                "distancia_sede_metros": dist_m,
                "origem": "KoboToolbox",
                "submetido_em": r.get("_submission_time", ""),
                "kobo_id": r.get("_id")
            }
            todas_vistorias.append(v_item)
    else:
        print(f"-> [AVISO] Kobo retornou status {resp.status_code}")
except Exception as e:
    print(f"-> [AVISO] Nao foi possivel conectar ao Kobo online: {e}")

# Se nao conseguiu baixar online, tenta ler vistorias_kobo.json anterior
if not kobo_online and os.path.exists(out_vistorias_json):
    try:
        with open(out_vistorias_json, "r", encoding="utf-8") as f:
            todas_vistorias = json.load(f)
        print(f"-> [OFFLINE] {len(todas_vistorias)} vistoria(s) carregadas do cache local.")
    except Exception as e:
        print(f"-> [AVISO] Falha ao ler cache local de vistorias: {e}")

# Salvar arquivo JSON estático das vistorias
if todas_vistorias:
    try:
        with open(out_vistorias_json, "w", encoding="utf-8") as f:
            json.dump(todas_vistorias, f, ensure_ascii=False, indent=2)
        print(f"-> [ARQUIVO] Salvo {out_vistorias_json} ({len(todas_vistorias)} vistorias)")
    except Exception as e:
        print(f"-> [AVISO] Falha ao gravar vistorias_kobo.json: {e}")

# Ordenar vistorias por data decrescente e mapear a mais recente de cada sede
def parse_data_sort(v):
    d_str = v.get("data_vistoria", "")
    sub_str = v.get("submetido_em", "")
    try:
        if "/" in d_str:
            d_parts = d_str.split("/")
            if len(d_parts) == 3:
                return datetime(int(d_parts[2]), int(d_parts[1]), int(d_parts[0]))
        elif "-" in d_str:
            d_parts = d_str.split("-")
            if len(d_parts) == 3:
                if len(d_parts[0]) == 4:
                    return datetime(int(d_parts[0]), int(d_parts[1]), int(d_parts[2]))
                else:
                    return datetime(int(d_parts[2]), int(d_parts[1]), int(d_parts[0]))
    except Exception:
        pass
    if sub_str:
        try:
            return datetime.fromisoformat(sub_str.replace("Z", "+00:00"))
        except Exception:
            pass
    return datetime(1900, 1, 1)

# Agrupar por num_sede e selecionar a última
for v in todas_vistorias:
    s_id = str(v.get("num_sede", "")).strip()
    if not s_id:
        continue
    if s_id not in visitas_registradas:
        visitas_registradas[s_id] = v
    else:
        if parse_data_sort(v) >= parse_data_sort(visitas_registradas[s_id]):
            visitas_registradas[s_id] = v

print(f"-> Total de sedes distintas vistoriadas: {len(visitas_registradas)}")

# -------------------------------------------------------------
# 3. PROCESSAR PONTOS PARA DADOS_MAPA.JS E EXCEL
# -------------------------------------------------------------
pts_features = []
sedes_export_excel = []
hoje = datetime.now().date()

for idx, row in pts.iterrows():
    num_sede = str(row.get("NUM_SEDE", idx + 1)).strip()
    nome_sede = str(row.get("NOME_SEDE", "")).strip()
    fa = str(row.get("FA", "")).strip()
    fazenda = str(row.get("FAZENDA", "")).strip()
    proprietario = str(row.get("proprietar", "")).strip()
    regiao = str(row.get("regiao", "")).strip()

    if fa in ["nan", "None", "-"]:
        fa = ""
    if fazenda in ["nan", "None"]:
        fazenda = ""

    if fa and fazenda:
        fa_completo = f"{fa} - {fazenda}"
    elif fazenda:
        fa_completo = fazenda
    elif fa:
        fa_completo = f"FA {fa}"
    else:
        fa_completo = "Sem Informacao"

    v_info = visitas_registradas.get(num_sede, {})
    is_visitado = bool(v_info)

    dias_vistoria = None
    status_tematico = "Vermelho (Sem Vistoria)"
    if is_visitado:
        dt_v = parse_data_sort(v_info)
        if dt_v.year > 1900:
            diff_d = max(0, (hoje - dt_v.date()).days)
            dias_vistoria = diff_d
            if diff_d <= 7:
                status_tematico = "Verde (<= 7 dias)"
            elif diff_d <= 15:
                status_tematico = "Amarelo (8 a 15 dias)"
            else:
                status_tematico = "Vermelho (>= 16 dias)"

    try:
        id_int = int(float(num_sede))
    except Exception:
        id_int = idx + 1

    lat_pt = round(row.geometry.y, 6) if row.geometry else None
    lon_pt = round(row.geometry.x, 6) if row.geometry else None
    loc_pt = f"{lat_pt}, {lon_pt}" if lat_pt and lon_pt else ""
    link_m = f"https://www.google.com/maps?q={lat_pt},{lon_pt}" if lat_pt and lon_pt else ""

    props = {
        "id": id_int,
        "num_sede": num_sede,
        "nome_sede": nome_sede if nome_sede else f"Sede {num_sede}",
        "fa": fa,
        "fazenda": fazenda,
        "fa_completo": fa_completo,
        "proprietario": proprietario if proprietario != "nan" else "",
        "regiao": regiao if regiao != "nan" else "",
        "lat": lat_pt,
        "lon": lon_pt,
        "lat_sede": lat_pt,
        "lon_sede": lon_pt,
        "localizacao_sede": loc_pt,
        "link_google_maps_sede": link_m,
        "visitado": is_visitado,
        "dias_vistoria": dias_vistoria,
        "status_tematico": status_tematico,
        "data_vistoria": v_info.get("data_vistoria", ""),
        "hora_vistoria": v_info.get("hora_vistoria", ""),
        "matricula": v_info.get("matricula", ""),
        "pressao": v_info.get("pressao", ""),
        "bandeiras": v_info.get("bandeiras", ""),
        "larva_pupa": v_info.get("larva_pupa", ""),
        "inseticida": v_info.get("inseticida", ""),
        "possiveis_causadores": v_info.get("possiveis_causadores", ""),
        "observacoes": v_info.get("observacoes", "")
    }

    pts_features.append({
        "type": "Feature",
        "geometry": {
            "type": "Point",
            "coordinates": [lon_pt, lat_pt]
        },
        "properties": props
    })

    sedes_export_excel.append({
        "NUM_SEDE": num_sede,
        "NOME_SEDE": props["nome_sede"],
        "FA": fa,
        "FAZENDA": fazenda,
        "PROPRIETARIO": props["proprietario"],
        "REGIAO": props["regiao"],
        "LATITUDE_SEDE": lat_pt,
        "LONGITUDE_SEDE": lon_pt,
        "LOCALIZACAO_SEDE": loc_pt,
        "LINK_GOOGLE_MAPS_SEDE": link_m,
        "STATUS_VISTORIA": "Vistoriada" if is_visitado else "Pendente",
        "STATUS_TEMATICO": status_tematico,
        "DIAS_DESDE_VISTORIA": dias_vistoria if dias_vistoria is not None else "",
        "ULTIMA_DATA": props["data_vistoria"],
        "ULTIMA_HORA": props["hora_vistoria"],
        "AVALIADOR": props["matricula"],
        "PRESSAO": props["pressao"],
        "BANDEIRAS": props["bandeiras"],
        "LARVA_PUPA": props["larva_pupa"],
        "INSETICIDA": props["inseticida"],
        "CAUSADORES": props["possiveis_causadores"],
        "OBSERVACOES": props["observacoes"]
    })

print(f"   [OK] {len(pts_features)} sedes processadas.")

# -------------------------------------------------------------
# 4. CARREGAR E PROCESSAR SHAPEFILE DE AREAS (POLIGONOS)
# -------------------------------------------------------------
cand_areas = [
    os.path.join(root_dir, "banco de dados geograficos", "AREAS.shp"),
    os.path.join(root_dir, "SHP", "AREAS.shp"),
    os.path.join(root_dir, "AREAS.shp")
]
areas_shp = next((p for p in cand_areas if os.path.exists(p)), None)

areas_features = []
if areas_shp and os.path.exists(areas_shp):
    print(f"-> Lendo areas: {os.path.basename(areas_shp)}...")
    areas = gpd.read_file(areas_shp, encoding="utf-8").to_crs(epsg=4326)
    areas["geometry"] = areas.geometry.simplify(0.0001, preserve_topology=True)
    for idx, row in areas.iterrows():
        if row.geometry is not None and not row.geometry.is_empty:
            geom_json = row.geometry.__geo_interface__
            faz = str(row.get("FAZENDA", "")).strip()
            fa_cod = str(row.get("FA", "")).strip()
            num_s = str(row.get("NUM_SEDE", "")).strip()
            areas_features.append({
                "type": "Feature",
                "geometry": geom_json,
                "properties": {
                    "fa": fa_cod if fa_cod != "nan" else "",
                    "fazenda": faz if faz != "nan" else "",
                    "num_sede": num_s if num_s != "nan" else ""
                }
            })
    print(f"   [OK] {len(areas_features)} poligonos de fazendas/areas processados.")

# -------------------------------------------------------------
# 5. GRAVAR DADOS_MAPA.JS COM VISTORIAS E COORDENADAS DA SEDE
# -------------------------------------------------------------
print("-> Gravando dados_mapa.js...")
with open(out_js, "w", encoding="utf-8") as f:
    fc_pontos = json.dumps({"type": "FeatureCollection", "features": pts_features}, ensure_ascii=False)
    fc_perims = json.dumps({"type": "FeatureCollection", "features": areas_features}, ensure_ascii=False)
    fc_vistorias = json.dumps(todas_vistorias, ensure_ascii=False)
    
    f.write("// BASE DE DADOS COMPILADA DO MONITORAMENTO DE SEDES CLEALCO\n")
    f.write(f"// Atualizado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n")
    f.write("const DADOS_VISTORIAS_KOBO = " + fc_vistorias + ";\n\n")
    f.write("const DADOS_SEDES = " + fc_pontos + ";\n\n")
    f.write("const DADOS_AREAS = " + fc_perims + ";\n\n")
    f.write("const DADOS_PONTOS = DADOS_SEDES;\n")
    f.write("const DADOS_PERIMETROS = DADOS_AREAS;\n")

tam_mb = os.path.getsize(out_js) / (1024 * 1024)
print(f"   [OK] dados_mapa.js atualizado com sucesso ({tam_mb:.2f} MB)!")

# -------------------------------------------------------------
# 6. ATUALIZAR PLANILHA EXCEL CONSOLIDADA
# -------------------------------------------------------------
try:
    # Preparar df_visitas com colunas bem organizadas
    rows_excel_visitas = []
    for v in todas_vistorias:
        rows_excel_visitas.append({
            "ORIGEM": v.get("origem", "KoboToolbox"),
            "DATA_VISTORIA": v.get("data_vistoria", ""),
            "HORA": v.get("hora_vistoria", ""),
            "NUM_SEDE": v.get("num_sede", ""),
            "NOME_SEDE": v.get("nome_sede", ""),
            "FA": v.get("fa", ""),
            "FAZENDA": v.get("fazenda", ""),
            "PROPRIETARIO": v.get("proprietario", ""),
            "REGIAO": v.get("regiao", ""),
            "LATITUDE_SEDE": v.get("lat_sede", ""),
            "LONGITUDE_SEDE": v.get("lon_sede", ""),
            "LOCALIZACAO_SEDE": v.get("localizacao_sede", ""),
            "LINK_GOOGLE_MAPS_SEDE": v.get("link_google_maps_sede", ""),
            "MATRICULA_AVALIADOR": v.get("matricula", ""),
            "PRESSAO": v.get("pressao", ""),
            "NUM_BANDEIRAS": v.get("bandeiras", ""),
            "LARVA_PUPA": v.get("larva_pupa", ""),
            "INSETICIDA": v.get("inseticida", ""),
            "POSSIVEIS_CAUSADORES": v.get("possiveis_causadores", ""),
            "OBSERVACOES": v.get("observacoes", ""),
            "LATITUDE_GPS_LANCAMENTO": v.get("latitude_gps_lancamento", v.get("lat", "")),
            "LONGITUDE_GPS_LANCAMENTO": v.get("longitude_gps_lancamento", v.get("lon", "")),
            "DISTANCIA_SEDE_METROS": v.get("distancia_sede_metros", ""),
            "DATA_ENVIO": v.get("submetido_em", "")
        })

    df_v = pd.DataFrame(rows_excel_visitas)
    df_s = pd.DataFrame(sedes_export_excel)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        if not df_v.empty:
            df_v.to_excel(writer, index=False, sheet_name="Vistorias_Kobo")
        if not df_s.empty:
            df_s.to_excel(writer, index=False, sheet_name="Status_Sedes")
    print(f"   [OK] Planilha Excel sincronizada com localização das sedes: {out_xlsx}")
except Exception as e:
    print(f"   [AVISO] Nao foi possivel atualizar Excel: {e}")

print("=" * 70)
print(f"  [CONCLUIDO] {len(visitas_registradas)} sedes com vistoria ativa no mapa!")
print("=" * 70)
