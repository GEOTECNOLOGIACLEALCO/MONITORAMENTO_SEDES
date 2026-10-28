# -*- coding: utf-8 -*-
import os
import sys
import json
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

print("=" * 65)
print("  CLEALCO - ATUALIZADOR DA BASE WEB DE SEDES (KOBOTOOLBOX & GIS)")
print("=" * 65)

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

# -------------------------------------------------------------
# 1. OBTER VISTORIAS DO KOBOTOOLBOX
# -------------------------------------------------------------
url_kobo = f"https://{KOBO_SERVER}/api/v2/assets/{ASSET_UID}/data.json"
headers_kobo = {"Authorization": f"Token {KOBO_TOKEN}"}

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
            if lat is None and r.get("latitude"):
                try:
                    lat = float(r.get("latitude"))
                except Exception:
                    pass
            if lon is None and r.get("longitude"):
                try:
                    lon = float(r.get("longitude"))
                except Exception:
                    pass

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
                "lat": lat,
                "lon": lon,
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

# Salvar arquivo JSON estatico das vistorias
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
    h_str = v.get("hora_vistoria", "")
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
        # Se a atual for mais nova que a guardada
        if parse_data_sort(v) >= parse_data_sort(visitas_registradas[s_id]):
            visitas_registradas[s_id] = v

print(f"-> Total de sedes distintas vistoriadas: {len(visitas_registradas)}")

# -------------------------------------------------------------
# 2. CARREGAR E PROCESSAR SHAPEFILE DE SEDES (PONTOS)
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

    # Calcular dias e status temático
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

    props = {
        "id": id_int,
        "num_sede": num_sede,
        "nome_sede": nome_sede if nome_sede else f"Sede {num_sede}",
        "fa": fa,
        "fazenda": fazenda,
        "fa_completo": fa_completo,
        "proprietario": proprietario if proprietario != "nan" else "",
        "regiao": regiao if regiao != "nan" else "",
        "lat": round(row.geometry.y, 6) if row.geometry else None,
        "lon": round(row.geometry.x, 6) if row.geometry else None,
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
            "coordinates": [round(row.geometry.x, 6), round(row.geometry.y, 6)]
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
        "OBSERVACOES": props["observacoes"],
        "LATITUDE": props["lat"],
        "LONGITUDE": props["lon"]
    })

print(f"   [OK] {len(pts_features)} sedes processadas.")

# -------------------------------------------------------------
# 3. CARREGAR E PROCESSAR SHAPEFILE DE AREAS (POLIGONOS)
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
# 4. GRAVAR DADOS_MAPA.JS COM VISTORIAS EMBUTIDAS
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
# 5. ATUALIZAR PLANILHA EXCEL CONSOLIDADA
# -------------------------------------------------------------
try:
    df_v = pd.DataFrame(todas_vistorias)
    df_s = pd.DataFrame(sedes_export_excel)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        if not df_v.empty:
            df_v.to_excel(writer, index=False, sheet_name="Vistorias_Kobo")
        if not df_s.empty:
            df_s.to_excel(writer, index=False, sheet_name="Status_Sedes")
    print(f"   [OK] Planilha Excel sincronizada: {out_xlsx}")
except Exception as e:
    print(f"   [AVISO] Nao foi possivel atualizar Excel: {e}")

print("=" * 65)
print(f"  [CONCLUIDO] {len(visitas_registradas)} sedes com vistoria ativa no mapa!")
print("=" * 65)
