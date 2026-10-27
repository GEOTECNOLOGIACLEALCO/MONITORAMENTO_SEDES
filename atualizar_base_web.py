# -*- coding: utf-8 -*-
import os
import json
import geopandas as gpd

print("=" * 60)
print("  CLEALCO - ATUALIZADOR DA BASE WEB DE SEDES")
print("=" * 60)

base_dir = os.path.dirname(os.path.abspath(__file__))
out_js = os.path.join(base_dir, "dados_mapa.js")

visitas_registradas = {}
hist_json = os.path.join(base_dir, "visitas_historico.json")
if os.path.exists(hist_json):
    try:
        with open(hist_json, "r", encoding="utf-8") as f:
            v_list = json.load(f)
            for v in v_list:
                s_id = str(v.get("num_sede", v.get("pontoId", ""))).strip()
                if s_id:
                    visitas_registradas[s_id] = v
        print(f"-> [HISTORICO] {len(visitas_registradas)} vistoria(s) carregada(s) do historico local.")
    except Exception as e:
        print(f"-> [AVISO] Falha ao ler historico de vistorias: {e}")

xlsx_visitas = os.path.join(base_dir, "Visitas_Monitoramento_Sede.xlsx")
if os.path.exists(xlsx_visitas):
    try:
        import pandas as pd
        df_v = pd.read_excel(xlsx_visitas, sheet_name="Historico_Vistorias")
        for _, row in df_v.iterrows():
            s_id = str(row.get("NUM_SEDE", "")).strip()
            if s_id and s_id not in visitas_registradas:
                visitas_registradas[s_id] = {
                    "num_sede": s_id,
                    "data_vistoria": str(row.get("DATA_VISTORIA", "")),
                    "hora_vistoria": str(row.get("HORA", "")),
                    "matricula": str(row.get("MATRICULA_AVALIADOR", "")),
                    "pressao": str(row.get("PRESSAO", "")),
                    "bandeiras": row.get("NUM_BANDEIRAS", ""),
                    "larva_pupa": str(row.get("LARVA_PUPA", "")),
                    "inseticida": str(row.get("INSETICIDA", "")),
                    "observacoes": str(row.get("OBSERVACOES", ""))
                }
        print(f"-> [EXCEL] Total de {len(visitas_registradas)} sede(s) com registro de vistoria.")
    except Exception:
        pass

sede_shp = os.path.join(base_dir, "Sede.shp")
if not os.path.exists(sede_shp):
    sede_shp = os.path.join(base_dir, "SHP", "Sede.shp")

if not os.path.exists(sede_shp):
    print(f"[ERRO] Arquivo de sedes nao encontrado: {sede_shp}")
    exit(1)

print(f"-> Lendo sedes: {os.path.basename(sede_shp)}...")
pts = gpd.read_file(sede_shp, encoding="utf-8").to_crs(epsg=4326)
pts_features = []

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

    try:
        id_int = int(float(num_sede))
    except Exception:
        id_int = idx + 1

    pts_features.append({
        "type": "Feature",
        "geometry": {
            "type": "Point",
            "coordinates": [round(row.geometry.x, 6), round(row.geometry.y, 6)]
        },
        "properties": {
            "id": id_int,
            "num_sede": num_sede,
            "nome_sede": nome_sede if nome_sede else f"Sede {num_sede}",
            "fa": fa,
            "fazenda": fazenda,
            "fa_completo": fa_completo,
            "proprietario": proprietario if proprietario != "nan" else "",
            "regiao": regiao if regiao != "nan" else "",
            "lat": round(row.geometry.y, 6),
            "lon": round(row.geometry.x, 6),
            "visitado": is_visitado,
            "data_vistoria": v_info.get("data_vistoria", v_info.get("data", "")),
            "hora_vistoria": v_info.get("hora_vistoria", v_info.get("hora", "")),
            "matricula": v_info.get("matricula", ""),
            "pressao": v_info.get("pressao", ""),
            "bandeiras": v_info.get("bandeiras", ""),
            "larva_pupa": v_info.get("larva_pupa", ""),
            "inseticida": v_info.get("inseticida", ""),
            "observacoes": v_info.get("observacoes", ""),
            "fotos": v_info.get("fotos", 0)
        }
    })

print(f"   [OK] {len(pts_features)} sedes processadas.")

areas_shp = os.path.join(base_dir, "AREAS.shp")
if not os.path.exists(areas_shp):
    areas_shp = os.path.join(base_dir, "SHP", "AREAS.shp")

areas_features = []
if os.path.exists(areas_shp):
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

print("-> Gravando dados_mapa.js...")
with open(out_js, "w", encoding="utf-8") as f:
    fc_pontos = json.dumps({"type": "FeatureCollection", "features": pts_features}, ensure_ascii=False)
    fc_perims = json.dumps({"type": "FeatureCollection", "features": areas_features}, ensure_ascii=False)
    f.write("const DADOS_SEDES = " + fc_pontos + ";\n")
    f.write("const DADOS_AREAS = " + fc_perims + ";\n")
    f.write("const DADOS_PONTOS = DADOS_SEDES;\n")
    f.write("const DADOS_PERIMETROS = DADOS_AREAS;\n")

tam_mb = os.path.getsize(out_js) / (1024 * 1024)
print(f"[CONCLUIDO] Base web atualizada com sucesso ({tam_mb:.2f} MB)!")
print("=" * 60)
