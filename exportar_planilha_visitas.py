# -*- coding: utf-8 -*-
import os
import json
import pandas as pd
from datetime import datetime

print("=" * 60)
print("  CLEALCO - EXPORTADOR DE VISTORIAS DAS SEDES EM EXCEL")
print("=" * 60)

base_dir = os.path.dirname(os.path.abspath(__file__))
out_xlsx = os.path.join(base_dir, "Visitas_Monitoramento_Sede.xlsx")
hist_json = os.path.join(base_dir, "visitas_historico.json")

visitas = []
if os.path.exists(hist_json):
    try:
        with open(hist_json, "r", encoding="utf-8") as f:
            visitas = json.load(f)
        print(f"-> Total de vistorias carregadas do historico: {len(visitas)}")
    except Exception as e:
        print(f"[ERRO] Falha ao ler arquivo de vistorias: {e}")

# Ler Sedes para consolidar status geral de todas as 77 sedes
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
        print(f"[AVISO] Nao foi possivel carregar Sedes: {e}")

fotos_dir = os.path.join(base_dir, "Fotos_Visitas")
os.makedirs(fotos_dir, exist_ok=True)

rows_visitas = []
for v in visitas:
    num_s = str(v.get("num_sede", v.get("pontoId", ""))).strip()
    nome_s = str(v.get("nome_sede", v.get("ponto", ""))).strip()
    fa = str(v.get("fa", "")).strip()
    fazenda = str(v.get("fazenda", "")).strip()
    prop = str(v.get("proprietario", "")).strip()
    regiao = str(v.get("regiao", "")).strip()
    
    data_v = str(v.get("data_vistoria", v.get("data", ""))).strip()
    hora_v = str(v.get("hora_vistoria", v.get("hora", ""))).strip()
    matricula = str(v.get("matricula", "")).strip()
    
    pressao = str(v.get("pressao", "")).strip()
    bandeiras = v.get("bandeiras", "")
    larva_pupa = str(v.get("larva_pupa", "")).strip()
    inseticida = str(v.get("inseticida", "")).strip()
    
    causador = str(v.get("possiveis_causadores", "")).strip()
    obs = str(v.get("observacoes", "")).strip()
    lat = v.get("lat", "")
    lon = v.get("lon", "")

    rows_visitas.append({
        "DATA_VISTORIA": data_v,
        "HORA": hora_v,
        "NUM_SEDE": num_s,
        "NOME_SEDE": nome_s,
        "FA": fa,
        "FAZENDA": fazenda,
        "PROPRIETARIO": prop,
        "REGIAO": regiao,
        "MATRICULA_AVALIADOR": matricula,
        "PRESSAO": pressao,
        "NUM_BANDEIRAS": bandeiras,
        "LARVA_PUPA": larva_pupa,
        "INSETICIDA": inseticida,
        "POSSIVEIS_CAUSADORES": causador,
        "OBSERVACOES": obs,
        "LATITUDE": lat,
        "LONGITUDE": lon,
        "REGISTRADO_EM": v.get("salvoEm", "")
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

        # Calcular dias e status tematico
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

df_visitas = pd.DataFrame(rows_visitas)
if not df_visitas.empty:
    df_visitas.sort_values(by=["DATA_VISTORIA", "HORA"], ascending=[False, False], inplace=True)

df_sedes = pd.DataFrame(list(sedes_dict.values()))

while True:
    try:
        with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
            if not df_visitas.empty:
                df_visitas.to_excel(writer, index=False, sheet_name="Historico_Vistorias")
            else:
                df_empty = pd.DataFrame(columns=[
                    "DATA_VISTORIA", "HORA", "NUM_SEDE", "NOME_SEDE", "FA", "FAZENDA",
                    "PROPRIETARIO", "REGIAO", "MATRICULA_AVALIADOR", "PRESSAO",
                    "NUM_BANDEIRAS", "LARVA_PUPA", "INSETICIDA", "POSSIVEIS_CAUSADORES",
                    "OBSERVACOES", "LATITUDE", "LONGITUDE", "REGISTRADO_EM"
                ])
                df_empty.to_excel(writer, index=False, sheet_name="Historico_Vistorias")

            if not df_sedes.empty:
                df_sedes.to_excel(writer, index=False, sheet_name="Status_Sedes")
        print(f"[CONCLUIDO] Planilha Excel gerada com sucesso:")
        print(f"  -> {out_xlsx}")
        break
    except PermissionError:
        print(f"\n[AVISO] O arquivo '{os.path.basename(out_xlsx)}' esta aberto no Excel!")
        print("Por favor, FECHE a planilha no Excel e pressione ENTER para continuar...")
        input()

print("=" * 60)
