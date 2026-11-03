# -*- coding: utf-8 -*-
"""
Baixa as vistorias do KoboToolbox e gera o Excel consolidado
(relatorios e formularios/Visitas_Monitoramento_Sedes_Kobo.xlsx).
Nao altera os arquivos do site. O token fica em scripts/.env.
"""
import sys
import nofly_core as core

core.preparar_console()
print("=" * 70)
print("  CLEALCO - EXPORTADOR DE VISTORIAS DAS SEDES (KOBOTOOLBOX -> EXCEL)")
print("=" * 70)

cfg = core.carregar_config()
try:
    _, sedes_ref = core.carregar_sedes()
    print(f"-> Base geografica de referencia: {len(sedes_ref)} sedes carregadas.")
except Exception as e:
    print(f"[AVISO] Falha ao ler shapefile de sedes: {e}")
    sedes_ref = {}

vistorias = core.obter_vistorias(sedes_ref, cfg)
ultimas = core.ultima_por_sede(vistorias)
ok = core.gravar_excel(vistorias, core.linhas_status_sedes(sedes_ref, ultimas, cfg))
print("=" * 70)
sys.exit(0 if ok else 1)
