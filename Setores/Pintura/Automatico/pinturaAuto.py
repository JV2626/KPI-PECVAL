# -*- coding: utf-8 -*-
"""
=============================================================================================
🤖 SETOR 04: PINTURA AUTOMÁTICA (MODO 100% COLAR VALORES)
=============================================================================================
Caminho na Estrutura: PECVAL/AutomacaoKPI/Setores/Pintura/Automatico/pinturaAuto.py
Unidade: Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)

Alinhamento Preciso:
  1. Extrai as 145 colunas úteis VISÍVEIS (ignorando as colunas ocultas do apontamento).
  2. Alinha Turno em BU (73), Horários de Produção em BV..BW (74..75), Setup em BX..BY (76..77).
  3. Coluna BZ (78) oculta no KPI mantida vazia para cravar Paradas em CA (79) em diante!
  4. ZERO colunas protegidas: cola 100% como VALOR PURO no KPI.
=============================================================================================
"""

import sys
import datetime as dt
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

DIRETORIO_AUTOMACAO_KPI = Path(__file__).resolve().parents[3]
if str(DIRETORIO_AUTOMACAO_KPI) not in sys.path:
    sys.path.insert(0, str(DIRETORIO_AUTOMACAO_KPI))

from Caminhos.config_caminho import (
    ESTRUTURA_SETORES,
    listar_apontamentos_disponiveis,
    obter_caminho_master,
)
from ValidadorDia.validador_kpi import (
    diagnosticar_kpi_setor,
    exibir_relatorio_validador,
    inserir_abaixo_ultima_linha_kpi,
    ler_linhas_planilha,
    salvar_planilha_master,
    sanitizar_valor_excel,
)

SETOR_ID = "PINTURA_AUTOMATICA"
TOTAL_COLUNAS_KPI_PINTURA_AUTO = 146  # Coluna A (Data) + 145 colunas de dados (B..EP)

# Mapeamento oficial 1-Based (Alinha colunas visíveis às colunas B..EP do KPI):
MAPA_DE_PARA_PINTURA_AUTO: Dict[int, int] = {
    **{c: c + 1 for c in range(1, 38)},      # Zona 1: 1..37 -> 2..38   [Shift +1]
    **{c: c for c in range(39, 42)},         # Zona 2: 39..41 -> 39..41 [Shift  0]
    **{c: c + 1 for c in range(42, 53)},     # Zona 3: 42..52 -> 43..53 [Shift +1]
    **{c: c for c in range(54, 66)},         # Zona 4: 54..65 -> 54..65 [Shift  0]
    66: 67,                                  # Zona 5: 66..67 -> 67..68 [Shift +1]
    67: 68,
    **{c: c for c in range(70, 78)},         # Zona 6: Col 71 Boas->71(BT), Col 72 Turno->72(BU), Col 73 Ini->73(BV), Col 74 Fim->74(BW), Col 75 Ini Set->75(BX), Col 76 Fim Set->76(BY)
    **{c: c - 2 for c in range(80, 147)},    # Zona 7: Paradas Programadas (CA=79 em diante)
}


def _eh_valor_turno(val: Any) -> bool:
    if val is None:
        return False
    s = str(val).strip().upper()
    return s.startswith("TURNO") or s in ("A", "B", "C", "1º", "2º", "3º", "ADM", "COMERCIAL")


def transformar_linha_pintura_auto(row_bruta: List[Any]) -> List[Any]:
    linha_kpi = [""] * (TOTAL_COLUNAS_KPI_PINTURA_AUTO - 1)
    val_71 = row_bruta[71] if len(row_bruta) > 71 else None

    for col_orig, col_dest in MAPA_DE_PARA_PINTURA_AUTO.items():
        orig_idx = col_orig - 1
        dest_idx = col_dest - 2
        if orig_idx < len(row_bruta) and 0 <= dest_idx < len(linha_kpi):
            val = row_bruta[orig_idx]
            if isinstance(val, str) and val.strip().startswith("#"):
                linha_kpi[dest_idx] = ""
            else:
                linha_kpi[dest_idx] = sanitizar_valor_excel(val)

    # Ancoragem do Turno na Coluna BU (dest_idx 71):
    if len(linha_kpi) > 76 and _eh_valor_turno(linha_kpi[72]) and not _eh_valor_turno(linha_kpi[71]):
        linha_kpi[71:76] = linha_kpi[72:77]
        linha_kpi[76] = ""
    elif _eh_valor_turno(val_71) and not _eh_valor_turno(linha_kpi[71]):
        for offset, idx_ap in enumerate(range(70, 76)):
            if idx_ap < len(row_bruta) and (70 + offset) < len(linha_kpi):
                linha_kpi[70 + offset] = sanitizar_valor_excel(row_bruta[idx_ap])

    return linha_kpi


def extrair_apontamentos_pintura_automatica(
    arquivos_dia: List[Path],
    data_ref: dt.date,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    cfg = ESTRUTURA_SETORES[SETOR_ID]
    registros: List[Dict[str, Any]] = []

    auditoria: Dict[str, Any] = {
        "setor_id": SETOR_ID,
        "nome_setor": cfg["nome_exibicao"],
        "data": data_ref.strftime("%d/%m/%Y"),
        "data_iso": data_ref.strftime("%Y-%m-%d"),
        "arquivos_lidos": [a.name for a in arquivos_dia],
        "linhas_brutas_lidas": 0,
        "linhas_validas": 0,
        "linhas_descartadas_vazias": 0,
        "total_produzido": 0.0,
        "pecas_refugo": 0.0,
        "pecas_boas": 0.0,
        "detalhes_abas": {},
    }

    for arq in arquivos_dia:
        if not arq.exists():
            continue
        for nome_aba in cfg["abas_apontamento"]:
            rows = ler_linhas_planilha(arq, nome_aba, max_col=155)
            if len(rows) <= 11:
                continue

            auditoria["linhas_brutas_lidas"] += (len(rows) - 11)
            auditoria["detalhes_abas"].setdefault(nome_aba, {"linhas": 0, "pecas": 0.0, "refugo": 0.0})

            for r_idx in range(11, len(rows)):
                r = list(rows[r_idx])
                if not r or len(r) < 7:
                    auditoria["linhas_descartadas_vazias"] += 1
                    continue

                kanban  = str(r[2] if len(r) > 2 and r[2] is not None else "").strip()
                maquina = str(r[5] if len(r) > 5 and r[5] is not None else "").strip()
                if not kanban or not maquina:
                    auditoria["linhas_descartadas_vazias"] += 1
                    continue

                qtd_raw  = r[6] if len(r) > 6 else 0.0
                ref_raw  = r[65] if len(r) > 65 else 0.0
                boas_raw = r[70] if len(r) > 70 else (r[71] if len(r) > 71 else 0.0)

                qtd_val = float(qtd_raw) if isinstance(qtd_raw, (int, float)) and qtd_raw > 0 else 0.0
                refugo  = float(ref_raw) if isinstance(ref_raw, (int, float)) else 0.0
                boas    = float(boas_raw) if isinstance(boas_raw, (int, float)) else max(0.0, qtd_val - refugo)

                auditoria["total_produzido"] += qtd_val
                auditoria["pecas_refugo"] += refugo
                auditoria["pecas_boas"] += boas
                auditoria["detalhes_abas"][nome_aba]["linhas"] += 1
                auditoria["detalhes_abas"][nome_aba]["pecas"] += qtd_val
                auditoria["detalhes_abas"][nome_aba]["refugo"] += refugo

                registros.append({
                    "setor": SETOR_ID,
                    "planta": f"LINHA ROBÓTICA ({nome_aba})",
                    "arquivo": arq.name,
                    "aba": nome_aba,
                    "linha_excel": r_idx + 1,
                    "data": data_ref,
                    "valores_kpi": transformar_linha_pintura_auto(r),
                })
                auditoria["linhas_validas"] += 1

    return registros, auditoria


def processar_setor_pintura_automatica(
    mes_alvo: Optional[int] = None,
    dia_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
    modo_oficial: bool = False,
    permitir_atualizar_mesmo_dia: bool = False,
    apenas_validar: bool = False,
) -> Dict[str, Any]:
    cfg = ESTRUTURA_SETORES[SETOR_ID]
    cfg["colunas_formula_protegidas"] = set()

    diag = diagnosticar_kpi_setor(SETOR_ID, mes_alvo=mes_alvo, ano_alvo=ano_alvo)
    mes_efetivo = diag.get("mes_detectado") or mes_alvo
    apontamentos_map = listar_apontamentos_disponiveis(SETOR_ID, mes_alvo=mes_efetivo, ano_alvo=ano_alvo)

    if not apontamentos_map:
        return {
            "ok": False,
            "setor_id": SETOR_ID,
            "nome_setor": cfg["nome_exibicao"],
            "motivo": f"Nenhum arquivo de apontamento (.xlsm) encontrado para Pintura Automática no mês {mes_efetivo:02d}/{ano_alvo}.",
            "diagnostico_kpi": diag,
        }

    data_escolhida: Optional[dt.date] = None
    if dia_alvo is not None:
        for dt_k in apontamentos_map:
            if dt_k.day == int(dia_alvo):
                data_escolhida = dt_k
                break
        if data_escolhida is None:
            return {
                "ok": False,
                "setor_id": SETOR_ID,
                "nome_setor": cfg["nome_exibicao"],
                "motivo": f"Dia {int(dia_alvo):02d}/{mes_efetivo:02d}/{ano_alvo} não encontrado nas pastas de apontamento da Pintura Automática.",
                "dias_disponiveis": [d.strftime("%d/%m/%Y") for d in sorted(apontamentos_map.keys())],
                "diagnostico_kpi": diag,
            }
    else:
        pendentes = diag.get("dias_pendentes", [])
        data_escolhida = pendentes[-1] if pendentes else sorted(apontamentos_map.keys())[-1]

    arquivos_dia = apontamentos_map[data_escolhida]
    registros, auditoria = extrair_apontamentos_pintura_automatica(arquivos_dia, data_escolhida)

    if not registros:
        return {
            "ok": False,
            "setor_id": SETOR_ID,
            "nome_setor": cfg["nome_exibicao"],
            "motivo": f"Nenhuma linha válida encontrada nos arquivos de {data_escolhida.strftime('%d/%m/%Y')}.",
            "auditoria": auditoria,
            "diagnostico_kpi": diag,
        }

    dias_no_kpi = diag.get("dias_no_kpi", {})
    ja_no_kpi = data_escolhida in dias_no_kpi
    info_existente = dias_no_kpi.get(data_escolhida)
    eh_linha_template = bool(ja_no_kpi and info_existente and info_existente.get("qtd_linhas", 0) <= 2 and len(registros) > 2)

    if ja_no_kpi and (permitir_atualizar_mesmo_dia or eh_linha_template):
        linha_ini_prevista = info_existente["linha_inicio"]
        acao_prevista = "SOBRESCREVER_TEMPLATE_INICIO_MES" if eh_linha_template else "ATUALIZAR_DIA_EXISTENTE"
    else:
        linha_ini_prevista = diag["proxima_linha_livre"]
        acao_prevista = "BLOQUEADO_JA_EXISTE" if ja_no_kpi else "INSERIR_ABAIXO_ULTIMA_LINHA"

    linha_fim_prevista = linha_ini_prevista + len(registros) - 1

    pre_validacao = {
        "data_escolhida": data_escolhida.strftime("%d/%m/%Y"),
        "data_iso": data_escolhida.strftime("%Y-%m-%d"),
        "kpi_alvo": diag["caminho_kpi"].name if diag.get("caminho_kpi") else "Não encontrado",
        "caminho_kpi": str(diag["caminho_kpi"]) if diag.get("caminho_kpi") else "",
        "aba_kpi": cfg["aba_kpi"],
        "ja_consta_no_kpi": ja_no_kpi,
        "eh_linha_template": eh_linha_template,
        "acao_prevista": acao_prevista,
        "ultima_linha_atual_kpi": diag["ultima_linha_preenchida"],
        "linha_inicio_prevista": linha_ini_prevista,
        "linha_fim_prevista": linha_fim_prevista,
        "qtd_linhas_a_inserir": len(registros),
        "total_pecas_a_inserir": auditoria["total_produzido"],
        "pecas_boas_a_inserir": auditoria["pecas_boas"],
        "refugo_a_inserir": auditoria["pecas_refugo"],
        "colunas_formula_protegidas": 0,
    }

    if apenas_validar:
        return {
            "ok": True,
            "modo": "PRE_VALIDACAO",
            "setor_id": SETOR_ID,
            "nome_setor": cfg["nome_exibicao"],
            "pre_validacao": pre_validacao,
            "auditoria": auditoria,
            "diagnostico_kpi": diag,
        }

    caminho_master = obter_caminho_master(SETOR_ID, data_escolhida)
    salvar_planilha_master(registros, caminho_master, "BASE_MASTER_PINTURA_AUTO", cfg["total_colunas_kpi"] - 1)

    res_gravacao = inserir_abaixo_ultima_linha_kpi(
        SETOR_ID,
        registros,
        modo_oficial=modo_oficial,
        colunas_formula_protegidas=set(),
        permitir_atualizar_mesmo_dia=(permitir_atualizar_mesmo_dia or eh_linha_template),
        caminho_kpi_custom=diag.get("caminho_kpi"),
    )

    return {
        "ok": res_gravacao.get("ok", False),
        "setor_id": SETOR_ID,
        "nome_setor": cfg["nome_exibicao"],
        "pre_validacao": pre_validacao,
        "auditoria": auditoria,
        "resultado_gravacao": res_gravacao,
        "caminho_master": str(caminho_master),
    }


if __name__ == "__main__":
    args = [a.strip().lower() for a in sys.argv[1:]]
    nums = [int(a) for a in args if a.isdigit()]
    mes_arg = nums[0] if len(nums) >= 1 else None
    dia_arg = nums[1] if len(nums) >= 2 else None
    eh_oficial = "oficial" in args
    eh_teste = "teste" in args
    forcar_att = "atualizar" in args or "overwrite" in args

    diag_inicial = diagnosticar_kpi_setor(SETOR_ID, mes_alvo=mes_arg)
    exibir_relatorio_validador(diag_inicial)

    res = processar_setor_pintura_automatica(
        mes_alvo=mes_arg,
        dia_alvo=dia_arg,
        modo_oficial=eh_oficial,
        permitir_atualizar_mesmo_dia=forcar_att,
        apenas_validar=not (eh_oficial or eh_teste),
    )

    print("\n" + "=" * 88)
    if not res.get("ok"):
        print(f"  ⚠️ {res.get('motivo') or res.get('resultado_gravacao', {}).get('motivo')}")
    elif res.get("modo") == "PRE_VALIDACAO":
        pv = res["pre_validacao"]
        print(f"  ✅ PRÉ-VALIDAÇÃO CONCLUÍDA — {res['nome_setor']}")
        print(f"     • Data Selecionada:     {pv['data_escolhida']}")
        print(f"     • KPI Alvo:             {pv['kpi_alvo']}")
        print(f"     • Linhas Extraídas:     {pv['qtd_linhas_a_inserir']} linhas ({pv['total_pecas_a_inserir']:,.0f} peças)".replace(",", "."))
        print(f"     • Faixa de Destino:     Linhas {pv['linha_inicio_prevista']} até {pv['linha_fim_prevista']} ({pv['acao_prevista']})")
        print("     • Modo de Colagem:      100% VALORES (Sem colunas ocultas)")
        print("     💡 Para gravar em TESTE, adicione 'teste' ao comando. Para OFICIAL, adicione 'oficial'.")
    else:
        rg = res["resultado_gravacao"]
        print(f"  🎉 GRAVAÇÃO {rg['modo']} CONCLUÍDA — {res['nome_setor']}")
        print(f"     • Arquivo Gravado:      {rg['arquivo_gravado']}")
        print(f"     • Data Inserida:        {rg['data_inserida'].strftime('%d/%m/%Y')}")
        print(f"     • Faixa Gravada:        Linhas {rg['linha_inicio']} até {rg['linha_fim']} ({rg['linhas_inseridas']} linhas)")
        print(f"     • Motor Utilizado:      {rg['motor_gravacao']} (100% Colar Valores)")
    print("=" * 88)