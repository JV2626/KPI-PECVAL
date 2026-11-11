# -*- coding: utf-8 -*-
"""
=============================================================================================
⚙️ SETOR 01: INJEÇÃO PLÁSTICA (F1 / LADO A / LADO B -> KPI INJEÇÃO)
=============================================================================================
Caminho na Estrutura: PECVAL/AutomacaoKPI/Setores/Injecao/setor_injecao.py
Unidade: Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)

Regras de Engenharia:
  1. Consolida até 3 arquivos de apontamento (.xlsm) do dia: F1, F2-(LADO A) e F2-(LADO B).
  2. Varre as abas ['ABA 1', 'ABA 2', 'ABA 3', 'ABA 4', 'ABA 5'] a partir da Linha 12 (índice 11).
  3. Aplica validação Poka-Yoke: exige Kanban (Col C), Máquina (Col F) ou Operador (Col B).
  4. Extrai as 138 colunas de dados (alinhadas às colunas B..EI da aba 'APONTAR AQUI').
  5. Valida o dia e a última linha preenchida via ValidadorDia/validador_kpi.py:
     - Se o KPI tiver apenas 1 ou 2 linhas template pré-preenchidas da mesma data (ex: 03/11),
       sobrescreve automaticamente o template na Linha 3 sem bloquear por duplicidade.
=============================================================================================
"""

import sys
import datetime as dt
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

DIRETORIO_AUTOMACAO_KPI = Path(__file__).resolve().parent
while (
    not (DIRETORIO_AUTOMACAO_KPI / "Caminhos").exists()
    and DIRETORIO_AUTOMACAO_KPI.name != "AutomacaoKPI"
    and DIRETORIO_AUTOMACAO_KPI.parent != DIRETORIO_AUTOMACAO_KPI
):
    DIRETORIO_AUTOMACAO_KPI = DIRETORIO_AUTOMACAO_KPI.parent
if str(DIRETORIO_AUTOMACAO_KPI) not in sys.path:
    sys.path.insert(0, str(DIRETORIO_AUTOMACAO_KPI))

from Caminhos.config_caminho import (
    ESTRUTURA_SETORES,
    listar_apontamentos_disponiveis,
    localizar_kpi_oficial,
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

SETOR_ID = "INJECAO"


def identificar_planta_injecao(caminho_arq: Path) -> str:
    """Identifica se o arquivo pertence à F1, LADO A ou LADO B pelo nome ou pasta."""
    texto = f"{caminho_arq.parent.name} {caminho_arq.name}".upper()
    if "F1" in texto:
        return "F1"
    if "LADO A" in texto or "LADO_A" in texto or "LADOA" in texto:
        return "LADO A"
    if "LADO B" in texto or "LADO_B" in texto or "LADOB" in texto:
        return "LADO B"
    return "INJEÇÃO"


def extrair_apontamentos_injecao(
    arquivos_dia: List[Path],
    data_ref: dt.date,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Extrai todas as linhas válidas dos arquivos de apontamento (.xlsm) da Injeção para a data_ref.
    Retorna (registros, resumo_auditoria).
    """
    cfg = ESTRUTURA_SETORES[SETOR_ID]
    num_cols_dados = cfg["total_colunas_kpi"] - 1  # 138 colunas de dados + 1 coluna de Data = 139
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
        "detalhes_plantas": {},
    }

    for arq in arquivos_dia:
        if not arq.exists():
            continue
        planta = identificar_planta_injecao(arq)
        auditoria["detalhes_plantas"].setdefault(
            planta, {"arquivo": arq.name, "linhas": 0, "pecas": 0.0, "refugo": 0.0}
        )

        for nome_aba in cfg["abas_apontamento"]:
            rows = ler_linhas_planilha(arq, nome_aba, max_col=160)
            if len(rows) <= 11:
                continue

            auditoria["linhas_brutas_lidas"] += (len(rows) - 11)
            for r_idx in range(11, len(rows)):
                r = rows[r_idx]
                if not r or len(r) < 7:
                    auditoria["linhas_descartadas_vazias"] += 1
                    continue

                operador = str(r[1] if len(r) > 1 and r[1] is not None else "").strip()
                kanban   = str(r[2] if len(r) > 2 and r[2] is not None else "").strip()
                maquina  = str(r[5] if len(r) > 5 and r[5] is not None else "").strip()

                # Verifica se há parada, motivo ou horário preenchido para não descartar apontamento de setup/parada
                tem_parada_ou_horario = any(
                    (len(r) > col_i and r[col_i] not in (None, "", "0", 0, "0:00", 0.0))
                    for col_i in range(70, min(len(r), 138))
                )

                if not kanban and not maquina and not operador and not tem_parada_ou_horario:
                    auditoria["linhas_descartadas_vazias"] += 1
                    continue

                qtd_raw = r[6] if len(r) > 6 else 0.0
                ref_raw = r[65] if len(r) > 65 else 0.0
                qtd = float(qtd_raw) if isinstance(qtd_raw, (int, float)) else 0.0
                ref = float(ref_raw) if isinstance(ref_raw, (int, float)) else 0.0

                auditoria["total_produzido"] += qtd
                auditoria["pecas_refugo"] += ref
                auditoria["detalhes_plantas"][planta]["pecas"] += qtd
                auditoria["detalhes_plantas"][planta]["refugo"] += ref
                auditoria["detalhes_plantas"][planta]["linhas"] += 1

                valores_brutos = list(r[:num_cols_dados])
                if len(valores_brutos) < num_cols_dados:
                    valores_brutos += [""] * (num_cols_dados - len(valores_brutos))
                valores_kpi = [sanitizar_valor_excel(v) for v in valores_brutos]

                registros.append({
                    "setor": SETOR_ID,
                    "planta": planta,
                    "arquivo": arq.name,
                    "aba": nome_aba,
                    "linha_excel": r_idx + 1,
                    "data": data_ref,
                    "valores_kpi": valores_kpi,
                })
                auditoria["linhas_validas"] += 1

    auditoria["pecas_boas"] = max(0.0, auditoria["total_produzido"] - auditoria["pecas_refugo"])
    return registros, auditoria


def processar_setor_injecao(
    mes_alvo: Optional[int] = None,
    dia_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
    modo_oficial: bool = False,
    permitir_atualizar_mesmo_dia: bool = False,
    apenas_validar: bool = False,
    caminho_kpi_custom: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Executa o fluxo completo do Setor 01 (Injeção Plástica):
      1. Localiza o KPI Oficial do mês escolhido (ex: Outubro=10 ou Novembro=11).
      2. Lista todos os dias de apontamento (.xlsm) disponíveis na pasta do mês.
      3. Seleciona o dia solicitado (ou o último dia pendente/disponível).
      4. Extrai os apontamentos, valida contra o KPI e grava (em Modo Teste ou Oficial).
    """
    cfg = ESTRUTURA_SETORES[SETOR_ID]
    diag = diagnosticar_kpi_setor(
        SETOR_ID,
        caminho_kpi_custom=caminho_kpi_custom,
        mes_alvo=mes_alvo,
        ano_alvo=ano_alvo,
    )
    mes_efetivo = diag.get("mes_detectado") or mes_alvo
    apontamentos_map = listar_apontamentos_disponiveis(SETOR_ID, mes_alvo=mes_efetivo, ano_alvo=ano_alvo)

    if not apontamentos_map:
        return {
            "ok": False,
            "setor_id": SETOR_ID,
            "nome_setor": cfg["nome_exibicao"],
            "motivo": f"Nenhum arquivo de apontamento (.xlsm) encontrado para Injeção no mês {mes_efetivo:02d}/{ano_alvo}.",
            "diagnostico_kpi": diag,
        }

    # Seleciona a data alvo
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
                "motivo": f"Dia {int(dia_alvo):02d}/{mes_efetivo:02d}/{ano_alvo} não encontrado nas pastas de apontamento da Injeção.",
                "dias_disponiveis": [d.strftime("%d/%m/%Y") for d in sorted(apontamentos_map.keys())],
                "diagnostico_kpi": diag,
            }
    else:
        pendentes = diag.get("dias_pendentes", [])
        data_escolhida = pendentes[-1] if pendentes else sorted(apontamentos_map.keys())[-1]

    arquivos_dia = apontamentos_map[data_escolhida]
    registros, auditoria = extrair_apontamentos_injecao(arquivos_dia, data_escolhida)

    if not registros:
        return {
            "ok": False,
            "setor_id": SETOR_ID,
            "nome_setor": cfg["nome_exibicao"],
            "motivo": f"Nenhuma linha válida encontrada nos arquivos de {data_escolhida.strftime('%d/%m/%Y')}.",
            "auditoria": auditoria,
            "diagnostico_kpi": diag,
        }

    # Verifica se a data já existe no KPI (e se é apenas uma linha template de início de mês)
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

    # Salva Base Master do mês
    caminho_master = obter_caminho_master(SETOR_ID, data_escolhida)
    salvar_planilha_master(registros, caminho_master, "BASE_MASTER_INJECAO", cfg["total_colunas_kpi"] - 1)

    # Grava no KPI (Teste ou Oficial)
    res_gravacao = inserir_abaixo_ultima_linha_kpi(
        SETOR_ID,
        registros,
        modo_oficial=modo_oficial,
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
    # Uso via terminal:
    #   python setor_injecao.py               -> Pré-valida o mês/dia atual
    #   python setor_injecao.py 10 30         -> Pré-valida Mês 10, Dia 30
    #   python setor_injecao.py 10 30 teste   -> Grava no KPI Teste (Mês 10, Dia 30)
    #   python setor_injecao.py 10 30 oficial -> Grava no KPI Oficial com Backup (Mês 10, Dia 30)
    args = [a.strip().lower() for a in sys.argv[1:]]
    nums = [int(a) for a in args if a.isdigit()]
    mes_arg = nums[0] if len(nums) >= 1 else None
    dia_arg = nums[1] if len(nums) >= 2 else None
    eh_oficial = "oficial" in args
    eh_teste = "teste" in args
    forcar_att = "atualizar" in args or "overwrite" in args

    diag_inicial = diagnosticar_kpi_setor(SETOR_ID, mes_alvo=mes_arg)
    exibir_relatorio_validador(diag_inicial)

    res = processar_setor_injecao(
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
        print("     💡 Para gravar em TESTE, adicione 'teste' ao comando. Para OFICIAL, adicione 'oficial'.")
    else:
        rg = res["resultado_gravacao"]
        print(f"  🎉 GRAVAÇÃO {rg['modo']} CONCLUÍDA — {res['nome_setor']}")
        print(f"     • Arquivo Gravado:      {rg['arquivo_gravado']}")
        print(f"     • Data Inserida:        {rg['data_inserida'].strftime('%d/%m/%Y')}")
        print(f"     • Faixa Gravada:        Linhas {rg['linha_inicio']} até {rg['linha_fim']} ({rg['linhas_inseridas']} linhas)")
        print(f"     • Motor Utilizado:      {rg['motor_gravacao']}")
    print("=" * 88)