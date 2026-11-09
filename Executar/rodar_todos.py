# -*- coding: utf-8 -*-
"""
=============================================================================================
HUB INTERATIVO E ORQUESTRADOR UNIFICADO — PECVAL (WEB COCKPIT + MENU CLI)
=============================================================================================
Caminho na Estrutura: PECVAL/AutomacaoKPI/Executar/rodar_todos.py
Unidade: Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)

Funcionalidades:
  1. HUB WEB INTERATIVO COMPLETO (http://127.0.0.1:8085):
     - Painel 1: Rastreador de Rede & Caminhos (visão ao vivo do config_caminho.py).
     - Painel 2: Validador de Dia & Transição de Mês (visão ao vivo do validador_kpi.py).
     - Painel 3: Lançador Cirúrgico com Pré-Validador Poka-Yoke.
  2. MODO TERMINAL INTERATIVO (--cli):
     - Lançamento 100% como VALOR em todos os setores (igual ao copiar/colar manual).
=============================================================================================
"""

import os
import sys
import time
import json
import subprocess
import threading
import webbrowser
import importlib.util
import datetime as dt
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

DIRETORIO_AUTOMACAO_KPI = Path(__file__).resolve().parent.parent
if str(DIRETORIO_AUTOMACAO_KPI) not in sys.path:
    sys.path.insert(0, str(DIRETORIO_AUTOMACAO_KPI))

from Caminhos.config_caminho import (
    ANO_OPERACAO_PADRAO,
    DRIVE_V_ATIVO,
    ESTRUTURA_SETORES,
    MESES_PT,
    PASTA_BACKUPS,
    PASTA_MASTERS,
    PASTA_TESTES_KPI,
    RAIZES_APONTAMENTOS_REDE,
    RAIZES_KPIS_OFICIAIS_REDE,
    REDE_UNC_ATIVA,
    _extrair_chave_ano_mes_pasta,
    descobrir_estrutura_rede_automatica,
    limpar_caches_caminho,
    listar_apontamentos_disponiveis,
    localizar_kpi_oficial,
    obter_caminho_master,
    obter_contexto_meses_operacao,
)
from ValidadorDia.validador_kpi import (
    diagnosticar_kpi_setor,
    exibir_relatorio_validador,
    inserir_abaixo_ultima_linha_kpi,
    ler_linhas_planilha,
    salvar_planilha_master,
    sanitizar_valor_excel,
)

# Cache em memória com TTL para navegação instantânea entre meses e abas
_CACHE_ESTADO_HUB: Dict[str, Tuple[float, Dict[str, Any]]] = {}
CACHE_TTL_SEGUNDOS = 300


def limpar_cache_hub() -> None:
    """Limpa o cache em memória após qualquer operação de gravação ou atualização manual."""
    _CACHE_ESTADO_HUB.clear()
    limpar_caches_caminho()


# ==============================================================================
# 0. MOTORES EMBUTIDOS E DE-PARA 100% VALOR (SEM PULAR COLUNAS)
# ==============================================================================
MAPA_DE_PARA_MONTAGEM: List[Tuple[int, int]] = [
    (0, 1), (1, 2), (2, 3), (4, 4), (5, 5), (6, 6),
    (33, 7), (34, 8), (35, 9), (36, 10), (41, 17),
    (44, 18), (45, 19), (46, 20), (47, 21), (48, 22), (49, 23), (50, 24), (53, 28),
    (56, 29), (57, 30), (58, 31), (59, 32), (60, 33), (61, 34), (62, 35), (63, 36),
    (65, 39), (66, 40), (67, 41), (68, 42), (69, 43), (70, 44), (71, 45), (72, 46),
    (73, 47), (74, 48), (75, 49),
    *[(ap_i, ap_i - 26) for ap_i in range(76, 128)],
    (128, 102), (130, 103), (131, 104), (132, 105), (133, 106), (134, 107), (135, 108),
    (136, 109), (138, 111), (139, 112), (140, 113), (142, 115),
]

# ZERO colunas bloqueadas: tudo entra como VALOR no KPI
COLUNAS_FORMULA_PROTEGIDAS_MANUAL: Set[int] = set()

MAPA_DE_PARA_PINTURA_AUTO: Dict[int, int] = {
    **{c: c + 1 for c in range(1, 38)},
    **{c: c for c in range(39, 42)},
    **{c: c + 1 for c in range(42, 53)},
    **{c: c for c in range(54, 66)},
    66: 67,
    67: 68,
    **{c: c for c in range(70, 78)},
    **{c: c - 2 for c in range(80, 147)},
}


def _identificar_planta_injecao(caminho_arq: Path) -> str:
    texto = f"{caminho_arq.parent.name} {caminho_arq.name}".upper()
    if "F1" in texto:
        return "F1"
    if "LADO A" in texto or "LADO_A" in texto or "LADOA" in texto:
        return "LADO A"
    if "LADO B" in texto or "LADO_B" in texto or "LADOB" in texto:
        return "LADO B"
    return "INJEÇÃO"


def _eh_valor_turno(val: Any) -> bool:
    if val is None:
        return False
    s = str(val).strip().upper()
    return s.startswith("TURNO") or s in ("A", "B", "C", "1º", "2º", "3º", "ADM", "COMERCIAL")


def _limpar_valor_colagem(val: Any, eh_coluna_tempo: bool = False) -> Any:
    """Limpa erros visuais do Excel (#VALOR!, #DIV/0!) ao colar como valor."""
    if val is None or val == "":
        return "0:00" if eh_coluna_tempo else ""
    if isinstance(val, str) and val.strip().startswith("#"):
        return "0:00" if eh_coluna_tempo else ""
    return sanitizar_valor_excel(val)


def _transformar_linha_montagem(row_bruta: List[Any]) -> List[Any]:
    vetor = [""] * 115
    for ap_i, kpi_i in MAPA_DE_PARA_MONTAGEM:
        if ap_i < len(row_bruta):
            val = row_bruta[ap_i]
            vetor[kpi_i - 1] = "" if (kpi_i == 17 and val in (0, 0.0, "0")) else sanitizar_valor_excel(val)
    return vetor


def _transformar_linha_pintura_auto(row_bruta: List[Any]) -> List[Any]:
    linha_kpi = [""] * 145
    val_71 = row_bruta[71] if len(row_bruta) > 71 else None

    for col_orig, col_dest in MAPA_DE_PARA_PINTURA_AUTO.items():
        orig_idx = col_orig - 1
        dest_idx = col_dest - 2
        if orig_idx < len(row_bruta) and 0 <= dest_idx < len(linha_kpi):
            linha_kpi[dest_idx] = sanitizar_valor_excel(row_bruta[orig_idx])

    if len(linha_kpi) > 76 and _eh_valor_turno(linha_kpi[72]) and not _eh_valor_turno(linha_kpi[71]):
        linha_kpi[71:76] = linha_kpi[72:77]
        linha_kpi[76] = ""
    elif _eh_valor_turno(val_71) and not _eh_valor_turno(linha_kpi[71]):
        for offset, idx_ap in enumerate(range(70, 76)):
            if idx_ap < len(row_bruta) and (70 + offset) < len(linha_kpi):
                linha_kpi[70 + offset] = sanitizar_valor_excel(row_bruta[idx_ap])

    return linha_kpi


def _transformar_linha_pintura_manual(row_bruta: List[Any]) -> List[Any]:
    """
    Modo 100% Colar Valores da Pintura Manual (Colunas B..EQ do KPI):
      - Colunas B..BT   = Dados iniciais, Kanban, Máquina, Qtd, Defeitos, Totais e Peças Boas
      - Coluna  BU (73) = Turno
      - Colunas BV..BW  = Horário de Produção (Início / Fim)
      - Colunas BX..BY  = Tempo de Setup (Início / Fim)
      - Coluna  BZ (78) = Oculta no KPI ("")
      - Colunas CA..CL  = 4 Paradas Programadas (Início, Fim, Motivo)
      - Colunas CM..DZ  = 10 Paradas Não Programadas (Início, Fim, Categoria, Motivo)
      - Colunas EA..EG  = 7 Tempos Calculados (Setup, Start, Paradas Prog, Não Prog, Sem Demanda, Produtivo, Retrabalho)
      - Coluna  EH (138)= Observação
    """
    num_cols_dados = 146
    linha_kpi = [""] * num_cols_dados

    t_idx = 71
    if len(row_bruta) > 72 and _eh_valor_turno(row_bruta[72]) and not _eh_valor_turno(row_bruta[71]):
        t_idx = 72
    elif len(row_bruta) > 71 and _eh_valor_turno(row_bruta[71]):
        t_idx = 71

    if t_idx == 71:
        for i in range(min(76, len(row_bruta))):
            linha_kpi[i] = _limpar_valor_colagem(row_bruta[i])

        linha_kpi[76] = ""  # Coluna BZ oculta no KPI

        tem_tempo_em_137 = (
            len(row_bruta) > 137
            and row_bruta[137] not in (None, "")
            and (isinstance(row_bruta[137], (dt.time, dt.datetime)) or ":" in str(row_bruta[137]) or str(row_bruta[137]).startswith("#"))
        )

        if tem_tempo_em_137:
            for orig_i in range(79, len(row_bruta)):
                dest_i = orig_i - 2
                if 77 <= dest_i < num_cols_dados:
                    eh_tempo = (131 <= orig_i <= 137)
                    linha_kpi[dest_i] = _limpar_valor_colagem(row_bruta[orig_i], eh_coluna_tempo=eh_tempo)
        else:
            for orig_i in range(78, len(row_bruta)):
                dest_i = orig_i - 1
                if 77 <= dest_i < num_cols_dados:
                    eh_tempo = (130 <= orig_i <= 136)
                    linha_kpi[dest_i] = _limpar_valor_colagem(row_bruta[orig_i], eh_coluna_tempo=eh_tempo)
    else:
        for i in range(min(70, len(row_bruta))):
            linha_kpi[i] = _limpar_valor_colagem(row_bruta[i])
        if len(row_bruta) > 71:
            linha_kpi[70] = _limpar_valor_colagem(row_bruta[71])
        for offset in range(5):
            orig_i = 72 + offset
            if orig_i < len(row_bruta):
                linha_kpi[71 + offset] = _limpar_valor_colagem(row_bruta[orig_i])
        linha_kpi[76] = ""
        for orig_i in range(79, len(row_bruta)):
            dest_i = orig_i - 2
            if 77 <= dest_i < num_cols_dados:
                eh_tempo = (131 <= orig_i <= 137)
                linha_kpi[dest_i] = _limpar_valor_colagem(row_bruta[orig_i], eh_coluna_tempo=eh_tempo)

    return linha_kpi


def _extrair_registros_setor_padrao(
    setor_id: str,
    arquivos_dia: List[Path],
    data_ref: dt.date,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    cfg = ESTRUTURA_SETORES[setor_id]
    num_cols_dados = cfg["total_colunas_kpi"] - 1
    registros: List[Dict[str, Any]] = []
    auditoria: Dict[str, Any] = {
        "setor_id": setor_id,
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
        planta_inj = _identificar_planta_injecao(arq) if setor_id == "INJECAO" else cfg["nome_exibicao"]
        for nome_aba in cfg["abas_apontamento"]:
            # Lê até a coluna 160 para capturar perfeitamente todas as paradas, tempos e observação
            rows = ler_linhas_planilha(arq, nome_aba, max_col=160)
            if len(rows) <= 11:
                continue
            auditoria["linhas_brutas_lidas"] += (len(rows) - 11)
            auditoria["detalhes_abas"].setdefault(nome_aba, {"linhas": 0, "pecas": 0.0, "refugo": 0.0})

            for r_idx in range(11, len(rows)):
                r = rows[r_idx]
                if not r or len(r) < 7:
                    auditoria["linhas_descartadas_vazias"] += 1
                    continue

                operador = str(r[1] if len(r) > 1 and r[1] is not None else "").strip()
                kanban   = str(r[2] if len(r) > 2 and r[2] is not None else "").strip()
                maquina  = str(r[5] if len(r) > 5 and r[5] is not None else "").strip()

                if setor_id == "INJECAO":
                    if not kanban and not maquina and not operador:
                        auditoria["linhas_descartadas_vazias"] += 1
                        continue
                else:
                    if not kanban or not maquina:
                        auditoria["linhas_descartadas_vazias"] += 1
                        continue

                if setor_id == "MONTAGEM":
                    if len(r) < 72:
                        auditoria["linhas_descartadas_vazias"] += 1
                        continue
                    qtd_raw = r[6] if len(r) > 6 else None
                    turno   = str(r[71] if len(r) > 71 and r[71] is not None else "").strip()
                    inicio  = r[72] if len(r) > 72 else None
                    if qtd_raw in (None, "") and not turno and inicio in (None, ""):
                        auditoria["linhas_descartadas_vazias"] += 1
                        continue

                qtd_raw = r[6] if len(r) > 6 else 0.0
                ref_raw = r[65] if len(r) > 65 else 0.0
                qtd_val = float(qtd_raw) if isinstance(qtd_raw, (int, float)) else 0.0
                ref_val = float(ref_raw) if isinstance(ref_raw, (int, float)) else 0.0

                if setor_id == "PINTURA_MANUAL":
                    hora_ini = (r[72] if len(r) > 72 and r[72] else (r[73] if len(r) > 73 else None))
                    eh_parada = any(
                        k in kanban.upper()
                        for k in ["SEM DEMANDA", "MANUTENCAO", "TRY OUT", "PARADA", "SETUP"]
                    )
                    if qtd_val <= 0 and not (eh_parada and hora_ini):
                        auditoria["linhas_descartadas_vazias"] += 1
                        continue

                auditoria["total_produzido"] += max(0.0, qtd_val)
                auditoria["pecas_refugo"] += max(0.0, ref_val)
                auditoria["detalhes_abas"][nome_aba]["linhas"] += 1
                auditoria["detalhes_abas"][nome_aba]["pecas"] += max(0.0, qtd_val)
                auditoria["detalhes_abas"][nome_aba]["refugo"] += max(0.0, ref_val)

                if setor_id == "MONTAGEM":
                    valores_kpi = _transformar_linha_montagem(r)
                elif setor_id == "PINTURA_MANUAL":
                    valores_kpi = _transformar_linha_pintura_manual(r)
                elif setor_id == "PINTURA_AUTOMATICA":
                    valores_kpi = _transformar_linha_pintura_auto(r)
                else:
                    valores_brutos = list(r[:num_cols_dados])
                    if len(valores_brutos) < num_cols_dados:
                        valores_brutos += [""] * (num_cols_dados - len(valores_brutos))
                    valores_kpi = [sanitizar_valor_excel(v) for v in valores_brutos]

                registros.append({
                    "setor": setor_id,
                    "planta": planta_inj if setor_id == "INJECAO" else f"{setor_id} ({nome_aba})",
                    "arquivo": arq.name,
                    "aba": nome_aba,
                    "linha_excel": r_idx + 1,
                    "data": data_ref,
                    "valores_kpi": valores_kpi,
                })
                auditoria["linhas_validas"] += 1

    auditoria["pecas_boas"] = max(0.0, auditoria["total_produzido"] - auditoria["pecas_refugo"])
    return registros, auditoria


def _executar_pipeline_setor_padrao(
    setor_id: str,
    mes_alvo: Optional[int] = None,
    dia_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
    modo_oficial: bool = False,
    permitir_atualizar_mesmo_dia: bool = False,
    apenas_validar: bool = False,
) -> Dict[str, Any]:
    cfg = ESTRUTURA_SETORES[setor_id]
    if setor_id == "PINTURA_MANUAL":
        cfg["colunas_formula_protegidas"] = set()

    diag = diagnosticar_kpi_setor(setor_id, mes_alvo=mes_alvo, ano_alvo=ano_alvo)
    mes_efetivo = diag.get("mes_detectado") or mes_alvo or 11
    apontamentos_map = listar_apontamentos_disponiveis(setor_id, mes_alvo=mes_efetivo, ano_alvo=ano_alvo)

    if not apontamentos_map:
        return {
            "ok": False,
            "setor_id": setor_id,
            "nome_setor": cfg["nome_exibicao"],
            "motivo": f"Nenhum arquivo de apontamento (.xlsm) encontrado para {cfg['nome_exibicao']} no mês {int(mes_efetivo):02d}/{ano_alvo}.",
            "diagnostico_kpi": diag,
        }

    data_escolhida: Optional[dt.date] = None
    if dia_alvo is not None:
        for dt_k in apontamentos_map:
            if dt_k.day == int(dia_alvo):
                data_escolhida = dt_k
                break
        if data_escolhida is None:
            dias_validos = [f"{d.day:02d}" for d in sorted(apontamentos_map.keys())]
            return {
                "ok": False,
                "setor_id": setor_id,
                "nome_setor": cfg["nome_exibicao"],
                "motivo": f"Dia {int(dia_alvo):02d}/{int(mes_efetivo):02d}/{ano_alvo} não encontrado nas pastas de apontamento. Dias existentes: [{', '.join(dias_validos)}].",
                "dias_disponiveis": [d.strftime("%d/%m/%Y") for d in sorted(apontamentos_map.keys())],
                "diagnostico_kpi": diag,
            }
    else:
        pendentes = diag.get("dias_pendentes", [])
        data_escolhida = pendentes[-1] if pendentes else sorted(apontamentos_map.keys())[-1]

    arquivos_dia = apontamentos_map[data_escolhida]
    registros, auditoria = _extrair_registros_setor_padrao(setor_id, arquivos_dia, data_escolhida)
    if not registros:
        return {
            "ok": False,
            "setor_id": setor_id,
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
        "colunas_formula_protegidas": 0 if setor_id == "PINTURA_MANUAL" else len(cfg.get("colunas_formula_protegidas") or []),
    }

    if apenas_validar:
        return {
            "ok": True,
            "modo": "PRE_VALIDACAO",
            "setor_id": setor_id,
            "nome_setor": cfg["nome_exibicao"],
            "pre_validacao": pre_validacao,
            "auditoria": auditoria,
            "diagnostico_kpi": diag,
        }

    caminho_master = obter_caminho_master(setor_id, data_escolhida)
    salvar_planilha_master(registros, caminho_master, f"BASE_MASTER_{setor_id}", cfg["total_colunas_kpi"] - 1)

    # 100% como valor na Pintura Manual (sem pular colunas)
    cols_protegidas = set() if setor_id == "PINTURA_MANUAL" else (cfg.get("colunas_formula_protegidas") or None)

    res_gravacao = inserir_abaixo_ultima_linha_kpi(
        setor_id,
        registros,
        modo_oficial=modo_oficial,
        colunas_formula_protegidas=cols_protegidas,
        permitir_atualizar_mesmo_dia=(permitir_atualizar_mesmo_dia or eh_linha_template),
        caminho_kpi_custom=diag.get("caminho_kpi"),
    )

    return {
        "ok": res_gravacao.get("ok", False),
        "setor_id": setor_id,
        "nome_setor": cfg["nome_exibicao"],
        "pre_validacao": pre_validacao,
        "auditoria": auditoria,
        "resultado_gravacao": res_gravacao,
        "caminho_master": str(caminho_master),
    }


def _carregar_processador_setor(
    setor_id: str,
    caminho_relativo: str,
    nome_funcao: str,
) -> Callable[..., Dict[str, Any]]:
    # Para PINTURA_MANUAL, executa diretamente o pipeline padronizado garantindo 100% valores
    if setor_id == "PINTURA_MANUAL":
        def _runner_manual(mes_alvo=None, dia_alvo=None, ano_alvo=2026, modo_oficial=False, permitir_atualizar_mesmo_dia=False, apenas_validar=False):
            return _executar_pipeline_setor_padrao(
                setor_id="PINTURA_MANUAL",
                mes_alvo=mes_alvo,
                dia_alvo=dia_alvo,
                ano_alvo=ano_alvo,
                modo_oficial=modo_oficial,
                permitir_atualizar_mesmo_dia=permitir_atualizar_mesmo_dia,
                apenas_validar=apenas_validar,
            )
        return _runner_manual

    caminho_arq = DIRETORIO_AUTOMACAO_KPI / caminho_relativo
    if not caminho_arq.exists():
        pasta_setores = DIRETORIO_AUTOMACAO_KPI / "Setores"
        nome_alvo_lower = Path(caminho_relativo).name.lower()
        if pasta_setores.exists():
            for cand in pasta_setores.rglob("*.py"):
                if cand.name.lower() == nome_alvo_lower:
                    caminho_arq = cand
                    break

    if caminho_arq.exists() and caminho_arq.stat().st_size > 50:
        try:
            spec = importlib.util.spec_from_file_location(f"mod_{setor_id.lower()}", str(caminho_arq))
            if spec and spec.loader:
                modulo = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(modulo)
                fn = getattr(modulo, nome_funcao, None)
                if callable(fn):
                    return fn
        except Exception as exc:
            print(f"  ⚠️ [{setor_id}] Usando motor embutido para {caminho_relativo} ({exc})")

    def _fallback_runner(
        mes_alvo: Optional[int] = None,
        dia_alvo: Optional[int] = None,
        ano_alvo: int = 2026,
        modo_oficial: bool = False,
        permitir_atualizar_mesmo_dia: bool = False,
        apenas_validar: bool = False,
    ) -> Dict[str, Any]:
        return _executar_pipeline_setor_padrao(
            setor_id=setor_id,
            mes_alvo=mes_alvo,
            dia_alvo=dia_alvo,
            ano_alvo=ano_alvo,
            modo_oficial=modo_oficial,
            permitir_atualizar_mesmo_dia=permitir_atualizar_mesmo_dia,
            apenas_validar=apenas_validar,
        )

    return _fallback_runner


processar_setor_injecao = _carregar_processador_setor(
    "INJECAO", "Setores/Injecao/setor_injecao.py", "processar_setor_injecao"
)
processar_setor_montagem = _carregar_processador_setor(
    "MONTAGEM", "Setores/Montagem/setor_montagem.py", "processar_setor_montagem"
)
processar_setor_pintura_manual = _carregar_processador_setor(
    "PINTURA_MANUAL", "Setores/Pintura/pinturaManu.py", "processar_setor_pintura_manual"
)
processar_setor_pintura_automatica = _carregar_processador_setor(
    "PINTURA_AUTOMATICA", "Setores/Pintura/Automatico/pinturaAuto.py", "processar_setor_pintura_automatica"
)

MAPA_PROCESSADORES = {
    "INJECAO": processar_setor_injecao,
    "MONTAGEM": processar_setor_montagem,
    "PINTURA_MANUAL": processar_setor_pintura_manual,
    "PINTURA_AUTOMATICA": processar_setor_pintura_automatica,
}


# ==============================================================================
# 1. SERIALIZAÇÃO DE DADOS DO CONFIG_CAMINHO E VALIDADOR_KPI PARA O HUB WEB
# ==============================================================================
def _serializar_diagnostico(diag: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not diag:
        return None
    dias_kpi_lista = []
    for dt_k, info_k in diag.get("dias_no_kpi", {}).items():
        dias_kpi_lista.append({
            "data_br": dt_k.strftime("%d/%m/%Y"),
            "data_iso": dt_k.strftime("%Y-%m-%d"),
            "dia": dt_k.day,
            "mes": dt_k.month,
            "linha_inicio": info_k["linha_inicio"],
            "linha_fim": info_k["linha_fim"],
            "qtd_linhas": info_k["qtd_linhas"],
        })

    dias_apt_lista = []
    for dt_a, info_a in diag.get("dias_apontamentos", {}).items():
        faixa = diag.get("dias_no_kpi", {}).get(dt_a)
        eh_template = bool(faixa and faixa.get("qtd_linhas", 0) <= 2)
        dias_apt_lista.append({
            "data_br": dt_a.strftime("%d/%m/%Y"),
            "data_iso": dt_a.strftime("%Y-%m-%d"),
            "dia": dt_a.day,
            "mes": dt_a.month,
            "ano": dt_a.year,
            "ja_no_kpi": info_a["ja_no_kpi"],
            "eh_template": eh_template,
            "linha_inicio_kpi": faixa["linha_inicio"] if faixa else None,
            "linha_fim_kpi": faixa["linha_fim"] if faixa else None,
            "qtd_linhas_kpi": faixa["qtd_linhas"] if faixa else 0,
            "arquivos": [p.name for p in info_a["arquivos"]],
            "caminhos": [str(p) for p in info_a["arquivos"]],
        })

    return {
        "setor_id": diag["setor_id"],
        "nome_setor": diag["nome_setor"],
        "kpi_encontrado": diag["kpi_encontrado"],
        "nome_kpi": diag["caminho_kpi"].name if diag.get("caminho_kpi") else "Não encontrado",
        "caminho_kpi": str(diag["caminho_kpi"]) if diag.get("caminho_kpi") else "",
        "aba_kpi": diag["aba_kpi"],
        "linha_cabecalho": diag["linha_cabecalho"],
        "mes_detectado": diag["mes_detectado"],
        "nome_mes_detectado": MESES_PT.get(diag.get("mes_detectado") or 0, ""),
        "ano_detectado": diag["ano_detectado"],
        "primeiro_dia_kpi": diag["primeiro_dia_kpi"].strftime("%d/%m/%Y") if diag.get("primeiro_dia_kpi") else None,
        "ultimo_dia_kpi": diag["ultimo_dia_kpi"].strftime("%d/%m/%Y") if diag.get("ultimo_dia_kpi") else None,
        "ultima_linha_preenchida": diag["ultima_linha_preenchida"],
        "proxima_linha_livre": diag["proxima_linha_livre"],
        "total_dias_kpi": len(diag.get("dias_no_kpi", {})),
        "total_linhas_com_dados": diag["total_linhas_com_dados"],
        "status_mes": diag["status_mes"],
        "dias_no_kpi": dias_kpi_lista,
        "dias_apontamentos": dias_apt_lista,
        "dias_pendentes": [d.strftime("%d/%m/%Y") for d in diag.get("dias_pendentes", [])],
        "qtd_pendentes": len(diag.get("dias_pendentes", [])),
        "mes_anterior": _serializar_diagnostico(diag.get("mes_anterior")),
    }


def _diagnosticar_setor_paralelo(
    args: Tuple[str, Dict[str, Any], Optional[int]]
) -> Tuple[str, Dict[str, Any], Dict[str, Any], Optional[Dict[str, Any]]]:
    s_id, cfg, mes_filtro = args
    ctx = obter_contexto_meses_operacao(s_id, ano_alvo=ANO_OPERACAO_PADRAO, mes_referencia=mes_filtro)
    diag = diagnosticar_kpi_setor(s_id, mes_alvo=mes_filtro, ano_alvo=ANO_OPERACAO_PADRAO, incluir_mes_anterior=False)
    diag_ser = _serializar_diagnostico(diag)
    return s_id, cfg, ctx, diag_ser


def coletar_estado_completo_hub(
    mes_filtro: Optional[int] = None,
    forcar_recarregar: bool = False,
) -> Dict[str, Any]:
    """Coleta com Cache e Execução Paralela todos os dados do config_caminho.py e validador_kpi.py."""
    chave_cache = f"mes_{mes_filtro if mes_filtro is not None else 'auto'}"
    if not forcar_recarregar and chave_cache in _CACHE_ESTADO_HUB:
        t_grav, dados = _CACHE_ESTADO_HUB[chave_cache]
        if (time.time() - t_grav) < CACHE_TTL_SEGUNDOS:
            return dados

    pastas_meses_rede = []
    for raiz in RAIZES_APONTAMENTOS_REDE:
        try:
            if raiz.exists():
                for pm in sorted([p for p in raiz.iterdir() if p.is_dir()], key=lambda p: _extrair_chave_ano_mes_pasta(p.name), reverse=True)[:8]:
                    ano_p, mes_p, _ = _extrair_chave_ano_mes_pasta(pm.name)
                    if mes_p > 0:
                        pastas_meses_rede.append({
                            "mes": mes_p,
                            "ano": ano_p,
                            "label": f"{mes_p:02d}/{ano_p} - {MESES_PT.get(mes_p, pm.name)}",
                            "nome_pasta": pm.name,
                            "caminho": str(pm),
                        })
                break
        except OSError:
            continue

    setores_dados = {}
    todos_dias_disponiveis: Dict[str, Dict[str, Any]] = {}

    tarefas = [(s_id, cfg, mes_filtro) for s_id, cfg in ESTRUTURA_SETORES.items()]
    with ThreadPoolExecutor(max_workers=4) as executor:
        resultados = list(executor.map(_diagnosticar_setor_paralelo, tarefas))

    for s_id, cfg, ctx, diag_ser in resultados:
        setores_dados[s_id] = {
            "setor_id": s_id,
            "nome_exibicao": cfg["nome_exibicao"],
            "aba_kpi": cfg["aba_kpi"],
            "linha_cabecalho_kpi": cfg["linha_cabecalho_kpi"],
            "total_colunas_kpi": cfg["total_colunas_kpi"],
            "contexto_caminhos": {
                "mes_atual": ctx["mes_atual"],
                "nome_mes_atual": ctx["nome_mes_atual"],
                "ano_atual": ctx["ano_atual"],
                "kpi_atual_nome": ctx["kpi_atual"].name if ctx.get("kpi_atual") else "Não encontrado",
                "kpi_atual_caminho": str(ctx["kpi_atual"]) if ctx.get("kpi_atual") else "",
                "qtd_dias_apontamentos_atual": len(ctx.get("apontamentos_atual", {})),
                "mes_anterior": ctx["mes_anterior"],
                "nome_mes_anterior": ctx["nome_mes_anterior"],
                "ano_anterior": ctx["ano_anterior"],
                "kpi_anterior_nome": ctx["kpi_anterior"].name if ctx.get("kpi_anterior") else "Não encontrado",
                "kpi_anterior_caminho": str(ctx["kpi_anterior"]) if ctx.get("kpi_anterior") else "",
                "qtd_dias_apontamentos_anterior": len(ctx.get("apontamentos_anterior", {})),
            },
            "validador": diag_ser,
        }

        if diag_ser:
            for d_info in diag_ser.get("dias_apontamentos", []):
                chave = d_info["data_iso"]
                if chave not in todos_dias_disponiveis:
                    todos_dias_disponiveis[chave] = {
                        "data_iso": chave,
                        "data_br": d_info["data_br"],
                        "dia": d_info["dia"],
                        "mes": d_info["mes"],
                        "ano": d_info["ano"],
                        "setores_com_arquivo": [],
                        "setores_pendentes": [],
                    }
                todos_dias_disponiveis[chave]["setores_com_arquivo"].append(s_id)
                if not d_info["ja_no_kpi"] or d_info.get("eh_template"):
                    todos_dias_disponiveis[chave]["setores_pendentes"].append(s_id)

    resultado_final = {
        "timestamp": dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "rede_unc_ativa": REDE_UNC_ATIVA,
        "drive_v_ativo": DRIVE_V_ATIVO,
        "pasta_masters": str(PASTA_MASTERS),
        "pasta_backups": str(PASTA_BACKUPS),
        "pasta_testes_kpi": str(PASTA_TESTES_KPI),
        "mes_filtro_ativo": mes_filtro,
        "pastas_meses_rede": pastas_meses_rede,
        "setores": setores_dados,
        "calendario_dias": sorted(todos_dias_disponiveis.values(), key=lambda x: x["data_iso"], reverse=True),
    }

    _CACHE_ESTADO_HUB[chave_cache] = (time.time(), resultado_final)
    return resultado_final


def executar_lote_ou_validacao(
    setor_alvo: str,
    mes_alvo: Optional[int],
    dia_alvo: Optional[int],
    acao: str,
    permitir_atualizar: bool = False,
) -> Dict[str, Any]:
    """Executa a pré-validação, gravação em KPI Teste ou lançamento no KPI Oficial."""
    setores_lista = (
        ["INJECAO", "MONTAGEM", "PINTURA_MANUAL", "PINTURA_AUTOMATICA"]
        if setor_alvo == "TODOS"
        else [setor_alvo]
    )
    apenas_validar = (acao == "PRE_VALIDAR")
    modo_oficial = (acao == "OFICIAL")

    resultados_setores = []
    for s_id in setores_lista:
        proc_fn = MAPA_PROCESSADORES.get(s_id)
        if not proc_fn:
            continue
        res = proc_fn(
            mes_alvo=mes_alvo,
            dia_alvo=dia_alvo,
            ano_alvo=ANO_OPERACAO_PADRAO,
            modo_oficial=modo_oficial,
            permitir_atualizar_mesmo_dia=permitir_atualizar,
            apenas_validar=apenas_validar,
        )
        rg = res.get("resultado_gravacao", {})
        if rg:
            if isinstance(rg.get("arquivo_gravado"), Path):
                rg["arquivo_gravado"] = str(rg["arquivo_gravado"])
            if isinstance(rg.get("ultimo_dia_antes"), dt.date):
                rg["ultimo_dia_antes"] = rg["ultimo_dia_antes"].strftime("%d/%m/%Y")
            if isinstance(rg.get("data_inserida"), dt.date):
                rg["data_inserida"] = rg["data_inserida"].strftime("%d/%m/%Y")

        resultados_setores.append({
            "setor_id": s_id,
            "nome_setor": res.get("nome_setor", s_id),
            "ok": res.get("ok", False),
            "modo": res.get("modo", acao),
            "motivo": res.get("motivo") or rg.get("motivo", ""),
            "pre_validacao": res.get("pre_validacao"),
            "auditoria": res.get("auditoria"),
            "resultado_gravacao": rg,
            "caminho_master": res.get("caminho_master", ""),
        })

    if acao in ("TESTE", "OFICIAL"):
        limpar_cache_hub()

    return {
        "ok": any(r["ok"] for r in resultados_setores),
        "acao": acao,
        "mes_alvo": mes_alvo,
        "dia_alvo": dia_alvo,
        "resultados": resultados_setores,
    }


def abrir_caminho_no_sistema(caminho_str: str, modo: str = "excel") -> Dict[str, Any]:
    """Abre a planilha diretamente no Excel ou revela na pasta do Windows Explorer / Finder."""
    if not caminho_str:
        return {"ok": False, "mensagem": "Caminho vazio."}
    p = Path(caminho_str)
    if not p.exists():
        return {"ok": False, "mensagem": f"Arquivo/pasta não encontrado: {p}"}
    try:
        if os.name == "nt":
            if modo == "pasta":
                if p.is_file():
                    subprocess.Popen(["explorer", "/select,", str(p)])
                else:
                    os.startfile(str(p))
            else:
                os.startfile(str(p))
        elif sys.platform == "darwin":
            if modo == "pasta":
                subprocess.Popen(["open", "-R", str(p)])
            else:
                subprocess.Popen(["open", str(p)])
        return {"ok": True, "mensagem": f"Aberto com sucesso: {p.name}"}
    except Exception as e:
        return {"ok": False, "mensagem": f"Erro ao abrir: {e}"}


# ==============================================================================
# 2. INTERFACE WEB DO HUB INTERATIVO (COCKPIT PECVAL)
# ==============================================================================
HTML_HUB_INTERATIVO = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Hub Interativo Automação KPI — PECVAL (Toyoda Gosei)</title>
<style>
  :root {
    --bg: #0b0f19;
    --card: #111827;
    --card-2: #1e293b;
    --border: #334155;
    --text: #f8fafc;
    --muted: #94a3b8;
    --blue: #38bdf8;
    --emerald: #10b981;
    --amber: #f59e0b;
    --rose: #f43f5e;
    --purple: #a855f7;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
    padding: 20px;
  }
  .container { max-width: 1440px; margin: 0 auto; }
  .header {
    display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center;
    background: linear-gradient(135deg, #111827 0%, #1e293b 100%);
    border: 1px solid var(--border); border-radius: 14px; padding: 18px 24px; margin-bottom: 18px;
    gap: 16px;
  }
  .header h1 { font-size: 20px; font-weight: 800; display: flex; align-items: center; gap: 10px; }
  .badges { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
  .pill {
    padding: 5px 12px; border-radius: 999px; font-size: 12px; font-weight: 700;
    background: #1e293b; border: 1px solid var(--border); color: var(--text);
  }
  .pill.ok { background: rgba(16,185,129,0.15); border-color: #10b981; color: #34d399; }
  .pill.warn { background: rgba(245,158,11,0.15); border-color: #f59e0b; color: #fbbf24; }
  .pill.err { background: rgba(244,63,94,0.15); border-color: #f43f5e; color: #fb7185; }
  .pill.info { background: rgba(56,189,248,0.15); border-color: #38bdf8; color: #38bdf8; }
  .tabs { display: flex; gap: 10px; margin-bottom: 18px; flex-wrap: wrap; }
  .tab-btn {
    padding: 11px 20px; border-radius: 10px; border: 1px solid var(--border);
    background: var(--card); color: var(--muted); font-weight: 700; font-size: 13.5px;
    cursor: pointer; transition: 0.15s;
  }
  .tab-btn:hover { color: var(--text); border-color: var(--blue); }
  .tab-btn.active {
    background: rgba(56,189,248,0.16); color: var(--blue); border-color: var(--blue);
  }
  .tab-panel { display: none; }
  .tab-panel.active { display: block; }
  .control-bar {
    background: var(--card); border: 1px solid var(--border); border-radius: 14px;
    padding: 18px 22px; margin-bottom: 20px;
    display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; align-items: end;
  }
  .field label { display: block; font-size: 12px; font-weight: 700; color: var(--muted); margin-bottom: 6px; text-transform: uppercase; }
  .field select, .field input {
    width: 100%; padding: 10px 12px; border-radius: 8px; border: 1px solid var(--border);
    background: #0f172a; color: var(--text); font-size: 13.5px; font-weight: 600;
  }
  .btn {
    padding: 10px 16px; border-radius: 8px; border: 1px solid transparent;
    font-weight: 700; font-size: 13px; cursor: pointer; transition: 0.15s;
    display: inline-flex; align-items: center; justify-content: center; gap: 6px;
  }
  .btn:disabled { opacity: 0.5; cursor: not-allowed; }
  .btn-blue { background: #0284c7; color: #fff; }
  .btn-blue:hover:not(:disabled) { background: #0369a1; }
  .btn-emerald { background: #059669; color: #fff; }
  .btn-emerald:hover:not(:disabled) { background: #047857; }
  .btn-amber { background: #d97706; color: #fff; }
  .btn-amber:hover:not(:disabled) { background: #b45309; }
  .btn-ghost { background: #1e293b; border-color: var(--border); color: var(--text); }
  .btn-ghost:hover:not(:disabled) { border-color: var(--blue); color: var(--blue); }
  .btn-sm { padding: 5px 10px; font-size: 11.5px; border-radius: 6px; }
  .grid-2 { display: grid; grid-template-columns: repeat(auto-fit, minmax(600px, 1fr)); gap: 18px; }
  .card {
    background: var(--card); border: 1px solid var(--border); border-radius: 14px;
    padding: 18px 20px;
  }
  .card-title {
    font-size: 15px; font-weight: 800; margin-bottom: 12px;
    display: flex; justify-content: space-between; align-items: center; gap: 8px;
    border-bottom: 1px solid #1e293b; padding-bottom: 10px;
  }
  .sub-box {
    background: #0f172a; border: 1px solid #1e293b; border-radius: 10px;
    padding: 12px 14px; margin-bottom: 12px;
  }
  .sub-box h4 { font-size: 12.5px; text-transform: uppercase; color: var(--muted); margin-bottom: 8px; letter-spacing: 0.4px; }
  .kv-row { display: flex; justify-content: space-between; font-size: 13px; padding: 4px 0; border-bottom: 1px dashed #1e293b; }
  .kv-row:last-child { border-bottom: none; }
  .kv-row span:first-child { color: var(--muted); }
  .kv-row strong { color: var(--text); text-align: right; word-break: break-all; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { padding: 9px 10px; text-align: left; border-bottom: 1px solid #1e293b; }
  th { color: var(--muted); font-size: 11.5px; text-transform: uppercase; background: #0f172a; }
  tr:hover td { background: rgba(255,255,255,0.02); }
  code { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; color: #38bdf8; }
  .preview-box {
    background: linear-gradient(135deg, rgba(16,185,129,0.08) 0%, rgba(56,189,248,0.08) 100%);
    border: 1px solid var(--blue); border-radius: 12px; padding: 18px; margin-bottom: 20px;
    display: none;
  }
  .spinner {
    width: 20px; height: 20px;
    border: 3px solid rgba(56,189,248,0.25);
    border-top-color: #38bdf8;
    border-radius: 50%;
    animation: spin 0.8s linear infinite;
    display: inline-block;
  }
  @keyframes spin { 100% { transform: rotate(360deg); } }
  .pulse { animation: pulseAnim 1.4s ease-in-out infinite; }
  @keyframes pulseAnim {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.45; }
  }
  .toast {
    position: fixed; bottom: 22px; right: 22px; padding: 12px 20px; border-radius: 10px;
    background: #10b981; color: #fff; font-weight: 700; font-size: 13.5px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.45); display: none; z-index: 9999;
  }
</style>
</head>
<body>
<div class="container">
  <!-- TOPBAR -->
  <div class="header">
    <div>
      <h1>🏭 Hub Interativo de Automação KPI — PECVAL (Indaiatuba)</h1>
      <div style="font-size:12.5px; color:var(--muted); margin-top:4px;">
        Visualizador Unificado: <code>config_caminho.py</code> + <code>validador_kpi.py</code> + Lançador Cirúrgico com Pré-Validador Poka-Yoke
      </div>
    </div>
    <div class="badges">
      <span id="badge-rede" class="pill warn pulse">⏳ Verificando Rede...</span>
      <span id="badge-drive-v" class="pill warn pulse">⏳ Verificando V:\...</span>
      <button id="btn-atualizar-leitura" class="btn btn-ghost btn-sm" onclick="carregarEstado(true)">🔄 Atualizar Leitura Agora</button>
    </div>
  </div>

  <!-- BANNER DE STATUS E PROGRESSO AO VIVO -->
  <div id="loading-banner" style="display:none; background:rgba(2,132,199,0.15); border:1px solid #0284c7; padding:12px 18px; border-radius:12px; margin-bottom:18px; font-size:13.5px; color:#38bdf8; display:flex; align-items:center; gap:12px;">
    <span class="spinner"></span>
    <div>
      <strong id="loading-titulo">⏳ Lendo e auditando dados na rede corporativa...</strong>
      <div id="loading-desc" style="font-size:12px; color:var(--muted); margin-top:2px;">
        A primeira leitura de cada mês audita as bases de apontamentos e os KPIs oficiais. As próximas leituras serão instantâneas.
      </div>
    </div>
  </div>

  <!-- NAVEGAÇÃO DE ABAS -->
  <div class="tabs">
    <button class="tab-btn active" onclick="trocarAba('tab-lancador', this)">🚀 1. Lançador & Pré-Validador (Escolher Mês/Dia)</button>
    <button class="tab-btn" onclick="trocarAba('tab-validador', this)">📊 2. Visão do Validador de KPI (validador_kpi.py)</button>
    <button class="tab-btn" onclick="trocarAba('tab-caminhos', this)">📂 3. Rastreador de Caminhos e Meses (config_caminho.py)</button>
  </div>

  <!-- ABA 1: LANÇADOR INTERATIVO -->
  <div id="tab-lancador" class="tab-panel active">
    <div class="control-bar">
      <div class="field">
        <label>1. Escolha o Mês Operacional (2026)</label>
        <select id="sel-mes" onchange="aoMudarMes()">
          <option value="">Automático (Mês Atual da Rede)</option>
          <option value="11">11/2026 — NOVEMBRO (Mês Atual)</option>
          <option value="10">10/2026 — OUTUBRO (Mês Anterior / Fechamento)</option>
          <option value="9">09/2026 — SETEMBRO</option>
          <option value="12">12/2026 — DEZEMBRO</option>
        </select>
      </div>
      <div class="field">
        <label>2. Escolha o Setor</label>
        <select id="sel-setor" onchange="atualizarDropdownDias()">
          <option value="TODOS">🏭 TODOS OS 4 SETORES</option>
          <option value="INJECAO">01. Injeção Plástica (F1 / Lado A / Lado B)</option>
          <option value="MONTAGEM">02. Montagem Industrial (Lado A-B)</option>
          <option value="PINTURA_MANUAL">03. Pintura Manual (Leandro e Marcos)</option>
          <option value="PINTURA_AUTOMATICA">04. Pintura Automática (Rodrigo e Emerson)</option>
        </select>
      </div>
      <div class="field">
        <label>3. Escolha o Dia do Apontamento (.xlsm)</label>
        <select id="sel-dia">
          <option value="">Carregando dias disponíveis...</option>
        </select>
      </div>
      <div class="field">
        <label style="display:flex; align-items:center; gap:6px; cursor:pointer; text-transform:none;">
          <input type="checkbox" id="chk-atualizar" style="width:16px; height:16px;">
          Permitir atualizar se o dia já existir no KPI
        </label>
        <div style="display:flex; gap:8px; margin-top:6px;">
          <button class="btn btn-blue" style="flex:1;" onclick="executarAcao('PRE_VALIDAR')">
            🔍 1. Pré-Validar Dia
          </button>
          <button class="btn btn-amber" style="flex:1;" onclick="executarAcao('TESTE')">
            🧪 2. Gravar KPI Teste
          </button>
          <button class="btn btn-emerald" style="flex:1;" onclick="executarAcao('OFICIAL')">
            🚀 3. Lançar no KPI Oficial
          </button>
        </div>
      </div>
    </div>

    <div id="box-resultado-acao" class="preview-box"></div>
    <div class="grid-2" id="grid-cards-lancador"></div>
  </div>

  <!-- ABA 2: VALIDADOR -->
  <div id="tab-validador" class="tab-panel">
    <div class="grid-2" id="grid-validador-detalhado"></div>
  </div>

  <!-- ABA 3: CAMINHOS -->
  <div id="tab-caminhos" class="tab-panel">
    <div class="grid-2">
      <div class="card">
        <div class="card-title">
          <span>📂 Pastas de Meses Detectadas em APONTAMENTOS (Rede Pecval)</span>
          <span class="pill info">Zero Caminho Fixo</span>
        </div>
        <table>
          <thead>
            <tr><th>Mês/Ano</th><th>Nome da Pasta</th><th>Caminho na Rede</th><th>Ação</th></tr>
          </thead>
          <tbody id="tbody-pastas-meses"></tbody>
        </table>
      </div>
      <div class="card">
        <div class="card-title">
          <span>🛡️ Diretórios de Governança (Masters, Backups e Testes)</span>
          <span class="pill ok">Ativos na Rede V:\</span>
        </div>
        <div id="box-governanca-pastas"></div>
      </div>
    </div>
    <div class="card" style="margin-top:18px;">
      <div class="card-title">
        <span>🗺️ Mapeamento Automático por Setor (Mês Atual vs. Mês Anterior Consolidado)</span>
      </div>
      <table>
        <thead>
          <tr>
            <th>Setor</th>
            <th>📅 KPI Mês Atual</th>
            <th>Apontamentos (Atual)</th>
            <th>📦 KPI Mês Anterior (Consolidado)</th>
            <th>Apontamentos (Anterior)</th>
          </tr>
        </thead>
        <tbody id="tbody-caminhos-setores"></tbody>
      </table>
    </div>
  </div>
</div>

<div id="toast" class="toast"></div>

<script>
let ESTADO_HUB = null;

function mostrarToast(msg, cor = '#10b981') {
  const t = document.getElementById('toast');
  t.style.background = cor;
  t.textContent = msg;
  t.style.display = 'block';
  setTimeout(() => { t.style.display = 'none'; }, 4000);
}

function trocarAba(idAba, btnEl) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById(idAba).classList.add('active');
  btnEl.classList.add('active');
}

let ABORT_CTRL = null;

function setCarregando(ativo, msg = '') {
  const banner = document.getElementById('loading-banner');
  const selMes = document.getElementById('sel-mes');
  const selDia = document.getElementById('sel-dia');
  const btnAtualizar = document.getElementById('btn-atualizar-leitura');
  if (ativo) {
    if (banner) {
      banner.style.display = 'flex';
      if (msg) document.getElementById('loading-titulo').textContent = msg;
    }
    if (selMes) selMes.disabled = true;
    if (btnAtualizar) btnAtualizar.disabled = true;
    if (selDia && (!ESTADO_HUB || !ESTADO_HUB.setores)) {
      selDia.innerHTML = '<option value="">⏳ Lendo apontamentos da rede...</option>';
    }
  } else {
    if (banner) banner.style.display = 'none';
    if (selMes) selMes.disabled = false;
    if (btnAtualizar) btnAtualizar.disabled = false;
  }
}

async function aoMudarMes() {
  await carregarEstado(false);
}

async function carregarEstado(forcar = false) {
  if (ABORT_CTRL) {
    try { ABORT_CTRL.abort(); } catch (e) {}
  }
  ABORT_CTRL = new AbortController();

  const mesVal = document.getElementById('sel-mes').value;
  const mesNome = mesVal === '10' ? 'Outubro' : (mesVal === '11' ? 'Novembro' : (mesVal ? `Mês ${mesVal}` : 'Mês Atual'));
  setCarregando(true, `⏳ Auditando apontamentos e KPIs de ${mesNome} na rede...`);

  let url = mesVal ? `/api/estado?mes=${mesVal}` : '/api/estado';
  if (forcar) url += (url.includes('?') ? '&forcar=1' : '?forcar=1');

  try {
    const resp = await fetch(url, { signal: ABORT_CTRL.signal });
    if (!resp.ok) throw new Error('Servidor retornou status ' + resp.status);
    ESTADO_HUB = await resp.json();
    renderizarTudo();
    setCarregando(false);
  } catch (e) {
    if (e.name === 'AbortError') return;
    setCarregando(false);
    mostrarToast('Aviso ao carregar estado: ' + (e.message || e), '#f43f5e');
  }
}

function renderizarTudo() {
  if (!ESTADO_HUB) return;
  const bRede = document.getElementById('badge-rede');
  bRede.className = 'pill ' + (ESTADO_HUB.rede_unc_ativa ? 'ok' : 'warn');
  bRede.textContent = ESTADO_HUB.rede_unc_ativa ? '🟢 Rede \\\\10.72.110.56: ONLINE' : '🟡 Rede UNC: Offline (Local)';

  const bDrive = document.getElementById('badge-drive-v');
  bDrive.className = 'pill ' + (ESTADO_HUB.drive_v_ativo ? 'ok' : 'warn');
  bDrive.textContent = ESTADO_HUB.drive_v_ativo ? '🟢 Drive V:\\: ONLINE' : '🟡 Drive V:\\: Offline';

  atualizarDropdownDias();
  renderizarCardsLancador();
  renderizarAbaValidador();
  renderizarAbaCaminhos();
}

function atualizarDropdownDias() {
  if (!ESTADO_HUB) return;
  const selSetor = document.getElementById('sel-setor').value;
  const selDia = document.getElementById('sel-dia');
  const valorAnterior = selDia.value;
  selDia.innerHTML = '';

  const opAuto = document.createElement('option');
  opAuto.value = '';
  opAuto.textContent = '⚡ Automático (Último dia pendente na pasta)';
  selDia.appendChild(opAuto);

  if (selSetor === 'TODOS') {
    (ESTADO_HUB.calendario_dias || []).forEach(d => {
      const op = document.createElement('option');
      op.value = d.dia;
      const pend = d.setores_pendentes.length > 0
        ? `⏳ PENDENTE em ${d.setores_pendentes.length} setor(es)`
        : '✅ Já consta nos KPIs';
      op.textContent = `Dia ${d.data_br} — ${pend} (${d.setores_com_arquivo.length} setores com arquivo)`;
      selDia.appendChild(op);
    });
  } else {
    const s = ESTADO_HUB.setores[selSetor];
    const lista = s?.validador?.dias_apontamentos || [];
    lista.slice().reverse().forEach(d => {
      const op = document.createElement('option');
      op.value = d.dia;
      let status = '⏳ PENDENTE (Novo no KPI)';
      if (d.ja_no_kpi && d.eh_template) {
        status = `⚡ Linha Template no KPI (Linha ${d.linha_inicio_kpi}) -> Irá substituir automaticamente`;
      } else if (d.ja_no_kpi) {
        status = `✅ CONSTA NO KPI (Linhas ${d.linha_inicio_kpi}..${d.linha_fim_kpi})`;
      }
      op.textContent = `Dia ${d.data_br} — ${status} [${d.arquivos.join(', ')}]`;
      selDia.appendChild(op);
    });
  }
  if (valorAnterior) selDia.value = valorAnterior;
}

function selecionarDiaRapido(setorId, mesNum, diaNum) {
  document.getElementById('sel-mes').value = String(mesNum);
  document.getElementById('sel-setor').value = setorId;
  carregarEstado().then(() => {
    document.getElementById('sel-dia').value = String(diaNum);
    executarAcao('PRE_VALIDAR');
  });
}

function renderizarCardsLancador() {
  const grid = document.getElementById('grid-cards-lancador');
  grid.innerHTML = '';
  Object.values(ESTADO_HUB.setores).forEach(s => {
    const v = s.validador || {};
    const ant = v.mes_anterior;
    const pendentesAtual = v.dias_pendentes || [];
    const pendentesAnt = ant ? (ant.dias_pendentes || []) : [];

    let alertaFechamentoAnt = '';
    if (ant && pendentesAnt.length > 0) {
      const ultPend = pendentesAnt[pendentesAnt.length - 1];
      const diaNum = parseInt(ultPend.split('/')[0], 10);
      alertaFechamentoAnt = `
        <div class="sub-box" style="border-color:#f59e0b; background:rgba(245,158,11,0.08);">
          <div style="display:flex; justify-content:space-between; align-items:center; gap:8px;">
            <div>
              <strong style="color:#fbbf24;">⚠️ Mês Anterior (${ant.nome_mes_detectado}/${ant.ano_detectado}) tem ${pendentesAnt.length} dia(s) pendente(s): ${pendentesAnt.join(', ')}</strong>
              <div style="font-size:12px; color:var(--muted);">KPI Anterior parou em ${ant.ultimo_dia_kpi || '-'} (Linha ${ant.ultima_linha_preenchida}).</div>
            </div>
            <button class="btn btn-amber btn-sm" onclick="selecionarDiaRapido('${s.setor_id}', ${ant.mes_detectado}, ${diaNum})">
              🔍 Validar Dia ${ultPend}
            </button>
          </div>
        </div>
      `;
    }

    const card = document.createElement('div');
    card.className = 'card';
    card.innerHTML = `
      <div class="card-title">
        <span>${s.nome_exibicao}</span>
        <span class="pill ${v.total_linhas_com_dados === 0 ? 'warn' : 'ok'}">
          ${v.nome_mes_detectado || '-'}/${v.ano_detectado || 2026}: ${v.total_linhas_com_dados === 0 ? 'KPI Vazio (Início de Mês)' : v.total_dias_kpi + ' dia(s) no KPI'}
        </span>
      </div>
      ${alertaFechamentoAnt}
      <div class="sub-box">
        <h4>📅 Status do KPI Selecionado (${v.nome_mes_detectado || '-'}/${v.ano_detectado || 2026})</h4>
        <div class="kv-row"><span>Arquivo KPI Oficial:</span><strong>${v.nome_kpi}</strong></div>
        <div class="kv-row"><span>Aba / Cabeçalho:</span><strong>'${v.aba_kpi}' (Cabeçalho na Linha ${v.linha_cabecalho})</strong></div>
        <div class="kv-row"><span>Último Dia Apontado no KPI:</span><strong style="color:#38bdf8;">${v.ultimo_dia_kpi || 'Nenhum (KPI Vazio)'}</strong></div>
        <div class="kv-row"><span>Última Linha Preenchida:</span><strong>Linha ${v.ultima_linha_preenchida} (${v.total_linhas_com_dados} linhas com dados)</strong></div>
        <div class="kv-row"><span>Próxima Linha Livre:</span><strong style="color:#34d399;">Linha ${v.proxima_linha_livre}</strong></div>
        <div class="kv-row"><span>Apontamentos (.xlsm) na Pasta do Mês:</span><strong>${(v.dias_apontamentos || []).length} dia(s) encontrado(s) | Pendentes: ${pendentesAtual.length}</strong></div>
      </div>
      <div style="display:flex; gap:8px; flex-wrap:wrap;">
        <button class="btn btn-ghost btn-sm" onclick="abrirKpiSetor('${s.setor_id}', 'excel')">📊 Abrir KPI no Excel</button>
        <button class="btn btn-ghost btn-sm" onclick="abrirKpiSetor('${s.setor_id}', 'pasta')">📂 Abrir Pasta do KPI</button>
      </div>
    `;
    grid.appendChild(card);
  });
}

function renderizarAbaValidador() {
  const grid = document.getElementById('grid-validador-detalhado');
  grid.innerHTML = '';
  Object.values(ESTADO_HUB.setores).forEach(s => {
    const v = s.validador || {};
    const ant = v.mes_anterior;
    let blocoAnt = '<div class="sub-box"><h4>📦 Mês Anterior Consolidado</h4><div style="color:var(--muted); font-size:12.5px;">Sem mês anterior carregado para este filtro.</div></div>';
    if (ant && ant.kpi_encontrado) {
      blocoAnt = `
        <div class="sub-box">
          <h4>📦 STATUS DO MÊS ANTERIOR CONSOLIDADO (${ant.nome_mes_detectado}/${ant.ano_detectado})</h4>
          <div class="kv-row"><span>KPI Anterior:</span><strong>${ant.nome_kpi}</strong></div>
          <div class="kv-row"><span>Período Consolidado:</span><strong>${ant.primeiro_dia_kpi || '-'} até ${ant.ultimo_dia_kpi || '-'}</strong></div>
          <div class="kv-row"><span>Total de Dias no KPI:</span><strong>${ant.total_dias_kpi} dias (${ant.total_linhas_com_dados} linhas preenchidas)</strong></div>
          <div class="kv-row"><span>Última Linha Preenchida:</span><strong>Linha ${ant.ultima_linha_preenchida}</strong></div>
          <div class="kv-row">
            <span>Situação de Fechamento:</span>
            <strong style="color:${ant.qtd_pendentes === 0 ? '#34d399' : '#fbbf24'};">
              ${ant.qtd_pendentes === 0 ? '✅ FECHAMENTO COMPLETO' : '⚠️ ATENÇÃO: ' + ant.qtd_pendentes + ' dia(s) pendente(s) (' + ant.dias_pendentes.join(', ') + ')'}
            </strong>
          </div>
        </div>
      `;
    }

    const linhasDias = (v.dias_apontamentos || []).slice().reverse().map(d => `
      <tr>
        <td><strong>${d.data_br}</strong></td>
        <td>
          ${d.ja_no_kpi
            ? (d.eh_template
                ? `<span class="pill warn">Linha Template (Linha ${d.linha_inicio_kpi})</span>`
                : `<span class="pill ok">✅ no KPI (Linhas ${d.linha_inicio_kpi}..${d.linha_fim_kpi})</span>`)
            : `<span class="pill info">⏳ PENDENTE -> Linha ${v.proxima_linha_livre}</span>`
          }
        </td>
        <td><code>${d.arquivos.join(', ')}</code></td>
        <td>
          <button class="btn btn-blue btn-sm" onclick="selecionarDiaRapido('${s.setor_id}', ${d.mes}, ${d.dia})">
            Pré-Validar
          </button>
        </td>
      </tr>
    `).join('');

    const card = document.createElement('div');
    card.className = 'card';
    card.innerHTML = `
      <div class="card-title">
        <span>📊 ${s.nome_exibicao}</span>
        <span class="pill info">Aba: '${v.aba_kpi}'</span>
      </div>
      ${blocoAnt}
      <div class="sub-box">
        <h4>📅 STATUS DO MÊS ANALISADO (${v.nome_mes_detectado || '-'}/${v.ano_detectado || 2026})</h4>
        <div class="kv-row"><span>KPI Atual:</span><strong>${v.nome_kpi}</strong></div>
        <div class="kv-row"><span>Período no KPI:</span><strong>${v.primeiro_dia_kpi ? (v.primeiro_dia_kpi + ' até ' + v.ultimo_dia_kpi) : 'VAZIO / Aguardando 1º Apontamento'}</strong></div>
        <div class="kv-row"><span>Última Linha Preenchida:</span><strong>Linha ${v.ultima_linha_preenchida}</strong></div>
        <div class="kv-row"><span>Próxima Linha de Entrada:</span><strong style="color:#34d399;">Linha ${v.proxima_linha_livre}</strong></div>
      </div>
      <div style="max-height:240px; overflow-y:auto;">
        <table>
          <thead><tr><th>Data</th><th>Status no KPI</th><th>Arquivos (.xlsm)</th><th>Ação</th></tr></thead>
          <tbody>
            ${linhasDias || '<tr><td colspan="4" style="text-align:center; color:var(--muted);">Nenhum arquivo .xlsm encontrado na pasta deste mês.</td></tr>'}
          </tbody>
        </table>
      </div>
    `;
    grid.appendChild(card);
  });
}

function renderizarAbaCaminhos() {
  const tbMeses = document.getElementById('tbody-pastas-meses');
  tbMeses.innerHTML = (ESTADO_HUB.pastas_meses_rede || []).map((pm, idx) => `
    <tr>
      <td><span class="pill info">${pm.mes.toString().padStart(2, '0')}/${pm.ano}</span></td>
      <td><strong>${pm.nome_pasta}</strong></td>
      <td><code>${pm.caminho}</code></td>
      <td><button class="btn btn-ghost btn-sm" onclick="abrirPastaMesIdx(${idx})">📂 Abrir</button></td>
    </tr>
  `).join('') || '<tr><td colspan="4">Usando Fallback Local</td></tr>';

  document.getElementById('box-governanca-pastas').innerHTML = `
    <div class="sub-box">
      <div class="kv-row"><span>📘 Pasta de Bases Master:</span><strong>${ESTADO_HUB.pasta_masters}</strong></div>
      <div class="kv-row"><span>🛡️ Pasta de Backups Automáticos:</span><strong>${ESTADO_HUB.pasta_backups}</strong></div>
      <div class="kv-row"><span>🧪 Pasta de KPIs de Homologação (Teste):</span><strong>${ESTADO_HUB.pasta_testes_kpi}</strong></div>
    </div>
  `;

  const tbSetores = document.getElementById('tbody-caminhos-setores');
  tbSetores.innerHTML = Object.values(ESTADO_HUB.setores).map(s => {
    const c = s.contexto_caminhos;
    return `
      <tr>
        <td><strong>${s.nome_exibicao}</strong></td>
        <td>
          <div><strong>${c.kpi_atual_nome}</strong></div>
          <div style="font-size:11.5px; color:var(--muted);">Mês ${c.mes_atual}/${c.ano_atual} (${c.nome_mes_atual})</div>
        </td>
        <td><span class="pill ok">${c.qtd_dias_apontamentos_atual} dia(s)</span></td>
        <td>
          <div><strong>${c.kpi_anterior_nome}</strong></div>
          <div style="font-size:11.5px; color:var(--muted);">Mês ${c.mes_anterior}/${c.ano_anterior} (${c.nome_mes_anterior})</div>
        </td>
        <td><span class="pill info">${c.qtd_dias_apontamentos_anterior} dia(s)</span></td>
      </tr>
    `;
  }).join('');
}

async function executarAcao(acao) {
  const mes = document.getElementById('sel-mes').value;
  const setor = document.getElementById('sel-setor').value;
  const dia = document.getElementById('sel-dia').value;
  const atualizar = document.getElementById('chk-atualizar').checked;

  if (acao === 'OFICIAL') {
    const confirma = confirm(
      `⚠️ CONFIRMAÇÃO DE LANÇAMENTO OFICIAL NA REDE:\n\n` +
      `Setor: ${setor}\nMês: ${mes || 'Automático'}\nDia: ${dia || 'Último pendente'}\n\n` +
      `O sistema criará um Backup automático em Backups_KPI e gravará abaixo da última linha do KPI Oficial. Deseja prosseguir?`
    );
    if (!confirma) return;
  }

  mostrarToast('⏳ Processando solicitação...', '#0284c7');
  try {
    const resp = await fetch('/api/executar', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        setor: setor,
        mes: mes ? parseInt(mes, 10) : null,
        dia: dia ? parseInt(dia, 10) : null,
        acao: acao,
        permitir_atualizar: atualizar
      })
    });
    const res = await resp.json();
    renderizarResultadoAcao(res);
    if (acao !== 'PRE_VALIDAR') {
      await carregarEstado(true);
    }
  } catch (e) {
    mostrarToast('Erro ao executar ação: ' + e, '#f43f5e');
  }
}

let ULTIMO_RESULTADO_ACAO = null;

function renderizarResultadoAcao(res) {
  ULTIMO_RESULTADO_ACAO = res;
  const box = document.getElementById('box-resultado-acao');
  box.style.display = 'block';

  const blocos = (res.resultados || []).map((r, idx) => {
    if (!r.ok) {
      return `
        <div class="sub-box" style="border-color:#f43f5e;">
          <strong style="color:#fb7185;">❌ ${r.nome_setor}: ${r.motivo}</strong>
        </div>
      `;
    }
    const pv = r.pre_validacao || {};
    const rg = r.resultado_gravacao || {};

    if (res.acao === 'PRE_VALIDAR') {
      const corBadge = pv.acao_prevista === 'BLOQUEADO_JA_EXISTE' ? 'err' : 'ok';
      return `
        <div class="sub-box">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
            <strong style="font-size:14px; color:#38bdf8;">🔍 PRÉ-VALIDAÇÃO: ${r.nome_setor} — Dia ${pv.data_escolhida}</strong>
            <span class="pill ${corBadge}">${pv.acao_prevista}</span>
          </div>
          <div class="kv-row"><span>KPI Alvo:</span><strong>${pv.kpi_alvo} (Aba '${pv.aba_kpi}')</strong></div>
          <div class="kv-row"><span>Última Linha Atual no KPI:</span><strong>Linha ${pv.ultima_linha_atual_kpi}</strong></div>
          <div class="kv-row"><span>🎯 Onde Vai Entrar no KPI:</span><strong style="color:#34d399;">Linhas ${pv.linha_inicio_prevista} até ${pv.linha_fim_prevista} (${pv.qtd_linhas_a_inserir} linhas)</strong></div>
          <div class="kv-row"><span>Volume Extraído do Apontamento:</span><strong>${Math.round(pv.total_pecas_a_inserir).toLocaleString('pt-BR')} peças produzidas (${Math.round(pv.pecas_boas_a_inserir).toLocaleString('pt-BR')} boas | ${Math.round(pv.refugo_a_inserir).toLocaleString('pt-BR')} refugo)</strong></div>
        </div>
      `;
    } else {
      return `
        <div class="sub-box" style="border-color:#10b981;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
            <strong style="font-size:14px; color:#34d399;">✅ GRAVAÇÃO ${rg.modo} CONCLUÍDA: ${r.nome_setor} — Dia ${rg.data_inserida}</strong>
            <span class="pill ok">Motor: ${rg.motor_gravacao}</span>
          </div>
          <div class="kv-row"><span>Arquivo Gravado:</span><strong>${rg.arquivo_gravado}</strong></div>
          <div class="kv-row"><span>Faixa de Linhas Gravada:</span><strong style="color:#34d399;">Linhas ${rg.linha_inicio} até ${rg.linha_fim} (${rg.linhas_inseridas} linhas inseridas)</strong></div>
          <div class="kv-row"><span>Planilha Master Atualizada:</span><strong>${r.caminho_master}</strong></div>
          <div style="margin-top:8px; display:flex; gap:8px;">
            <button class="btn btn-emerald btn-sm" onclick="abrirGravadoIdx(${idx}, 'excel')">📊 Abrir KPI Gravado no Excel</button>
            <button class="btn btn-ghost btn-sm" onclick="abrirGravadoIdx(${idx}, 'pasta')">📂 Mostrar na Pasta</button>
          </div>
        </div>
      `;
    }
  }).join('');

  box.innerHTML = `
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
      <h3 style="font-size:15px;">📋 Resultado da Operação (${res.acao})</h3>
      <button class="btn btn-ghost btn-sm" onclick="document.getElementById('box-resultado-acao').style.display='none'">Fechar</button>
    </div>
    ${blocos}
  `;
  mostrarToast('✅ Operação concluída com sucesso!');
}

function abrirKpiSetor(setorId, modo) {
  const s = ESTADO_HUB?.setores?.[setorId];
  const cam = s?.validador?.caminho_kpi || s?.contexto_caminhos?.kpi_atual_caminho;
  if (!cam) return mostrarToast('Caminho do KPI não encontrado.', '#f43f5e');
  abrirCaminho(cam, modo);
}

function abrirPastaMesIdx(idx) {
  const pm = ESTADO_HUB?.pastas_meses_rede?.[idx];
  if (!pm?.caminho) return mostrarToast('Pasta não encontrada.', '#f43f5e');
  abrirCaminho(pm.caminho, 'pasta');
}

function abrirGravadoIdx(idx, modo) {
  const item = ULTIMO_RESULTADO_ACAO?.resultados?.[idx];
  const cam = item?.resultado_gravacao?.arquivo_gravado;
  if (!cam) return mostrarToast('Arquivo gravado não encontrado.', '#f43f5e');
  abrirCaminho(cam, modo);
}

async function abrirCaminho(caminho, modo) {
  if (!caminho) return;
  try {
    const resp = await fetch('/api/abrir', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ caminho, modo })
    });
    const res = await resp.json();
    mostrarToast(res.mensagem, res.ok ? '#10b981' : '#f43f5e');
  } catch (e) {
    mostrarToast('Erro ao abrir: ' + e, '#f43f5e');
  }
}

window.addEventListener('DOMContentLoaded', () => {
  carregarEstado();
});
</script>
</body>
</html>
"""


class HubAutomacaoHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        pass

    def _enviar_json(self, dados: Any, status: int = 200) -> None:
        corpo = json.dumps(dados, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(corpo)

    def _enviar_html(self, html_str: str) -> None:
        corpo = html_str.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(corpo)))
        self.end_headers()
        self.wfile.write(corpo)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        if parsed.path in ("/", "/index.html"):
            self._enviar_html(HTML_HUB_INTERATIVO)
            return
        if parsed.path == "/api/estado":
            mes_str = qs.get("mes", [""])[0]
            mes_val = int(mes_str) if mes_str.isdigit() else None
            forcar = qs.get("forcar", ["0"])[0] in ("1", "true")
            dados = coletar_estado_completo_hub(mes_filtro=mes_val, forcar_recarregar=forcar)
            self._enviar_json(dados)
            return
        self._enviar_json({"erro": "Rota não encontrada"}, status=404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        tam = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(tam).decode("utf-8") if tam > 0 else "{}"
        payload = json.loads(raw) if raw else {}

        if parsed.path == "/api/executar":
            setor = payload.get("setor", "TODOS")
            mes = payload.get("mes")
            dia = payload.get("dia")
            acao = payload.get("acao", "PRE_VALIDAR")
            atualizar = bool(payload.get("permitir_atualizar", False))
            res = executar_lote_ou_validacao(
                setor_alvo=setor,
                mes_alvo=int(mes) if mes else None,
                dia_alvo=int(dia) if dia else None,
                acao=acao,
                permitir_atualizar=atualizar,
            )
            self._enviar_json(res)
            return

        if parsed.path == "/api/abrir":
            res = abrir_caminho_no_sistema(payload.get("caminho", ""), payload.get("modo", "excel"))
            self._enviar_json(res)
            return

        self._enviar_json({"erro": "Rota POST não encontrada"}, status=404)


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True
    daemon_threads = True


def iniciar_hub_web(porta_inicial: int = 8085, abrir_navegador: bool = True) -> None:
    porta = porta_inicial
    servidor = None
    for p in range(porta_inicial, porta_inicial + 20):
        try:
            servidor = ReusableThreadingHTTPServer(("127.0.0.1", p), HubAutomacaoHandler)
            porta = p
            break
        except OSError:
            continue

    if servidor is None:
        print("❌ Não foi possível alocar uma porta livre entre 8085 e 8105.")
        sys.exit(1)

    url = f"http://127.0.0.1:{porta}"
    print("\n" + "=" * 90)
    print(" 🏭 HUB INTERATIVO DE AUTOMAÇÃO KPI PECVAL — ONLINE!")
    print("=" * 90)
    print(f" 🌐 Acesse no navegador: {url}")
    print(" 💡 No Hub você visualiza config_caminho.py + validador_kpi.py, escolhe Mês/Dia,")
    print("    faz a Pré-Validação Poka-Yoke e lança nos KPIs (Modo Teste ou Oficial com Backup).")
    print(" ⏹️  Pressione CTRL+C neste terminal quando quiser encerrar o Hub.")
    print("=" * 90 + "\n")

    if abrir_navegador:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        servidor.serve_forever()
    except KeyboardInterrupt:
        print("\n⏹️ Hub Interativo encerrado pelo usuário.")
        servidor.server_close()


def menu_interativo_terminal() -> None:
    while True:
        print("\n" + "=" * 88)
        print(" 🏭 MENU INTERATIVO NO TERMINAL — AUTOMAÇÃO KPI PECVAL")
        print("=" * 88)
        print("  [1] Ver Rastreador de Caminhos e Meses (config_caminho.py)")
        print("  [2] Ver Validador de Dia e Transição de Mês (validador_kpi.py)")
        print("  [3] Escolher Mês/Setor/Dia -> Pré-Validar e Lançar no KPI")
        print("  [4] Abrir Hub Web Visual no Navegador")
        print("  [0] Sair")
        opc = input("\nEscolha uma opção: ").strip()

        if opc == "0":
            break
        elif opc == "1":
            m_in = input("Filtrar mês específico? (Digite 10, 11 ou ENTER para automático): ").strip()
            descobrir_estrutura_rede_automatica(mes_alvo=int(m_in) if m_in.isdigit() else None)
        elif opc == "2":
            m_in = input("Filtrar mês específico? (Digite 10, 11 ou ENTER para automático): ").strip()
            m_val = int(m_in) if m_in.isdigit() else None
            for s_id in ESTRUTURA_SETORES:
                exibir_relatorio_validador(diagnosticar_kpi_setor(s_id, mes_alvo=m_val))
        elif opc == "3":
            print("\nSetores: [1] TODOS | [2] INJEÇÃO | [3] MONTAGEM | [4] PINTURA MANUAL | [5] PINTURA AUTOMÁTICA")
            s_in = input("Escolha o setor [Padrão=1 TODOS]: ").strip()
            mapa_s = {"2": "INJECAO", "3": "MONTAGEM", "4": "PINTURA_MANUAL", "5": "PINTURA_AUTOMATICA"}
            setor_esc = mapa_s.get(s_in, "TODOS")

            m_in = input("Digite o número do Mês (ex: 10 para Outubro, 11 para Novembro) [ENTER=10]: ").strip()
            mes_val = int(m_in) if m_in.isdigit() else 10

            s_consulta = "INJECAO" if setor_esc == "TODOS" else setor_esc
            aponts_disp = listar_apontamentos_disponiveis(s_consulta, mes_alvo=mes_val)
            dias_lista = [f"{d.day:02d}" for d in aponts_disp.keys()]
            ult_sugestao = dias_lista[-1] if dias_lista else ""

            if dias_lista:
                print(f"\n📅 Dias reais com apontamentos em {s_consulta} (Mês {mes_val:02d}): [{', '.join(dias_lista)}]")
                d_in = input(f"Digite o Dia [ENTER = Dia {ult_sugestao} (Automático / Recomendado)]: ").strip()
            else:
                d_in = input("Digite o Dia do Apontamento [ENTER=Último pendente]: ").strip()

            dia_val = int(d_in) if d_in.isdigit() else (int(ult_sugestao) if ult_sugestao.isdigit() and not d_in else None)

            res_pre = executar_lote_ou_validacao(
                setor_alvo=setor_esc,
                mes_alvo=mes_val,
                dia_alvo=dia_val,
                acao="PRE_VALIDAR",
            )
            print("\n--- RESULTADO DA PRÉ-VALIDAÇÃO ---")
            tem_ok = False
            precisa_confirmar_overwrite = False
            for r in res_pre["resultados"]:
                if not r["ok"]:
                    print(f"  ❌ {r['nome_setor']}: {r['motivo']}")
                else:
                    tem_ok = True
                    pv = r["pre_validacao"]
                    if pv.get("acao_prevista") == "BLOQUEADO_JA_EXISTE":
                        precisa_confirmar_overwrite = True
                    print(
                        f"  ✅ {r['nome_setor']} | Dia {pv['data_escolhida']} -> KPI: {pv['kpi_alvo']} | "
                        f"Linhas {pv['linha_inicio_prevista']}..{pv['linha_fim_prevista']} "
                        f"({pv['qtd_linhas_a_inserir']} linhas, {pv['total_pecas_a_inserir']:,.0f} pçs) [{pv['acao_prevista']}]"
                    )
            if not tem_ok:
                print("\n⚠️ Nenhum apontamento válido foi localizado para a data informada.")
                print("💡 Dica: Escolha um dos dias listados entre colchetes acima ou aperte ENTER para pegar o dia automático.")
                continue

            permitir_att_cli = False
            if precisa_confirmar_overwrite:
                resp_att = input("\n⚠️ O dia selecionado JÁ CONSTA no KPI. Deseja sobrescrever/atualizar este dia? [S/N]: ").strip().upper()
                permitir_att_cli = resp_att in ("S", "SIM", "Y", "YES")

            print("\nDeseja gravar agora? [1] Gravar em KPI TESTE | [2] Lançar no KPI OFICIAL | [ENTER] Cancelar")
            g_in = input("Opção: ").strip()
            if g_in in ("1", "2"):
                acao_g = "OFICIAL" if g_in == "2" else "TESTE"
                res_g = executar_lote_ou_validacao(
                    setor_alvo=setor_esc,
                    mes_alvo=mes_val,
                    dia_alvo=dia_val,
                    acao=acao_g,
                    permitir_atualizar=permitir_att_cli,
                )
                for r in res_g["resultados"]:
                    if r["ok"]:
                        rg = r["resultado_gravacao"]
                        print(f"  🚀 [{rg['modo']}] {r['nome_setor']}: Gravado nas linhas {rg['linha_inicio']}..{rg['linha_fim']} em {rg['arquivo_gravado']}")
                        if r.get("caminho_master"):
                            print(f"     📘 Master Atualizada: {r['caminho_master']}")
                    else:
                        print(f"  ❌ {r['nome_setor']}: {r['motivo']}")
        elif opc == "4":
            url_hub = "http://127.0.0.1:8085"
            print(f"\n🌐 Abrindo Hub Web no navegador: {url_hub}")
            webbrowser.open(url_hub)


if __name__ == "__main__":
    if "--web-only" in sys.argv:
        iniciar_hub_web(abrir_navegador=("--no-browser" not in sys.argv))
    elif "--check" in sys.argv:
        estado = coletar_estado_completo_hub()
        print("CHECK_HUB_OK:", len(estado["setores"]), "setores mapeados.")
    else:
        t_web = threading.Thread(target=iniciar_hub_web, kwargs={"abrir_navegador": False}, daemon=True)
        t_web.start()
        time.sleep(0.3)
        menu_interativo_terminal()