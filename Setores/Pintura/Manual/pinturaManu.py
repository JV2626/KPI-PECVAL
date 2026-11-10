# -*- coding: utf-8 -*-
"""
=============================================================================================
🎨 SETOR 03: PINTURA MANUAL (FILTRO DE COLUNAS OCULTAS + 100% COLAR VALORES)
=============================================================================================
Caminho na Estrutura: PECVAL/AutomacaoKPI/Setores/Pintura/pinturaManu.py
Unidade: Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)

Regras de Engenharia:
  1. Detecta dinamicamente e remove qualquer coluna OCULTA das abas do Apontamento (.xlsm).
  2. Extrai apenas as colunas VISÍVEIS (exatamente como o operador vê na tela).
  3. Mapeia e alinha do início até a coluna Observação (EH no KPI).
  4. ZERO colunas protegidas: cola 100% dos dados como VALOR PURO no KPI Oficial.
=============================================================================================
"""

import os
import re
import sys
import zipfile
import datetime as dt
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

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

SETOR_ID = "PINTURA_MANUAL"
TOTAL_COLUNAS_KPI_PINTURA_MANUAL = 147

# ZERO colunas bloqueadas: tudo entra como VALOR no KPI
COLUNAS_FORMULA_PROTEGIDAS_MANUAL: Set[int] = set()


def obter_colunas_ocultas_aba(caminho_arquivo: Path, nome_aba: str) -> Set[int]:
    """
    Retorna o conjunto de índices 0-based de colunas OCULTAS da aba no Apontamento (.xlsm).
    Lê diretamente as tags <col min="X" max="Y" hidden="1"/> ou width="0" do XML da aba.
    """
    colunas_ocultas: Set[int] = set()
    if not caminho_arquivo or not caminho_arquivo.exists():
        return colunas_ocultas
    try:
        with zipfile.ZipFile(caminho_arquivo, "r") as z:
            wb_xml = z.read("xl/workbook.xml").decode("utf-8", errors="ignore")
            rels_xml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")

            rid_target = {}
            for r in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml):
                rid_target[r[0]] = r[1]
            for r in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels_xml):
                rid_target[r[1]] = r[0]

            target_sheet = None
            alvo_norm = nome_aba.strip().upper()
            for s in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml):
                if s[0].strip().upper() == alvo_norm:
                    target_sheet = rid_target.get(s[1])
                    break

            if not target_sheet:
                return colunas_ocultas

            ts = target_sheet.lstrip("/")
            caminho_sheet = ts if ts.startswith("xl/") else "xl/" + ts
            if caminho_sheet not in z.namelist():
                return colunas_ocultas

            sheet_xml = z.read(caminho_sheet).decode("utf-8", errors="ignore")
            for m in re.finditer(r'<col\s+([^>]+)/?>', sheet_xml):
                attrs = m.group(1)
                m_min = re.search(r'min="(\d+)"', attrs)
                m_max = re.search(r'max="(\d+)"', attrs)
                m_hid = re.search(r'hidden="(1|true|True)"', attrs)
                m_w = re.search(r'width="([^"]+)"', attrs)

                is_hidden = bool(m_hid)
                if not is_hidden and m_w:
                    try:
                        if float(m_w.group(1)) == 0.0:
                            is_hidden = True
                    except ValueError:
                        pass

                if is_hidden and m_min and m_max:
                    for c in range(int(m_min.group(1)), int(m_max.group(1)) + 1):
                        colunas_ocultas.add(c - 1)  # 0-based
    except Exception:
        pass
    return colunas_ocultas


def _eh_valor_turno(val: Any) -> bool:
    if val is None:
        return False
    s = str(val).strip().upper()
    return s.startswith("TURNO") or s in ("A", "B", "C", "1º", "2º", "3º", "ADM", "COMERCIAL")


def _eh_valor_horario_ou_tempo(val: Any) -> bool:
    if val is None or val == "":
        return False
    if isinstance(val, (dt.time, dt.datetime, dt.timedelta)):
        return True
    if isinstance(val, (int, float)) and 0.0 <= float(val) < 2.0:
        return True
    s = str(val).strip()
    return (":" in s) or s.startswith("#")


def _limpar_valor_colagem(val: Any, eh_coluna_tempo: bool = False) -> Any:
    """
    Limpa erros visuais (#VALOR!, #DIV/0!) e converte números decimais de tempo
    para formato HH:MM:SS ou mantém vazio para preservar fórmulas nativas do KPI.
    """
    if val is None or val == "":
        return ""
    if isinstance(val, str) and val.strip().startswith("#"):
        return "00:00:00" if eh_coluna_tempo else ""
    if eh_coluna_tempo and isinstance(val, (int, float)):
        total_sec = max(0, int(round(float(val) * 86400.0)))
        hh = total_sec // 3600
        mm = (total_sec % 3600) // 60
        ss = total_sec % 60
        return f"{hh:02d}:{mm:02d}:{ss:02d}"
    return sanitizar_valor_excel(val)


def transformar_linha_pintura_manual(row_visivel: List[Any]) -> List[Any]:
    """
    Alinha as colunas VISÍVEIS do Apontamento com as 146 colunas de dados do KPI (B..EQ):
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
    num_cols_dados = TOTAL_COLUNAS_KPI_PINTURA_MANUAL - 1  # 146
    linha_kpi = [""] * num_cols_dados

    # Localiza dinamicamente a coluna de Turno na linha já filtrada (visível)
    t_idx = 71
    if len(row_visivel) > 72 and _eh_valor_turno(row_visivel[72]) and not _eh_valor_turno(row_visivel[71]):
        t_idx = 72
    elif len(row_visivel) > 71 and _eh_valor_turno(row_visivel[71]):
        t_idx = 71

    if t_idx == 71:
        for i in range(min(76, len(row_visivel))):
            linha_kpi[i] = _limpar_valor_colagem(row_visivel[i])

        linha_kpi[76] = ""  # Coluna BZ oculta no KPI

        tem_tempo_em_137 = (
            len(row_visivel) > 137
            and _eh_valor_horario_ou_tempo(row_visivel[137])
        )

        if tem_tempo_em_137:
            for orig_i in range(79, len(row_visivel)):
                dest_i = orig_i - 2
                if 77 <= dest_i < num_cols_dados:
                    eh_tempo = (131 <= orig_i <= 137)
                    linha_kpi[dest_i] = _limpar_valor_colagem(row_visivel[orig_i], eh_coluna_tempo=eh_tempo)
        else:
            for orig_i in range(78, len(row_visivel)):
                dest_i = orig_i - 1
                if 77 <= dest_i < num_cols_dados:
                    eh_tempo = (130 <= orig_i <= 136)
                    linha_kpi[dest_i] = _limpar_valor_colagem(row_visivel[orig_i], eh_coluna_tempo=eh_tempo)
    else:
        for i in range(min(70, len(row_visivel))):
            linha_kpi[i] = _limpar_valor_colagem(row_visivel[i])
        if len(row_visivel) > 71:
            linha_kpi[70] = _limpar_valor_colagem(row_visivel[71])
        for offset in range(5):
            orig_i = 72 + offset
            if orig_i < len(row_visivel):
                linha_kpi[71 + offset] = _limpar_valor_colagem(row_visivel[orig_i])
        linha_kpi[76] = ""
        for orig_i in range(79, len(row_visivel)):
            dest_i = orig_i - 2
            if 77 <= dest_i < num_cols_dados:
                eh_tempo = (131 <= orig_i <= 137)
                linha_kpi[dest_i] = _limpar_valor_colagem(row_visivel[orig_i], eh_coluna_tempo=eh_tempo)

    return linha_kpi


def extrair_apontamentos_pintura_manual(
    arquivos_dia: List[Path],
    data_ref: dt.date,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Extrai apenas as linhas e colunas VISÍVEIS de todas as abas da Pintura Manual."""
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
            colunas_ocultas = obter_colunas_ocultas_aba(arq, nome_aba)
            rows = ler_linhas_planilha(arq, nome_aba, max_col=160)
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
                hora_ini = r[72] if (len(r) > 72 and r[72]) else (r[73] if len(r) > 73 else None)
                eh_parada = any(
                    k in kanban.upper()
                    for k in ["SEM DEMANDA", "MANUTENCAO", "TRY OUT", "PARADA", "SETUP"]
                )

                qtd_val = float(qtd_raw) if isinstance(qtd_raw, (int, float)) and qtd_raw > 0 else 0.0
                if qtd_val <= 0 and not (eh_parada and hora_ini):
                    auditoria["linhas_descartadas_vazias"] += 1
                    continue

                # FILTRAGEM DAS COLUNAS OCULTAS DO APONTAMENTO
                if colunas_ocultas:
                    r_visivel = [v for idx_col, v in enumerate(r) if idx_col not in colunas_ocultas]
                else:
                    r_visivel = r

                refugo = float(r[65]) if len(r) > 65 and isinstance(r[65], (int, float)) else 0.0
                boas   = float(r[70]) if len(r) > 70 and isinstance(r[70], (int, float)) else max(0.0, qtd_val - refugo)

                auditoria["total_produzido"] += qtd_val
                auditoria["pecas_refugo"] += refugo
                auditoria["pecas_boas"] += boas
                auditoria["detalhes_abas"][nome_aba]["linhas"] += 1
                auditoria["detalhes_abas"][nome_aba]["pecas"] += qtd_val
                auditoria["detalhes_abas"][nome_aba]["refugo"] += refugo

                valores_kpi = transformar_linha_pintura_manual(r_visivel)

                registros.append({
                    "setor": SETOR_ID,
                    "planta": f"CABINE MANUAL ({nome_aba})",
                    "arquivo": arq.name,
                    "aba": nome_aba,
                    "linha_excel": r_idx + 1,
                    "data": data_ref,
                    "valores_kpi": valores_kpi,
                })
                auditoria["linhas_validas"] += 1

    return registros, auditoria


def processar_setor_pintura_manual(
    mes_alvo: Optional[int] = None,
    dia_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
    modo_oficial: bool = False,
    permitir_atualizar_mesmo_dia: bool = False,
    apenas_validar: bool = False,
) -> Dict[str, Any]:
    """Executa o fluxo da Pintura Manual colando 100% como valores."""
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
            "motivo": f"Nenhum arquivo de apontamento (.xlsm) encontrado para Pintura Manual no mês {mes_efetivo:02d}/{ano_alvo}.",
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
                "motivo": f"Dia {int(dia_alvo):02d}/{mes_efetivo:02d}/{ano_alvo} não encontrado nas pastas de apontamento da Pintura Manual.",
                "dias_disponiveis": [d.strftime("%d/%m/%Y") for d in sorted(apontamentos_map.keys())],
                "diagnostico_kpi": diag,
            }
    else:
        pendentes = diag.get("dias_pendentes", [])
        data_escolhida = pendentes[-1] if pendentes else sorted(apontamentos_map.keys())[-1]

    arquivos_dia = apontamentos_map[data_escolhida]
    registros, auditoria = extrair_apontamentos_pintura_manual(arquivos_dia, data_escolhida)

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
    salvar_planilha_master(registros, caminho_master, "BASE_MASTER_PINTURA_MANUAL", cfg["total_colunas_kpi"] - 1)

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

    res = processar_setor_pintura_manual(
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