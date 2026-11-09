# -*- coding: utf-8 -*-
"""
=============================================================================================
MÓDULO 2: VALIDADOR DE DIA, ÚLTIMA LINHA E INJETOR CIRÚRGICO DE KPI (VERSÃO BLINDADA V2)
=============================================================================================
Caminho na Estrutura: PECVAL/AutomacaoKPI/ValidadorDia/validador_kpi.py
Unidade: Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)
=============================================================================================
"""

import os
import re
import sys
import time
import uuid
import io
import shutil
import zipfile
import tempfile
import warnings
import datetime as dt
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

# Silencia avisos internos do OpenPyXL caso seja usado como fallback de leitura
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Blindagem contra UnicodeEncodeError no terminal do Windows (CMD / PowerShell / VS Code)
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    
import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill

try:
    from python_calamine import CalamineWorkbook
    TEM_CALAMINE = True
except ImportError:
    TEM_CALAMINE = False

try:
    import win32com.client
    TEM_WIN32COM = True
except ImportError:
    TEM_WIN32COM = False

# Adiciona a raiz AutomacaoKPI ao sys.path para permitir importação limpa entre pastas
DIRETORIO_AUTOMACAO_KPI = Path(__file__).resolve().parent.parent
if str(DIRETORIO_AUTOMACAO_KPI) not in sys.path:
    sys.path.insert(0, str(DIRETORIO_AUTOMACAO_KPI))

from Caminhos.config_caminho import (
    ESTRUTURA_SETORES,
    PASTA_BACKUPS,
    _eh_kpi_do_setor,
    _extrair_chave_ano_mes_pasta,
    calcular_mes_ano_anterior,
    eh_diretorio_legado_ou_ignorado,
    listar_apontamentos_disponiveis,
    localizar_kpi_oficial,
    normalizar_texto,
    obter_caminho_kpi_teste,
)


# ==============================================================================
# 1. LEITURA RÁPIDA DE PLANILHAS (.XLSM / .XLSX) E PARSER DE DATAS
# ==============================================================================
def converter_valor_para_data(val: Any) -> Optional[dt.date]:
    """Converte qualquer representação de célula Excel na Coluna A para dt.date."""
    if val is None or val == "":
        return None
    if isinstance(val, dt.datetime):
        return val.date()
    if isinstance(val, dt.date):
        return val
    if isinstance(val, (int, float)) and 40000 <= val <= 60000:
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(val))

    texto = str(val).strip()
    m_iso = re.match(r"^(\d{4})-(\d{2})-(\d{2})", texto)
    if m_iso:
        try:
            return dt.date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
        except ValueError:
            return None

    m_br = re.match(r"^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})", texto)
    if m_br:
        d, m, y = int(m_br.group(1)), int(m_br.group(2)), int(m_br.group(3))
        if y < 100:
            y += 2000
        try:
            return dt.date(y, m, d)
        except ValueError:
            return None
    return None


def sanitizar_valor_excel(val: Any) -> Any:
    """Limpa valores vindos do apontamento antes de gravar no Excel."""
    if val is None:
        return ""
    if isinstance(val, dt.time):
        return val.strftime("%H:%M:%S")
    if isinstance(val, dt.datetime):
        return val.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(val, dt.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str) and val.strip().startswith("#"):
        return ""
    return val


def resolver_nome_aba(nomes_disponiveis: List[str], aba_esperada: str) -> Optional[str]:
    """Encontra o nome exato da aba mesmo que haja espaço invisível no final ('APONTAR AQUI ')."""
    if aba_esperada in nomes_disponiveis:
        return aba_esperada
    alvo_limpo = aba_esperada.strip().upper()
    for nome in nomes_disponiveis:
        if nome.strip().upper() == alvo_limpo:
            return nome
    return None


def ler_linhas_planilha(caminho_arquivo: Path, nome_aba: str, max_col: Optional[int] = None) -> List[List[Any]]:
    """
    Lê todas as linhas de uma aba usando o motor Rust Calamine (ultrarrápido)
    ou OpenPyXL em modo read_only como fallback seguro.
    """
    if not caminho_arquivo or not caminho_arquivo.exists():
        return []
    try:
        if caminho_arquivo.stat().st_size == 0:
            return []
    except OSError:
        return []

    if TEM_CALAMINE:
        try:
            dados_ram = caminho_arquivo.read_bytes()
            wb = CalamineWorkbook.from_filelike(io.BytesIO(dados_ram))
            aba_real = resolver_nome_aba(wb.sheet_names, nome_aba)
            if aba_real:
                sh = wb.get_sheet_by_name(aba_real)
                if not (sh.height and sh.height > 25000):
                    rows = sh.to_python(skip_empty_area=True)
                    if sh.start and sh.start[0] > 0:
                        rows = [[] for _ in range(sh.start[0])] + rows
                    return rows
        except Exception:
            pass

    try:
        wb_ox = openpyxl.load_workbook(str(caminho_arquivo), read_only=True, data_only=True)
    except Exception:
        return []

    try:
        aba_real = resolver_nome_aba(wb_ox.sheetnames, nome_aba)
        if not aba_real:
            return []
        ws = wb_ox[aba_real]
        linhas: List[List[Any]] = []
        vazias_seguidas = 0
        for row in ws.iter_rows(max_col=max_col, values_only=True):
            lista_r = list(row)
            linhas.append(lista_r)
            if len(linhas) > 15 and all(v is None or v == "" for v in lista_r[:10]):
                vazias_seguidas += 1
                if vazias_seguidas >= 150:
                    break
            else:
                vazias_seguidas = 0
        return linhas
    except Exception:
        return []
    finally:
        try:
            wb_ox.close()
        except Exception:
            pass


_CACHE_ZIP_DATAS: Dict[Tuple[str, float, str], List[Tuple[int, dt.date]]] = {}


def extrair_datas_coluna_a_zip(
    caminho_kpi: Path,
    nome_aba: str = "APONTAR AQUI",
) -> Optional[List[Tuple[int, dt.date]]]:
    """Leitor Cirúrgico OpenXML (Buffer em Memória RAM)."""
    if not caminho_kpi or not caminho_kpi.exists():
        return None
    try:
        mtime = caminho_kpi.stat().st_mtime
    except OSError:
        mtime = 0.0

    chave = (str(caminho_kpi), mtime, nome_aba.strip().upper())
    if chave in _CACHE_ZIP_DATAS:
        return _CACHE_ZIP_DATAS[chave]

    try:
        dados_bytes = caminho_kpi.read_bytes()
        with zipfile.ZipFile(io.BytesIO(dados_bytes)) as z:
            wb_xml = z.read("xl/workbook.xml")
            rels_xml = z.read("xl/_rels/workbook.xml.rels")

            rid_target: Dict[str, str] = {}
            for r in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml.decode("utf-8", errors="ignore")):
                rid_target[r[0]] = r[1]
            for r in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels_xml.decode("utf-8", errors="ignore")):
                rid_target[r[1]] = r[0]

            target_sheet = None
            alvo_norm = nome_aba.strip().upper()
            for s in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml.decode("utf-8", errors="ignore")):
                if s[0].strip().upper() == alvo_norm:
                    target_sheet = rid_target.get(s[1])
                    break
            if not target_sheet:
                return None

            ts = target_sheet.lstrip("/")
            caminho_sheet = ts if ts.startswith("xl/") else "xl/" + ts
            raw_xml = z.read(caminho_sheet)

            sst: List[str] = []
            if "xl/sharedStrings.xml" in z.namelist():
                sst_raw = z.read("xl/sharedStrings.xml").decode("utf-8", errors="ignore")
                sst = re.findall(r"<t[^>]*>([^<]*)</t>", sst_raw)

            padrao = re.compile(rb'<c\s+r="A(\d+)"([^>]*)>(?:<f[^>]*>.*?</f>)?<v>([^<]+)</v>')
            linhas_data: List[Tuple[int, dt.date]] = []
            for r_num_b, attrs_b, val_b in padrao.findall(raw_xml):
                r_num = int(r_num_b)
                val_str = val_b.decode("utf-8", errors="ignore").strip()
                if b't="s"' in attrs_b:
                    idx_sst = int(val_str) if val_str.isdigit() else -1
                    val_real = sst[idx_sst] if 0 <= idx_sst < len(sst) else ""
                else:
                    try:
                        val_real = float(val_str)
                    except ValueError:
                        val_real = val_str
                data_obj = converter_valor_para_data(val_real)
                if data_obj:
                    linhas_data.append((r_num, data_obj))

            _CACHE_ZIP_DATAS[chave] = linhas_data
            return linhas_data
    except Exception:
        return None


# ==============================================================================
# 2. VALIDADOR DE DIA E ÚLTIMA LINHA DO KPI ORIGINAL
# ==============================================================================
def diagnosticar_kpi_setor(
    setor_id: str,
    caminho_kpi_custom: Optional[Path] = None,
    mes_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
    incluir_mes_anterior: bool = True,
) -> Dict[str, Any]:
    cfg = ESTRUTURA_SETORES[setor_id]
    kpi_path = caminho_kpi_custom or localizar_kpi_oficial(setor_id, ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    apontamentos_pasta = listar_apontamentos_disponiveis(setor_id, mes_alvo=mes_alvo, ano_alvo=ano_alvo)

    mes_detectado = mes_alvo
    ano_detectado = ano_alvo
    if kpi_path:
        ano_k, mes_k, _ = _extrair_chave_ano_mes_pasta(kpi_path.name, ano_fallback=ano_alvo)
        if 1 <= mes_k <= 12:
            mes_detectado = mes_k
        if ano_k != 0:
            ano_detectado = ano_k

    kpi_existe = bool(kpi_path and kpi_path.exists())

    resultado: Dict[str, Any] = {
        "setor_id": setor_id,
        "nome_setor": cfg["nome_exibicao"],
        "kpi_encontrado": kpi_existe,
        "caminho_kpi": kpi_path,
        "aba_kpi": cfg["aba_kpi"],
        "linha_cabecalho": cfg["linha_cabecalho_kpi"],
        "mes_detectado": mes_detectado,
        "ano_detectado": ano_detectado,
        "primeiro_dia_kpi": None,
        "ultimo_dia_kpi": None,
        "ultima_linha_preenchida": cfg["linha_cabecalho_kpi"],
        "proxima_linha_livre": cfg["linha_cabecalho_kpi"] + 1,
        "total_linhas_com_dados": 0,
        "status_mes": "NAO_ENCONTRADO" if not kpi_existe else "EM_ANDAMENTO",
        "status_operacional": "NAO_ENCONTRADO" if not kpi_existe else "EM_ANDAMENTO",
        "dias_no_kpi": {},
        "dias_apontamentos": {},
        "dias_pendentes": [],
        "linhas_orfas": [],
        "mes_anterior": None,
        "transicao_ativa": False,
    }

    if not kpi_existe:
        return resultado

    linha_hdr = cfg["linha_cabecalho_kpi"]
    dias_map: Dict[dt.date, Dict[str, int]] = {}
    ultima_linha_continua = linha_hdr
    total_linhas_validas = 0
    linhas_orfas: List[int] = []

    res_zip = extrair_datas_coluna_a_zip(kpi_path, cfg["aba_kpi"])
    if res_zip is not None:
        ultimo_r_valido = linha_hdr
        for num_linha_excel, data_celula in res_zip:
            if num_linha_excel <= linha_hdr:
                continue
            if (num_linha_excel - ultimo_r_valido) > 25 and ultimo_r_valido > linha_hdr:
                linhas_orfas.append(num_linha_excel)
            else:
                ultima_linha_continua = num_linha_excel
                ultimo_r_valido = num_linha_excel
                total_linhas_validas += 1
                if data_celula not in dias_map:
                    dias_map[data_celula] = {
                        "linha_inicio": num_linha_excel,
                        "linha_fim": num_linha_excel,
                        "qtd_linhas": 1,
                    }
                else:
                    dias_map[data_celula]["linha_fim"] = num_linha_excel
                    dias_map[data_celula]["qtd_linhas"] += 1
    else:
        linhas = ler_linhas_planilha(kpi_path, cfg["aba_kpi"], max_col=10)
        vazias_consecutivas = 0
        for idx in range(linha_hdr, len(linhas)):
            row = linhas[idx]
            num_linha_excel = idx + 1
            data_celula = converter_valor_para_data(row[0] if (row and len(row) > 0) else None)
            if data_celula is not None:
                if vazias_consecutivas > 25 and ultima_linha_continua > linha_hdr:
                    linhas_orfas.append(num_linha_excel)
                else:
                    ultima_linha_continua = num_linha_excel
                    total_linhas_validas += 1
                    vazias_consecutivas = 0
                    if data_celula not in dias_map:
                        dias_map[data_celula] = {
                            "linha_inicio": num_linha_excel,
                            "linha_fim": num_linha_excel,
                            "qtd_linhas": 1,
                        }
                    else:
                        dias_map[data_celula]["linha_fim"] = num_linha_excel
                        dias_map[data_celula]["qtd_linhas"] += 1
            else:
                vazias_consecutivas += 1

    dias_ordenados = sorted(dias_map.keys())
    resultado["dias_no_kpi"] = {d: dias_map[d] for d in dias_ordenados}
    resultado["total_linhas_com_dados"] = total_linhas_validas
    resultado["ultima_linha_preenchida"] = ultima_linha_continua
    resultado["proxima_linha_livre"] = ultima_linha_continua + 1
    resultado["linhas_orfas"] = linhas_orfas

    if dias_ordenados:
        resultado["primeiro_dia_kpi"] = dias_ordenados[0]
        resultado["ultimo_dia_kpi"] = dias_ordenados[-1]
        if mes_detectado is None:
            mes_detectado = dias_ordenados[-1].month
            resultado["mes_detectado"] = mes_detectado
        if ano_detectado is None:
            ano_detectado = dias_ordenados[-1].year
            resultado["ano_detectado"] = ano_detectado

    for dt_apt, arqs in apontamentos_pasta.items():
        ja_existe = dt_apt in dias_map
        resultado["dias_apontamentos"][dt_apt] = {
            "arquivos": arqs,
            "ja_no_kpi": ja_existe,
        }
        if not ja_existe:
            resultado["dias_pendentes"].append(dt_apt)

    if total_linhas_validas == 0:
        resultado["status_mes"] = "VAZIO_AGUARDANDO_APONTAMENTOS"
        resultado["status_operacional"] = "VAZIO_AGUARDANDO_APONTAMENTOS"
    else:
        resultado["status_mes"] = "EM_ANDAMENTO"
        resultado["status_operacional"] = "EM_ANDAMENTO"

    if incluir_mes_anterior and mes_detectado:
        ano_ant, mes_ant = calcular_mes_ano_anterior(ano_detectado, mes_detectado)
        kpi_ant = localizar_kpi_oficial(setor_id, ano_alvo=ano_ant, mes_alvo=mes_ant)

        if (not kpi_ant or not kpi_ant.exists()) and kpi_path and kpi_path.parent.exists():
            for item in kpi_path.parent.iterdir():
                if item.is_file() and not item.name.startswith("~$") and item.suffix.lower() == ".xlsx":
                    if eh_diretorio_legado_ou_ignorado(item, ano_vigente=ano_ant):
                        continue
                    n_norm = normalizar_texto(item.name)
                    if "apontamento" in n_norm:
                        continue
                    if not (_eh_kpi_do_setor(n_norm, setor_id, item.suffix.lower()) or (caminho_kpi_custom is not None and "kpi" in n_norm)):
                        continue
                    ano_item, m_item, _ = _extrair_chave_ano_mes_pasta(item.name, ano_fallback=None)
                    if ano_item != 0 and ano_item != ano_ant:
                        continue
                    if m_item == mes_ant:
                        kpi_ant = item
                        break

        if kpi_ant and kpi_ant.exists() and (kpi_path is None or kpi_ant.resolve() != kpi_path.resolve()):
            diag_ant = diagnosticar_kpi_setor(
                setor_id,
                caminho_kpi_custom=kpi_ant,
                mes_alvo=mes_ant,
                ano_alvo=ano_ant,
                incluir_mes_anterior=False,
            )
            resultado["mes_anterior"] = diag_ant
            resultado["transicao_ativa"] = (total_linhas_validas == 0)

    return resultado


def exibir_relatorio_validador(diag: Dict[str, Any]) -> None:
    """Imprime no terminal o relatório integrado com status do Mês Anterior e Mês Atual."""
    print("\n" + "=" * 88)
    print(f"  VALIDADOR DE DIA & TRANSIÇÃO DE MÊS — {diag['nome_setor']}")
    print("=" * 88)

    diag_ant = diag.get("mes_anterior")
    if diag_ant and diag_ant.get("kpi_encontrado") and diag_ant.get("caminho_kpi"):
        ult_d_ant = diag_ant.get("ultimo_dia_kpi")
        pri_d_ant = diag_ant.get("primeiro_dia_kpi")
        dias_kpi_ant = len(diag_ant.get("dias_no_kpi", {}))
        linhas_ant = diag_ant.get("total_linhas_com_dados", 0)
        ult_l_ant = diag_ant.get("ultima_linha_preenchida")
        pend_ant = diag_ant.get("dias_pendentes", [])

        print("  📦 STATUS DO MÊS ANTERIOR (CONSOLIDADO):")
        print(f"       KPI Anterior:            {diag_ant['caminho_kpi'].name}")
        if pri_d_ant and ult_d_ant:
            print(f"       Período Consolidado:     {pri_d_ant.strftime('%d/%m/%Y')} até {ult_d_ant.strftime('%d/%m/%Y')}")
        print(f"       Total de Dias no KPI:    {dias_kpi_ant} dias ({linhas_ant} linhas preenchidas)")
        print(f"       Última Linha Preenchida: Linha {ult_l_ant}")
        if not pend_ant:
            print("       Situação de Fechamento:  ✅ FECHAMENTO COMPLETO (sem apontamentos pendentes)")
        else:
            print(f"       Situação de Fechamento:  ⚠️ ATENÇÃO: {len(pend_ant)} dia(s) pendente(s) na pasta!")
        print("  " + "-" * 84)

    print("  📅 STATUS DO MÊS ATUAL:")
    if not diag.get("kpi_encontrado"):
        print("     ❌ KPI Oficial do mês atual não encontrado nas pastas configuradas!")
        print("=" * 88)
        return

    print(f"       KPI Atual:               {diag['caminho_kpi'].name}")
    print(f"       Aba de Destino:          '{diag['aba_kpi']}' (Cabeçalho na Linha {diag['linha_cabecalho']})")

    if diag.get("total_linhas_com_dados", 0) == 0:
        print("       Situação do Mês:         ⏳ VAZIO / AGUARDANDO PRIMEIROS APONTAMENTOS DO MÊS")
        print(f"       Total de Dias no KPI:    0 dias (0 linhas preenchidas)")
        print(f"       Linha de Cabeçalho:      Linha {diag['linha_cabecalho']}")
        print(f"       PRÓXIMA LINHA ENTRADA:   Linha {diag['proxima_linha_livre']} (Pronto para receber o 1º apontamento)")
    else:
        ult_d = diag.get("ultimo_dia_kpi")
        pri_d = diag.get("primeiro_dia_kpi")
        dias_kpi = diag.get("dias_no_kpi", {})
        if ult_d and ult_d in dias_kpi:
            info_ult = dias_kpi[ult_d]
            if pri_d and ult_d:
                print(f"       Período Atual no KPI:    {pri_d.strftime('%d/%m/%Y')} até {ult_d.strftime('%d/%m/%Y')}")
            print(f"       Total de Dias no KPI:    {len(dias_kpi)} dias ({diag['total_linhas_com_dados']} linhas preenchidas)")
            print(f"       ÚLTIMO DIA APONTADO:     {ult_d.strftime('%d/%m/%Y')} (Linhas {info_ult['linha_inicio']}..{info_ult['linha_fim']})")
        print(f"       ÚLTIMA LINHA PREENCHIDA: Linha {diag['ultima_linha_preenchida']}")
        print(f"       PRÓXIMA LINHA ENTRADA:   Linha {diag['proxima_linha_livre']}")

    print("  " + "-" * 84)
    if diag.get("dias_apontamentos"):
        print("  📂 Apontamentos (.xlsm) identificados na pasta do mês atual:")
        for dt_a, info_a in diag["dias_apontamentos"].items():
            nomes = ", ".join(p.name for p in info_a["arquivos"])
            if info_a["ja_no_kpi"]:
                faixa = diag["dias_no_kpi"].get(dt_a, {})
                status_str = f"✅ CONSTA NO KPI (Linhas {faixa.get('linha_inicio', '?')}..{faixa.get('linha_fim', '?')})"
            else:
                status_str = f"⏳ PENDENTE -> Entrará na Linha {diag['proxima_linha_livre']}"
            print(f"       Dia {dt_a.strftime('%d/%m/%Y')}: {status_str} [{nomes}]")
    else:
        print("  ⚠️ Nenhum arquivo de apontamento (.xlsm) encontrado na pasta para este setor no mês atual.")
    print("=" * 88)


# ==============================================================================
# 3. GRAVADOR BLINDADO SEMPRE EMBAIXO DA ÚLTIMA LINHA + GERADOR DE MASTER
# ==============================================================================
def verificar_lock_excel(caminho_kpi: Path) -> Optional[str]:
    """Verifica se existe arquivo temporário de bloqueio (~$*.xlsx) e retorna o dono."""
    arq_lock = caminho_kpi.parent / f"~${caminho_kpi.name}"
    if not arq_lock.exists():
        return None
    try:
        dados = arq_lock.read_bytes()
        if dados and 0 < dados[0] < len(dados):
            usuario = dados[1 : 1 + dados[0]].decode("latin-1", errors="ignore").strip()
            if usuario:
                return usuario
    except Exception:
        pass
    return "Outro usuário na rede"


def _calcular_faixas_colunas_livres(total_colunas: int, cols_protegidas: Set[int]) -> List[Tuple[int, int]]:
    """
    Divide as colunas (1..total_colunas) em faixas contíguas que NÃO são protegidas por fórmula.
    Exemplo: se a coluna 5 é fórmula protegida, retorna [(1, 4), (6, ...)].
    Permite gravar via Excel COM em alta velocidade preservando 100% das fórmulas e gráficos!
    """
    faixas: List[Tuple[int, int]] = []
    ini: Optional[int] = None
    for c in range(1, total_colunas + 1):
        if c not in cols_protegidas:
            if ini is None:
                ini = c
        else:
            if ini is not None:
                faixas.append((ini, c - 1))
                ini = None
    if ini is not None:
        faixas.append((ini, total_colunas))
    return faixas


def salvar_planilha_master(
    registros: List[Dict[str, Any]],
    caminho_master: Path,
    titulo_aba: str,
    total_colunas_dados: int,
) -> Path:
    """
    Gera ou atualiza a planilha Master consolidada do setor de forma ACUMULATIVA no mês.
    Salva primeiro em arquivo temporário local (C:\\Temp) e depois copia para a rede V:\\.
    """
    caminho_master.parent.mkdir(parents=True, exist_ok=True)
    aba_limpa = titulo_aba[:31]
    datas_lote = {reg["data"].strftime("%Y-%m-%d") for reg in registros}

    linhas_antigas_preservadas: List[List[Any]] = []
    if caminho_master.exists():
        try:
            linhas_existentes = ler_linhas_planilha(caminho_master, aba_limpa)
            for row in linhas_existentes[1:]:
                if not row or not row[0]:
                    continue
                dt_obj = converter_valor_para_data(row[0])
                dt_iso = dt_obj.strftime("%Y-%m-%d") if dt_obj else str(row[0]).strip()[:10]
                if dt_iso not in datas_lote:
                    linhas_antigas_preservadas.append(list(row))
        except Exception:
            pass

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = aba_limpa

    cabecalho = [
        "DATA_APONTAMENTO",
        "ORIGEM_PLANTA",
        "ARQUIVO_ORIGEM",
        "ABA_ORIGEM",
        "LINHA_ORIGEM",
    ] + [f"COL_KPI_{i+2}" for i in range(total_colunas_dados)]
    ws.append(cabecalho)

    fill_hdr = PatternFill("solid", fgColor="1F4E78")
    font_hdr = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, len(cabecalho) + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = fill_hdr
        c.font = font_hdr

    novas_linhas: List[List[Any]] = []
    for reg in registros:
        novas_linhas.append([
            reg["data"].strftime("%Y-%m-%d"),
            reg.get("planta", ""),
            reg.get("arquivo", ""),
            reg.get("aba", ""),
            reg.get("linha_excel", 0),
        ] + reg["valores_kpi"][:total_colunas_dados])

    todas_linhas = linhas_antigas_preservadas + novas_linhas
    todas_linhas.sort(
        key=lambda r: (
            str(r[0]),
            str(r[1]),
            str(r[3]),
            int(r[4]) if (len(r) > 4 and str(r[4]).isdigit()) else 0,
        )
    )

    resumo_agrupado: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for linha in todas_linhas:
        ws.append(linha)
        dt_str = str(linha[0]) if len(linha) > 0 else ""
        planta_str = str(linha[1]) if len(linha) > 1 else ""
        arq_str = str(linha[2]) if len(linha) > 2 else ""
        qtd_val = 0.0
        if len(linha) > 11 and isinstance(linha[11], (int, float)):
            qtd_val = float(linha[11])
        chave_res = (dt_str, planta_str)
        if chave_res not in resumo_agrupado:
            resumo_agrupado[chave_res] = {"arquivo": arq_str, "linhas": 0, "pecas": 0.0}
        resumo_agrupado[chave_res]["linhas"] += 1
        resumo_agrupado[chave_res]["pecas"] += qtd_val

    ws_res = wb.create_sheet(title="RESUMO_VALIDACAO")
    hdr_res = ["DATA_APONTAMENTO", "ORIGEM_PLANTA", "ARQUIVO_ORIGEM", "QTD_LINHAS", "TOTAL_PECAS_PRODUZIDAS"]
    ws_res.append(hdr_res)
    for col_idx in range(1, len(hdr_res) + 1):
        c = ws_res.cell(row=1, column=col_idx)
        c.fill = fill_hdr
        c.font = font_hdr

    for (dt_s, pl_s), info_r in sorted(resumo_agrupado.items()):
        ws_res.append([dt_s, pl_s, info_r["arquivo"], info_r["linhas"], round(info_r["pecas"], 2)])

    # Salva primeiro no disco local temporário e depois copia para a rede V:\
    fd_tmp, nome_tmp_local = tempfile.mkstemp(suffix=".xlsx", prefix="master_pecval_")
    os.close(fd_tmp)
    caminho_tmp_local = Path(nome_tmp_local)
    try:
        wb.save(caminho_tmp_local)
        wb.close()
        shutil.copy2(caminho_tmp_local, caminho_master)
    except Exception:
        pass
    finally:
        if caminho_tmp_local.exists():
            caminho_tmp_local.unlink(missing_ok=True)

    return caminho_master


def inserir_abaixo_ultima_linha_kpi(
    setor_id: str,
    registros: List[Dict[str, Any]],
    modo_oficial: bool = False,
    colunas_formula_protegidas: Optional[Set[int]] = None,
    permitir_atualizar_mesmo_dia: bool = False,
    caminho_kpi_custom: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Grava os registros extraídos SEMPRE EMBAIXO DA ÚLTIMA LINHA PREENCHIDA do KPI.
      - Usa prioritariamente o Excel COM nativo no Windows para TODOS os setores (inclusive Pintura Manual),
        preservando 100% dos botões/shapes/validações de dados e protegendo as colunas de fórmula!
    """
    if not registros:
        return {"ok": False, "motivo": "Lista de registros vazia."}

    diag = diagnosticar_kpi_setor(setor_id, caminho_kpi_custom=caminho_kpi_custom, incluir_mes_anterior=False)
    if not diag["kpi_encontrado"]:
        return {"ok": False, "motivo": f"KPI Oficial do setor {setor_id} não encontrado."}

    kpi_oficial: Path = diag["caminho_kpi"]
    data_lote: dt.date = registros[0]["data"]
    cfg = ESTRUTURA_SETORES[setor_id]
    total_colunas = cfg["total_colunas_kpi"]
    cols_prot: Set[int] = set(
        colunas_formula_protegidas
        if colunas_formula_protegidas is not None
        else (cfg.get("colunas_formula_protegidas") or set())
    )

    # 1. Validação de Dia Duplicado no KPI
    if data_lote in diag["dias_no_kpi"] and not permitir_atualizar_mesmo_dia:
        faixa = diag["dias_no_kpi"][data_lote]
        return {
            "ok": False,
            "bloqueado_duplicidade": True,
            "motivo": (
                f"O dia {data_lote.strftime('%d/%m/%Y')} JÁ ESTÁ APONTADO no KPI "
                f"(Linhas {faixa['linha_inicio']} a {faixa['linha_fim']}). "
                f"Operação bloqueada para evitar duplicidade!"
            ),
            "linha_inicio": faixa["linha_inicio"],
            "linha_fim": faixa["linha_fim"],
        }

    # 2. Define o arquivo de destino (KPI Teste vs. KPI Oficial)
    if modo_oficial:
        usuario_lock = verificar_lock_excel(kpi_oficial)
        if usuario_lock:
            return {
                "ok": False,
                "motivo": f"KPI Oficial está aberto no Excel por: '{usuario_lock}'. Feche antes de gravar.",
            }
        PASTA_BACKUPS.mkdir(parents=True, exist_ok=True)
        carimbo = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_backup = PASTA_BACKUPS / f"{kpi_oficial.stem}_BACKUP_{carimbo}{kpi_oficial.suffix}"
        shutil.copy2(kpi_oficial, caminho_backup)
        kpi_destino = kpi_oficial
    else:
        kpi_destino = obter_caminho_kpi_teste(kpi_oficial)
        shutil.copy2(kpi_oficial, kpi_destino)

    # 3. Define a linha exata de início
    if data_lote in diag["dias_no_kpi"] and permitir_atualizar_mesmo_dia:
        linha_inicio = diag["dias_no_kpi"][data_lote]["linha_inicio"]
    else:
        linha_inicio = diag["proxima_linha_livre"]

    linha_fim = linha_inicio + len(registros) - 1

    # 4. Gravação em Staging Temporário Local (C:\Temp)
    temp_dir = Path(tempfile.mkdtemp(prefix=f"pecval_{setor_id.lower()}_"))
    temp_kpi = temp_dir / kpi_destino.name

    try:
        shutil.copy2(kpi_destino, temp_kpi)
        formulas_preservadas = 0
        gravou_via_com = False

        # Prepara a matriz completa de linhas (Coluna 1 = Data ISO, Colunas 2..N = Dados)
        matriz_completa: List[List[Any]] = []
        for reg in registros:
            vals = list(reg["valores_kpi"][: total_colunas - 1])
            if len(vals) < total_colunas - 1:
                vals.extend([""] * (total_colunas - 1 - len(vals)))
            linha_completa = [reg["data"].strftime("%Y-%m-%d")] + [
                (v if v != "" else None) for v in vals
            ]
            matriz_completa.append(linha_completa)

        # 4A. MOTOR PRINCIPAL NO WINDOWS: EXCEL COM (Suporta qualquer setor, inclusive Pintura Manual com cols_prot!)
        if TEM_WIN32COM:
            excel_app = None
            wb_com = None
            try:
                excel_app = win32com.client.DispatchEx("Excel.Application")
                excel_app.Visible = False
                excel_app.DisplayAlerts = False
                excel_app.ScreenUpdating = False
                excel_app.EnableEvents = False

                wb_com = excel_app.Workbooks.Open(str(temp_kpi.resolve()), UpdateLinks=0, ReadOnly=False)
                nomes_abas = [wb_com.Sheets(i).Name for i in range(1, wb_com.Sheets.Count + 1)]
                aba_real = resolver_nome_aba(nomes_abas, cfg["aba_kpi"])

                if aba_real:
                    ws_com = wb_com.Sheets(aba_real)
                    faixas_livres = _calcular_faixas_colunas_livres(total_colunas, cols_prot)

                    # Escreve apenas nos blocos de colunas que NÃO têm fórmulas protegidas
                    for col_ini, col_fim in faixas_livres:
                        sub_matriz = [
                            tuple(linha[col_ini - 1 : col_fim])
                            for linha in matriz_completa
                        ]
                        cel_ini = ws_com.Cells(linha_inicio, col_ini)
                        cel_fim = ws_com.Cells(linha_fim, col_fim)
                        ws_com.Range(cel_ini, cel_fim).Value = tuple(sub_matriz)

                    # Nas colunas de fórmula protegidas, propaga a fórmula da linha modelo caso alguma linha nova esteja vazia
                    if cols_prot:
                        linha_modelo_f = cfg["linha_cabecalho_kpi"] + 1
                        for col_f in sorted(cols_prot):
                            try:
                                cel_mod = ws_com.Cells(linha_modelo_f, col_f)
                                form_str = str(cel_mod.Formula or "")
                                if form_str.startswith("=") and linha_fim > linha_modelo_f:
                                    rng_f = ws_com.Range(
                                        ws_com.Cells(linha_modelo_f, col_f),
                                        ws_com.Cells(linha_fim, col_f),
                                    )
                                    cel_mod.Copy(rng_f)
                            except Exception:
                                pass
                        formulas_preservadas = len(cols_prot) * len(registros)

                    wb_com.Save()
                    gravou_via_com = True
            except Exception:
                gravou_via_com = False
            finally:
                if wb_com is not None:
                    try:
                        wb_com.Close(SaveChanges=False)
                    except Exception:
                        pass
                if excel_app is not None:
                    try:
                        excel_app.Quit()
                    except Exception:
                        pass

        # 4B. Fallback OpenPyXL (Apenas se o PC não tiver Microsoft Excel instalado)
        if not gravou_via_com:
            wb = openpyxl.load_workbook(temp_kpi, data_only=False, keep_vba=False)
            aba_real = resolver_nome_aba(wb.sheetnames, cfg["aba_kpi"])
            if not aba_real:
                wb.close()
                return {"ok": False, "motivo": f"Aba '{cfg['aba_kpi']}' não encontrada em {kpi_destino.name}."}

            ws = wb[aba_real]
            for offset, reg in enumerate(registros):
                r_alvo = linha_inicio + offset
                valores_kpi = reg["valores_kpi"]

                cel_data = ws.cell(row=r_alvo, column=1)
                if not isinstance(cel_data, MergedCell):
                    cel_data.value = reg["data"]
                    cel_data.number_format = "dd/mm/yyyy"

                for idx_val in range(total_colunas - 1):
                    col_excel = idx_val + 2
                    cel = ws.cell(row=r_alvo, column=col_excel)
                    if isinstance(cel, MergedCell):
                        continue

                    if cols_prot and col_excel in cols_prot:
                        if cel.value is not None and str(cel.value).startswith("="):
                            formulas_preservadas += 1
                            continue

                    val_bruto = valores_kpi[idx_val] if idx_val < len(valores_kpi) else ""
                    cel.value = val_bruto if val_bruto != "" else None

            wb.save(temp_kpi)
            wb.close()

        # 5. Transferência Segura para a Rede V:\ (Sem os.fsync que causa Errno 9 no Windows SMB)
        temp_no_destino = kpi_destino.parent / f".tmp_{uuid.uuid4().hex[:6]}_{kpi_destino.name}"
        ultimo_erro = None

        for tentativa in range(4):
            try:
                shutil.copy2(temp_kpi, temp_no_destino)
                os.replace(temp_no_destino, kpi_destino)
                ultimo_erro = None
                break
            except Exception as exc_rep:
                ultimo_erro = exc_rep
                # Fallback direto se o SMB bloquear o rename do arquivo temporário oculto
                try:
                    shutil.copy2(temp_kpi, kpi_destino)
                    if temp_no_destino.exists():
                        temp_no_destino.unlink(missing_ok=True)
                    ultimo_erro = None
                    break
                except Exception as exc_copy:
                    ultimo_erro = exc_copy
                    time.sleep(0.6 * (tentativa + 1))

        if temp_no_destino.exists():
            try:
                temp_no_destino.unlink(missing_ok=True)
            except Exception:
                pass

        if ultimo_erro is not None:
            return {
                "ok": False,
                "motivo": f"Não foi possível gravar em '{kpi_destino.name}' (verifique se está aberto no Excel): {ultimo_erro}",
            }

        _CACHE_ZIP_DATAS.clear()

        return {
            "ok": True,
            "modo": "OFICIAL" if modo_oficial else "TESTE",
            "motor_gravacao": "EXCEL_COM" if gravou_via_com else "OPENPYXL",
            "arquivo_gravado": kpi_destino,
            "ultimo_dia_antes": diag["ultimo_dia_kpi"],
            "ultima_linha_antes": diag["ultima_linha_preenchida"],
            "data_inserida": data_lote,
            "linha_inicio": linha_inicio,
            "linha_fim": linha_fim,
            "linhas_inseridas": len(registros),
            "formulas_preservadas": formulas_preservadas,
        }
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


if __name__ == "__main__":
    mes_cli = int(sys.argv[1]) if (len(sys.argv) > 1 and sys.argv[1].isdigit()) else None
    for s_id in ["INJECAO", "MONTAGEM", "PINTURA_MANUAL", "PINTURA_AUTOMATICA"]:
        diag_setor = diagnosticar_kpi_setor(s_id, mes_alvo=mes_cli)
        exibir_relatorio_validador(diag_setor)