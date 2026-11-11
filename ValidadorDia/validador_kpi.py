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
import threading
import datetime as dt
import xml.etree.ElementTree as ET
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
from openpyxl.utils import get_column_letter

try:
    from python_calamine import CalamineWorkbook
    TEM_CALAMINE = True
except ImportError:
    TEM_CALAMINE = False

try:
    import pythoncom
    TEM_PYTHONCOM = True
except ImportError:
    TEM_PYTHONCOM = False

try:
    import win32com.client
    TEM_WIN32COM = True
except ImportError:
    TEM_WIN32COM = False

try:
    import wolfxl
    TEM_WOLFXL = True
except ImportError:
    TEM_WOLFXL = False

_COM_EXCEL_LOCK = threading.Lock()

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

try:
    from Caminhos.config_caminho import PASTA_OVERRIDE_KPIS
except ImportError:
    PASTA_OVERRIDE_KPIS = None

try:
    from Caminhos.config_caminho import listar_backups_kpi
except ImportError:
    def listar_backups_kpi(setor_id: Optional[str] = None, limite: int = 30) -> List[Dict[str, Any]]:
        """Fallback embutido caso config_caminho.py ainda não tenha listar_backups_kpi sincronizado."""
        if not PASTA_BACKUPS.exists():
            return []
        try:
            arquivos = sorted(
                [p for p in PASTA_BACKUPS.glob("*.xlsx") if not p.name.startswith("~$")],
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
        except Exception:
            return []
        resultado: List[Dict[str, Any]] = []
        for arq in arquivos:
            nome_norm = normalizar_texto(arq.name)
            setor_detectado = "DESCONHECIDO"
            if "INJECAO" in nome_norm or "INJECTION" in nome_norm:
                setor_detectado = "INJECAO"
            elif "MONTAGEM" in nome_norm or "ASSEMBLY" in nome_norm:
                setor_detectado = "MONTAGEM"
            elif "MANUAL" in nome_norm or "BANHO" in nome_norm:
                setor_detectado = "PINTURA_MANUAL"
            elif "AUTOMATICA" in nome_norm or "ROBO" in nome_norm or "PAINTING" in nome_norm:
                setor_detectado = "PINTURA_AUTO"
            if setor_id and setor_detectado != setor_id:
                continue
            try:
                st = arq.stat()
                dt_mod = dt.datetime.fromtimestamp(st.st_mtime)
                tamanho_mb = round(st.st_size / (1024 * 1024), 2)
            except Exception:
                dt_mod = dt.datetime.now()
                tamanho_mb = 0.0
            resultado.append({
                "nome": arq.name,
                "caminho": str(arq),
                "setor_id": setor_detectado,
                "data_hora": dt_mod.strftime("%d/%m/%Y %H:%M:%S"),
                "tamanho_mb": tamanho_mb,
            })
            if len(resultado) >= limite:
                break
        return resultado


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
    """Limpa valores vindos do apontamento antes de gravar no Excel, preservando tempos e paradas."""
    if val is None:
        return ""
    if isinstance(val, dt.timedelta):
        total_sec = int(val.total_seconds())
        if total_sec < 0:
            return "00:00:00"
        horas = total_sec // 3600
        minutos = (total_sec % 3600) // 60
        segundos = total_sec % 60
        # Poka-Yoke: Se horas >= 35088 (1462 dias decorrentes da época 1904 do Excel),
        # subtrai a época para obter a duração pura da parada (ex: 35090h -> 2h)
        if horas >= 35088:
            horas -= 35088
        return f"{horas:02d}:{minutos:02d}:{segundos:02d}"
    if isinstance(val, dt.time):
        return val.strftime("%H:%M:%S")
    if isinstance(val, dt.datetime):
        # No Excel, horários/paradas/durações com ano base menor que 1950 (1899, 1900 ou 1904)
        # são HORÁRIOS PUROS e JAMAIS devem ser gravados com ano de calendário (o que causava 35.100 horas)!
        if val.year < 1950:
            if val.year == 1904 and (val.month == 1 and val.day == 1):
                return val.strftime("%H:%M:%S")
            elif val.year == 1904:
                delta = val - dt.datetime(1904, 1, 1)
                total_sec = max(0, int(round(delta.total_seconds())))
                horas = total_sec // 3600
                minutos = (total_sec % 3600) // 60
                segundos = total_sec % 60
                if horas >= 35088:
                    horas -= 35088
                return f"{horas:02d}:{minutos:02d}:{segundos:02d}"
            elif val.year in (1899, 1900):
                if (val.year == 1899 and val.month == 12 and val.day in (30, 31)) or (val.year == 1900 and val.month == 1 and val.day == 1):
                    return val.strftime("%H:%M:%S")
                delta = val - dt.datetime(1899, 12, 30)
                total_sec = max(0, int(round(delta.total_seconds())))
                horas = total_sec // 3600
                minutos = (total_sec % 3600) // 60
                segundos = total_sec % 60
                if horas >= 35088:
                    horas -= 35088
                return f"{horas:02d}:{minutos:02d}:{segundos:02d}"
            else:
                return val.strftime("%H:%M:%S")
        return val.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(val, dt.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        # Se for número com época 1904 (1460 <= val <= 1465 com parte decimal):
        # 1462 dias = 01/01/1904 e a parte decimal é o horário puro!
        if 1460.0 <= val <= 1465.0:
            frac = float(val) - int(val)
            if frac > 0.000001:
                total_sec = int(round(frac * 86400))
                horas = total_sec // 3600
                minutos = (total_sec % 3600) // 60
                segundos = total_sec % 60
                return f"{horas:02d}:{minutos:02d}:{segundos:02d}"
    if isinstance(val, str):
        s = val.strip()
        if s.startswith("#"):
            return ""
        # Intercepta strings com data base 1899, 1900 ou 1904 (ex: '1904-01-01 18:00:00' -> '18:00:00')
        m_t = re.match(r"^(?:1899-12-30|1900-01-01|1904-01-01)[T\s]+(\d{1,2}:\d{2}(?::\d{2})?)", s)
        if m_t:
            return m_t.group(1)
        # Intercepta strings com formato de duração com mais de 35000 horas (ex: '35090:10' ou '35090:10:00')
        m_h = re.match(r"^(\d{4,6}):(\d{2})(?::(\d{2}))?$", s)
        if m_h:
            h_val = int(m_h.group(1))
            m_val = int(m_h.group(2))
            s_val = int(m_h.group(3)) if m_h.group(3) else 0
            if h_val >= 35088:
                h_val -= 35088
                return f"{h_val:02d}:{m_val:02d}:{s_val:02d}"
        return s
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
    Lê todas as linhas de uma aba usando o motor Rust Calamine (com skip_empty_area=False
    para JAMAIS deslocar colunas quando a Coluna A estiver vazia) ou OpenPyXL em modo
    read_only como fallback seguro.
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
                    # CRÍTICO: skip_empty_area=False impede que colunas vazias à esquerda (ex: Coluna A)
                    # desloquem todos os índices das colunas de Paradas, Horas Não Programadas e Tempos!
                    rows = sh.to_python(skip_empty_area=False)
                    if max_col is not None:
                        rows = [list(r[:max_col]) for r in rows]
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


def _converter_valor_para_xml_celula(val: Any, col_idx_1based: int) -> Tuple[str, str]:
    """
    Converte um valor Python para (tipo_xml, valor_texto) nativo do OpenXML (ECMA-376):
      - Datas (dt.date / ISO YYYY-MM-DD na Coluna 1) -> Serial OADate numérico ('n')
      - Horários ("HH:MM" / "HH:MM:SS") -> Fração do dia OADate ('n') para somar perfeitamente no KPI!
      - Números (int / float) -> ('n', str(val))
      - Strings -> ('inlineStr', texto_escapado)
    """
    if val is None or val == "":
        return ("empty", "")

    if isinstance(val, bool):
        return ("b", "1" if val else "0")

    if isinstance(val, (int, float)):
        return ("n", str(val))

    if isinstance(val, dt.datetime):
        if val.year in (1899, 1900):
            delta = val - dt.datetime(1899, 12, 30)
            return ("n", f"{delta.total_seconds() / 86400.0:.10f}")
        delta = val - dt.datetime(1899, 12, 30)
        return ("n", f"{delta.days + (delta.seconds / 86400.0):.10f}")

    if isinstance(val, dt.date):
        serial = (val - dt.date(1899, 12, 30)).days
        return ("n", str(serial))

    if isinstance(val, dt.time):
        frac = (val.hour * 3600 + val.minute * 60 + val.second) / 86400.0
        return ("n", f"{frac:.10f}")

    s = str(val).strip()
    if not s or s.startswith("#"):
        return ("empty", "")

    # Data ISO na Coluna A (ex: '2026-10-30') -> Serial numérico do Excel (mantém formatação dd/mm/yyyy original)
    if col_idx_1based == 1:
        dt_parsed = converter_valor_para_data(s)
        if dt_parsed:
            serial = (dt_parsed - dt.date(1899, 12, 30)).days
            return ("n", str(serial))

    # Horário ou duração (ex: '08:30', '01:15:00', '0:00') -> Fração numérica do dia no Excel!
    m_time = re.match(r"^(\d{1,3}):(\d{2})(?::(\d{2}))?$", s)
    if m_time:
        hh = int(m_time.group(1))
        mm = int(m_time.group(2))
        ss = int(m_time.group(3) or 0)
        if mm < 60 and ss < 60:
            frac = (hh * 3600 + mm * 60 + ss) / 86400.0
            return ("n", f"{frac:.10f}")

    return ("inlineStr", s)


def _gravar_cirurgico_openxml_zip(
    caminho_xlsx: Path,
    nome_aba_esperada: str,
    linha_inicio: int,
    matriz_completa: List[List[Any]],
    cols_protegidas: Set[int],
    linha_fim_antiga_limpar: int = 0,
) -> Tuple[bool, int, str]:
    """
    MOTOR CIRÚRGICO OPENXML (Inspirado em editpyxl / xlsxedit / wolfxl):
    Edita APENAS o XML da aba alvo ('APONTAR AQUI') dentro do pacote ZIP (.xlsx),
    mantendo 100% de todas as outras abas, Tabelas Dinâmicas, Gráficos, Slicers,
    Drawings, Botões, Estilos e CalcChain INTACTOS (Zero corrupção e Zero 'Arquivo Restaurado')!
    Além disso, preserva fórmulas existentes na planilha quando cols_protegidas está ativo
    ou quando o valor de entrada é vazio e a célula já possui fórmula (<f>).
    """
    try:
        with zipfile.ZipFile(caminho_xlsx, "r") as zin:
            wb_xml = zin.read("xl/workbook.xml").decode("utf-8", errors="ignore")
            rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8", errors="ignore")

            rid_target: Dict[str, str] = {}
            for r in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_xml):
                rid_target[r[0]] = r[1]
            for r in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels_xml):
                rid_target[r[1]] = r[0]

            target_sheet = None
            alvo_norm = nome_aba_esperada.strip().upper()
            for s in re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml):
                if s[0].strip().upper() == alvo_norm:
                    target_sheet = rid_target.get(s[1])
                    break

            if not target_sheet:
                return (False, 0, f"Aba '{nome_aba_esperada}' não localizada no XML.")

            ts = target_sheet.lstrip("/")
            caminho_sheet = ts if ts.startswith("xl/") else f"xl/{ts}"
            if caminho_sheet not in zin.namelist():
                return (False, 0, f"XML '{caminho_sheet}' não encontrado no arquivo.")

            sheet_bytes = zin.read(caminho_sheet)
            todos_arquivos = {nome: zin.read(nome) for nome in zin.namelist()}

        ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ET.register_namespace("", ns_main)
        ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
        ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
        ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
        ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
        ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

        root = ET.fromstring(sheet_bytes)
        tag_sheet_data = f"{{{ns_main}}}sheetData"
        tag_row = f"{{{ns_main}}}row"
        tag_c = f"{{{ns_main}}}c"
        tag_f = f"{{{ns_main}}}f"
        tag_v = f"{{{ns_main}}}v"
        tag_is = f"{{{ns_main}}}is"
        tag_t = f"{{{ns_main}}}t"

        sheet_data = root.find(tag_sheet_data)
        if sheet_data is None:
            return (False, 0, "Elemento <sheetData> não encontrado na aba.")

        mapa_linhas: Dict[int, ET.Element] = {}
        estilo_padrao_col: Dict[int, str] = {}

        for row_el in sheet_data.findall(tag_row):
            r_attr = row_el.get("r")
            if r_attr and r_attr.isdigit():
                r_num = int(r_attr)
                mapa_linhas[r_num] = row_el
                if r_num >= 3 and len(estilo_padrao_col) < 150:
                    for c_el in row_el.findall(tag_c):
                        ref = c_el.get("r", "")
                        s_attr = c_el.get("s")
                        if s_attr and ref:
                            m_col = re.match(r"^([A-Z]+)", ref)
                            if m_col:
                                col_let = m_col.group(1)
                                col_idx = 0
                                for ch in col_let:
                                    col_idx = col_idx * 26 + (ord(ch) - ord("A") + 1)
                                if col_idx not in estilo_padrao_col:
                                    estilo_padrao_col[col_idx] = s_attr

        formulas_preservadas = 0
        linha_fim = linha_inicio + len(matriz_completa) - 1

        for offset, valores_linha in enumerate(matriz_completa):
            r_alvo = linha_inicio + offset
            row_el = mapa_linhas.get(r_alvo)
            if row_el is None:
                row_el = ET.Element(tag_row, {"r": str(r_alvo)})
                # Insere em ordem crescente de linha no sheetData
                inserido = False
                for idx_child, existing_row in enumerate(list(sheet_data)):
                    r_ex = existing_row.get("r")
                    if r_ex and r_ex.isdigit() and int(r_ex) > r_alvo:
                        sheet_data.insert(idx_child, row_el)
                        inserido = True
                        break
                if not inserido:
                    sheet_data.append(row_el)
                mapa_linhas[r_alvo] = row_el

            cels_existentes: Dict[str, ET.Element] = {}
            for c_el in row_el.findall(tag_c):
                ref = c_el.get("r")
                if ref:
                    cels_existentes[ref] = c_el

            for col_idx_1based, val_bruto in enumerate(valores_linha, start=1):
                col_letra = get_column_letter(col_idx_1based)
                ref_cel = f"{col_letra}{r_alvo}"
                c_el = cels_existentes.get(ref_cel)

                # Verifica se a célula já possui fórmula (<f>) no KPI
                tem_formula = (c_el is not None and c_el.find(tag_f) is not None)
                if tem_formula and (col_idx_1based in cols_protegidas or val_bruto in (None, "")):
                    # Remove apenas o valor em cache (<v>) para forçar o Excel a recalcular a fórmula ao abrir
                    v_old = c_el.find(tag_v)
                    if v_old is not None:
                        c_el.remove(v_old)
                    formulas_preservadas += 1
                    continue

                if col_idx_1based in cols_protegidas:
                    continue

                tipo_xml, val_xml = _converter_valor_para_xml_celula(val_bruto, col_idx_1based)

                if tipo_xml == "empty":
                    if c_el is not None and not tem_formula:
                        for sub in list(c_el):
                            c_el.remove(sub)
                        if "t" in c_el.attrib:
                            del c_el.attrib["t"]
                    continue

                if c_el is None:
                    atributos_c = {"r": ref_cel}
                    if col_idx_1based in estilo_padrao_col:
                        atributos_c["s"] = estilo_padrao_col[col_idx_1based]
                    c_el = ET.SubElement(row_el, tag_c, atributos_c)
                    cels_existentes[ref_cel] = c_el
                else:
                    for sub in list(c_el):
                        c_el.remove(sub)

                if tipo_xml == "n":
                    if "t" in c_el.attrib:
                        del c_el.attrib["t"]
                    v_el = ET.SubElement(c_el, tag_v)
                    v_el.text = val_xml
                elif tipo_xml == "b":
                    c_el.set("t", "b")
                    v_el = ET.SubElement(c_el, tag_v)
                    v_el.text = val_xml
                else:
                    c_el.set("t", "inlineStr")
                    is_el = ET.SubElement(c_el, tag_is)
                    t_el = ET.SubElement(is_el, tag_t)
                    if val_xml.startswith(" ") or val_xml.endswith(" "):
                        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_el.text = val_xml

            # Reordena as células da linha por índice de coluna (exigência do schema OpenXML)
            cels_ordenadas = sorted(
                row_el.findall(tag_c),
                key=lambda el: (
                    len(re.match(r"^([A-Z]+)", el.get("r", "A1")).group(1)),
                    el.get("r", "A1"),
                ),
            )
            for c_el in list(row_el):
                row_el.remove(c_el)
            for c_el in cels_ordenadas:
                row_el.append(c_el)

        # Se for uma atualização de dia onde o lote novo é menor que o antigo, limpa as linhas excedentes daquele dia
        if linha_fim_antiga_limpar > linha_fim:
            for r_limpar in range(linha_fim + 1, linha_fim_antiga_limpar + 1):
                row_old = mapa_linhas.get(r_limpar)
                if row_old is not None:
                    for c_el in row_old.findall(tag_c):
                        if c_el.find(tag_f) is None:
                            for sub in list(c_el):
                                c_el.remove(sub)
                            if "t" in c_el.attrib:
                                del c_el.attrib["t"]

        # Atualiza a tag <dimension ref="A1:EQ..."/> se presente para que tabelas e filtros vejam as novas linhas
        tag_dim = f"{{{ns_main}}}dimension"
        dim_el = root.find(tag_dim)
        if dim_el is not None:
            ref_dim = dim_el.get("ref", "")
            m_dim = re.match(r"^([A-Z]+\d+):([A-Z]+)(\d+)$", ref_dim)
            if m_dim and int(m_dim.group(3)) < linha_fim:
                dim_el.set("ref", f"{m_dim.group(1)}:{m_dim.group(2)}{linha_fim}")

        novo_sheet_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        todos_arquivos[caminho_sheet] = novo_sheet_xml

        # Remove calcChain.xml se fórmulas foram limpas para o Excel reconstruir a cadeia sem aviso de reparo
        if "xl/calcChain.xml" in todos_arquivos:
            # Mantém calcChain intacto se não removemos nenhuma fórmula; se o Excel abrir, ele recalcula
            pass

        with zipfile.ZipFile(caminho_xlsx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for nome_arq, conteudo_bytes in todos_arquivos.items():
                zout.writestr(nome_arq, conteudo_bytes)

        return (True, formulas_preservadas, "OPENXML_CIRURGICO")
    except Exception as exc:
        return (False, 0, str(exc))


def restaurar_ultimo_backup_kpi(
    setor_id: str,
    caminho_backup_especifico: Optional[str] = None,
    mes_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
) -> Dict[str, Any]:
    """
    Restaura com segurança um Backup Oficial da pasta Backups_KPI de volta para o KPI Oficial da rede.
    Cria antes um snapshot de segurança do arquivo atual para que nada seja perdido.
    """
    backups = listar_backups_kpi(setor_id=setor_id, limite=50)
    if not backups and not caminho_backup_especifico:
        return {
            "ok": False,
            "motivo": f"Nenhum arquivo de backup encontrado em {PASTA_BACKUPS} para o setor {setor_id}.",
        }

    arq_backup = Path(caminho_backup_especifico) if caminho_backup_especifico else Path(backups[0]["caminho"])
    if not arq_backup.exists():
        return {"ok": False, "motivo": f"Arquivo de backup não encontrado: {arq_backup}"}

    kpi_oficial = localizar_kpi_oficial(setor_id, ano_alvo=ano_alvo, mes_alvo=mes_alvo)
    if not kpi_oficial:
        return {"ok": False, "motivo": f"KPI Oficial de destino do setor {setor_id} não foi localizado na rede."}

    usuario_lock = verificar_lock_excel(kpi_oficial)
    if usuario_lock:
        return {
            "ok": False,
            "motivo": f"O KPI Oficial '{kpi_oficial.name}' está aberto no Excel por '{usuario_lock}'. Feche-o antes de restaurar.",
        }

    try:
        shutil.copy2(arq_backup, kpi_oficial)
        _CACHE_ZIP_DATAS.clear()
        return {
            "ok": True,
            "setor_id": setor_id,
            "backup_utilizado": arq_backup.name,
            "caminho_backup": str(arq_backup),
            "kpi_restaurado": kpi_oficial.name,
            "caminho_kpi": str(kpi_oficial),
            "mensagem": f"KPI '{kpi_oficial.name}' restaurado com sucesso a partir de '{arq_backup.name}'!",
        }
    except Exception as exc:
        return {"ok": False, "motivo": f"Erro ao restaurar backup: {exc}"}


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
    Arquitetura Tripla de Gravadores Não-Destrutivos (Inspirada em xlwings, wolfxl e editpyxl):
      1º EXCEL_COM_STA: Usa Microsoft Excel via COM com `pythoncom.CoInitialize()` + Lock de Thread
         (garante funcionamento 100% dentro do Hub Web multi-thread sem cair por engano no fallback!).
      2º WOLFXL_MODIFY: Se `wolfxl` estiver instalado, usa o motor Rust `load_workbook(modify=True)`.
      3º OPENXML_CIRURGICO: Edita cirurgicamente apenas as células da aba 'APONTAR AQUI' no XML/ZIP,
         mantendo 100% das abas de Dashboard, Gráficos, Tabelas Dinâmicas e Botões intactos!
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

    # 1B. Proteção contra sobrescrita de dias posteriores quando atualizando um dia no meio do KPI
    linha_fim_antiga_dia = 0
    if data_lote in diag["dias_no_kpi"] and permitir_atualizar_mesmo_dia:
        faixa_dia = diag["dias_no_kpi"][data_lote]
        linha_inicio = faixa_dia["linha_inicio"]
        linha_fim_antiga_dia = faixa_dia["linha_fim"]
        qtd_antiga = faixa_dia["qtd_linhas"]
        dias_posteriores = [d for d in diag["dias_no_kpi"] if d > data_lote]
        if dias_posteriores and len(registros) > qtd_antiga and not TEM_WIN32COM:
            # Se houver dias posteriores e o novo lote for maior sem COM para inserir linhas,
            # bloqueia para jamais sobrescrever os dias seguintes!
            return {
                "ok": False,
                "motivo": (
                    f"O dia {data_lote.strftime('%d/%m/%Y')} ocupava {qtd_antiga} linhas ({linha_inicio}..{linha_fim_antiga_dia}) "
                    f"e possui dias posteriores já lançados abaixo dele. Como o novo lote tem {len(registros)} linhas, "
                    f"a atualização requer o Excel instalado para inserir as {len(registros) - qtd_antiga} linhas extras sem sobrescrever o dia seguinte."
                ),
            }
    else:
        linha_inicio = diag["proxima_linha_livre"]

    linha_fim = linha_inicio + len(registros) - 1

    # 2. Define o arquivo de destino (KPI Teste vs. KPI Oficial vs. KPI Zerado de Homologação)
    caminho_backup: Optional[Path] = None
    eh_teste_zerado = bool(PASTA_OVERRIDE_KPIS or caminho_kpi_custom)

    if modo_oficial or eh_teste_zerado:
        usuario_lock = verificar_lock_excel(kpi_oficial)
        if usuario_lock:
            return {
                "ok": False,
                "motivo": f"KPI está aberto no Excel por: '{usuario_lock}'. Feche antes de gravar.",
            }
        PASTA_BACKUPS.mkdir(parents=True, exist_ok=True)
        carimbo = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        caminho_backup = PASTA_BACKUPS / f"{kpi_oficial.stem}_BACKUP_{carimbo}{kpi_oficial.suffix}"
        shutil.copy2(kpi_oficial, caminho_backup)
        kpi_destino = kpi_oficial
    else:
        kpi_destino = obter_caminho_kpi_teste(kpi_oficial)
        shutil.copy2(kpi_oficial, kpi_destino)

    # 3. Gravação em Staging Temporário Local (C:\Temp)
    temp_dir = Path(tempfile.mkdtemp(prefix=f"pecval_{setor_id.lower()}_"))
    temp_kpi = temp_dir / kpi_destino.name

    try:
        shutil.copy2(kpi_destino, temp_kpi)
        formulas_preservadas = 0
        motor_usado = ""

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

        # 4A. MOTOR 1 (WINDOWS): EXCEL COM NATIVO COM COINITIALIZE E THREAD LOCK (100% FIEL AO EXCEL)
        if TEM_WIN32COM:
            with _COM_EXCEL_LOCK:
                if TEM_PYTHONCOM:
                    try:
                        pythoncom.CoInitialize()
                    except Exception:
                        pass
                excel_app = None
                wb_com = None
                try:
                    excel_app = win32com.client.DispatchEx("Excel.Application")
                    excel_app.Visible = False
                    excel_app.DisplayAlerts = False
                    excel_app.ScreenUpdating = False
                    excel_app.EnableEvents = False
                    try:
                        excel_app.Calculation = -4135  # xlCalculationManual: suspende recálculo pesado durante colagem
                    except Exception:
                        pass

                    wb_com = excel_app.Workbooks.Open(str(temp_kpi.resolve()), UpdateLinks=0, ReadOnly=False)
                    nomes_abas = [wb_com.Sheets(i).Name for i in range(1, wb_com.Sheets.Count + 1)]
                    aba_real = resolver_nome_aba(nomes_abas, cfg["aba_kpi"])

                    if aba_real:
                        ws_com = wb_com.Sheets(aba_real)

                        # Se estiver atualizando um dia existente e houver diferença de linhas, ajusta a faixa
                        if linha_fim_antiga_dia > 0:
                            qtd_antiga = linha_fim_antiga_dia - linha_inicio + 1
                            diff_linhas = len(registros) - qtd_antiga
                            dias_posteriores = [d for d in diag["dias_no_kpi"] if d > data_lote]
                            if diff_linhas > 0 and dias_posteriores:
                                ws_com.Rows(f"{linha_fim_antiga_dia + 1}:{linha_fim_antiga_dia + diff_linhas}").Insert()
                            elif diff_linhas < 0 and dias_posteriores:
                                ws_com.Rows(f"{linha_fim + 1}:{linha_fim_antiga_dia}").Delete()

                        # Identifica dinamicamente colunas que possuem fórmula na linha modelo ou na linha alvo
                        # quando o valor a colar está vazio (evita apagar fórmulas de Horas e Indicadores do KPI!)
                        cols_prot_efetivas = set(cols_prot)
                        linha_modelo_f = cfg["linha_cabecalho_kpi"] + 1
                        try:
                            formulas_mod = ws_com.Range(
                                ws_com.Cells(linha_modelo_f, 1),
                                ws_com.Cells(linha_modelo_f, total_colunas),
                            ).Formula
                            if formulas_mod and len(formulas_mod) > 0:
                                linha_f_tuple = formulas_mod[0]
                                for c_idx, f_val in enumerate(linha_f_tuple, start=1):
                                    if isinstance(f_val, str) and f_val.startswith("="):
                                        # Se todas as linhas do lote tiverem vazio nessa coluna, protege a fórmula!
                                        if all(linha[c_idx - 1] is None for linha in matriz_completa):
                                            cols_prot_efetivas.add(c_idx)
                        except Exception:
                            pass

                        faixas_livres = _calcular_faixas_colunas_livres(total_colunas, cols_prot_efetivas)

                        for col_ini, col_fim in faixas_livres:
                            sub_matriz = [
                                tuple(linha[col_ini - 1 : col_fim])
                                for linha in matriz_completa
                            ]
                            cel_ini = ws_com.Cells(linha_inicio, col_ini)
                            cel_fim = ws_com.Cells(linha_fim, col_fim)
                            ws_com.Range(cel_ini, cel_fim).Value = tuple(sub_matriz)

                        if cols_prot_efetivas:
                            for col_f in sorted(cols_prot_efetivas):
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
                            formulas_preservadas = len(cols_prot_efetivas) * len(registros)
                        try:
                            excel_app.Calculation = -4105  # xlCalculationAutomatic: recalcula uma única vez ao final
                            excel_app.Calculate()
                        except Exception:
                            pass
                        wb_com.Save()
                        motor_usado = "EXCEL_COM"
                except Exception:
                    motor_usado = ""
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
                    if TEM_PYTHONCOM:
                        try:
                            pythoncom.CoUninitialize()
                        except Exception:
                            pass

        # 4B. MOTOR 2 (RUST WOLFXL MODIFY MODE - SE INSTALADO)
        if not motor_usado and TEM_WOLFXL:
            try:
                wb_w = wolfxl.load_workbook(str(temp_kpi), modify=True)
                aba_real = resolver_nome_aba(wb_w.sheetnames, cfg["aba_kpi"])
                if aba_real:
                    ws_w = wb_w[aba_real]
                    for offset, linha_vals in enumerate(matriz_completa):
                        r_alvo = linha_inicio + offset
                        for c_idx, val_c in enumerate(linha_vals, start=1):
                            if c_idx in cols_prot or val_c is None:
                                continue
                            ws_w.cell(row=r_alvo, column=c_idx, value=val_c)
                    wb_w.save(str(temp_kpi))
                    motor_usado = "WOLFXL_RUST"
            except Exception:
                motor_usado = ""

        # 4C. MOTOR 3 (CIRÚRGICO OPENXML ZIP - 100% PRESERVAÇÃO DE ABAS, GRÁFICOS E TABELAS DINÂMICAS)
        if not motor_usado:
            ok_xml, form_xml, info_xml = _gravar_cirurgico_openxml_zip(
                caminho_xlsx=temp_kpi,
                nome_aba_esperada=cfg["aba_kpi"],
                linha_inicio=linha_inicio,
                matriz_completa=matriz_completa,
                cols_protegidas=cols_prot,
                linha_fim_antiga_limpar=linha_fim_antiga_dia,
            )
            if ok_xml:
                motor_usado = info_xml
                formulas_preservadas = form_xml
            else:
                return {
                    "ok": False,
                    "motivo": f"Falha ao gravar de forma segura no KPI ({info_xml}). Operação abortada para proteger as abas e gráficos do arquivo!",
                }

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
            "motor_gravacao": motor_usado,
            "arquivo_gravado": kpi_destino,
            "caminho_backup": str(caminho_backup) if caminho_backup else "",
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