# -*- coding: utf-8 -*-
"""
=============================================================================================
📂 MÓDULO 1: CONFIGURAÇÃO CENTRAL DE CAMINHOS (APONTAMENTOS vs. KPIs vs. MASTERS)
=============================================================================================
Caminho na Estrutura: PECVAL/AutomacaoKPI/Caminhos/config_caminho.py
Unidade: Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)
Diretório Oficial na Rede: V:\\Joao ADM-Prod\\PECVAL\\AutomacaoKPI

Responsabilidades (Clean Architecture / sk-backend):
  1. Separar rigorosamente os caminhos de APONTAMENTOS (leitura) dos caminhos de KPIs (escrita).
  2. Centralizar as Bases Master e os Backups dentro de 'V:\\Joao ADM-Prod\\PECVAL'.
  3. Detectar automaticamente o mês/ano vigente na rede (ex: '10- OUTUBRO - 2026', '11- NOVEMBRO - 2026')
     sem depender de nomes de arquivos fixos.
  4. Possuir fallback transparente para a pasta local caso esteja rodando fora da rede (em casa).
=============================================================================================
"""

import os
import re
import sys
import time
import socket
import datetime as dt
import unicodedata
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

# Blindagem contra UnicodeEncodeError no terminal do Windows (CMD / PowerShell / VS Code)
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# ==============================================================================
# 1. DETECÇÃO RÁPIDA DE REDE CORPORATIVA & VPN (Evita falso negativo em VPN)
# ==============================================================================
def _rede_pecval_acessivel() -> bool:
    """Testa em até 2.5s se o servidor 10.72.110.56 (porta 445 SMB via LAN ou VPN) está acessível."""
    if os.name != "nt":
        return False
    try:
        with socket.create_connection(("10.72.110.56", 445), timeout=2.5):
            return True
    except OSError:
        try:
            return Path(r"\\10.72.110.56\Departamentos").exists()
        except OSError:
            return False

def _drive_v_acessivel() -> bool:
    """Verifica se a unidade de rede V:\\ está montada no Windows."""
    if os.name != "nt":
        return False
    try:
        return Path("V:/").exists()
    except OSError:
        return False

REDE_UNC_ATIVA = _rede_pecval_acessivel()
DRIVE_V_ATIVO  = _drive_v_acessivel()

DIRETORIO_DESTE_ARQUIVO = Path(__file__).resolve().parent          # .../AutomacaoKPI/Caminhos
DIRETORIO_AUTOMACAO_KPI = DIRETORIO_DESTE_ARQUIVO.parent           # .../AutomacaoKPI
DIRETORIO_PECVAL_LOCAL  = DIRETORIO_AUTOMACAO_KPI.parent           # .../PECVAL

PASTA_REDE_JOAO_PECVAL = Path(r"V:\Joao ADM-Prod\PECVAL")

# Define onde salvar as Bases Master e os Backups de segurança
PASTA_BASE_OPERACAO = (
    PASTA_REDE_JOAO_PECVAL
    if (DRIVE_V_ATIVO and PASTA_REDE_JOAO_PECVAL.exists())
    else DIRETORIO_PECVAL_LOCAL
)
PASTA_MASTERS       = PASTA_BASE_OPERACAO / "Masters"
PASTA_BACKUPS       = PASTA_BASE_OPERACAO / "Backups_KPI"
PASTA_TESTES_KPI    = PASTA_BASE_OPERACAO / "KPIs_Teste"

# ==============================================================================
# 2. SEPARAÇÃO ESTRITA: CAMINHOS DE APONTAMENTOS vs. CAMINHOS DOS KPIs OFICIAIS
# ==============================================================================
# Raízes na Rede Corporativa PECVAL (Suporta tanto unidade mapeada V:/ quanto UNC \\10.72.110.56)
RAIZES_APONTAMENTOS_REDE = [
    Path(r"\\10.72.110.56\Departamentos\REPORT DE PRODUCAO\1. INDICADORES PRODUTIVOS PECVAL\APONTAMENTOS"),
    Path(r"V:\REPORT DE PRODUCAO\1. INDICADORES PRODUTIVOS PECVAL\APONTAMENTOS"),
    Path(r"V:\Departamentos\REPORT DE PRODUCAO\1. INDICADORES PRODUTIVOS PECVAL\APONTAMENTOS"),
]

RAIZES_KPIS_OFICIAIS_REDE = [
    Path(r"\\10.72.110.56\Departamentos\Producao\Confidential\1. INDICADORES PRODUTIVOS PECVAL"),
    Path(r"V:\Producao\Confidential\1. INDICADORES PRODUTIVOS PECVAL"),
    Path(r"V:\Departamentos\Producao\Confidential\1. INDICADORES PRODUTIVOS PECVAL"),
]

# Subpastas padrão de cada um dos 4 setores (Injeção, Montagem, Pintura Manual, Pintura Automática)
ESTRUTURA_SETORES = {
    "INJECAO": {
        "nome_exibicao": "01. Injeção Plástica (F1 / Lado A / Lado B)",
        "subpastas_apontamento": [
            r"INJEÇÃO\F1",
            r"INJEÇÃO\F2-( LADO A )",
            r"INJEÇÃO\F2-( LADO B )",
            r"INJEÇÃO",
        ],
        "subpasta_kpi": "INJETORA",
        "aba_kpi": "APONTAR AQUI",
        "linha_cabecalho_kpi": 2,
        "total_colunas_kpi": 139,
        "abas_apontamento": ["ABA 1", "ABA 2", "ABA 3", "ABA 4", "ABA 5"],
        "nome_master": "BASE_MASTER_INJECAO.xlsx",
    },
    "MONTAGEM": {
        "nome_exibicao": "02. Montagem Industrial (Lado A-B)",
        "subpastas_apontamento": [
            r"MONTAGEM\Montagem LADO A LADO B",
            r"MONTAGEM",
        ],
        "subpasta_kpi": "MONTAGEM",
        "aba_kpi": "APONTAR AQUI",
        "linha_cabecalho_kpi": 3,
        "total_colunas_kpi": 116,
        "abas_apontamento": ["ABA 1", "ABA 2", "ABA 3"],
        "nome_master": "BASE_MASTER_MONTAGEM.xlsx",
    },
    "PINTURA_MANUAL": {
        "nome_exibicao": "03. Pintura Manual (Leandro e Marcos Manual)",
        "subpastas_apontamento": [
            r"PINTURA\LEANDRO  E Marcos Manual",
            r"PINTURA\LEANDRO E Marcos Manual",
            r"PINTURA MANUAL",
            r"PINTURA",
        ],
        "subpasta_kpi": "PINTURA",
        "subpastas_kpi_alternativas": [
            r"PINTURA\PINTURA MANUAL",
            r"PINTURA\Manual",
            r"PINTURA\LEANDRO  E Marcos Manual",
            r"PINTURA\LEANDRO E Marcos Manual",
            "PINTURA MANUAL",
            "PINTURA",
        ],
        "aba_kpi": "APONTAR AQUI ",  # Possui espaço no final no arquivo original
        "linha_cabecalho_kpi": 3,
        "total_colunas_kpi": 147,
        "colunas_formula_protegidas": {
            5, 6, 9, 32, 43, 55, 67, 68, 69, 70, 73,
            95, 99, 103, 107, 111, 115, 119, 123, 127, 131,
            133, 134, 135, 136, 137, 138, 139, 144
        },
        "abas_apontamento": ["ABA 1", "ABA 2", "ABA 3", "ABA RETRABALHO"],
        "nome_master": "BASE_MASTER_PINTURA_MANUAL.xlsx",
    },
    "PINTURA_AUTOMATICA": {
        "nome_exibicao": "04. Pintura Automática (Rodrigo e Emerson Automático)",
        "subpastas_apontamento": [
            r"PINTURA\Rodrigo e Emerson Automático",
            r"PINTURA\Rodrigo e Emerson Automatico",
            r"PINTURA AUTOMATICA",
            r"PINTURA AUTOMÁTICA",
            r"PINTURA",
        ],
        "subpasta_kpi": "PINTURA",
        "subpastas_kpi_alternativas": [
            r"PINTURA\PINTURA AUTOMATICA",
            r"PINTURA\PINTURA AUTOMÁTICA",
            r"PINTURA\Automático",
            r"PINTURA\Rodrigo e Emerson Automático",
            "PINTURA AUTOMATICA",
            "PINTURA AUTOMÁTICA",
            "PINTURA",
        ],
        "aba_kpi": "APONTAR AQUI ",  # Possui espaço no final no arquivo original
        "linha_cabecalho_kpi": 3,
        "total_colunas_kpi": 146,
        "abas_apontamento": ["ABA 1", "ABA 2", "ABA 3", "ABA RETRABALHO"],
        "nome_master": "BASE_MASTER_PINTURA_AUTOMATICA.xlsx",
    },
}

MESES_PT = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
}

ANO_OPERACAO_PADRAO = 2026
ANOS_LEGADOS_IGNORAR = {2020, 2021, 2022, 2023, 2024, 2025}


# ==============================================================================
# 3. FUNÇÕES UTILITÁRIAS DE NORMALIZAÇÃO E RESOLUÇÃO DE DATAS/ARQUIVOS
# ==============================================================================
def normalizar_texto(texto: str) -> str:
    """Remove acentos, espaços e caracteres especiais para comparação à prova de falhas."""
    if not texto:
        return ""
    sem_acento = unicodedata.normalize("NFKD", str(texto)).encode("ASCII", "ignore").decode("ASCII")
    return re.sub(r"[^a-z0-9]", "", sem_acento.lower())


def _extrair_chave_ano_mes_pasta(
    nome_pasta: str,
    ano_fallback: Optional[int] = None,
) -> Tuple[int, int, str]:
    """
    Extrai (ano, mes, nome) do nome da pasta ou arquivo.
    Suporta:
      - Anos de 4 dígitos (ex: '2026', '2025', '2024')
      - Anos de 2 dígitos (ex: 'DEZEMBRO 25' -> 2025, 'OUTUBRO - 26' -> 2026, 'kpi ... 25.xlsx' -> 2025)
      - Meses por extenso (JANEIRO..DEZEMBRO) e abreviações (JAN..DEZ)
      - Prefixos numéricos de mês (ex: '10- OUTUBRO', '12 . kpi ...')
    Não assume ano corrente cegamente caso não encontre correspondência no nome.
    """
    if not nome_pasta:
        return (ano_fallback or 0, 0, "")

    n_norm = normalizar_texto(nome_pasta)

    # 1. Busca ano de 4 dígitos (2010 a 2039)
    ano: Optional[int] = None
    m_ano_4 = re.search(r"(?<!\d)(20[1-3]\d)(?!\d)", nome_pasta)
    if m_ano_4:
        ano = int(m_ano_4.group(1))

    # 2. Busca mês por extenso (MESES_PT)
    mes = 0
    for m_num, m_nome in MESES_PT.items():
        if normalizar_texto(m_nome) in n_norm:
            mes = m_num
            break

    # 2b. Se não achou mês por extenso, busca abreviações de 3 letras (jan..dez)
    if mes == 0:
        meses_abrev_map = {
            "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
            "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
        }
        m_abrev = re.search(
            r"(?<![a-zA-Z])(jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)(?![a-zA-Z])",
            nome_pasta,
            re.IGNORECASE,
        )
        if m_abrev:
            mes = meses_abrev_map.get(m_abrev.group(1).lower(), 0)

    # 3. Busca ano de 2 dígitos se ano de 4 dígitos não foi encontrado
    if ano is None:
        nomes_meses_lista = list(MESES_PT.values()) + [
            "JAN", "FEV", "MAR", "MARCO", "ABR", "MAI", "JUN",
            "JUL", "AGO", "SET", "OUT", "NOV", "DEZ",
        ]
        nomes_meses_regex = "|".join(nomes_meses_lista)

        # Padrão A1: 2 dígitos seguidos de delimitador e mês (ex: '25 DEZ', '25-DEZ', '25_DEZ', '25 DEZEMBRO', '26 OUT', '25 OUTUBRO')
        m_ano_pre_mes = re.search(
            rf"(?<!\d)(2[0-9])[\s\-_/.]+(?:{nomes_meses_regex})(?![a-zA-Z])",
            nome_pasta,
            re.IGNORECASE,
        )
        if m_ano_pre_mes:
            ano = 2000 + int(m_ano_pre_mes.group(1))
        else:
            # Padrão A2: Mês (extenso ou abrev) seguido de delimitador e 2 dígitos (ex: 'DEZEMBRO 25', 'DEZEMBRO - 25', 'DEZ 25', 'DEZ_25', 'OUTUBRO-26')
            m_ano_pos_mes = re.search(
                rf"(?:{nomes_meses_regex})[\s\-_/.]*(?<!\d)(2[0-9])(?!\d)",
                nome_pasta,
                re.IGNORECASE,
            )
            if m_ano_pos_mes:
                ano = 2000 + int(m_ano_pos_mes.group(1))
            else:
                # Padrão B: Sufixo de 2 dígitos no final do nome ou do stem (ex: '... - 25.xlsx', '..._25', '... 25', 'kpi_25', 'kpi - 24')
                m_ano_fim = re.search(r"[\s\-_/](2[0-9])(?:\.[a-zA-Z0-9]+)?$", nome_pasta)
                if m_ano_fim:
                    ano = 2000 + int(m_ano_fim.group(1))
                else:
                    # Padrão C: Delimitador seguido de 2 dígitos quando o mês já foi detectado (ex: '10 - 25')
                    m_ano_delim = re.search(r"[\s\-_/]+(2[0-9])(?!\d)", nome_pasta)
                    if m_ano_delim and mes != 0:
                        cand_ano = 2000 + int(m_ano_delim.group(1))
                        if cand_ano != (2000 + mes):
                            ano = cand_ano

    # 4. Se o mês ainda for 0, tenta prefixo numérico no início (ex: '10- OUTUBRO', '12 . kpi ...', '01_Janeiro')
    if mes == 0:
        m_num_re = re.match(r"^\s*0?(\d{1,2})\s*[\.\-_]", nome_pasta)
        if m_num_re and 1 <= int(m_num_re.group(1)) <= 12:
            mes = int(m_num_re.group(1))

    ano_final = ano if ano is not None else (ano_fallback if ano_fallback is not None else 0)
    return (ano_final, mes, nome_pasta)


def eh_diretorio_legado_ou_ignorado(caminho: Union[Path, str], ano_vigente: int = 2026) -> bool:
    """
    Retorna True se o diretório (ou qualquer pasta pai) for de um ano anterior
    (ex: 2025, 2024, DEZEMBRO 25, 25 DEZ, 25_DEZ, INJETORA_2025) ou pasta arquivada/morta/backup/legada.
    """
    if caminho is None:
        return False
    p = Path(caminho)
    ignorar_tokens = {
        "old", "antigo", "antigos", "historico", "historicos",
        "legado", "legados", "lixeira", "desativado", "desativados",
    }
    for part in p.parts:
        part_norm = normalizar_texto(part)
        tokens = set(re.findall(r"[a-z0-9]+", unicodedata.normalize("NFKD", str(part)).encode("ASCII", "ignore").decode("ASCII").lower()))
        if any(tok in ignorar_tokens for tok in tokens):
            return True
        if "arquivomorto" in part_norm or ("arquivo" in tokens and "morto" in tokens):
            return True
        if "backupkpi" in part_norm or ("backup" in tokens and "kpi" in tokens):
            return True
        # 1. Ano de 4 dígitos isolado ou com delimitador (2010 a ano_vigente - 1)
        m_ano = re.search(r"(?<!\d)(20[1-2]\d)(?!\d)", part)
        if m_ano:
            ano_val = int(m_ano.group(1))
            if 2010 <= ano_val < ano_vigente:
                return True
        # 2. Extração de ano via parser (incluindo 2 dígitos como 'DEZEMBRO 25', '25 DEZ', '25_DEZ')
        ano_p, _, _ = _extrair_chave_ano_mes_pasta(part, ano_fallback=None)
        if ano_p != 0 and 2010 <= ano_p < ano_vigente:
            return True

    return False


def extrair_data_do_nome(nome_arquivo: str, ano_padrao: Optional[int] = None, mes_esperado: Optional[int] = None) -> Optional[dt.date]:
    """
    Extrai a data real (dt.date) a partir do nome do arquivo de apontamento.
    Suporta padrões como:
      - '09-10-APONTAMENTO INJEÇÃO F1.xlsm' -> 2026-10-09
      - '13.10.2026APONTAMENTO MONTAGEM LADO A-B.xlsm' -> 2026-10-13
      - 'APONTAMENTO  PINTURA  13.10.2026.xlsm' -> 2026-10-13
      - '13.10- APONTAMENTO PINTURA AUTOMÁTICA.xlsm' -> 2026-10-13
    Protegido contra erros de digitação de ano no nome do arquivo (ex: '2028' em vez de '2026').
    """
    ano_atual = dt.date.today().year
    ano_ref = ano_padrao or ano_atual
    m_ano = re.search(r"(?<!\d)(202[4-9]|203\d)(?!\d)", nome_arquivo)
    if m_ano:
        ano_lido = int(m_ano.group(1))
        if 2024 <= ano_lido <= ano_atual + 1:
            ano_ref = ano_lido

    m_data = re.search(r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})(?:[.\-_](\d{2,4}))?(?!\d)", nome_arquivo)
    if not m_data:
        return None

    dia = int(m_data.group(1))
    mes = int(m_data.group(2))
    ano = ano_ref
    if m_data.group(3):
        ano_grp = int(m_data.group(3))
        if ano_grp < 100:
            ano_grp += 2000
        if 2024 <= ano_grp <= ano_atual + 1:
            ano = ano_grp

    if mes_esperado and mes != mes_esperado:
        return None

    try:
        return dt.date(ano, mes, dia)
    except ValueError:
        return None


def obter_pastas_busca_apontamentos(
    setor_id: str,
    mes_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
) -> List[Path]:
    """
    Retorna a lista ordenada de pastas onde procurar os apontamentos (.xlsm) do setor:
    1º Nas pastas oficiais da rede (filtrando pela pasta do mês se existir);
    2º Na pasta V:\\Joao ADM-Prod\\PECVAL;
    3º Na pasta local (fallback para testes fora da fábrica).
    Filtra estritamente diretórios legados de anos anteriores (ex: 2025, 2024).
    """
    if 1 <= ano_alvo <= 12 and mes_alvo is None:
        mes_alvo = ano_alvo
        ano_alvo = 2026

    cfg = ESTRUTURA_SETORES[setor_id]
    pastas: List[Path] = []

    for raiz in RAIZES_APONTAMENTOS_REDE:
        try:
            if str(raiz).startswith("\\\\") and not REDE_UNC_ATIVA:
                continue
            if str(raiz).upper().startswith("V:") and not DRIVE_V_ATIVO:
                continue
            if not raiz.exists():
                continue
            if eh_diretorio_legado_ou_ignorado(raiz, ano_vigente=ano_alvo):
                continue

            # Descobre pastas de mês dentro de APONTAMENTOS (ex: '10- OUTUBRO - 2026')
            subpastas_mes = [
                p for p in raiz.iterdir()
                if p.is_dir() and not eh_diretorio_legado_ou_ignorado(p, ano_vigente=ano_alvo)
            ]
            if mes_alvo:
                nome_mes_norm = normalizar_texto(MESES_PT.get(mes_alvo, ""))
                subpastas_mes_filtradas = [
                    p for p in subpastas_mes
                    if f"{mes_alvo:02d}" in p.name or nome_mes_norm in normalizar_texto(p.name)
                ]
                if subpastas_mes_filtradas:
                    subpastas_mes = subpastas_mes_filtradas

            for p_mes in sorted(subpastas_mes, key=lambda p: _extrair_chave_ano_mes_pasta(p.name), reverse=True):
                for sub_setor in cfg["subpastas_apontamento"]:
                    cand = p_mes / sub_setor
                    if cand.exists() and not eh_diretorio_legado_ou_ignorado(cand, ano_vigente=ano_alvo):
                        pastas.append(cand)
                        # Também inclui subpastas internas (ex: dentro de PINTURA ou MONTAGEM)
                        try:
                            for sub_interna in cand.iterdir():
                                if sub_interna.is_dir() and not eh_diretorio_legado_ou_ignorado(sub_interna, ano_vigente=ano_alvo):
                                    if sub_interna not in pastas:
                                        pastas.append(sub_interna)
                        except OSError:
                            pass
                if p_mes not in pastas and not eh_diretorio_legado_ou_ignorado(p_mes, ano_vigente=ano_alvo):
                    pastas.append(p_mes)
        except OSError:
            continue

    # Pastas de fallback (V:\Joao ADM-Prod\PECVAL e diretório local)
    fallbacks = ([PASTA_REDE_JOAO_PECVAL] if DRIVE_V_ATIVO else []) + [DIRETORIO_PECVAL_LOCAL, DIRETORIO_AUTOMACAO_KPI]
    for fb in fallbacks:
        try:
            if not fb.exists() or eh_diretorio_legado_ou_ignorado(fb, ano_vigente=ano_alvo):
                continue
            for sub_setor in cfg["subpastas_apontamento"]:
                cand = fb / sub_setor
                if cand.exists() and cand not in pastas and not eh_diretorio_legado_ou_ignorado(cand, ano_vigente=ano_alvo):
                    pastas.append(cand)
            if fb not in pastas:
                pastas.append(fb)
        except OSError:
            continue

    return pastas


def obter_pastas_busca_kpi(setor_id: str, ano_alvo: int = 2026) -> List[Path]:
    """
    Retorna a lista ordenada de pastas onde procurar o KPI Oficial (.xlsx) do setor:
    1º Na pasta Confidential do setor na rede (e subpastas dentro de Confidential);
    2º Na pasta V:\\Joao ADM-Prod\\PECVAL;
    3º Na pasta local (fallback para testes).
    Filtra estritamente subpastas arquivadas de anos legados (ex: 2025, 2024).
    """
    cfg = ESTRUTURA_SETORES[setor_id]
    pastas: List[Path] = []
    opcoes_sub = [cfg["subpasta_kpi"]] + cfg.get("subpastas_kpi_alternativas", [])

    for raiz in RAIZES_KPIS_OFICIAIS_REDE:
        try:
            if str(raiz).startswith("\\\\") and not REDE_UNC_ATIVA:
                continue
            if str(raiz).upper().startswith("V:") and not DRIVE_V_ATIVO:
                continue
            if not raiz.exists() or eh_diretorio_legado_ou_ignorado(raiz, ano_vigente=ano_alvo):
                continue

            # 1. Tenta subpastas configuradas (ex: PINTURA, PINTURA MANUAL, INJETORA, MONTAGEM)
            for sub_k in opcoes_sub:
                cand_setor = raiz / sub_k
                if cand_setor.exists() and not eh_diretorio_legado_ou_ignorado(cand_setor, ano_vigente=ano_alvo):
                    if cand_setor not in pastas:
                        pastas.append(cand_setor)
                    try:
                        for sub_interna in cand_setor.iterdir():
                            if sub_interna.is_dir() and not eh_diretorio_legado_ou_ignorado(sub_interna, ano_vigente=ano_alvo):
                                if sub_interna not in pastas:
                                    pastas.append(sub_interna)
                    except OSError:
                        pass

            # 2. Varre todas as pastas de 1º e 2º nível dentro de Confidential (caso o nome da pasta tenha mudado)
            try:
                for item_dir in raiz.iterdir():
                    if item_dir.is_dir() and not eh_diretorio_legado_ou_ignorado(item_dir, ano_vigente=ano_alvo):
                        if item_dir not in pastas:
                            pastas.append(item_dir)
                        try:
                            for sub_dir in item_dir.iterdir():
                                if sub_dir.is_dir() and not eh_diretorio_legado_ou_ignorado(sub_dir, ano_vigente=ano_alvo):
                                    if sub_dir not in pastas:
                                        pastas.append(sub_dir)
                        except OSError:
                            pass
            except OSError:
                pass

            if raiz not in pastas and not eh_diretorio_legado_ou_ignorado(raiz, ano_vigente=ano_alvo):
                pastas.append(raiz)
        except OSError:
            continue

    fallbacks = ([PASTA_REDE_JOAO_PECVAL] if DRIVE_V_ATIVO else []) + [DIRETORIO_PECVAL_LOCAL, DIRETORIO_AUTOMACAO_KPI]
    for fb in fallbacks:
        try:
            if not fb.exists() or eh_diretorio_legado_ou_ignorado(fb, ano_vigente=ano_alvo):
                continue
            for sub_k in opcoes_sub:
                cand = fb / sub_k
                if cand.exists() and cand not in pastas and not eh_diretorio_legado_ou_ignorado(cand, ano_vigente=ano_alvo):
                    pastas.append(cand)
            if fb not in pastas:
                pastas.append(fb)
        except OSError:
            continue

    return pastas


def _eh_apontamento_do_setor(nome_norm: str, setor_id: str, sufixo: str, pasta_norm: str = "") -> bool:
    """
    Verifica se um arquivo (.xlsm/.xlsx) pertence aos apontamentos do setor,
    considerando tanto o nome do arquivo quanto a pasta onde ele está salvo
    (ex: 'LEANDRO E Marcos Manual', 'Rodrigo e Emerson Automático', 'Montagem LADO A LADO B').
    """
    if sufixo not in (".xlsm", ".xlsx"):
        return False
    if any(bloq in nome_norm for bloq in ["kpi", "master", "backup", "teste", "config"]):
        return False

    if setor_id == "INJECAO":
        if any(p in pasta_norm for p in ["f1", "f2ladoa", "f2ladob", "injec"]):
            return "apontamento" in nome_norm or "injec" in nome_norm
        return "injec" in nome_norm and any(p in nome_norm for p in ["f1", "ladoa", "ladob", "f2"])

    if setor_id == "MONTAGEM":
        if "montag" in pasta_norm:
            return True
        return "montag" in nome_norm

    if setor_id == "PINTURA_MANUAL":
        # Se estiver na pasta do Rodrigo/Emerson Automático, não é Pintura Manual
        if any(a in pasta_norm for a in ["rodrigo", "emerson", "automatic"]):
            return False
        if any(m in pasta_norm for m in ["leandro", "marcos", "manual"]):
            return True
        return "pintura" in nome_norm and "auto" not in nome_norm

    if setor_id == "PINTURA_AUTOMATICA":
        # Se estiver na pasta do Leandro/Marcos Manual, não é Pintura Automática
        if any(m in pasta_norm for m in ["leandro", "marcos", "manual"]):
            return False
        if any(a in pasta_norm for a in ["rodrigo", "emerson", "automatic"]):
            return True
        return "pintura" in nome_norm and "auto" in nome_norm

    return False


def _eh_kpi_do_setor(nome_norm: str, setor_id: str, sufixo: str) -> bool:
    """Verifica se o nome normalizado de um arquivo é o KPI Oficial do setor."""
    if sufixo != ".xlsx" or "kpi" not in nome_norm:
        return False
    if any(bloq in nome_norm for bloq in ["teste", "master", "backup"]):
        return False

    if setor_id == "INJECAO":
        return "injec" in nome_norm
    if setor_id == "MONTAGEM":
        return "montag" in nome_norm
    if setor_id == "PINTURA_MANUAL":
        return "pintura" in nome_norm and "manual" in nome_norm
    if setor_id == "PINTURA_AUTOMATICA":
        return "pintura" in nome_norm and ("auto" in nome_norm or "robo" in nome_norm)
    return False


def _chave_ordenacao_kpi(p: Path, ano_alvo: int = 2026) -> Tuple[int, int, int, float]:
    """
    Ordena KPIs com prioridade estrita:
    1º Coincidência com ano_alvo (arquivos de 2026 têm prioridade absoluta sobre legados como 2025);
    2º Ano mais recente (determinado pelo nome do arquivo ou herdado das pastas pai);
    3º Mês mais recente (Novembro > Outubro);
    4º Timestamp de modificação mais recente.
    """
    # 1. Tenta extrair ano e mês do nome do arquivo
    ano_k, mes_k, _ = _extrair_chave_ano_mes_pasta(p.name, ano_fallback=None)

    # 2. Se o nome do arquivo não trouxer o ano, herda das pastas pai (da mais próxima à raiz)
    if ano_k == 0:
        for parent_part in reversed(p.parent.parts):
            ano_parent, mes_parent, _ = _extrair_chave_ano_mes_pasta(parent_part, ano_fallback=None)
            if ano_parent != 0:
                ano_k = ano_parent
                if mes_k == 0 and mes_parent != 0:
                    mes_k = mes_parent
                break

    if ano_k == 0:
        if eh_diretorio_legado_ou_ignorado(p.parent, ano_vigente=ano_alvo):
            ano_k = 2020
        else:
            ano_k = ano_alvo

    if mes_k == 0:
        for parent_part in reversed(p.parent.parts):
            _, mes_parent, _ = _extrair_chave_ano_mes_pasta(parent_part, ano_fallback=None)
            if 1 <= mes_parent <= 12:
                mes_k = mes_parent
                break

    match_ano = 1 if (ano_k == ano_alvo) else 0

    try:
        mtime = p.stat().st_mtime
    except OSError:
        mtime = 0.0

    return (match_ano, ano_k, mes_k, mtime)


_CACHE_LOCALIZAR_KPI: Dict[Tuple[str, int, Optional[int]], Tuple[float, Optional[Path]]] = {}
_CACHE_APONTAMENTOS: Dict[Tuple[str, Optional[int], int], Tuple[float, Dict[dt.date, List[Path]]]] = {}
TTL_CACHE_REDE_SEGUNDOS = 120.0


def limpar_caches_caminho() -> None:
    """Limpa os caches em memória de caminhos e arquivos de rede."""
    _CACHE_LOCALIZAR_KPI.clear()
    _CACHE_APONTAMENTOS.clear()


def localizar_kpi_oficial(
    setor_id: str,
    ano_alvo: int = 2026,
    mes_alvo: Optional[int] = None,
) -> Optional[Path]:
    """
    Localiza dinamicamente o arquivo de KPI Oficial (.xlsx) do setor para o ano_alvo (padrão 2026).
    Rejeita estritamente subpastas arquivadas (como 2025\\, 2024\\) e arquivos legados.
    Prioriza o arquivo do mês alvo (se informado) ou o mais recente (Ano/Mês/Data).
    Possui cache de 120s para acelerar em 100x acessos via VPN.
    """
    # Suporta chamada posicional legada onde mes_alvo era passado como 2º argumento
    if 1 <= ano_alvo <= 12 and mes_alvo is None:
        mes_alvo = ano_alvo
        ano_alvo = 2026

    chave_cache = (setor_id, ano_alvo, mes_alvo, tuple(str(r) for r in RAIZES_KPIS_OFICIAIS_REDE))
    agora = time.time()
    if chave_cache in _CACHE_LOCALIZAR_KPI:
        t_grav, res_cache = _CACHE_LOCALIZAR_KPI[chave_cache]
        if (agora - t_grav) < TTL_CACHE_REDE_SEGUNDOS:
            return res_cache

    candidatos: List[Path] = []
    pastas = obter_pastas_busca_kpi(setor_id, ano_alvo=ano_alvo)

    for pasta in pastas:
        try:
            if eh_diretorio_legado_ou_ignorado(pasta, ano_vigente=ano_alvo):
                continue
            for item in pasta.iterdir():
                if item.name.startswith("~$") or not item.is_file():
                    continue
                if eh_diretorio_legado_ou_ignorado(item, ano_vigente=ano_alvo):
                    continue

                n_norm = normalizar_texto(item.name)
                if not _eh_kpi_do_setor(n_norm, setor_id, item.suffix.lower()):
                    continue

                # Determina o ano do arquivo inspecionando nome e pastas pai
                ano_item, mes_item, _ = _extrair_chave_ano_mes_pasta(item.name, ano_fallback=None)
                if ano_item == 0:
                    for parent_part in reversed(item.parent.parts):
                        ano_p, mes_p, _ = _extrair_chave_ano_mes_pasta(parent_part, ano_fallback=None)
                        if ano_p != 0:
                            ano_item = ano_p
                            if mes_item == 0:
                                mes_item = mes_p
                            break

                # Se o arquivo declarar explicitamente outro ano (ex: 2025, 2024), rejeita
                if ano_item != 0 and ano_item != ano_alvo:
                    continue

                candidatos.append(item)
        except OSError:
            continue

    if not candidatos:
        _CACHE_LOCALIZAR_KPI[chave_cache] = (agora, None)
        return None

    if mes_alvo:
        nome_mes_norm = normalizar_texto(MESES_PT.get(mes_alvo, ""))
        cands_mes = []
        for c in candidatos:
            _, mes_c, _ = _extrair_chave_ano_mes_pasta(c.name, ano_fallback=None)
            if mes_c == mes_alvo or nome_mes_norm in normalizar_texto(c.name) or re.match(rf"^\s*0?{mes_alvo}\s*[\.\-_]", c.name):
                cands_mes.append(c)
        if cands_mes:
            candidatos = cands_mes
        else:
            _CACHE_LOCALIZAR_KPI[chave_cache] = (agora, None)
            return None

    candidatos.sort(key=lambda p: _chave_ordenacao_kpi(p, ano_alvo=ano_alvo), reverse=True)
    res_kpi = candidatos[0]
    _CACHE_LOCALIZAR_KPI[chave_cache] = (agora, res_kpi)
    return res_kpi


def _extrair_contexto_ano_mes_do_caminho(
    caminho: Path,
    ano_padrao: int = 2026,
) -> Tuple[Optional[int], Optional[int]]:
    """
    Descobre em qual pasta de mês/ano (ex: '10- OUTUBRO - 2026') o arquivo está salvo.
    Percorre as partes do caminho mantendo memória do ano definido em pastas superiores (ex: '2025')
    e vinculando ao mês definido nas subpastas (ex: '10- OUTUBRO'), evitando atribuição incorreta
    do ano corrente a pastas históricas.
    """
    p = Path(caminho)
    ano_acumulado: Optional[int] = None
    mes_acumulado: Optional[int] = None

    for parte in p.parts:
        ano_p, mes_p, _ = _extrair_chave_ano_mes_pasta(parte, ano_fallback=None)
        if ano_p != 0:
            ano_acumulado = ano_p
        if 1 <= mes_p <= 12:
            mes_acumulado = mes_p

    ano_final = ano_acumulado if ano_acumulado is not None else ano_padrao
    return (ano_final, mes_acumulado)


def listar_apontamentos_disponiveis(
    setor_id: str,
    mes_alvo: Optional[int] = None,
    ano_alvo: int = 2026,
) -> Dict[dt.date, List[Path]]:
    """
    Varre as pastas de apontamento do setor e agrupa todos os arquivos (.xlsm)
    encontrados por data (dt.date), permitindo saber exatamente quais dias existem na pasta.
    Se mes_alvo não for informado, prioriza automaticamente o mês do KPI Oficial do setor
    (ou o mês vigente), filtrando erros de digitação de datas nas pastas da rede.
    Garante que arquivos em pastas legadas ou de outros anos sejam ignorados.
    Possui cache de 120s para acelerar em 100x acessos via VPN.
    """
    if 1 <= ano_alvo <= 12 and mes_alvo is None:
        mes_alvo = ano_alvo
        ano_alvo = 2026

    chave_cache = (setor_id, mes_alvo, ano_alvo, tuple(str(r) for r in RAIZES_APONTAMENTOS_REDE))
    agora = time.time()
    if chave_cache in _CACHE_APONTAMENTOS:
        t_grav, res_cache = _CACHE_APONTAMENTOS[chave_cache]
        if (agora - t_grav) < TTL_CACHE_REDE_SEGUNDOS:
            return res_cache

    if mes_alvo is None:
        kpi_ref = localizar_kpi_oficial(setor_id, ano_alvo=ano_alvo)
        if kpi_ref:
            _, mes_kpi, _ = _extrair_chave_ano_mes_pasta(kpi_ref.name, ano_fallback=ano_alvo)
            if 1 <= mes_kpi <= 12:
                mes_alvo = mes_kpi

    por_data: Dict[dt.date, List[Path]] = {}
    nomes_ja_vistos = set()

    for pasta in obter_pastas_busca_apontamentos(setor_id, mes_alvo=mes_alvo, ano_alvo=ano_alvo):
        try:
            if eh_diretorio_legado_ou_ignorado(pasta, ano_vigente=ano_alvo):
                continue
            pasta_norm = normalizar_texto(pasta.name)
            ano_ctx, mes_ctx = _extrair_contexto_ano_mes_do_caminho(pasta, ano_padrao=ano_alvo)
            if ano_ctx and ano_ctx != ano_alvo:
                continue
            mes_filtro = mes_alvo or mes_ctx

            for item in pasta.iterdir():
                if item.name.startswith("~$") or not item.is_file():
                    continue
                if eh_diretorio_legado_ou_ignorado(item, ano_vigente=ano_alvo):
                    continue
                chave_nome = item.name.lower()
                if chave_nome in nomes_ja_vistos:
                    continue

                n_norm = normalizar_texto(item.name)
                if not _eh_apontamento_do_setor(n_norm, setor_id, item.suffix.lower(), pasta_norm=pasta_norm):
                    continue

                data_arq = extrair_data_do_nome(item.name, ano_padrao=ano_ctx, mes_esperado=mes_filtro)
                if not data_arq:
                    continue
                if data_arq.year != ano_alvo:
                    continue
                if mes_alvo and data_arq.month != mes_alvo:
                    continue

                nomes_ja_vistos.add(chave_nome)
                por_data.setdefault(data_arq, []).append(item)
        except OSError:
            continue

    # Ordena os arquivos dentro de cada data (ex: F1, Lado A, Lado B)
    for dt_k in por_data:
        por_data[dt_k].sort(key=lambda p: p.name)

    res_final = dict(sorted(por_data.items()))
    _CACHE_APONTAMENTOS[chave_cache] = (agora, res_final)
    return res_final


def obter_caminho_master(setor_id: str, data_ref: Optional[dt.date] = None) -> Path:
    """Retorna o caminho onde a planilha Master consolidada do setor será salva."""
    PASTA_MASTERS.mkdir(parents=True, exist_ok=True)
    mes_num = data_ref.month if data_ref else dt.date.today().month
    ano_num = data_ref.year if data_ref else dt.date.today().year
    nome_mes = MESES_PT.get(mes_num, "MES")
    return PASTA_MASTERS / f"BASE_MASTER_{setor_id}_{nome_mes}_{ano_num}.xlsx"


def obter_caminho_kpi_teste(kpi_oficial: Path) -> Path:
    """Retorna o caminho para salvar uma cópia de homologação (- TESTE.xlsx) sem tocar no oficial."""
    PASTA_TESTES_KPI.mkdir(parents=True, exist_ok=True)
    meses_regex = "|".join(MESES_PT.values())
    stem_limpo = re.sub(rf"\s*-\s*({meses_regex}).*", "", kpi_oficial.stem, flags=re.IGNORECASE).strip()
    return PASTA_TESTES_KPI / f"{stem_limpo} - TESTE.xlsx"


def calcular_mes_ano_anterior(ano: int, mes: int) -> Tuple[int, int]:
    """
    Dado (ano, mes), calcula (ano_anterior, mes_anterior).
    Exemplos:
      (2026, 11) -> (2026, 10)
      (2026, 1)  -> (2025, 12)
    """
    if mes <= 1:
        return (ano - 1, 12)
    return (ano, mes - 1)


def obter_contexto_meses_operacao(
    setor_id: str,
    ano_alvo: int = 2026,
    mes_referencia: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Retorna o contexto operacional completo dos meses corrente/ativo e anterior:
      - 'setor_id': Identificador do setor
      - 'ano_atual': Ano alvo (padrão 2026)
      - 'mes_atual': Mês de referência ou mês mais recente com KPI/apontamento
      - 'nome_mes_atual': Nome do mês atual por extenso
      - 'kpi_atual': Path do KPI oficial do mês atual (se existir)
      - 'ano_anterior': Ano do mês anterior (calculado via calcular_mes_ano_anterior)
      - 'mes_anterior': Mês imediatamente anterior
      - 'nome_mes_anterior': Nome do mês anterior por extenso
      - 'kpi_anterior': Path do KPI oficial do mês anterior (se existir)
      - 'apontamentos_atual': Apontamentos agrupados por data para o mês atual
      - 'apontamentos_anterior': Apontamentos agrupados por data para o mês anterior
    """
    mes_atual = mes_referencia
    if mes_atual is None:
        kpi_ref = localizar_kpi_oficial(setor_id, ano_alvo=ano_alvo)
        if kpi_ref:
            _, mes_detectado, _ = _extrair_chave_ano_mes_pasta(kpi_ref.name, ano_fallback=ano_alvo)
            if 1 <= mes_detectado <= 12:
                mes_atual = mes_detectado
        if mes_atual is None:
            mes_atual = dt.date.today().month

    ano_atual = ano_alvo
    kpi_atual = localizar_kpi_oficial(setor_id, ano_alvo=ano_atual, mes_alvo=mes_atual)

    ano_ant, mes_ant = calcular_mes_ano_anterior(ano_atual, mes_atual)
    kpi_anterior = localizar_kpi_oficial(setor_id, ano_alvo=ano_ant, mes_alvo=mes_ant)

    aponts_atual = listar_apontamentos_disponiveis(setor_id, mes_alvo=mes_atual, ano_alvo=ano_atual)
    aponts_ant = listar_apontamentos_disponiveis(setor_id, mes_alvo=mes_ant, ano_alvo=ano_ant)

    return {
        "setor_id": setor_id,
        "ano_atual": ano_atual,
        "mes_atual": mes_atual,
        "nome_mes_atual": MESES_PT.get(mes_atual, ""),
        "kpi_atual": kpi_atual,
        "ano_anterior": ano_ant,
        "mes_anterior": mes_ant,
        "nome_mes_anterior": MESES_PT.get(mes_ant, ""),
        "kpi_anterior": kpi_anterior,
        "apontamentos_atual": aponts_atual,
        "apontamentos_anterior": aponts_ant,
    }


def localizar_kpis_vigentes(setor_id: str, ano_alvo: int = 2026) -> Dict[str, Optional[Path]]:
    """
    Localiza os KPIs vigentes de 2026 para o setor:
      - 'kpi_atual': KPI do mês mais recente no ano alvo
      - 'kpi_anterior': KPI do mês imediatamente anterior no ano alvo
    """
    ctx = obter_contexto_meses_operacao(setor_id, ano_alvo=ano_alvo)
    return {
        "kpi_atual": ctx["kpi_atual"],
        "kpi_anterior": ctx["kpi_anterior"],
    }


def descobrir_estrutura_rede_automatica(mes_alvo: Optional[int] = None) -> None:
    r"""
    Rastreador Automático de Rede (Scanner):
    Varre a pasta mãe '\\10.72.110.56\...\APONTAMENTOS' e a pasta de KPIs,
    lista automaticamente todas as pastas de meses existentes (Outubro, Novembro, Dezembro, etc.)
    e mostra exatamente onde estão os apontamentos e o KPI de cada setor!
    """
    print("=" * 92)
    print(" 🔎 RASTREADOR AUTOMÁTICO DE MESES, PASTAS E KPIs — PECVAL (ZERO CAMINHO FIXO DE MÊS)")
    print("=" * 92)
    print(f" 🌐 Rede UNC (\\\\10.72.110.56): {'ONLINE ✅' if REDE_UNC_ATIVA else 'OFFLINE (Usando Fallback Local) 🏠'}")
    print(f" 💾 Drive V:\\ Mapeado        : {'ONLINE ✅' if DRIVE_V_ATIVO else 'OFFLINE (Usando Fallback Local) 🏠'}")

    # Lista pastas de meses encontradas na raiz de APONTAMENTOS
    for raiz in RAIZES_APONTAMENTOS_REDE:
        try:
            if (str(raiz).startswith("\\\\") and REDE_UNC_ATIVA) or (str(raiz).upper().startswith("V:") and DRIVE_V_ATIVO):
                if raiz.exists():
                    pastas_mes = sorted(
                        [p for p in raiz.iterdir() if p.is_dir()],
                        key=lambda p: _extrair_chave_ano_mes_pasta(p.name),
                        reverse=True,
                    )
                    print(f"\n 📂 Pastas de Meses detectadas automaticamente em {raiz}:")
                    for pm in pastas_mes[:6]:
                        ano_p, mes_p, _ = _extrair_chave_ano_mes_pasta(pm.name)
                        print(f"    • [{mes_p:02d}/{ano_p}] -> {pm}")
                    break
        except OSError:
            continue

    # Lista as pastas de KPIs encontradas dentro de Confidential
    for raiz_kpi in RAIZES_KPIS_OFICIAIS_REDE:
        try:
            if (str(raiz_kpi).startswith("\\\\") and REDE_UNC_ATIVA) or (str(raiz_kpi).upper().startswith("V:") and DRIVE_V_ATIVO):
                if raiz_kpi.exists():
                    pastas_kpi = sorted([p for p in raiz_kpi.iterdir() if p.is_dir()], key=lambda p: p.name)
                    print(f"\n 🔒 Pastas de KPIs detectadas automaticamente em {raiz_kpi}:")
                    for pk in pastas_kpi:
                        arqs_xlsx = [f.name for f in pk.iterdir() if f.is_file() and f.suffix.lower() == ".xlsx" and not f.name.startswith("~$")]
                        print(f"    • {pk.name:<25} -> {len(arqs_xlsx)} arquivo(s) .xlsx | Ex: {arqs_xlsx[:2]}")
                    break
        except OSError:
            continue

    print("\n 📊 RESUMO DE LOCALIZAÇÃO AUTOMÁTICA POR SETOR (TRANSIÇÃO DE MÊS 2026):")
    for s_id, cfg in ESTRUTURA_SETORES.items():
        if mes_alvo:
            kpi = localizar_kpi_oficial(s_id, mes_alvo=mes_alvo)
            aponts = listar_apontamentos_disponiveis(s_id, mes_alvo=mes_alvo)
            datas = sorted(aponts.keys())
            ult_dia = datas[-1].strftime("%d/%m/%Y") if datas else "Nenhum"
            print(f"  ▸ {cfg['nome_exibicao']} [Mês filtrado: {mes_alvo:02d}/2026]")
            print(f"    ├── KPI Localizado     : {kpi if kpi else 'Não encontrado'}")
            print(f"    └── Apontamentos Lidos : {len(datas)} dia(s) encontrado(s) | Último arquivo: {ult_dia}")
        else:
            ctx = obter_contexto_meses_operacao(s_id, ano_alvo=ANO_OPERACAO_PADRAO)
            kpi_atu = ctx["kpi_atual"]
            aponts_atu = ctx["apontamentos_atual"]
            datas_atu = sorted(aponts_atu.keys())
            ult_atu = datas_atu[-1].strftime("%d/%m/%Y") if datas_atu else "Nenhum (Mês recém-iniciado)"

            kpi_ant = ctx["kpi_anterior"]
            aponts_ant = ctx["apontamentos_anterior"]
            datas_ant = sorted(aponts_ant.keys())
            ult_ant = datas_ant[-1].strftime("%d/%m/%Y") if datas_ant else "Nenhum"

            print(f"  ▸ {cfg['nome_exibicao']}")
            print(f"    ├── 🟢 MÊS ATUAL    [{ctx['mes_atual']:02d}/{ctx['ano_atual']} - {ctx['nome_mes_atual']}]:")
            print(f"    │      • KPI Oficial   : {kpi_atu.name if kpi_atu else 'Ainda não criado na rede'}")
            print(f"    │      • Caminho KPI   : {kpi_atu if kpi_atu else '---'}")
            print(f"    │      • Apontamentos  : {len(datas_atu)} dia(s) | Último: {ult_atu}")
            if kpi_ant or datas_ant:
                print(f"    └── 🔵 MÊS ANTERIOR [{ctx['mes_anterior']:02d}/{ctx['ano_anterior']} - {ctx['nome_mes_anterior']} (Consolidado)]:")
                print(f"           • KPI Anterior  : {kpi_ant.name if kpi_ant else 'Não encontrado'}")
                print(f"           • Apontamentos  : {len(datas_ant)} dia(s) | Último: {ult_ant}")
            else:
                print(f"    └── 🔵 MÊS ANTERIOR [{ctx['mes_anterior']:02d}/{ctx['ano_anterior']}]: Sem KPI/Apontamentos anteriores localizados")
    print("=" * 92)


if __name__ == "__main__":
    mes_cli = int(sys.argv[1]) if (len(sys.argv) > 1 and sys.argv[1].isdigit()) else None
    descobrir_estrutura_rede_automatica(mes_alvo=mes_cli)

