# -*- coding: utf-8 -*-
"""
=============================================================================================
🧪 AMBIENTE DE HOMOLOGAÇÃO & TESTES EM KPIS ZERADOS — PECVAL (ZERO RISCO À PRODUÇÃO)
=============================================================================================
Caminho: PECVAL/AutomacaoKPI/Executar/testar_kpi_zerado.py
Unidade: Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)

Objetivo:
  Permitir testar a automação completa (extração de apontamentos reais, cálculo de horas
  não programadas, filtros de colunas ocultas e gravação) em arquivos de KPI ZERADOS ou
  cópias de homologação, com 100% de isolamento:
  - NENHUM arquivo da pasta oficial de produção (Confidential) é lido ou modificado como KPI.
  - O resultado é gravado exclusivamente na sua pasta de testes com os KPIs zerados.
  - Você pode abrir a planilha no Excel logo após o teste para conferir abas, fórmulas e totais!
=============================================================================================
"""

import os
import sys
import subprocess
import datetime as dt
from pathlib import Path
from typing import Any, Dict, List, Optional

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
    ANO_OPERACAO_PADRAO,
    DRIVE_V_ATIVO,
    ESTRUTURA_SETORES,
    PASTA_BASE_OPERACAO,
    PASTA_KPIS_ZERADOS,
    ativar_ambiente_teste,
    desativar_ambiente_teste,
    listar_apontamentos_disponiveis,
    localizar_kpi_oficial,
    normalizar_texto,
    obter_caminho_master,
)
from ValidadorDia.validador_kpi import (
    diagnosticar_kpi_setor,
    exibir_relatorio_validador,
    salvar_planilha_master,
)
try:
    from Executar.rodar_todos import MAPA_PROCESSADORES
except ImportError:
    from rodar_todos import MAPA_PROCESSADORES


def abrir_arquivo_excel(caminho: Path) -> None:
    """Abre o arquivo no Excel nativo do Windows para inspeção visual."""
    try:
        if sys.platform == "win32":
            os.startfile(str(caminho))
        elif sys.platform == "darwin":
            subprocess.run(["open", str(caminho)], check=False)
        else:
            subprocess.run(["xdg-open", str(caminho)], check=False)
    except Exception as exc:
        print(f"  ⚠️ Não foi possível abrir o Excel automaticamente: {exc}")


def listar_kpis_na_pasta(pasta: Path) -> Dict[str, Path]:
    """Varre a pasta de teste e identifica qual KPI pertence a qual setor."""
    kpis_encontrados: Dict[str, Path] = {}
    if not pasta.exists():
        return kpis_encontrados

    for arq in sorted(pasta.glob("*.xlsx")):
        if arq.name.startswith("~$"):
            continue
        nome_norm = normalizar_texto(arq.name)
        if "master" in nome_norm or "backup" in nome_norm:
            continue

        if "injec" in nome_norm or "injet" in nome_norm:
            kpis_encontrados["INJECAO"] = arq
        elif "montag" in nome_norm:
            kpis_encontrados["MONTAGEM"] = arq
        elif "pintura" in nome_norm and "manual" in nome_norm:
            kpis_encontrados["PINTURA_MANUAL"] = arq
        elif "pintura" in nome_norm and ("auto" in nome_norm or "robo" in nome_norm):
            kpis_encontrados["PINTURA_AUTOMATICA"] = arq

    return kpis_encontrados


def executar_teste_kpi_zerado() -> None:
    print("\n" + "=" * 92)
    print(" 🧪 AMBIENTE DE HOMOLOGAÇÃO & TESTES EM KPIS ZERADOS — PECVAL")
    print("=" * 92)
    print(" 🔒 SEGURANÇA MÁXIMA:")
    print("    • Os arquivos oficiais da rede (V:\\Producao\\Confidential\\...) NÃO SERÃO TOCADOS.")
    print("    • A gravação ocorrerá APENAS na sua pasta de testes com os KPIs zerados.")
    print("=" * 92)

    # 1. Definição da pasta com os KPIs zerados
    pasta_padrao = PASTA_KPIS_ZERADOS
    pasta_padrao.mkdir(parents=True, exist_ok=True)

    print(f"\n📂 Pasta padrão para KPIs zerados:")
    print(f"   👉 {pasta_padrao}")
    print("\nOpções de pasta:")
    print("  [ENTER] Usar a pasta padrão acima")
    print("  [Caminho] Digite ou cole outro caminho onde estão seus KPIs zerados")
    resp_pasta = input("\nEscolha a pasta: ").strip().strip('"').strip("'")

    pasta_teste = Path(resp_pasta) if resp_pasta else pasta_padrao
    if not pasta_teste.exists():
        try:
            pasta_teste.mkdir(parents=True, exist_ok=True)
            print(f"  📁 Pasta criada: {pasta_teste}")
        except Exception as exc:
            print(f"  ❌ Erro ao acessar/criar pasta: {exc}")
            return

    # Ativa o isolamento de ambiente
    ativar_ambiente_teste(pasta_teste)

    try:
        # 2. Varredura dos arquivos zerados
        kpis_detectados = listar_kpis_na_pasta(pasta_teste)
        todos_arqs = [f for f in pasta_teste.glob("*.xlsx") if not f.name.startswith("~$")]

        print(f"\n🔍 Verificando arquivos na pasta de testes:")
        if not todos_arqs:
            print(f"  ⚠️ Nenhum arquivo .xlsx encontrado em: {pasta_teste}")
            print("\n  👉 COMO PROCEDER:")
            print(f"     1. Abra a pasta: {pasta_teste}")
            print("     2. Cole lá dentro seus arquivos de KPI zerados (modelos em branco).")
            print("     3. Exemplos de nomes reconhecidos automaticamente:")
            print("        • KPI INJEÇÃO 2026.xlsx")
            print("        • KPI MONTAGEM 2026.xlsx")
            print("        • KPI PINTURA MANUAL 2026.xlsx")
            print("        • KPI PINTURA AUTOMÁTICA 2026.xlsx")
            input("\nDepois de colar os arquivos na pasta, aperte ENTER para continuar...")
            kpis_detectados = listar_kpis_na_pasta(pasta_teste)
            todos_arqs = [f for f in pasta_teste.glob("*.xlsx") if not f.name.startswith("~$")]

        if not todos_arqs:
            print("  ❌ Nenhum arquivo .xlsx foi colocado na pasta. Teste cancelado.")
            return

        print(f"\n📋 Arquivos identificados na pasta de testes:")
        for s_id, cfg in ESTRUTURA_SETORES.items():
            arq_kpi = kpis_detectados.get(s_id)
            if arq_kpi:
                print(f"  ✅ [{s_id}] {cfg['nome_exibicao']}:")
                print(f"     Arquivo: {arq_kpi.name}")
            else:
                print(f"  ⏳ [{s_id}] {cfg['nome_exibicao']}: Nenhum KPI com nome reconhecido.")

        # 3. Escolha do Setor a Testar
        print("\n" + "-" * 92)
        print("Qual setor deseja testar agora?")
        print("  [1] INJEÇÃO")
        print("  [2] MONTAGEM")
        print("  [3] PINTURA MANUAL")
        print("  [4] PINTURA AUTOMÁTICA")
        print("  [5] TODOS OS SETORES QUE TÊM ARQUIVO ZERADO")
        print("  [0] Cancelar e Sair")
        opc_s = input("\nEscolha o setor [Padrão=3 PINTURA MANUAL]: ").strip()

        mapa_opc = {
            "1": ["INJECAO"],
            "2": ["MONTAGEM"],
            "3": ["PINTURA_MANUAL"],
            "4": ["PINTURA_AUTOMATICA"],
            "5": [s for s in ESTRUTURA_SETORES if s in kpis_detectados],
            "": ["PINTURA_MANUAL"],
        }
        setores_a_rodar = mapa_opc.get(opc_s, [])
        if not setores_a_rodar:
            print("⏹️ Operação cancelada.")
            return

        # 4. Escolha do Mês e Dia do Apontamento
        setor_ref = setores_a_rodar[0]
        m_in = input("\nDigite o número do Mês do apontamento real (ex: 10 para Outubro, 11 para Novembro) [ENTER=10]: ").strip()
        mes_val = int(m_in) if m_in.isdigit() else 10

        aponts = listar_apontamentos_disponiveis(setor_ref, mes_alvo=mes_val)
        dias_lista = [f"{d.day:02d}" for d in sorted(aponts.keys())]
        ult_dia_sugestao = dias_lista[-1] if dias_lista else ""

        if dias_lista:
            print(f"📅 Dias reais com apontamentos na rede ({setor_ref}): [{', '.join(dias_lista)}]")
            d_in = input(f"Digite o Dia a testar [ENTER = Dia {ult_dia_sugestao} (Automático)]: ").strip()
        else:
            d_in = input("Digite o Dia do apontamento [ENTER=Último disponível]: ").strip()

        dia_val = int(d_in) if d_in.isdigit() else (int(ult_dia_sugestao) if ult_dia_sugestao.isdigit() and not d_in else None)

        print("\n" + "=" * 92)
        print(" 🚀 INICIANDO TESTE CONTROLADO NO KPI ZERADO...")
        print("=" * 92)

        arquivos_para_abrir: List[Path] = []

        for s_id in setores_a_rodar:
            cfg = ESTRUTURA_SETORES[s_id]
            kpi_alvo = kpis_detectados.get(s_id)
            if not kpi_alvo:
                print(f"\n⚠️ [{cfg['nome_exibicao']}] Não foi encontrado arquivo zerado para este setor. Pulando...")
                continue

            print(f"\n--- SETOR: {cfg['nome_exibicao']} ---")
            print(f"  • KPI Zerado Alvo:  {kpi_alvo.name}")
            print(f"  • Caminho Completo: {kpi_alvo}")

            # Diagnóstico prévio
            diag = diagnosticar_kpi_setor(s_id, caminho_kpi_custom=kpi_alvo, mes_alvo=mes_val)
            print(f"  • Estado Inicial:   Linha Cabeçalho={diag['linha_cabecalho']} | Última Linha Preenchida={diag['ultima_linha_preenchida']} | Próxima Livre={diag['proxima_linha_livre']}")

            proc_fn = MAPA_PROCESSADORES.get(s_id)
            if not proc_fn:
                print(f"  ❌ Processador para {s_id} não encontrado.")
                continue

            # Executa com modo_oficial=True apontando para o arquivo de teste (isolado!)
            res = proc_fn(
                mes_alvo=mes_val,
                dia_alvo=dia_val,
                ano_alvo=ANO_OPERACAO_PADRAO,
                modo_oficial=True,
                permitir_atualizar_mesmo_dia=True,
                apenas_validar=False,
            )

            if not res.get("ok"):
                print(f"  ❌ FALHA NO TESTE: {res.get('motivo')}")
                continue

            rg = res.get("resultado_gravacao", {})
            print(f"  🎉 TESTE CONCLUÍDO COM SUCESSO!")
            print(f"     • Arquivo Gravado:      {rg.get('arquivo_gravado')}")
            print(f"     • Faixa Preenchida:     Linhas {rg.get('linha_inicio')} até {rg.get('linha_fim')} ({rg.get('linhas_inseridas')} linhas)")
            print(f"     • Motor Utilizado:      {rg.get('motor_gravacao')}")
            print(f"     • Fórmulas Preservadas: {rg.get('formulas_preservadas', 0)} célula(s)")
            print("     🔒 CONFIRMAÇÃO: A pasta de produção Confidential permaneceu 100% INTOCADA.")

            if kpi_alvo not in arquivos_para_abrir:
                arquivos_para_abrir.append(kpi_alvo)

        print("\n" + "=" * 92)
        print(" 🎯 TESTE FINALIZADO!")
        print("=" * 92)

        if arquivos_para_abrir:
            resp_abrir = input("\n👀 Deseja abrir a planilha de teste no Excel para conferir agora? [S/N]: ").strip().upper()
            if resp_abrir in ("S", "SIM", "Y", "YES"):
                for arq in arquivos_para_abrir:
                    print(f"  📂 Abrindo {arq.name} no Excel...")
                    abrir_arquivo_excel(arq)

    finally:
        desativar_ambiente_teste()
        print("\n🔒 Ambiente de testes desativado. Caminhos oficiais restaurados com segurança.")


if __name__ == "__main__":
    executar_teste_kpi_zerado()
