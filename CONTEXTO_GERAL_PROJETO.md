# 🏭 DOCUMENTO DE CONTEXTO MESTRE & ARQUITETURA TÉCNICA (HANDOFF PARA IA)
> **Projeto:** Automação e Consolidação Cirúrgica de KPIs Industriais PECVAL  
> **Unidade:** Toyoda Gosei / Pecval Indústria Ltda. (Indaiatuba/SP)  
> **Repositório GitHub:** `JV2626/KPI-PECVAL`  
> **Destinatário:** Antigravity / Agentes de IA / Desenvolvedores Autônomos  
> **Versão da Arquitetura:** V3 — Blindada, Poka-Yoke, Multi-Engine (Novembro/2026)

---

## 1. VISÃO GERAL DO PROJETO

### 1.1 Objetivo do Sistema
Este sistema automatiza diariamente a coleta, validação Poka-Yoke, sanitização de dados, unificação e injeção cirúrgica de apontamentos de produção diária (`.xlsm`) nos arquivos oficiais de **Indicadores Produtivos (KPIs)** (`.xlsx`) da fábrica da PECVAL (Toyoda Gosei), cobrindo quatro setores industriais essenciais:
1. **01. Injeção Plástica** (Plantas F1, F2-Lado A e F2-Lado B)
2. **02. Montagem Industrial** (Bancadas e Linhas Lado A-B)
3. **03. Pintura Manual** (Cabines de Banho e Pintura Manual)
4. **04. Pintura Automática** (Linhas Robóticas de Pintura)

### 1.2 O Maior Desafio Técnico & A Regra Fundamental
Os arquivos de KPI Oficial da PECVAL (`.xlsx`) são planilhas corporativas extremamente complexas, contendo dezenas de abas, **Dashboards Gerenciais, Gráficos de Pareto, Tabelas Dinâmicas (Pivot Tables), Slicers, Botões VBA e cadeias de cálculo dinâmicas**.
* **REGRA FUNDAMENTAL:** **JAMAIS salvar esses arquivos usando `openpyxl.save()` diretamente!** O OpenPyXL reconstrói o arquivo do zero e corrompe/elimina gráficos, conexões de dados e tabelas dinâmicas, gerando o erro de "Planilha Danificada / Recuperada" no Excel.
* **SOLUÇÃO ADOTADA:** Arquitetura Tripla de Gravadores Não-Destrutivos:
  1. **Motor 1 (Padrão Windows): `EXCEL_COM` via `win32com.client`** com threading lock (`_COM_EXCEL_LOCK`), inicialização STA (`pythoncom.CoInitialize()`) e suspensão temporária de recálculo (`excel_app.Calculation = -4135`).
  2. **Motor 2: `WOLFXL_RUST` (`wolfxl.load_workbook(modify=True)`)** caso disponível.
  3. **Motor 3 (Fallback Universal): `_gravar_cirurgico_openxml_zip`**, que descompacta o `.xlsx` como ZIP em memória RAM, altera **APENAS o XML da aba `'APONTAR AQUI'`** via `xml.etree.ElementTree` e remonta o ZIP mantendo 100% de todos os outros arquivos, gráficos, estilos e relações intactos.

---

## 2. ESTRUTURA COMPLETA DE ARQUIVOS E RESPONSABILIDADES

```
AutomacaoKPI/
├── Caminhos/
│   ├── __init__.py
│   └── config_caminho.py          # Localizador resiliente de rede V:/UNC, normalizador e mapeador de setores
├── ValidadorDia/
│   ├── __init__.py
│   └── validador_kpi.py           # Diagnóstico de KPIs, sanitização universal, Poka-Yoke e injeção não-destrutiva
├── Setores/
│   ├── Injecao/
│   │   ├── __init__.py
│   │   └── setor_injecao.py       # Extração consolidada F1/Lado A/Lado B (139 colunas)
│   ├── Montagem/
│   │   ├── __init__.py
│   │   └── setor_montagem.py      # Extração da Montagem (116 colunas, preservação Bancada 30)
│   └── Pintura/
│       ├── __init__.py
│       ├── Manual/
│       │   └── pinturaManu.py     # Extração Pintura Manual (147 colunas, mapeamento de turnos e tempos)
│       └── Automatico/
│           └── pinturaAuto.py     # Extração Pintura Automática Robótica (146 colunas)
└── Executar/
    ├── __init__.py
    ├── rodar_todos.py             # Cockpit Central: Hub Web (Porta 8085) + Menu Terminal CLI + APIs REST
    └── testar_kpi_zerado.py       # Módulo standalone para homologação em pasta de KPIs zerados
```

---

## 3. DETALHAMENTO DE CADA ARQUIVO & SUAS REGRAS

### 3.1 `Caminhos/config_caminho.py`
* **Caminhos de Rede Suportados:**
  - Unidade Mapeada: `V:\` (ex: `V:\REPORT DE PRODUCAO\1. INDICADORES PRODUTIVOS PECVAL\APONTAMENTOS` e `V:\Producao\Confidential\...`).
  - Caminho UNC Direto: `\\10.72.110.56\Departamentos\...`.
  - Fallbacks locais e rede João: `V:\Joao ADM-Prod\PECVAL`.
* **Funções-Chave:**
  - `obter_pastas_busca_apontamentos(setor_id, mes, ano)`: Localiza onde estão os `.xlsm` de cada setor.
  - `localizar_kpi_oficial(setor_id, mes, ano)`: Encontra a planilha oficial de cada setor para o mês/ano vigente.
  - `listar_apontamentos_disponiveis(setor_id, mes, ano)`: Varre os arquivos do mês e agrupa por data (`dt.date`), ignorando arquivos temporários `~$` e pastas de backup/antigas via `eh_diretorio_legado_ou_ignorado()`.
  - `normalizar_texto(texto)`: Remove acentos e converte para caixa alta, permitindo busca resiliente mesmo com variações de digitação (`INJEÇÃO` == `INJECAO`).
  - `obter_caminho_kpi_teste(kpi_oficial)`: Gera o caminho seguro em `PASTA_TESTES_KPI` com o sufixo `- TESTE.xlsx`.
  - `ativar_ambiente_teste(pasta)` / `desativar_ambiente_teste()`: Redireciona a busca e gravação para uma pasta local isolada (`PASTA_OVERRIDE_KPIS`).

### 3.2 `ValidadorDia/validador_kpi.py`
O cérebro do sistema:
* **`sanitizar_valor_excel(val)` (CRÍTICO!):**
  - Trata conversão de tipos de dados para o Excel sem distorções de tempo.
  - **Blindagem do Sistema de Datas 1904:** No Excel, se uma duração ou fórmula for gerada com o calendário de 1904 (Mac ou fórmulas de apontamento), $1.462\text{ dias} \times 24\text{h} = 35.088\text{ horas}$ são indevidamente somadas. O sanitizador intercepta:
    1. `dt.datetime` com ano $< 1950$ (anos 1899, 1900, 1904) $\rightarrow$ converte para horário puro `HH:MM:SS`.
    2. `dt.timedelta` com horas $\ge 35.088\text{h}$ $\rightarrow$ desconta $35.088\text{h}$ e retorna `HH:MM:SS`.
    3. `float` em torno de 1.462 (ex: $1.462,090278$) $\rightarrow$ extrai apenas a fração decimal como segundos do dia.
    4. `str` com mais de 35.000 horas (ex: `'35090:10'`) ou strings com datas legadas $\rightarrow$ normaliza para o horário real.
  - Converte erros de células (`#VALOR!`, `#DIV/0!`, `#REF!`) para string vazia `""` para não corromper fórmulas do KPI.
* **`diagnosticar_kpi_setor()`:**
  - Inspeciona o KPI sem abrir no Excel: usa leitor OpenXML cirúrgico (`extrair_datas_coluna_a_zip`) para ler a Coluna A direto do XML.
  - Retorna `dias_no_kpi`, `ultima_linha_preenchida`, `proxima_linha_livre`, `dias_pendentes` e `status_mes` (`VAZIO_AGUARDANDO_APONTAMENTOS` vs `EM_ANDAMENTO`).
* **`inserir_abaixo_ultima_linha_kpi()`:**
  - Valida duplicidade do dia: se o dia já existe e `permitir_atualizar_mesmo_dia` for `False`, bloqueia com `bloqueado_duplicidade = True`.
  - Cria backup prévio com timestamp em `PASTA_BACKUPS` (`Backups_KPI/`).
  - Executa a gravação em staging local (`tempfile.mkdtemp` em `C:\Temp`) para não travar a rede SMB durante a escrita.
  - Durante o motor Win32COM: define `excel_app.Calculation = -4135` (`xlCalculationManual`) para não disparar recálculos pesados a cada célula, recalculando apenas ao final antes de salvar.
  - Move de volta para a rede via atomic replace (`os.replace`).

### 3.3 `Setores/Injecao/setor_injecao.py`
* **Setor:** 01. Injeção Plástica (F1 / Lado A / Lado B).
* **Consolidação:** Reúne até 3 arquivos diários (`F1.xlsm`, `LADO A.xlsm`, `LADO B.xlsm`), varrendo as abas `ABA 1` a `ABA 5` a partir da linha 12.
* **Colunas:** 139 colunas no KPI (Coluna A = Data ISO, Colunas B..EI = Dados).
* **Regra Poka-Yoke & Desempenho:**
  - Exige `kanban`, `maquina` ou `operador`.
  - **Parada inteligente:** Se encontrar 100 linhas vazias consecutivas (`vazias_consecutivas >= 100`), dá `break` imediatamente, impedindo a varredura desnecessária de milhares de linhas em branco do final da planilha.
  - **Tempo de execução:** Reduzido de 15 minutos para menos de 45 segundos.

### 3.4 `Setores/Montagem/setor_montagem.py`
* **Setor:** 02. Montagem Industrial (Lado A-B).
* **Abas:** `ABA 1`, `ABA 2`, `ABA 3` a partir da linha 12.
* **Colunas:** 116 colunas no KPI (Coluna A = Data, Colunas B..DL = 115 colunas úteis visíveis).
* **Mapeamento De-Para:** `MAPA_DE_PARA_MONTAGEM` suprime com precisão as 27 colunas ocultas do apontamento.
* **Regra de Preservação da Bancada 30 (Regra de Ouro):**
  - Se a coluna de quantidade tiver peças produzidas (`qtd_val > 0`), **A LINHA NUNCA É DESCARTADA**, mesmo que o operador tenha deixado o campo de Kanban em branco.
  - Reconhece identificação em `bancada` (Col E), `maquina` (Col F), `kanban` (Col C) e `operador` (Col B).
  - Garante a contagem perfeita de todas as 21.555 peças, 22 peças esperadas e 42 minutos de tempo produtivo.
* **Modo 100% Colar Valores:** Zero fórmulas coladas do apontamento; tudo entra limpo e sanitizado.

### 3.5 `Setores/Pintura/Manual/pinturaManu.py`
* **Setor:** 03. Pintura Manual (Leandro e Marcos Manual).
* **Colunas:** 147 colunas no KPI (A = Data, B..EQ = 146 colunas de dados).
* **Detecção Dinâmica de Colunas Ocultas:** `obter_colunas_ocultas_aba()` lê as tags `<col>` no XML da aba para ignorar colunas ocultas no `.xlsm`.
* **Alinhamento Dinâmico de Turno:** Detecta se o turno está na coluna 71 ou 72 e ajusta o vetor de destino para não deslocar paradas e horários.
* **Fórmula Protegida:** A Coluna 77 (BZ no KPI) é sempre forçada para `""` (coluna inativa).

### 3.6 `Setores/Pintura/Automatico/pinturaAuto.py`
* **Setor:** 04. Pintura Automática Robótica (Rodrigo e Emerson Automático).
* **Colunas:** 146 colunas no KPI.
* **Mapeamento:** `MAPA_DE_PARA_PINTURA_AUTO` mapeia as colunas de robôs e cabines sem pular dados de refugo ou produção.

### 3.7 `Executar/rodar_todos.py`
O orquestrador unificado:
* **Interface Web (Cockpit Glassmorphism):**
  - Servidor embutido `http://127.0.0.1:8085` usando `http.server` puro (sem dependência de Flask/Django).
  - Seleção interativa de Mês, Setor e Dia com status de cada dia (`✅ CONSTA NO KPI` vs `⏳ PENDENTE`).
  - 4 Cards visuais com diagnóstico em tempo real dos 4 setores.
  - 3 Botões de Ação: `1. Pré-Validar Dia`, `2. Gravar KPI Teste`, `3. Lançar no KPI Oficial`.
  - Botão de restauração rápida de backups (`Restaurar Último Backup`).
* **Menu CLI no Terminal:** Opções interativas numeradas de `[1]` a `[6]`, incluindo `[6] 🧪 Testar em Pasta de KPIs Zerados`.

---

## 4. CASOS DE BORDA HISTÓRICOS & SUAS SOLUÇÕES DEFINITIVAS

### Caso 1: As 35.100 Horas nas Paradas (`35106:00`, `35100:20`, `35090:10`)
* **Sintoma:** O KPI Teste mostrava paradas absurdas de 35.100 horas, totalizando 105.379 horas e destruindo o gráfico de Pareto.
* **Causa Raiz:** O sistema de data 1904 do Excel ($1.462\text{ dias} \times 24\text{h} = 35.088\text{ horas}$). Uma parada de 2h10min era gravada como $35.088 + 2\text{h }10\text{min} = 35.090:10$.
* **Solução:** `sanitizar_valor_excel` no `validador_kpi.py` desconta 35.088 horas de qualquer `timedelta`, `float`, `datetime` ou `string` que exceda 35.000 horas, retornando a duração real (`18:25`, `12:20`, `02:10`).

### Caso 2: Sumiço das 3 Peças da Bancada 30 (21.555 vs 21.552 peças)
* **Sintoma:** Na aba de qualidade da Montagem, a produção do dia 03/11 caiu de 21.555 para 21.552 (3 peças a menos).
* **Causa Raiz:** O filtro `if not kanban or not maquina:` descartava a linha inteira da Bancada 30 porque o operador deixou o Kanban vazio.
* **Solução:** Regra Poka-Yoke: se `qtd_val > 0`, a linha **JAMAIS é descartada**. Todas as 21.555 peças, as 22 esperadas e os 42 minutos de tempo produtivo foram recuperados.

### Caso 3: Travamento de 15 Minutos na Injeção
* **Sintoma:** Ao rodar Injeção ou "Todos de uma vez", o script ficava mais de 15 minutos rodando sem dar erro.
* **Causa Raiz:** Uma checagem de 68 colunas (`col 70 a 138`) dava falso positivo nas fórmulas vazias do template da Injeção em 15 abas, fazendo o Python extrair mais de 50.000 linhas em branco; somado ao Excel COM recalculando todas as fórmulas a cada colagem.
* **Solução:** Remoção do scan falso de 68 colunas; adição de cutoff de 100 linhas em branco consecutivas (`vazias_consecutivas >= 100: break`); e ativação de `excel_app.Calculation = -4135` (`xlCalculationManual`) durante a escrita.

### Caso 4: Alerta Falso de "KPI Vazio (Início de Mês)"
* **Sintoma:** No início do mês (ex: 11/2026), o Hub mostrava "KPI Vazio (0 linhas com dados)" e o usuário achava que os dados dos meses anteriores tinham sumido.
* **Causa Raiz:** Cada mês possui um arquivo `.xlsx` separado na rede. No início do mês, o arquivo de Novembro realmente só possui as linhas 1 a 3 (cabeçalho) e zero apontamentos lançados.
* **Solução:** O Hub informa com clareza que o aviso refere-se exclusivamente ao mês selecionado, mantendo os meses anteriores intactos em seus próprios arquivos.

---

## 5. RECOMENDAÇÕES PARA FUTUROS DESENVOLVIMENTOS & AGENTES DE IA

1. **Nunca alterar nomes de colunas ou abas fixas:**
   - A aba alvo em todos os KPIs é `'APONTAR AQUI'` (algumas planilhas antigas possuem um espaço no final `'APONTAR AQUI '`, que a função `resolver_nome_aba()` já resolve automaticamente).
2. **Ao modificar leitores de planilha:**
   - Se utilizar Calamine, utilize sempre `skip_empty_area=False`. Se for `True`, colunas vazias à esquerda (como a Coluna A) são suprimidas e deslocam todas as colunas de dados para a esquerda!
3. **Ao manipular arquivos de rede (`V:\` ou `\\10.72.110.56`):**
   - Não use `os.fsync()` em compartilhamentos SMB/CIFS no Windows (causa `Errno 9 - Bad file descriptor`).
   - Use cópia em staging local (`C:\Temp`) e transferência atômica com retentativas.
4. **Sincronização com o repositório:**
   - A branch padrão é `main`. Todos os commits devem ser testados via `python3 -m py_compile` antes de dar push.
