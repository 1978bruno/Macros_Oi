# Macros_Oi
Repositório de códigos VBA das macros desenvolvidas entre 2001 e 2008 para a Telemar Oi.

##**Macro - Internacional**
<img width="1048" height="602" alt="image" src="https://github.com/user-attachments/assets/a7528fee-078c-4e1a-a78c-40b360d5ade5" />
# Documentação Técnica e Funcional: Aplicação VBA Excel

---

## 1. Visão Geral da Solução

* **Arquivo:** `Macro - Internacional (Tudo).xls`
* **Domínio de Negócio:** Gestão e Monitoramento de Tráfego de Voz Internacional (*Backbone Internacional / Telecomunicações*).
* **Finalidade Técnica:** Atua como o **motor de consolidação de séries temporais (Data Mart / Repositório Histórico)** que alimenta a suíte de gráficos executivos (como o `InterGraf`). O sistema realiza ingestão incremental diária de dados operacionais brutos gerados pelo COI (*Centro de Operações de Rede*), calcula a distribuição percentual de rotas (**LCR Share**) e constrói rankings dos principais destinos ofensores em tráfego de saída (`SAINTE`), entrada (`ENTRANTE`) e tráfego visitante (`ROOMING`).

---

## 2. Modelagem e Estrutura de Dados (Abas)

```
       [ Relatório Externo COI ]
                   │
                   ▼ (Módulo2 - Ingestão Incremental Diária)
  ┌────────────────┬────────────────┬────────────────┐
  │     SAINTE     │    ENTRANTE    │    ROOMING     │  <- Séries Temporais (Colunas Triplas / Dia)
  └────────────────┴────────────────┴────────────────┘
                   │
                   ▼ (Módulo3 - Consolidação & Agregação)
           ┌───────────────┐
           │     FACE      │  <- Painel Executivo / Rankings TOP 20 e TOP 10
           └───────────────┘

```

| Aba | Função Principal | Esquema de Dados e Dinâmica de Colunas |
| --- | --- | --- |
| **`FACE`** | **Painel Executivo / Front-end** | Apresenta os rankings dos principais destinos de tráfego, metadados da última execução (data/hora, usuário e hostname da estação). |
| **`SAINTE`** | **Histórico de Tráfego Internacional de Saída** | • **Chaves Primárias (A:C):** `[País, Rota de saída, Telefonia]` (ex.: Fixa/Móvel).<br>

<br>• **Série Temporal (D em diante):** A cada novo dia importado, são adicionadas **3 colunas contíguas**: `[OK% (Completamento), TCH (Tentativas), LCR (% Share da Rota)]` sob o cabeçalho da data. |
| **`ROOMING`** | **Histórico de Tráfego Internacional de Roaming** | Mesma estrutura multidimensional da aba `SAINTE`, segmentando chamadas de usuários em roaming. |
| **`ENTRANTE`** | **Histórico de Tráfego Internacional Entrante** | • **Chaves (A:B):** `[Localidade, Operadora]` (ex.: `AC-RBO-RIO BRANCO`, `BRASIL TELECOM`).<br>

<br>• **Série Temporal:** Adiciona blocos diários de colunas para métricas de chamadas recebidas do exterior. |

---

## 3. Arquitetura de Módulos VBA

### 3.1. `Módulo1` — Inicialização da Aplicação

* **`Sub auto_open()`:** Executado no momento da abertura do arquivo. Dispara a exibição da tela inicial (`UserForm1.Show 0`), pausando por 3 segundos via `Application.Wait` antes de descarregá-la (`Unload UserForm1`).

---

### 3.2. `Módulo2` — Ingestão Incremental Diária e Cálculo de Roteamento

Este módulo implementa o processo de ETL diário. Ele recebe os arquivos brutos do COI e os anexa às colunas da série temporal correspondente.

#### Sub-rotinas de Gatilho:

* **`Sub sainte()`:** Define `PLANREF = "SAINTE"` e chama `internacional`.
* **`Sub entrante()`:** Define `PLANREF = "ENTRANTE"` e chama `internacional`.
* **`Sub pan()`:** Define `PLANREF = "ROOMING"` e chama `internacional`.

#### Detalhamento do Processo `Sub internacional()`:

```
[ Início ]
    │
    ▼
1. Solicita arquivo diário COI via diálogo do Windows (Application.GetOpenFilename)
    │
    ▼
2. Copia aba de dados do arquivo externo para a pasta de trabalho atual como aba temporária
    │
    ▼
3. Localiza a data de referência no arquivo e valida consistência cronológica:
   - Checa se o dia importado é exatamente D+1 do último dia registrado
   - Emite alerta ao operador caso haja quebra de sequência de datas
    │
    ▼
4. Identifica dinamicamente as colunas de origem por cabeçalho:
   - "País" / "Localidade", "Rota de saída" / "Operadora", "Telefonia", "OK%", "TCH"
    │
    ▼
5. Cria o novo bloco de 3 colunas na aba de destino (SAINTE, ENTRANTE ou ROOMING)
    │
    ▼
6. Loop de Upsert (Linha a Linha):
   - Descarta agregados e registros inválidos ([TOTAL], NI-NI-NÃO IDENTIFICADA, EX-EXTERIOR, ADMI)
   - Localiza a combinação chave existente OU insere novo registro ao final
   - Grava OK% e TCH do dia
    │
    ▼
7. Cálculo Matemático do LCR Share (% de tráfego por operadora parceira):
   - Totaliza TTCH = ∑(TCH do destino)
   - Calcula LCR = (TCH da Rota / TTCH) formatado como "0.00%"
    │
    ▼
8. Limpa aba temporária e remove avisos de tela

```

---

### 3.3. `Módulo3` — Agregação e Atualização do Painel Executivo (`FACE`)

Responsável por sintetizar os dados analíticos das abas históricas e compor as tabelas de ranking.

#### Declarações de API do Windows (Win32):

```vba
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

```

#### Procedimentos:

* **`Sub atu_sainte_geral()` / `Sub atu_sainte_pan()`:** Direcionam a consolidação para as abas `SAINTE` ou `ROOMING`.
* **`Sub atualiza_sainte()`:**
1. Cria uma aba temporária para isolar pares distintos de `[País, Telefonia]`.
2. Localiza a coluna com a data mais recente (`UTCOL`).
3. Soma o volume total de chamadas (`TTCH`) somando todas as rotas de cada destino.
4. Ordena os destinos pelo volume de tráfego de forma decrescente (`xlDescending`).
5. Copia os **TOP 20** maiores ofensores para o painel `FACE`:
* Para `SAINTE`: Cola valores a partir de `R10:T29`, grava data em `Q8` e timestamp em `Q30`.
* Para `ROOMING`: Cola valores em `C10:D29`, reorganiza colunas e grava data em `B8` e timestamp em `B30`.


6. Deleta a aba temporária e invoca `Ult_Atua`.


* **`Sub atualiza_entrante()`:**
* Realiza a mesma agregação para tráfego de entrada, gerando o **TOP 10** de localidades em `FACE!R35:S44`.


* **`Sub Ult_Atua()` (Auditoria e Rastreabilidade):**
* Invoca `GetUserNameA` e `GetComputerNameA`.
* Preenche a célula `FACE!A4` com o registro de auditoria:
> *"{Usuário} usando a máquina {Nome_do_Computador}"*





---

### 3.4. `Módulo4` — Ferramentas de Filtragem Rápida

Implementa atalhos de filtragem rápida por contexto de seleção ativa na planilha:

* **`Sub Filtro_Sainte()`:** Captura o valor da célula atualmente selecionada pelo usuário (`Busca = Selection`), navega até a aba `SAINTE` e aplica o `AutoFilter` na Coluna 1 (`Field:=1`).
* **`Sub Filtro_Entrante()`:** Aplica o mesmo comportamento de filtro contextual sobre a aba `ENTRANTE`.

---

## 4. Dicionário de Termos e Regras de Negócio de Telecom

* **TCH (*Total Calls Handled*):** Total de tentativas de chamadas telefônicas direcionadas à rota/destino.
* **OK% (*Answer-Seizure Ratio - ASR / Completamento*):** Percentual de chamadas efetivamente completadas e atendidas em relação às tentadas.
* **LCR (*Least Cost Routing Share*):** Percentual do tráfego do país que foi roteado através daquela operadora parceira específica no dia.
* **Ofensores:** Destinos internacionais ou rotas que concentram os maiores volumes de chamadas ou que apresentam anomalias/degradação de completamento.

---

## 5. Diagnóstico Técnico e Pontos de Modernização

1. **Compatibilidade com MS Office de 64 bits:**
* As declarações `GetComputerName` e `GetUserName` no `Módulo3` utilizam chamadas Win32 de 32 bits. Em ambientes Excel de 64 bits modernos, isso causa erro de compilação.
* *Solução recomendada:* Adicionar o modificador `PtrSafe` ou substituir integralmente pelas funções nativas do ambiente VBA:
```vba
Comp_Name = Environ("COMPUTERNAME")
USU_Name = Environ("USERNAME")

```




2. **Desempenho de Carga (Otimização de Operações em Planilha):**
* As rotinas `internacional`, `atualiza_sainte` e `atualiza_entrante` iteram linha a linha sobre a planilha utilizando `.Select` e loops aninhados `Do Until`.
* *Solução recomendada:* Carregar os intervalos de dados em matrizes de memória (`Variant Arrays`) ou utilizar o objeto `Scripting.Dictionary` para realizar as operações de agrupamento (*GROUP BY*) e agregação de dados (*SUM*), o que reduz o tempo de execução de vários minutos para frações de segundo.


3. **Tratamento de Exceções em Manipulação de Abas:**
* A macro cria e exclui abas temporárias (`Sheets.Add`, `Sheets(ATUA).Delete`). Caso ocorra um erro antes do encerramento da sub-rotina, a planilha fica suja com abas órfãs. Recomenda-se adicionar um bloco de tratamento estruturado `On Error GoTo TrataErro` garantindo a limpeza e a restauração de `Application.DisplayAlerts = True` e `Application.ScreenUpdating = True`.

##**Macro - Grafico**
<img width="904" height="736" alt="image" src="https://github.com/user-attachments/assets/c719913a-9a6c-4bd6-9764-48990e35e2a9" />
# Documentação Técnica e Funcional: Aplicação VBA Excel

---

## 1. Identificação do Projeto e Metadados

* **Nome da Aplicação:** `INTERGRAF-V.2.2.1` *(conforme registrado na `Application.StatusBar` no `Módulo3`)*
* **Arquivo:** `Macro - Grafico (Tudo).xls`
* **Domínio de Negócio:** Gestão e Monitoração de Tráfego Telefônico Internacional (*Backbone Internacional / Telecom*)
* **Propósito:** Consolidação de dados diários de tráfego saindo do Brasil (`SAINTE`), geração de ranking com o **TOP 20 de Países/Destinos Ofensores**, cruzamento com metas de qualidade (**ASR**) e regras de roteamento de menor custo (**LCR**), alimentando um painel gráfico com filtros em cascata.

---

## 2. Visão Geral da Arquitetura de Dados (Abas)

| Nome da Aba | Status de Exibição | Descrição e Papel no Sistema |
| --- | --- | --- |
| **`Grafico`** | **Visível** | **Dashboard Principal da Aplicação.** Contém: <br>

<br>• Tabela do ranking `TOP20 - PAISES` (`[Posição, País, Telefonia, TCH]`)<br>

<br>• Objeto de visualização `ChartObjects("Gráfico 1")`<br>

<br>• Três controles suspensos de formulário: `Drop Down 5` (País), `Drop Down 6` (Telefonia) e `Drop Down 7` (Rota de Saída)<br>

<br>• Tabela de opções prioritárias de rota (**LCR**) nas linhas 29 a 33 |
| **`Base`** | **Visível** | **Tabela de Fatos Histórica.** Armazena a massa de dados consolidada importada do relatório COI. Colunas: `[Data, Pais, Rsaida, Telefonia, OK%, TCH, LCR, ASR]`. Alimenta diretamente as séries dinâmicas do gráfico e as seleções de autofiltro. |
| **`ASR`** | **Oculta** (`Visible = False`) | **Matriz de Metas de Completamento.** Título de referência: *"POP DE MONITORAÇÃO DE TRÁFEGO DO BACK BONE INTERNACIONAL - METAS POR TELEFONIA PARA DESTINOS INTERNACIONAIS - ASR'S"*. Chave de busca composta: `PAIS & TELE` (ex.: `ALEMANHAFIXA`). Contém o percentual alvo de completamento (`44.4%`, `60.0%`, etc.). |
| **`ASR - ANTIGA`** | **Visível** (Apoio legado) | Tabela histórica prévia de ASR e compromissos com carriers parceiros (`SPRI`, `KPNH`, `IDTE`). |
| **`LCR`** | **Oculta** (`Visible = False`) | **Matriz de Menor Custo (Least Cost Routing).** Contém o plano ordenado de contingência e roteamento por destino (`1ª Opção`, `2ª Opção`, `3ª Opção`, `4ª Opção`, `5ª Opção`) e observações/comentários de split (ex.: `50% MCIP / 50% MCI4`, `DEUT`, `TLGL`, etc.). |

---

## 3. Análise Detalhada dos Módulos VBA

### 3.1. `Módulo1` — Mecanismo Principal de ETL (`Sub Import`)

Rotina central responsável pela extração, transformação e carga dos dados operacionais.

#### Fluxo Operacional:

1. **Seleção da Fonte Externa:**
* Abre o diálogo de seleção de arquivo (`Application.GetOpenFilename`) buscando o arquivo diário `Macro - Internacional.xls` gerado pelo COI (*Centro de Operações de Rede*).
* Exibe o formulário de carregamento `UserForm2` de forma não modal (`UserForm2.Show 0`).


2. **Setup do Ambiente e Purga:**
* Desativa atualização de tela e alertas (`ScreenUpdating = False`, `DisplayAlerts = False`).
* Torna a aba `ASR` visível para permitir operações de busca.
* Limpa os dados históricos remanescentes na aba `Base` (`Range("6:65000").Delete`).


3. **Extração em Lote (Loop de 20 Destinos Ofensores):**
* Lê a aba `FACE` do arquivo de origem nas posições `R10C18:R29C19` para capturar os 20 principais destinos (`PAIS` e `TELE`).
* Localiza o bloco de dados correspondente na aba `SAINTE`.
* Itera em janelas de 8 dias/períodos retroativos (`Do Until DIA_ANTE = 24`, com passo de 3 colunas) copiando:
* Chaves: `País`, `Rota de Saída (Rsaida)`, `Telefonia`.
* Métricas: `Data`, `TCH` (Tentativas de Chamada), `OK%` (Completamento), `LCR%`.


* **Cruzamento com Metas de ASR:** Concatena `BUSCA = PAIS & TELE`, pesquisa na coluna A da aba `ASR` e injeta a meta na coluna 8 (`ASR`) da `Base`.
* **Detecção de Fins de Semana:** Avalia `Weekday(DT, vbMonday)` para identificar sábados (`6`) e domingos (`7`), guardando os índices de barra (`BL1`, `BL2`, `BL3`) para estilização posterior do gráfico.


4. **Tratamento e Higienização de Dados:**
* Ordena a base por `[Data, País, Rota de Saída]`.
* Converte strings percentuais da coluna `G` para valores numéricos reais (`Selection.Replace "%" -> ""`).
* **Purga de Inatividade (Clean Up de Rotas sem Tráfego):** Identifica e exclui sequências de registros onde o tráfego foi zero/nulo ao longo de todo o período analisado (`If cont = 7 Then Rows.Delete`).


5. **Atualização do Dashboard e Gráficos:**
* Copia a nova tabela do TOP 20 para a aba `Grafico` (`Range("Q7:T30")` para `A1:D24`).
* Fecha o arquivo de origem (`Workbooks(RELCOO).Close`).
* Extrai a lista exclusiva e ordenada de países para as células de apoio (`Grafico!$V$2:$V$n`).
* Configura os intervalos e células de ligação dos Dropdowns nos controles ActiveX/Formulário:
* `Drop Down 5` (Países): Lista `$V$2:$V$n`, vinculada à célula `$V$1`.
* `Drop Down 6` (Telefonia): Lista `$X$2:$X$10`, vinculada à célula `$X$1`.
* `Drop Down 7` (Rota de Saída): Lista `$Z$2:$Z$10`, vinculada à célula `$Z$1`.


* Redefine os limites dinâmicos de dados das séries do `Gráfico 1`:
* **Série 1:** `TCH` (Volume de tráfego) — Estilizada em tom de destaque (cor 46) com barras de fim de semana diferenciadas em tom dourado/amarelo (cor 44).
* **Série 2:** `OK%` (Completamento real).
* **Série 3:** `LCR%` (Participação do tráfego).
* **Série 4:** `ASR` (Meta de qualidade).


* Oculta novamente a aba `ASR`.
* Fecha `UserForm2` e exibe `UserForm1` (splash screen) durante 3 segundos antes de concluir.



---

### 3.2. `Módulo2` — Interatividade do Dashboard e Regras de LCR

Gerencia a navegação dinâmica do operador no painel e o vínculo com as rotas de menor custo.

* **Variáveis Públicas Globais:** `PAIS`, `RSAI`, `TELE`.
* **`Sub filtro_1()` (Evento do Drop Down de País):**
* Disparada quando o operador seleciona um novo país.
* Reseta seleções dependentes (`Cells(1, 24) = "1"`, `Cells(1, 26) = "1"`).
* Filtra a aba `Base` por `PAIS` e extrai os tipos de telefonia disponíveis (`FIXA` ou `MOVEL`) para a coluna `X` (`Grafico!$X$2`), ajustando dinamicamente o preenchimento do segundo dropdown (`Drop Down 6`).
* Dispara `Sub LCR()`.


* **`Sub filtro_2()` (Evento do Drop Down de Telefonia):**
* Disparada quando o operador altera a telefonia (`FIXA` / `MOVEL`).
* Reseta o terceiro nível (`Cells(1, 26) = "1"`).
* Aplica filtro duplo na base (`PAIS` e `TELE`) e extrai as rotas parceiras ativas (`Rsaida`) para a coluna `Z` (`Grafico!$Z$2`), atualizando o dropdown de rotas (`Drop Down 7`).
* **Realce Visual na Tabela TOP 20:** Varre o intervalo de linhas 4 a 23 da tabela resumo; a linha correspondente ao país/tipo selecionado recebe preenchimento cinza (`ColorIndex = 15`), restaurando as demais para branco (`ColorIndex = 2`).
* Dispara `Sub LCR()`.


* **`Sub filtro_3()` (Evento do Drop Down de Rota de Saída):**
* Aplica o filtro triplo simultâneo na base: `Field 2 = PAIS`, `Field 3 = RSAI` e `Field 4 = TELE`.
* Isola a performance específica daquela operadora parceira nas séries temporais do gráfico.


* **`Sub LCR()` (Consulta à Matriz de Roteamento):**
* Torna a aba `LCR` temporariamente visível.
* Localiza a célula correspondente à concatenação `PAIS & TELE` na coluna C.
* Extrai as 5 opções de rota contratadas (`OP1` a `OP5`, colunas 6 a 10) e os seus comentários explicativos embutidos (`.Comment.Text`).
* Despeja os dados diretamente no dashboard da aba `Grafico`:
* Nomes das operadoras: `Cells(29, 7)` até `Cells(33, 7)` (G29:G33).
* Observações/Percentuais de split: `Cells(29, 9)` até `Cells(33, 9)` (I29:I33).


* Oculta a aba `LCR` novamente.



---

### 3.3. `Módulo3` — Ciclo de Vida da Aplicação

* **`Sub auto_open()`:** Injeta a assinatura institucional do autor na barra inferior do Excel durante a inicialização:
```vba
Application.StatusBar = "INTERGRAF-V.2.2.1 - Copyright © 2007 - Bruno Rabelo Rocha - Todos os direitos reservados"

```


* **`Sub auto_close()`:** Devolve o controle da barra de status ao Excel (`Application.StatusBar = ""`).

---

### 3.4. Formulários de Interface (`UserForms`)

1. **`UserForm1`:** Tela de conclusão de carga ("*Concluído / Sucesso*"). Permanece em tela por 3 segundos via timer assíncrono `Application.Wait (Now + TimeValue("0:00:03"))`.
2. **`UserForm2`:** Caixa de diálogo modal sem botões exibida durante a extração dos dados brutos para indicar processamento em andamento ("*Aguarde, processando...*").

---

## 4. Diferenciais Notáveis (Comparativo com Versões Anteriores / Posteriores)

Observando a evolução da família `InterGraf`:

1. **Escopo Fixo de 20 Destinos:** Nesta versão 2.2.1, o loop de importação é parametrizado estaticamente para `NREF = 20` (TOP 20), focando estritamente na rota de tráfego de saída (`SAINTE`), enquanto revisões posteriores (ex.: v2.5.5) passaram a suportar até 70 países dinamicamente com suporte adicional a tráfego de `ROAMING`.
2. **Período Fixo de Importação:** A janela de amostragem de dados é pré-calculada em blocos de colunas (`DIA_ANTE = 24`), sem o formulário manual de seleção de datas `UserForm3` introduzido em versões mais avançadas.
3. **Ausência de Trava Win32:** Esta versão não implementa validação de matrícula de rede via `advapi32.dll`, permitindo execução irrestrita em qualquer estação local.

---

## 5. Avaliação Técnica e Oportunidades de Refatoração

1. **Desacoplamento de Linhas Fixas e Magic Numbers:**
* O código referencia posições absolutas como `For COR_ATI = 4 To 23` e `Range("Q7:T30")`. A utilização de intervalos nomeados (*Named Ranges*) ou busca dinâmica da última linha (`xlUp`) evitará falhas caso o layout da tabela seja expandido.


2. **Eliminação de `.Select` e Operações de Clipboard:**
* A rotina utiliza exaustivamente sequências de `.Copy` e `ActiveSheet.Paste`. A atribuição direta de matrizes em memória (`Worksheets("Base").Range(...).Value = Worksheets("SAINTE").Range(...).Value`) reduz o tempo total de processamento em mais de 70% e evita poluição da área de transferência do Windows.


3. **Tratamento de Comentários de Célula no LCR:**
* Na rotina `LCR()`, o acesso a `Cells(LINR, ...).Comment.Text` gera erro de execução caso a célula não possua comentário cadastrado. Embora haja um `On Error Resume Next`, é boa prática programática verificar explicitamente:
```vba
If Not Cells(LINR, col).Comment Is Nothing Then
    COP = Cells(LINR, col).Comment.Text
Else
    COP = ""
End If

```

##**Macro - Intergraf**
<img width="782" height="676" alt="image" src="https://github.com/user-attachments/assets/a71d5fdc-34cc-466e-882d-d7d4654593a1" />
# Documentação Técnica e Funcional: Aplicação VBA InterGraf

---

## 1. Visão Geral do Sistema

* **Nome da Aplicação:** `INTERGRAF-V.2.5.5`
* **Tipo:** Ferramenta de Análise, Monitoração e Visualização de Tráfego de Telecomunicações Internacionais (*Backbone Internacional*)
* **Ambiente de Execução:** Microsoft Excel (.xls - Excel 97-2003 / VBA 6.0)
* **Objetivo Principal:** Automatizar a importação, normalização e cruzamento de relatórios operacionais diários de tráfego de telecom (rotas de saída, destinos, chamadas completadas, ASR e LCR), gerando dashboards dinâmicos com gráficos de evolução temporal e identificação de rotas críticas / ofensoras.

---

## 2. Arquitetura de Dados (Abas da Planilha)

| Aba | Visibilidade Padrão | Função Técnica |
| --- | --- | --- |
| **`Grafico`** | Visível | **Dashboard principal**. Contém o gráfico interativo de desempenho, controles de formulário (*Comboboxes / Dropdowns*) integrados, tabela de ranking (`TOP - PAÍSES`) e resumo de opções de roteamento (**LCR**). |
| **`Base`** | Visível | **Tabela de Fatos / Dados Históricos**. Estrutura: `[Data, Pais, Rsaida, Telefonia, OK%, TCH, TTC (Min), LCR, ASR]`. Alimenta os cálculos e as séries do gráfico. |
| **`ASR`** | Oculta via macro (`xlSheetHidden`) | **Metas de ASR** (*Answer-Seizure Ratio*). Matriz de referência contendo as metas de completamento percentual por destino e tipo de tráfego (`FIXA` / `MOVEL`). |
| **`LCR`** | Oculta via macro (`xlSheetHidden`) | **Matriz de Roteamento de Menor Custo** (*Least Cost Routing*). Contém a ordem de prioridade das rotas de saída (`1ª Opção` a `5ª Opção`), status de validação e operadoras parceiras (`MCI`, `DEUT`, `SPRI`, `TLGL`, etc.). |

---

## 3. Módulos e Componentes de Código

### 3.1. `Módulo1` — Núcleo de ETL e Importação de Dados

* **Declaração Win32 API:**
```vba
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

```


* **Controle de Acesso (RBAC simples):**
* Captura o usuário autenticado na sessão Windows (`advapi32.dll`).
* Valida a execução contra uma *whitelist* de matrículas/IDs autorizados (`92033`, `84078`, `josearf`, `T87529`, `0928`). Caso o usuário não conste na lista, aborta a execução via `MsgBox`.


* **Fluxo da Rotina `Sub Import()`:**
1. **Seleção de Fonte Externa:** Invoca `Application.GetOpenFilename` solicitando o arquivo diário/consolidado do COI (*Centro de Operações de Rede*).
2. **Parametrização:**
* Exibe o formulário modal `UserForm4` para seleção do tipo de tráfego: `SAINTE` (*Tráfego Internacional de Saída*) ou `ROAMING`.
* Exibe `UserForm3` para definição da janela temporal de análise (`DTI` - Data Inicial e `DTF` - Data Final), validando limites do intervalo com `DateDiff`.


3. **Extração e Tratamento:**
* Itera sobre a lista de países ofensores presentes na aba `FACE` da planilha de origem.
* Copia blocos de métricas (`TCH`, `TTC`, taxa de completamento, etc.) para a aba `Base`.
* Faz cruzamento (*VLOOKUP programático*) com a aba `ASR` para preencher as metas de qualidade.


4. **Higienização e Limpeza:**
* Converte strings numéricas em tipos de ponto flutuante.
* Higieniza caracteres `%` na coluna de completamento.
* **Purga de Ociosidade:** Remove rotas e carriers sem tráfego (`cont = DDFFT + 1`).
* Ordena registros crescentemente por `[Data, País, Rota]`.


5. **Atualização do Dashboard e Gráfico:**
* Atualiza a tabela Top-N na aba `Grafico`.
* Preenche as listas exclusivas e dinâmicas nas colunas de apoio `V`, `X` e `Z`.
* Configura o objeto `ChartObjects("Gráfico 1")`: vincula fontes de dados (`XValues` e `Values` das `SeriesCollection 1 a 5`), define legendas dinâmicas e aplica destaque visual diferenciado (cor laranja/ouro) para dias de final de semana (`SMN = 6 Or SMN = 7`).





---

### 3.2. `Módulo2` — Lógica de Interatividade e Filtragem

Contém as rotinas disparadas pelos controles de seleção ou eventos de planilha:

* **`Sub filtro_1()`:**
* Disparada pela mudança de **País** (`Drop Down 5`).
* Aplica `AutoFilter` na aba `Base` filtrando pelo país escolhido.
* Extrai valores distintos de **Tipo de Telefonia** (`FIXA` / `MOVEL`), popula a coluna `X` e ajusta o `ListFillRange` do segundo Dropdown (`Drop Down 6`).
* Invoca a busca de parâmetros do LCR.


* **`Sub filtro_2()`:**
* Disparada pela mudança de **Telefonia** (`Drop Down 6`).
* Filtra a base por `País` e `Telefonia`.
* Extrai as **Rotas de Saída** ativas (`Rsaida`), popula a coluna `Z` e atualiza o terceiro Dropdown (`Drop Down 7`).
* **Feedback Visual:** Realça em cinza (`ColorIndex = 15`) a linha correspondente na tabela TOP de países e redefine as demais.
* Executa a sub-rotina `LCR`.


* **`Sub filtro_3()`:**
* Filtra a base pela combinação completa: `[País] + [Rota de Saída] + [Telefonia]`.


* **`Sub LCR()`:**
* Realiza busca na aba oculta `LCR` com a chave composta `PAIS & TELE`.
* Extrai as rotas prioritárias (1ª a 5ª opção) e os comentários/observações associados às células, exibindo-os nos campos `I29:J33` da aba `Grafico`.



---

### 3.3. `Módulo3` — Eventos de Ciclo de Vida da Aplicação

* **`Sub auto_open()`:** Injeta a assinatura na barra de status do Excel:
```
INTERGRAF-V.2.5.5 - Copyright © 2007 - Bruno Rabelo Rocha - Todos os direitos reservados

```


* **`Sub auto_close()`:** Restaura a `Application.StatusBar = ""`.

---

### 3.4. `Módulo4` — Módulo Utilitário / Rascunho de Testes

* **`Sub FillArrayMulti()`:** Procedimento residual de desenvolvimento utilizado para testes de iteração de matrizes multidimensionais e depuração via `Debug.Print`.

---

### 3.5. Componentes de Interface (`UserForms`)

1. **`UserForm1`:** Tela splash / modal informativa de sucesso de execução (fechamento temporizado de 3 segundos via `Application.Wait`).
2. **`UserForm2`:** Tela de carregamento / progresso durante o processamento do arquivo externo.
3. **`UserForm3`:** Janela de entrada de parâmetros de datas (`TextBox1` a `TextBox6`), validando coerência temporal antes de avançar.
4. **`UserForm4`:** Modal de seleção de contexto de negócio (`OptionButton1` = `SAINTE`, `OptionButton2` = `ROAMING`).

---

## 4. Dicionário de Métricas e Negócio de Telecom

* **TCH (*Total Calls Handled / Tráfego Tentado*):** Volume bruto de chamadas geradas para o destino.
* **TTC (*Total Time of Conversation*):** Minutos totais faturados / conversados no enlace.
* **ASR (*Answer-Seizure Ratio*):** Relação percentual entre chamadas atendidas e tentativas de chamada (métrica padrão ITU-T para saúde de rota).
* **LCR (*Least Cost Routing*):** Tabela de decisão dinâmica para envio preferencial do tráfego através de parceiros com melhor custo-benefício.

---

## 5. Diagnóstico e Recomendações de Modernização

1. **Eliminar Dependência de DLL de 32 bits:**
* A chamada `advapi32.dll` sem a diretiva `PtrSafe` quebra a execução em versões modernas do Office de 64 bits.
* *Correção sugerida:* Substituir a API pela variável nativa do ambiente VBA: `Environ("USERNAME")`.


2. **Otimização de Performance (Remoção de `.Select` e `.Activate`):**
* O código faz uso frequente de `Sheets(...).Select` e `ActiveCell.Paste`. O acesso direto a ranges (`Worksheets("Base").Range(...).Value = ...`) tornará a rotina de importação até 10x mais rápida.


3. **Parametrização da Lista de Acesso:**
* Migrar os IDs da lista de autorização do código-fonte (*hardcoded*) para uma aba de configuração com proteção de planilha, facilitando a governança sem necessidade de alteração de código.
