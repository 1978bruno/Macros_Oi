# Macros_Oi
Repositório de códigos VBA das macros desenvolvidas entre 2001 e 2008 para a Telemar Oi.
````
**Índice**
       - Macro Internacional -
       - Macro Gráfico -
       - Intergraf -
       - Base_Intergraf -
       - Macro Oficio Consolidação LDN Evol - 
````

## **Macro - Internacional**
<img width="1048" height="602" alt="image" src="https://github.com/user-attachments/assets/a7528fee-078c-4e1a-a78c-40b360d5ade5" />

## Documentação Técnica e Funcional: Aplicação VBA Excel

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

## **Macro - Grafico**
<img width="904" height="736" alt="image" src="https://github.com/user-attachments/assets/c719913a-9a6c-4bd6-9764-48990e35e2a9" />
## Documentação Técnica e Funcional: Aplicação VBA Excel

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

## **Macro - Intergraf**
<img width="782" height="676" alt="image" src="https://github.com/user-attachments/assets/a71d5fdc-34cc-466e-882d-d7d4654593a1" />

## Documentação Técnica e Funcional: Aplicação VBA InterGraf

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

## **Base_Intergraf**
<img width="877" height="734" alt="image" src="https://github.com/user-attachments/assets/39196d75-1486-4217-b724-16f1acc23206" />

## Documentação Técnica e Funcional: Aplicação VBA Excel

---

## 1. Visão Geral da Solução

* **Arquivo:** `Base_InterGraf.xls`
* **Domínio de Negócio:** Monitoramento e Análise de Tráfego Telefônico Internacional (*Backbone Internacional / Telecomunicações*).
* **Papel no Ecossistema:** Atua como o **Data Mart Operacional Centralizado** da suíte *InterGraf*. É a base histórica robusta responsável por armazenar séries temporais diárias, calcular indicadores de desempenho e gerar a visualização consolidada na aba `FACE` (que é posteriormente consumida pelo frontend gráfico executivo `InterGraf.xls`).
* **Principais Recursos:**
* Ingestão incremental automatizada de relatórios diários de tráfego do COI (*Centro de Operações de Rede*).
* Gestão de **janela móvel histórica de 60 dias** (purga automática de datas antigas).
* Consolidação diária ou **por período customizado** de rankings parametrizáveis (**TOP 20, 30, 40, 50, 60 ou 70** destinos).
* Tráfego segmentado em: Saída Internacional (`SAINTE`), Tráfego de Entrada (`ENTRANTE`) e Usuários em Roaming Internacional (`ROAMING`).



---

## 2. Arquitetura do Modelo de Dados (Abas)

```
        [ Relatório Bruto COI (.xls) ]
                     │
                     ▼ (Módulo2: Ingestão Incremental & Upsert)
  ┌──────────────────┬──────────────────┬──────────────────┐
  │      SAINTE      │     ROAMING      │     ENTRANTE     │  <- Séries Temporais (4 Colunas/Dia)
  └──────────────────┴──────────────────┴──────────────────┘
                     │
                     ▼ (Módulo3: Agregação Diária ou por Período)
             ┌───────────────┐
             │     FACE      │  <- Rankings TOP 20 a 70 & Auditoria
             └───────────────┘

```

### Detalhamento das Abas:

| Aba | Papel / Função | Estrutura de Colunas e Dinâmica Temporal |
| --- | --- | --- |
| **`FACE`** | **Dashboard Executivo e Metadados** | • Colunas `B:F`: Ranking `TOP N ROAMING - PAISES`<br>

<br>• Colunas `H:K`: Ranking `TOP 10 - Localidades` (Entrante)<br>

<br>• Colunas `M:Q`: Ranking `TOP N - PAISES` (Sainte)<br>

<br>• Cabeçalho com data, período e registro de auditoria do operador. |
| **`SAINTE`** | **Série Temporal de Tráfego de Saída** | • **Chaves Primárias (`A:C`):** `[País, Rota de saída, Telefonia]` (Fixa/Móvel).<br>

<br>• **Bloco Diário (4 colunas contíguas por data):** `[OK%, TCH, TTC(min), LCR]`.<br>

<br>*(Armazena mais de 60 dias em colunas expandidas horizontalmente).* |
| **`ROAMING`** | **Série Temporal de Tráfego em Roaming** | Mesma estrutura multidimensional de 4 colunas por data da aba `SAINTE`, filtrando especificamente tráfego originado por parceiros/visitantes (foco em `KPNO` / `MOVEL`). |
| **`ENTRANTE`** | **Série Temporal de Tráfego Recebido** | • **Chaves Primárias (`A:B`):** `[Localidade, Operadora]` (ex.: `MG-BHE-BELO HORIZONTE`, `TELEMAR`).<br>

<br>• Colunas diárias para métricas de tráfego de entrada. |

---

## 3. Estrutura de Código e Módulos VBA

### 3.1. `Módulo1` — Ciclo de Inicialização

* **`Sub auto_open()`:** Invocado ao abrir o arquivo. Exibe a tela de abertura `UserForm1` em modo não modal (`Show 0`) por 3 segundos utilizando `Application.Wait`, finalizando com `Unload UserForm1`.

---

### 3.2. `Módulo2` — Ingestão Incremental Diária (`Sub internacional`)

Responsável pela extração do relatório externo diário do COI, validação cronológica, manutenção da janela de 60 dias, concatenação dos dados e cálculo de percentual de rota (*LCR Share*).

#### Procedimentos de Entrada:

* **`Sub sainte()`:** Processa `PLANREF = "SAINTE"`.
* **`Sub pan()`:** Processa `PLANREF = "ROAMING"`.
* **`Sub entrante()`:** Processa `PLANREF = "ENTRANTE"`.

#### Fluxo Operacional:

1. **Captura do Arquivo:** Abre a caixa de diálogo `Application.GetOpenFilename` para seleção do relatório diário do COI e exibe tela de espera `UserForm3`.
2. **Cópia da Aba Externa:** Copia a planilha para a pasta de trabalho atual como uma aba de trabalho temporária (`REFCOMP`).
3. **Validação Cronológica Estrita:**
* Localiza a data no cabeçalho do arquivo importado (`DT`).
* Compara com a data da última coluna presente na base (`Data1`).
* Valida via `DateDiff("d", Data1, DT)` se o arquivo é rigorosamente o dia seguinte (`D+1`). Caso não seja, emite alerta com opção de abortar (`vbOKCancel`).
* Suporta inserção retroativa no meio do grid caso seja importado um dia faltante anterior.


4. **Política de Retenção de Dados (Janela Móvel de 60 Dias):**
```vba
NREF = DateDiff("d", Cells(1, 4), Cells(1, CLFM - 4))
If NREF > 60 Then
    Range(Columns(4), Columns(7)).Delete
    CLFM = CLFM - 4
End If

```


Garante que a planilha não ultrapasse os limites de colunas do Excel legado (.xls suporta até 256 colunas).
5. **Mapeamento Dinâmico de Colunas:** Localiza via `Find` a posição exata no relatório de origem dos campos: `País/Localidade`, `Rota de saída/Operadora`, `Telefonia/PRD`, `OK%`, `TCH` e `TTC(Min)`.
6. **Inserção / Atualização (Upsert):**
* Descarta registros agregadores (`[TOTAL]`, `NI-NI-NÃO IDENTIFICADA`, `EX-EXT-EXTERIOR`, `ADMI`).
* No caso de `ROAMING`, filtra apenas `RSA = "KPNO"` e `TEL = "MOVEL"`.
* Realiza busca linear na base: se a combinação existir, preenche as métricas do dia; caso não exista, insere um novo destino no final da tabela.


7. **Cálculo Matemático do LCR Share (% de Distribuição da Rota):**
Para cada país e tipo de telefonia, calcula a soma total de chamadas (`TTCH`) e preenche a 4ª coluna do dia:

$$\text{LCR} = \left(\frac{\text{TCH da Rota}}{\sum \text{TCH do País e Telefonia}}\right) \times 100$$


8. **Encadeamento Automático:** Ao concluir a atualização da aba `SAINTE`, a rotina altera automaticamente `PLANREF = "ROAMING"` e executa o processamento do roaming para o mesmo arquivo antes de finalizar.
9. **Finalização:** Exibe a mensagem de sucesso `UserForm4` por 3 segundos.

---

### 3.3. `Módulo3` — Consolidação, Agregações e Painel Executivo

Contém a inteligência analítica para sintetizar os dados das séries temporais e atualizar os quadros de ranking da aba `FACE`.

#### Declarações de API do Windows (Win32):

```vba
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, nSize As Long) As Long

```

#### Procedimentos Principais:

* **`Sub atualiza_sainte()` (Consolidação Diária):**
1. Abre `UserForm2` para o usuário escolher o tamanho do ranking desejado: **TOP 20, 30, 40, 50, 60 ou 70** (`TOPREF`).
2. Cria uma aba temporária e extrai os pares únicos de `[País, Telefonia]`.
3. Varre a última data disponível na base (`UTCOL`), somando o volume de tráfego (`TCH`) de todas as rotas e calculando o volume global (`TCHF`).
4. Ordena os destinos de forma decrescente pelo tráfego total.
5. Trunca a lista na quantidade selecionada (`TOPREF`) e calcula a coluna `% Int.` (participação do país sobre o tráfego internacional total):

$$\% \text{ Int.} = \left(\frac{\text{TCH do País}}{\text{TCH Total Internacional}}\right) \times 100$$


6. Transfere os dados consolidados para a aba `FACE`:
* Para `SAINTE`: Coluna 13 (`M:Q`).
* Para `ROAMING`: Coluna 2 (`B:F`).


7. Formata dinamicamente as bordas e numeração ordinal de 1 a N via `AutoFill`.
8. Deleta a aba temporária e invoca `Ult_Atua`.


* **`Sub atualiza_periodo()` (Consolidação Multi-Diária por Intervalo Customizado):**
* Executa a consolidação para ambos os tráfegos (`SAINTE` e `ROAMING`) de forma iterativa (`For ST = 1 To 2`).
* Abre o formulário `UserForm5`, permitindo ao operador digitar a Data Inicial (`DTI`) e a Data Final (`DTF`).
* Identifica os índices de coluna inicial (`CLFI`) e final (`CLFF`).
* Itera ao longo das colunas do intervalo com passo 4 (`COLU = COLU + 4`), acumulando a soma do `TCH` de todos os dias selecionados.
* Gera o ranking acumulado do período na aba `FACE` com a legenda: `"Periodo: {DTI} à {DTF}"`.


* **`Sub atualiza_entrante()`:**
* Totaliza as tentativas de chamadas por localidade na aba `ENTRANTE`, gerando o ranking **TOP 10** em `FACE!R35:S44`.


* **`Sub Ult_Atua()` (Auditoria de Operação):**
* Recupera o login de rede e o hostname da estação de trabalho via APIs do Windows.
* Grava em `FACE!A4`:
> *"{USU_Name} usando a máquina {Comp_Name}"*





---

### 3.4. `Módulo4` — Atalhos de Filtro em Tabela

* **`Sub Filtro_Sainte()` / `Sub Filtro_Entrante()`:** Procedimentos utilitários que capturam o valor da célula atualmente selecionada (`Selection`) e ativam imediatamente o `AutoFilter` na coluna 1 da respectiva aba.

---

### 3.5. Componentes de Interface (`UserForms`)

| Formulário | Descrição e Componentes |
| --- | --- |
| **`UserForm1`** | Splash screen de abertura do arquivo (exibido por 3 segundos). |
| **`UserForm2`** | Modal de parametrização do tamanho do ranking com botões de opção (`OptionButton1` a `6` para TOP 20, 30, 40, 50, 60 ou 70). |
| **`UserForm3`** | Tela de bloqueio e aviso visual durante o processamento da macro ("Aguarde..."). |
| **`UserForm4`** | Splash screen temporizado de conclusão do processo de importação. |
| **`UserForm5`** | Diálogo de parametrização de intervalo de datas (`TextBox1:3` para DTI e `TextBox4:6` para DTF), com validação de consistência cronológica via `DateDiff`. |

---

## 4. Dicionário de Termos Técnicos e Negócio

* **TCH (*Total Calls Handled*):** Volume total de tentativas de chamadas telefônicas.
* **TTC (*Total Time of Conversation*):** Minutos totais tarifados/conversados.
* **OK% (*Answer-Seizure Ratio - ASR*):** Taxa de completamento percentual de chamadas atendidas sobre tentadas.
* **LCR (*Least Cost Routing Share*):** Percentual do volume total do país que foi roteado pela operadora parceira indicada.
* **% Int.:** Participação percentual do tráfego do país específico sobre a volumetria total do tráfego internacional brasileiro monitorado.

---

## 5. Avaliação Técnica e Recomendações de Engenharia

1. **Compatibilidade com Excel 64 bits:**
* As declarações `GetComputerNameA` e `GetUserNameA` no `Módulo3` utilizam declarações legadas de 32 bits. Para rodar sem erros no Office 64 bits, devem ser adaptadas com a diretiva `#If VBA7` e `PtrSafe`, ou substituídas nativamente por `Environ("USERNAME")` e `Environ("COMPUTERNAME")`.


2. **Eliminação do Limite de 256 Colunas:**
* A aplicação foi projetada sob a limitação do formato Excel 97-2003 (.xls com 256 colunas máximas), exigindo a deleção das colunas com mais de 60 dias (`Columns(4:7).Delete`). Ao converter a pasta para o formato moderno `.xlsm` (que suporta até 16.384 colunas), essa restrição torna-se desnecessária, permitindo histórico anual ou plurianual.


3. **Otimização Drástica de Performance:**
* O código itera sobre células da planilha com loops aninhados `Do Until` e manipulação visual de seleção (`.Select`, `.Copy`, `.PasteSpecial`).
* *Refatoração recomendada:* O carregamento das colunas para matrizes em memória (`Variant Arrays`) ou o uso do objeto `Scripting.Dictionary` para agregação de chaves acelera as rotinas `atualiza_sainte` e `atualiza_periodo` de minutos para milissegundos.

## **Macro Oficio Consolidação LDN Evol**
<img width="953" height="521" alt="image" src="https://github.com/user-attachments/assets/0d21155c-7393-4216-aeb5-a1da8402d04b" />

## Documentação Técnica e Funcional: Aplicação VBA Excel

---

## 1. Visão Geral da Solução

* **Arquivo:** `Macro_Oficio_Consolidação_LDN_Evol.xls`
* **Domínio de Negócio:** Telecomunicações / Regulatório (**Anatel**). Auditoria, consolidação e expurgo de indicadores de qualidade de chamadas de Longa Distância Nacional (**LDN**) associados ao cumprimento do **Ofício 745** e metas regulatórias de PMM (*Período de Maior Movimento*).
* **Finalidade Técnica:** A aplicação automatiza a extração e carga de bilhetagem/CDR (*Call Detail Records*) oriundos do sistema SGD (*Sistema de Gestão de Desempenho*), correlaciona registros por chaves concatenadas (`DDD + CSP + Indicador + PMM`), simula cenários iterativos de expurgo de chamadas não completadas (expurgo percentual de **NR - Não Responde** e **LO - Linha Ocupada**) e gera planilhas de consolidação prontas para envio aos órgãos reguladores e operadoras parceiras (Telemar, Brasil Telecom, Sercomtel, Telefônica).

---

## 2. Modelagem e Estrutura de Dados (Abas)

```
        [ Arquivo Externo: SGD / CDR (*.xls) ]
                         │
                         ▼
        ┌──────────────────────────────────┐
        │               Base               │  <- Repositório bruto de CDR (transposto com chave composta)
        └──────────────────────────────────┘
                         │
                         ▼
        ┌──────────────────────────────────┐
        │             Calculo              │  <- Matriz de agregação, simulação e expurgo (PMM1, 2 e 3)
        └──────────────────────────────────┘
                         │
                         ▼
        ┌──────────────────────────────────┐
        │               SGD                │  <- Relatório final formatado e exportado via SaveAs
        └──────────────────────────────────┘

```

### Detalhamento das Abas:

| Aba | Visibilidade Padrão | Função Técnica e Esquema de Dados |
| --- | --- | --- |
| **`Calculo`** | **Visível** | **Coração do Motor de Regras.** Contém: <br>

<br>• Matrizes base de rotas por operadora e região (colunas `BT` a `CN`).<br>

<br>• Linhas 63 a 68: Parâmetros de formulação dinâmica por PMM e níveis de totalização (`[TOTAL]`, `[TOTAL] nivel2`, `[TOTAL] nivel3`).<br>

<br>• Células `B64:F65`: Percentuais de expurgo aplicados a **NR** e **LO** para PMM1, PMM2 e PMM3.<br>

<br>• Linha 80 em diante: Grid analítico comparando chamadas tentadas, completadas e causas de não completamento. |
| **`Base`** | **Oculta** (`xlSheetHidden`) | **Staging de Bilhetagem.** Recebe a massa bruta de CDR importada do SGD (até ~50.000 registros). Cria e armazena na Coluna A a chave sintética de cruzamento: `DDD & CSP & IND & "'" & PMM`. |
| **`SGD`** | **Oculta** (`xlSheetHidden`) | **Template de Exportação Executiva.** Contém o layout oficial de reporte com cabeçalho de prestadora (ex.: `31 - Telemar`), operadora de origem (`BRASIL TELECOM`, `SERCOMTEL`, `TELEFONICA`) e colunas de índices percentuais: `[PMM, Área DD, LDN, OK'%, PRD'%, PAB'%, NR'%, LO'%, CO'%, OU'%]`. |

---

## 3. Arquitetura e Engenharia de Módulos VBA

### 3.1. `Módulo1` — Fluxo Principal de Processamento (`Sub auto_open`)

O procedimento `auto_open` orquestra a interface do operador e o fluxo de extração e transformação:

#### 1. Roteamento de Negócio e Seleção de Escopo:

* Exibe `UserForm1` (splash inicial) na primeira execução.
* Aciona `UserForm3` para selecionar o tipo de tratamento:
* **Opção 1:** Consolidação Regional (`DDD-X FCN7 LDN`) $\rightarrow$ Direciona para `inicio3`.
* **Opção 2:** Consolidação de Operadora sob Ofício 745 (`FCN7 LDN`) $\rightarrow$ Direciona para `inicio2`.
* **Opção 3:** Sair.


* **Seletor de Operadora (`UserForm4`):**
* `1`: Brasil Telecom (carrega template das colunas `BT` / índice 72).
* `2`: Sercomtel (colunas `CF` / índice 84).
* `3`: Telefônica (colunas `CB` / índice 80).
* `4`: CTBC (emite alerta de indisponibilidade momentânea do layout).


* **Seletor de Região (`UserForm5`):**
* `1`: Região II (colunas `CK` / índice 89).
* `2`: Região III (colunas `CN` / índice 92).



#### 2. Carga Dinâmica de Fórmulas de Totalização:

O código copia o esqueleto selecionado para `Calculo!A80` e varre as linhas aplicando dinamicamente blocos de fórmulas matriciais pré-configuradas:

* Registros `PMM1`: Copia fórmulas da linha `Calculo!63` (I63:...).
* Registros `PMM2`: Copia fórmulas da linha `Calculo!64`.
* Registros `PMM3`: Copia fórmulas da linha `Calculo!65`.
* Subtotais de Nível 1, 2 e 3 (`[TOTAL]`): Copia fórmulas das linhas 66, 67 e 68 respectivamente.

#### 3. Ingestão e Enriquecimento do CDR Bruto:

1. Abre diálogo `Application.GetOpenFilename` solicitando o arquivo diário gerado pelo SGD.
2. Abre o arquivo externo e insere uma nova coluna `A` no início.
3. Localiza a linha inicial do grid através da cor cinza (`Interior.ColorIndex = 15`).
4. **Criação da Chave Primária Sintética:** Itera até o fim da massa de dados gerando a chave:

$$\text{Chave} = \text{DDD (pos 1-2)} + \text{CSP} + \text{Indicador} + \text{"'"} + \text{PMM (pos 12-13)}$$


5. Transpõe os dados tratados para a aba `Base` da pasta principal e encerra a planilha externa sem salvar alterações.

#### 4. Motor Iterativo de Simulação e Expurgo (Loop `Do Until REP = 7`):

* Abre o modal `UserForm2` para entrada dos percentuais de corte de **NR** e **LO** por PMM.
* O motor recalcula as fórmulas na planilha `Calculo`, recalculando os indicadores `OK'%` (Completamento) e `CO'%` (Chamadas Ocupadas).
* **Exportação do Relatório Consolidado:**
1. Torna a aba `SGD` visível.
2. Ajusta o cabeçalho de prestadora de origem.
3. Replica as linhas consolidadas de `Calculo` para `SGD`.
4. Converte fórmulas em valores puros (`PasteSpecial xlValues`).
5. Copia a aba `SGD` para uma nova pasta de trabalho desvinculada.
6. Dispara `Application.GetSaveAsFilename` para que o operador salve o arquivo oficial consolidado.
7. Fecha a pasta exportada e repete o ciclo para ajustes finos se necessário.



#### 5. Teardown e Limpeza de Segurança:

* Ao término, limpa todos os dados sensíveis da aba `Base` e `SGD` (`Selection.Clear`).
* Oculta novamente as abas `Base` e `SGD` (`Visible = False`).
* Oculta as colunas de template `BS:CF` e as linhas intermediárias de processamento (`Rows("61:3848").Hidden = True`), mantendo a integridade visual da planilha.

---

### 3.2. Módulos de Interface (`UserForms`)

| Formulário | Função Técnica e Lógica de Controle |
| --- | --- |
| **`UserForm1`** | Splash screen institucional exibida apenas na primeira abertura da sessão (`Static pula`). |
| **`UserForm2`** | **Painel de Simulação de Expurgo:** Permite input manual de `NR1`, `NR2`, `NR3` e `LO1`, `LO2`, `LO3`. Possui rotina (`CommandButton2_Click`) que efetua contagem prévia de rotas e indicadores fora da meta regulatória (`Cells(OK, 24) < "70"`). Grava os parâmetros em `Calculo!B64:F65`. |
| **`UserForm3`** | Menu de seleção do escopo da consolidação (`1`: Região / `2`: Ofício 745 Operadora / `3`: Sair). Grava a escolha na célula de apoio `Calculo!A72`. |
| **`UserForm4`** | Menu de seleção de operadoras sob o Ofício 745 (`1`: Brasil Telecom, `2`: Sercomtel, `3`: Telefônica, `4`: CTBC, `5`: Cancelar). Grava o resultado em `Calculo!A72`. |
| **`UserForm5`** | Menu de seleção de região regulatória (`1`: Região II, `2`: Região III, `3`: Cancelar). Grava o resultado em `Calculo!A72`. |

---

## 4. Dicionário de Termos e Regras de Telecom / Regulatório

* **LDN (*Longa Distância Nacional*):** Tráfego telefônico interurbano inter-regional ou interestadual.
* **PMM (*Período de Maior Movimento*):** Janelas horárias de pico de tráfego regulamentadas pela Anatel (geralmente PMM1, PMM2 e PMM3 ao longo do dia comercial) onde o índice de completamento não pode sofrer degradação.
* **Ofício 745:** Ato regulatório/diretriz técnica estabelecendo critérios de expurgo de chamadas não atribuíveis à rede da prestadora de transporte (ex.: falha de terminal ou abandono do usuário).
* **OK% / OK'%:** Taxa de chamadas atendidas e completadas antes e após os expurgos regulatórios.
* **NR (*Não Responde*):** Chamadas completadas até o destino que tocaram até o tempo limite (*timeout*) sem que o usuário atendesse.
* **LO (*Linha Ocupada*):** Tentativas bloqueadas devido ao terminal de destino estar ocupado.
* **CO (*Causa Operacional / Congestionamento*):** Bloqueios gerados por saturação de enlace ou indisponibilidade de rota.
* **CSP (*Código de Seleção de Prestadora*):** Dígito de seleção da operadora de longa distância (ex.: 31, 14, 15, 21).

---

## 5. Diagnóstico Técnico e Oportunidades de Refatoração

1. **Eliminação de `.Select`, `.Activate` e Área de Transferência:**
* O código usa comandos pesados como `Cells.Select`, `Selection.Copy` e `ActiveSheet.Paste` em tabelas com mais de 45.000 linhas na aba `Base`, o que pode travar o Excel ou causar *Out of Memory*.
* *Refatoração sugerida:* Substituir a cópia gráfica pela manipulação direta via array ou leitura direta de intervalos:
```vba
Workbooks(MACRO).Sheets("Base").Range("A1:X" & LFM2).Value = _
    Workbooks(PLANCDR).Sheets(1).Range("A1:X" & LFM2).Value

```




2. **Correção Lógica nas Comparações do `UserForm2`:**
* No `CommandButton2_Click`, a condição de avaliação compara tipos incompatíveis como strings alfanuméricas: `Cells(OK, 24) < "70"`. Valores numéricos formatados como percentual devem ser comparados como `Val(Cells(OK, 24).Value) < 0.70`.


3. **Desacoplamento de Células de Apoio (`Calculo!A72`):**
* Os formulários `UserForm3`, `UserForm4` e `UserForm5` utilizam a célula física `Cells(72, 1)` da planilha ativa como memória global temporária. Se a planilha ativa não for `Calculo`, o valor sobrescreve dados de outras abas. Recomenda-se declarar uma variável pública tipada no módulo (`Public TipoConsolidacao As Integer`).
