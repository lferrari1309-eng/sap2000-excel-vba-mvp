# MVP Excel VBA → SAP2000 OAPI

Este repositório contém um MVP em VBA para validar o fluxo **Excel → VBA → SAP2000**. A planilha recebe dados básicos de um pórtico plano simples e a macro cria automaticamente o modelo no SAP2000 usando a OAPI.

## O que o MVP cria no SAP2000

A rotina cria um pórtico plano 2D com:

- Dois pilares verticais:
  - Pilar esquerdo: `(0, 0, 0)` até `(0, 0, H)`
  - Pilar direito: `(L, 0, 0)` até `(L, 0, H)`
- Uma viga superior: `(0, 0, H)` até `(L, 0, H)`
- Apoios engastados na base dos pilares
- Unidade do modelo: **tf, m, C**
- Material de concreto para viga e pilares
- Seções retangulares para viga e pilares
- Casos/padrões de carga:
  - `PP`
  - `SOBRECARGA`
  - `VENTO`
- Carga distribuída vertical na viga no padrão `SOBRECARGA`
- Carga horizontal no nó superior direito no padrão `VENTO`
- Registro de execução na aba `Log`

## Arquivos VBA

Importe estes módulos padrão (`.bas`) no Editor do VBA:

- `modPrincipal.bas`: macro principal `ExportarParaSAP2000`
- `modPlanilha.bas`: criação da aba `Entrada`, leitura dos campos e botões
- `modValidacao.bas`: validação de campos vazios, zerados ou negativos
- `modSAP2000.bas`: conexão com SAP2000 e chamadas OAPI
- `modModeloPortico.bas`: constantes, nomes e tipo `DadosPortico`
- `modLog.bas`: criação/limpeza da aba `Log` e registro de mensagens

## Como criar o arquivo Excel `.xlsm`

1. Abra o Excel no Windows.
2. Crie uma pasta de trabalho em branco.
3. Salve como **Pasta de Trabalho Habilitada para Macro do Excel (`.xlsm`)**.
4. Use um nome como `sap2000_portico_mvp.xlsm`.

## Como importar os módulos no Excel

1. Abra o arquivo `.xlsm`.
2. Pressione `ALT + F11` para abrir o Editor do VBA.
3. No painel do projeto, clique com o botão direito em `VBAProject (seu_arquivo.xlsm)`.
4. Escolha **Importar Arquivo...**.
5. Importe, um por um, todos os arquivos `.bas` deste repositório.
6. Salve o arquivo `.xlsm`.

## Como ativar macros

1. Feche e reabra o arquivo `.xlsm`.
2. Se aparecer a barra amarela de segurança, clique em **Habilitar Conteúdo**.
3. Se o Excel bloquear o arquivo por origem da internet:
   - Feche o Excel.
   - Clique com o botão direito no arquivo `.xlsm`.
   - Abra **Propriedades**.
   - Marque **Desbloquear**, se disponível.
   - Abra novamente o arquivo e habilite macros.

## Como conectar ao SAP2000

A rotina usa **late binding**, portanto normalmente **não é necessário** ativar referência do SAP2000 no VBA. O código tenta:

1. Conectar a uma instância aberta do SAP2000 com `GetObject`.
2. Se não encontrar, abrir o SAP2000 via `SAP2000v1.Helper` e `CSI.SAP2000.API.SapObject`.

Pré-requisitos:

- SAP2000 instalado no Windows.
- SAP2000 registrado corretamente como servidor COM/OAPI.
- Licença válida do SAP2000.
- Excel e SAP2000 com compatibilidade de arquitetura recomendada (ambos 64 bits, ou ambos 32 bits, quando aplicável).

### Referência opcional do SAP2000 no VBA

Como o MVP usa late binding, esta etapa é opcional. Se você quiser desenvolver com IntelliSense e tipos fortes:

1. No Editor do VBA, acesse **Tools/Ferramentas → References/Referências**.
2. Procure e marque a biblioteca equivalente à sua instalação, geralmente algo como:
   - `SAP2000v1 1.0 Type Library`
   - `CSI SAP2000 API`
3. O nome exato varia conforme a versão instalada.
4. Se ativar referência e trocar para early binding, será necessário adaptar declarações `As Object` para os tipos da biblioteca.

## Como criar os botões

Há duas opções.

### Opção recomendada: macro automática

1. No Excel, pressione `ALT + F8`.
2. Execute a macro `CriarOuAtualizarLayout`.
3. A macro criará a aba `Entrada`, a aba `Log` e três botões:
   - **Criar/Atualizar Layout**
   - **Exportar para SAP2000**
   - **Limpar Log**

### Opção manual

1. Ative a guia **Desenvolvedor** no Excel.
2. Clique em **Inserir → Botão (Controle de Formulário)**.
3. Crie três botões e associe as macros:
   - Botão `Criar/Atualizar Layout` → macro `CriarOuAtualizarLayout`
   - Botão `Exportar para SAP2000` → macro `ExportarParaSAP2000`
   - Botão `Limpar Log` → macro `LimparLog`

## Como usar a planilha

1. Execute `CriarOuAtualizarLayout` para montar a aba `Entrada`.
2. Preencha os campos:
   - Vão/comprimento do pórtico `[cm]`
   - Pé-direito `[cm]`
   - Largura da viga `[cm]`
   - Altura da viga `[cm]`
   - Fck da viga `[MPa]`
   - Dimensão X do pilar `[cm]`
   - Dimensão Y do pilar `[cm]`
   - Fck dos pilares `[MPa]`
   - Carga distribuída vertical na viga `[tf/m]`
   - Carga horizontal de vento `[tf]`
3. Clique em **Exportar para SAP2000**.
4. A rotina valida os campos, abre/conecta ao SAP2000 e cria o pórtico.
5. Consulte a aba `Log` se houver erro ou para acompanhar o passo a passo.

## Lógica resumida da rotina

1. `ExportarParaSAP2000` inicia a execução e registra no `Log`.
2. `LerDadosEntrada` lê a aba `Entrada` e converte dimensões de centímetros para metros.
3. `ValidarDadosEntrada` impede valores vazios, negativos ou iguais a zero.
4. `ConectarOuAbrirSAP2000` tenta conectar ao SAP2000 aberto ou iniciar uma nova instância.
5. `CriarModeloPorticoSAP2000` cria um modelo novo, define unidades, materiais, seções, barras, restrições e cargas.
6. A janela do SAP2000 é atualizada e a execução é finalizada com mensagem amigável.

## Conversões adotadas

- Dimensões: `cm × 0,01 = m`
- Cargas: informadas diretamente em `tf` e `tf/m`
- Resistência `fck`: informada em `MPa`
- Módulo de elasticidade didático: `Eci = 5600 × √fck` em MPa, convertido para `tf/m²`

## Partes que podem precisar de ajuste conforme a versão do SAP2000

Algumas versões do SAP2000 podem ter pequenas diferenças na OAPI. Se ocorrer erro, revise principalmente em `modSAP2000.bas`:

- `SAP_HELPER_PROG_ID = "SAP2000v1.Helper"`
- `SAP_PROG_ID = "CSI.SAP2000.API.SapObject"`
- Código da unidade `SAP_UNITS_TON_M_C = 12`
- Códigos dos tipos de carga (`DEAD`, `LIVE`, `WIND`)
- Ordem de argumentos de:
  - `FrameObj.AddByCoord`
  - `FrameObj.SetLoadDistributed`
  - `PointObj.SetLoadForce`
  - `PropFrame.SetRectangle`

Se sua versão exigir referência COM explícita, ative a referência da biblioteca do SAP2000 descrita acima e consulte o arquivo de ajuda da OAPI instalado com o SAP2000.

## Limitações do MVP

Este MVP é propositalmente simples. Ainda **não** inclui:

- Múltiplos pavimentos
- Múltiplos vãos
- Combinações de carga
- Dimensionamento de concreto armado
- Exportação de resultados
- Verificação normativa
- Tratamento avançado de eixos locais
- Modelagem de diafragmas, lajes, rótulas ou fundações
- Interface avançada no Excel
- Persistência de versões do modelo

O foco desta primeira versão é validar o conceito: **digitar dados no Excel e criar um pórtico simples no SAP2000**.
