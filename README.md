# SAP2000 Pórtico MVP para Excel

Este projeto gera uma planilha Excel habilitada para macros (`.xlsm`) que cria automaticamente um pórtico simples no SAP2000 a partir de dados digitados no Excel.

A versão atual **não exporta dados de um modelo SAP2000 existente**. O objetivo é criar o modelo do zero via OAPI/COM usando VBA com *late binding*.

## Arquivos do projeto

| Arquivo | Finalidade |
| --- | --- |
| `Sap2000ExcelMvp.bas` | Módulo VBA completo com as macros da planilha e a automação do SAP2000. |
| `CriarPlanilhaSAP2000.vbs` | Script Windows/VBS que abre o Excel, cria a planilha, importa o módulo VBA, cria o layout/botões e salva `SAP2000_Portico_MVP.xlsm`. |
| `README.md` | Este passo a passo. |


## Como baixar/gerar sem arquivos binários no PR

Este repositório mantém apenas arquivos texto versionáveis (`.bas`, `.vbs`, `.md`). Arquivos Excel gerados (`.xlsm`, `.xlsx`) e pacotes `.zip` não ficam no Git para evitar problemas com PRs e revisão de binários.

Para obter a planilha funcional, baixe ou clone o repositório no Windows e execute `CriarPlanilhaSAP2000.vbs`. O script abrirá o Excel, importará `Sap2000ExcelMvp.bas`, criará a aba `Entrada`, criará os botões reais e salvará `SAP2000_Portico_MVP.xlsm` localmente.

## Pré-requisitos

O teste real precisa ser feito no **Windows** com:

1. Microsoft Excel instalado.
2. SAP2000 instalado e registrado corretamente no Windows.
3. Macros habilitadas no Excel.
4. A opção do Excel **Trust access to the VBA project object model** habilitada para o script importar automaticamente o `.bas`:
   - `File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model`.

> Este repositório não possui comandos de build ou instalação. O arquivo `.xlsm` é gerado localmente pelo script VBS no Windows e fica ignorado pelo Git.

## Como gerar a planilha `.xlsm`

1. Copie estes arquivos para uma pasta no Windows:
   - `Sap2000ExcelMvp.bas`
   - `CriarPlanilhaSAP2000.vbs`
2. Dê duplo clique em `CriarPlanilhaSAP2000.vbs` ou execute pelo Prompt de Comando:

   ```cmd
   cscript CriarPlanilhaSAP2000.vbs
   ```

3. O script abrirá o Excel, criará uma nova pasta de trabalho, importará o módulo VBA, executará `CriarLayout` e salvará:

   ```text
   SAP2000_Portico_MVP.xlsm
   ```

4. Abra `SAP2000_Portico_MVP.xlsm` e habilite macros, se o Excel solicitar.

## Aba `Entrada`

A macro `CriarLayout` cria uma aba chamada `Entrada` com estes campos:

### Dados gerais

- Vão do pórtico `[cm]`
- Pé-direito `[cm]`

### Dados da viga

- Largura da viga `[cm]`
- Altura da viga `[cm]`
- Fck da viga `[MPa]`

### Dados dos pilares

- Dimensão X do pilar `[cm]`
- Dimensão Y do pilar `[cm]`
- Fck dos pilares `[MPa]`

### Carregamentos

- Carga distribuída vertical na viga `[tf/m]`
- Carga horizontal de vento `[tf]`

## Botões da planilha

A aba `Entrada` contém três botões:

| Botão | Macro | O que faz |
| --- | --- | --- |
| `Criar Layout` | `CriarLayout` | Recria os campos, formatação, botões e cabeçalho do log. |
| `Exportar para SAP2000` | `ExportarParaSAP2000` | Valida os dados e cria o pórtico do zero no SAP2000. |
| `Limpar Log` | `LimparLog` | Limpa a aba `Log` e recria o cabeçalho. |

## O que a macro `ExportarParaSAP2000` faz

Ao clicar em `Exportar para SAP2000`, o VBA:

1. Valida todos os campos da aba `Entrada` como números maiores que zero.
2. Converte dimensões em centímetros para metros.
3. Conecta a uma instância aberta do SAP2000 ou inicia uma nova via OAPI/COM.
4. Cria um modelo novo em branco.
5. Define unidades em `tf, m, C`.
6. Cria materiais de concreto para viga e pilares.
7. Cria seção retangular dos pilares.
8. Cria seção retangular da viga.
9. Lança a geometria:
   - Pilar esquerdo: `(0,0,0)` até `(0,0,H)`.
   - Pilar direito: `(L,0,0)` até `(L,0,H)`.
   - Viga superior: `(0,0,H)` até `(L,0,H)`.
10. Aplica engaste na base dos pilares.
11. Cria os casos/padrões de carga:
    - `PP`
    - `SOBRECARGA`
    - `VENTO`
12. Aplica carga distribuída vertical na viga no caso `SOBRECARGA`.
13. Aplica carga horizontal de vento no nó superior direito no caso `VENTO`.
14. Atualiza a visualização do SAP2000.
15. Registra todas as etapas na aba `Log`.

## Observações técnicas

- O VBA usa `CreateObject("CSI.SAP2000.API.SapObject")` e `GetObject` para evitar referência manual à biblioteca do SAP2000.
- As unidades do SAP2000 são definidas com o valor de enumeração `Ton_m_C = 12`, que corresponde ao uso prático de tonelada-força, metro e Celsius.
- A carga vertical da viga é aplicada no eixo global `Z` com sinal negativo, ou seja, para baixo.
- A carga de vento é aplicada no nó superior direito no sentido global `+X`.
- O peso próprio é representado pelo caso/padrão `PP` com multiplicador de peso próprio igual a `1`.

## Solução de problemas

### O script não importa o módulo VBA

Habilite no Excel a opção **Trust access to the VBA project object model** e execute o script novamente.

### O SAP2000 não abre ou não conecta

Verifique se o SAP2000 está instalado e se a API COM está registrada no Windows. Em algumas instalações pode ser necessário abrir o SAP2000 uma vez como administrador para completar registros do COM.

### A macro para com código de retorno do SAP2000

Consulte a aba `Log`. Cada chamada crítica à OAPI registra a etapa que falhou e o código retornado pelo SAP2000.

### Não consigo testar neste ambiente Linux/CI

Isso é esperado. A validação real depende do Excel e do SAP2000 no Windows. Neste repositório só é possível validar estaticamente a presença dos arquivos e das macros.
