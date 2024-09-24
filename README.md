---

# Lista de Materiais - Cálculo de Quantidade de Importação

Este projeto tem como objetivo automatizar o cálculo da **quantidade de importação** de materiais a partir de uma lista gerada de um projeto. Utilizando duas planilhas (lista de materiais e banco de dados de materiais), o programa realiza a união dos dados e calcula a quantidade necessária para importação, levando em consideração diferentes unidades de compra, como **peso**, **metros**, **unidades** e **barra** (com cada barra representando 6 metros).

## Funcionalidades:
- **Carregamento automático de planilhas**: Leitura das planilhas fornecidas em formato Excel, utilizando `pandas`.
- **Cálculo dinâmico da quantidade de importação**:
  - Para itens comprados por **peso**, **metros** ou **unidade**, o cálculo da quantidade de importação corresponde à quantidade solicitada no projeto.
  - Para itens comprados por **barra**, o programa calcula a quantidade necessária com base em barras de 6 metros, arredondando para cima (caso a quantidade do projeto exija uma fração de barra).
- **Exportação do resultado**: Geração de um novo arquivo Excel contendo a coluna "QUANTIDADE DE IMPORTAÇÃO" com os valores calculados.
  
## Tecnologias Utilizadas:
- **Python**: Linguagem de programação utilizada para a lógica do cálculo e manipulação dos dados.
- **Pandas**: Biblioteca utilizada para a manipulação e processamento dos dados contidos nas planilhas Excel.
- **Openpyxl**: Biblioteca responsável por ler e salvar os arquivos Excel.
- **Math**: Utilizada para operações de arredondamento (como o `ceil` para itens comprados por barra).

## Como Utilizar:
1. Clone este repositório para o seu ambiente local:
   ```bash
   git clone https://github.com/usuario/repo.git
   ```

2. Instale as dependências necessárias:
   ```bash
   pip install pandas openpyxl
   ```

3. Coloque as planilhas de **lista de materiais** e **banco de dados de materiais** na mesma pasta do script Python.

4. Execute o script principal:
   ```bash
   python calcular_quantidade_importacao.py
   ```

5. O resultado final será gerado em um novo arquivo Excel com a coluna "QUANTIDADE DE IMPORTAÇÃO" calculada.

## Possíveis Melhorias:
- Suporte a diferentes comprimentos de barras.
- Otimização para processamento em massa, permitindo manipulação de grandes volumes de dados de forma mais eficiente.
- Interface gráfica para facilitar a interação com o usuário.

## Contribuições:
Sinta-se à vontade para abrir **issues** ou enviar **pull requests** para sugerir melhorias ou correções.

---

Essa descrição é clara, objetiva e informativa. Ela explica as funcionalidades do projeto, as tecnologias utilizadas e como utilizá-lo, além de sugerir melhorias futuras e incentivar contribuições da comunidade.
