# Análise de Unidades Contratadas por Ano

Este projeto contém um script Python que realiza a análise de dados de unidades contratadas por ano a partir de um arquivo Excel e gera uma visualização em forma de gráfico de barras, além de salvar os resultados em um arquivo CSV.

## Descrição do Código

O script realiza as seguintes etapas:

1. **Importação de Bibliotecas**:
   - `os`: Para manipulação de caminhos de arquivos.
   - `pandas`: Para leitura e manipulação de dados.
   - `matplotlib.pyplot`: Para criação de visualizações gráficas.

2. **Leitura do Arquivo Excel**:
   - O script lê o arquivo `base_ogu_202503_snh.xlsx` localizado na mesma pasta do script, utilizando a aba `base_ogu`.
   - A coluna `Unidades Contratadas` é convertida para valores numéricos, com valores inválidos sendo tratados como `NaN`.

3. **Processamento dos Dados**:
   - A coluna `Ano` é convertida para valores inteiros (tipo `Int64` para suportar valores nulos).
   - Os dados são agrupados por `Ano`, somando as `Unidades Contratadas` para cada ano.
   - O DataFrame resultante é renomeado com colunas `Rótulos de Linha` (para os anos) e `Soma de mcmv_ogu_22_qtd_uh` (para a soma das unidades contratadas).
   - Linhas com valores nulos nas colunas principais são removidas.
   - Os dados são ordenados por ano em ordem crescente.

4. **Visualização**:
   - Um gráfico de barras é gerado com os anos no eixo X e a soma das unidades contratadas no eixo Y.
   - O gráfico é configurado com título, rótulos nos eixos e dimensões definidas (10x6 polegadas).
   - O gráfico é salvo como `Tabela.png` na mesma pasta do script.

5. **Exportação**:
   - O DataFrame processado é salvo como `Tabela.csv` na mesma pasta do script, incluindo o índice.

6. **Exibição**:
   - O gráfico é exibido na tela (quando executado em um ambiente que suporta visualização gráfica).

## Requisitos

- Python 3.x
- Bibliotecas necessárias:
  - `pandas`
  - `matplotlib`
  - `openpyxl` (para leitura de arquivos Excel)
- Arquivo de entrada: `base_ogu_202503_snh.xlsx` (deve estar na mesma pasta do script)

## Como Executar

1. Certifique-se de que o arquivo `base_ogu_202503_snh.xlsx` está na mesma pasta do script.
2. Instale as dependências:
   ```bash
   pip install pandas matplotlib openpyxl
   ```
3. Execute o script:
   ```bash
   python script.py
   ```

## Saídas

- **Gráfico**: Um arquivo `Tabela.png` contendo o gráfico de barras.
- **Arquivo CSV**: Um arquivo `Tabela.csv` com os dados processados.

## Observações

- O script assume que o arquivo Excel possui uma aba chamada `base_ogu` e as colunas `Ano` e `Unidades Contratadas`.
- Caso o arquivo Excel não esteja presente ou o formato esteja incorreto, o script pode gerar erros.
- A exibição do gráfico (`plt.show()`) pode não funcionar em ambientes sem interface gráfica.