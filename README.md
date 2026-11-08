# Relatório de Vendas Automático

Script em Python que transforma uma planilha de vendas (`.xlsx` ou `.csv`) em um relatório pronto em Excel, com indicadores, tabelas resumidas e gráficos. O que normalmente leva horas de trabalho manual sai em poucos segundos.

## O que o relatório traz

| Aba | Conteúdo |
|---|---|
| **Resumo** | Faturamento total, nº de vendas, unidades, ticket médio, produto campeão, vendedor destaque, melhor e pior mês |
| **Por Mês** | Faturamento, variação em relação ao mês anterior e gráfico de linha |
| **Por Produto** | Faturamento, participação, unidades, ticket médio e gráfico de barras |
| **Por Categoria** | Mesmas métricas, agrupadas por categoria |
| **Por Vendedor** | Mesmas métricas, agrupadas por vendedor |
| **Dados** | Planilha original limpa, com a coluna Total calculada e filtros ligados |

## Instalação

```bash
git clone https://github.com/LeviAmorim/relatorio-vendas.git
cd relatorio-vendas
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Como usar

```bash
# Gerar o relatório a partir de uma planilha
python gerar_relatorio.py exemplos/vendas_exemplo.xlsx

# Escolher o nome do arquivo de saída
python gerar_relatorio.py minhas_vendas.csv -o relatorio_outubro.xlsx
```

Há um resultado pronto para conferir em `exemplos/relatorio_exemplo.xlsx`. Para testar sem dados reais, gere uma planilha fictícia:

```bash
python gerar_dados_exemplo.py              # cria exemplos/vendas_exemplo.xlsx
python gerar_dados_exemplo.py -n 1000 -o outra.xlsx
```

## Formato da planilha de entrada

A planilha precisa ter estas colunas (a ordem não importa, maiúsculas e acentos são ignorados):

| Coluna | Exemplo |
|---|---|
| `Data` | 15/03/2026 |
| `Produto` | Camiseta Básica |
| `Categoria` | Roupas |
| `Vendedor` | Ana |
| `Quantidade` | 2 |
| `Preco_Unitario` | 49,90 |

Também são aceitos os nomes `Qtd`, `Qtde`, `Preço`, `Valor Unitário`. Valores como `R$ 1.234,50` são convertidos automaticamente. Linhas com data, quantidade ou preço inválidos são ignoradas, e o script avisa quantas foram.

## Como funciona

1. **Leitura:** `pandas` lê o arquivo (o separador do CSV é detectado sozinho).
2. **Limpeza:** padroniza os nomes das colunas, converte tipos e descarta linhas inválidas.
3. **Cálculo:** agrupa por mês, produto, categoria e vendedor.
4. **Saída:** `openpyxl` monta o Excel formatado (moeda em R$, percentuais, filtros e gráficos).

## Tecnologias

Python 3 · pandas · openpyxl · numpy (apenas para os dados de exemplo)

## Possíveis evoluções

- Exportar também em PDF
- Comparar dois períodos (mês atual vs. anterior)
- Enviar o relatório por e-mail automaticamente
- Interface web simples com Flask para enviar a planilha pelo navegador

## Licença

MIT. Veja o arquivo [LICENSE](LICENSE).
