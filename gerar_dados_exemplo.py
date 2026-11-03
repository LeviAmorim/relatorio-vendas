#!/usr/bin/env python3
"""Gera uma planilha de vendas fictícia para testar o gerar_relatorio.py.

Uso:
    python gerar_dados_exemplo.py
    python gerar_dados_exemplo.py -o minha_planilha.xlsx -n 500
"""

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

# (produto, categoria, preço base)
CATALOGO = [
    ("Camiseta Básica", "Roupas", 49.90),
    ("Calça Jeans", "Roupas", 139.90),
    ("Jaqueta", "Roupas", 289.90),
    ("Meia (par)", "Roupas", 19.90),
    ("Tênis Casual", "Calçados", 249.90),
    ("Chinelo", "Calçados", 39.90),
    ("Mochila", "Acessórios", 119.90),
    ("Boné", "Acessórios", 59.90),
]
VENDEDORES = ["Ana", "Bruno", "Carla", "Diego"]


def gerar(n: int, semente: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(semente)

    datas = pd.to_datetime("2026-01-01") + pd.to_timedelta(rng.integers(0, 304, n), unit="D")
    indices = rng.choice(len(CATALOGO), n, p=[0.22, 0.14, 0.06, 0.18, 0.10, 0.14, 0.08, 0.08])
    promocao = rng.choice([1.0, 0.9], n, p=[0.8, 0.2])

    df = pd.DataFrame(
        {
            "Data": datas,
            "Produto": [CATALOGO[i][0] for i in indices],
            "Categoria": [CATALOGO[i][1] for i in indices],
            "Vendedor": rng.choice(VENDEDORES, n, p=[0.35, 0.25, 0.25, 0.15]),
            "Quantidade": rng.choice([1, 2, 3, 4, 5], n, p=[0.45, 0.25, 0.15, 0.10, 0.05]),
            "Preco_Unitario": [round(CATALOGO[i][2] * p, 2) for i, p in zip(indices, promocao)],
        }
    )
    return df.sort_values("Data").reset_index(drop=True)


def main():
    parser = argparse.ArgumentParser(description="Gera planilha de vendas fictícia.")
    parser.add_argument("-o", "--saida", type=Path, default=Path("exemplos/vendas_exemplo.xlsx"))
    parser.add_argument("-n", "--linhas", type=int, default=400, help="número de vendas (padrão: 400)")
    args = parser.parse_args()

    args.saida.parent.mkdir(parents=True, exist_ok=True)
    df = gerar(args.linhas)
    df.to_excel(args.saida, index=False)
    print(f"Planilha de exemplo criada: {args.saida} ({len(df)} vendas)")


if __name__ == "__main__":
    main()
