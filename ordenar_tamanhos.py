import pandas as pd

# Definição das ordens corretas para tamanhos padrão
ordem_feminino = ["Tamanho: PP", "Tamanho: P", "Tamanho: M", "Tamanho: G", "Tamanho: GG"]
ordem_masculino = ["Tamanho: P", "Tamanho: M", "Tamanho: G", "Tamanho: GG", "Tamanho: GGG"]

# Função para ordenar os tamanhos conforme a regra
def ordenar_tamanhos(variations):
    if pd.isna(variations):  # Tratamento para células vazias
        return ""

    tamanhos = str(variations).split(";")  # Converte para string e separa os tamanhos
    
    # Verifica se os tamanhos são numéricos (ex: 36, 38, 40)
    if any(char.isdigit() for char in tamanhos[0]):  
        tamanhos_ordenados = sorted(tamanhos, key=lambda x: int(''.join(filter(str.isdigit, x))))
    else:
        # Ordenação para tamanhos padrão (feminino e masculino)
        if set(ordem_feminino).intersection(tamanhos):  # Verifica se há tamanhos femininos
            tamanhos_ordenados = sorted(tamanhos, key=lambda x: ordem_feminino.index(x) if x in ordem_feminino else len(ordem_feminino))
        elif set(ordem_masculino).intersection(tamanhos):  # Verifica se há tamanhos masculinos
            tamanhos_ordenados = sorted(tamanhos, key=lambda x: ordem_masculino.index(x) if x in ordem_masculino else len(ordem_masculino))
        else:
            tamanhos_ordenados = tamanhos  # Mantém a ordem original se não corresponder às regras

    return ";".join(tamanhos_ordenados)

# Caminho do arquivo de entrada
file_path = "tamanhos.xlsx"  # Substitua pelo nome correto do seu arquivo

# Carregar a planilha Excel
df = pd.read_excel(file_path)

# Verifica se a coluna "variations" existe na planilha
if "variations" in df.columns:
    df["variations"] = df["variations"].apply(ordenar_tamanhos)

    # Salvar as alterações em um novo arquivo Excel
    df.to_excel("tamanhos_ordenados.xlsx", index=False)

    print("Arquivo processado com sucesso: 'tamanhos_ordenados.xlsx'")
else:
    print("Erro: A coluna 'variations' não foi encontrada no arquivo.")
