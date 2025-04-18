import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Carregar o arquivo Excel
caminho_arquivo = "dados_papelaria.xlsx"  # Se o arquivo estiver no mesmo diretório do script

# Importar as planilhas
df_vendas = pd.read_excel(caminho_arquivo, sheet_name="Vendas")
df_clientes = pd.read_excel(caminho_arquivo, sheet_name="Clientes")
df_estoque = pd.read_excel(caminho_arquivo, sheet_name="Estoque")
df_financeiro = pd.read_excel(caminho_arquivo, sheet_name="Financeiro")

# Exibir as primeiras linhas de cada planilha
print(" Dados de Vendas:")
print(df_vendas.head(), "\n")

print(" Dados de Clientes:")
print(df_clientes.head(), "\n")

print(" Dados de Estoque:")
print(df_estoque.head(), "\n")

print(" Dados Financeiros:")
print(df_financeiro.head(), "\n")

# Estatísticas gerais das vendas
print("Estatísticas gerais sobre vendas:\n")
print(df_vendas.describe())

# Calculando o faturamento total
faturamento_total = df_vendas["Valor_Total"].sum()
print(f" Faturamento total: R$ {faturamento_total:.2f}")

# Convertendo datas
if 'Data' in df_vendas.columns:
    df_vendas['Data'] = pd.to_datetime(df_vendas['Data'])

# Faturamento ao longo do tempo
plt.figure(figsize=(10, 5))
sns.lineplot(x=df_vendas['Data'], y=df_vendas['Valor_Total'], marker='o', color='b')
plt.title('Faturamento ao Longo do Tempo')
plt.xlabel('Data')
plt.ylabel('Faturamento (R$)')
plt.xticks(rotation=45)
plt.grid()
plt.show()
# Distribuição de vendas por categoria
if 'Categoria' in df_vendas.columns:
    plt.figure(figsize=(10, 5))
    sns.countplot(y=df_vendas['Categoria'], palette='pastel')
    plt.title('Distribuição de Vendas por Categoria')
    plt.xlabel('Número de Vendas')
    plt.ylabel('Categoria')
    plt.show()

# Estoque por categoria
if 'Categoria' in df_estoque.columns and 'Quantidade' in df_estoque.columns:
    plt.figure(figsize=(10, 5))
    sns.barplot(y=df_estoque['Categoria'], x=df_estoque['Quantidade'], palette='coolwarm')
    plt.title('Estoque por Categoria')
    plt.xlabel('Quantidade em Estoque')
    plt.ylabel('Categoria')
    plt.show()

# Receita por mês
if 'Data' in df_vendas.columns:
    df_vendas['Mes_Ano'] = df_vendas['Data'].dt.to_period('M')
    receita_mensal = df_vendas.groupby('Mes_Ano')['Valor_Total'].sum()
    plt.figure(figsize=(10, 5))
    receita_mensal.plot(kind='bar', color='purple')
    plt.title('Receita Mensal')
    plt.xlabel('Mês')
    plt.ylabel('Faturamento (R$)')
    plt.xticks(rotation=45)
    plt.show()

print("Análises e gráficos gerados com sucesso!")



