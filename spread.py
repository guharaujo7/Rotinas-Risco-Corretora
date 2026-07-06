import pandas as pd


###################     DATAFRAME DO QUICK SEPARADO POR SACADO + TAXA (AJUSTAR)       ################
df_quicksight_path = r'C:\Users\ghamkjx\Downloads\Analítico_de_Notas_D_1783345993183.xlsx'
df_quicksight = pd.read_excel(df_quicksight_path, engine="openpyxl")

###################     LEITURA DA CURVA NA REDE       ################
df_curva_spot_path = r'\\bbaprod3\fo\diretoria de produtos ativos\ativos em reais\risco sacado\cotacoes\curva_diaria.xls'
df_curva_spot = pd.read_excel(df_curva_spot_path, engine="xlrd", usecols="D,E,G,J", skiprows=1, nrows=150)
df_curva_spot["Vencimento"] = pd.to_datetime(df_curva_spot["Vencimento"], dayfirst=True)
df_curva_spot = df_curva_spot.rename(columns={'Vencimento': 'Data Vencimento'})
df_curva_spot = df_curva_spot.rename(columns={'Dar': 'Funding'})

def calcular_spread(df):
    #separa as colunas necessárias Valor - Vencimento - Taxa (ajustar)
    df = df[['Valor', 'Data Vencimento', 'Taxa']]

    taxa_alvo = df["Taxa"].iloc[0]
    taxa_alvo = taxa_alvo.round(4)
    df = df.drop(columns=["Taxa"])

    spread = 0
    taxa_media = 0

    while taxa_alvo != taxa_media:

        if taxa_alvo > taxa_media:
            spread = spread + 0.001
        else:
            spread = spread - 0.001

        spread = round(spread, 3)
        taxa_media = calcular_taxa(df, spread)
    
    return spread

def calcular_taxa(df, spread):

    df["Data Vencimento"] = pd.to_datetime(df["Data Vencimento"], dayfirst=True)
    df = pd.merge(df, df_curva_spot, on='Data Vencimento', how='left')

    df["Funding"] = df["Funding"] + 1
    df["Taxa Efetiva"] = (df["Funding"] * (spread/100+1))**(df["DU"]/252)
    df["Taxa a.m linear"] = ((df["Taxa Efetiva"]-1)/df["Taxa Efetiva"])/df["DC"]*30*100
    df["Taxa a.m linear"] = df["Taxa a.m linear"].round(4)

    #calculo taxa média
    df["VF x Pz x Tx"] = df["Valor"] * df["DC"] * df["Taxa a.m linear"]
    df["VF x Pz"] = df["Valor"] * df["DC"]

    taxa_media = df["VF x Pz x Tx"].sum() / df["VF x Pz"].sum()
    taxa_media = taxa_media.round(4)

    return taxa_media

teste = calcular_spread(df_quicksight)
print(teste)
