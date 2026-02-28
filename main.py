import pandas as pd
import glob
import os
import json

PASTA_DADOS = "planilhas"
PASTA_SAIDA = "saida"

def carregar_arquivos():
    arquivos = glob.glob(f"{PASTA_DADOS}/*.xlsx")
    dfs = []
    for arquivo in arquivos:
        # Ler com header=1 pois a linha 1 contém os nomes reais das colunas
        df = pd.read_excel(arquivo, header=1)
        df.columns = df.columns.str.strip().str.lower()

        # Converter valor_repassado de string com vírgula para numérico
        if 'valor_repassado' in df.columns:
            df['valor_repassado'] = pd.to_numeric(
                df['valor_repassado'].astype(str).str.replace(',', '.'),
                errors='coerce'
            ).fillna(0)

        dfs.append(df)
    return pd.concat(dfs, ignore_index=True)

def crescimento_percentual(atual, anterior):
    if anterior == 0:
        return None
    return ((atual - anterior) / anterior) * 100

def calcular_indicadores(df):

    total_socios = df['ativos'].sum()

    # ACI pode ser calculado como valor_repassado ou usar 0 se não existir
    if 'valor_repassado' in df.columns:
        total_contribuicao = df['valor_repassado'].sum()
    else:
        total_contribuicao = 0

    if total_socios == 0:
        return {}

    jovens = df['menor19'].sum() + df['de19a23'].sum()

    deficiencias = df[['surdos','auditiva','cegos','baixa_visao',
                       'fisica_inferior','fisica_superior',
                       'neurologico','intelectual']].sum().sum()

    # Calcular total de UMPs (estrutura) como soma de organizadas e não organizadas
    total_umps = 0
    if 'ump_organizada' in df.columns:
        total_umps += df['ump_organizada'].sum()
    if 'ump_nao_organizada' in df.columns:
        total_umps += df['ump_nao_organizada'].sum()

    indicadores = {
        "total_socios": int(total_socios),
        "total_contribuicao": float(total_contribuicao),
        "media_contribuicao_por_socio": float(total_contribuicao / total_socios) if total_socios > 0 else 0.0,
        "indice_renovacao_geracional": float(jovens / total_socios * 100),
        "percentual_inclusao": float(deficiencias / total_socios * 100),
        "total_umps": int(total_umps)
    }

    return indicadores

def gerar_alertas(indicadores, crescimento_socios=None, crescimento_contrib=None):

    alertas = []

    if crescimento_socios and crescimento_socios < -10:
        alertas.append("Queda superior a 10% no número de sócios")

    if crescimento_contrib and crescimento_contrib < -15:
        alertas.append("Queda superior a 15% na contribuição")

    if indicadores.get("indice_renovacao_geracional", 100) < 25:
        alertas.append("Baixa renovação geracional")

    if indicadores.get("percentual_inclusao", 100) < 2:
        alertas.append("Baixo índice de inclusão")

    return alertas

def gerar_estatisticas():

    df = carregar_arquivos()
    # Converter anos para inteiro para garantir ordenação correta
    df['ano_referencia'] = pd.to_numeric(df['ano_referencia'], errors='coerce')
    df = df.dropna(subset=['ano_referencia'])  # Remover linhas com anos inválidos
    anos = sorted(df['ano_referencia'].unique().astype(int))

    resultado_final = {}

    for i, ano in enumerate(anos):

        df_ano = df[df['ano_referencia'] == ano]

        nacional = calcular_indicadores(df_ano)

        resultado_final[str(ano)] = {
            "nacional": nacional,
            "regioes": {},
            "crescimento_nacional": {},
            "alertas_nacionais": []
        }

        # Crescimento nacional
        if i > 0:
            ano_anterior = anos[i-1]
            anterior = resultado_final[str(ano_anterior)]["nacional"]

            crescimento_socios = crescimento_percentual(
                nacional["total_socios"],
                anterior["total_socios"]
            )

            crescimento_contrib = crescimento_percentual(
                nacional["total_contribuicao"],
                anterior["total_contribuicao"]
            )

            resultado_final[str(ano)]["crescimento_nacional"] = {
                "socios_percentual": crescimento_socios,
                "contribuicao_percentual": crescimento_contrib
            }

            resultado_final[str(ano)]["alertas_nacionais"] = gerar_alertas(
                nacional,
                crescimento_socios,
                crescimento_contrib
            )

        # REGIÕES
        for regiao in df_ano['regiao'].unique():

            df_reg = df_ano[df_ano['regiao'] == regiao]
            indicadores_reg = calcular_indicadores(df_reg)

            resultado_final[str(ano)]["regioes"][regiao] = {
                "indicadores": indicadores_reg,
                "contribuicao_per_capita": indicadores_reg["media_contribuicao_por_socio"],
                "alertas": gerar_alertas(indicadores_reg)
            }

    os.makedirs(PASTA_SAIDA, exist_ok=True)

    with open(f"{PASTA_SAIDA}/base_ia_estatistica.json", "w", encoding="utf-8") as f:
        json.dump(resultado_final, f, indent=4, ensure_ascii=False)

if __name__ == "__main__":
    gerar_estatisticas()