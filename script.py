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
        return 0
    return ((atual - anterior) / anterior) * 100

def classificar_crescimento(valor):
    if valor > 8:
        return "Crescimento Forte"
    elif valor > 2:
        return "Crescimento Moderado"
    elif valor >= -2:
        return "Estável"
    else:
        return "Declínio"

def calcular_indicadores(df):

    total_socios = df['ativos'].sum()

    # ACI pode ser calculado como valor_repassado ou usar 0 se não existir
    if 'valor_repassado' in df.columns:
        total_contribuicao = df['valor_repassado'].sum()
    else:
        total_contribuicao = 0

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

    if total_socios == 0:
        return {}

    return {
        "total_socios": int(total_socios),
        "total_contribuicao": float(total_contribuicao),
        "media_contribuicao_por_socio": float(total_contribuicao / total_socios) if total_socios > 0 else 0.0,
        "indice_renovacao_geracional": float(jovens / total_socios * 100),
        "percentual_inclusao": float(deficiencias / total_socios * 100),
        "total_umps": int(total_umps),
        "indice_maturidade_estrutural": float(total_socios / total_umps) if total_umps > 0 else 0
    }

def calcular_score(indicadores, crescimento):

    score = 0

    # Crescimento (30)
    score += min(max(crescimento + 10, 0), 20) * 1.5

    # Renovação (25)
    score += min(indicadores["indice_renovacao_geracional"], 50) * 0.5

    # Inclusão (20)
    score += min(indicadores["percentual_inclusao"], 10) * 2

    # Contribuição (25)
    score += min(indicadores["media_contribuicao_por_socio"] / 100, 25)

    return round(min(score, 100), 2)

def gerar_prompt_ia(base):

    return f"""
Com base nos dados estatísticos institucionais abaixo, gere um relatório analítico formal, técnico e estratégico.

Destaque:
- Tendências nacionais
- Crescimento ou declínio
- Regiões em destaque
- Alertas estruturais
- Recomendações estratégicas

Dados:
{json.dumps(base, indent=2, ensure_ascii=False)}
"""

def gerar_estatisticas():

    df = carregar_arquivos()
    # Converter anos para inteiro para garantir ordenação correta
    df['ano_referencia'] = pd.to_numeric(df['ano_referencia'], errors='coerce')
    df = df.dropna(subset=['ano_referencia'])  # Remover linhas com anos inválidos
    df['ano_referencia'] = df['ano_referencia'].astype(int)  # Converter para int
    anos = sorted(df['ano_referencia'].unique())

    resultado_final = {}

    for i, ano in enumerate(anos):

        df_ano = df[df['ano_referencia'] == ano]
        nacional = calcular_indicadores(df_ano)

        resultado_final[str(ano)] = {
            "nacional": nacional,
            "regioes": {},
            "ranking_top10_sinodais": [],
            "crescimento_nacional": {},
            "classificacao_nacional": ""
        }

        # Crescimento Nacional
        if i > 0:
            anterior = resultado_final[str(anos[i-1])]["nacional"]

            crescimento = crescimento_percentual(
                nacional["total_socios"],
                anterior["total_socios"]
            )

            resultado_final[str(ano)]["crescimento_nacional"] = crescimento
            resultado_final[str(ano)]["classificacao_nacional"] = classificar_crescimento(crescimento)

        # Ranking Sinodais
        ranking = df_ano.groupby('nome')['ativos'].sum().sort_values(ascending=False).head(10)
        resultado_final[str(ano)]["ranking_top10_sinodais"] = ranking.to_dict()

        # Regiões
        for regiao in df_ano['regiao'].unique():

            df_reg = df_ano[df_ano['regiao'] == regiao]
            indicadores = calcular_indicadores(df_reg)

            crescimento_reg = 0
            if i > 0:
                df_reg_ant = df[df['ano_referencia'] == anos[i-1]]
                df_reg_ant = df_reg_ant[df_reg_ant['regiao'] == regiao]
                if not df_reg_ant.empty:
                    ant = calcular_indicadores(df_reg_ant)
                    crescimento_reg = crescimento_percentual(
                        indicadores["total_socios"],
                        ant["total_socios"]
                    )

            score = calcular_score(indicadores, crescimento_reg)

            resultado_final[str(ano)]["regioes"][regiao] = {
                "indicadores": indicadores,
                "crescimento_percentual": crescimento_reg,
                "classificacao": classificar_crescimento(crescimento_reg),
                "score_institucional": score
            }

    os.makedirs(PASTA_SAIDA, exist_ok=True)

    caminho_base = f"{PASTA_SAIDA}/base_ia_estatistica.json"
    with open(caminho_base, "w", encoding="utf-8") as f:
        json.dump(resultado_final, f, indent=4, ensure_ascii=False)

    prompt = gerar_prompt_ia(resultado_final)
    with open(f"{PASTA_SAIDA}/prompt_relatorio_ia.txt", "w", encoding="utf-8") as f:
        f.write(prompt)

if __name__ == "__main__":
    gerar_estatisticas()