"""Dimensionamento de Equipe de CallCenter

# Dimensionamento de Equipe

Para realizar é preciso definir os paramentros

ITERVALO_MIN = É o tamanho do bloco de tempo usado no cálculo. EX : 30 calcula volume, TMA e HC a cada 30 minutos.

SLA_ALVO = Nivel de Serviço desejado. Ex: 0.80 é 80%

TEMPO_SLA = Tempo maximo de espera

SHRINKAGE = tempo que o tecnico se encotra indiponivel como pausas, saidas ao banheiro
"""

import pandas as pd
import os

INTERVALO_MIN = 30
SLA_ALVO = 0.95
TEMPO_SLA = 20
SHRINKAGE = 0.25

# pasta onde o Power BI salva
pasta = r"C:\\Users\\caminho\\Documentos\\Dimensionamento Mensal"

# pegar todos arquivos Excel da pasta
arquivos = [
    os.path.join(pasta, f)
    for f in os.listdir(pasta)
    if f.endswith(".xlsx") and "Geral_" in f
]


# pegar o mais recente
arquivo_mais_recente = max(arquivos, key=os.path.getmtime)

print("Arquivo carregado:", arquivo_mais_recente)

# ler arquivo
df = pd.read_excel(arquivo_mais_recente)

#df =  pd.read_excel("C:\\Users\\raquel.costa\\Downloads\\data.xlsx")

# limpeza pesada de nomes de colunas
df.columns = (
    df.columns
    .str.replace("﻿", "", regex=False)  # remove BOM
    .str.strip()                         # remove espaços
    .str.lower()
    .str.replace("[", "", regex=False)   # remove [
    .str.replace("]", "", regex=False)   # remove ]
)
print(df.columns)

import math

# -----------------------------
# Cálculo de tráfego (Erlangs)
# -----------------------------
def calcular_trafego(volume, tma_seg, intervalo_min):
    if volume <= 0 or tma_seg <= 0 or intervalo_min <= 0:
        return 0
    return (volume * tma_seg) / (intervalo_min * 60)


# ------------------------------------------------
# Erlang C estável (sem fatorial / sem potência)
# ------------------------------------------------
def erlang_c(trafego, agentes):
    # Proteções básicas
    if agentes <= trafego or trafego == 0:
        return 1.0

    # Erlang B por recorrência
    b = 1.0
    for k in range(1, agentes + 1):
        b = (trafego * b) / (k + trafego * b)

    # Conversão Erlang B → Erlang C
    return b / (1 - (trafego / agentes))


# ------------------------------------------------
# Cálculo de agentes com SLA
# ------------------------------------------------
def calcular_agentes(volume, tma, intervalo):
    trafego = calcular_trafego(volume, tma, intervalo)

    if trafego == 0:
        return 0

    agentes = max(1, math.ceil(trafego))

    # Limite de segurança (evita loop infinito)
    LIMITE_AGENTES = 300

    while agentes <= LIMITE_AGENTES:
        ec = erlang_c(trafego, agentes)

        sla = 1 - (
            ec * math.exp(
                -(agentes - trafego) * (TEMPO_SLA / tma)
            )
        )

        if sla >= SLA_ALVO:
            return agentes

        agentes += 1

    # Se estourar o limite, retorna o máximo calculado
    return agentes


# ------------------------------------------------
# Aplicar shrinkage
# ------------------------------------------------
def aplicar_shrinkage(agentes):
    if agentes <= 0:
        return 0
    return math.ceil(agentes / (1 - SHRINKAGE))

print(df.columns.tolist())

df = df[(df["volume"] > 0) & (df["tma"] > 0)]

resultados = []

for _, row in df.iterrows():
    agentes_base = calcular_agentes(
        row["volume"],
        row["tma"],
        INTERVALO_MIN
    )

    agentes_final = aplicar_shrinkage(agentes_base)

    resultados.append({
        "data": row["data"],
        "intervalo": row["intervalo"],
        "equipes": row["equipes"],
        "volume": row["volume"],
        "tma": row["tma"],
        "agentes_necessarios": agentes_final
    })

df_resultado = pd.DataFrame(resultados)

# Data
df_resultado["data"] = pd.to_datetime(
    df_resultado["data"],
    errors="coerce"
)

 #Intervalo - pega só a hora inicial (ex: "08:00 - 08:30")
df_resultado["intervalo"] = (
    df_resultado["intervalo"]
    .astype(str)
    .str.slice(0, 5)
)

'''df_resultado["intervalo"] = pd.to_datetime(
    df_resultado["intervalo"],
    format="%H:%M",
    errors="coerce"
)'''

df_resultado["intervalo"] = pd.to_datetime(
    df_resultado["intervalo"],
    errors="coerce"
)

df_resultado = df_resultado.dropna(subset=["intervalo"])

df_resultado = df_resultado.sort_values(
    by=["equipes", "intervalo"],
    ascending=[True, True]
)

df_resultado = df_resultado.sort_values(
    by=["equipes", "intervalo"],
    ascending=[True, True]
)

df_resultado["tma_x_volume"] = df_resultado["tma"] * df_resultado["volume"]

resumo = (
    df_resultado
    .groupby("equipes", as_index=False)
    .agg(
        Qtd_Tecnicos=("agentes_necessarios", "max"),
        tma_ponderado=("tma_x_volume", "sum"),
        volume_total=("volume", "sum")
    )
)

resumo["tma_ponderado"] = resumo["tma_ponderado"] / resumo["volume_total"]
resumo = resumo.drop(columns="volume_total")



Qtd_Tecnicos_equipe = (
    df_resultado
    .groupby("equipes", as_index=False)
    .agg(
        hc_pico=("agentes_necessarios", "max"),
        hc_medio=("agentes_necessarios", "mean"),
        hc_p90=("agentes_necessarios", lambda x: x.quantile(0.9))
    )
)

Qtd_Tecnicos_equipe = Qtd_Tecnicos_equipe.sort_values("equipes")

resumo = resumo.sort_values(
    by=["equipes"]
)

HORARIOS_ENTRADA = {
    "drive": ["07:30", "08:00", "09:00", "10:30", "12:00", "13:30"],
    "elevate": ["07:00", "08:00", "09:00", "10:30", "12:00", "13:00"],
    "prime": ["00:00", "06:00", "07:30", "08:30", "09:30",
              "10:00", "11:00", "12:00", "13:00","14:00", "19:00"]
}

DURACAO_TURNO = pd.Timedelta(hours=6)

def gerar_escala(df):

    escalas = []
    

    for equipe, grupo in df.groupby("equipes"):

        equipe_key = equipe.lower()
        grupo = grupo.sort_values("intervalo")

        hc_max = grupo["agentes_necessarios"].max()
        horarios = HORARIOS_ENTRADA.get(equipe_key, [])

        tecnicos_alocados = []

        for horario in horarios:

            inicio_turno = pd.to_datetime(horario, format="%H:%M")
            fim_turno = inicio_turno + DURACAO_TURNO

            # PRIME regra fixa de 2 tecnicos durante a madrugada 
            if equipe_key == "prime" and horario in ["00:00", "06:00", "19:00"]:
                for _ in range(2):
                    tecnicos_alocados.append(
                        {"inicio": inicio_turno, "fim": fim_turno}
                    )
                continue

            for _, row in grupo.iterrows():

                intervalo = row["intervalo"]

                ativos = sum(
                    1 for t in tecnicos_alocados
                    if t["inicio"].time() <= intervalo.time() < t["fim"].time()
                )

                necessidade = row["agentes_necessarios"]

                if ativos < necessidade and len(tecnicos_alocados) < hc_max:
                    tecnicos_alocados.append(
                        {"inicio": inicio_turno, "fim": fim_turno}
                    )

        # Consolidar por horário
        for horario in horarios:

            inicio_turno = pd.to_datetime(horario, format="%H:%M")

            qtd = sum(
                1 for t in tecnicos_alocados
                if t["inicio"] == inicio_turno
            )

            if qtd > 0:
                escalas.append({
                    "Equipe": equipe,
                    "Inicio_Turno": horario,
                    "Fim_Turno": (
                        inicio_turno + DURACAO_TURNO
                    ).strftime("%H:%M"),
                    "Qtd_Tecnicos": qtd
                })

    return pd.DataFrame(escalas)

df_escala = gerar_escala(df_resultado)

from datetime import datetime

data_str = datetime.now().strftime("%Y-%m")

arquivo_saida = os.path.join(pasta,f"dimensionamento_{data_str}.xlsx"
)

with pd.ExcelWriter(arquivo_saida) as writer:
    #df_resultado.to_excel(writer, sheet_name="Detalhado", index=False)
    resumo.to_excel(writer, sheet_name="Resultado", index=False)
    df_escala.to_excel(writer, sheet_name="Escala", index=False)

print("Dimensionamento Finalizado")
print("Arquivo salvo em:", arquivo_saida)
