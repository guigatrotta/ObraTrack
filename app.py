from dash import Dash, dcc, html, Output, Input, State, no_update
import plotly.graph_objects as go
import pandas as pd
import datetime
import webbrowser
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re
from dateutil import parser

from dateutil.parser import parse
import pandas as pd

def parse_data(data_str):
    try:
        # Trata números do Excel (ex: 45500 -> datetime)
        num = float(data_str)
        if 10000 < num < 60000:  # Faixa típica de datas Excel
            return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(num), unit="D")
    except:
        pass
    try:
        return pd.to_datetime(parse(str(data_str).strip(), dayfirst=False))
    except:
        return pd.NaT



# === Autenticação Google Sheets ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
import json
import os

credentials_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)

client = gspread.authorize(creds)

# === Leitura da planilha de dados principais ===
sheet = client.open("01_ Fluxo de Caixa e Compras").worksheet("Acumulado por Categoria")
data = sheet.get_all_values()
df = pd.DataFrame(data[1:], columns=["Data", "Categoria", "SomaDiaria", "Acumulado"])
df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors='coerce')
df["Categoria"] = df["Categoria"].astype(str).str.strip()
df["Acumulado"] = pd.to_numeric(
    df["Acumulado"].str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
    errors='coerce'
)
df = df.dropna(subset=["Data", "Categoria", "Acumulado"])
df_pivot = df.pivot_table(index="Data", columns="Categoria", values="Acumulado", aggfunc="last").ffill()

# Reordena categorias por valor final acumulado
df_final = df_pivot.iloc[-1].sort_values(ascending=False)
categorias = df_final.index.tolist()

# === Leitura da planilha de marcos ===
sheet_marcos = client.open("02_ Cronograma Físico").worksheet("REAL")
dados_marcos = sheet_marcos.get_all_values()

for i, linha in enumerate(dados_marcos):
    if "ID" in linha and "FASES/SUBFASES" in linha and any(re.match(r"FINAL|Data\\s*final", col, re.IGNORECASE) for col in linha):
        header_row = i
        break
else:
    raise ValueError("Cabeçalho com colunas esperadas não encontrado.")

df_marcos = pd.DataFrame(dados_marcos[header_row + 1:], columns=dados_marcos[header_row])
df_marcos.columns = df_marcos.columns.str.strip().str.upper()
df_marcos = df_marcos[["ID", "FASES/SUBFASES", "FINAL"]]
df_marcos = df_marcos.rename(columns={"ID": "Codigo", "FASES/SUBFASES": "Descricao", "FINAL": "Data"})

def parse_data(data_str):
    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return pd.to_datetime(data_str.strip(), format=fmt)
        except:
            continue
    return pd.NaT

from pandas.errors import ParserError

def robust_excel_date(val):
    try:
        # Tenta como número (data serial do Excel)
        num = float(val)
        if 10000 < num < 60000:
            return pd.to_datetime("1899-12-30") + pd.to_timedelta(int(num), unit="D")
    except:
        pass
    try:
        # Tenta como string
        return pd.to_datetime(val, dayfirst=False, errors="coerce")
    except (ParserError, ValueError, TypeError):
        return pd.NaT

df_marcos["Data"] = df_marcos["Data"].map(robust_excel_date)
df_marcos = df_marcos.dropna(subset=["Data"])
print("📌 Total de marcos com data válida:", len(df_marcos))
print(df_marcos[["Codigo", "Descricao", "Data"]])


print("📌 Marcos identificados:")
for _, row in df_marcos.iterrows():
    print(f"- {row['Codigo']:>5} | {row['Descricao']:<60} | {row['Data'].strftime('%d/%m/%Y') if pd.notna(row['Data']) else 'N/A'}")

print(">>> Tipo da coluna Data:", df_marcos["Data"].dtype)
print(df_marcos["Data"].head())

# === Criação da lista de opções de marcos ===
df_marcos["Opcao"] = df_marcos.apply(
    lambda row: f'{row["Codigo"]} - {row["Descricao"][:40]}', axis=1
)
df_marcos["EhNumerico"] = df_marcos["Codigo"].astype(str).str.strip().str.match(r"^\d+$")
opcoes_marcos = df_marcos["Opcao"].tolist()
opcoes_marcos_visiveis = df_marcos[df_marcos["EhNumerico"]]["Opcao"].tolist()



# === Função principal de geração do gráfico ===
def criar_fig(visiveis=None, marcos_visiveis=None):
    fig = go.Figure()
    for i, cat in enumerate(categorias):
        visible = True if visiveis is None else visiveis[i]
        fig.add_trace(go.Scatter(
            x=df_pivot.index,
            y=df_pivot[cat],
            mode="lines+markers",
            name=cat,
            visible=visible,
            hovertemplate=f"{cat}<br>%{{x|%d/%m/%Y}}<br>R$ %{{y:,.2f}}<extra></extra>"
        ))

    y_max = max([df_pivot[cat].max() for i, cat in enumerate(categorias) if visiveis is None or visiveis[i]])

    for _, row in df_marcos.iterrows():
        if marcos_visiveis and row["Opcao"] not in marcos_visiveis:
            continue
        fig.add_vline(x=row["Data"], line_dash="dot", line_color="gray")
        fig.add_annotation(
            x=row["Data"],
            y=y_max,
            text=f'{row["Codigo"]} {row["Descricao"]}',
            showarrow=True,
            arrowhead=1,
            ax=0,
            ay=-40,
            font=dict(color="gray"),
            bgcolor="white",
            bordercolor="gray"
        )

    fig.update_layout(
        title="Gasto acumulado por categoria (com fases da obra)",
        xaxis_title="Data",
        yaxis_title="Valor acumulado (R$)",
        hovermode="x unified",
        template="plotly_white",
        yaxis_tickprefix="R$ ",
        margin=dict(l=40, r=220, t=60, b=40),
    )
    return fig


# === App ===
app = Dash(__name__)
app.title = "Gráfico Interativo de Gastos"

app.layout = html.Div([
    html.H2("Gasto acumulado por categoria"),
    dcc.Graph(id="grafico-interativo"),
    html.Div([
        html.Button("Ajustar eixo Y", id="ajustar-eixo", n_clicks=0),
        html.Button("Iniciar gráfico", id="resetar", n_clicks=0)
    ], style={"marginTop": "10px", "marginBottom": "10px"}),
    dcc.Store(id="visibilidade", data=[True]*len(categorias)),
    dcc.Dropdown(
        id="dropdown-marcos",
        options=[{"label": opcao, "value": opcao} for opcao in opcoes_marcos],
        value=opcoes_marcos_visiveis,
        multi=True,
        placeholder="Filtrar marcos a exibir...",
        style={"marginTop": "10px", "marginBottom": "10px"}
    ),

])

@app.callback(
    Output("grafico-interativo", "figure"),
    Input("ajustar-eixo", "n_clicks"),
    State("visibilidade", "data"),
    State("dropdown-marcos", "value"),
    prevent_initial_call=True
)
def ajustar_y(n_clicks, visiveis, marcos_visiveis):
    if n_clicks > 0:
        # Usa as categorias visíveis no Store para reconstruir
        fig = criar_fig(visiveis=visiveis, marcos_visiveis=marcos_visiveis)
        visiveis_series = [df_pivot[cat] for cat, vis in zip(categorias, visiveis) if vis]
        y_max = max([serie.max() for serie in visiveis_series]) if visiveis_series else 1
        fig.update_yaxes(range=[0, y_max * 1.05])
        return fig
    return no_update



@app.callback(
    Output("grafico-interativo", "figure", allow_duplicate=True),
    Input("resetar", "n_clicks"),
    State("dropdown-marcos", "value"),
    prevent_initial_call=True
)
def resetar_grafico(n, marcos_visiveis):
    if n > 0:
        return criar_fig(marcos_visiveis=marcos_visiveis)
    return no_update

@app.callback(
    Output("visibilidade", "data"),
    Input("grafico-interativo", "restyleData"),
    State("visibilidade", "data"),
    prevent_initial_call=True
)
def atualizar_visiveis(restyle_data, estado_atual):
    if restyle_data and "visible" in restyle_data[0]:
        indices = restyle_data[1]
        novo_estado = list(estado_atual)
        novo_valor = restyle_data[0]["visible"][0]
        for i in indices:
            novo_estado[i] = (novo_valor != "legendonly")
        return novo_estado
    return estado_atual


@app.callback(
    Output("grafico-interativo", "figure", allow_duplicate=True),
    Input("dropdown-marcos", "value"),
    State("visibilidade", "data"),
    prevent_initial_call=True
)
def atualizar_marcos(marcos_visiveis, visiveis):
    return criar_fig(visiveis, marcos_visiveis)


def resetar_grafico(n):
    if n > 0:
        return criar_fig()
    return no_update



# === Imprime os marcos identificados ===
print("\n📌 Marcos identificados:")
for _, row in df_marcos.iterrows():
    print(f"- {row['Codigo']:<6} | {row['Descricao']:<50} | {row['Data'].strftime('%d/%m/%Y')}")

print("\n📌 Marcos identificados:")
for _, row in df_marcos.iterrows():
    print(f"- {row.get('Codigo', '').strip():<6} | {row['Descricao']:<50} | {row['Data'].strftime('%d/%m/%Y')}")


if __name__ == "__main__":
    app.run_server(host="0.0.0.0", port=8080)


