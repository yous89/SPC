import dash
import pandas as pd
import plotly.graph_objs as go
from dash import Input, Output, dcc, html

# Einlesen der Daten
excel_file = r"C:\Users\youse\OneDrive\Desktop\prisma\BE01-BE13_Batchdata (1).xlsx"
xlsx = pd.ExcelFile(excel_file)

# Erstelle ein Dictionary, um die DataFrames zu speichern
df_processes = {}

# Iterieren über die 13 Batches, sheet 14 ist Controlline Information
for i, sheet_name in enumerate(xlsx.sheet_names[:13], start=1):
    df_processes[f"data_process{i:02}"] = xlsx.parse(sheet_name)
    # Die Spaltennamen als Strings schreiben
    df_processes[f"data_process{i:02}"].columns = df_processes[f"data_process{i:02}"].columns.astype(str)

# Initialisierung der Dash-App und Setzen des Titels
app = dash.Dash(__name__)
app.title = "Control Chart Analyse"

# Layout der App definieren
app.layout = html.Div(
    [
        html.H1("Control Chart Analyse"),
        # Regeln als Checkboxen
        html.Div(
            [
                html.Label("Ankreuzen, welche Trendregeln geprüft werden sollen"),
                dcc.Checklist(
                    id="rule_selector",
                    options=[{"label": f"Rule {i}", "value": f"rule{i}"} for i in range(1, 11)],
                    value=[],
                    inline=True,
                ),
            ]
        ),
        # Prozessschritt auswählen (entspricht sheet der Exceldatei)
        html.Div(
            [
                html.Label("Wählen Sie einen Prozessschritt aus"),
                dcc.Dropdown(
                    id="process_selector",
                    options=[{"label": key, "value": key} for key in df_processes.keys()],
                    value="data_process01",  # Standardwert
                    clearable=False,
                ),
            ],
            style={"width": "50%"},
        ),
        # Dropdownmenü für die zu analysierende Spalte
        html.Div(
            [
                html.Label("Wählen Sie eine Spalte aus:"),
                dcc.Dropdown(id="column_selector"),
            ],
            style={"margin-top": "20px"},
        ),
        # Auswahl der Straßen
        html.Div(
            [
                html.Label("Straßen auswählen:"),
                dcc.Checklist(id="street_selector", inline=True),
            ],
            style={"margin-top": "20px"},
        ),
        # Eingabefelder für Spezifikationslimits
        html.Div(
            [
                html.Label("Definieren Sie Spezifikationslimits:"),
                html.Div(
                    [
                        html.Label("USL (Upper Specification Limit):"),
                        dcc.Input(id="usl_input", type="number", value=None, placeholder="USL", step=0.1),
                    ],
                    style={"display": "inline-block", "margin-right": "20px"},
                ),
                html.Div(
                    [
                        html.Label("LSL (Lower Specification Limit):"),
                        dcc.Input(id="lsl_input", type="number", value=None, placeholder="LSL", step=0.1),
                    ],
                    style={"display": "inline-block"},
                ),
            ],
            style={"margin-top": "20px"},
        ),
        dcc.Graph(id="line_chart"),
    ]
)


# Callback, der die Auswahl der Spalten und Straßen aktualisiert
@app.callback(
    [
        Output("column_selector", "options"),
        Output("column_selector", "value"),
        Output("street_selector", "options"),
        Output("street_selector", "value"),
    ],
    Input("process_selector", "value"),
)
def update_column_and_street_dropdown(selected_process):
    df = df_processes[selected_process]  # Dataframe des ausgewählten Prozesses
    excluded_cols = ["Straße", "BatchID", "Teilcharge"]

    # Spaltenoptionen
    column_options = [
        {"label": col, "value": col} for col in df.columns if col not in excluded_cols
    ]
    default_column = column_options[0]["value"] if column_options else None

    # Straßenoptionen
    street_options = [{"label": str(street), "value": str(street)} for street in df["Straße"].unique()]
    default_streets = [opt["value"] for opt in street_options]

    return column_options, default_column, street_options, default_streets


# Callback: Aktualisiere den Plot basierend auf der Auswahl
@app.callback(
    Output("line_chart", "figure"),
    [
        Input("process_selector", "value"),
        Input("column_selector", "value"),
        Input("street_selector", "value"),
        Input("rule_selector", "value"),
        Input("usl_input", "value"),
        Input("lsl_input", "value"),
    ],
)
def update_control_charts(selected_process, selected_column, selected_streets, rule_selection, usl, lsl):
    df = df_processes[selected_process]

    # Filter basierend auf den ausgewählten Straßen
    if selected_streets:
        df = df[df["Straße"].astype(str).isin(selected_streets)]

    mean_value = df[selected_column].mean()
    std_dev = df[selected_column].std()

    data = [
        go.Scatter(
            x=df["BatchID"],
            y=df[selected_column],
            mode="lines+markers",
            name=selected_column,
            line=dict(color="black"),
        ),
        go.Scatter(
            x=df["BatchID"],
            y=[mean_value] * len(df),
            mode="lines",
            line=dict(dash="dash", color="rgb(25,232,44)"),
            name="Mean",
        ),
        go.Scatter(
            x=df["BatchID"],
            y=[mean_value + 3 * std_dev] * len(df),
            mode="lines",
            line=dict(dash="dash", color="rgb(9,70,14)"),
            name="UCL (3σ)",
        ),
        go.Scatter(
            x=df["BatchID"],
            y=[mean_value - 3 * std_dev] * len(df),
            mode="lines",
            line=dict(dash="dash", color="rgb(9,70,14)"),
            name="LCL (3σ)",
        ),
    ]

    # Spezifikationslimits hinzufügen, falls definiert
    if usl is not None:
        data.append(
            go.Scatter(
                x=df["BatchID"],
                y=[usl] * len(df),
                mode="lines",
                line=dict(dash="dashdot", color="red"),
                name="USL",
            )
        )
    if lsl is not None:
        data.append(
            go.Scatter(
                x=df["BatchID"],
                y=[lsl] * len(df),
                mode="lines",
                line=dict(dash="dashdot", color="blue"),
                name="LSL",
            )
        )

    layout = go.Layout(
        title=f"{selected_column} in {selected_process} (Straßen: {', '.join(selected_streets)})",
        xaxis=dict(title="Batchnummer"),
        yaxis=dict(title=selected_column),
    )

    fig = {"data": data, "layout": layout}
    return fig


# Starten der Dash-App
if __name__ == "__main__":
    app.run_server(debug=True, port=8051)
