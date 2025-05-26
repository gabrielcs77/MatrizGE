import dash
from dash import dcc, html, Input, Output, State, callback_context
import dash_mantine_components as dmc
from dash_iconify import DashIconify
import plotly.graph_objects as go
import pandas as pd
from PIL import Image
import io
import base64

# ==== Constantes Globais para Dimensões da Imagem ====
IMAGE_ORIGINAL_WIDTH = 1705
IMAGE_ORIGINAL_HEIGHT = 1650
TARGET_ASPECT_RATIO = IMAGE_ORIGINAL_WIDTH / IMAGE_ORIGINAL_HEIGHT # Largura / Altura

# ==== Carregar Dados ====
arquivo_excel = 'Matriz GE - EP.xlsx'
arquivo_imagem_fundo = 'Fundo GE.png'
arquivo_imagem_explicacao = 'explicacao.png'

try:
    df = pd.read_excel(arquivo_excel, sheet_name='Análise', skiprows=2)
except FileNotFoundError:
    print(f"Erro: O arquivo Excel '{arquivo_excel}' não foi encontrado.")
    df = pd.DataFrame(columns=[
        "Ignorar", "Modalidade", "Grande Área", "Produto", "Hora Aluno",
        "Faturamento", "Satisfação", "Capacidade de Oferta", "Facilidade de Adesão", "Posição Competitiva",
        "Tamanho Mercado", "Crescimento Mercado", "Vulnerabilidade", "Concorrentes", "Atratividade Mercado",
        "Quadrante"
    ])

df.columns = [
    "Ignorar", "Modalidade", "Grande Área", "Produto", "Hora Aluno",
    "Faturamento", "Satisfação", "Capacidade de Oferta", "Facilidade de Adesão", "Posição Competitiva",
    "Tamanho Mercado", "Crescimento Mercado", "Vulnerabilidade", "Concorrentes", "Atratividade Mercado",
    "Quadrante"
]
if "Ignorar" in df.columns:
    df = df.drop(columns=["Ignorar"])

essential_cols = {
    "Produto": object, "Posição Competitiva": float, "Atratividade Mercado": float,
    "Hora Aluno": float, "Grande Área": object, "Quadrante": object
}
for col, dtype in essential_cols.items():
    if col not in df.columns:
        if dtype == float:
            df[col] = pd.Series(dtype='float64') if col != "Hora Aluno" else 1.0
        else:
            df[col] = pd.Series(dtype='object')
if "Produto" in df.columns:
    df = df[df["Produto"].notnull()]

df["Posição Competitiva"] = pd.to_numeric(df["Posição Competitiva"], errors="coerce")
df["Atratividade Mercado"] = pd.to_numeric(df["Atratividade Mercado"], errors="coerce")
df["Hora Aluno"] = pd.to_numeric(df["Hora Aluno"], errors="coerce").fillna(1)

encoded_image_fundo = None
try:
    image_fundo_pil = Image.open(arquivo_imagem_fundo)
    buffer_fundo = io.BytesIO()
    image_fundo_pil.save(buffer_fundo, format="PNG")
    encoded_image_fundo = base64.b64encode(buffer_fundo.getvalue()).decode()
except FileNotFoundError:
    print(f"Erro: O arquivo de imagem de fundo '{arquivo_imagem_fundo}' não foi encontrado.")
except Exception as e:
    print(f"Erro ao carregar a imagem de fundo: {e}")

encoded_image_explicacao = None
try:
    image_explicacao_pil = Image.open(arquivo_imagem_explicacao)
    buffer_explicacao = io.BytesIO()
    image_explicacao_pil.save(buffer_explicacao, format="PNG")
    encoded_image_explicacao = base64.b64encode(buffer_explicacao.getvalue()).decode()
except FileNotFoundError:
    print(f"Alerta: O arquivo de imagem de explicação '{arquivo_imagem_explicacao}' não foi encontrado.")
except Exception as e:
    print(f"Erro ao carregar a imagem de explicação: {e}")


app = dash.Dash(__name__, suppress_callback_exceptions=True)
server = app.server

def create_filter_popover(filter_id, label_text, button_id):
    fixed_button_width = "210px" 
    return html.Div([
        dmc.Text(label_text, size="md", fw=500, mb=4),
        dmc.Popover(
            withArrow=True, shadow="md", position="bottom-start", trapFocus=False, closeOnClickOutside=True,
            children=[
                dmc.PopoverTarget(
                    dmc.Button(
                        id=button_id, 
                        children=f"Selecionar {label_text.lower()}", 
                        variant="default", 
                        rightSection=DashIconify(icon="carbon:chevron-down", width=16),
                        styles={
                            "width": fixed_button_width, 
                            "textAlign": "left", 
                            "justifyContent": "space-between", 
                            "paddingRight": "10px", 
                            "height":"36px"
                        },
                        size="sm"
                    )
                ),
                dmc.PopoverDropdown(
                    style={
                        "maxHeight": "250px", 
                        "overflowY": "auto", 
                        "minWidth": fixed_button_width, 
                        "padding":"8px"
                    },
                    children=[dmc.CheckboxGroup(id=filter_id, children=[], value=[])]
                )
            ]
        ),
        dmc.Space(h="sm")
    ])

app.layout = dmc.MantineProvider(children=[
    dmc.Container([
        dmc.Title("📊 Matriz GE - Análise de Cursos", order=2, ta="center", my="lg"),
        html.Div([ # Contêiner Flex principal para 3 colunas
            # Coluna 1: Sidebar
            html.Div([
                dmc.Stack([
                    dmc.Title("🎯 Filtros", order=3, mb="sm"),
                    create_filter_popover(filter_id="filtro-area", label_text="Área", button_id="botao-popover-area"),
                    create_filter_popover(filter_id="filtro-quadrante", label_text="Quadrante", button_id="botao-popover-quadrante"),
                    create_filter_popover(filter_id="filtro-produto", label_text="Produto", button_id="botao-popover-produto"),
                    dmc.Group([dmc.Button("🔄 Resetar Filtros", id="botao-reset", color="red", variant="light", size="sm")], mt="md"),
                    dmc.Divider(my="md"),
                    dmc.Checkbox(label="Exibir nomes dos produtos", id="exibir-texto", checked=True, mb="sm", size="sm"),
                    dmc.Button("⬇️ Baixar Gráfico (PNG)", id="botao-download", variant="outline", mt="md", size="sm"),
                    dcc.Download(id="download-imagem")
                ])
            ], style={ 
                'width': '20%', 
                'minWidth': '260px', 
                'padding': '20px', 
                'borderRight': '1px solid #eee'
            }),
            # Coluna 2: Área do Gráfico
            html.Div([
                dcc.Graph(id='grafico-matriz', config={'displayModeBar': False})
            ], style={ 
                'width': '50%', 
                'display': 'flex',
                'justifyContent': 'center',
                'alignItems': 'flex-start',
                'padding': '0 10px 0 5px', 
            }),
            # Coluna 3: Área da Imagem de Explicação
            html.Div([
                (html.Img(src=f"data:image/png;base64,{encoded_image_explicacao}", 
                         style={'maxWidth': '100%', 
                                'height': 'auto', 
                                'display': 'block', 
                                'marginTop': '30px'}) 
                 if encoded_image_explicacao 
                 else dmc.Text("Imagem 'explicacao.png' não encontrada.", color="red", align="center", mt="xl"))
            ], style={ 
                'width': '30%', 
                'padding': '0 40px 0 20px', 
                'display': 'flex',
                'alignItems': 'flex-start', 
                'justifyContent': 'center' 
            })
        ], style={ 
            'display': 'flex', 
            'flexDirection': 'row',
            'alignItems': 'flex-start' 
        })
    ], fluid=True, px="xl")
])

@app.callback(
    [Output('filtro-area', 'children'), Output('filtro-quadrante', 'children'), Output('filtro-produto', 'children'),
     Output('grafico-matriz', 'figure'),
     Output('filtro-area', 'value'), Output('filtro-quadrante', 'value'), Output('filtro-produto', 'value'),
     Output('botao-popover-area', 'children'), Output('botao-popover-quadrante', 'children'), Output('botao-popover-produto', 'children')],
    [Input('filtro-area', 'value'), 
     Input('filtro-quadrante', 'value'), 
     Input('filtro-produto', 'value'),
     Input('exibir-texto', 'checked'),
     Input('botao-reset', 'n_clicks')],
    prevent_initial_call=False
)
def atualizar_tudo(selected_areas, selected_quadrantes, selected_produtos, 
                   exibir_texto, reset_n_clicks):
    triggered_input = callback_context.triggered_id
    areas_val = selected_areas if selected_areas is not None else []
    quadrantes_val = selected_quadrantes if selected_quadrantes is not None else []
    produtos_val = selected_produtos if selected_produtos is not None else []

    fixed_cor_bolha = "#FFFFFF" 
    fixed_transparencia_percent = 55 

    min_bubble_size_pref = 1
    target_max_bubble_size_pref = 150

    df_plot = df.copy()
    df_options_source = df.copy()

    if triggered_input == "botao-reset":
        areas_val, quadrantes_val, produtos_val = [], [], []
    else:
        if areas_val: df_plot = df_plot[df_plot['Grande Área'].isin(areas_val)]
        if quadrantes_val: df_plot = df_plot[df_plot['Quadrante'].isin(quadrantes_val)]
        if produtos_val: df_plot = df_plot[df_plot['Produto'].isin(produtos_val)]

    df_temp_area_opts = df_options_source.copy();
    if quadrantes_val: df_temp_area_opts = df_temp_area_opts[df_temp_area_opts['Quadrante'].isin(quadrantes_val)]
    if produtos_val: df_temp_area_opts = df_temp_area_opts[df_temp_area_opts['Produto'].isin(produtos_val)]
    unique_areas = sorted(df_temp_area_opts["Grande Área"].dropna().unique()) if "Grande Área" in df_temp_area_opts else []
    children_areas = [dmc.Checkbox(label=str(i), value=i, styles={"labelWrapper": {"width": "100%"}}) for i in unique_areas]

    df_temp_quad_opts = df_options_source.copy();
    if areas_val: df_temp_quad_opts = df_temp_quad_opts[df_temp_quad_opts['Grande Área'].isin(areas_val)]
    if produtos_val: df_temp_quad_opts = df_temp_quad_opts[df_temp_quad_opts['Produto'].isin(produtos_val)]
    unique_quadrantes = sorted(df_temp_quad_opts["Quadrante"].dropna().unique()) if "Quadrante" in df_temp_quad_opts else []
    children_quadrantes = [dmc.Checkbox(label=str(i), value=i, styles={"labelWrapper": {"width": "100%"}}) for i in unique_quadrantes]

    df_temp_prod_opts = df_options_source.copy();
    if areas_val: df_temp_prod_opts = df_temp_prod_opts[df_temp_prod_opts['Grande Área'].isin(areas_val)]
    if quadrantes_val: df_temp_prod_opts = df_temp_prod_opts[df_temp_prod_opts['Quadrante'].isin(quadrantes_val)]
    unique_produtos = sorted(df_temp_prod_opts["Produto"].dropna().unique()) if "Produto" in df_temp_prod_opts else []
    children_produtos = [dmc.Checkbox(label=str(i), value=i, styles={"labelWrapper": {"width": "100%"}}) for i in unique_produtos]
    
    def get_button_text(selected_vals, default_single, default_plural_prefix):
        if not selected_vals: return f"Selecionar {default_single.lower()}"
        if len(selected_vals) == 1:
            txt = str(selected_vals[0]); max_l = 25 
            return txt if len(txt) <= max_l else txt[:max_l-3] + "..."
        return f"{len(selected_vals)} {default_plural_prefix} selecionad{'as' if default_single.endswith('a') else 'os'}"
    texto_botao_area = get_button_text(areas_val, "Área", "áreas")
    texto_botao_quadrante = get_button_text(quadrantes_val, "Quadrante", "quadrantes")
    texto_botao_produto = get_button_text(produtos_val, "Produto", "produtos")
    
    fig = go.Figure()
    cols_for_plot = ["Hora Aluno", "Quadrante", "Posição Competitiva", "Atratividade Mercado", "Produto"]
    for col_plot in cols_for_plot:
        if col_plot not in df_plot.columns:
            if col_plot == "Hora Aluno": df_plot[col_plot] = 1.0
            elif col_plot in ["Posição Competitiva", "Atratividade Mercado"]: df_plot[col_plot] = pd.NA
            else: df_plot[col_plot] = pd.Series(dtype='object')
    
    max_hora_aluno = df_plot["Hora Aluno"].max() if not (df_plot.empty or df_plot["Hora Aluno"].isnull().all() or "Hora Aluno" not in df_plot.columns) else 1.0
    
    fator_escala = max_hora_aluno / target_max_bubble_size_pref if pd.notna(max_hora_aluno) and max_hora_aluno > 0 else 1.0
    if fator_escala == 0: fator_escala = 1.0

    if not df_plot.empty and "Quadrante" in df_plot.columns:
        for quadrante_unico in df_plot["Quadrante"].dropna().unique():
            df_q = df_plot[df_plot["Quadrante"] == quadrante_unico].copy()
            if not df_q.empty and all(col in df_q.columns for col in cols_for_plot):
                df_q["Produto_Formatado"] = df_q["Produto"].astype(str).str.replace("TÉCNICO EM ", "TÉCNICO EM<br>", regex=False)
                posicoes_texto = []
                for _, row in df_q.iterrows():
                    x_pv, y_am = row["Posição Competitiva"], row["Atratividade Mercado"]
                    pos = "middle center"
                    if pd.notna(x_pv) and pd.notna(y_am):
                        pos_x = "center"; pos_y = "middle"
                        if x_pv <= 2.0: pos_x = "right"
                        elif x_pv >= 8.0: pos_x = "left"
                        if y_am <= 2.0: pos_y = "top"
                        elif y_am >= 8.0: pos_y = "bottom"
                        pos = f"{pos_y} {pos_x}"
                    posicoes_texto.append(pos)
                df_q["posicao_texto_calculada"] = posicoes_texto
                df_q["marker_size_calculada"] = (df_q["Hora Aluno"] / fator_escala).apply(lambda s: max(s if pd.notna(s) else 0, min_bubble_size_pref))
                
                fig.add_trace(go.Scatter(
                    x=df_q["Posição Competitiva"], y=df_q["Atratividade Mercado"],
                    mode="markers+text" if exibir_texto else "markers", name=str(quadrante_unico if pd.notna(quadrante_unico) else "N/A"),
                    marker=dict(size=df_q["marker_size_calculada"], color=fixed_cor_bolha, opacity=max(0, min(1, (100 - fixed_transparencia_percent) / 100))),
                    text=df_q["Produto_Formatado"] if exibir_texto else None, textposition=df_q["posicao_texto_calculada"] if exibir_texto else None,
                    textfont=dict(color="black", size=14, family="Bahnschrift, Arial, sans-serif"), customdata=df_q['Produto'],
                    hovertemplate="<b>%{customdata}</b><br><b>Posição Competitiva:</b> %{x:.1f}<br><b>Atratividade:</b> %{y:.1f}<extra></extra>"
                ))

    if encoded_image_fundo:
        fig.add_layout_image(dict(source="data:image/png;base64," + encoded_image_fundo, xref="x domain", yref="y domain", x=0, y=1, sizex=1, sizey=1, sizing="stretch", opacity=1, layer="below"))

    graph_pixel_width = 800
    graph_pixel_height = int(graph_pixel_width / TARGET_ASPECT_RATIO)

    fig.update_layout(
        width=graph_pixel_width, height=graph_pixel_height, autosize=False,
        xaxis=dict(
            range=[0, 10], autorange=False, fixedrange=True, 
            title=dict(
                text="<b>Posição Competitiva</b>",  # MODIFICADO: Seta removida
                font=dict(size=17)
            ), 
            tickfont=dict(size=14),
            showgrid=False, zeroline=False, tickmode="linear", tick0=0, dtick=1, ticks="outside", ticklen=6, tickwidth=1, tickcolor='grey', linecolor='lightgrey'
        ),
        yaxis=dict(
            range=[0, 10], autorange=False, fixedrange=True, 
            title=dict(
                text="<b>Atratividade de Mercado</b>",  # MODIFICADO: Seta removida
                font=dict(size=17)
            ), 
            tickfont=dict(size=14),
            scaleanchor="x", scaleratio=(IMAGE_ORIGINAL_HEIGHT / IMAGE_ORIGINAL_WIDTH),
            showgrid=False, zeroline=False, tickmode="linear", tick0=0, dtick=1, ticks="outside", ticklen=6, tickwidth=1, tickcolor='grey', linecolor='lightgrey'
        ),
        plot_bgcolor='rgba(255,255,255,0.1)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False, 
        margin=dict(l=60, r=30, t=30, b=60), 
        hovermode="closest"
    )
    if df_plot.empty and (areas_val or quadrantes_val or produtos_val):
         fig.update_layout(xaxis_visible=False, yaxis_visible=False, annotations=[dict(text="Nenhum curso para os filtros.", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=16))])
    return (children_areas, children_quadrantes, children_produtos, fig, areas_val, quadrantes_val, produtos_val, texto_botao_area, texto_botao_quadrante, texto_botao_produto)

@app.callback(
    Output("download-imagem", "data"),
    Input("botao-download", "n_clicks"),
    State("grafico-matriz", "figure"),
    prevent_initial_call=True
)
def baixar_grafico(n_clicks, current_figure_dict):
    if n_clicks and current_figure_dict and current_figure_dict.get('data'):
        fig_layout = current_figure_dict.get('layout', {})
        on_screen_width = fig_layout.get('width')
        on_screen_height = fig_layout.get('height')
        
        if not isinstance(on_screen_width, (int, float)) or on_screen_width <= 0: 
            on_screen_width = 800 
        if not isinstance(on_screen_height, (int, float)) or on_screen_height <= 0:
            on_screen_height = int(on_screen_width / TARGET_ASPECT_RATIO) if TARGET_ASPECT_RATIO != 0 else int(on_screen_width * (IMAGE_ORIGINAL_HEIGHT / IMAGE_ORIGINAL_WIDTH))
        
        scale_factor = 1.0
        if on_screen_width != 0 :
             scale_factor = IMAGE_ORIGINAL_WIDTH / on_screen_width
        else:
            print("Aviso: on_screen_width é zero no callback de download, usando scale_factor padrão de 1.0.")

        img_bytes = go.Figure(current_figure_dict).to_image(
            format="png",
            width=int(on_screen_width) if on_screen_width else IMAGE_ORIGINAL_WIDTH,
            height=int(on_screen_height) if on_screen_height else IMAGE_ORIGINAL_HEIGHT,
            scale=scale_factor
        )
        return dcc.send_bytes(img_bytes, "matriz_ge.png")
    return dash.no_update

if __name__ == '__main__':
    app.run(debug=True)
