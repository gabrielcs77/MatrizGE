import dash
from dash import dcc, html, Input, Output, State, callback_context
import dash_mantine_components as dmc
from dash_iconify import DashIconify # Importação necessária
import plotly.graph_objects as go
import pandas as pd
from PIL import Image
import io
import base64
import os

# ==== Define o diretório base do script para caminhos de arquivo robustos ====
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ==== Constantes Globais para Dimensões da Imagem ====
IMAGE_ORIGINAL_WIDTH = 1705
IMAGE_ORIGINAL_HEIGHT = 1650
TARGET_ASPECT_RATIO = IMAGE_ORIGINAL_WIDTH / IMAGE_ORIGINAL_HEIGHT if IMAGE_ORIGINAL_HEIGHT != 0 else 16/9 # Largura / Altura

# ==== Constantes para Dimensionamento Responsivo do Gráfico ====
GRAPH_CONTAINER_MIN_WIDTH = 400  # px, Largura mínima para o contêiner do gráfico
GRAPH_CONTAINER_MAX_WIDTH = 850  # px, Largura máxima para o contêiner do gráfico (ajuste se necessário)

# ==== Carregar Dados ====
arquivo_excel = os.path.join(BASE_DIR, 'Matriz GE - EP.xlsx')
arquivo_imagem_fundo = os.path.join(BASE_DIR, 'Fundo GE.png')
arquivo_imagem_explicacao = os.path.join(BASE_DIR, 'explicacao.png')

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


app = dash.Dash(__name__, suppress_callback_exceptions=True, title="Matriz GE - Educação Profissional")
server = app.server

# ==== Função para criar Popover de Filtro ====
def create_filter_popover(filter_id, label_text, button_id):
    fixed_button_width = "300px"
    return html.Div([
        dcc.Store(id=f"store-label-{filter_id}", data=f"Selecionar {label_text.lower()}"),
        dmc.Text(label_text, size="md", fw=500, mb=4),
        dmc.Popover(
            withArrow=True,
            shadow="md",
            position="bottom-start",
            trapFocus=False,
            closeOnClickOutside=True,
            children=[
                dmc.PopoverTarget(
                    dmc.Button(
                        id=button_id,
                        variant="default",
                        size="sm",
                        style={
                            "width": fixed_button_width,
                            "height": "36px",
                            "padding": "0 12px",
                            "display": "flex",
                            "alignItems": "center",
                            "justifyContent": "space-between",
                            "textAlign": "left",
                            "overflow": "hidden"
                        },
                        children=html.Div([
                            html.Div(
                                id=f"label-{filter_id}",
                                children=f"Selecionar {label_text.lower()}",
                                style={
                                    "flexGrow": 1,
                                    "whiteSpace": "nowrap",
                                    "overflow": "hidden",
                                    "textOverflow": "ellipsis",
                                    "textAlign": "left",
                                    "paddingRight": "8px"
                                }
                            ),
                            html.Div(
                                "▼", # Seta para baixo
                                style={
                                    "flexShrink": 0,
                                    "minWidth": "16px",
                                    "textAlign": "right"
                                }
                            )
                        ], style={
                            "width": "100%",
                            "display": "flex",
                            "alignItems": "center"
                        })
                    )
                ),
                dmc.PopoverDropdown(
                    style={
                        "maxHeight": "250px",
                        "overflowY": "auto",
                        "minWidth": fixed_button_width,
                        "padding": "8px"
                    },
                    children=[
                        dmc.CheckboxGroup(id=filter_id, children=[], value=[])
                    ]
                )
            ]
        ),
        dmc.Space(h="sm")
    ])

# ==== Layout da Página Principal (Matriz GE) ====
def create_layout_matriz_ge():
    return dmc.Container([
        dmc.Title("📊 Ciclo de Vida de Produtos - Educação Profissional", order=2, ta="center", my="lg"),
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
            
            # Coluna 2: Área do Gráfico --- AJUSTES PARA RESPONSIVIDADE E PROPORÇÃO ---
            html.Div(
                id='graph-column-container', 
                children=[
                    dcc.Graph(
                        id='grafico-matriz', 
                        config={'displayModeBar': False},
                        style={'width': '100%', 'height': '100%'} # Gráfico preenche o contêiner
                    )
                ], 
                style={
                    'width': '50%', 
                    'minWidth': f'{GRAPH_CONTAINER_MIN_WIDTH}px',
                    'maxWidth': f'{GRAPH_CONTAINER_MAX_WIDTH}px',
                    'aspectRatio': str(TARGET_ASPECT_RATIO), 
                    'margin': '0 auto', 
                    'overflow': 'hidden', 
                    'padding': '0 10px 0 5px', 
                    'display': 'flex', 
                    'flexDirection': 'column'
                }
            ),
            # --- FIM DOS AJUSTES NA COLUNA DO GRÁFICO ---

            # Coluna 3: Área da Imagem de Explicação
            html.Div([
                (html.Img(src=f"data:image/png;base64,{encoded_image_explicacao}",
                          style={'maxWidth': '100%', 'height': 'auto', 'display': 'block', 'marginTop': '30px'})
                 if encoded_image_explicacao
                 else dmc.Text("Imagem 'explicacao.png' não encontrada.", c="red", ta="center", mt="xl"))
            ], style={
                'width': '30%', 
                'minWidth': '200px', # Adicionado minWidth
                'padding': '0 40px 0 20px',
                'display': 'flex',
                'alignItems': 'flex-start',
                'justifyContent': 'center'
            })
        ], style={
            'display': 'flex',
            'flexDirection': 'row',
            'alignItems': 'flex-start',
            'flexWrap': 'nowrap' # Para garantir que as colunas não quebrem linha prematuramente
        })
    ], fluid=True, px="xl")

# ==== Layout da Página de Nota Técnica ====
def create_layout_nota_tecnica():
    texto_da_nota = """
    A **Matriz GE** é uma ferramenta de **análise estratégica** desenvolvida para ajudar empresas a **avaliar o portfólio de produtos**. Ela foi criada pela consultoria McKinsey em parceria com a General Electric como uma **evolução da Matriz BCG**. É composta por uma **grade 3x3** (nove quadrantes), que avalia cada produto com base em dois critérios principais:

    &nbsp;

    1. **Atratividade de Mercado** (eixo vertical):

        • Avalia o quão atraente é o mercado para este produto.
        
        • Fatores considerados: tamanho de mercado, crescimento de mercado, vulnerabilidade de mercado e volume de concorrentes.

    2. **Posição Competitiva** (eixo horizontal):

        • Mede o quão bem-posicionado o produto está em relação ao mercado.

        • Fatores considerados: faturamento, satisfação de clientes, capacidade de oferta e facilidade de adesão.

        &nbsp;

    Esses dois eixos são divididos em **baixo**, **médio** e **alto**, gerando nove zonas diferentes.

    &nbsp;

    • **Investir / Crescer (zona superior direita)**:
        • Alta atratividade + alta força competitiva.
        • Estratégia: expandir, alocar recursos, inovar.

    • **Selecionar / Manter (zonas centrais)**:
        • Atração e força competitiva medianas.
        • Estratégia: manter o desempenho atual, avaliar oportunidades com cautela.

    • **Desinvestir / Colher (zona inferior esquerda)**:
        • Baixa atratividade + baixa força competitiva.
        • Estratégia: reduzir investimentos, sair do mercado.

    &nbsp;
    
    Para os itens de análise interna foram considerados os semestres 2024.1, 2024.2 e 2025.1. Os intervalos de classificação são:

    **Faturamento**

    Nota 1: até R$ 100.000,00

    Nota 2: até R$ 250.000,00

    Nota 3: até R$ 350.000,00

    Nota 4: até R$ 700.000,00

    Nota 5: acima de R$ 700.000,00

    **Satisfação**

    Nota 1: NPS até 20%

    Nota 2: NPS entre 21% e 40%

    Nota 3: NPS entre 41% e 60%

    Nota 4: NPS entre 61% e 80%

    Nota 5: NPS acima de 81%

    **Capacidade de Oferta (execução do planejamento)**

    Nota 1: Menor ou igual que 50%

    Nota 2: Maior que 50% e menor ou igual que 70%

    Nota 3: Maior que 70% e menor ou igual que 80%

    Nota 4: Maior que 80% e menor ou igual que 90%

    Nota 5: Maior que 90%

    **Facilidade de Adesão (fechamento de turmas)**

    Nota 1: Maior que 120 dias

    Nota 2: Maior que 100 dias e menor ou igual que 120 dias

    Nota 3: Maior que 80 dias e menor ou igual que 100 dias

    Nota 4: Maior que 60 dias e menor ou igual que 80 dias

    Nota 5: Menor ou igual que 60 dias

    **Tamanho de Mercado (estoque de empregos)**

    Nota 1: Saldo menor ou igual a 500

    Nota 2: Saldo maior que 500 e menor ou igual a 1000

    Nota 3: Saldo maior que 1000 e menor ou igual a 2000

    Nota 4: Saldo maior que 2000 e menor ou igual a 3000

    Nota 5: Saldo maior que 3000

    **Crescimento de Mercado (produto entre o estoque de empregos e o mapa do trabalho)**

    Nota 1: Resultado menor ou igual a 20000

    Nota 2: Resultado maior que 20000 e menor ou igual a 40000

    Nota 3: Resultado maior que 40000 e menor ou igual a 60000

    Nota 4: Resultado maior que 60000 e menor ou igual a 80000

    Nota 5: Saldo maior que 80000

    **Vulnerabilidade (sensibilidade aos fatores externos – análise PESTEL)**

    Nota 5: cenário muito favorável

    Nota 4: cenário favorável

    Nota 3: cenário neutro

    Nota 2: cenário desfavorável

    Nota 1: cenário muito desfavorável

    **Volume de Concorrentes**

    Nota 5: nenhum concorrente relevante

    Nota 4: há concorrentes, mas nenhum relevante

    Nota 3: há concorrentes, mas somente 1 é relevante

    Nota 2: há concorrentes e 2 são relevantes

    nota 1: há 3 ou mais concorrentes relevantes
    """
    return dmc.Container([
        dmc.Title("📄 Nota Técnica", order=2, ta="center", my="lg"),
        dmc.Paper(
            shadow="xs",
            p="xl",
            children=[
                dcc.Markdown(texto_da_nota, dangerously_allow_html=False)
            ]
        )
    ], fluid=True, px="xl")

# ==== Layout Principal da Aplicação (com Navegação e Switch no Header) ====
app.layout = dmc.MantineProvider(
    id="mantine-provider",
    forceColorScheme="light", 
    children=[
        dcc.Location(id='url', refresh=False),
        dmc.Paper(
            shadow="xs", p="sm", withBorder=True, style={"marginBottom": "20px"},
            children=[
                dmc.Group(
                    justify="space-between", align="center",
                    children=[
                        dmc.Group(
                            gap="xl",
                            children=[
                                dmc.Anchor(dmc.Text("📊 Matriz GE", fw=500, size="lg"), href="/"),
                                dmc.Anchor(dmc.Text("📄 Nota Técnica", fw=500, size="lg"), href="/nota-tecnica"),
                            ]
                        ),
                        dmc.Switch(
                            id="dark-mode-switch", size="md", checked=False,
                            offLabel=DashIconify(icon="radix-icons:sun", width=18),
                            onLabel=DashIconify(icon="radix-icons:moon", width=18)
                        ),
                    ]
                )
            ]
        ),
        html.Div(id='page-content', style={"padding": "0 20px 20px 20px"})
    ]
)

# ==== Callback para Atualizar o Conteúdo da Página com Base na URL ====
@app.callback(Output('page-content', 'children'),
              [Input('url', 'pathname')])
def display_page(pathname):
    if pathname == '/nota-tecnica':
        return create_layout_nota_tecnica()
    elif pathname == '/': 
        return create_layout_matriz_ge()
    else:
        return dmc.Center(dmc.Text("Página não encontrada (404)", size="xl", c="red"), style={"height": "50vh"})

# ==== Callback para Alternar Modo Escuro ====
@app.callback(
    Output('mantine-provider', 'forceColorScheme'),
    Input('dark-mode-switch', 'checked'),
    prevent_initial_call=True
)
def toggle_dark_mode(checked):
    return 'dark' if checked else 'light'


# ==== Callback para Atualizar Filtros e Gráfico ====
@app.callback(
    [Output('filtro-area', 'children', allow_duplicate=True), Output('filtro-quadrante', 'children', allow_duplicate=True), Output('filtro-produto', 'children', allow_duplicate=True),
     Output('grafico-matriz', 'figure', allow_duplicate=True),
     Output('filtro-area', 'value', allow_duplicate=True), Output('filtro-quadrante', 'value', allow_duplicate=True), Output('filtro-produto', 'value', allow_duplicate=True),
     Output('botao-popover-area', 'children', allow_duplicate=True), Output('botao-popover-quadrante', 'children', allow_duplicate=True), Output('botao-popover-produto', 'children', allow_duplicate=True)],
    [Input('filtro-area', 'value'),
     Input('filtro-quadrante', 'value'),
     Input('filtro-produto', 'value'),
     Input('exibir-texto', 'checked'),
     Input('botao-reset', 'n_clicks'),
     Input('url', 'pathname'),
     Input('dark-mode-switch', 'checked')],
    prevent_initial_call='initial_duplicate'
)
def atualizar_tudo(selected_areas, selected_quadrantes, selected_produtos,
                   exibir_texto, reset_n_clicks, pathname,
                   dark_mode_checked):

    if pathname != '/':
        empty_fig = go.Figure()
        axis_font_color_empty_fig = 'white' if dark_mode_checked else '#333'
        empty_fig.update_layout(
            xaxis_visible=False, yaxis_visible=False,
            annotations=[dict(text="Navegue para a página da Matriz GE para visualizar os dados.", 
                              xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, 
                              font=dict(size=16, color=axis_font_color_empty_fig))],
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        texto_botao_area = html.Div([html.Div("Selecionar área", style={"flexGrow": 1, "whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "left", "paddingRight": "8px"}), html.Div("▼", style={"flexShrink": 0, "minWidth": "16px", "textAlign": "right"})], style={"width": "100%", "display": "flex", "alignItems": "center"})
        texto_botao_quadrante = html.Div([html.Div("Selecionar quadrante", style={"flexGrow": 1, "whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "left", "paddingRight": "8px"}), html.Div("▼", style={"flexShrink": 0, "minWidth": "16px", "textAlign": "right"})], style={"width": "100%", "display": "flex", "alignItems": "center"})
        texto_botao_produto = html.Div([html.Div("Selecionar produto", style={"flexGrow": 1, "whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "left", "paddingRight": "8px"}), html.Div("▼", style={"flexShrink": 0, "minWidth": "16px", "textAlign": "right"})], style={"width": "100%", "display": "flex", "alignItems": "center"})
        return ([], [], [], empty_fig, [], [], [], texto_botao_area, texto_botao_quadrante, texto_botao_produto)


    triggered_input_obj = callback_context.triggered[0] if callback_context.triggered else None
    triggered_input = triggered_input_obj['prop_id'].split('.')[0] if triggered_input_obj else None

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

    def get_button_children(selected_vals, default_single_text):
        base_style_text = {"flexGrow": 1, "whiteSpace": "nowrap", "overflow": "hidden", "textOverflow": "ellipsis", "textAlign": "left", "paddingRight": "8px"}
        base_style_arrow = {"flexShrink": 0, "minWidth": "16px", "textAlign": "right"}
        container_style = {"width": "100%", "display": "flex", "alignItems": "center"}

        if not selected_vals:
            text_content = f"Selecionar {default_single_text.lower()}"
        elif len(selected_vals) == 1:
            txt = str(selected_vals[0])
            max_l = 25 
            text_content = txt if len(txt) <= max_l else txt[:max_l-3] + "..."
        else:
            plural_suffix = "s"
            if default_single_text.lower() == "área": plural_suffix = "s selecionadas"
            elif default_single_text.lower() == "quadrante": plural_suffix = "s selecionados"
            elif default_single_text.lower() == "produto": plural_suffix = "s selecionados"
            else: plural_suffix = f"s selecionad{'as' if default_single_text.endswith('a') else 'os'}"
            text_content = f"{len(selected_vals)} {default_single_text.lower().replace('ã','a')}{plural_suffix}"
        return html.Div([ html.Div(text_content, style=base_style_text), html.Div("▼", style=base_style_arrow)], style=container_style)

    texto_botao_area_children = get_button_children(areas_val, "Área")
    texto_botao_quadrante_children = get_button_children(quadrantes_val, "Quadrante")
    texto_botao_produto_children = get_button_children(produtos_val, "Produto")

    fig = go.Figure()
    cols_for_plot = ["Hora Aluno", "Quadrante", "Posição Competitiva", "Atratividade Mercado", "Produto"]
    for col_plot in cols_for_plot:
        if col_plot not in df_plot.columns:
            if col_plot == "Hora Aluno": df_plot[col_plot] = 1.0
            elif col_plot in ["Posição Competitiva", "Atratividade Mercado"]: df_plot[col_plot] = pd.NA
            else: df_plot[col_plot] = pd.Series(dtype='object')

    max_hora_aluno = df_plot["Hora Aluno"].max() if not (df_plot.empty or df_plot["Hora Aluno"].isnull().all() or "Hora Aluno" not in df_plot.columns) else 1.0
    if pd.isna(max_hora_aluno): max_hora_aluno = 1.0

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
                    pos = "middle center" # Padrão
                    if pd.notna(x_pv) and pd.notna(y_am):
                        # Lógica de posicionamento do texto para evitar sobreposição com a bolha
                        # Vertical: top, middle, bottom. Horizontal: left, center, right.
                        # O texto é posicionado em relação ao *centro* do marcador.
                        pos_y, pos_x = "middle", "center"

                        if y_am < 2.5: pos_y = "top"       # Bolha em baixo, texto em cima
                        elif y_am > 7.5: pos_y = "bottom"    # Bolha em cima, texto em baixo
                        
                        if x_pv < 2.5: pos_x = "right"     # Bolha à esquerda, texto à direita
                        elif x_pv > 7.5: pos_x = "left"      # Bolha à direita, texto à esquerda
                        
                        # Se estiver centralizado em um eixo, ajustar o outro para não ficar sobre a bolha
                        if pos_y == "middle" and pos_x == "center":
                            if x_pv > 5: pos_x = "left" # Default para direita do centro da bolha
                            else: pos_x = "right"      # Default para esquerda do centro da bolha
                        
                        pos = f"{pos_y} {pos_x}".strip() # Remove espaços extras se um for 'middle' ou 'center' por padrão

                    posicoes_texto.append(pos)
                df_q["posicao_texto_calculada"] = posicoes_texto
                df_q["marker_size_calculada"] = (df_q["Hora Aluno"] / fator_escala).apply(lambda s: max(s if pd.notna(s) else 0, min_bubble_size_pref))

                fig.add_trace(go.Scatter(
                    x=df_q["Posição Competitiva"], y=df_q["Atratividade Mercado"],
                    mode="markers+text" if exibir_texto else "markers", name=str(quadrante_unico if pd.notna(quadrante_unico) else "N/A"),
                    marker=dict(size=df_q["marker_size_calculada"], color=fixed_cor_bolha, opacity=max(0, min(1, (100 - fixed_transparencia_percent) / 100.0))),
                    text=df_q["Produto_Formatado"] if exibir_texto else None, textposition=df_q["posicao_texto_calculada"] if exibir_texto else None,
                    textfont=dict(
                        color='white' if dark_mode_checked else 'black',
                        size=14, family="Bahnschrift, Arial, sans-serif"
                    ),
                    customdata=df_q['Produto'],
                    hovertemplate="<b>%{customdata}</b><br><b>Posição Competitiva:</b> %{x:.1f}<br><b>Atratividade:</b> %{y:.1f}<extra></extra>"
                ))

    if encoded_image_fundo:
        fig.add_layout_image(dict(source="data:image/png;base64," + encoded_image_fundo, xref="x domain", yref="y domain", x=0, y=1, sizex=1, sizey=1, sizing="stretch", opacity=1, layer="below"))

    axis_font_color = 'white' if dark_mode_checked else '#333'
    axis_tick_color = '#777' if dark_mode_checked else 'grey'
    axis_line_color = '#555' if dark_mode_checked else 'lightgrey'

    fig.update_layout(
        autosize=True, 
        xaxis=dict(
            range=[0, 10], autorange=False, fixedrange=True,
            title=dict(text="<b>Posição Competitiva</b>", font=dict(size=17, color=axis_font_color)),
            tickfont=dict(size=14, color=axis_font_color),
            showgrid=False, zeroline=False, tickmode="linear", tick0=0, dtick=1, ticks="outside", ticklen=6, tickwidth=1,
            tickcolor=axis_tick_color, linecolor=axis_line_color
        ),
        yaxis=dict(
            range=[0, 10], autorange=False, fixedrange=True,
            title=dict(text="<b>Atratividade de Mercado</b>", font=dict(size=17, color=axis_font_color)),
            tickfont=dict(size=14, color=axis_font_color),
            scaleanchor="x", scaleratio=(IMAGE_ORIGINAL_HEIGHT / IMAGE_ORIGINAL_WIDTH) if IMAGE_ORIGINAL_WIDTH !=0 else 1,
            showgrid=False, zeroline=False, tickmode="linear", tick0=0, dtick=1, ticks="outside", ticklen=6, tickwidth=1,
            tickcolor=axis_tick_color, linecolor=axis_line_color
        ),
        plot_bgcolor='rgba(0,0,0,0)', 
        paper_bgcolor='rgba(0,0,0,0)', 
        showlegend=False,
        margin=dict(l=60, r=30, t=30, b=60),
        hovermode="closest"
    )
    if df_plot.empty and (bool(areas_val) or bool(quadrantes_val) or bool(produtos_val)):
        fig.update_layout(
            xaxis_visible=False, yaxis_visible=False,
            annotations=[dict(
                text="Nenhum curso para os filtros.", xref="paper", yref="paper", x=0.5, y=0.5,
                showarrow=False, font=dict(size=16, color=axis_font_color) 
            )]
        )
    return (children_areas, children_quadrantes, children_produtos, fig, areas_val, quadrantes_val, produtos_val, texto_botao_area_children, texto_botao_quadrante_children, texto_botao_produto_children)

# ==== Callback para Baixar Gráfico ====
@app.callback(
    Output("download-imagem", "data"),
    Input("botao-download", "n_clicks"),
    State("grafico-matriz", "figure"),
    prevent_initial_call=True
)
def baixar_grafico(n_clicks, current_figure_dict):
    if n_clicks and current_figure_dict and current_figure_dict.get('data'):
        fig_to_download = go.Figure(current_figure_dict)
        
        fig_layout = current_figure_dict.get('layout', {})
        export_width = IMAGE_ORIGINAL_WIDTH
        export_height = IMAGE_ORIGINAL_HEIGHT

        if not fig_to_download.data or not encoded_image_fundo:
            on_screen_width = fig_layout.get('width')
            on_screen_height = fig_layout.get('height')
            if not isinstance(on_screen_width, (int, float)) or on_screen_width <= 0: on_screen_width = 800
            if not isinstance(on_screen_height, (int, float)) or on_screen_height <= 0:
                on_screen_height = int(on_screen_width / TARGET_ASPECT_RATIO) if TARGET_ASPECT_RATIO != 0 else int(on_screen_width * (IMAGE_ORIGINAL_HEIGHT / (IMAGE_ORIGINAL_WIDTH if IMAGE_ORIGINAL_WIDTH != 0 else 1)))
            
            export_width = int(on_screen_width) if on_screen_width and on_screen_width > 0 else 800
            export_height = int(on_screen_height) if on_screen_height and on_screen_height > 0 else 600
            
        fig_to_download.update_layout(paper_bgcolor='white', plot_bgcolor='white')

        img_bytes = fig_to_download.to_image(
            format="png", width=export_width, height=export_height, scale=1
        )
        return dcc.send_bytes(img_bytes, "matriz_ge_app.png")
    return dash.no_update


if __name__ == '__main__':
    app.run(debug=True)
