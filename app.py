import streamlit as st
import ezdxf
from shapely.geometry import LineString, Polygon
import os
import pandas as pd
import logging

# Configurar o logger para o ezdxf
logging.getLogger('ezdxf').setLevel(logging.ERROR)

def carregar_dxf(uploaded_file):
    # Carregar o arquivo DXF diretamente da memória
    with open(uploaded_file.name, 'wb') as f:
        f.write(uploaded_file.getvalue())
    return ezdxf.readfile(uploaded_file.name)

def calcular_metros_lineares_e_contar_circulos_por_layer(doc, layer_names_excluidos):
    metros_por_layer = {}
    contagem_circulos_por_layer = {}

    for entity in doc.modelspace():
        layer_name = entity.dxf.layer

        if layer_name in layer_names_excluidos:
            continue
        
        if entity.dxftype() == 'LINE':
            line = LineString([entity.dxf.start, entity.dxf.end])
            length_in_meters = line.length
            metros_por_layer[layer_name] = metros_por_layer.get(layer_name, 0) + length_in_meters

        elif entity.dxftype() in {'LWPOLYLINE', 'POLYLINE'}:
            points = entity.get_points('xy') if entity.dxftype() == 'LWPOLYLINE' else [point for point in entity.points()]
            line = LineString(points)
            length_in_meters = line.length
            metros_por_layer[layer_name] = metros_por_layer.get(layer_name, 0) + length_in_meters

        elif entity.dxftype() == 'CIRCLE':
            contagem_circulos_por_layer[layer_name] = contagem_circulos_por_layer.get(layer_name, 0) + 1

    return metros_por_layer, contagem_circulos_por_layer

def contar_lwpolylines(doc, layer_name):
    return sum(1 for entity in doc.modelspace() if entity.dxftype() == 'LWPOLYLINE' and entity.dxf.layer == layer_name)

def calcular_area_total_em_hectares_por_layers(doc, layer_names):
    areas_por_layer = {layer: 0.0 for layer in layer_names}

    for entity in doc.modelspace():
        layer_name = entity.dxf.layer
        if layer_name not in areas_por_layer:
            continue
        
        if entity.dxftype() in {'LWPOLYLINE', 'POLYLINE'}:
            points = entity.get_points('xy') if entity.dxftype() == 'LWPOLYLINE' else [point for point in entity.points()]
            polygon = Polygon(points)
            area_m2 = polygon.area
            areas_por_layer[layer_name] += area_m2

    return {layer: area / 10000.0 for layer, area in areas_por_layer.items()}

def calcular_area_diferenca(doc, layer_total, layers_subtracao):
    area_total_hectares = calcular_area_total_em_hectares_por_layers(doc, [layer_total]).get(layer_total, 0.0)
    area_subtracao_hectares = sum(calcular_area_total_em_hectares_por_layers(doc, layers_subtracao).values())
    return area_total_hectares - area_subtracao_hectares

def salvar_em_excel(metros_por_layer, contagem_circulos_por_layer, areas_totais_hectares, area_carreadores, arquivo_excel):
    df_metros = pd.DataFrame({
        'Layer': list(metros_por_layer.keys()),
        'Valor': [round(value, 2) for value in metros_por_layer.values()],
        'Unidade': 'metros lineares'
    })

    df_circulos = pd.DataFrame({
        'Layer': list(contagem_circulos_por_layer.keys()),
        'Valor': list(contagem_circulos_por_layer.values()),
        'Unidade': 'unidades'
    })

    df_areas = pd.DataFrame({
        'Layer': list(areas_totais_hectares.keys()),
        'Valor': [round(value, 2) for value in areas_totais_hectares.values()],
        'Unidade': 'ha'
    })

    df_carreadores = pd.DataFrame({
        'Layer': ['Carreador Estimado'],
        'Valor': [round(area_carreadores, 2)],
        'Unidade': 'ha'
    })

    contador_coroamento_postes = contar_lwpolylines(doc, '0 -Coroamento Postes')
    df_circulos_coroamento = pd.DataFrame({
        'Layer': ['0 - Coroamento Postes'],
        'Valor': [contador_coroamento_postes],
        'Unidade': 'unidades'
    })

    with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
        start_col = 0
        df_metros.to_excel(writer, sheet_name='Resumo', startrow=0, startcol=start_col, index=False)
        start_col += len(df_metros.columns) + 1
        df_combined1 = pd.concat([df_circulos, df_circulos_coroamento], ignore_index=True)
        df_combined1.to_excel(writer, sheet_name='Resumo', startrow=0, startcol=start_col, index=False)
        start_col += len(df_circulos.columns) + 1
        df_combined2 = pd.concat([df_areas, df_carreadores], ignore_index=True)
        df_combined2.to_excel(writer, sheet_name='Resumo', startrow=0, startcol=start_col, index=False)

st.title('Inteligência Artificial - Mapa de Sistematização')

pasta_dados = 'Dados'
os.makedirs(pasta_dados, exist_ok=True)

uploaded_file = st.file_uploader("Escolha um arquivo DXF", type=["dxf"])

if uploaded_file:
    doc = carregar_dxf(uploaded_file)
    dxf_file_path = os.path.join(pasta_dados, uploaded_file.name)

    layer_names_excluidos = ['827 - Perímetro Cadastro', 'Não Reforma', "818 - Linha Transmissão_Alta Tensão", "0 -Coroamento Postes",
                         "819 - Linha Distribuição_Rede Eletrica", "815 - Sem Cana", "TALHÕES", "ÁREAS", "Texto", "801 - Talhões", "0"]

    layer_names = ['827 - Perímetro Cadastro', 'Não Reforma', "818 - Linha Transmissão_Alta Tensão",
                   "819 - Linha Distribuição_Rede Eletrica", "815 - Sem Cana", "TALHÕES"]

    layer_total = '827 - Perímetro Cadastro'
    layers_subtracao = ['Não Reforma', "818 - Linha Transmissão_Alta Tensão",
                        "819 - Linha Distribuição_Rede Eletrica", "815 - Sem Cana", "TALHÕES"]
    
    if st.button("Calcular"):
        metros_por_layer, contagem_circulos_por_layer = calcular_metros_lineares_e_contar_circulos_por_layer(doc, layer_names_excluidos)
        areas_totais_hectares = calcular_area_total_em_hectares_por_layers(doc, layer_names)
        area_carreadores = calcular_area_diferenca(doc, layer_total, layers_subtracao)

        arquivo_excel = os.path.join(pasta_dados, "resultado.xlsx")
        salvar_em_excel(metros_por_layer, contagem_circulos_por_layer, areas_totais_hectares, area_carreadores, arquivo_excel)

        st.write("Planilha gerada com sucesso!")
        st.download_button(label="Baixar Planilha Excel", data=open(arquivo_excel, 'rb').read(), file_name="resultado.xlsx")