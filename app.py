import streamlit as st
import ezdxf
from shapely.geometry import LineString, Polygon
import os
import pandas as pd
import logging

# Configurar o logger para o ezdxf
logging.getLogger('ezdxf').setLevel(logging.ERROR)

def calcular_metros_lineares_e_contar_circulos_por_layer(dxf_file_path, layer_names_excluidos):
    # Lê o arquivo DXF a partir do caminho do arquivo
    doc = ezdxf.readfile(dxf_file_path)
    
    # Dicionários para armazenar metros lineares e contagem de círculos por layer
    metros_por_layer = {}
    contagem_circulos_por_layer = {}

    # Iterar sobre as entidades do modelo
    for entity in doc.modelspace():
        layer_name = entity.dxf.layer

        if entity.dxftype() == 'LINE':
            # Verifica se a layer não está na lista de exclusão
            if layer_name not in layer_names_excluidos:
                # Cria uma linha com os pontos de início e fim
                line = LineString([entity.dxf.start, entity.dxf.end])
                length_in_meters = line.length

                # Acumula o comprimento no layer correspondente
                if layer_name in metros_por_layer:
                    metros_por_layer[layer_name] += length_in_meters
                else:
                    metros_por_layer[layer_name] = length_in_meters

        elif entity.dxftype() in {'LWPOLYLINE', 'POLYLINE'}:
            # Verifica se a layer não está na lista de exclusão
            if layer_name not in layer_names_excluidos:
                # Extrai os pontos da polilinha
                if entity.dxftype() == 'LWPOLYLINE':
                    points = entity.get_points('xy')
                else:  # POLYLINE
                    points = [point for point in entity.points()]
                
                line = LineString(points)
                length_in_meters = line.length

                # Acumula o comprimento no layer correspondente
                if layer_name in metros_por_layer:
                    metros_por_layer[layer_name] += length_in_meters
                else:
                    metros_por_layer[layer_name] = length_in_meters

        elif entity.dxftype() == 'CIRCLE':
            # Verifica se a layer não está na lista de exclusão
            if layer_name not in layer_names_excluidos:
                # Conta o número de círculos por layer
                if layer_name in contagem_circulos_por_layer:
                    contagem_circulos_por_layer[layer_name] += 1
                else:
                    contagem_circulos_por_layer[layer_name] = 1

        else:
            continue  # Ignorar outras entidades

    return metros_por_layer, contagem_circulos_por_layer

def contar_lwpolylines(dxf_file_path, layer_name):
    # Lê o arquivo DXF a partir do caminho do arquivo
    doc = ezdxf.readfile(dxf_file_path)
    
    # Contador para a quantidade de LWPOLYLINE na layer especificada
    contador = 0

    # Iterar sobre as entidades do modelo
    for entity in doc.modelspace():
        if entity.dxftype() == 'LWPOLYLINE':
            if entity.dxf.layer == layer_name:
                contador += 1

    return contador

def calcular_area_total_em_hectares_por_layers(dxf_file_path, layer_names):
    # Lê o arquivo DXF a partir do caminho do arquivo
    doc = ezdxf.readfile(dxf_file_path)
    
    # Dicionário para armazenar a área total por layer
    areas_por_layer = {layer: 0.0 for layer in layer_names}

    # Iterar sobre as entidades do modelo
    for entity in doc.modelspace():
        layer_name = entity.dxf.layer
        if layer_name in areas_por_layer:
            if entity.dxftype() in {'LWPOLYLINE', 'POLYLINE'}:
                # Extrai os pontos da polilinha
                if entity.dxftype() == 'LWPOLYLINE':
                    points = entity.get_points('xy')
                else:  # POLYLINE
                    points = [point for point in entity.points()]
                
                # Criar um polígono com os pontos
                polygon = Polygon(points)
                area_m2 = polygon.area
                areas_por_layer[layer_name] += area_m2

    # Converter as áreas de metros quadrados para hectares
    areas_por_layer_hectares = {layer: area / 10000.0 for layer, area in areas_por_layer.items()}

    return areas_por_layer_hectares

def calcular_area_diferenca(dxf_file_path, layer_total, layers_subtracao):
    # Calcular a área total da layer principal
    areas_totais_hectares = calcular_area_total_em_hectares_por_layers(dxf_file_path, [layer_total])
    area_total_hectares = areas_totais_hectares.get(layer_total, 0.0)

    # Calcular a soma das áreas das layers a serem subtraídas
    areas_subtracao_hectares = calcular_area_total_em_hectares_por_layers(dxf_file_path, layers_subtracao)
    area_subtracao_hectares = sum(areas_subtracao_hectares.values())

    # Calcular a área dos carreadores
    area_carreadores_hectares = area_total_hectares - area_subtracao_hectares

    return area_carreadores_hectares

def salvar_em_excel(metros_por_layer, contagem_circulos_por_layer, areas_totais_hectares, area_carreadores, arquivo_excel):
    # Criar DataFrames individuais com a unidade de medida
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

    contador_coroamento_postes = contar_lwpolylines(dxf_file_path, '0 -Coroamento Postes')

    df_circulos_coroamento = pd.DataFrame({
        'Layer': ['0 - Coroamento Postes'],
        'Valor': [contador_coroamento_postes],
        'Unidade': 'unidades'
    })

    # Criar uma planilha com todas as tabelas na mesma aba
    with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
        start_col = 0

        # Escrever DataFrame de metros lineares
        df_metros.to_excel(writer, sheet_name='Resumo', startrow=0, startcol=start_col, index=False)
        start_col += len(df_metros.columns) + 1  # Deixar espaço entre tabelas

        # Escrever DataFrame de círculos
        df_combined1 = pd.concat([df_circulos, df_circulos_coroamento], ignore_index=True)
        df_combined1.to_excel(writer, sheet_name='Resumo', startrow=0, startcol=start_col, index=False)
        start_col += len(df_circulos.columns) + 1  # Deixar espaço entre tabelas

        # Concatenar os DataFrames de áreas e carreadores
        df_combined2 = pd.concat([df_areas, df_carreadores], ignore_index=True)
        df_combined2.to_excel(writer, sheet_name='Resumo', startrow=0, startcol=start_col, index=False)

st.title('Inteligência Artificial - Mapa de Sistematização')

# Pasta para salvar arquivos
pasta_dados = 'Dados'
os.makedirs(pasta_dados, exist_ok=True)

# Upload do arquivo DXF
uploaded_file = st.file_uploader("Escolha um arquivo DXF", type=["dxf"])

if uploaded_file:
    dxf_file_path = os.path.join(pasta_dados, uploaded_file.name)
    with open(dxf_file_path, 'wb') as f:
        f.write(uploaded_file.read())

    layer_names_excluidos = ['827 - Perímetro Cadastro', 'Não Reforma', "818 - Linha Transmissão_Alta Tensão", "0 -Coroamento Postes",
                         "819 - Linha Distribuição_Rede Eletrica", "815 - Sem Cana", "TALHÕES", "ÁREAS", "Texto", "801 - Talhões", "0"]

    layer_names = ['827 - Perímetro Cadastro', 'Não Reforma', "818 - Linha Transmissão_Alta Tensão",
                   "819 - Linha Distribuição_Rede Eletrica", "815 - Sem Cana", "TALHÕES"]

    layer_total = '827 - Perímetro Cadastro'
    layer_name = '0 -Coroamento Postes'
    layers_subtracao = ['Não Reforma', "818 - Linha Transmissão_Alta Tensão",
                        "819 - Linha Distribuição_Rede Eletrica", "815 - Sem Cana", "TALHÕES"]
    
    layer_name = '0 -Coroamento Postes'
    quantidade = contar_lwpolylines(dxf_file_path, layer_name)

    if st.button("Calcular"):
        metros_por_layer, contagem_circulos_por_layer = calcular_metros_lineares_e_contar_circulos_por_layer(dxf_file_path, layer_names_excluidos)
        areas_totais_hectares = calcular_area_total_em_hectares_por_layers(dxf_file_path, layer_names)
        area_carreadores = calcular_area_diferenca(dxf_file_path, layer_total, layers_subtracao)

        arquivo_excel = os.path.join(pasta_dados, "resultado.xlsx")
        salvar_em_excel(metros_por_layer, contagem_circulos_por_layer, areas_totais_hectares, area_carreadores, arquivo_excel)

        st.write("Planilha gerada com sucesso!")
        st.download_button(label="Baixar Planilha Excel", data=open(arquivo_excel, 'rb').read(), file_name="resultado.xlsx")