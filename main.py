from flask import Flask, render_template, request,send_file
import json
import pandas as pd
import os
from datetime import datetime
import os
import json
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
from werkzeug.datastructures import ImmutableMultiDict
from acessoBD import *

uri = f"mongodb+srv://{usuario}:{senha}@{BD}.{complemento}.mongodb.net/?retryWrites=true&w=majority"
client = MongoClient(uri, server_api=ServerApi('1'))
db = client['produtosML']
collection = db.get_collection('produtosML')
arquivo_download = ''
quantidade = 0
quant_m = 0
categoria = []
dias = 0
dias_m = 0
palavra = ''
nome = ''
dados_filtrados_ref = None
total_produtos = ''
dados_filtrados = []    
todas_cat = ['Acessórios para Veículos', 'Agro', 'Alimentos e Bebidas', 'Animais', 'Antiguidades e Coleções', 'Arte, Papelaria e Armarinho', 'Bebês', 'Beleza e Cuidado Pessoal', 'Brinquedos e Hobbies', 'Calçados, Roupas e Bolsas', 'Carros, Motos e Outros', 'Casa, Móveis e Decoração', 'Celulares e Telefones', 'Construção', 'Câmeras e Acessórios', 'Eletrodomésticos', 'Eletrônicos, Áudio e Vídeo', 'Esportes e Fitness', 'Ferramentas', 'Festas e Lembrancinhas', 'Games', 'Imóveis', 'Indústria e Comércio', 'Informática', 'Instrumentos Musicais', 'Joias e Relógios', 'Livros, Revistas e Comics', 'Musica, Filmes, Seriados', 'Serviços', 'Saúde', 'Mais Categorias', 'Tendencias']

app = Flask(__name__)


def filtrar(quant_m,dias_m,dias,quantidade,palavra,categoria,nome):
    filtro = {'dias':{'$lte':int(dias),"$gte":int(dias_m)},
              'quantidade':{'$lte':int(quantidade),"$gte":int(quant_m)}
              }
    if palavra != '':
        filtro['palavra chave'] =  {"$regex": palavra,'$options':'i'}
    if categoria != []:
        if len(categoria) == 1:
            filtro['categoria'] = {"$regex": categoria[0] ,'$options':'i'}
        else:
            fil = '|'.join(categoria)
            filtro['categoria'] = {"$regex": fil ,'$options':'i'}
    if nome != '':
        filtro['nome'] = {"$regex": nome ,'$options':'i'}
    return filtro
    

def apagar_arquivo(extensao):
    pasta_atual = os.getcwd()  
    # Verifique cada arquivo na pasta atual
    for arquivo in os.listdir(pasta_atual):
        if arquivo.endswith(extensao):  # Verifica se o arquivo tem a extensão .xlsx
            caminho_completo = os.path.join(pasta_atual, arquivo)  # Obtém o caminho completo do arquivo
            os.remove(caminho_completo)  # Apaga o arquivo



# Rota para a página principal
@app.route('/')
def index():
    global dados_filtrados_ref
    global dias_m
    global quant_m
    global dias
    global quantidade
    global categoria
    apagar_arquivo('.xlsx')
    return render_template('index.html',
                           dados=[],
                           dias=dias,
                           dias_m = dias_m,
                           quantidade=quantidade,
                           quant_m=quant_m,
                           categoria=categoria,
                           palavra=palavra,
                           nome = nome,
                           total_produtos=total_produtos)#, dados=dados)

# Rota para a busca com filtros

@app.route('/buscar', methods=['POST'])
def buscar():

    global collection
    global arquivo_download
    global dias_m
    global dias
    global quant_m
    global quantidade
    global categoria
    global dados_filtrados
    global palavra
    global nome
    global dados_filtrados_ref

    apagar_arquivo('.xlsx')
    

    # Obter filtros dos parâmetros da solicitação POST
    categoria = []
    try:
        categoria = request.form.getlist('categs')
    except:
        pass
            
    
    quant_m =request.form['quant_m']
    dias_m = request.form['dias_m']
    dias = request.form['dias']
    quantidade = request.form['quantidade']
    nome = request.form['nome']
    #categoria = request.form['todas_cat']
    palavra = request.form['palavra']
    if nome == '':
        nome = ''
    if palavra == '':
        palavra = ''

    hoje = datetime.now().strftime("%d-%m-%Y")
    print(filtrar(quant_m,dias_m,dias,quantidade,palavra,categoria,nome))
    filtro = filtrar(quant_m,dias_m,dias,quantidade,palavra,categoria,nome)
    #calcular bytes:
    máx_bytes = 33554432
                 #1140568
    dados_filtrados = list(collection.find(filtro))
    
    
    if len(categoria) == 1:
        try:
            dados_filtrados.sort(key = lambda x: x['nome'])
        except:
            dados_filtrados.sort(key = lambda x: x['nome'])
            
    else:
        dados_filtrados.sort(key = lambda x: x['categoria'])
    dados_filtrados_ref= []
    urls_filtro = []
    
    for dados in dados_filtrados:
        if dados['url'] not in urls_filtro:
            dados_filtrados_ref.append(dados)
            urls_filtro.append(dados['url'])
    
    total_produtos = f'Total de Produtos: {len(dados_filtrados_ref)}'
    
    arquivo_download = f'nome_{nome}_cat_{"_".join(categoria)}_dias_{dias_m}_{dias}_vendas_{quant_m}_{quantidade}_palavra_chave_{palavra}_{hoje}'

    pd.DataFrame(dados_filtrados_ref).to_excel(f'{arquivo_download}.xlsx')
    
    return render_template('index.html',
                           dados=dados_filtrados_ref,
                           dias=dias,
                           dias_m = dias_m,
                           quantidade=quantidade,
                           quant_m=quant_m,
                           categoria=categoria,
                           palavra=palavra,
                           nome = nome,
                           baixar=f'{arquivo_download}.xlsx',
                           total_produtos=total_produtos)

@app.route('/download')
def download_arquivo():
   
    global arquivo_download
    
    caminho_arquivo = f'{arquivo_download}.xlsx'
    
    return send_file(caminho_arquivo, as_attachment=True)



if __name__ == '__main__':
    app.run(debug=True,use_reloader=False)

