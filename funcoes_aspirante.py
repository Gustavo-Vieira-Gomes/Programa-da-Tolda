import pandas as pd
from aspirante import Aspirante
import re

aspirante_nulo = Aspirante()

aspirante_nulo.numero_atual = 'XXXX'
aspirante_nulo.numero_interno_atual = 'XXXX'
aspirante_nulo.nome_guerra = 'Nome não encontrado'
aspirante_nulo.pelotao = 'NA'
aspirante_nulo.companhia = 'NA'

def cria_aspirantes(pben):
    aspirantes = []
    '''
    Adiciona no array aspirante todos os aspirantes como objetos, assim como suas propriedades
    '''
    for i in range(len(pben)):
        aspirante = Aspirante()

        aspirante.numero_interno_atual = pben['Número Interno Atual'][i]
        #Alterar datas
        aspirante.numero_interno_2023 = pben['Número Interno 2023'][i]
        aspirante.numero_interno_2022 = pben['Número Interno 2022'][i]
        aspirante.numero_interno_2021 = pben['Número Interno 2021'][i]
        aspirante.numero_interno_2020 = pben['Número Interno 2020'][i]
        aspirante.nome_guerra = pben['Nome de Guerra'][i]
        aspirante.nome_completo = pben['Nome Completo'][i]
        aspirante.companhia = pben['Companhia'][i]
        aspirante.pelotao = pben['Pelotão'][i]
        aspirante.alojamento = pben['Alojamento/Camarote'][i]
        aspirante.nip = pben['N.I.P.'][i]
        aspirante.id_militar = pben['Número da ID. Militar'][i]
        aspirante.quarto_habilitacao = pben['Quarto/Habilitação'][i]
        aspirante.telefone = pben['Telefone de Contato'][i]
        aspirante.celular = pben['Celular de Contato'][i]
        aspirante.data_nascimento = pben['Data de Nascimento'][i]
        aspirante.sangue = pben['Tipo Sanguíneo +Fator RH'][i]
        aspirante.equipe = pben['Equipe'][i]
        aspirante.email = pben['E-mail'][i]
        aspirante.religiao = pben['Religião'][i]
        aspirante.cidade = pben['Cidade'][i]
        aspirante.estado = pben['Estado'][i]
        aspirante.bairro = pben['Bairro'][i]
        aspirante.endereco = pben['Endereço'][i]
        aspirante.cep = pben['CEP'][i]
        aspirante.nome_pai = pben['Nome do Pai'][i]
        aspirante.profissao_pai = pben['Profissão do Pai'][i]
        aspirante.forca_militar_pai = pben['Caso o Pai Seja Militar- Força Armada/Força Auxiliar'][i]
        aspirante.cargo_militar_pai = pben['Caso o Pai Seja Militar- Posto ou Graduação'][i]
        aspirante.nome_mae = pben['Nome da Mãe'][i]
        aspirante.profissao_mae = pben['Profissão da Mãe'][i]
        aspirante.forca_militar_mae = pben['Caso a Mãe Seja Militar- Força Armada/Força Auxiliar'][i]
        aspirante.cargo_militar_mae = pben['Caso a Mãe Seja Militar- Posto ou Graduação'][i]
        aspirante.adicional = pben['Descrição'][i]

        aspirantes.append(aspirante)
    
    return aspirantes

def busca_aspirante(aspirantes, valor_buscado):
    '''Retorna o objeto do aspirante desejado'''
    padrao_numero = '([0-9]{4})'
    padrao_im = '(IM-[0-9]{3})'
    padrao_fn = '(FN-[0-9]{3})'
    valor_buscado = valor_buscado.upper()

    '''
    Verifica se o valor buscado bate com uma RE de número de aspirante
    Se não bater, considera que é o nome de um aspirante
    '''
    
    if re.search(padrao_numero,valor_buscado) or re.search(padrao_im,valor_buscado) or re.search(padrao_fn,valor_buscado):
        for aspirante in aspirantes:
            if str(aspirante.numero_interno_atual) == str(valor_buscado):
                return aspirante
        return aspirante_nulo
    else:
        for aspirante in aspirantes:
            if aspirante.nome_guerra.upper() == str(valor_buscado).upper():
                return aspirante
        return aspirante_nulo


def busca_licenca(numero_interno, dataframe):
    dataframe['Número Interno'] = dataframe['Número Interno'].astype('str')
    index = dataframe.index[dataframe['Número Interno'] == numero_interno].to_list()
    info_licenca = []
    try:
        index = index[0]
        info_licenca.append(dataframe.at[dataframe.index[index],'Situação'])
        info_licenca.append(dataframe.at[dataframe.index[index],'Última Alteração'])
    except:
        info_licenca = ['Não encontrado','Não encontrado']
    return info_licenca


def busca_chave(numero_chave, dataframe):
    numero_chave = str(numero_chave)
    dataframe['Numero da Chave'] = dataframe['Numero da Chave'].astype('str')
    index = dataframe.index[dataframe['Numero da Chave'] == numero_chave].to_list()
    info_chave = []
    try:
        index = index[0]
        info_chave.append(str(dataframe.at[dataframe.index[index],'Numero da Chave']))
        info_chave.append(str(dataframe.at[dataframe.index[index],'Nome da Chave']))
        info_chave.append(str(dataframe.at[dataframe.index[index],'Anterior']))
        info_chave.append(str(dataframe.at[dataframe.index[index],'Atual']))
        info_chave.append(str(dataframe.at[dataframe.index[index],'Última Alteração']))
    except:
        info_chave = ['Chave não encontrada','Chave não encontrada','Chave não encontrada','Chave não encontrada','Chave não encontrada']
    return info_chave
