"""
    Projeto: Controle Geral da Tolda
    Suporte: Asp Hartmann, Asp Lucas Moraes, Asp Nunes Trindade, Asp Wagner Souza
    Data da última modificação: 30/12/2023
    Resumo da última modificação: Adicionou o Nome da chave junto a sua pesquisa, retirada ou devolução, além de
correção de erros na página claviculário.

    Objetivos:
-3 botões pagian Clavi (OK)

-Salvamento autoamtico  -Pagina Scrollable "RegistroLicenças"   -Trocar Excel por BD
-Função e arquivo "HistoricoSemanalLicenças"    - **ETCS**
"""
import os, sys
from kivy.resources import resource_add_path, resource_find
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.properties import StringProperty, ObjectProperty, ListProperty
import pandas as pd
from funcoes_aspirante import *
from kivy.uix.boxlayout import BoxLayout
from kivy.core.window import Window
from datetime import datetime, timedelta
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from kivy.uix.dropdown import DropDown
from kivy.uix.button import Button
from kivy.uix.scrollview import ScrollView
from kivy.uix.recycleview import RecycleView
from kivy.uix.recycleview.views import RecycleDataViewBehavior

EXTERN_FILE      = 'registro.txt'
FILE_NAME        = 'Plano de Busca do Corpo de Aspirantes 2024.ods'
SHEET_CHEFE_DIA  = 'ChefeDia'
SHEET_LICENCAS   = 'Licenças'
SHEET_PBEN       = 'PBEN'
SHEET_CHAVES     = 'Chaves'
SHEET_PARTE_ALTA = 'ParteAlta'

Window.size = (1250, 650)

class MenuScreen(Screen):
    pass

class PbenScreen(Screen):
    chave_pesquisa = ObjectProperty(None)


class LicencaScreen(Screen):
    chave_pesquisa = ObjectProperty(None)


class RegistroLicencasScreen(Screen):
    pass

class ChavesScreen(Screen):
    pass


class ParteAltaScreen(Screen):
    pass


class ChefeDiaScreen(Screen):
    pass

class VerChefesDiaScreen(Screen):
    pass

class SuporteScreen(Screen):
    pass



class ControleGeralApp(App):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        global planilha
        self.pben = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_PBEN)

        self.aspirantes = cria_aspirantes(self.pben)

        self.licencas = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_LICENCAS)
        self.licencas['Última Alteração'] = self.licencas['Última Alteração'].astype('str')
        self.licencas['Número Interno']   = self.licencas['Número Interno'].astype('str')
        planilha = self.licencas.copy()

        self.popup_content = Label(text='Salvando...', color=(0.18, 0.28, 0.40, 1), bold=True)
        self.popup_salvando_licenca = Popup(title='Salvando', content=self.popup_content,
                                            size_hint=(None, None), size=(300, 300))

        self.organiza_controle_geral_licenca()
        self.organiza_primeiro_licenca()
        self.organiza_segundo_licenca()
        self.organiza_terceiro_licenca()
        self.organiza_quarto_licenca()

        self.chaves = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_CHAVES)
        self.chaves['Última Alteração'] = self.chaves['Última Alteração'].astype('str')
        self.chaves['Anterior'] = self.chaves['Anterior'].astype('str')
        self.chaves['Atual'] = self.chaves['Atual'].astype('str')
        self.organiza_claviculario()

        self.partealta = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_PARTE_ALTA)
        self.partealta['Última Alteração'] = self.partealta['Última Alteração'].astype('str')
        self.partealta['Número Interno'] = self.partealta['Número Interno'].astype('str')

        self.organiza_controle_geral_partealta()
        self.organiza_primeiro_partealta()
        self.organiza_segundo_partealta()
        self.organiza_terceiro_partealta()
        self.organiza_quarto_partealta()
        self.chefedia = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_CHEFE_DIA)

    """
        Ao iniciar o programa, isso deixará os campos de informação limpos
    """

    nome_guerra  = StringProperty()
    numero_atual = StringProperty()
    # Alterar datas
    numero_interno_2023 = StringProperty()
    numero_interno_2022 = StringProperty()
    numero_interno_2021 = StringProperty()
    numero_interno_2020 = StringProperty()
    nome_completo       = StringProperty()
    nascimento          = StringProperty()
    telefone            = StringProperty()
    celular             = StringProperty()
    email               = StringProperty()
    companhia           = StringProperty()
    pelotao             = StringProperty()
    camarote            = StringProperty()
    quarto              = StringProperty()
    nip                 = StringProperty()
    sangue              = StringProperty()

    # Licenca
    situacao_atual_licenca   = StringProperty()
    ultima_alteracao_licenca = StringProperty()
    texto_input_licenca      = StringProperty()
    licenca_salvou           = StringProperty('Sem alterações')
    # LicencaGeral
    abordo_licenca         = StringProperty('0')
    baixado_licenca        = StringProperty('0')
    crestricao_licenca     = StringProperty('0')
    dispdomiciliar_licenca = StringProperty('0')
    hnmd_licenca           = StringProperty('0')
    lts_licenca            = StringProperty('0')
    licenciados_licenca    = StringProperty('0')
    stgt_licenca           = StringProperty('0')

    # LicencaPrimeiroAno
    abordo_licenca1         = StringProperty('0')
    baixado_licenca1        = StringProperty('0')
    crestricao_licenca1     = StringProperty('0')
    dispdomiciliar_licenca1 = StringProperty('0')
    hnmd_licenca1           = StringProperty('0')
    lts_licenca1            = StringProperty('0')
    licenciados_licenca1    = StringProperty('0')
    stgt_licenca1           = StringProperty('0')
    resumo_licenca1         = ListProperty([])
    # LicencaSegundoAno
    abordo_licenca2         = StringProperty('0')
    baixado_licenca2        = StringProperty('0')
    crestricao_licenca2     = StringProperty('0')
    dispdomiciliar_licenca2 = StringProperty('0')
    hnmd_licenca2           = StringProperty('0')
    lts_licenca2            = StringProperty('0')
    licenciados_licenca2    = StringProperty('0')
    stgt_licenca2           = StringProperty('0')
    resumo_licenca2         = ListProperty([])
    # LicencaTerceiroAno
    abordo_licenca3         = StringProperty('0')
    baixado_licenca3        = StringProperty('0')
    crestricao_licenca3     = StringProperty('0')
    dispdomiciliar_licenca3 = StringProperty('0')
    hnmd_licenca3           = StringProperty('0')
    lts_licenca3            = StringProperty('0')
    licenciados_licenca3    = StringProperty('0')
    stgt_licenca3           = StringProperty('0')
    resumo_licenca3         = ListProperty([])
    # LicencaQuartoAno
    abordo_licenca4         = StringProperty('0')
    baixado_licenca4        = StringProperty('0')
    crestricao_licenca4     = StringProperty('0')
    dispdomiciliar_licenca4 = StringProperty('0')
    hnmd_licenca4           = StringProperty('0')
    lts_licenca4            = StringProperty('0')
    licenciados_licenca4    = StringProperty('0')
    stgt_licenca4           = StringProperty('0')
    resumo_licenca4         = ListProperty([])

    # Claviculário
    chaves_salvou           = StringProperty('Sem alterações')
    chave_input             = StringProperty() #Objeto de pesquisa
    chave_nome              = StringProperty() #Nuero e Nome da chave, aparece na tela
    chave_atualmente_com    = StringProperty()
    chave_anteriormente_com = StringProperty()
    chave_ultima_alteracao  = StringProperty()
    chaves_claviculario     = StringProperty()
    chaves_fora             = StringProperty()

    # Parte Baixa
    partealta_salvou           = StringProperty('Sem alterações')
    situacao_atual_partealta   = StringProperty()
    ultima_alteracao_partealta = StringProperty()
    partealta_input            = StringProperty()
    # ParteBaixa1Ano
    partealta_partelta1     = StringProperty()
    tfm_partealta1          = StringProperty()
    saladeestado_partealta1 = StringProperty()
    enfermaria_partealta1   = StringProperty()
    banco_partealta1        = StringProperty()
    biblioteca_partealta1   = StringProperty()
    # ParteBaixa2Ano
    partealta_partelta2     = StringProperty()
    tfm_partealta2          = StringProperty()
    saladeestado_partealta2 = StringProperty()
    enfermaria_partealta2   = StringProperty()
    banco_partealta2        = StringProperty()
    biblioteca_partealta2   = StringProperty()
    # ParteBaixa3Ano
    partealta_partelta3     = StringProperty()
    tfm_partealta3          = StringProperty()
    saladeestado_partealta3 = StringProperty()
    enfermaria_partealta3   = StringProperty()
    banco_partealta3        = StringProperty()
    biblioteca_partealta3   = StringProperty()
    # ParteBaixa4Ano
    partealta_partelta4     = StringProperty()
    tfm_partealta4          = StringProperty()
    saladeestado_partealta4 = StringProperty()
    enfermaria_partealta4   = StringProperty()
    banco_partealta4        = StringProperty()
    biblioteca_partealta4   = StringProperty()
    # ParteBaixaGeral
    partealta_partelta     = StringProperty()
    tfm_partealta          = StringProperty()
    saladeestado_partealta = StringProperty()
    enfermaria_partealta   = StringProperty()
    banco_partealta        = StringProperty()
    biblioteca_partealta   = StringProperty()

    # Chefe do Dia
    chefedodia_registrando1 = StringProperty()
    chefedodia_registrando2 = StringProperty()
    numero_ajosca_cd        = StringProperty()
    quarto_de_serviço_cd    = StringProperty()
    numero_chefe_cd         = StringProperty()
    companhia_cd            = StringProperty()
    pelotao_cd              = StringProperty()
    cintos_cd               = StringProperty()
    computador_cd           = StringProperty()
    mapamundi_cd            = StringProperty()
    bandeira_cd             = StringProperty()
    licenciados_cd          = StringProperty()
    regressos_cd            = StringProperty()
    

    def build(self):
        # Create the screen manager
        self.sm = ScreenManager()
        self.sm.add_widget(MenuScreen(name='menu'))
        self.sm.add_widget(PbenScreen(name='pben'))
        self.sm.add_widget(LicencaScreen(name='licenca'))
        self.sm.add_widget(RegistroLicencasScreen(name='registrolicencas'))
        self.sm.add_widget(ChavesScreen(name='chaves'))
        self.sm.add_widget(ParteAltaScreen(name='partealta'))
        self.sm.add_widget(ChefeDiaScreen(name='chefedia'))
        self.sm.add_widget(VerChefesDiaScreen(name='verchefesdia'))
        self.sm.add_widget(SuporteScreen(name='suporte'))

        return self.sm

    def consulta_pben(self, chave_pesquisa):
        aspirante = busca_aspirante(self.aspirantes, chave_pesquisa)
        self.numero_atual = str(aspirante.numero_interno_atual)
        self.nome_guerra  = str(aspirante.nome_guerra)
        # Alterar datas
        self.numero_interno_2023 = str(aspirante.numero_interno_2023)
        self.numero_interno_2022 = str(aspirante.numero_interno_2022)
        self.numero_interno_2021 = str(aspirante.numero_interno_2021)
        self.numero_interno_2020 = str(aspirante.numero_interno_2020)
        self.nascimento = str(aspirante.data_nascimento)
        self.celular    = str(aspirante.celular)
        self.telefone   = str(aspirante.telefone)
        self.email      = str(aspirante.email)
        self.companhia  = str(aspirante.companhia)
        self.pelotao    = str(aspirante.pelotao)
        self.camarote   = str(aspirante.alojamento)
        self.quarto     = str(aspirante.quarto_habilitacao)
        self.nip        = str(aspirante.nip)
        self.sangue     = str(aspirante.sangue)
        self.nome_completo = str(aspirante.nome_completo)

    BOTAO_PRESSIONADO = StringProperty('')

    def consultar_licenca(self, chave_pesquisa):
        self.consulta_pben(chave_pesquisa)

        info_licencas = busca_licenca(self.numero_atual, self.licencas)
        self.situacao_atual_licenca = info_licencas[0]
        self.ultima_alteracao_licenca = info_licencas[1]
        print("Botão pressionado: " + self.BOTAO_PRESSIONADO)

    def registro_externo_regs_lics(self, situacao, horario):
        arq = open(EXTERN_FILE, 'a')
        arq.write(situacao+' '+str(horario.year)+' '+str(horario.month)+' '+str(horario.day)+' '+str(horario.hour)+' '+str(horario.minute)+'\n')

        arq.close()

    def atualiza_licenca(self, button_text):
        print("Botão pressionado: " + button_text)

        if button_text == "":
            pass
        else:
            try:
                index = self.licencas.query('`Número Interno` == @self.numero_atual').index.tolist()[0]
                if button_text == 'Regresso':
                    self.licencas['Situação'][index] = 'A bordo'
                    self.registro_externo_regs_lics("REG", datetime.now())

                else:
                    self.licencas['Situação'][index] = button_text
                    if button_text == "Licença":
                        self.registro_externo_regs_lics("LIC", datetime.now())


                self.licencas['Última Alteração'][index] = datetime.now().strftime('%d/%m/%Y %H:%M')

                self.licenca_salvou = 'Alterações pendentes'

            except:
                self.numero_atual = 'Selecione um aspirante'
                self.nome_guerra = ''

        info_licencas = busca_licenca(self.numero_atual, self.licencas)
        self.situacao_atual_licenca = info_licencas[0]
        self.ultima_alteracao_licenca = info_licencas[1]

        self.organiza_controle_geral_licenca()
        self.organiza_primeiro_licenca()
        self.organiza_segundo_licenca()
        self.organiza_terceiro_licenca()
        self.organiza_quarto_licenca()

    def salvar_alteracoes(self):

        with pd.ExcelWriter(resource_path(FILE_NAME)) as writer:
            self.licencas.to_excel(  writer, sheet_name=SHEET_LICENCAS  , index=False)
            self.pben.to_excel(      writer, sheet_name=SHEET_PBEN      , index=False)
            self.chaves.to_excel(    writer, sheet_name=SHEET_CHAVES    , index=False)
            self.partealta.to_excel( writer, sheet_name=SHEET_PARTE_ALTA, index=False)
            self.chefedia.to_excel(  writer, sheet_name=SHEET_CHEFE_DIA , index=False)

        self.licencas = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_LICENCAS  )
        self.chaves = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_CHAVES    )
        self.partealta = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_PARTE_ALTA)
        self.chefedia = pd.read_excel( resource_path(FILE_NAME), sheet_name=SHEET_CHEFE_DIA )

        self.licencas['Última Alteração']  = self.licencas['Última Alteração'].astype('str')
        self.licencas['Número Interno']    = self.licencas['Número Interno'].astype('str')
        self.partealta['Última Alteração'] = self.partealta['Última Alteração'].astype('str')
        self.partealta['Número Interno']   = self.partealta['Número Interno'].astype('str')

        self.chaves_salvou    = 'Sem alterações'
        self.licenca_salvou   = 'Sem alterações'
        self.partealta_salvou = 'Sem alterações'

        self.popup_salvando_licenca.dismiss()

    def organiza_controle_geral_licenca(self):

        dicionario_licencas_geral = dict(self.licencas['Situação'].value_counts())

        try:
            self.abordo_licenca = str(dicionario_licencas_geral['A bordo'])
        except:
            self.abordo_licenca = '0'

        try:
            self.baixado_licenca = str(dicionario_licencas_geral['Baixado'])
        except:
            self.baixado_licenca = '0'

        try:
            self.crestricao_licenca = str(dicionario_licencas_geral['C/ Restrição'])
        except:
            self.crestricao_licenca = '0'

        try:
            self.dispdomiciliar_licenca = str(dicionario_licencas_geral['Disp. Domiciliar'])
        except:
            self.dispdomiciliar_licenca = '0'

        try:
            self.hnmd_licenca = str(dicionario_licencas_geral['HNMD'])
        except:
            self.hnmd_licenca = '0'

        try:
            self.lts_licenca = str(dicionario_licencas_geral['LTS'])
        except:
            self.lts_licenca = '0'

        try:
            self.licenciados_licenca = str(dicionario_licencas_geral['Licença'])
        except:
            self.licenciados_licenca = '0'

        try:
            self.stgt_licenca = str(dicionario_licencas_geral['ST/GT'])
        except:
            self.stgt_licenca = '0'

    def organiza_primeiro_licenca(self):
        query = self.licencas.query('`Ano` == 1')
        dicionario_primeiro_licenca = dict(query['Situação'].value_counts())
        try:
            self.abordo_licenca1 = str(dicionario_primeiro_licenca['A bordo'])
        except:
            self.abordo_licenca1 = '0'
        try:
            self.baixado_licenca1 = str(dicionario_primeiro_licenca['Baixado'])
        except:
            self.baixado_licenca1 = '0'
        try:
            self.crestricao_licenca1 = str(dicionario_primeiro_licenca['C/ Restrição'])
        except:
            self.crestricao_licenca1 = '0'
        try:
            self.dispdomiciliar_licenca1 = str(dicionario_primeiro_licenca['Disp. Domiciliar'])
        except:
            self.dispdomiciliar_licenca1 = '0'
        try:
            self.hnmd_licenca1 = str(dicionario_primeiro_licenca['HNMD'])
        except:
            self.hnmd_licenca1 = '0'
        try:
            self.lts_licenca1 = str(dicionario_primeiro_licenca['LTS'])
        except:
            self.lts_licenca1 = '0'
        try:
            self.licenciados_licenca1 = str(dicionario_primeiro_licenca['Licença'])
        except:
            self.licenciados_licenca1 = '0'
        try:
            self.stgt_licenca1 = str(dicionario_primeiro_licenca['ST/GT'])
        except:
            self.stgt_licenca1 = '0'

    def organiza_segundo_licenca(self):
        query = self.licencas.query('`Ano` == 2')
        dicionario_segundo_licenca = dict(query['Situação'].value_counts())
        try:
            self.abordo_licenca2 = str(dicionario_segundo_licenca['A bordo'])
        except:
            self.abordo_licenca2 = '0'
        try:
            self.baixado_licenca2 = str(dicionario_segundo_licenca['Baixado'])
        except:
            self.baixado_licenca2 = '0'
        try:
            self.crestricao_licenca2 = str(dicionario_segundo_licenca['C/ Restrição'])
        except:
            self.crestricao_licenca2 = '0'
        try:
            self.dispdomiciliar_licenca2 = str(dicionario_segundo_licenca['Disp. Domiciliar'])
        except:
            self.dispdomiciliar_licenca2 = '0'
        try:
            self.hnmd_licenca2 = str(dicionario_segundo_licenca['HNMD'])
        except:
            self.hnmd_licenca2 = '0'
        try:
            self.lts_licenca2 = str(dicionario_segundo_licenca['LTS'])
        except:
            self.lts_licenca2 = '0'
        try:
            self.licenciados_licenca2 = str(dicionario_segundo_licenca['Licença'])
        except:
            self.licenciados_licenca2 = '0'
        try:
            self.stgt_licenca2 = str(dicionario_segundo_licenca['ST/GT'])
        except:
            self.stgt_licenca2 = '0'

    def organiza_terceiro_licenca(self):
        query = self.licencas.query('`Ano` == 3')
        dicionario_terceiro_licenca = dict(query['Situação'].value_counts())
        try:
            self.abordo_licenca3 = str(dicionario_terceiro_licenca['A bordo'])
        except:
            self.abordo_licenca3 = '0'
        try:
            self.baixado_licenca3 = str(dicionario_terceiro_licenca['Baixado'])
        except:
            self.baixado_licenca3 = '0'
        try:
            self.crestricao_licenca3 = str(dicionario_terceiro_licenca['C/ Restrição'])
        except:
            self.crestricao_licenca3 = '0'
        try:
            self.dispdomiciliar_licenca3 = str(dicionario_terceiro_licenca['Disp. Domiciliar'])
        except:
            self.dispdomiciliar_licenca3 = '0'
        try:
            self.hnmd_licenca3 = str(dicionario_terceiro_licenca['HNMD'])
        except:
            self.hnmd_licenca3 = '0'
        try:
            self.lts_licenca3 = str(dicionario_terceiro_licenca['LTS'])
        except:
            self.lts_licenca3 = '0'
        try:
            self.licenciados_licenca3 = str(dicionario_terceiro_licenca['Licença'])
        except:
            self.licenciados_licenca3 = '0'
        try:
            self.stgt_licenca3 = str(dicionario_terceiro_licenca['ST/GT'])
        except:
            self.stgt_licenca3 = '0'

    def organiza_quarto_licenca(self):
        query = self.licencas.query('`Ano` == 4')
        dicionario_quarto_licenca = dict(query['Situação'].value_counts())
        try:
            self.abordo_licenca4 = str(dicionario_quarto_licenca['A bordo'])
        except:
            self.abordo_licenca4 = '0'
        try:
            self.baixado_licenca4 = str(dicionario_quarto_licenca['Baixado'])
        except:
            self.baixado_licenca4 = '0'
        try:
            self.crestricao_licenca4 = str(dicionario_quarto_licenca['C/ Restrição'])
        except:
            self.crestricao_licenca4 = '0'
        try:
            self.dispdomiciliar_licenca4 = str(dicionario_quarto_licenca['Disp. Domiciliar'])
        except:
            self.dispdomiciliar_licenca4 = '0'
        try:
            self.hnmd_licenca4 = str(dicionario_quarto_licenca['HNMD'])
        except:
            self.hnmd_licenca4 = '0'
        try:
            self.lts_licenca4 = str(dicionario_quarto_licenca['LTS'])
        except:
            self.lts_licenca4 = '0'
        try:
            self.licenciados_licenca4 = str(dicionario_quarto_licenca['Licença'])
        except:
            self.licenciados_licenca4 = '0'
        try:
            self.stgt_licenca4 = str(dicionario_quarto_licenca['ST/GT'])
        except:
            self.stgt_licenca4 = '0'

    def atualiza_resumo_licenciamentos(self):
        self.licencas = pd.read_excel(resource_path(FILE_NAME), sheet_name=SHEET_LICENCAS)
        self.licencas['Última Alteração'] = self.licencas['Última Alteração'].astype('str')
        self.licencas['Número Interno']   = self.licencas['Número Interno'].astype('str')
        resumo1, resumo2, resumo3, resumo4 = [], [], [], []
        df_final = self.licencas[self.licencas['Número Interno'].apply(lambda x: x[0]) == '1'][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            resumo1.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        df_final = self.licencas[self.licencas['Número Interno'].apply(lambda x: x[0]) == '2'][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            resumo2.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        df_final = self.licencas[self.licencas['Número Interno'].apply(lambda x: x[0]) == '3'][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            resumo3.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        df_final = self.licencas[(self.licencas['Número Interno'].apply(lambda x: x[0]) == '4') | (self.licencas['Número Interno'].apply(lambda x: x[0]) == 'F') | (self.licencas['Número Interno'].apply(lambda x: x[0]) == 'I')][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            resumo4.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        
        #self.resumo_licenca1, self.resumo_licenca2, self.resumo_licenca3, self.resumo_licenca4 = resumo1.copy(), resumo2.copy(), resumo3.copy(), resumo4.copy()
        registro_licencas = self.sm.get_screen('registrolicencas')
        scroller1_instance = registro_licencas.ids.scroller1
        scroller2_instance = registro_licencas.ids.scroller2
        scroller3_instance = registro_licencas.ids.scroller3
        scroller4_instance = registro_licencas.ids.scroller4
        scroller1_instance.atualizar(resumo1)
        scroller2_instance.atualizar(resumo2)
        scroller3_instance.atualizar(resumo3)
        scroller4_instance.atualizar(resumo4)
        

    def input_texto_chave(self, chave_pesquisa):

        self.chave_input = chave_pesquisa

        info_chaves = busca_chave(self.chave_input, self.chaves)
        self.chave_input = info_chaves[0]

        if info_chaves[0] == 'Chave não encontrada':
            self.chave_nome = 'Chave não encontrada'
        else:
            self.chave_nome = info_chaves[0] + ' - ' + info_chaves[1]
        self.chave_anteriormente_com = info_chaves[2]
        self.chave_atualmente_com = info_chaves[3]
        self.chave_ultima_alteracao = info_chaves[4]

    def dar_chave(self, chave_pesquisa, asp_retirando):
        try:
            self.chave_input = chave_pesquisa

            busca_chave(self.chave_input, self.chaves)
            self.chaves_salvou = 'Alterações pendentes'

            index = self.chaves[self.chaves['Numero da Chave'] == self.chave_input].index.tolist()[0]
            self.chaves['Anterior'][index] = self.chaves['Atual'][index]
            self.chaves['Atual'][index] = asp_retirando
            self.chaves['Última Alteração'][index] = datetime.now().strftime('%d/%m/%Y %H:%M')

            info_chaves = busca_chave(self.chave_input, self.chaves)

            if info_chaves[0] == 'Chave não encontrada':
                self.chave_nome = 'Chave não encontrada'
            else:
                self.chave_nome = info_chaves[0] + ' - ' + info_chaves[1]
            self.chave_anteriormente_com = info_chaves[2]
            self.chave_atualmente_com = info_chaves[3]
            self.chave_ultima_alteracao = info_chaves[4]

        except:
            #self.chave_input = 'Chave não encontrada'
            self.chave_nome = 'Chave não encontrada'
            self.chave_atualmente_com = 'Chave não encontrada'
            self.chave_anteriormente_com = 'Chave não encontrada'
            self.chave_ultima_alteracao = 'Chave não encontrada'

        self.organiza_claviculario()

    def retorna_chave(self, Chave_Pesquisada):
        try:
            self.chave_input = Chave_Pesquisada.text

            busca_chave(self.chave_input, self.chaves)
            self.chaves_salvou = 'Alterações pendentes'

            index = self.chaves[self.chaves['Numero da Chave'] == self.chave_input].index.tolist()[0]
            self.chaves['Anterior'][index] = self.chaves['Atual'][index]
            self.chaves['Atual'][index] = 'Claviculário'
            self.chaves['Última Alteração'][index] = datetime.now().strftime('%d/%m/%Y %H:%M')

            #O 'Info_Chaves' deve estar logo apos todas as alterações/ logo antes de ser escrito para evitar erros
            info_chaves = busca_chave(self.chave_input, self.chaves)

            if info_chaves[0] == 'Chave não encontrada':
                self.chave_nome = 'Chave não encontrada'
            else:
                self.chave_nome = info_chaves[0] + ' - ' + info_chaves[1]
            self.chave_atualmente_com = info_chaves[3]
            self.chave_anteriormente_com = info_chaves[2]
            self.chave_ultima_alteracao = info_chaves[4]

        except:
            #self.chave_input = 'Chave não encontrada'
            self.chave_nome = 'Chave não encontrada'
            self.chave_atualmente_com = 'Chave não encontrada'
            self.chave_anteriormente_com = 'Chave não encontrada'
            self.chave_ultima_alteracao = 'Chave não encontrada'

        self.organiza_claviculario()

    def organiza_claviculario(self):
        dicionario = dict(self.chaves.Atual.value_counts())
        self.chaves_claviculario = str(dicionario['Claviculário'])
        self.chaves_fora = str(len(self.chaves.Atual) - dicionario['Claviculário'])

    def consultar_partealta(self, chave_pesquisa):
        self.consulta_pben(chave_pesquisa)

        info_partealta = busca_licenca(self.numero_atual, self.partealta)
        self.situacao_atual_partealta = info_partealta[0]
        self.ultima_alteracao_partealta = info_partealta[1]

    botao_parte_alta = StringProperty('')



    def atualiza_partealta(self, button_text):
        if button_text == '':
            pass
        else:
            try:
                index = self.partealta.query('`Número Interno` == @self.numero_atual').index.tolist()[0]
                self.partealta['Situação'][index] = button_text
                self.partealta['Última Alteração'][index] = datetime.now().strftime('%d/%m/%Y %H:%M')
                self.partealta_salvou = 'Alterações pendentes'
            except:
                self.numero_atual = 'Selecione um aspirante'
                self.nome_guerra = ''

        info_partealta = busca_licenca(self.numero_atual, self.partealta)
        self.situacao_atual_partealta = info_partealta[0]
        self.ultima_alteracao_partealta = info_partealta[1]

        self.organiza_controle_geral_partealta()
        self.organiza_primeiro_partealta()
        self.organiza_segundo_partealta()
        self.organiza_terceiro_partealta()
        self.organiza_quarto_partealta()

    def organiza_controle_geral_partealta(self):
        dicionario_partealta_geral = dict(self.partealta['Situação'].value_counts())

        try:
            self.partealta_partelta = str(dicionario_partealta_geral['Parte Alta'])
        except:
            self.partealta_partelta = '0'
        try:
            self.tfm_partealta = str(dicionario_partealta_geral['TFM'])
        except:
            self.tfm_partealta = '0'
        try:
            self.saladeestado_partealta = str(dicionario_partealta_geral['Sala de Estado'])
        except:
            self.saladeestado_partealta = '0'
        try:
            self.enfermaria_partealta = str(dicionario_partealta_geral['Enfermaria'])
        except:
            self.enfermaria_partealta = '0'
        try:
            self.banco_partealta = str(dicionario_partealta_geral['Banco'])
        except:
            self.banco_partealta = '0'
        try:
            self.biblioteca_partealta = str(dicionario_partealta_geral['Biblioteca'])
        except:
            self.biblioteca_partealta = '0'

    def organiza_primeiro_partealta(self):
        query = self.partealta.query('`Ano` == 1')
        dicionario_primeiro_partealta = dict(query['Situação'].value_counts())
        try:
            self.partealta_partelta1 = str(dicionario_primeiro_partealta['Parte Alta'])
        except:
            self.partealta_partelta1 = '0'
        try:
            self.tfm_partealta1 = str(dicionario_primeiro_partealta['TFM'])
        except:
            self.tfm_partealta1 = '0'
        try:
            self.saladeestado_partealta1 = str(dicionario_primeiro_partealta['Sala de Estado'])
        except:
            self.saladeestado_partealta1 = '0'
        try:
            self.enfermaria_partealta1 = str(dicionario_primeiro_partealta['Enfermaria'])
        except:
            self.enfermaria_partealta1 = '0'
        try:
            self.banco_partealta1 = str(dicionario_primeiro_partealta['Banco'])
        except:
            self.banco_partealta1 = '0'
        try:
            self.biblioteca_partealta1 = str(dicionario_primeiro_partealta['Biblioteca'])
        except:
            self.biblioteca_partealta1 = '0'

    def organiza_segundo_partealta(self):
        query = self.partealta.query('`Ano` == 2')
        dicionario_segundo_partealta = dict(query['Situação'].value_counts())
        try:
            self.partealta_partelta2 = str(dicionario_segundo_partealta['Parte Alta'])
        except:
            self.partealta_partelta2 = '0'
        try:
            self.tfm_partealta2 = str(dicionario_segundo_partealta['TFM'])
        except:
            self.tfm_partealta2 = '0'
        try:
            self.saladeestado_partealta2 = str(dicionario_segundo_partealta['Sala de Estado'])
        except:
            self.saladeestado_partealta2 = '0'
        try:
            self.enfermaria_partealta2 = str(dicionario_segundo_partealta['Enfermaria'])
        except:
            self.enfermaria_partealta2 = '0'
        try:
            self.banco_partealta2 = str(dicionario_segundo_partealta['Banco'])
        except:
            self.banco_partealta2 = '0'
        try:
            self.biblioteca_partealta2 = str(dicionario_segundo_partealta['Biblioteca'])
        except:
            self.biblioteca_partealta2 = '0'

    def organiza_terceiro_partealta(self):
        query = self.partealta.query('`Ano` == 3')
        dicionario_terceiro_partealta = dict(query['Situação'].value_counts())
        try:
            self.partealta_partelta3 = str(dicionario_terceiro_partealta['Parte Alta'])
        except:
            self.partealta_partelta3 = '0'
        try:
            self.tfm_partealta3 = str(dicionario_terceiro_partealta['TFM'])
        except:
            self.tfm_partealta3 = '0'
        try:
            self.saladeestado_partealta3 = str(dicionario_terceiro_partealta['Sala de Estado'])
        except:
            self.saladeestado_partealta3 = '0'
        try:
            self.enfermaria_partealta3 = str(dicionario_terceiro_partealta['Enfermaria'])
        except:
            self.enfermaria_partealta3 = '0'
        try:
            self.banco_partealta3 = str(dicionario_terceiro_partealta['Banco'])
        except:
            self.banco_partealta3 = '0'
        try:
            self.biblioteca_partealta3 = str(dicionario_terceiro_partealta['Biblioteca'])
        except:
            self.biblioteca_partealta3 = '0'

    def organiza_quarto_partealta(self):
        query = self.partealta.query('`Ano` == 4')
        dicionario_quarto_partealta = dict(query['Situação'].value_counts())
        try:
            self.partealta_partelta4 = str(dicionario_quarto_partealta['Parte Alta'])
        except:
            self.partealta_partelta4 = '0'
        try:
            self.tfm_partealta4 = str(dicionario_quarto_partealta['TFM'])
        except:
            self.tfm_partealta4 = '0'
        try:
            self.saladeestado_partealta4 = str(dicionario_quarto_partealta['Sala de Estado'])
        except:
            self.saladeestado_partealta4 = '0'
        try:
            self.enfermaria_partealta4 = str(dicionario_quarto_partealta['Enfermaria'])
        except:
            self.enfermaria_partealta4 = '0'
        try:
            self.banco_partealta4 = str(dicionario_quarto_partealta['Banco'])
        except:
            self.banco_partealta4 = '0'
        try:
            self.biblioteca_partealta4 = str(dicionario_quarto_partealta['Biblioteca'])
        except:
            self.biblioteca_partealta4 = '0'

    def atualiza_chefedodia_registrando(self):
        self.chefedodia_registrando1 = (
                                           f'AjOSCA: {self.numero_ajosca_cd} | Quarto de Serviço: {self.quarto_de_serviço_cd} | Chefe de dia: {self.numero_chefe_cd} | '
                                       ) + (
                                           f'Cia.: {self.companhia_cd} | Pelotão: {self.pelotao_cd} | '
                                       )

        self.chefedodia_registrando2 = (
                                           f'Computador/Alarme: {self.computador_cd} | Mapa-Mundi: {self.mapamundi_cd} | Bandeira Asp. Nasc.: {self.bandeira_cd} | '
                                       ) + (
                                           f'Cintos: {self.cintos_cd} | Licenciados: {self.licenciados_cd} | Regressos: {self.regressos_cd}')

    def att_numero_ajosca_cd(self, textinput):
        self.numero_ajosca_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def att_quarto_de_serviço_cd(self, textinput):
        self.quarto_de_serviço_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def att_numero_chefe_cd(self, textinput):
        self.numero_chefe_cd = textinput.text

        aspirante = busca_aspirante(self.aspirantes, self.numero_chefe_cd)

        self.numero_chefe_cd = str(aspirante.numero_interno_atual)
        self.pelotao_cd = aspirante.pelotao
        self.companhia_cd = aspirante.companhia

        self.atualiza_chefedodia_registrando()

    def att_cintos_cd(self, textinput):
        self.cintos_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def att_regressos_cd(self, textinput):
        self.regressos_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def att_licenciados_cd(self, textinput):
        self.licenciados_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def att_computador_cd(self, textinput):
        self.computador_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def att_mapamundi_cd(self, textinput):
        self.mapamundi_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def att_bandeira_cd(self, textinput):
        self.bandeira_cd = textinput.text
        self.atualiza_chefedodia_registrando()

    def registra_chefe_dia(self):
        agora = datetime.now().strftime('%d/%m/%Y %H:%M')
        self.chefedia.loc[self.chefedia.index.max() + 1] = [agora,
                                                            self.numero_ajosca_cd,
                                                            self.numero_chefe_cd,
                                                            self.quarto_de_serviço_cd,
                                                            self.companhia_cd,
                                                            self.pelotao_cd,
                                                            self.cintos_cd,
                                                            self.bandeira_cd,
                                                            self.licenciados_cd,
                                                            self.regressos_cd,
                                                            self.computador_cd,
                                                            self.mapamundi_cd]

        self.salvar_alteracoes()
        self.popup_salvando_licenca.dismiss()

    def count_regs_lics(self,quarto_servico, lic, reg):
        """
            Padrão de escrita do quarto de serviço:
                    DDHHMMF MES-DDHHMMF MES
            , onde DD é o dia, HH é a hora, MM é o minuto, e MES é o mês
        """

        # cria um dicionário que associa a sigla do mês a um valor numérico inteiro, que é o número do mês no calendário
        calend = {"JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4,
                  "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
                  "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12}

        try:
            # separa o quarto de serviço em início e fim
            quarto = quarto_servico.text.split('-')

            # cria uma variável para o início do quarto
            data_inicio = datetime(datetime.now().year,
                                   calend[quarto[0][8:11]],
                                   int(quarto[0][0:2]),
                                   int(quarto[0][2:4]),
                                   int(quarto[0][4:6])
                                   )
            # cria uma variável para o término do quarto
            data_termino = datetime(datetime.now().year,
                                    calend[quarto[1][8:11]],
                                    int(quarto[1][0:2]),
                                    int(quarto[1][2:4]),
                                    int(quarto[1][4:6]),
                                    )

            # Zera os contadores para que seja feita a contagem
            licencas = 0
            regressos = 0

            # Estabelece momento mais antigo em que um registro que está no EXTERN_FILE pode ter sido gravado
            piso = datetime.now() - timedelta(days=3)

            """
                Abre o EXTERN_FILE e grava, na variável linhas, todo o conteúdo dele
            """
            with open(EXTERN_FILE, 'r') as file:
                linhas = file.readlines()

            """
                Abre o  arquivo no modo ESCRITA para que possam ser regravadas as linhas com registros mais modernos
            do que o momento estabelecido pela variável PISO. Além disso, é feita a contagem do número de licenças e
            de regressos dentro do quarto de serviço estipulado.
            """
            with open(EXTERN_FILE, 'w') as file:

                # Faz a análise dos registros linha por linha
                for linha_inteira in linhas:
                    linha = linha_inteira.split()

                    # transforma o conteúdo da linha em uma variável do tipo DATETIME
                    data_registro = datetime(int(linha[1]), int(linha[2]), int(linha[3]), int(linha[4]), int(linha[5]))

                    # se o registro é mais moderno que o momento PISO:
                    if data_registro > piso:

                        # se o registro está dentro do quarto de serviço em questão:
                        if data_inicio <= data_registro <= data_termino:

                            if linha[0] == "LIC":
                                licencas = licencas + 1
                            elif linha[0] == "REG":
                                regressos = regressos + 1
                        # regrava os registros desejados no EXTERN_FILE
                        file.write(linha_inteira)
            # fecha o EXTERN_FILE
            file.close()

            # atualiza o texto das caixas de texto de Licenciados e Regressos para os valores calculados
            lic.text = str(licencas)
            reg.text = str(regressos)

        except:
            # dado que ocorreu um erro na execução por conta do formato da entrada do quarto de serviço, informa isso ao usuário
            lic.text = ''
            reg.text = ''
            quarto_servico.text = 'FORMATO FORA DO PADRÃO'

class ScrollerPage1(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        data = []
        df_final = planilha[planilha['Número Interno'].apply(lambda x: x[0]) == '1'][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        self.data = data

    def atualizar(self, data):
        self.data = data
        self.refresh_from_data()

class ScrollerPage2(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        data = []
        df_final = planilha[planilha['Número Interno'].apply(lambda x: x[0]) == '2'][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        self.data = data

    def atualizar(self, data):
        self.data = data
        self.refresh_from_data()

class ScrollerPage3(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        data = []
        df_final = planilha[planilha['Número Interno'].apply(lambda x: x[0]) == '3'][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        self.data = data

    def atualizar(self, data):
        self.data = data
        self.refresh_from_data()

class ScrollerPage4(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        data = []
        df_final = planilha[(planilha['Número Interno'].apply(lambda x: x[0]) == '4') | (planilha['Número Interno'].apply(lambda x: x[0]) == 'F') | (planilha['Número Interno'].apply(lambda x: x[0]) == 'I')][['Número Interno', 'Nome de Guerra','Situação']]
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}'})
        self.data = data

    def atualizar(self, data):
        self.data = data
        self.refresh_from_data()


def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("")

    return os.path.join(base_path, relative_path)



if __name__ == '__main__':
    # se o arquivo não existir, cria-o. Dessa maneira, evita-se erros relacionados à abertura futura do EXTERN_FILE
    arquivo = open(EXTERN_FILE,'a')
    arquivo.close()

    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))
    ControleGeralApp().run()
