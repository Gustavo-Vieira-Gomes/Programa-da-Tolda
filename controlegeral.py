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
from kivy.resources import resource_add_path
from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.properties import StringProperty, ObjectProperty, ListProperty
from funcoes_aspirante import *
from kivy.core.window import Window
from datetime import datetime, timedelta
from kivy.uix.recycleview import RecycleView
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from banco_de_dados import DatabaseTolda
from sqlite3 import connect
import os
import time
from banco_de_dados import resource_path

EXTERN_FILE      = 'registro.txt'
SHEET_LICENCAS   = 'Licenças'
SHEET_PBEN       = 'PBEN'
SHEET_CHAVES     = 'Chaves'

Window.size = (1250, 650)

class MenuScreen(Screen):
    pass

class PbenScreen(Screen):
    chave_pesquisa = ObjectProperty(None)


class LicencaScreen(Screen):
    chave_pesquisa_lic = ObjectProperty(None)


class RegistroLicencasScreen(Screen):
    pass

class ChavesScreen(Screen):
    pass

class SuporteScreen(Screen):
    pass

class AdminScreen(Screen):
    pass

class IniciarScreen(Screen):
    pass

class FileChooserScreen(Screen):
    pass

class ControleGeralApp(App):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.pben = DatabaseTolda().pegar_table(SHEET_PBEN)

        self.aspirantes = cria_aspirantes(self.pben)

        self.organiza_controle_geral_licenca()
        self.organiza_primeiro_licenca()
        self.organiza_segundo_licenca()
        self.organiza_terceiro_licenca()
        self.chaves = DatabaseTolda().pegar_table(SHEET_CHAVES)
        self.organiza_claviculario()

        self.popup_errolog= Label(text='Usuário/Senha incorretos',color = (0.18, 0.28, 0.40, 1), bold= True)
        self.popup_erro_login = Popup(title ='ERRO!',
                                            content = self.popup_errolog,
                                            size_hint=(None, None), size=(300, 300))
        
        self.popup_successlog = Label(text='Operação Realizada\n com Sucesso',color = (0.18, 0.28, 0.40, 1), bold= True)
        self.popup_success_operation = Popup(title ='Sucesso!',
                                            content = self.popup_successlog,
                                            size_hint=(None, None), size=(300, 300))
        self.popup_operation_errorlog = Label(text='Erro Ao realizar a Operação,\n tente novamente.',color = (0.18, 0.28, 0.40, 1), bold= True)
        self.popup_operation_error = Popup(title ='ERRO!',
                                            content = self.popup_operation_errorlog,
                                            size_hint=(None, None), size=(300, 300))
        

    """
        Ao iniciar o programa, isso deixará os campos de informação limpos
    """

    nome_guerra  = StringProperty()
    numero_atual = StringProperty()
    tela_antiga  = StringProperty()
    # Alterar datas
    numero_interno_2023 = StringProperty()
    numero_interno_2022 = StringProperty()
    numero_interno_2021 = StringProperty()
    nome_completo       = StringProperty()
    nascimento          = StringProperty()
    telefone            = StringProperty()
    celular             = StringProperty()
    email               = StringProperty()
    companhia           = StringProperty()
    pelotao             = StringProperty()
    equipe              = StringProperty()
    tel_emergencia      = StringProperty()
    nip                 = StringProperty()
    sangue              = StringProperty()

    # Licenca
    situacao_atual_licenca   = StringProperty()
    ultima_alteracao_licenca = StringProperty()
    texto_input_licenca      = StringProperty()
    # LicencaGeral
    abordo_licenca            = StringProperty('0')
    baixado_licenca           = StringProperty('0')
    visitacao_licenca         = StringProperty('0')
    dispdomiciliar_licenca    = StringProperty('0')
    hnmd_licenca              = StringProperty('0')
    lts_licenca               = StringProperty('0')
    stgt_licenca              = StringProperty('0')
    uism_licenca              = StringProperty('0')
    licenciados_rio_licenca   = StringProperty('0')
    licenciados_angra_licenca = StringProperty('0')

    # LicencaPrimeiroAno
    abordo_licenca1            = StringProperty('0')
    baixado_licenca1           = StringProperty('0')
    visitacao_licenca1         = StringProperty('0')
    dispdomiciliar_licenca1    = StringProperty('0')
    hnmd_licenca1              = StringProperty('0')
    lts_licenca1               = StringProperty('0')
    stgt_licenca1              = StringProperty('0')
    uism_licenca1              = StringProperty('0')
    licenciados_rio_licenca1   = StringProperty('0')
    licenciados_angra_licenca1 = StringProperty('0')
    resumo_licenca1            = ListProperty([])
    # LicencaSegundoAno
    abordo_licenca2           = StringProperty('0')
    baixado_licenca2          = StringProperty('0')
    visitacao_licenca2        = StringProperty('0')
    dispdomiciliar_licenca2   = StringProperty('0')
    hnmd_licenca2             = StringProperty('0')
    lts_licenca2              = StringProperty('0')
    stgt_licenca2             = StringProperty('0')
    uism_licenca2             = StringProperty('0')
    licenciados_rio_licenca2  = StringProperty('0')
    licenciados_angra_licenca2= StringProperty('0')
    resumo_licenca2         = ListProperty([])
    # LicencaTerceiroAno
    abordo_licenca3           = StringProperty('0')
    baixado_licenca3          = StringProperty('0')
    visitacao_licenca3        = StringProperty('0')
    dispdomiciliar_licenca3   = StringProperty('0')
    hnmd_licenca3             = StringProperty('0')
    lts_licenca3              = StringProperty('0')
    stgt_licenca3             = StringProperty('0')
    uism_licenca3             = StringProperty('0')
    licenciados_rio_licenca3  = StringProperty('0')
    licenciados_angra_licenca3= StringProperty('0')
    resumo_licenca3         = ListProperty([])

    # Claviculário
    chave_input             = StringProperty() #Objeto de pesquisa
    chave_nome              = StringProperty() #Nuero e Nome da chave, aparece na tela
    chave_atualmente_com    = StringProperty()
    chave_anteriormente_com = StringProperty()
    chave_ultima_alteracao  = StringProperty()
    chaves_claviculario     = StringProperty()
    chaves_fora             = StringProperty()

    def build(self):
        # Create the screen manager
        self.sm = ScreenManager()
        self.sm.add_widget(IniciarScreen(name='inicioscreen'))
        self.sm.add_widget(AdminScreen(name='adminscreen'))
        self.sm.add_widget(FileChooserScreen(name='filechooserscreen'))
        self.sm.add_widget(MenuScreen(name='menu'))
        self.sm.add_widget(PbenScreen(name='pben'))
        self.sm.add_widget(LicencaScreen(name='licenca'))
        self.sm.add_widget(RegistroLicencasScreen(name='registrolicencas'))
        self.sm.add_widget(ChavesScreen(name='chaves'))
        self.sm.add_widget(SuporteScreen(name='suporte'))

        return self.sm

    def verifica_user(self, login, senha):
        if login=="tolda" and senha=="tolda":
            self.user_menu = "menu"
        elif login=="admin" and senha=="admin":
            self.user_menu = "adminscreen"
        else:
            self.user_menu = "inicioscreen"
            self.popup_erro_login.open()

    def erro_login(self):
        time.sleep(0.5)
        self.popup_erro_login.dismiss()

    def atualizar_bd(self, arquivo_excel):
        try:
            DatabaseTolda().criar_database(arquivo_excel[0])
            print(arquivo_excel[0])
        except:
            self.popup_operation_error.open()
        else:
            self.refresh_app()
            self.popup_success_operation.open()


    def excluir_registros(self):
        pass

    def consulta_pben(self, chave_pesquisa):
        aspirante = busca_aspirante(self.aspirantes, chave_pesquisa)
        self.numero_atual = str(aspirante.numero_interno_atual)
        self.nome_guerra  = str(aspirante.nome_guerra)
        # Alterar datas
        self.numero_interno_2023 = str(aspirante.numero_interno_2023)
        self.numero_interno_2022 = str(aspirante.numero_interno_2022)
        self.numero_interno_2021 = str(aspirante.numero_interno_2021)
        self.nascimento = str(aspirante.data_nascimento)
        self.celular    = str(aspirante.celular)
        self.telefone   = str(aspirante.telefone)
        self.email      = str(aspirante.email)
        self.companhia  = str(aspirante.companhia)
        self.pelotao    = str(aspirante.pelotao)
        self.equipe   = str(aspirante.equipe)
        self.tel_emergencia     = str(aspirante.tel_emergencia)
        self.nip        = str(aspirante.nip)
        self.sangue     = str(aspirante.sangue)
        self.nome_completo = str(aspirante.nome_completo)

    BOTAO_PRESSIONADO = StringProperty('')

    def consultar_licenca(self, chave_pesquisa):
        self.consulta_pben(chave_pesquisa)
        info_licencas = busca_licenca(self.numero_atual, DatabaseTolda().pegar_table(SHEET_LICENCAS))
        self.situacao_atual_licenca = info_licencas[0]
        self.ultima_alteracao_licenca = info_licencas[1]
        print("Botão pressionado: " + self.BOTAO_PRESSIONADO)

    def registro_externo_regs_lics(self, situacao, horario):
        arq = open(EXTERN_FILE, 'a')
        arq.write(situacao+' '+str(horario.year)+' '+str(horario.month)+' '+str(horario.day)+' '+str(horario.hour)+' '+str(horario.minute)+'\n')

        arq.close()

    def refresh_app(self):
        self.organiza_controle_geral_licenca()
        self.organiza_primeiro_licenca()
        self.organiza_segundo_licenca()
        self.organiza_terceiro_licenca()
        self.organiza_claviculario()
        self.sm.get_screen('registrolicencas').ids.scroller1.atualizar()
        self.sm.get_screen('registrolicencas').ids.scroller2.atualizar()
        self.sm.get_screen('registrolicencas').ids.scroller3.atualizar()

    def reiniciar_licenca_reg(self, situação):
        try:
            conn = connect(resource_path('database_tolda.db'))
            c = conn.cursor()
            sql = f"UPDATE Licenças SET [Situação] = '{situação}' WHERE  [Situação] <> 'BAIXA'"
            c.execute(sql)
            sql = f"UPDATE Licenças SET [Última Alteração] = '{datetime.now().strftime('%d/%m/%Y %H:%M')}' WHERE  [Situação] <> 'BAIXA'"
            c.execute(sql)
            conn.commit()
            c.close()
            self.refresh_app()
            self.popup_success_operation.open()
        except Exception as e:
            print(e)
            self.popup_operation_error.open()


    def atualiza_licenca(self, button_text):
        print("Botão pressionado: " + button_text)
        if button_text == "":
            pass
        else:
            if DatabaseTolda().pegar_info(SHEET_LICENCAS, 'Situação', 'Número Interno', self.numero_atual) != 'BAIXA':
                try:
                    if button_text == 'Regresso':
                        DatabaseTolda().update_info(SHEET_LICENCAS, 'Situação', 'A Bordo', 'Número Interno', self.numero_atual)
                        self.registro_externo_regs_lics("REG", datetime.now())

                    else:
                        DatabaseTolda().update_info(SHEET_LICENCAS, 'Situação', button_text, 'Número Interno', self.numero_atual)
                        if button_text == "Licença Angra":
                            self.registro_externo_regs_lics("LIC ANGRA", datetime.now())
                        elif button_text == 'Licença Rio':
                            self.registro_externo_regs_lics('LIC RIO', datetime.now())
                    DatabaseTolda().update_info(SHEET_LICENCAS, 'Última Alteração', datetime.now().strftime('%d/%m/%Y %H:%M'), 'Número Interno', self.numero_atual)

                except:
                    self.numero_atual = 'Selecione um aspirante'
                    self.nome_guerra = ''

        info_licencas = busca_licenca(self.numero_atual, DatabaseTolda().pegar_table(SHEET_LICENCAS))
        self.situacao_atual_licenca = info_licencas[0]
        self.ultima_alteracao_licenca = info_licencas[1]

        self.organiza_controle_geral_licenca()
        self.organiza_primeiro_licenca()
        self.organiza_segundo_licenca()
        self.organiza_terceiro_licenca()

    def organiza_controle_geral_licenca(self):
        try:
            self.abordo_licenca =  DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'A Bordo')
        except:
            self.abordo_licenca = '0'

        try:
            self.baixado_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Baixado')
        except:
            self.baixado_licenca = '0'

        try:
            self.visitacao_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Visitação')
        except:
            self.visitacao_licenca = '0'

        try:
            self.dispdomiciliar_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Disp. Domiciliar')
        except:
            self.dispdomiciliar_licenca = '0'

        try:
            self.hnmd_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'HNMD')
        except:
            self.hnmd_licenca = '0'

        try:
            self.lts_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'LTS')
        except:
            self.lts_licenca = '0'

        try:
            self.licenciados_angra_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Angra')
        except:
            self.licenciados_angra_licenca = '0'

        try:
            self.licenciados_rio_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Rio')
        except:
            self.licenciados_rio_licenca = '0'

        try:
            self.stgt_licenca = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'ST/GT')
        except:
            self.stgt_licenca = '0'

    def organiza_primeiro_licenca(self):
        try:
            self.abordo_licenca1 =  DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'A Bordo', 1)
        except:
            self.abordo_licenca1 = '0'

        try:
            self.baixado_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Baixado', 1)
        except:
            self.baixado_licenca1 = '0'

        try:
            self.visitacao_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Visitação', 1)
        except:
            self.visitacao_licenca1 = '0'

        try:
            self.dispdomiciliar_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Disp. Domiciliar', 1)
        except:
            self.dispdomiciliar_licenca1 = '0'

        try:
            self.hnmd_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'HNMD', 1)
        except:
            self.hnmd_licenca1 = '0'

        try:
            self.lts_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'LTS', 1)
        except:
            self.lts_licenca1 = '0'

        try:
            self.licenciados_angra_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Angra', 1)
        except:
            self.licenciados_angra_licenca1 = '0'

        try:
            self.licenciados_rio_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Rio', 1)
        except:
            self.licenciados_rio_licenca1 = '0'

        try:
            self.stgt_licenca1 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'ST/GT', 1)
        except:
            self.stgt_licenca1 = '0'

    def organiza_segundo_licenca(self):
        try:
            self.abordo_licenca2 =  DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'A Bordo', 2)
        except:
            self.abordo_licenca2 = '0'

        try:
            self.baixado_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Baixado', 2)
        except:
            self.baixado_licenca2 = '0'

        try:
            self.visitacao_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Visitação', 2)
        except:
            self.visitacao_licenca2 = '0'

        try:
            self.dispdomiciliar_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Disp. Domiciliar', 2)
        except:
            self.dispdomiciliar_licenca2 = '0'

        try:
            self.hnmd_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'HNMD', 2)
        except:
            self.hnmd_licenca2 = '0'

        try:
            self.lts_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'LTS', 2)
        except:
            self.lts_licenca2 = '0'

        try:
            self.licenciados_angra_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Angra', 2)
        except:
            self.licenciados_angra_licenca2 = '0'

        try:
            self.licenciados_rio_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Rio', 2)
        except:
            self.licenciados_rio_licenca2 = '0'

        try:
            self.stgt_licenca2 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'ST/GT', 2)
        except:
            self.stgt_licenca2 = '0'

    def organiza_terceiro_licenca(self):
        try:
            self.abordo_licenca3 =  DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'A Bordo', 3)
        except:
            self.abordo_licenca3 = '0'

        try:
            self.baixado_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Baixado', 3)
        except:
            self.baixado_licenca3 = '0'

        try:
            self.visitacao_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Visitação', 3)
        except:
            self.visitacao_licenca3 = '0'

        try:
            self.dispdomiciliar_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Disp. Domiciliar', 3)
        except:
            self.dispdomiciliar_licenca3 = '0'

        try:
            self.hnmd_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'HNMD', 3)
        except:
            self.hnmd_licenca3 = '0'

        try:
            self.lts_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'LTS', 3)
        except:
            self.lts_licenca3 = '0'

        try:
            self.licenciados_angra_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Angra', 3)
        except:
            self.licenciados_angra_licenca3 = '0'

        try:
            self.licenciados_rio_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'Licença Rio', 3)
        except:
            self.licenciados_rio_licenca3 = '0'

        try:
            self.stgt_licenca3 = DatabaseTolda().contar_info(SHEET_LICENCAS, 'Situação', 'ST/GT', 3)
        except:
            self.stgt_licenca3 = '0'

    def input_texto_chave(self, chave_pesquisa):

        self.chave_input = chave_pesquisa

        info_chaves = busca_chave(self.chave_input, DatabaseTolda().pegar_table(SHEET_CHAVES))
        self.chave_input = info_chaves[0]

        if info_chaves[0] == 'Chave não encontrada':
            self.chave_nome = 'Chave não encontrada'
        else:
            if len(info_chaves[1]) > 45:
                palavras = info_chaves[1].split(' ')
                primeira_metade = ' '.join(palavras[0: len(palavras)//2+1])
                segunda_metade = ' '.join(palavras[len(palavras)//2+1:])
                self.chave_nome = info_chaves[0] + ' - ' + '\n'.join([primeira_metade, segunda_metade])
            else:
                self.chave_nome = info_chaves[0] + ' - ' + info_chaves[1]
        self.chave_anteriormente_com = info_chaves[2]
        self.chave_atualmente_com = info_chaves[3]
        self.chave_ultima_alteracao = info_chaves[4]

    def dar_chave(self, chave_pesquisa, asp_retirando):
        try:
            self.chave_input = chave_pesquisa

            info_chave = busca_chave(self.chave_input, DatabaseTolda().pegar_table(SHEET_CHAVES))

            DatabaseTolda().update_info(SHEET_CHAVES, 'Anterior', info_chave[3], 'Numero da Chave', self.chave_input)
            DatabaseTolda().update_info(SHEET_CHAVES, 'Atual', asp_retirando, 'Numero da Chave', self.chave_input)
            momento_retirada = datetime.now().strftime('%d/%m/%Y %H:%M')
            DatabaseTolda().update_info(SHEET_CHAVES, 'Última Alteração', momento_retirada, 'Numero da Chave', self.chave_input)

            if info_chave[0] == 'Chave não encontrada':
                self.chave_nome = 'Chave não encontrada'
            else:
                self.chave_nome = info_chave[0] + ' - ' + info_chave[1]
            self.chave_anteriormente_com = info_chave[3]
            self.chave_atualmente_com = asp_retirando
            self.chave_ultima_alteracao = momento_retirada

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

            info_chave = busca_chave(self.chave_input, DatabaseTolda().pegar_table('Chaves'))

            DatabaseTolda().update_info('Chaves', 'Anterior', info_chave[3], 'Numero da Chave', self.chave_input)
            DatabaseTolda().update_info('Chaves', 'Atual', 'Claviculário', 'Numero da Chave', self.chave_input)
            momento_devolucao = datetime.now().strftime('%d/%m/%Y %H:%M')
            DatabaseTolda().update_info('Chaves', 'Última Alteração', momento_devolucao, 'Numero da Chave', self.chave_input)

            #O 'Info_Chaves' deve estar logo apos todas as alterações/ logo antes de ser escrito para evitar erros

            if info_chave[0] == 'Chave não encontrada':
                self.chave_nome = 'Chave não encontrada'
            else:
                self.chave_nome = info_chave[0] + ' - ' + info_chave[1]
            self.chave_atualmente_com = 'Claviculário'
            self.chave_anteriormente_com = info_chave[3]
            self.chave_ultima_alteracao = momento_devolucao

        except:
            #self.chave_input = 'Chave não encontrada'
            self.chave_nome = 'Chave não encontrada'
            self.chave_atualmente_com = 'Chave não encontrada'
            self.chave_anteriormente_com = 'Chave não encontrada'
            self.chave_ultima_alteracao = 'Chave não encontrada'

        self.organiza_claviculario()

    def organiza_claviculario(self):
        data = DatabaseTolda()
        self.chaves_claviculario = str(data.contar_info('Chaves', 'Atual', 'Claviculário'))
        self.chaves_fora = str(len(self.chaves.Atual) - int(self.chaves_claviculario))
        

class ScrollerPage1(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'A Bordo': 'green', 'Licença Rio': 'red', 'Licença Angra': 'yellow', 'Disp. Domiciliar': 'purple', 'LTS': 'blue', 'HNMD': 'pink', 'Baixado': 'orange', 'ST/GT': 'grey', 'BAIXA': 'black', 'Visitação': 'grey', 'UISM': 'brown'}
        data = []
        df_final = DatabaseTolda().pegar_table(SHEET_LICENCAS).query('Ano == 1')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data

    def atualizar(self):
        data = []
        print('cheguei aqui')
        df_final = DatabaseTolda().pegar_table(SHEET_LICENCAS).query('Ano == 1')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data
        self.refresh_from_data()

class ScrollerPage2(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'A Bordo': 'green', 'Licença Rio': 'red', 'Licença Angra': 'yellow', 'Disp. Domiciliar': 'purple', 'LTS': 'blue', 'HNMD': 'pink', 'Baixado': 'orange', 'ST/GT': 'grey', 'BAIXA': 'black', 'Visitação': 'grey', 'UISM': 'brown'}
        data = []
        df_final = DatabaseTolda().pegar_table(SHEET_LICENCAS).query('Ano == 2')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table(SHEET_LICENCAS).query('Ano == 2')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data
        self.refresh_from_data()

class ScrollerPage3(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'A Bordo': 'green', 'Licença Rio': 'red', 'Licença Angra': 'yellow', 'Disp. Domiciliar': 'purple', 'LTS': 'blue', 'HNMD': 'pink', 'Baixado': 'orange', 'ST/GT': 'grey', 'BAIXA': 'black', 'Visitação': 'grey', 'UISM': 'brown'}
        data = []
        df_final = DatabaseTolda().pegar_table(SHEET_LICENCAS).query('Ano == 3')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table(SHEET_LICENCAS).query('Ano == 3')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data
        self.refresh_from_data()

if __name__ == '__main__':
    # se o arquivo não existir, cria-o. Dessa maneira, evita-se erros relacionados à abertura futura do EXTERN_FILE
    arquivo = open(EXTERN_FILE,'a')
    arquivo.close()

    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))
    ControleGeralApp().run()
