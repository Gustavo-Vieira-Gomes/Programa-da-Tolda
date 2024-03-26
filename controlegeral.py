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
from datetime import datetime, timedelta, date
from kivy.uix.recycleview import RecycleView
from kivy.uix.popup import Popup
from kivy.uix.label import Label
from banco_de_dados import DatabaseTolda
from sqlite3 import connect
import os
import pdb
import time
from fpdf import FPDF

EXTERN_FILE      = 'registro.txt'
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


class SuporteScreen(Screen):
    pass

class RegistroParteAltaScreen(Screen):
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
        self.organiza_quarto_licenca()

        self.chaves = DatabaseTolda().pegar_table(SHEET_CHAVES)
        self.organiza_claviculario()

        self.organiza_controle_geral_partealta()
        self.organiza_primeiro_partealta()
        self.organiza_segundo_partealta()
        self.organiza_terceiro_partealta()
        self.organiza_quarto_partealta()
        self.chefedia = DatabaseTolda().pegar_table(SHEET_CHEFE_DIA)

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
    chave_input             = StringProperty() #Objeto de pesquisa
    chave_nome              = StringProperty() #Nuero e Nome da chave, aparece na tela
    chave_atualmente_com    = StringProperty()
    chave_anteriormente_com = StringProperty()
    chave_ultima_alteracao  = StringProperty()
    chaves_claviculario     = StringProperty()
    chaves_fora             = StringProperty()

    # Parte Baixa
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
    status_download_cd      = StringProperty()
    

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
        self.sm.add_widget(ParteAltaScreen(name='partealta'))
        self.sm.add_widget(RegistroParteAltaScreen(name='registropartealta'))
        self.sm.add_widget(ChefeDiaScreen(name='chefedia'))
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
        info_licencas = busca_licenca(self.numero_atual, DatabaseTolda().pegar_table('Licenças'))
        self.situacao_atual_licenca = info_licencas[0]
        self.ultima_alteracao_licenca = info_licencas[1]
        print("Botão pressionado: " + self.BOTAO_PRESSIONADO)

    def registro_externo_regs_lics(self, num_int, nome, situacao, day):
        with open('registro.txt', 'r', encoding='utf-8') as file:
            lines = file.readlines()
        dias_da_semana = ['SEGUNDA','TERÇA', 'QUARTA', 'QUINTA', 'SEXTA', 'SÁBADO', 'DOMINGO']
        for index, line in enumerate(lines):
            if line.split(' - ')[0] == dias_da_semana[datetime.weekday(day)]:
                if line.split(' - ')[1][:-1] == day.strftime('%d/%m/%Y'):
                    for line_after in lines[index+1:]:
                        if line_after.split(' - ')[0] == dias_da_semana[datetime.weekday(day-timedelta(1))]:
                            lines.insert(lines.index(line_after)-1, f'{str(num_int).upper()} {str(nome).upper()} - {str(situacao).upper()} - {datetime.now().strftime('%d/%m/%Y  %H:%M:%S')}\n')
                else:
                    # Caso 1: 1º linha
                    # Caso 2: Meio do Texto
                    # Caso 3: última linha


                    lines.insert(0, '\n')
                    lines.insert(0, f'{str(num_int).upper()} - {str(nome).upper()} - {str(situacao).upper()} - {datetime.now().strftime('%d/%m/%Y  %H:%M:%S')}\n')
                    lines.insert(0, '\n')
                    lines.insert(0, f'{dias_da_semana[datetime.weekday(day)]} - {day.strftime("%d/%m/%Y")}\n'.upper())
        with open('registro.txt', 'w', encoding='utf-8') as file:
            file.writelines(lines)

    def create_pdf(self, input_file):
        # Create a new FPDF object
        pdf = FPDF()
        # Open the text file and read its contents
        with open(input_file, 'r', encoding='utf-8') as f:
            text = f.read()

        # Add a new page to the PDF
        pdf.add_page()

        # Set the font and font size
        pdf.set_font('times', size=10)

        # Write the text to the PDF
        pdf.write(5, text)

        # Save the PDF
        pdf.output('registros licencas.pdf')

    def refresh_app(self):
        self.organiza_controle_geral_licenca()
        self.organiza_primeiro_licenca()
        self.organiza_segundo_licenca()
        self.organiza_terceiro_licenca()
        self.organiza_quarto_licenca()
        self.organiza_claviculario()
        self.organiza_controle_geral_partealta()
        self.organiza_primeiro_partealta()
        self.organiza_segundo_partealta()
        self.organiza_terceiro_partealta()
        self.organiza_quarto_partealta()
        self.sm.get_screen('registrolicencas').ids.scroller1.atualizar()
        self.sm.get_screen('registrolicencas').ids.scroller2.atualizar()
        self.sm.get_screen('registrolicencas').ids.scroller3.atualizar()
        self.sm.get_screen('registrolicencas').ids.scroller4.atualizar()

    def reiniciar_licenca_reg(self, situação):
        try:
            conn = connect('database_tolda.db')
            c = conn.cursor()
            sql = f"UPDATE Licenças SET [Situação] = '{situação}' WHERE  [Situação] <> 'BAIXA'"
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
        elif (button_text == 'Regresso' and DatabaseTolda().pegar_info('Licenças', 'Situação', 'Número Interno', self.numero_atual) == 'A Bordo') or (button_text == DatabaseTolda().pegar_info('Licenças', 'Situação', 'Número Interno', self.numero_atual)):
            pass
        else:
            if DatabaseTolda().pegar_info('Licenças', 'Situação', 'Número Interno', self.numero_atual) != 'BAIXA':
                try:
                    if button_text == 'Regresso':
                        DatabaseTolda().update_info('Licenças', 'Situação', 'A Bordo', 'Número Interno', self.numero_atual)
                        try:
                            self.registro_externo_regs_lics(self.numero_atual, DatabaseTolda().pegar_info('Licenças', 'Nome de Guerra', 'Número Interno', self.numero_atual), 'A Bordo', datetime.now())
                        except Exception as e:
                            print(e)
                    else:
                        DatabaseTolda().update_info('Licenças', 'Situação', button_text, 'Número Interno', self.numero_atual)
                        try:    
                            self.registro_externo_regs_lics(self.numero_atual,DatabaseTolda().pegar_info('Licenças', 'Nome de Guerra', 'Número Interno', self.numero_atual), str(button_text), datetime.now())
                        except Exception as e:
                            print(e)    
                    DatabaseTolda().update_info('Licenças', 'Última Alteração', datetime.now().strftime('%d/%m/%Y %H:%M'), 'Número Interno', self.numero_atual)

                except:
                    self.numero_atual = 'Selecione um aspirante'
                    self.nome_guerra = ''

        info_licencas = busca_licenca(self.numero_atual, DatabaseTolda().pegar_table('Licenças'))
        self.situacao_atual_licenca = info_licencas[0]
        self.ultima_alteracao_licenca = info_licencas[1]

        self.organiza_controle_geral_licenca()
        self.organiza_primeiro_licenca()
        self.organiza_segundo_licenca()
        self.organiza_terceiro_licenca()
        self.organiza_quarto_licenca()

    def organiza_controle_geral_licenca(self):
        try:
            self.abordo_licenca =  DatabaseTolda().contar_info('Licenças', 'Situação', 'A Bordo')
        except:
            self.abordo_licenca = '0'

        try:
            self.baixado_licenca = DatabaseTolda().contar_info('Licenças', 'Situação', 'Baixado')
        except:
            self.baixado_licenca = '0'

        try:
            self.crestricao_licenca = DatabaseTolda().contar_info('Licenças', 'Situação', 'C/ Restrição')
        except:
            self.crestricao_licenca = '0'

        try:
            self.dispdomiciliar_licenca = DatabaseTolda().contar_info('Licenças', 'Situação', 'Disp. Domiciliar')
        except:
            self.dispdomiciliar_licenca = '0'

        try:
            self.hnmd_licenca = DatabaseTolda().contar_info('Licenças', 'Situação', 'HNMD')
        except:
            self.hnmd_licenca = '0'

        try:
            self.lts_licenca = DatabaseTolda().contar_info('Licenças', 'Situação', 'LTS')
        except:
            self.lts_licenca = '0'

        try:
            self.licenciados_licenca = DatabaseTolda().contar_info('Licenças', 'Situação', 'Licença')
        except Exception as e:
            print(e)
            self.licenciados_licenca = '0'

        try:
            self.stgt_licenca = DatabaseTolda().contar_info('Licenças', 'Situação', 'ST/GT')
        except:
            self.stgt_licenca = '0'

    def organiza_primeiro_licenca(self):
        try:
            self.abordo_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'A Bordo', 1)
        except:
            self.abordo_licenca1 = '0' 
        try:
            self.baixado_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Baixado', 1)
        except:
            self.baixado_licenca1 = '0'
        try:
            self.crestricao_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'C/ Restrição', 1)
        except:
            self.crestricao_licenca1 = '0'
        try:
            self.dispdomiciliar_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Disp. Domiciliar', 1)
        except:
            self.dispdomiciliar_licenca1 = '0'
        try:
            self.hnmd_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'HNMD', 1)
        except:
            self.hnmd_licenca1 = '0'
        try:
            self.lts_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'LTS', 1)
        except:
            self.lts_licenca1 = '0'
        try:
            self.licenciados_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Licença', 1)
        except:
            self.licenciados_licenca1 = '0'
        try:
            self.stgt_licenca1 = DatabaseTolda().contar_info('Licenças', 'Situação', 'ST/GT', 1)
        except:
            self.stgt_licenca1 = '0'

    def organiza_segundo_licenca(self):
        try:
            self.abordo_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'A Bordo', 2)
        except:
            self.abordo_licenca2 = '0'
        try:
            self.baixado_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Baixado', 2)
        except:
            self.baixado_licenca2 = '0'
        try:
            self.crestricao_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'C/ Restrição', 2)
        except:
            self.crestricao_licenca2 = '0'
        try:
            self.dispdomiciliar_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Disp. Domiciliar', 2)
        except:
            self.dispdomiciliar_licenca2 = '0'
        try:
            self.hnmd_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'HNMD', 2)
        except:
            self.hnmd_licenca2 = '0'
        try:
            self.lts_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'LTS', 2)
        except:
            self.lts_licenca2 = '0'
        try:
            self.licenciados_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Licença', 2)
        except:
            self.licenciados_licenca2 = '0'
        try:
            self.stgt_licenca2 = DatabaseTolda().contar_info('Licenças', 'Situação', 'ST/GT', 2)
        except:
            self.stgt_licenca2 = '0'

    def organiza_terceiro_licenca(self):
        try:
            self.abordo_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'A Bordo', 3)
        except:
            self.abordo_licenca3 = '0'
        try:
            self.baixado_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Baixado', 3)
        except:
            self.baixado_licenca3 = '0'
        try:
            self.crestricao_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'C/ Restrição', 3)
        except:
            self.crestricao_licenca3 = '0'
        try:
            self.dispdomiciliar_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Disp. Domiciliar', 3)
        except:
            self.dispdomiciliar_licenca3 = '0'
        try:
            self.hnmd_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'HNMD', 3)
        except:
            self.hnmd_licenca3 = '0'
        try:
            self.lts_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'LTS', 3)
        except:
            self.lts_licenca3 = '0'
        try:
            self.licenciados_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Licença', 3)
        except:
            self.licenciados_licenca3 = '0'
        try:
            self.stgt_licenca3 = DatabaseTolda().contar_info('Licenças', 'Situação', 'ST/GT', 3)
        except:
            self.stgt_licenca3 = '0'

    def organiza_quarto_licenca(self):
        try:
            self.abordo_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'A Bordo', 4)
        except:
            self.abordo_licenca4 = '0'
        try:
            self.baixado_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Baixado', 4)
        except:
            self.baixado_licenca4 = '0'
        try:
            self.crestricao_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'C/ Restrição', 4)
        except:
            self.crestricao_licenca4 = '0'
        try:
            self.dispdomiciliar_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Disp. Domiciliar', 4)
        except:
            self.dispdomiciliar_licenca4 = '0'
        try:
            self.hnmd_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'HNMD', 4)
        except:
            self.hnmd_licenca4 = '0'
        try:
            self.lts_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'LTS', 4)
        except:
            self.lts_licenca4 = '0'
        try:
            self.licenciados_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'Licença', 4)
        except:
            self.licenciados_licenca4 = '0'
        try:
            self.stgt_licenca4 = DatabaseTolda().contar_info('Licenças', 'Situação', 'ST/GT', 4)
        except:
            self.stgt_licenca4 = '0'
        
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

    def consultar_partealta(self, chave_pesquisa):
        self.consulta_pben(chave_pesquisa)

        info_partealta = busca_licenca(self.numero_atual, DatabaseTolda().pegar_table(SHEET_PARTE_ALTA))
        self.situacao_atual_partealta = info_partealta[0]
        self.ultima_alteracao_partealta = info_partealta[1]

    botao_parte_alta = StringProperty('')



    def atualiza_partealta(self, button_text):
        if button_text == '':
            return
        else:
            try:
                DatabaseTolda().update_info(SHEET_PARTE_ALTA, 'Situação', button_text, 'Número Interno', self.numero_atual)
                momento_alt_situacao = datetime.now().strftime('%d/%m/%Y %H:%M')
                DatabaseTolda().update_info(SHEET_PARTE_ALTA, 'Última Alteração', momento_alt_situacao, 'Número Interno', self.numero_atual)
            except:
                self.numero_atual = 'Selecione um aspirante'
                self.nome_guerra = ''

        self.situacao_atual_partealta = button_text
        self.ultima_alteracao_partealta = momento_alt_situacao

        self.organiza_controle_geral_partealta()
        self.organiza_primeiro_partealta()
        self.organiza_segundo_partealta()
        self.organiza_terceiro_partealta()
        self.organiza_quarto_partealta()

    def organiza_controle_geral_partealta(self):
        try:
            self.partealta_partelta = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Parte Alta')
        except:
            self.partealta_partelta = '0'
        try:
            self.tfm_partealta = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'TFM')
        except:
            self.tfm_partealta = '0'
        try:
            self.saladeestado_partealta = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Sala de Estado')
        except:
            self.saladeestado_partealta = '0'
        try:
            self.enfermaria_partealta = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Enfermaria')
        except:
            self.enfermaria_partealta = '0'
        try:
            self.banco_partealta = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Banco')
        except:
            self.banco_partealta = '0'
        try:
            self.biblioteca_partealta = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Biblioteca')
        except:
            self.biblioteca_partealta = '0'

    def organiza_primeiro_partealta(self):
        try:
            self.partealta_partelta1 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Parte Alta', 1)
        except:
            self.partealta_partelta1 = '0'
        try:
            self.tfm_partealta1 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'TFM', 1)
        except:
            self.tfm_partealta1 = '0'
        try:
            self.saladeestado_partealta1 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Sala de Estado', 1)
        except:
            self.saladeestado_partealta1 = '0'
        try:
            self.enfermaria_partealta1 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Enfermaria', 1)
        except:
            self.enfermaria_partealta1 = '0'
        try:
            self.banco_partealta1 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Banco', 1)
        except:
            self.banco_partealta1 = '0'
        try:
            self.biblioteca_partealta1 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Biblioteca', 1)
        except:
            self.biblioteca_partealta1 = '0'

    def organiza_segundo_partealta(self):
        try:
            self.partealta_partelta2 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Parte Alta', 2)
        except:
            self.partealta_partelta2 = '0'
        try:
            self.tfm_partealta2 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'TFM', 2)
        except:
            self.tfm_partealta2 = '0'
        try:
            self.saladeestado_partealta2 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Sala de Estado', 2)
        except:
            self.saladeestado_partealta2 = '0'
        try:
            self.enfermaria_partealta2 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Enfermaria', 2)
        except:
            self.enfermaria_partealta2 = '0'
        try:
            self.banco_partealta2 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Banco', 2)
        except:
            self.banco_partealta2 = '0'
        try:
            self.biblioteca_partealta2 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Biblioteca', 2)
        except:
            self.biblioteca_partealta2 = '0'

    def organiza_terceiro_partealta(self):
        try:
            self.partealta_partelta3 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Parte Alta', 3)
        except:
            self.partealta_partelta3 = '0'
        try:
            self.tfm_partealta3 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'TFM', 3)
        except:
            self.tfm_partealta3 = '0'
        try:
            self.saladeestado_partealta3 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Sala de Estado', 3)
        except:
            self.saladeestado_partealta3 = '0'
        try:
            self.enfermaria_partealta3 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Enfermaria', 3)
        except:
            self.enfermaria_partealta3 = '0'
        try:
            self.banco_partealta3 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Banco', 3)
        except:
            self.banco_partealta3 = '0'
        try:
            self.biblioteca_partealta3 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Biblioteca', 3)
        except:
            self.biblioteca_partealta3 = '0'

    def organiza_quarto_partealta(self):
        try:
            self.partealta_partelta4 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Parte Alta', 4)
        except:
            self.partealta_partelta4 = '0'
        try:
            self.tfm_partealta4 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'TFM', 4)
        except:
            self.tfm_partealta4 = '0'
        try:
            self.saladeestado_partealta4 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Sala de Estado', 4)
        except:
            self.saladeestado_partealta4 = '0'
        try:
            self.enfermaria_partealta4 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Enfermaria', 4)
        except:
            self.enfermaria_partealta4 = '0'
        try:
            self.banco_partealta4 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Banco', 4)
        except:
            self.banco_partealta4 = '0'
        try:
            self.biblioteca_partealta4 = DatabaseTolda().contar_info('ParteAlta', 'Situação', 'Biblioteca', 4)
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
        DatabaseTolda().insert_row('ChefeDia', DATAHORA=agora, Nx_Ajosca=self.numero_ajosca_cd, Nx_CHEFE_DE_DIA=self.numero_chefe_cd, QUARTO_DE_SERVIÇO=self.quarto_de_serviço_cd, COMPANHIA=self.companhia_cd, PELOTÃO=self.pelotao_cd, QTD_CINTOS=self.cintos_cd, BANDEIRA_ASPz_NASCIMENTO=self.bandeira_cd, QTD_ASP_LICENCIADOS=self.licenciados_cd, QTD_REGRESSOS=self.regressos_cd, SITUAÇÃO_COMPUTADORyALARME=self.computador_cd, SITUAÇÃO_MAPA_MUNDI=self.mapamundi_cd)
        #self.chefedia.loc[self.chefedia.index.max() + 1] = [agora,
        #                                                    self.numero_ajosca_cd,
        #                                                    self.numero_chefe_cd,
        #                                                    self.quarto_de_serviço_cd,
        #                                                    self.companhia_cd,
        #                                                    self.pelotao_cd,
        #                                                    self.cintos_cd,
        #                                                    self.bandeira_cd,
        #                                                    self.licenciados_cd,
        #                                                    self.regressos_cd,
        #                                                    self.computador_cd,
        #                                                    self.mapamundi_cd]

        #self.salvar_alteracoes()

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

    def download_registros_cd(self):
        df = DatabaseTolda().pegar_table(SHEET_CHEFE_DIA)
        pasta_atual = os.getcwd()
        caminho_arquivo = os.path.join(pasta_atual, 'registros chefe de dia.xlsx')
        try:
            df.to_excel(caminho_arquivo, sheet_name='registros CD', )
        except:
            self.status_download_cd = 'Não foi possível fazer o download'
        else:
            self.status_download_cd = 'Download concluído com Sucesso'

class ScrollerPage1(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'A Bordo': 'green', 'Licença': 'red', 'ST/GT': 'yellow', 'Disp. Domiciliar': 'purple', 'LTS': 'blue', 'HNMD': 'pink', 'Baixado': 'orange', 'C/ Restrição': 'grey', 'BAIXA': 'black'}
        data = []
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 1')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data

    def atualizar(self):
        data = []
        print('cheguei aqui')
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 1')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data
        self.refresh_from_data()

class ScrollerPage2(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'A Bordo': 'green', 'Licença': 'red', 'ST/GT': 'yellow', 'Disp. Domiciliar': 'purple', 'LTS': 'blue', 'HNMD': 'pink', 'Baixado': 'orange', 'C/ Restrição': 'grey', 'BAIXA': 'black'}
        data = []
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 2')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 2')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data
        self.refresh_from_data()

class ScrollerPage3(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'A Bordo': 'green', 'Licença': 'red', 'ST/GT': 'yellow', 'Disp. Domiciliar': 'purple', 'LTS': 'blue', 'HNMD': 'pink', 'Baixado': 'orange', 'C/ Restrição': 'grey', 'BAIXA': 'black'}
        data = []
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 3')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 3')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data
        self.refresh_from_data()

class ScrollerPage4(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'A Bordo': 'green', 'Licença': 'red', 'ST/GT': 'yellow', 'Disp. Domiciliar': 'purple', 'LTS': 'blue', 'HNMD': 'pink', 'Baixado': 'orange', 'C/ Restrição': 'grey', 'BAIXA': 'black'}
        data = []
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 4')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table('Licenças').query('Ano == 4')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {row.values[1]} - {row.values[2]}', 'background_color': self.dict_colors[row.values[2]]})
        self.data = data
        self.refresh_from_data()

class ScrollerParteAlta1(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'Parte Alta': 'green', 'Enfermaria': 'red', 'Biblioteca': 'purple', 'TFM': 'blue', 'Sala de Estado': 'orange', 'Banco': 'grey', 'BAIXA': 'black', 'LTS': 'pink'}
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 1')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 1')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data
        self.refresh_from_data()

class ScrollerParteAlta2(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'Parte Alta': 'green', 'Enfermaria': 'red', 'Biblioteca': 'purple', 'TFM': 'blue', 'Sala de Estado': 'orange', 'Banco': 'grey', 'BAIXA': 'black', 'LTS': 'pink'}
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 2')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 2')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data
        self.refresh_from_data()

class ScrollerParteAlta3(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'Parte Alta': 'green', 'Enfermaria': 'red', 'Biblioteca': 'purple', 'TFM': 'blue', 'Sala de Estado': 'orange', 'Banco': 'grey', 'BAIXA': 'black', 'LTS': 'pink' }
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 3')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 3')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data
        self.refresh_from_data()

class ScrollerParteAlta4(RecycleView):
    def __init__(self, **kwargs, ):
        super().__init__(**kwargs)
        self.dict_colors = {'Parte Alta': 'green', 'Enfermaria': 'red', 'Biblioteca': 'purple', 'TFM': 'blue', 'Sala de Estado': 'orange', 'Banco': 'grey', 'BAIXA': 'black', 'LTS': 'pink'}
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 4')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data

    def atualizar(self):
        data = []
        df_final = DatabaseTolda().pegar_table('ParteAlta').query('Ano == 4')
        for index, row in df_final.iterrows():
            data.append({'text': f'{row.values[0]} {DatabaseTolda().pegar_info('PBEN', 'Nome de Guerra', 'Número Interno Atual', row.values[0])} - {row.values[1]}', 'background_color': self.dict_colors[row.values[1]]})
        self.data = data
        self.refresh_from_data()


def criar_registro_semanal_txt():
    def string_dia(dia_atual: datetime, delta: timedelta):
            data_final = dia_atual - delta
            return data_final.strftime('%d/%m/%Y')
        
    if 'registros.txt' not in os.listdir():
        with open('registro.txt', 'w', encoding='utf-8') as file:
            day = datetime.now()
            #day = date(2024, 3, 24)
            wk_day = datetime.weekday(day)
            dias_da_semana = ['Segunda','Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo']
            linhas = [f"{dias_da_semana[wk_day- delta]} - {string_dia(day, timedelta(days=(indice)))}\n\n\n".upper() for indice, delta in enumerate(range(0, wk_day+1))]
            linhas += [f"{dias_da_semana[6 - delta + wk_day]} - {string_dia(day, timedelta(days=wk_day + indice + 1))}\n\n\n".upper() for indice, delta in enumerate(range(wk_day, 6))]
            file.writelines(linhas)

if __name__ == '__main__':
    if hasattr(sys, '_MEIPASS'):
        resource_add_path(os.path.join(sys._MEIPASS))
    criar_registro_semanal_txt()
    ControleGeralApp().run()
