import sqlite3
import pandas as pd
import datetime
import numpy as np

class DatabaseTolda:
    def __init__(self) -> None:
        self.conn = sqlite3.connect('database_tolda.db')
        self.c = self.conn.cursor()
    
    def criar_database(self, arquivo_excel):
        licencas = pd.read_excel(arquivo_excel, 'Licenças')
        pben = pd.read_excel(arquivo_excel, 'PBEN')
        chaves = pd.read_excel(arquivo_excel, 'Chaves')
        parte_alta = pd.read_excel(arquivo_excel, 'ParteAlta')
        chefe_de_dia = pd.read_excel(arquivo_excel, 'ChefeDia')

        pben['Data de Nascimento'] = pben['Data de Nascimento'].astype('string').apply(lambda x: x.split(' ')[0])
        pben.drop('Carimbo de data/hora', axis=1, inplace=True)

        licencas.to_sql('Licenças', con=self.conn, if_exists='replace', index=False)
        pben.to_sql('PBEN', con=self.conn, if_exists='replace', index=False)
        chaves.to_sql('Chaves', con=self.conn, if_exists='replace', index=False)
        parte_alta.to_sql('ParteAlta', con=self.conn, index=False, if_exists='replace')
        chefe_de_dia.to_sql('ChefeDia', con=self.conn, index=False, if_exists='replace')


    def pegar_table(self, table):
        self.c.execute(f'SELECT * FROM {table}')
        dados_table = self.c.fetchall()
        self.c.execute(f"PRAGMA table_info({table})")
        colunas = [row[1] for row in self.c.fetchall()]
        self.c.close()
        return pd.DataFrame(dados_table, columns=colunas)
    
    def contar_info(self, table, coluna, valor, year=None):
        if year:
            self.c.execute(f"SELECT COUNT({coluna}) AS Valor FROM {table} WHERE [{coluna}] = '{valor}' AND Ano = '{year}'")
        else:
            self.c.execute(f"SELECT COUNT({coluna}) AS Valor FROM {table} WHERE [{coluna}] = '{valor}'")
    
        data = str(self.c.fetchone()[0])
        self.c.close()
        return data
    
    def update_info(self, table, coluna, valor, key_column, parametro):
        sql = f"UPDATE {table} SET [{coluna}] = (?) WHERE  [{key_column}] = (?)"
        self.c.execute(sql, (valor, parametro))
        self.conn.commit()
        self.c.close()

    def pegar_info(self, table, coluna, key_column, parametro):
        sql = f'SELECT [{coluna}] from {table} WHERE [{key_column}] = "{parametro}"'
        self.c.execute(sql)
        informacao = self.c.fetchone()
        if informacao is None:
            return 'Não Encontrado'
        else:
            return informacao[0]
        
    def insert_row(self, table,**values):
        data = {key.replace('_', ' ').replace('z', '.').replace('y', '/').replace('x', 'º'): [value] for key, value in values.items()}
        df = pd.DataFrame(data)
        df.to_sql(table, con=self.conn, if_exists='append', index=False)
        self.c.close()
        self.conn.close()    
        

    










