import sqlite3
import pandas as pd
import os
import sys

class DatabaseTolda:
    def __init__(self) -> None:
        self.conn = sqlite3.connect(resource_path('database_tolda.db'))
        self.c = self.conn.cursor()
    
    def criar_database(self, arquivo_excel):
        licencas = pd.read_excel(arquivo_excel, 'Licenças')
        pben = pd.read_excel(arquivo_excel, 'PBEN')
        chaves = pd.read_excel(arquivo_excel, 'Chaves')

        pben['NASC'] = pben['NASC'].astype('string').apply(lambda x: x.split(' ')[0])

        licencas.to_sql('Licenças', con=self.conn, if_exists='replace', index=False)
        pben.to_sql('PBEN', con=self.conn, if_exists='replace', index=False)
        chaves.to_sql('Chaves', con=self.conn, if_exists='replace', index=False)



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
        

def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath("")

    return os.path.join(base_path, relative_path)

