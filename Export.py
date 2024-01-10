
from time import sleep
import openpyxl as xl
from datetime import datetime
import os
import locale
import pandas as pd


locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')


class Export_ZCP015():
    
    def __init__(self, path, set_process_status, process_status) -> None:
        self.set_msg = set_process_status
        self.set_process_status = process_status
        # Rename file
        self.path = path
        self.rename_path = path.replace('XLSX', 'xlsx')
        os.rename(self.path, self.rename_path)

        self.file_name = self.rename_path.split("/")[-1]
        try:
            self.set_msg(f'Lendo base: {self.file_name}... {self.set_process_status(1)}')
            self.wb = xl.load_workbook(self.rename_path) 
    
            self.Sheet=self.wb.sheetnames[0]
            self.df = pd.read_excel(self.rename_path, sheet_name=self.Sheet)

        except Exception as e:
            self.set_msg(f"Erro na leitura da base: {self.file_name}")
            

    def remove_null_romaneio(self):
        try:
            self.set_msg(f'Removendo Romaneios nulos... {self.set_process_status(2)}')

            self.df.dropna(subset=['Romaneio'], how='all', inplace=True)
        except Exception as e:
            self.set_msg('Erro ao remover romaneios nulos.')

    def GfConvert(self):
        self.set_msg(f'Convertendo Guia Florestal... {self.set_process_status(3)}')
        self.df['Guia Florestal'] = pd.to_numeric(self.df['Guia Florestal'], errors='coerce')
        self.set_msg(f"Concluido. {self.set_process_status(3)}")

    def dateTime_format(self):
        self.set_msg(f'Ajustando formado de data... {self.set_process_status(4)}')
        self.df['Dt. Agendamento'] = pd.to_datetime(self.df['Dt. Agendamento'], format='%d %b %Y', errors='coerce').dt.date
        self.df['Dt. Pesagem Inicial'] = pd.to_datetime(self.df['Dt. Pesagem Inicial'], format='%d %b %Y', errors='coerce').dt.date
        
        self.df['Data Nota Fiscal'] = pd.to_datetime(self.df['Data Nota Fiscal'], format='%d %b %Y', errors='coerce').dt.date
        self.df['Data de criação'] = pd.to_datetime(self.df['Data de criação'], format='%d %b %Y', errors='coerce').dt.date
        
        self.set_msg('Formato de data ajustado.')  

    def sort_data_pesagem(self):
        try:
            self.set_msg(f'Ordenando coluna "Dt. Pesagem Inicial"... {self.set_process_status(5)}')
            self.df.sort_values(by=['Dt. Pesagem Inicial', 'Hora Pesagem Inicial'], inplace=True)
            
        except Exception as e:
            self.set_msg('Erro ao ordenar coluna "Dt. Pesagem Inicial"...')


    def set_data_pesagem(self):
        try:
            self.set_msg(f'Tratando romaneios sem "Dt. Pesagem Inicial"... {self.set_process_status(6)}')
            if len(self.df['Dt. Pesagem Inicial']) == len(self.df['Data de criação']):
                for index in range(0,len(self.df['Dt. Pesagem Inicial'])):
                    if self.df.iat[index, 4] is pd.NaT:
                        self.df.iat[index, 4] = self.df.iat[index, 32]
        except Exception as e:
            self.set_msg('Erro ao Tratar romaneios sem "Dt. Pesagem Inicial"...')

  

    def auto_adjust_column(self, df):
        try:
            self.set_msg(f'Ajustando tamanho das colunas... {self.set_process_status(9)}')
            for column in df:
                column_length = max(self.df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column)
                self.writer.sheets[self.Sheet].set_column(col_idx, col_idx, column_length+2)
        except Exception as e:
            self.set_msg(f'Colunas ajustadas. {self.set_process_status(9)}')

    def remove_duplicates(self):
        try:
            self.set_msg(f'Removendo Linhas duplicadas... {self.set_process_status(7)}')
            self.df.drop_duplicates()
        except Exception as e:
            self.set_msg("Erro ao remover linhas duplicadas...")

    def save_file(self):
        try:
            self.writer = pd.ExcelWriter(self.rename_path, engine='xlsxwriter', date_format='d/m/yyyy')
            self.set_msg(f"Salvando Base tratada... {self.set_process_status(8)} ")

            self.df.to_excel(self.writer, sheet_name=self.Sheet, index=False, header=True)
            self.auto_adjust_column(self.df)
            
            self.writer.close()
            self.set_msg(f'Concluido. {self.set_process_status(10)}')
        except Exception as e:
            self.set_msg("Erro ao salvar base tratada...")
            self.set_msg(e)

    def update(self):
        self.remove_null_romaneio()
        self.set_data_pesagem()
        self.sort_data_pesagem()
        self.remove_duplicates()
        self.GfConvert()
        self.dateTime_format()
        self.save_file()
        



    
