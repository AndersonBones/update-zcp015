from Export import Export_ZCP015
from Base import Base_ZCP015
import PySimpleGUI as sg
from pathlib import Path
import os
import threading
from time import sleep
from sapGui import sapconnection
from process import process_status



# this specifies the number of
# CPUs to use. 
os.environ["MODIN_CPUS"] = "2"
sg.theme('Reddit')


# sg.OK(button_text='Exportar Base SAP', button_color="#1D976C", s=15, pad=(5, (15, 0)))
class Gui():
    def __init__(self) -> None:
        self.popup_status = False
        self.layout = [[sg.Text('Atualização de Base de dados ZCP015', text_color="#2F80ED", font=('Arial Bold', 15))],
                [sg.Text('Selecione a base com os novos dados')],
                [sg.Input( size=70, key='-IN-', default_text='C:\\TEMP\\EXPORT.XLSX'), sg.FileBrowse(button_text='Procurar', s=15, file_types=(("Excel files","*.xlsx"),))],
                [sg.Checkbox("Atualizar Base ZCP015", key='s1')],
                [sg.Exit(button_text='Sair', button_color="red", s=6, pad=(5, (15, 0))), sg.OK(button_text='Processar', s=10, pad=(5, (15, 0))), sg.Text('', key='RUN', s=100, pad=(15, (20, 0)))]]

        self.window = sg.Window('ATUALIZAÇÃO ZCP015', self.layout, size=(600, 180))

    def set_process_status(self, msg):
        self.window['RUN'].update(msg)

    def run_export_process(self, path, s1): 
        

        if s1:
            Export_ZCP015(path, self.set_process_status, process_status).update()
            Base_ZCP015(path, self.set_process_status, process_status).start_update()
        
        if not s1:
            Export_ZCP015(path,self.set_process_status, process_status).update()


        self.popup_status = True
    
    def run_sap_gui(self):
        sapconnection()
       
    def main(self):
       
        while True:
            self.event, self.values = self.window.read(timeout=200)
        
            if self.event in (sg.WINDOW_CLOSED, 'Sair'):
                break
            
            if os.path.exists(self.values['-IN-']) == False and self.event == 'Processar':
                sg.popup_error('Selecione um arquivo válido!', title='Error')
            

            if os.path.exists(self.values['-IN-']) == True and self.event == "Processar" and not self.popup_status:
                thr1 = threading.Thread(target=self.run_export_process, args=(self.values['-IN-'],self.values['s1']), daemon=True)
                thr1.start()
            
            # if self.event == 'Exportar Base SAP' and not self.popup_status:
            #     thr2 = threading.Thread(target=self.run_sap_gui, daemon=True)
            #     thr2.start()
                    
            if self.popup_status:
                sg.popup_ok(f'Tarefa Concluida com sucesso :)\n{"-"*50}', title='Salvo com sucesso')
                self.popup_status = False
        self.window.close()


if __name__ == '__main__':
    Gui().main()


