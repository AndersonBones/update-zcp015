from PyQt5 import QtWidgets
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import QObject, Qt, QThread, pyqtSignal, QCoreApplication
from PyQt5.QtWidgets import QDialog, QApplication, QFileDialog, QWidget, QMessageBox, QProgressBar, QMainWindow
from PyQt5.uic import loadUi
import sys
from Export import Export_ZCP015
from Base import    Base_ZCP015 
import threading
from time import sleep


class MainWindow(QDialog):
    def __init__(self) -> None:
        super(MainWindow, self).__init__()
        loadUi('gui.ui', self)
  
        self.search_file_btn.clicked.connect(self.browsefiles)
        self.run_process_btn.clicked.connect(self.run_thread)
        self.exit_button.clicked.connect(self.exit)
    
        
        self.movie = QMovie("C:\\TEMP\\Loading_2.gif")
        self.process.setMovie(self.movie)
        self.movie.start()

        self.msg = QMessageBox()
        self.msg.setWindowTitle('Tarefa concluida')
        self.msg.setText('Tarefa concluida com sucesso.')
        
        
    
    def success_msg(self):
        
        #self.msg.setIcon(QMessageBox.SP_DialogApplyButton)
        self.msg.exec_()
        
    def close_success_msg(self):
        self.msg.close()
    
    def browsefiles(self):
        self.file_path = QFileDialog.getOpenFileName(self,'Procurar Arquivo','C:\\TEMP', 'Excel Files (*.xls *.xlsx)')
        self.input_path.setText(self.file_path[0])
    
  
    def run_process(self):
        Export_ZCP015(self.file_path[0]).update()
        #Base_ZCP015(self.file_path[0]).start_update()
    
        self.success_msg()
        
    def run_thread(self):
        self.thr = threading.Thread(target=self.run_process, daemon=True)
        self.thr.start()
    
    def exit(self):
        self.hide()
        QCoreApplication.closeAllWindows()  # <- Crashes here, if I comment this app closes but doesnt kill process nor release terminal
        QCoreApplication.quit()
        

        
        
    
def run():
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    widget = QtWidgets.QStackedWidget()
    widget.addWidget(mainWindow)
    widget.setFixedWidth(665)
    widget.setFixedHeight(300)
    widget.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    run()
