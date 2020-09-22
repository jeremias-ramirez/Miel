import sys
from PyQt5.QtWidgets import QApplication
import view 
  
if __name__ == '__main__':

    app = QApplication(sys.argv)
    ex = view.App()
    sys.exit(app.exec_())


