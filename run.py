from mains_part2 import Ui_Form
from PyQt5.QtWidgets import *


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    Form = QWidget()
    ui = Ui_Form()
    ui.setupUi(Form, file_selected=True)
    ui.show_charts()
    Form.show()
    sys.exit(app.exec_())
