from PyQt5 import QtCore
from PyQt5.QtWidgets import QTableWidget

class dragDropTable(QTableWidget):
    '''
    This class adds drag and drop functionality to QTableWidget. tbl_fileList is promoted to dragDropTable in Qt Designer.
    Only accepts *.max file drops. dropEvent executes addFiles(files) method implemented in maxConverter.py.
    '''
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(QtCore.Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile().endswith('.max')]
        self.parent.addFiles(files) # method implemented in maxConverter.py
        
if __name__ == '__main__':
    import sys
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    window = dragDropTable()
    window.show()
    sys.exit(app.exec_())