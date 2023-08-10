import sys
import os
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import pandas as pd


class Form(QMainWindow):

    def __init__(self, parent=None):
        super().__init__(parent)

        self.plainTextEdit = QPlainTextEdit()
        self.plainTextEdit.setFont(QFont('Arial', 11))

        openDirButton = QPushButton("Select Directory with xlsx")
        openDirButton.clicked.connect(self.getDirectory)

        getFileNameButton = QPushButton("Select File with price")
        getFileNameButton.clicked.connect(self.getFileName)

        saveFileNameButton = QPushButton("Finish this Sheets!")
        saveFileNameButton.clicked.connect(self.saveFile)

        layoutV = QVBoxLayout()
        layoutV.addWidget(openDirButton)
        layoutV.addWidget(getFileNameButton)
        layoutV.addWidget(saveFileNameButton)

        layoutH = QHBoxLayout()
        layoutH.addLayout(layoutV)
        layoutH.addWidget(self.plainTextEdit)

        centerWidget = QWidget()
        centerWidget.setLayout(layoutH)
        self.setCentralWidget(centerWidget)

        self.resize(740, 300)
        self.setWindowTitle("Comrade")
        self.setWindowIcon(QIcon("./images/background.png"))
        self.setStyleSheet('.QWidget {border-image: url(./images/back_flag.svg) 0 0 0 0 stretch stretch;}')

        self.directory_path = ""
        self.file_path = ""
        self.common_dataframe = pd.DataFrame()

    def getDirectory(self):  # <-----
        dirlist = QFileDialog.getExistingDirectory(self, "Select directory", ".")
        if dirlist != '':
            self.plainTextEdit.appendHtml("<br>Directory was selected: <b>{}</b>".format(dirlist))
            self.directory_path = "{}".format(dirlist)

    def getFileName(self):
        filename, filetype = QFileDialog.getOpenFileName(self,
                                                         "Select file",
                                                         ".",
                                                         "All Files(*)")
        if filename != '':
            self.plainTextEdit.appendHtml("<br>File was selected: <b>{}</b>"
                                          "".format(filename, filetype))
            self.file_path = "{}".format(filename)

    def saveFile(self):

        if self.directory_path == "":
            self.plainTextEdit.appendHtml("<br><b>Select directory, please!</b>")
            return 0

        with os.scandir(self.directory_path) as files:
            list_of_files = [file.path for file in files if file.is_file() if file.path[-5:] == ".xlsx"]

        self.common_dataframe = (pd.concat([pd.read_excel(file_path) for file_path in list_of_files],
                                           ignore_index=True))

        if self.file_path == '':
            self.common_dataframe = (self.common_dataframe
                                     [['Appt.Provider Name', 'CPT Code(s)']]
                                     .rename(columns={'Appt.Provider Name': 'appt_provider_name', 'CPT Code(s)': 'cpt_codes'})
                                     .query("cpt_codes.notnull()")
                                     .assign(cpt_codes=lambda x: x.cpt_codes.str.split(','))
                                     .explode('cpt_codes')
                                     .groupby(['appt_provider_name', 'cpt_codes'])
                                     .agg({'cpt_codes': ['count']})
                                     )
        else:
            df_price = pd.read_excel(self.file_path)

            df_price = (df_price
                        [['Appt.Provider Name', 'CPT Code(s)', 'Price']]
                        .rename(columns={'Appt.Provider Name': 'appt_provider_name', 'CPT Code(s)': 'cpt_codes', 'Price': 'price'})
                        .astype({'appt_provider_name': 'string', 'cpt_codes': 'string', 'price': 'float32'})
                        )

            self.common_dataframe = (self.common_dataframe
                                     [['Appt.Provider Name', 'CPT Code(s)']]
                                     .rename(columns={'Appt.Provider Name': 'appt_provider_name', 'CPT Code(s)': 'cpt_codes'})
                                     .query("cpt_codes.notnull()")
                                     .assign(cpt_codes=lambda x: x.cpt_codes.str.split(','))
                                     .explode('cpt_codes')
                                     .merge(df_price, how="left")
                                     .groupby(['appt_provider_name', 'cpt_codes'])
                                     .agg({'cpt_codes': ['count'], 'price': ['max', 'sum']})
                                     .reset_index()
                                     )

            self.common_dataframe.columns = [tup[1] if tup[1] else tup[0] for tup in self.common_dataframe.columns]

            df1 = self.common_dataframe.copy()

            df1 = (df1
                   .groupby(['appt_provider_name'])
                   .agg({'sum': ['sum']})
                   .reset_index().reset_index()
                   .rename(columns={'sum': 'total_sum'})
                   )
            df1.columns = [tup[1] if tup[1] else tup[0] for tup in df1.columns]

            self.common_dataframe = (self.common_dataframe
                                     .merge(df1, how='left')
                                     .groupby(['appt_provider_name', 'total_sum', 'cpt_codes', 'count', 'max', 'sum'])
                                     .agg({'total_sum': ['max']})
                                     .drop(columns=[('total_sum', 'max')])
                                     )
            self.common_dataframe.columns = [tup[1] if tup[1] else tup[0] for tup in self.common_dataframe.columns]

        file_directory = self.directory_path + "/ready.xlsx"
        self.common_dataframe.to_excel(file_directory)
        self.plainTextEdit.appendHtml("<br>File was saved as: <b>{}</b>"
                                      "".format(file_directory))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Form()
    ex.show()
    sys.exit(app.exec_())
