from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtGui import QDoubleValidator
from UIpy import Main_ui, MsgBoxOk, MsgBoxOkCancel, MsgBoxAutoClose
from threading import Thread


# Main UI
class MainUI(QMainWindow, Main_ui.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.default_configuration()
        self.checkBox_ExtSupUsed.clicked.connect(self.Ext_Supply_used_checked)
        self.checkBox_ReadIntVol_PSEquip.clicked.connect(self.DMM1_used_checked)
        self.checkBox_ReadIntCur_PSEquip.clicked.connect(self.DMM2_used_checked)
        self.checkBox_ReadOutVol_ELoad.clicked.connect(self.DMM3_used_checked)
        self.checkBox_ReadOutCur_ELoad.clicked.connect(self.DMM4_used_checked)

    def default_configuration(self):
        self.Ext_Supply_used_checked(False)
        self.onlyFloat = QDoubleValidator()
        self.lineEdit_PS_VIN.setValidator(self.onlyFloat)
        self.lineEdit_PS_VIN_Limit.setValidator(self.onlyFloat)
        self.lineEdit_PS_Cur_Limit.setValidator(self.onlyFloat)
        self.lineEdit_IMAX.setValidator(self.onlyFloat)
        self.lineEdit_ISTEP.setValidator(self.onlyFloat)
        self.lineEdit_ISTART.setValidator(self.onlyFloat)

    def Ext_Supply_used_checked(self, status=True):
        self.comboBox_PS_Address.setEnabled(status)
        self.lineEdit_PS_VIN.setEnabled(status)
        self.lineEdit_PS_VIN_Limit.setEnabled(status)
        self.lineEdit_PS_Cur_Limit.setEnabled(status)
        self.checkBox_ReadIntCur_PSEquip.setEnabled(status)
        self.label_23.setEnabled(status)
        self.checkBox_ReadIntVol_PSEquip.setEnabled(status)
        self.label_19.setEnabled(status)

    def DMM1_used_checked(self):
        if self.checkBox_ReadIntVol_PSEquip.isChecked():
            self.comboBox_DMM_VI.setEnabled(False)
        else:
            self.comboBox_DMM_VI.setEnabled(True)

    def DMM2_used_checked(self):
        if self.checkBox_ReadIntCur_PSEquip.isChecked():
            self.comboBox_DMM_CI.setEnabled(False)
        else:
            self.comboBox_DMM_CI.setEnabled(True)

    def DMM3_used_checked(self):
        if self.checkBox_ReadOutVol_ELoad.isChecked():
            self.comboBox_DMM_VO.setEnabled(False)
        else:
            self.comboBox_DMM_VO.setEnabled(True)

    def DMM4_used_checked(self):
        if self.checkBox_ReadOutCur_ELoad.isChecked():
            self.comboBox_DMM_CO.setEnabled(False)
        else:
            self.comboBox_DMM_CO.setEnabled(True)


