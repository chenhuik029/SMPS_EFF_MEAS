# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'UI\Support.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog_support(object):
    def setupUi(self, Dialog_support):
        Dialog_support.setObjectName("Dialog_support")
        Dialog_support.resize(758, 354)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Dialog_support.sizePolicy().hasHeightForWidth())
        Dialog_support.setSizePolicy(sizePolicy)
        Dialog_support.setMinimumSize(QtCore.QSize(758, 354))
        Dialog_support.setMaximumSize(QtCore.QSize(758, 354))
        self.verticalLayout = QtWidgets.QVBoxLayout(Dialog_support)
        self.verticalLayout.setObjectName("verticalLayout")
        self.frame = QtWidgets.QFrame(Dialog_support)
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.frame)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.frame)
        font = QtGui.QFont()
        font.setPointSize(11)
        font.setBold(True)
        font.setWeight(75)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.verticalLayout_2.addWidget(self.label_2)
        self.label_email = QtWidgets.QLabel(self.frame)
        self.label_email.setOpenExternalLinks(False)
        self.label_email.setTextInteractionFlags(QtCore.Qt.TextBrowserInteraction)
        self.label_email.setObjectName("label_email")
        self.verticalLayout_2.addWidget(self.label_email)
        self.label = QtWidgets.QLabel(self.frame)
        self.label.setObjectName("label")
        self.verticalLayout_2.addWidget(self.label)
        self.verticalLayout.addWidget(self.frame)
        self.buttonBox = QtWidgets.QDialogButtonBox(Dialog_support)
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.verticalLayout.addWidget(self.buttonBox)

        self.retranslateUi(Dialog_support)
        self.buttonBox.accepted.connect(Dialog_support.accept)
        self.buttonBox.rejected.connect(Dialog_support.reject)
        QtCore.QMetaObject.connectSlotsByName(Dialog_support)

    def retranslateUi(self, Dialog_support):
        _translate = QtCore.QCoreApplication.translate
        Dialog_support.setWindowTitle(_translate("Dialog_support", "Support"))
        self.label_2.setText(_translate("Dialog_support", "This is an automated power supply efficiency measurement tools.\n"
"For any support, please contact the following provided contact."))
        self.label_email.setText(_translate("Dialog_support", "<html><head/><body><p><a href=\"https://outlook.live.com/mail/0/inbox\"><span style=\" text-decoration: underline; color:#0000ff;\">If there is any technical support or further understanding, please contact: chenhui_k029@hotmail.com"))
        self.label.setText(_translate("Dialog_support", "Created @2021/5 version 1.0"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog_support = QtWidgets.QDialog()
    ui = Ui_Dialog_support()
    ui.setupUi(Dialog_support)
    Dialog_support.show()
    sys.exit(app.exec_())
