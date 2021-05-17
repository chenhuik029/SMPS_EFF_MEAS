import math
from PyQt5.QtWidgets import QMainWindow
from PyQt5.QtCore import QThread, QObject, pyqtSignal
from PyQt5.QtGui import QDoubleValidator
from UIpy import Main_ui
from function_msgbox import msg_box_ok, msg_box_auto_close, msg_box_ok_cancel
from Instrument_PyVisa import Basic_PyVisa, PS_Kikusui_PyVisa, Eload_Chroma_PyVisa, DMM_Keysight_PyVisa
from threading import Thread
import threading
import re
import time
thread_stop = "FALSE"


# Main UI for Fixed VIN, Variable Vout Test
class FixedVIN_VarVOUT_UI(QMainWindow, Main_ui.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.default_configuration()
        self.list_instrument = Basic_PyVisa.Basic_PyVisa()
        self.combobox_equipment_list()

        self.error_count = 0
        self.PS_address = ""
        self.PS_target_Vin = 0
        self.PS_limit_Vin = 0
        self.PS_limit_Cin = 0
        self.ELoad_IMAX = 0
        self.ELoad_ISTEP = 0
        self.ELoad_ISTART = 0
        self.ELoad_address = ""
        self.DMM_VIN_address = ""
        self.DMM_CIN_address = ""
        self.DMM_VOUT_address = ""
        self.DMM_COUT_address = ""
        self.thread_count = 0

        self.checkBox_ExtSupUsed.clicked.connect(self.Ext_Supply_used_checked)
        self.checkBox_ReadIntVol_PSEquip.clicked.connect(self.DMM1_used_checked)
        self.checkBox_ReadIntCur_PSEquip.clicked.connect(self.DMM2_used_checked)
        self.checkBox_ReadOutVol_ELoad.clicked.connect(self.DMM3_used_checked)
        self.checkBox_ReadOutCur_ELoad.clicked.connect(self.DMM4_used_checked)
        self.pushButton_StartTest.clicked.connect(self.start_test)
        self.pushButton_AbortTest.clicked.connect(self.stop_test)

    # Default configuration setting
    def default_configuration(self):
        self.Ext_Supply_used_checked(False)
        self.onlyFloat = QDoubleValidator()
        self.lineEdit_PS_VIN.setValidator(self.onlyFloat)
        self.lineEdit_PS_VIN_Limit.setValidator(self.onlyFloat)
        self.lineEdit_PS_Cur_Limit.setValidator(self.onlyFloat)
        self.lineEdit_IMAX.setValidator(self.onlyFloat)
        self.lineEdit_ISTEP.setValidator(self.onlyFloat)
        self.lineEdit_ISTART.setValidator(self.onlyFloat)

    # Initial check box status
    def Ext_Supply_used_checked(self, status=True):
        self.comboBox_PS_Address.setEnabled(status)
        self.lineEdit_PS_VIN.setEnabled(status)
        self.lineEdit_PS_VIN_Limit.setEnabled(status)
        self.lineEdit_PS_Cur_Limit.setEnabled(status)
        self.checkBox_ReadIntCur_PSEquip.setEnabled(status)
        self.label_23.setEnabled(status)
        self.checkBox_ReadIntVol_PSEquip.setEnabled(status)
        self.label_19.setEnabled(status)

    # Initial check box status
    def DMM1_used_checked(self):
        if self.checkBox_ReadIntVol_PSEquip.isChecked():
            self.comboBox_DMM_VI.setEnabled(False)
        else:
            self.comboBox_DMM_VI.setEnabled(True)

    # Initial check box status
    def DMM2_used_checked(self):
        if self.checkBox_ReadIntCur_PSEquip.isChecked():
            self.comboBox_DMM_CI.setEnabled(False)
        else:
            self.comboBox_DMM_CI.setEnabled(True)

    # Initial check box status
    def DMM3_used_checked(self):
        if self.checkBox_ReadOutVol_ELoad.isChecked():
            self.comboBox_DMM_VO.setEnabled(False)
        else:
            self.comboBox_DMM_VO.setEnabled(True)

    # Initial check box status
    def DMM4_used_checked(self):
        if self.checkBox_ReadOutCur_ELoad.isChecked():
            self.comboBox_DMM_CO.setEnabled(False)
        else:
            self.comboBox_DMM_CO.setEnabled(True)

    # List the equipment list to drop box
    def combobox_equipment_list(self):
        self.comboBox_PS_Address.clear()
        self.comboBox_ELoad_Address.clear()
        self.comboBox_DMM_VI.clear()
        self.comboBox_DMM_CI.clear()
        self.comboBox_DMM_VO.clear()
        self.comboBox_DMM_CO.clear()
        self.scan_equipment_list()

        if self.comboBox_PS_Address.currentIndex() < 0:
            self.comboBox_PS_Address.addItem("Please select the targeted instrument")
            self.comboBox_PS_Address.setCurrentIndex(0)

        if self.comboBox_ELoad_Address.currentIndex() < 0:
            self.comboBox_ELoad_Address.addItem("Please select the targeted instrument")
            self.comboBox_ELoad_Address.setCurrentIndex(0)

        if self.comboBox_DMM_VI.currentIndex() < 0:
            self.comboBox_DMM_VI.addItem("Please select the targeted instrument")
            self.comboBox_DMM_VI.setCurrentIndex(0)

        if self.comboBox_DMM_CI.currentIndex() < 0:
            self.comboBox_DMM_CI.addItem("Please select the targeted instrument")
            self.comboBox_DMM_CI.setCurrentIndex(0)

        if self.comboBox_DMM_VO.currentIndex() < 0:
            self.comboBox_DMM_VO.addItem("Please select the targeted instrument")
            self.comboBox_DMM_VO.setCurrentIndex(0)

        if self.comboBox_DMM_CO.currentIndex() < 0:
            self.comboBox_DMM_CO.addItem("Please select the targeted instrument")
            self.comboBox_DMM_CO.setCurrentIndex(0)

        for equipment in self.equipment_list:

            # To filter USB connection
            pattern_USB = re.compile("USB[0-9]")
            if pattern_USB.match(equipment.split("::", 4)[0]):
                # Filter Kikusui PBZ20-20 equipment for PS combobox
                if equipment.split("::", 4)[1] == "0x0B3E" and equipment.split("::", 4)[2] == "0x1012":
                    self.comboBox_PS_Address.addItem(equipment)
                # Add the remaining list to all the other DMM combobox
                else:
                    self.comboBox_DMM_VI.addItem(equipment)
                    self.comboBox_DMM_CI.addItem(equipment)
                    self.comboBox_DMM_VO.addItem(equipment)
                    self.comboBox_DMM_CO.addItem(equipment)

            # To filter non-USB connection
            pattern_ASRL = re.compile("ASRL[0-9]")
            if pattern_ASRL.match(equipment.split("::", 4)[0]):
                # Filter ELoad for ELoad combobox
                if equipment.split("::", 4)[1] == "INSTR":
                    self.comboBox_ELoad_Address.addItem(equipment)

    # Scan the equipment list
    def scan_equipment_list(self):
        # self.equipment_list = self.list_instrument.list_connected_devices()
        # To be deleted when actual instrument was used.
        self.equipment_list = ('USB0::0x0B3E::0x1012::XF001773::0::INSTR',
                               'ASRL4::INSTR', 'ASRL8::INSTR',
                               'USB0::2391::1543::MY53020107::INSTR',
                               'USB0::0x2A8D::0x1301::MY53218004::0::INSTR',
                               'USB0::0x0699::0x0408::C014709::0::INSTR',
                               'USB0::0x0AAD::0x0197::1329.7002k44-320094::0::INSTR')

    # Start Test
    def start_test(self):

        # Check input parameter
        self.check_input_parameter()

        # --------------------------------------------------------------------------------------------------------
        # Check existing thread
        if not threading.activeCount() > 1:

            global thread_stop
            thread_stop = "FALSE"

            if self.error_count <= 0:
                self.eff_meas = Eff_Measurement(PS_USED=self.checkBox_ExtSupUsed.isChecked(), PS_ADD=self.PS_address,
                                                PS_VSTART=self.PS_target_Vin, PS_VMAX=self.PS_limit_Vin,
                                                PS_IMAX=self.PS_limit_Cin, ELOAD_ADD=self.ELoad_address,
                                                ELOAD_START=self.ELoad_ISTART, ELOAD_MAX=self.ELoad_IMAX,
                                                ELOAD_STEP=self.ELoad_ISTEP,
                                                DMM_VIN_XUSED=self.checkBox_ReadIntVol_PSEquip.isChecked(),
                                                DMM_VIN_ADD=self.DMM_VIN_address,
                                                DMM_CIN_XUSED=self.checkBox_ReadIntCur_PSEquip.isChecked(),
                                                DMM_CIN_ADD=self.DMM_CIN_address,
                                                DMM_VOUT_XUSED=self.checkBox_ReadOutVol_ELoad.isChecked(),
                                                DMM_VOUT_ADD=self.DMM_VOUT_address,
                                                DMM_COUT_XUSED=self.checkBox_ReadOutCur_ELoad.isChecked(),
                                                DMM_COUT_ADD=self.DMM_COUT_address, OUTPUT_DIR=self.output_directory,
                                                OUTPUT_NAME=self.output_filename,
                                                EXP_PDF=self.checkBox_ExpToPdf.isChecked())

                # self.measurement_thread = QThread()
                # self.eff_meas.moveToThread(self.measurement_thread)
                # self.measurement_thread.started.connect(self.eff_meas.fixed_vin_meas)
                # self.eff_meas.finished.connect(self.measurement_thread.quit)
                # self.eff_meas.finished.connect(self.eff_meas.deleteLater)
                # self.measurement_thread.finished.connect(self.measurement_thread.deleteLater)
                # self.eff_meas.progress.connect(self.reportProgress)
                # self.measurement_thread.start()
                self.measurement_thread = Thread(target=self.eff_meas.fixed_vin_meas)
                self.measurement_thread.start()

            else:
                msg_box_ok("Please fill up all the required field")

        # If existing thread, kill it/ wait until current process finish
        else:
            msg_box_ok('Please wait the current measurement finish or click "Abort Test"')

    # Stop Test
    def stop_test(self):
        if threading.activeCount() <= 1:
            msg_box_ok("Info:\n"
                       "- Test is not running.\n"
                       "- No test case to abort!")
        else:
            global thread_stop
            thread_stop = "TRUE"

    # Check and verify input parameter
    def check_input_parameter(self):

        self.error_count = 0

        # ----------------------Get Info from External Power Supply (If selected)-------------------------------------
        if self.checkBox_ExtSupUsed.isChecked():
            # Check PS Address Input
            if self.comboBox_PS_Address.currentIndex() > 0:
                self.PS_address = self.comboBox_PS_Address.currentText()
            else:
                self.error_count += 1

            # Check PS VIN input
            if self.lineEdit_PS_VIN.text() == "":
                self.error_count += 1
                self.lineEdit_PS_VIN.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
            else:
                self.lineEdit_PS_VIN.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
                self.PS_target_Vin = float(self.lineEdit_PS_VIN.text())

            # Check PS VIN limit input
            if self.lineEdit_PS_VIN_Limit.text() == "":
                self.error_count += 1
                self.lineEdit_PS_VIN_Limit.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
            else:
                self.lineEdit_PS_VIN_Limit.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
                self.PS_limit_Vin = float(self.lineEdit_PS_VIN_Limit.text())

            # Check PS IIN limit input
            if self.lineEdit_PS_Cur_Limit.text() == "":
                self.error_count += 1
                self.lineEdit_PS_Cur_Limit.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
            else:
                self.lineEdit_PS_Cur_Limit.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
                self.PS_limit_Cin = float(self.lineEdit_PS_Cur_Limit.text())

        # ---------------------- Get Info from Electronic Load ------------------------------------------------------
        # Check Eload Address Input
        if self.comboBox_ELoad_Address.currentIndex() > 0:
            self.ELoad_address = self.comboBox_ELoad_Address.currentText()
        else:
            self.error_count += 1

        # Check Eload parameter
        if self.lineEdit_IMAX.text() == "":
            self.error_count += 1
            self.lineEdit_IMAX.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
        else:
            self.lineEdit_IMAX.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
            self.ELoad_IMAX = float(self.lineEdit_IMAX.text())

        if self.lineEdit_ISTEP.text() == "":
            self.error_count += 1
            self.lineEdit_ISTEP.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
        else:
            self.lineEdit_ISTEP.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
            self.ELoad_ISTEP = float(self.lineEdit_ISTEP.text())

        if self.lineEdit_ISTART.text() == "":
            self.error_count += 1
            self.lineEdit_ISTART.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
        else:
            self.lineEdit_ISTART.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
            self.ELoad_ISTART = float(self.lineEdit_ISTART.text())

        # ---------------------- Get Info from DMMs -------------------------------------
        # DMM1
        if not self.checkBox_ReadIntVol_PSEquip.isChecked():
            if self.comboBox_DMM_VI.currentIndex() > 0:
                self.DMM_VIN_address = self.comboBox_DMM_VI.currentText()
            else:
                self.error_count += 1

        # DMM2
        if not self.checkBox_ReadIntCur_PSEquip.isChecked():
            if self.comboBox_DMM_CI.currentIndex() > 0:
                self.DMM_CIN_address = self.comboBox_DMM_CI.currentText()
            else:
                self.error_count += 1

        # DMM3
        if not self.checkBox_ReadOutVol_ELoad.isChecked():
            if self.comboBox_DMM_VO.currentIndex() > 0:
                self.DMM_VOUT_address = self.comboBox_DMM_VO.currentText()
            else:
                self.error_count += 1

        # DMM4
        if not self.checkBox_ReadOutCur_ELoad.isChecked():
            if self.comboBox_DMM_CO.currentIndex() > 0:
                self.DMM_COUT_address = self.comboBox_DMM_CO.currentText()
            else:
                self.error_count += 1

        # ---------------------- Get info for output directory -------------------------------------
        if self.lineEdit_Out_Directory.text() == "":
            self.error_count += 1
            self.lineEdit_Out_Directory.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
        else:
            self.lineEdit_Out_Directory.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
            self.output_directory = self.lineEdit_Out_Directory.text()

        if self.lineEdit_Out_FileName.text() == "":
            self.error_count += 1
            self.lineEdit_Out_FileName.setStyleSheet("QLineEdit{background-color: rgb(255, 10, 10);}")
        else:
            self.lineEdit_Out_FileName.setStyleSheet("QLineEdit{background-color: rgb(255, 255, 255);}")
            self.output_filename = self.lineEdit_Out_FileName.text()


class Eff_Measurement(QObject):
    def __init__(self, PS_USED=False, PS_ADD="", PS_VSTART=0, PS_VMAX=0, PS_VSTEP=0,PS_IMAX=0,
                 ELOAD_ADD="", ELOAD_START=0, ELOAD_MAX=0, ELOAD_STEP=0,
                 DMM_VIN_XUSED=False, DMM_VIN_ADD="",
                 DMM_CIN_XUSED=False, DMM_CIN_ADD="",
                 DMM_VOUT_XUSED=False, DMM_VOUT_ADD="",
                 DMM_COUT_XUSED=False, DMM_COUT_ADD="",
                 OUTPUT_DIR="", OUTPUT_NAME="", EXP_PDF=True):
        super().__init__()

        self.ps_used = PS_USED
        self.ps_add = PS_ADD
        self.ps_vstart = PS_VSTART
        self.ps_vmax = PS_VMAX
        self.ps_vstep = PS_VSTEP
        self.ps_imax = PS_IMAX
        self.eload_add = ELOAD_ADD
        self.eload_istart = ELOAD_START
        self.eload_imax = ELOAD_MAX
        self.eload_istep = ELOAD_STEP
        self.dmm_vin_xused = DMM_VIN_XUSED
        self.dmm_vin_add = DMM_VIN_ADD
        self.dmm_cin_xused = DMM_CIN_XUSED
        self.dmm_cin_add = DMM_CIN_ADD
        self.dmm_vout_xused = DMM_VOUT_XUSED
        self.dmm_vout_add = DMM_VOUT_ADD
        self.dmm_cout_xused = DMM_COUT_XUSED
        self.dmm_cout_add = DMM_COUT_ADD
        self.output_dir = OUTPUT_DIR
        self.output_name = OUTPUT_NAME
        self.exp_pdf = EXP_PDF

        self.vin_measured = []
        self.iin_measured = []
        self.vout_measured = []
        self.iout_measured = []
        self.pin_calculated = []
        self.pout_calculated = []
        self.eff_calculated = []

        self.ext_power_supply = PS_Kikusui_PyVisa.Kikusui_features()  # Kikusui External Power Supply
        self.eload_command = Eload_Chroma_PyVisa.ELOAD_Chroma_features()    # Chroma ELOAD

    def fixed_vin_meas(self):
        print("crash here")
        msg_box_ok("Print")
        error = 0
        error_msg = []

        # connect ELOAD remotely
        self.eload_command.connect_equipment(target_resource_instr=self.eload_add)

        # Count the required loops required
        self.eload_steps = (self.eload_imax - self.eload_istart) / self.eload_istep
        self.eload_steps_round = math.floor(self.eload_steps)

        if self.eload_steps > self.eload_steps_round:
            self.eload_steps_round += 1

        # Turn on external power supply if required
        if self.ps_used:
            self.ext_power_supply.connect_equipment(self.ps_add)
            self.ext_power_supply.set_cv_output(mode='CV', polar='UNIP', out_vol=self.ps_vstart, out_vol_lim=self.ps_vmax, out_cur_lim=self.ps_imax)
            self.status = self.ext_power_supply.on_off_equipment(1)
            time.sleep(3)

            if self.status:
                print(f"Turn on external power supply {self.ps_add}")
            else:
                msg_box_ok(f"Error - function UI-> Eff_measurement -> fixed vin_meas: \n\n"
                           f"Failed to turn on external power supply {self.ps_add}")

        # Start looping test
        for i in range(self.eload_steps_round):

            if thread_stop == "FALSE":

                # Configure eload current
                eload_current = self.eload_istart + (self.eload_istep * i)
                eload_set_status = self.eload_command.static_load(1, eload_current, "ON")

                # If error on ELOAD
                if not eload_set_status:
                    error += 1
                    error_msg.append("Unable to configure ELOAD")
                    error_msg.append("asdaaca")
                    print("ELoad not working properly")
                    break

                time.sleep(3)   # For load current to stabilize
                print(eload_current)
                print(i)
                # # Configure measuring devices
                # if not self.dmm_vin_xused:
                #     print("DMM_VIN_Used")
                # else:
                #     print("PS_VIN_USED")

            else:
                print("Thread Stop")
                break

        if error:
            pass













