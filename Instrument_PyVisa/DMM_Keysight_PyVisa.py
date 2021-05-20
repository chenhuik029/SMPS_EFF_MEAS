from Instrument_PyVisa.Basic_PyVisa import Basic_PyVisa


# Only Keysight DMM specified PyVISA command place here.
class Keysight_DMM(Basic_PyVisa):
    def __init__(self):
        super().__init__()

    def connect_equipment(self, target_resource_instr):
        try:
            self.DMM_inst = self.rm.open_resource(target_resource_instr)
            print(f'Connected --> {self.DMM_inst.query("*IDN?")}')
            return True

        except:
            print("Error 002: \n"
                  "- Incorrect equipment used/ no connection detection.\n"
                  "- Please make sure the equipment is properly connected")
            return False

    def disconnect_device(self):
        try:
            self.DMM_inst.close()
            print('DMM session closed')
        except:
            print("Error 003: Unable to close DMM session")

    def display_session(self):
        return self.DMM_inst.session

    def meas_vdc(self):
        try:
            read_vdc = self.DMM_inst.query_ascii_values("MEAS:VOLT:DC?")
            return read_vdc
        except:
            print(f'Unable to read VDC')
            return [999]

    def meas_idc(self):
        try:
            read_idc = self.DMM_inst.query_ascii_values("MEAS:CURR:DC?")
            return read_idc
        except:
            print(f'Unable to read IDC')
            return [999]


if __name__ == "__main__":
    dmm = Keysight_DMM()
    dmm.connect_equipment("USB0::0x0957::0x1C07::MY53206078::0::INSTR")
    print(dmm.meas_vdc()[0])
