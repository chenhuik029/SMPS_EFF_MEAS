from Instrument_PyVisa.Basic_PyVisa import Basic_PyVisa
import time


# Only Chroma ELOAD related PyVISA command
class ELoad_Chroma_Core_PyVisa(Basic_PyVisa):
    def __init__(self):
        super().__init__()

    def connect_equipment(self, target_resource_instr):
        try:
            self.Eload_inst = self.rm.open_resource(target_resource_instr)
            print(f'Connected -->  {self.Eload_inst.query("*IDN?")}')
            return True

        # except pyvisa.VisaIOError as error:
        except:
            print("Error 002: \n"
                  "- Incorrect equipment used/ no connection detection.\n"
                  "- Please make sure the equipment is properly connected")
            return False

    def disconnect_device(self):
        try:
            self.Eload_inst.close()
            print('Eload session closed')
        except:
            print("Error 003: Unable to close ELOAD session")

    def display_session(self):
        return self.Eload_inst.session

    def config_channel(self, channel=1):
        try:
            self.Eload_inst.write(f"CHAN {channel}")
            return True
        except:
            print(f"Unable to configure ELOAD channel {channel}")
            return False

    def config_voltage(self, voltage=0):
        try:
            self.Eload_inst.write(f"CONF:VOLT:ON {voltage}")
            return True
        except:
            print(f"Unable to configure ELOAD {voltage}")
            return False

    def config_current(self, current=0):
        try:
            self.Eload_inst.write(f"CURR:STAT:L1 {current}")
            self.Eload_inst.write(f"CURR:STAT:L2 {current}")
            return True
        except:
            print(f"Unable to configure ELOAD {current}")
            return False

    def config_remote(self, remote="ON"):
        try:
            self.Eload_inst.write(f"CONF:REM {remote}")
            return True
        except:
            print(f"Unable to configure ELOAD remote")
            return False

    def config_onoff(self, status="ON"):
        try:
            self.Eload_inst.write(f"LOAD {status}")
            return True
        except:
            print(f"Unable to configure ELOAD ON/OFF")
            return False

    def current_read(self):
        try:
            read_curr = self.Eload_inst.query_ascii_values(f"MEAS:CURR?")
            return read_curr
        except:
            print(f"Unable to configure read ELOAD current")
            return 999

    def voltage_read(self):
        try:
            voltage_curr = self.Eload_inst.query_ascii_values(f"MEAS:VOLT?")
            return voltage_curr
        except:
            print(f"Unable to configure read ELOAD current")
            return 999


class ELOAD_Chroma_features(ELoad_Chroma_Core_PyVisa):
    def __init__(self):
        super().__init__()

    def static_load(self, channel, load_current, on_off):
        channel = self.config_channel(channel=channel)
        load_current = self.config_current(current=load_current)
        on_off = self.config_onoff(status=on_off)
        return channel and load_current and on_off


if __name__ == "__main__":
    eload = ELOAD_Chroma_features()
    eload.connect_equipment('ASRL8::INSTR')
    eload.static_load(3, 0.3, 1)
    time.sleep(2)
    eload.static_load(3, 0.3, 0)

