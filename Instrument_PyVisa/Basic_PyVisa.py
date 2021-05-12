import pyvisa
import time


# All the basic PyVISA command should be placed here.
class Basic_PyVisa:
    def __init__(self):
        self.rm = pyvisa.ResourceManager()

    def list_connected_devices(self):
        try:
            resource_list = self.rm.list_resources()

        except:
            resource_list = ""

        return resource_list


if __name__ == "__main__":
    rm = Basic_PyVisa()
    rm.connect_device('USB0::0x0B3E::0x1012::XF001773::0::INSTR')
    rm.disconnect_device()




