import winreg
import subprocess
import sys

REG_PATH = r'SOFTWARE\Nefarius Software Solutions e.U.\HidHide'

'''
Utilities for handling the HIDHide program
'''

QUADSTICK_VENDOR_ID_OLD = 0x1fc9
QUADSTICK_PRODUCT_ID_OLD = 0x205B

QUADSTICK_VENDOR_ID = 0x16D0
QUADSTICK_PRODUCT_ID = 0x092B

HORI_VENDOR_ID = 0x0F0D
HORI_PRODUCT_ID = 0x0066

class HIDHide(object):
    def __init__(self, mainWindow=None):
        # find HIDHide program location
        if mainWindow:
            self._log = mainWindow.SendConsoleMessage
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, REG_PATH, 0, winreg.KEY_READ)
            self.H_path = winreg.QueryValueEx(key, 'Path')[0]
            print(self.H_path) # 'C:\\Program Files\\Nefarius Software Solutions\\HidHide\\'
            self._log("HIDHide found")
            winreg.CloseKey(key)
        except Exception as e:
            print('HIDHide init exception ' + repr(e))
            self.H_path = None
            # self._log("HIDHide not found")
        if self.H_path:
            self.H_path = self.H_path + "x64\\HidHideCLI.exe"

    def check_for_quadstick_registration(self):
        try:
            app_list = subprocess.check_output([self.H_path, "--app-list"])
            if app_list.find(b'QuadStick') > 0:
                self._log("QC is registered with HIDHide OK")
                return True
            # try to register QC
            # get QC path
            my_path = str(sys.executable)
            print (my_path)
            # register QC with HIDHide
            print (subprocess.check_output([self.H_path, "--app-reg", my_path]))
            self._log("QC has registered with HIDHide OK")

            return True
        except Exception as e:
            if self.H_path:
                self._log("Unable to check HIDHide registration. Please close the HIDHide Configuration Client if it is open")
            print ('HIDHide check_for_quadstick_registration exception '+ repr(e))
            return False

    def is_installed(self):
        if self.H_path:
            return True
        return False

    def _get_quadstick_usb_path(self, qs):
        path = qs.get_path() # looks like '\\\\?\\hid#vid_16d0&pid_092b&mi_00#8&27ea8d27&0&0000#{4d1e55b2-f16f-11cf-88cb-001111000030}'
        print (path)
        # convert it to something like this: "HID\VID_16D0&PID_092B&MI_00\7&1c5044e4&0&0000"
        a, b = path.split('16d0')
        c,d,e = b.split("#")
        new_path = "HID\\VID_16D0" + c.upper() + "\\" + d
        print (new_path)
        return new_path

    def hide_quadstick(self, qs):
        if self.H_path:
            try:
                new_path = self._get_quadstick_usb_path(qs)
                print(subprocess.check_output([self.H_path, "--dev-hide", new_path]))
                print(subprocess.check_output([self.H_path, "--cloak-on"]))
                self._log("QuadStick hidden with HIDHide")
            except:
                self._log("Unable to hide QuadStick with HIDHide. Please close the HIDHide Configuration Client if it is open")

    def unhide_quadstick(self, qs):
        if self.H_path:
            try:
                new_path = self._get_quadstick_usb_path(qs)
                print(subprocess.check_output([self.H_path, "--dev-unhide", new_path]))
                self._log("QuadStick un-hidden with HIDHide")
            except:
                self._log("Unable to un-hide QuadStick with HIDHide. Please close the HIDHide Configuration Client if it is open")

    def is_hidden(self, qs):
        if self.H_path:
            new_path = self._get_quadstick_usb_path(qs)
            paths = subprocess.check_output([self.H_path, "--dev-list"])
            print (new_path, paths.decode())
            return paths.decode().find(new_path) > -1
        return False
