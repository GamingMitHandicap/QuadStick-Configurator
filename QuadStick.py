import socket
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox

import sys
import os
import platform
import subprocess
import shutil
from configparser import ConfigParser

import xlsx2csv
import qsflash
import HIDHide

from qsflash import *
from vocola import *
from microterm import microterm, has_serial_ports
from ViGEmBus import VirtualGamepadEmulator
from QuadStickHID import *
from ultrastik import *
#from trackir import *
from googledrive import *
from textstrings import *
from CTkToolTip import *

if platform.system() == "Windows":
    import winsound

cparser = ConfigParser(interpolation=None)
settings = qsflash.settings
app_title = "QuadStick Configurator"

global window_counter
window_counter= 0

global current_locale
current_locale = "en"

global current_appearance_mode
current_appearance_mode = "System"

global enable_log
enable_log = False

# Global variables for devices
VG = None
QS = None
US1 = None
US2 = None
TIR = None
MOUSE = None
H = None # HIDHide handler

SERIAL_PORT_SOCKET = None

# Set custom color theme
# https://github.com/TomSchimansky/CustomTkinter/wiki/Themes
ctk.set_default_color_theme("themes/theme.json") # TODO Edit color theme

# Deactivate HighDPI support
ctk.deactivate_automatic_dpi_awareness()

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def update_column_width(treeview, col, min_width=50):
    """ Set dynamic column width by text length for treeview """
    font = ctk.CTkFont(family="Arial", size=16)
    # Measure the header text
    max_width = font.measure(col)
    for item in treeview.get_children():
        text = treeview.set(item, col)
        max_width = max(max_width, font.measure(text))
    final_width = max(max_width + 10, min_width)
    treeview.column(col, minwidth=final_width)

def GetLocaleText(text):
    """ Get locale text from locale file """
    try:
        locale_text = cparser.get('locale', text)
        return locale_text.replace('\\n', '\n').replace('"', '')
    except Exception as e:
        print("GetLocaleText exception: ", repr(e))
        return text

def bring_window_to_front(self):
    self.lift()
    self.attributes('-topmost', True)
    self.after_idle(self.attributes, '-topmost', False) # Allow other windows to be on top later

    if platform.system() == 'Linux':
        # On Linux, this usually works fine
        self.focus_force()
    elif platform.system() == 'Darwin':
        # macOS may ignore focus_force; use AppleScript as workaround
        try:
            import subprocess
            app_name = "Python" # Might need to change based on your interpreter
            script = f'''
                tell application "{app_name}"
                    activate
                end tell
            '''
            subprocess.run(['osascript', '-e', script])
        except Exception as e:
            print("Could not run AppleScript: ", e)

def SetWindowGeometry(window, width, height):
    # Set window center
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width / 2) - (width / 2)
    y = (screen_height / 2) - (height / 2) - 30

    # Set window geometry
    window.geometry(f"{width}x{height}+{int(x)}+{int(y)}")

def center_window(window):
    window.update_idletasks() # Ensures the window has calculated size
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")

def get_system_theme():
    system = platform.system()

    if system == "Windows":
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                return "Light" if value == 1 else "Dark"
        except:
            return "Dark"

    elif system == "Darwin": # macOS
        try:
            result = subprocess.run(
                ['osascript', '-e', 'tell application "System Events" to tell appearance preferences to get dark mode'],
                capture_output=True, text=True)
            if "true" in result.stdout.lower():
                return "Dark"
        except:
            pass
        return "Dark"

    elif system == "Linux":
        try:
            desktop_env = os.environ.get("XDG_CURRENT_DESKTOP", "").lower()

            if "gnome" in desktop_env:
                result = subprocess.run(
                    ["gsettings", "get", "org.gnome.desktop.interface", "gtk-theme"],
                    capture_output=True, text=True)
                theme = result.stdout.strip().strip("'").lower()
                return "Dark" if "dark" in theme else "Light"

            elif "kde" in desktop_env or "plasma" in desktop_env:
                config_path = os.path.expanduser("~/.config/kdeglobals")
                if not os.path.exists(config_path):
                    return "Light"

                config = cparser
                config.read(config_path)
                color_scheme = config["General"].get("ColorScheme", "")
                return "Dark" if "dark" in color_scheme.lower() else "Light"

            elif "xfce" in desktop_env:
                result = subprocess.run(
                    ["xfconf-query", "-c", "xsettings", "-p", "/Net/ThemeName"],
                    capture_output=True, text=True)
                theme = result.stdout.strip().lower()
                return "Dark" if "dark" in theme else "Light"
        except:
            pass
        return "Dark"

    else:
        return "Dark"

def clear_treeview_rows(tree):
    for item in tree.get_children():
        tree.delete(item)

def on_validate_int(P):
    if P == "":
        return True
    try:
        int(P)
        return True
    except ValueError:
        return False

class ErrorDialog(ctk.CTkToplevel):
    def __init__(self, title=None, message=None, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title(title)
        self.after(200, lambda: self.iconbitmap(resource_path("quadstickx.ico")))
        self.after(10, lambda: center_window(self)) # Delay to allow proper sizing
        self.resizable(width=False, height=False)

        app_font = ctk.CTkFont(family="Arial", size=16)
        app_header_font = ctk.CTkFont(family="Arial", size=16, weight="bold")

        self.label = ctk.CTkLabel(self, text=message, font=app_font, justify="left", wraplength=350)
        self.label.pack(padx=20, pady=20)

        self.btn_close = ctk.CTkButton(self, text=GetLocaleText("OK"), cursor="hand2", height=40, command=self.close, font=app_header_font)
        self.btn_close.pack(side="right", anchor="e", padx=5, pady=5)

        if platform.system() == "Windows":
            winsound.MessageBeep(winsound.MB_ICONHAND) # Play the default Windows error sound

    def close(self):
        self.destroy()

class MessageDialog(ctk.CTkToplevel):
    def __init__(self, title=None, message=None, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title(title)
        self.after(200, lambda: self.iconbitmap(resource_path("quadstickx.ico")))
        self.after(10, lambda: center_window(self)) # Delay to allow proper sizing
        self.resizable(width=False, height=False)

        app_font = ctk.CTkFont(family="Arial", size=16)
        app_header_font = ctk.CTkFont(family="Arial", size=16, weight="bold")

        self.answer = "no"

        self.label = ctk.CTkLabel(self, text=message, font=app_font, justify="left", wraplength=350)
        self.label.pack(padx=20, pady=20)

        self.box_btns = ctk.CTkFrame(self)
        self.box_btns.pack(side="bottom", anchor="e")

        self.btn_cancel = ctk.CTkButton(self.box_btns, text=GetLocaleText("Cancel"), cursor="hand2", height=40, command=self.close, font=app_header_font)
        self.btn_cancel.pack(padx=2, pady=2, side="right")

        self.btn_submit = ctk.CTkButton(self.box_btns, text=GetLocaleText("OK"), cursor="hand2", height=40, command=self.submit, font=app_header_font)
        self.btn_submit.pack(padx=2, pady=2, side="right")

        self.protocol("WM_DELETE_WINDOW", self.close)

    def submit(self):
        self.answer = "yes"
        self.close()

    def close(self):
        self.destroy()
        global window_counter
        window_counter -= 1

    def get(self):
        return self.answer

class UserGoogleDriveFolder(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title(GetLocaleText("Add_User_Game_Configuration_File"))
        self.after(200, lambda: self.iconbitmap(resource_path("quadstickx.ico")))
        SetWindowGeometry(self, 700, 300)
        self.resizable(width=False, height=False)

        app_font = ctk.CTkFont(family="Arial", size=16)
        app_header_font = ctk.CTkFont(family="Arial", size=16, weight="bold")

        self.url = str("")

        self.label = ctk.CTkLabel(self, text=GetLocaleText("Add_User_Game_Configuration_File_Info"), font=app_font, justify="left", wraplength=550)
        self.label.pack(padx=20, pady=20)

        self.textbox = ctk.CTkTextbox(master=self, height=20, width=600, corner_radius=0, activate_scrollbars=False, font=app_font)
        self.textbox.pack(pady=(0, 20))
        self.textbox.focus()

        self.box_btns = ctk.CTkFrame(self)
        self.box_btns.pack(side="bottom", anchor="e")

        self.btn_cancel = ctk.CTkButton(self.box_btns, text=GetLocaleText("Cancel"), cursor="hand2", height=100, command=self.close, font=app_header_font)
        self.btn_cancel.pack(padx=2, pady=2, side="right")

        self.btn_submit = ctk.CTkButton(self.box_btns, text=GetLocaleText("OK"), cursor="hand2", height=100, command=self.submit, font=app_header_font)
        self.btn_submit.pack(padx=2, pady=2, side="right")

        self.protocol("WM_DELETE_WINDOW", self.close)

    def submit(self):
        self.url = self.textbox.get("1.0", "end-1c") # from line 1, char 0 to end minus last newline
        self.close()

    def close(self):
        self.destroy()
        global window_counter
        window_counter -= 1

    def get(self):
        return self.url

class MouseCapture(ctk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title(GetLocaleText("Mouse_capture"))
        self.after(200, lambda: self.iconbitmap(resource_path("quadstickx.ico")))
        SetWindowGeometry(self, 400, 200)
        self.resizable(width=False, height=False)

        # Set transparency (0.0 = fully transparent, 1.0 = fully opaque)
        self.attributes("-alpha", 0.8) # 80% opacity

        app_font = ctk.CTkFont(family="Arial", size=16)

        self.label = ctk.CTkLabel(self, text=GetLocaleText("Mouse_capture_active"), font=app_font, justify="left")
        self.label.pack(padx=20, pady=20)

        self.result = False

        self.protocol("WM_DELETE_WINDOW", self.close)
        self.bind("<Key>", self.KeyDownEvent)

        # setup mouse event variables
        self._mode = settings['mouse_capture_mode']
        # self._center = wx.Point(settings['mouse_center_x'], settings['mouse_center_y'])
        self._width = float(settings['mouse_width'])/2
        self._height = float(settings['mouse_height'])/2 
        self._limit_x = int(self._width * 0.75)
        self._limit_y = int(self._height * 0.75)
        self._gain_x = float(settings['mouse_gain_x'])
        self._gain_y = float(settings['mouse_gain_x'])
        self._delta_x = 0.0
        self._delta_y = 0.0
        self._last_x = 0.0
        self._last_y = 0.0
        self._last_time = time.time()
        self._dxdt = 0
        self._dydt = 0
        self._buttons = (0,0,0,0,0)
        # if self._mode == "Motion":
            #"""Hides the cursor."""
            # self.SetCursor(wx.Cursor(wx.CURSOR_BLANK))

    def close(self):
        self.destroy()
        global window_counter
        window_counter -= 1

    def get(self):
        return self.result

    def KeyDownEvent(self, event):
        # Look for the F10 character to exit capture mode
        value = event.keycode
        # print('KeyDownEvent ', repr(value))
        if value == 121: # F10
            self.result = False
            self.close()
            print("exit mouse capture")

    def MouseEvent(self, event):
        pass
        #print "mouse event: ", repr(event)
        #if event.ButtonDown():
            #print "mouse left button clicked", dir(event)
        # self._buttons = (event.LeftIsDown(),
        #                  event.MiddleIsDown(),
        #                  event.RightIsDown())
        #print "buttons: ", repr(self._buttons)
        # self.update_mouse((event.x, event.y) - self._center)

    # def MouseEventText(self, event):
    #     # the text message in the middle of the screen gets separate events from the rest of the screen
    #     #print "mouse event: ", repr(event)
    #     #if event.ButtonDown():
    #         #print "mouse left button clicked", dir(event)
    #     #if event.Moving() or event.Dragging():
    #     self._buttons = (event.LeftIsDown(),
    #                      event.MiddleIsDown(),
    #                      event.RightIsDown())
    #     #print "buttons: ", repr(self._buttons)
    #     self.update_mouse(((event.GetPosition() + self.message.GetPosition()) - self._center).Get())

    # def update_mouse(self, xy):
    #     if self._mode == "Motion":
    #         now = time.time()
    #         delta_time = now - self._last_time
    #         # print "delta_time: ", delta_time
    #         self._last_time = now
    #         if delta_time <= 0:
    #             return
    #         if self._timer.IsRunning():
    #             self._timer.Stop()
    #         # WarpPointer(self, x, y)
    #         self._delta_x = xy[0] - self._last_x
    #         self._delta_y = xy[1] - self._last_y
    #         # check if close to the edge of the window
    #         if abs(xy[0]) > self._limit_x or abs(xy[1]) > self._limit_y:
    #             # reset the point position to the center
    #             xy = (0,0)
    #             self.WarpPointer(int(self._width), int(self._height))
    #         self._last_x, self._last_y = xy
    #         # low pass filter
    #         self._dxdt = ((self._dxdt * MOUSE_CAPTURE_LOW_PASS_FILTER) + (self._delta_x / delta_time)) / (MOUSE_CAPTURE_LOW_PASS_FILTER + 1)
    #         self._dydt = ((self._dydt * MOUSE_CAPTURE_LOW_PASS_FILTER) + (self._delta_y / delta_time)) / (MOUSE_CAPTURE_LOW_PASS_FILTER + 1)
    #         xy = (self._dxdt, self._dydt)
    #         self._timer.Start(50, True) #one shot, not recurring

    #     if MOUSE:
    #         x = int((xy[0] * self._gain_x) / self._width)
    #         y = int((xy[1] * self._gain_y) / self._height)
    #         x = x if x < 100 else 100
    #         x = x if x > -100 else -100
    #         y = y if y < 100 else 100
    #         y = y if y > -100 else -100
    #         MOUSE.update_location(x, y, self._buttons)

    # def TimerEvent(self, event):
    #     # print "timer expired before a movement was detected.  Update mouse."
    #     self._dxdt = 0.0
    #     self._dydt = 0.0
    #     self.update_mouse((self._last_x,self._last_y))
    #     self._timer.Stop()  # just in case
    #     #if self._timer.IsRunning(): # kill timer and wait for actual change
    #         #self._timer.Stop()

    # def __set_properties(self, event):  # wxGlade: MouseCapture.<event_handler>
    #     self.SetSize((wx.DisplaySize()[0], wx.DisplaySize()[1]))
    #     self.SetWindowStyle(wx.STAY_ON_TOP) # go borderless
    #     self.SetExtraStyle(self.GetExtraStyle() & ~(win32con.WS_EX_DLGMODALFRAME |
    #        win32con.WS_EX_WINDOWEDGE | win32con.WS_EX_CLIENTEDGE | win32con.WS_EX_STATICEDGE))
    #     self.SetRect((0,0,wx.DisplaySize()[0], wx.DisplaySize()[1]))
# end of class MouseCapture

class QuadStickConfigurator(ctk.CTk): 
    def __init__(self):
        super().__init__()

        self.title(app_title)

        if platform.system() == "Windows":
            self.iconbitmap(resource_path("quadstickx.ico"))
        elif platform.system() == "Darwin": # macOS
            self.iconphoto(True, ctk.CTkImage(light_image=resource_path("quadstickx.ico"), dark_image=resource_path("quadstickx.ico"), size=(32, 32)))
        else: # Linux
            self.iconphoto(True, ctk.CTkImage(light_image=resource_path("quadstickx.ico"), dark_image=resource_path("quadstickx.ico"), size=(32, 32)))

        SetWindowGeometry(self, 1200, 700)
        self.resizable(width=True, height=True)

        bring_window_to_front(self)

        if settings.get("start_mimimized") == True:
            self.minimize_window()

        # Add default fonts
        self.app_font = ctk.CTkFont(family="Arial", size=16)
        self.app_header_font = ctk.CTkFont(family="Arial", size=16, weight="bold")

        # Add style
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.set_appearance_mode(settings.get("appearance_mode"))

        # Register the validation command
        self.vcmd = self.register(on_validate_int)

        # Add tabs
        self.tabs = ctk.CTkTabview(self, anchor="w", corner_radius=15)
        self.tabs.pack(side="top", expand=True, fill="both")
        self.tabs._segmented_button.configure(font=self.app_font)

        self.tab_game_files = self.build_game_files_tab(self.tabs.add(GetLocaleText("Game_Files")))
        self.tab_joystick = self.build_joystick_tab(self.tabs.add(GetLocaleText("Joystick")))
        self.tab_misc = self.build_misc_tab(self.tabs.add(GetLocaleText("Misc")))
        self.tab_firmware = self.build_firmware_tab(self.tabs.add(GetLocaleText("Firmware")))
        self.tab_voice_control = self.build_voice_control_tab(self.tabs.add(GetLocaleText("Voice_Control")))
        self.tab_voice_files = self.build_voice_files_tab(self.tabs.add(GetLocaleText("Voice_Files")))
        self.tab_external_pointers = self.build_external_pointers_tab(self.tabs.add(GetLocaleText("External_Pointers")))

        # Add footer
        self.frame_footer = ctk.CTkFrame(self, height=200)
        self.frame_footer.pack(side="top", fill="x", expand=False)

        self.box_messages = ctk.CTkTextbox(self.frame_footer, width=590, font=self.app_font)
        self.box_messages.configure(state="disabled")
        self.box_messages.pack(side="left", expand=True, fill="both")

        self.btn_save = ctk.CTkButton(self.frame_footer, text=GetLocaleText("Save_Preferences_to_QuadStick"), cursor="hand2", font=self.app_header_font, command=self.SavePreferences)
        self.btn_save.pack(padx=2, pady=2, side="left", fill="both")

        self.btn_close = ctk.CTkButton(self.frame_footer, text=GetLocaleText("Close"), cursor="hand2", font=self.app_header_font, command=self.close_window)
        self.btn_close.pack(padx=2, pady=2, side="left", fill="both")

        self.btn_reload = ctk.CTkButton(self.frame_footer, text=GetLocaleText("Reload_Preferences_from_QuadStick"), cursor="hand2", font=self.app_header_font, command=self.ReloadFromQuadStick)
        self.btn_reload.pack(padx=2, pady=2, side="left", fill="both")

        # Store original cursors here
        self.original_cursors = {}

        # Create the Microterm singleton used for voice and other commands
        self.microterm = None

        print("Load initial values")
        if self.load_initial_values():
            print("Initial values loaded")

            self.after(1, self.start_microterm)
            self.SendConsoleMessage("Retrieved " + str(len(self._game_profiles))+ " game files.")

            VG = None
            self.VG = None
            try:
                VG = VirtualGamepadEmulator(self) # Opens the DLL, regardless of the presence of a VG
                VG.DEBUG = DEBUG
                self.VG = VG
                settings['ViGEmBus'] = 'VIGEM_ERROR_NONE'
                try: # Set up HIDHide to allow QMP to see the Quadstick
                    H = HIDHide.HIDHide(self)
                    settings['HIDHide_registered'] = str(H.check_for_quadstick_registration())
                    settings['HIDHide_path'] = str(H.H_path)
                except Exception as e:
                    print ('HIDHide init error: ' + repr(e))
                print('ViGEmBus OK')
            except Exception as e:
                print(repr(e))
                settings['ViGEmBus'] = str(e)
                self.SendConsoleMessage('ViGEmBus driver not present')
                # self.checkbox_enable_vg_X360.Disable()
                # self.checkbox_enable_vg_DS4.Disable()

        # Check if minimize_to_tray is enabled
        if settings.get("minimize_to_tray") == True:
            self.protocol("WM_DELETE_WINDOW", self.minimize_window) # Override the close button behavior
        else:
            self.protocol("WM_DELETE_WINDOW", self.close_window)

    def minimize_window(self):
        self.iconify() # Minimizes to taskbar instead of closing
        self.iconbitmap(resource_path("quadstickx.ico")) # Set icon again to use custom icon on minimized window
        self.update() # Force update to ensure it shows in taskbar

    def close_window(self):
        settings['start_mimimized'] = self.checkbox_start_minimized_state.get()
        save_repr_file(settings)
        self.destroy()

    def SendConsoleMessage(self, msg):
        self.box_messages.configure(state="normal")
        self.box_messages.insert("end", msg + "\n")
        self.box_messages.configure(state="disabled")
        self.box_messages.see("end") # Auto-scroll to the end

        global enable_log
        if enable_log: self.Log(msg)

    def Log(self, msg):
        # Create log file if not exists
        with open("log.txt", "a") as log:
            log.write(msg + "\n")

    def set_cursor_all(self, cursor_type):
        """Save and apply cursor to all widgets"""
        self.original_cursors.clear()
        widgets = [self] + self.winfo_children()

        for widget in widgets:
            try:
                orig_cursor = widget.cget("cursor")
                self.original_cursors[widget] = orig_cursor
                widget.configure(cursor=cursor_type)
            except:
                pass
            # Check children of container widgets
            if hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    try:
                        orig_cursor = child.cget("cursor")
                        self.original_cursors[child] = orig_cursor
                        child.configure(cursor=cursor_type)
                    except:
                        pass

    def restore_cursors(self):
        """Restore saved cursors"""
        for widget, cursor in self.original_cursors.items():
            try:
                widget.configure(cursor=cursor)
            except:
                widget.configure(cursor="")
                pass

    def SavePreferences(self):
        print("Event handler 'SavePreferences'")

        # Test
        # self.set_cursor_all("wait")
        # self.after(3000, self.restore_cursors)
        # -----------

        # Update dictionary values
        self.calculate_joystick_preferences()

        # Set up sliders for D_Pad
        preferences['joystick_D_Pad_outer']         = str(self.slider_D_Pad_outer_ring.get())
        preferences['joystick_D_Pad_inner']         = str(self.slider_D_Pad_inner_ring.get())

        # Set up sliders and spinners for Sip/Puff
        preferences['sip_puff_threshold_soft']      = str(self.slider_SP_low.get())
        preferences['sip_puff_threshold']           = str(self.slider_SP_high.get())
        preferences['sip_puff_maximum']             = str(self.slider_SP_max.get())
        preferences['sip_puff_delay_soft']          = str(self.SP_low_delay_value.get())
        preferences['sip_puff_delay_hard']          = str(self.SP_high_delay_value.get())

        # Set up sliders for Lip
        preferences['lip_position_maximum']         = str(self.slider_Lip_max.get())
        preferences['lip_position_minimum']         = str(self.slider_Lip_min.get())

        # Set up mouse, volume and brightness
        preferences['mouse_speed']                  = str(self.slider_mouse_speed.get())
        preferences['brightness']                   = str(self.slider_brightness.get())
        preferences['volume']                       = str(self.slider_volume.get())

        # Set up digital outputs
        preferences['digital_out_1']                = str(self.checkbox_do_1_state.get())
        preferences['digital_out_2']                = str(self.checkbox_do_2_state.get())

        # Set up bluetooth
        preferences['bluetooth_device_mode']        = self.GetDeviceMode(self.choice_BT_device_mode.get())
        preferences['bluetooth_authentication_mode']= self.choice_BT_auth_mode.get()
        preferences['bluetooth_connection_mode']    = 'pair'
        preferences['bluetooth_remote_address']     = ''

        # Set up misc
        preferences['enable_select_files']          = str(self.checkbox_select_files_state.get())
        preferences['enable_swap_inputs']           = str(self.checkbox_swap_state.get())
        preferences['joystick_dead_zone_shape']     = str(self.checkbox_circular_deadzone_state.get())
        preferences['enable_auto_zero']             = '0'
        preferences['mouse_response_curve']         = str([GetLocaleText('Linear'), GetLocaleText('Mixed'), GetLocaleText('Parabolic')].index(self.choice_mouse_response.get()))
        preferences['enable_usb_comm']              = str(self.checkbox_usb_comm_state.get())
        preferences['enable_usb_a_host']            = str(self.checkbox_usb_A_host_mode_state.get())

        print(repr(preferences))

        if save_preferences_file(preferences) is None:
            dialog = ErrorDialog(title=GetLocaleText("Unable_to_save"), message=GetLocaleText("Unable_to_save_Description"))
            dialog.wm_transient(self) # Set window on top
            dialog.focus_force() # focus window
            self.SendConsoleMessage("Failed to save preferences")
        else: # Update status box
            self.SendConsoleMessage("Preferences saved OK")
        # self.log('saveprefs&' + urllib.parse.urlencode(preferences))

    def ReloadFromQuadStick(self):
        print("Event handler 'ReloadFromQuadStick'")
        if load_preferences_file(self) is not None: # Try both flash and ssp access to prefs files
            # Update status box
            self.SendConsoleMessage("Loaded preferences OK")
            settings["preferences"] = preferences
            self.updateControls()

    def DeleteFromQuadStickEvent(self):
        print("Event handler 'DeleteFromQuadStickEvent'")
        
        if self._last_game_list_selected == self.list_csv_files:
            confirm = MessageDialog(title=GetLocaleText("Please_Confirm"), message=GetLocaleText("Delete_the_selected_file_question"))
            confirm.wm_transient(self) # Set window on top
            self.wait_window(confirm)
            result = confirm.get()

            if result == "yes":
                selection = int(-1)
                selected_item = None

                try:
                    selected_item = self._last_game_list_selected.selection()
                except Exception:
                    return

                if selected_item is not None:
                    selection = int(selected_item[0]) # Get the first selected item's ID

                if selection >= 0:
                    gps = self._csv_files
                    gp = gps[selection]
                    filename = gp["filename"] # Get the filename
                    print(repr(filename))

                    if filename == 'default.csv' or filename == 'prefs.csv':
                        self.SendConsoleMessage("Sorry, cannot remove: " + filename)
                    else:
                        try:
                            d = find_quadstick_drive()
                            if d is None:
                                self.microterm.delete_file(filename)
                            else:
                                pathname = d + filename
                                os.remove(pathname)
                            self.SendConsoleMessage("Removed: " + filename)
                            self._last_game_list_selected = None
                            self.update_quadstick_flash_files_items() # Refresh list
                        except Exception as e:
                            print("DeleteFromQuadStickEvent exception: ", repr(e))
                            self.SendConsoleMessage("Exception while removing: " + filename)

    def DownloadToQuadStickEvent(self):
        print("Event handler 'DownloadToQuadStickEvent'")
        self.DownloadCSVFileEvent()

    def user_game_files_dropped(self, data):
        # Add a user game config file to the list
        # Check for validity
        # self.SetCursor(wx.Cursor(wx.CURSOR_WAIT))
        # try:
        info = xlsx2csv.get_config_profile_info_from_url(data)
        if info:
            self.SendConsoleMessage(f"File {info.get('name')} added to user list")
            gps = settings.get("user_game_profiles", [])
            # Check for pre-existing matching ID
            for gp in gps:
                if gp['id'] == info['id']:
                    # Duplicate
                    self.SendConsoleMessage("File found: " + info.get("name"))
                    if gp.get("name") != info.get("name"):
                        # Rename spreadsheet in list
                        if info.get("name") is not None:
                            self.SendConsoleMessage(f"File renamed from: {gp.get('name')} to: {info.get('name')}")
                            gp['name'] = info['name']
                            self.update_user_game_files_list_items()
                    break
            # Add new id and sort
            else:
                gps.append(info)
            settings["user_game_profiles"] = sorted(gps, key=lambda f: f['name'].lower()) # Update settings with sorted list
        else:
            self.SendConsoleMessage("Error: Google spreadsheet is not publicly shared or published")
        self.update_user_game_files_list_items()
        # finally:
        #     self.SetCursor(wx.Cursor(wx.CURSOR_DEFAULT))

    def DownloadCSVFileEvent(self):
        print("Event handler 'DownloadCSVFileEvent'")

        # Get the user's email address if necessary
        email = settings.get('user_email_address', "")
        if len(email) == 0:
            dlg = ctk.CTkInputDialog(title=GetLocaleText('Google_Spreadsheets_Account'), text=GetLocaleText('Enter_Your_Email_Address'))
            dlg_input = dlg.get_input() # Waits for input
            if dlg_input is not None and dlg_input != "":
                print(dlg_input)
                settings['user_email_address'] = dlg_input

        # Handle user list download
        item = int(-1)
        selected_item = None

        try:
            selected_item = self._last_game_list_selected.selection()
        except Exception:
            return

        if selected_item is not None:
            item = int(selected_item[0]) # Get the first selected item's ID

            if item >= 0:
                print("DownloadCSVFileEvent item ", repr(item))
                self.config(cursor="watch") # Set to wait cursor
                self.update()

                try:
                    id = ''
                    print("selected item: ", item)

                    if self._last_game_list_selected == self.list_user_game_files:
                        print("download user game")
                        gps = settings.get("user_game_profiles", [])
                    elif self._last_game_list_selected == self.list_csv_files:
                        gps = self._csv_files
                    else:
                        print("download factory game")
                        gps = self._game_profiles

                    gp = gps[item]
                    id = gp["id"]
                    d = find_quadstick_drive()
                    print("download csv: ", id, d)

                    if xlsx2csv.write_csv_file_for(id, d, self): # download and copy csv into quadstick
                        if self._last_game_list_selected == self.list_user_game_files:
                            info, wb = xlsx2csv.get_config_profile_info(id)
                            if info: # if the csv filename changed, update user list
                                if gp.get("name") != info.get("name"):
                                    gp["name"] = info["name"]
                                    self.update_user_game_files_list_items()
                        self.SendConsoleMessage(f"Downloaded {gp['name']} into QuadStick")
                        app.update() # This allows the GUI to refresh and process events

                    # Refresh list
                    self.update_quadstick_flash_files_items()
                except Exception as e:
                    print(repr(e))
                    return
                finally:
                    self.config(cursor="") # Reset cursor
                    self.update()

    def DownloadFirmwareEvent(self):
        global QuadStickDrive
        global QS
        import shutil
        import tempfile
        import win32file
        from zipfile import ZipFile
        print("Event handler 'DownloadFirmwareEvent'")

        item = int(-1)
        selected_item = None

        try:
            selected_item = self.list_ctrl_firmware.selection()
        except Exception:
            return

        if selected_item is not None:
            item = int(selected_item[0]) # Get the first selected item's ID
            print("item ", repr(item))

        if item >= 0:
            bld_version = self._builds[item]["version"]
            if self.build_number == int(bld_version):
                self.SendConsoleMessage("Sorry, selected build is already installed in QuadStick")
                return

            # Find joystick.bin file and download it
            # TODO self.SetCursor(wx.Cursor(wx.CURSOR_WAIT))
            self.SendConsoleMessage("Download new firmware file. Please wait...")
            firmware_image_zip = get_google_drive_file_by_id(self._builds[item]["id"])

            # Save contents of QuadStick
            tmp_folder_path = tempfile.gettempdir() + '\\quad_stick_temporary_files'
            shutil.rmtree(tmp_folder_path, True)
            tmp_folder = os.mkdir(tmp_folder_path)

            # Write zip file to temp folder and unzip it
            with open(tmp_folder_path + "\\Joystick.zip", "wb", 0) as zipFile:
                zipFile.write(firmware_image_zip)
                zipFile.flush()

            # Unzip file
            with ZipFile(tmp_folder_path + "\\Joystick.zip", "r", 0) as zipFile:
                firmware_image = zipFile.read("Joystick.bin")

            # Get quadstick folder
            qs = find_quadstick_drive()

            # Copy csv files to temp folder
            self.SendConsoleMessage("Backup game configuration files")
            csv_file_list = list_quadstick_csv_files(self)

            # Give prefs.csv file special handling if PS4 boot mode enabled
            try:
                if bld_version < "1215": # Old firmware can't work with PS4 boot mode
                    p = load_preferences_file(self)
                    p["enable_DS3_emulation"] = "0" # Make sure DS4 mode off
                    save_preferences_file(p)
            except Exception as e:
                print(repr(e))
            for file in csv_file_list:
                shutil.copyfile(qs + file[0], tmp_folder_path + "\\" + file[0])
                self.SendConsoleMessage(" " + file[0])
                app.update()
            self.SendConsoleMessage("CSV files backed up to: " + tmp_folder_path)
            self.SendConsoleMessage("Write new firmware file to QuadStick")
            handle = win32file.CreateFile(qs + "Joystick.tmp", win32file.GENERIC_WRITE, 0, None, win32file.CREATE_ALWAYS, win32file.FILE_FLAG_WRITE_THROUGH, None)
            print("file handle: ", handle)
            print("image size: ", len(firmware_image))
            win32file.WriteFile(handle, firmware_image)
            win32file.FlushFileBuffers(handle)
            win32file.CloseHandle(handle)
            time.sleep(5)
            try:
                os.remove(qs + "Joystick.bin")
            except:
                pass
            os.rename(qs + "Joystick.tmp",qs + "Joystick.bin")
            # del(fwfile)
            self.SendConsoleMessage("Wait for QuadStick to reboot...")
            app.update()
            if QS:
                QS.close()
                QS = None

            # Wait for quadstick to disappear from drive list
            for sec in range(40):
                time.sleep(1.0)
                app.update()
                print(sec)
                self.SendConsoleMessage(".")
                # Force actual search for QS
                if find_quadstick_drive(True) is None: break
            if find_quadstick_drive(True) is None:
                self.SendConsoleMessage("QuadStick rebooting..")
            else:
                self.SendConsoleMessage("QuadStick reboot not detected")
                messagebox.showerror(app_title, "The QuadStick did not automatically reboot!\n\nIn Windows Explorer 'Eject' the QuadStick drive.\nIf QuadStick does not restart in ten seconds,\nunplug it and plug it back in\nthen click OK")
            self.SendConsoleMessage("Waiting for QuadStick to install new firmware...")
            for i in range(5):
                for sec in range(60):
                    time.sleep(1.0)
                    app.update()
                    print(sec)
                    self.SendConsoleMessage(".")
                    if find_quadstick_drive(): break
                qs = find_quadstick_drive()
                if qs:
                    # Copy csv files back to QuadStick
                    self.SendConsoleMessage("Copy files back")
                    time.sleep(4)
                    for file in csv_file_list:
                        shutil.copyfile(tmp_folder_path + "\\" + file[0], qs + file[0])
                        self.SendConsoleMessage(" " + file[0])
                        app.update()
                    self.SendConsoleMessage("Done!")

                    # Reopen game controller interface
                    try:
                        QS = QuadStickHID(self, self.VG)
                        QS.enable(settings.get('enable_CM', True))
                        updater = None
                        # if CM: updater = CM.update # TODO Remove?
                        QS = QS.open(updater) # None if QS did not open
                        self.QS = QS # Used for checkbox event
                    except Exception as e:
                        print("reopen QS error: ", repr(e))
                    break
                else:
                    self.SendConsoleMessage("Not able to copy CSV files to QuadStick")
                    confirm = MessageDialog(title=GetLocaleText("Error_copying_CSV_files"), message=GetLocaleText("Unable_to_copy_CSV_files"))
                    confirm.wm_transient(self) # Set window on top
                    self.wait_window(confirm)
                    result = confirm.get()
                    # self.SendConsoleMessage(repr(result))
                    if result != "yes":
                        break
            # shutil.rmtree(tmp_folder_path, True)
            self.update_build_number(str(quadstick_drive_serial_number(self)))
            self.update_online_game_files_list_items()
            self.updateControls()
            # TODO self.SetCursor(wx.Cursor(wx.CURSOR_DEFAULT))

    # TODO
    # def on_entry_change(self, parent_value, *args):
        # if parent_value >= 0:
        # print(parent_value.get())
        # return

    def EnableSerialPortEvent(self, event):
        print("Event handler 'EnableSerialPortEvent'")
        widget = event.widget
        flag = widget.get()
        settings['enable_serial_port'] = flag
        save_repr_file(settings)

    def checkbox_minimize_to_tray_event(self):
        if self.checkbox_minimize_to_tray_state.get() == True:
            # Override the close button behavior
            self.protocol("WM_DELETE_WINDOW", self.minimize_window)
        else:
            # Reset the close button behavior
            self.protocol("WM_DELETE_WINDOW", self.close_window)
        
        settings['minimize_to_tray'] = self.checkbox_minimize_to_tray_state.get()

    def checkbox_enable_log_event(self):
        global enable_log
        enable_log = self.checkbox_log_state.get()
        settings['enable_log'] = enable_log

    def import_data(self):
        file_path = filedialog.askopenfilename(filetypes=[("REPR files", "*.repr")], title=GetLocaleText("Select_QMP_File"))
        if file_path:
            print(file_path)

            # TODO Open on first start of the program.

            # Copy to resource_path
            destination = os.path.join("./", "settings.repr") # Rename here
            shutil.copy(file_path, destination)

            read_repr_file() # Reload global settings from file

            print("Copied as: settings.repr")
            self.SendConsoleMessage("Settings successfully imported.")

    def UserGamesAdd(self):
        global window_counter

        if window_counter < 1:
            window_counter += 1

            dialog = UserGoogleDriveFolder(self)
            dialog.wm_transient(self) # Set window on top
            dialog.focus_force() # focus window
            self.wait_window(dialog)
            url = dialog.get()

            if url != "":
                self.user_game_files_dropped(url)

    def UserGamesRemove(self):
        print("Event handler 'UserGamesRemove'")

        if self._last_game_list_selected == self.list_user_game_files:
            item = int(-1)
            selected_item = None

            try:
                selected_item = self._last_game_list_selected.selection()
            except Exception:
                return

            if selected_item is not None:
                item = int(selected_item[0]) # Get the first selected item's ID
                print("item ", repr(item))
 
            if item >= 0:
                print("selected item: ", item)
                gps = settings.get("user_game_profiles", [])
                del gps[item]
                self.update_user_game_files_list_items() # TODO Error Item 0 already exists
                self._last_game_list_selected = None

    def _update_linked_joystick_slider(self, slider, all, vertical, horizontal):
        new_value = min((max((slider.get(), self.slider_NEUTRAL.get() + 5,)), 100,))

        if not hasattr(slider,'_qs_value'):
            slider._qs_value = new_value

        diff = new_value - slider._qs_value
        slider._qs_value = new_value

        if new_value != slider.get():
            slider.set(new_value)

        if diff == 0: return

        links = self.slider_linking_var.get()
        if links == 0:
            for s in all:
                s._qs_value = min((max((s.get() + diff, self.slider_NEUTRAL.get() + 5,)), 100,))
                s.set(s._qs_value)

                if s == self.slider_UP:
                    self.label_slider_UP_value.configure(text=int(s._qs_value))
                elif s == self.slider_DOWN:
                    self.label_slider_DOWN_value.configure(text=int(s._qs_value))
                elif s == self.slider_LEFT:
                    self.label_slider_LEFT_value.configure(text=int(s._qs_value))
                elif s == self.slider_RIGHT:
                    self.label_slider_RIGHT_value.configure(text=int(s._qs_value))

        if links == 1:
            for s in vertical:
                s._qs_value = min((max((s.get() + diff, self.slider_NEUTRAL.get() + 5,)), 100,))
                s.set(s._qs_value)

                if s == self.slider_UP:
                    self.label_slider_UP_value.configure(text=int(s._qs_value))
                elif s == self.slider_DOWN:
                    self.label_slider_DOWN_value.configure(text=int(s._qs_value))

        if links == 2:
            for s in horizontal:
                s._qs_value = min((max((s.get() + diff, self.slider_NEUTRAL.get() + 5,)), 100,))
                s.set(s._qs_value)

                if s == self.slider_LEFT:
                    self.label_slider_LEFT_value.configure(text=int(s._qs_value))
                elif s == self.slider_RIGHT:
                    self.label_slider_RIGHT_value.configure(text=int(s._qs_value))

    def joystick_slider_up_event(self, value):
        self._update_linked_joystick_slider(self.slider_UP,
            [self.slider_LEFT, self.slider_RIGHT, self.slider_DOWN],
            [self.slider_DOWN],
            []
        )
        self.label_slider_UP_value.configure(text=int(value))
        self.update_joystick_preference_grid()

    def joystick_slider_left_event(self, value):
        self._update_linked_joystick_slider(self.slider_LEFT,
            [self.slider_RIGHT, self.slider_UP, self.slider_DOWN],
            [],
            [self.slider_RIGHT]
        )
        self.label_slider_LEFT_value.configure(text=int(value))
        self.update_joystick_preference_grid()

    def joystick_slider_neutral_event(self, value):
        print("Event handler 'slider_NEUTRAL_event'")

        new_value = self.slider_NEUTRAL.get()

        if not hasattr(self.slider_NEUTRAL,'_qs_value'):
            self.slider_NEUTRAL._qs_value = new_value

        diff = new_value - self.slider_NEUTRAL._qs_value
        self.slider_NEUTRAL._qs_value = new_value

        if diff == 0: return

        for s in [self.slider_LEFT, self.slider_RIGHT, self.slider_UP, self.slider_DOWN]:
            if s.get() < new_value + 5:
                s._qs_value = new_value + 5
                s.set(s._qs_value)

        self.label_slider_NEUTRAL_value.configure(text=int(value))
        self.update_joystick_preference_grid()

    def joystick_slider_right_event(self, value):
        self._update_linked_joystick_slider(self.slider_RIGHT,
            [self.slider_LEFT, self.slider_UP, self.slider_DOWN],
            [],
            [self.slider_LEFT]
        )
        self.label_slider_RIGHT_value.configure(text=int(value))
        self.update_joystick_preference_grid()

    def joystick_slider_down_event(self, value):
        self._update_linked_joystick_slider(self.slider_DOWN,
            [self.slider_LEFT, self.slider_UP, self.slider_RIGHT],
            [self.slider_UP],
            []
        )
        self.label_slider_DOWN_value.configure(text=int(value))
        self.update_joystick_preference_grid()

    def slider_mouse_speed_event(self, value):
        self.label_mouse_speed_value.configure(text=int(value))

    def slider_brightness_event(self, value):
        self.label_brightness_value.configure(text=int(value))

    def slider_volume_event(self, value):
        self.label_volume_value.configure(text=int(value))

    def slider_D_Pad_outer_ring_event(self, value):
        self.label_D_Pad_outer_ring_value.configure(text=int(value))

    def slider_D_Pad_inner_ring_event(self, value):
        self.label_D_Pad_inner_ring_value.configure(text=int(value))

    def slider_SP_max_event(self, value):
        self.label_SP_max_value.configure(text=int(value))

    def slider_SP_high_event(self, value):
        self.label_SP_high_value.configure(text=int(value))

    def slider_SP_low_event(self, value):
        self.label_SP_low_value.configure(text=int(value))

    def slider_Lip_max_event(self, value):
        self.label_Lip_max_value.configure(text=int(value))

    def slider_Lip_min_event(self, value):
        self.label_Lip_min_value.configure(text=int(value))

    def calculate_joystick_preferences(self):
        up = self.slider_UP.get()
        down = self.slider_DOWN.get()
        left = self.slider_LEFT.get()
        right = self.slider_RIGHT.get()
        max_joy = max([up,down,left,right])
        preferences['joystick_deflection_maximum']  = str(int(max_joy))
        preferences['deflection_multiplier_up']     = str(int(max_joy * 100 / up))
        preferences['deflection_multiplier_down']   = str(int(max_joy * 100 / down))
        preferences['deflection_multiplier_left']   = str(int(max_joy * 100 / left))
        preferences['deflection_multiplier_right']  = str(int(max_joy * 100 / right))
        preferences['joystick_deflection_minimum']  = str(int(self.slider_NEUTRAL.get()))

    def update_joystick_preference_grid(self):
        self.calculate_joystick_preferences()

        clear_treeview_rows(self.list_joystick_preference)

        self.list_joystick_preference.insert('', 'end', values=('joystick_deflection_maximum', str(preferences['joystick_deflection_maximum'])))
        self.list_joystick_preference.insert('', 'end', values=('joystick_deflection_minimum', str(preferences['joystick_deflection_minimum'])))
        self.list_joystick_preference.insert('', 'end', values=('deflection_multiplier_up', str(preferences['deflection_multiplier_up'])))
        self.list_joystick_preference.insert('', 'end', values=('deflection_multiplier_down', str(preferences['deflection_multiplier_down'])))
        self.list_joystick_preference.insert('', 'end', values=('deflection_multiplier_left', str(preferences['deflection_multiplier_left'])))
        self.list_joystick_preference.insert('', 'end', values=('deflection_multiplier_right', str(preferences['deflection_multiplier_right'])))

    def GetDeviceMode(self, mode):
        if mode == GetLocaleText("Device_None"):
            return "none"
        elif mode == GetLocaleText("Keyboard"):
            return "keyboard"
        elif mode == GetLocaleText("Game_Pad"):
            return "game_pad"
        elif mode == GetLocaleText("Mouse"):
            return "mouse"
        elif mode == GetLocaleText("Combo"):
            return "combo"
        elif mode == GetLocaleText("Joystick"):
            return "joystick"
        elif mode == GetLocaleText("SSP"):
            return "ssp"

    def GetDeviceModeName(self, mode):
        if mode == "none":
            return GetLocaleText("Device_None")
        elif mode == "keyboard":
            return GetLocaleText("Keyboard")
        elif mode == "game_pad":
            return GetLocaleText("Game_Pad")
        elif mode == "mouse":
            return GetLocaleText("Mouse")
        elif mode == "combo":
            return GetLocaleText("Combo")
        elif mode == "joystick":
            return GetLocaleText("Joystick")
        elif mode == "ssp":
            return GetLocaleText("SSP")

    def GetLocales(self):
        locales = []

        for locale in os.listdir(os.path.join('./locales/')):
            if locale.endswith('.ini'):
                locale_name = locale.split(".ini")[0] # Get only the name of file without ending
                locales.append(locale_name)

        return locales

    def GetLocaleNames(self):
        locale_names = []

        for locale in self.GetLocales():
            locale_name = GetLocaleText(locale)
            locale_names.append(locale_name)

        return locale_names

    def change_locale(self, choice):
        # Get the index of locale
        index = self.GetLocaleNames().index(choice)

        global current_locale

        if self.GetLocales()[index] == current_locale:
            return

        # Get locale short name
        current_locale = self.GetLocales()[index]

        # Check if locale exists
        if cparser.read('./locales/' + current_locale + '.INI', 'utf-8'):
            settings['current_locale'] = current_locale

            self.recreate_tabs([
                (GetLocaleText("Game_Files"), self.build_game_files_tab),
                (GetLocaleText("Joystick"), self.build_joystick_tab),
                (GetLocaleText("Misc"), self.build_misc_tab),
                (GetLocaleText("Firmware"), self.build_firmware_tab),
                (GetLocaleText("Voice_Control"), self.build_voice_control_tab),
                (GetLocaleText("Voice_Files"), self.build_voice_files_tab),
                (GetLocaleText("External_Pointers"), self.build_external_pointers_tab),
            ])

            self.after(100, lambda:self.tabs.set(GetLocaleText("Misc")))

            # Update footer button texts
            self.btn_save.configure(text=GetLocaleText("Save_Preferences_to_QuadStick"))
            self.btn_close.configure(text=GetLocaleText("Close"))
            self.btn_reload.configure(text=GetLocaleText("Reload_Preferences_from_QuadStick"))

            self.SendConsoleMessage("Set language to: " + current_locale)
        else:
            self.SendConsoleMessage("Language file not found: " + current_locale)

    def set_appearance_mode(self, appearance_mode: str):
        ctk.set_appearance_mode(appearance_mode)

        if appearance_mode == "Light":
            self.style.configure("Treeview", background="white", foreground="black", fieldbackground="white", font=self.app_font)
            self.style.configure("Treeview.Heading", background="white", foreground="black", font=self.app_header_font)
            self.style.map('Treeview', background=[('selected','#3a7ebf')], foreground=[('selected','white')])
            self.style.map("Treeview.Heading", background=[("active", "gray")], foreground=[('active','white')])

        elif appearance_mode == "Dark":
            self.style.configure("Treeview", background="#1f1f1f", foreground="White", fieldbackground="#1f1f1f", font=self.app_font)
            self.style.configure("Treeview.Heading", background="#181818", foreground="white", font=self.app_header_font)
            self.style.map('Treeview', background=[('selected','#3c3737')], foreground=[('selected','white')])
            self.style.map("Treeview.Heading", background=[("active", "#3c3737")])

        elif appearance_mode == GetLocaleText("System"):
            current_system_theme = get_system_theme()

            if current_system_theme == "Light":
                self.style.configure("Treeview", background="white", foreground="black", fieldbackground="white", font=self.app_font)
                self.style.configure("Treeview.Heading", background="white", foreground="black", font=self.app_header_font)
                self.style.map('Treeview', background=[('selected','#3a7ebf')], foreground=[('selected','white')])
                self.style.map("Treeview.Heading", background=[("active", "gray")], foreground=[('active','white')])

            else: # Dark
                self.style.configure("Treeview", background="#1f1f1f", foreground="White", fieldbackground="#1f1f1f", font=self.app_font)
                self.style.configure("Treeview.Heading", background="#181818", foreground="white", font=self.app_header_font)
                self.style.map('Treeview', background=[('selected','#3c3737')], foreground=[('selected','white')])
                self.style.map("Treeview.Heading", background=[("active", "#3c3737")])

    def GetAppearanceMode(self, appearance_mode):
        if appearance_mode == GetLocaleText("Light"):
            return "Light"
        elif appearance_mode == GetLocaleText("Dark"):
            return "Dark"
        elif appearance_mode == GetLocaleText("System"):
            return "System"

    def change_appearance_mode(self, appearance_mode: str):
        new_appearance_mode = self.GetAppearanceMode(appearance_mode)

        # Check if appearance mode is already selected
        global current_appearance_mode
        if current_appearance_mode == new_appearance_mode:
            return

        current_appearance_mode = new_appearance_mode
        settings['appearance_mode'] = new_appearance_mode

        self.set_appearance_mode(new_appearance_mode)
        self.SendConsoleMessage(f"Set appearance mode to: {new_appearance_mode}")

    def GetCaptureMode(self, capture_mode):
        if capture_mode == GetLocaleText("Off"):
            return "Off"
        elif capture_mode == GetLocaleText("Position"):
            return "Position"
        elif capture_mode == GetLocaleText("Motion"):
            return "Motion"

    def change_capture_mode(self, capture_mode: str):
        print(self.GetCaptureMode(capture_mode))

    def update_quadstick_flash_files_items(self, fast = False):
        index = int(0)
        _csv_files = []
        files = list_quadstick_csv_files(self, fast) # a tuple of (csv, id, name)

        clear_treeview_rows(self.list_csv_files)

        for f in files:
            if f[0] == 'prefs.csv':
                self.list_csv_files.insert('', 'end', str(-1), values=(" ", f[0], f[2]))
            else:
                self.list_csv_files.insert('', 'end', str(index), values=(index, f[0], f[2]))

            index += 1
            _csv_files.append({"filename":f[0], "id":f[1], "name":f[2]})

        self._csv_files = _csv_files
        update_column_width(self.list_csv_files, "#3")
        self.label_csv_files.configure(text=GetLocaleText("In_QuadStick") + f" ({index})")

    def update_online_game_files_list_items(self): # Updates the display widget with the current game list
        games, voices = get_factory_game_and_voice_files()  # get csv and vch/vcl file info from Google
        self._game_profiles = games

        clear_treeview_rows(self.list_online_game_files)

        # try:
        #     self.SendConsoleMessage("Retrieved " + str(len(self._game_profiles))+ " game files")
        # except:
        #     pass

        index = int(0)
        self._game_profiles = sorted(self._game_profiles, key=lambda f: f['name'].lower())

        for f in self._game_profiles:
            # (game_name, folder, path, name, url)
            csv_name = f["csv_name"]
            name = f["name"]
            self.list_online_game_files.insert('', 'end', str(index), values=(csv_name, name))
            index += 1

        # update_column_width(self.list_online_game_files, "#1", 0) # resize column to match new items
        update_column_width(self.list_online_game_files, "#2") # resize column to match new items
        self.label_online_game_files.configure(text=GetLocaleText("QuadStick_Factory_profiles") + f" ({index})")

    def update_user_game_files_list_items(self):
        print("update_user_game_files_list_items")

        index = int(0)
        user_game_profiles = settings.get("user_game_profiles", [])

        clear_treeview_rows(self.list_user_game_files)

        if len(user_game_profiles) == 0:
            # if user game profiles empty, check for old 1.04 version files
            url = settings.get('profile_url', None)  # get old saved url
            # url = """https://googledrive.com/host/0BwJQJADcHggka2htZ0FlM2FMdTQ/"""
            if url: # aha! first time running after an update
                try:
                    files = get_game_profiles(url, self) # get csv and vch/vcl file info from Google
                    print("IMPORT OLD FILES: ", repr(files))
                    _game_profiles = [fn for fn in files if (fn[3].find(".csv") > 0)]
                    for f in _game_profiles:
                        (game_name, folder, path, name, url) = f
                        print("game name: ", game_name)
                        try:
                            if path:
                                text = read_google_drive_file(path, url) # read the csv file
                                if text.find("/spreadsheets/d/") > 0: # if has a spreadsheet key
                                    text = text.split("/spreadsheets/d/")[1]
                                    text = text.split("/")[0]  # we have the isolated ID
                                    info = {"name":game_name, "id": text, "csv_name":name}
                                    user_game_profiles.append(info)
                        except Exception as e:
                            print(repr(e))
                except Exception as e:
                    print(repr(e))
            if len(user_game_profiles) > 0: # something was imported, save list
                user_game_profiles = sorted(user_game_profiles, key=lambda f: f['name'].lower())
                settings["user_game_profiles"] = user_game_profiles
                settings["profile_url"] = None # flag that we have imported already

        # list of dictionary objects with: {"name":name, "id": id, "csv_name":csv_file_name}
        user_game_profiles = sorted(user_game_profiles, key=lambda f: f['name'].lower())
        for gp in user_game_profiles:
            game_name = gp['name']
            name = gp["csv_name"]
            # print("game name: ", game_name, " csv_name: ", name)
            self.list_user_game_files.insert('', 'end', str(index), values=(name, game_name))
            index += 1

        update_column_width(self.list_user_game_files, "#2")
        self.label_user_game_files.configure(text=GetLocaleText("User_Custom_profiles") + f" ({index})")

    def update_build_number(self, build_number):
        self.build_number_text.configure(state="normal")
        self.build_number_text.insert("end", build_number)
        self.build_number_text._textbox.tag_add("text", "0.0", "end") # set text position using tag
        self.build_number_text.configure(state="disabled")
        self.build_number = build_number

    def update_firmware_list(self):
        self._builds = get_firmware_versions()
        print(repr(self._builds))
        self._available_firmware_list = [bld.get("version","ukn") for bld in self._builds]

        clear_treeview_rows(self.list_ctrl_firmware)

        for bld in self._builds:
            build = ("     " + bld.get("version","ukn"))[-4:]
            self.list_ctrl_firmware.insert('', 'end', values=(build, bld.get("comment","")))

        update_column_width(self.list_ctrl_firmware, "#2") # Resize column to match new items

    def get_max_game_name_length(self, my_games): # In order to create fixed width output for games and URLs
        max_game_name_length = 0
        for my_game in my_games:
            if len(my_game) > max_game_name_length:
                max_game_name_length = len(my_game) # Get the maximum field length needed
        return max_game_name_length

    def StartMouseCaptureEvent(self):
        print("Event handler 'StartMouseCaptureEvent'")
        # global window_counter

        # if window_counter < 1:
        #     window_counter += 1

        #     dialog = MouseCapture(self)
        #     dialog.wm_transient(self) # Set window on top
        #     dialog.focus_force() # focus window
        #     self.wait_window(dialog)
        #     result = dialog.get()

    def start_microterm(self):
        t = threading.Thread(target=self._start_microterm)
        t.daemon = True
        t.start()

    def _start_microterm(self):
        global MT
        # self.SendConsoleMessage("Start serial connection search.")
        for i in range(4): # Try up to four times to find com port
            mt = microterm(self)
            if mt:
                self.microterm = mt
                MT = mt
                print("Microterm started")
                if find_quadstick_drive() is None: 
                    print("Load preference file over microterm")
                    load_preferences_file(self)
                    self.after(1, self.updateControls)
                    self.after(1, self.SendConsoleMessage("Loaded preferences OK"))
                return
        # self.SendConsoleMessage("No serial connection to the QuadStick was found")

    def PrintFileListEvent(self):
        print("Event handler 'PrintFileListEvent'")

        answer = [FILE_LIST_HTML_HEADER]

        base_varname = "FILE_LIST_HTML_HEADER"
        if current_locale != "en":
            # Access the variable dynamically by locale name with fallback
            answer = [globals()[f"{base_varname}_{current_locale}"]]

        files = list_quadstick_csv_files(self) 
        index = 1
        for f in files:
            if f[0] == 'prefs.csv': continue
            answer.append(FILE_LIST_COL_0)
            answer.append(LED_PATTERN.get(index, ""))
            answer.append(FILE_LIST_COL_1)
            answer.append(str(index))
            answer.append(FILE_LIST_COL_2)
            answer.append(str(f[0]))
            answer.append(FILE_LIST_COL_3)
            answer.append(str(f[2])) # xlsx2csv.get_name_from_csv(f,d))
            answer.append(FILE_LIST_COL_4)
            index += 1
        answer.append(FILE_LIST_HTML_FOOTER)
        answer = "".join(answer)
        tmp_file_path = xlsx2csv.write_temporary_file("file_list.html", answer)
        xlsx2csv.write_temporary_file("red.svg", RED_DOT, False)
        xlsx2csv.write_temporary_file("blue.svg", BLUE_DOT, False)
        xlsx2csv.write_temporary_file("purple.svg", PURPLE_DOT, False)
        xlsx2csv.write_temporary_file("grey.svg", GREY_DOT, False)
        import webbrowser
        url = """file:///""" + tmp_file_path
        webbrowser.open(url, new=2)

    def LoadAndRunEvent(self):
        print("Event handler 'LoadAndRunEvent'")

        if self._last_game_list_selected == self.list_csv_files:
            selection = int(-1)
            selected_item = None

            try:
                selected_item = self._last_game_list_selected.selection()
            except Exception:
                return

            if selected_item is not None:
                selection = int(selected_item[0]) # Get the first selected item's ID

            if selection >= 0:
                gps = self._csv_files
                gp = gps[selection]
                filename = gp["filename"] # Get the filename
                print(repr(filename))
                # if filename == 'prefs.csv':
                #     self.SendConsoleMessage("Sorry, that is not a game file.")
                #     return
                self.SendConsoleMessage("Load and Run " + filename + " in QuadStick")
                aString = "\rload," + filename + "\r"
                # Make use of existing vocola listening thread to handle this
                sock = socket.socket(socket.AF_INET, # Internet
                                    socket.SOCK_DGRAM) # UDP
                sock.sendto(aString.encode(), (UDP_IP, UDP_PORT))

    def OnEditSpreadsheet(self):
        print("Event handler 'EditGameFileEvent'")

        item = int(-1)
        selected_item = None

        try:
            selected_item = self._last_game_list_selected.selection()
        except Exception:
            return

        if selected_item is not None:
            item = int(selected_item[0]) # Get the first selected item's ID

        if item >= 0:
            print("item ", repr(item))
            id = None

            print("selected item: ", item)

            if self._last_game_list_selected == self.list_user_game_files:
                print("Edit user game")
                gps = settings.get("user_game_profiles", [])
            elif self._last_game_list_selected == self.list_online_game_files: # factory game
                print("Edit factory game")
                gps = self._game_profiles
            elif self._last_game_list_selected == self.list_csv_files:
                print("Edit csv file in quadstick flash")
                gps = self._csv_files

            gp = gps[item]
            id = gp["id"]
            game_name = gp["name"]

        if id:
            url = "https://docs.google.com/spreadsheets/d/" + id + "/"
            import webbrowser
            webbrowser.open(url, new=2)
            self.SendConsoleMessage("Opened: " + game_name)

    def GameListSelected(self, event):
        # save last selected game for Edit or Download button action
        print("Event handler 'GameListSelected'")

        current_tree = event.widget # The Treeview that triggered the event
        selected_item = current_tree.selection()

        # If nothing is selected, don't proceed
        if not selected_item:
            return

        if selected_item:
            self._last_game_list_selected = current_tree
            # print(f"Selected {selected_item}")

            # item_data = self._last_game_list_selected.item(selected_item)
            # print(f"Item Data {item_data}")

        # Save selected item ID (in case it's lost later)
        selected_item_id = selected_item[0]

        def deselect_others():
            all_trees = [self.list_online_game_files, self.list_csv_files, self.list_user_game_files]

            for tree in all_trees:
                if tree != current_tree:
                    selected = tree.selection()
                    if selected:
                        tree.selection_remove(selected)

            # Reapply selection to the current tree if needed
            if selected_item_id not in current_tree.selection():
                self.selection_set(selected_item_id)

        self.after(10, deselect_others)

        # TODO disable buttons
        # if current_tree == self.list_online_game_files:
            # self.button_delete_csv.Disable()
            # self.button_remove_user_game.Disable()
            # self.button_load_and_run.Disable()
            # self.button_download_csv.Enable()

        # elif current_tree == self.list_csv_files:
            # self.button_remove_user_game.Disable()
            # self.button_download_csv.Enable()

        # elif current_tree == self.list_user_game_files:
            # self.button_delete_csv.Disable()
            # self.button_remove_user_game.Enable()
            # self.button_load_and_run.Disable()
            # self.button_download_csv.Enable()

    def double_click_spreadsheet(self, event):
        self.OnEditSpreadsheet()

    def CopyGameListToClipboard(self):
        user_game_profiles = settings.get("user_game_profiles", [])
        self.game_dict = {game['name']: game['id'] for game in user_game_profiles} # Assuming each game is a dictionary
        self.max_game_name_length = self.get_max_game_name_length(self.game_dict.keys()) + 5 # Add spacing for the output
        clipboard_data = ""

        for my_game in self.game_dict.keys():
            clipboard_data += f"{my_game.ljust(self.max_game_name_length, ' ')} https://docs.google.com/spreadsheets/d/{self.game_dict[my_game]}\n"

        # Copy the data to the clipboard
        self.clipboard_clear()
        self.clipboard_append(clipboard_data)

    def build_game_files_tab(self, parent_frame):
        # Flash Drive Profiles
        self.box_quadstick_flash = ctk.CTkFrame(parent_frame)
        self.box_quadstick_flash.pack(side="left", expand=True, fill="both")

        self.box_quadstick_flash_main = ctk.CTkFrame(self.box_quadstick_flash, height=400)
        self.box_quadstick_flash_main.pack(side="top", expand=True, fill="both")
        
        self.label_csv_files = ctk.CTkLabel(self.box_quadstick_flash_main, text=GetLocaleText("In_QuadStick") + f" (0)", font=self.app_header_font)
        self.label_csv_files.pack(side="top")

        self.xscrollbar_csv_files = ctk.CTkScrollbar(self.box_quadstick_flash_main, orientation="horizontal")
        self.xscrollbar_csv_files.pack(side="bottom", fill="x")

        self.yscrollbar_csv_files = ctk.CTkScrollbar(self.box_quadstick_flash_main)
        self.yscrollbar_csv_files.pack(side="right", fill="y")

        self.cols_csv_files = (GetLocaleText("ID"), GetLocaleText("Filename"), GetLocaleText("Spreadsheet"))
        self.list_csv_files = ttk.Treeview(self.box_quadstick_flash_main, show="headings", columns=self.cols_csv_files, height=15, yscrollcommand=self.yscrollbar_csv_files.set, xscrollcommand=self.xscrollbar_csv_files.set)
        self.list_csv_files.heading("#1", text=self.cols_csv_files[0], anchor="w")
        self.list_csv_files.heading("#2", text=self.cols_csv_files[1], anchor="w")
        self.list_csv_files.heading("#3", text=self.cols_csv_files[2], anchor="w")
        self.list_csv_files.column("#1", minwidth=40, width=40)
        self.list_csv_files.column("#2", minwidth=140, width=140)
        self.list_csv_files.column("#3", minwidth=210, width=210)
        self.list_csv_files.pack(expand=True, fill="both")

        # Not work atm
        # self.tooltip_csv_files = CTkToolTip(self.list_csv_files, GetLocaleText("Double_click_to_edit_and_drag_over_to_download"))

        self.xscrollbar_csv_files.configure(command=self.list_csv_files.xview)
        self.yscrollbar_csv_files.configure(command=self.list_csv_files.yview)

        self.box_quadstick_flash_btns_top = ctk.CTkFrame(self.box_quadstick_flash)
        self.box_quadstick_flash_btns_top.pack(side="top", expand=True, fill="both")

        self.btn_load_and_run = ctk.CTkButton(self.box_quadstick_flash_btns_top, text=GetLocaleText("Load_and_Run_File"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.LoadAndRunEvent)
        self.btn_load_and_run.pack(padx=2, pady=2, side="left", expand=True, fill="both")

        self.tooltip_load_and_run = CTkToolTip(self.btn_load_and_run, GetLocaleText("Load_and_Run_File_Description"))

        self.btn_delete_csv = ctk.CTkButton(self.box_quadstick_flash_btns_top, text=GetLocaleText("Remove_from_QuadStick"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.DeleteFromQuadStickEvent)
        self.btn_delete_csv.pack(padx=2, pady=2, side="left", expand=True, fill="both")

        self.tooltip_delete_csv = CTkToolTip(self.btn_delete_csv, GetLocaleText("Remove_from_QuadStick_Description"))

        self.box_quadstick_flash_btns_bottom = ctk.CTkFrame(self.box_quadstick_flash)
        self.box_quadstick_flash_btns_bottom.pack(side="top", expand=True, fill="both")

        self.btn_print_file_list = ctk.CTkButton(self.box_quadstick_flash_btns_bottom, text=GetLocaleText("Print_file_list"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.PrintFileListEvent)
        self.btn_print_file_list.pack(padx=2, pady=2, side="left", expand=True, fill="both")

        self.tooltip_print_file_list = CTkToolTip(self.btn_print_file_list, GetLocaleText("Print_file_list_Description"))

        # Factory Profiles
        self.box_factory_profiles = ctk.CTkFrame(parent_frame)
        self.box_factory_profiles.pack(side="left", expand=True, fill="both")

        self.box_factory_profiles_main = ctk.CTkFrame(self.box_factory_profiles, height=400)
        self.box_factory_profiles_main.pack(side="top", expand=True, fill="both")

        self.label_online_game_files = ctk.CTkLabel(self.box_factory_profiles_main, text=GetLocaleText("QuadStick_Factory_profiles") + f" (0)", font=self.app_header_font)
        self.label_online_game_files.pack(side="top")

        self.xscrollbar_online_game_files = ctk.CTkScrollbar(self.box_factory_profiles_main, height=20, orientation="horizontal")
        self.xscrollbar_online_game_files.pack(side="bottom", fill="x")

        self.yscrollbar_online_game_files = ctk.CTkScrollbar(self.box_factory_profiles_main)
        self.yscrollbar_online_game_files.pack(side="right", fill="y")

        self.cols_online_game_files = (GetLocaleText("Filename"), GetLocaleText("Spreadsheet"))
        self.list_online_game_files = ttk.Treeview(self.box_factory_profiles_main, show="headings", columns=self.cols_online_game_files, height=15, yscrollcommand=self.yscrollbar_online_game_files.set, xscrollcommand=self.xscrollbar_online_game_files.set)
        self.list_online_game_files.heading("#1", text=self.cols_online_game_files[0], anchor="w")
        self.list_online_game_files.heading("#2", text=self.cols_online_game_files[1], anchor="w")
        self.list_online_game_files.column("#1", minwidth=200, width=200)
        self.list_online_game_files.column("#2", minwidth=170, width=170)
        self.list_online_game_files.pack(expand=True, fill="both")

        self.xscrollbar_online_game_files.configure(command=self.list_online_game_files.xview)
        self.yscrollbar_online_game_files.configure(command=self.list_online_game_files.yview)

        self.box_factory_profiles_btns = ctk.CTkFrame(self.box_factory_profiles)
        self.box_factory_profiles_btns.pack(side="top", expand=True, fill="both")

        self.btn_edit_spreadsheet = ctk.CTkButton(self.box_factory_profiles_btns, text=GetLocaleText("Open_Configuration_Spreadsheet"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.OnEditSpreadsheet)
        self.btn_edit_spreadsheet.pack(padx=2, pady=2, side="top", expand=True, fill="both")

        self.tooltip_edit_spreadsheet = CTkToolTip(self.btn_edit_spreadsheet, GetLocaleText("Open_Configuration_Spreadsheet_Description"))

        self.btn_download_csv = ctk.CTkButton(self.box_factory_profiles_btns, text=GetLocaleText("Download_to_QuadStick"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.DownloadToQuadStickEvent)
        self.btn_download_csv.pack(padx=2, pady=2, side="top", expand=True, fill="both")

        self.tooltip_download_csv = CTkToolTip(self.btn_download_csv, GetLocaleText("Download_to_QuadStick_Description"))

        # User Custom Profiles
        self.box_custom_profiles = ctk.CTkFrame(parent_frame)
        self.box_custom_profiles.pack(side="left", expand=True, fill="both")

        self.box_custom_profiles_main = ctk.CTkFrame(self.box_custom_profiles, height=400)
        self.box_custom_profiles_main.pack(side="top", expand=True, fill="both")

        self.label_user_game_files = ctk.CTkLabel(self.box_custom_profiles_main, text=GetLocaleText("User_Custom_profiles") + f" (0)", font=self.app_header_font)
        self.label_user_game_files.pack(side="top")

        self.xscrollbar_user_game_files = ctk.CTkScrollbar(self.box_custom_profiles_main, width=20, orientation="horizontal")
        self.xscrollbar_user_game_files.pack(side="bottom", fill="x")

        self.yscrollbar_user_game_files = ctk.CTkScrollbar(self.box_custom_profiles_main)
        self.yscrollbar_user_game_files.pack(side="right", fill="y")

        self.cols_user_game_files = (GetLocaleText("Filename"), GetLocaleText("Spreadsheet"))
        self.list_user_game_files = ttk.Treeview(self.box_custom_profiles_main, show="headings", columns=self.cols_user_game_files, height=15, yscrollcommand=self.yscrollbar_user_game_files.set, xscrollcommand=self.xscrollbar_user_game_files.set)
        self.list_user_game_files.heading("#1", text=self.cols_user_game_files[0], anchor="w")
        self.list_user_game_files.heading("#2", text=self.cols_user_game_files[1], anchor="w")
        self.list_user_game_files.column("#1", minwidth=140, width=140)
        self.list_user_game_files.column("#2", minwidth=210, width=210)
        self.list_user_game_files.pack(expand=True, fill="both")

        self.xscrollbar_user_game_files.configure(command=self.list_user_game_files.xview)
        self.yscrollbar_user_game_files.configure(command=self.list_user_game_files.yview)

        self.box_custom_profiles_btns_top = ctk.CTkFrame(self.box_custom_profiles)
        self.box_custom_profiles_btns_top.pack(side="top", expand=True, fill="both")

        self.btn_add_user_game = ctk.CTkButton(self.box_custom_profiles_btns_top, text=GetLocaleText("Add_Game_to_User_List"), width=200, height=50, cursor="hand2", command=self.UserGamesAdd, font=self.app_header_font)
        self.btn_add_user_game.pack(padx=2, pady=2, side="left", expand=True, fill="both")

        self.tooltip_add_user_game = CTkToolTip(self.btn_add_user_game, GetLocaleText("Add_Game_to_User_List_Description"))

        self.btn_remove_user_game = ctk.CTkButton(self.box_custom_profiles_btns_top, text=GetLocaleText("Remove_Game_from_User_List"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.UserGamesRemove)
        self.btn_remove_user_game.pack(padx=2, pady=2, side="left", expand=True, fill="both")

        self.tooltip_remove_user_game = CTkToolTip(self.btn_remove_user_game, GetLocaleText("Remove_Game_from_User_List_Description"))

        self.btn_copy_game_list = ctk.CTkButton(self.box_custom_profiles, text=GetLocaleText("Copy_Game_List"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.CopyGameListToClipboard)
        self.btn_copy_game_list.pack(padx=2, pady=2, side="top", expand=True, fill="both")

        self.tooltip_copy_game_list = CTkToolTip(self.btn_copy_game_list, GetLocaleText("Copy_Game_List_Description"))

        self._last_game_list_selected = None

        # Bind selection events
        self.list_csv_files.bind("<<TreeviewSelect>>", self.GameListSelected)
        self.list_online_game_files.bind("<<TreeviewSelect>>", self.GameListSelected)
        self.list_user_game_files.bind("<<TreeviewSelect>>", self.GameListSelected)

        # Bind double-click events
        self.list_csv_files.bind("<Double-1>", self.double_click_spreadsheet)
        self.list_online_game_files.bind("<Double-1>", self.double_click_spreadsheet)
        self.list_user_game_files.bind("<Double-1>", self.double_click_spreadsheet)

        self.update_quadstick_flash_files_items()
        self.update_online_game_files_list_items()
        self.update_user_game_files_list_items()

    def build_joystick_tab(self, parent_frame):
        # Joystick Tab left

        self.box_joystick = ctk.CTkFrame(parent_frame)
        self.box_joystick.pack(padx=5, pady=5, side="left", expand=True, fill="both")

        self.box_joystick_center = ctk.CTkFrame(self.box_joystick, height=400)
        self.box_joystick_center.pack(side="left", expand=True, fill="both")

        self.box_joystick_info = ctk.CTkFrame(self.box_joystick_center, height=400)
        self.box_joystick_info.pack(side="left", expand=True, fill="both")

        self.label_joystick_info = ctk.CTkLabel(self.box_joystick_info, text=GetLocaleText("Joystick_Info"), font=self.app_header_font, justify="left", wraplength=300)
        self.label_joystick_info.pack(side="top", anchor="w")

        # Slider UP
        self.box_slider_UP = ctk.CTkFrame(self.box_joystick_center)
        self.box_slider_UP.pack(pady=(0, 5), side="top", expand=True, fill="y")

        self.box_slider_UP_value = ctk.CTkFrame(self.box_slider_UP)
        self.box_slider_UP_value.pack(padx=(0, 5), side="left")

        self.label_slider_UP_value = ctk.CTkLabel(self.box_slider_UP_value, text="0", font=self.app_header_font)
        self.label_slider_UP_value.pack(side="left")

        self.box_slider_UP_main = ctk.CTkFrame(self.box_slider_UP)
        self.box_slider_UP_main.pack(side="left")

        self.label_slider_UP_max = ctk.CTkLabel(self.box_slider_UP_main, text="50", font=self.app_header_font)
        self.label_slider_UP_max.pack(side="top")

        self.slider_UP = ctk.CTkSlider(self.box_slider_UP_main, from_=10, to=50, command=self.joystick_slider_up_event, orientation="vertical", height=140, width=20)
        self.slider_UP.pack(side="top", expand=True, fill="y")

        self.tooltip_slider_UP = CTkToolTip(self.slider_UP, GetLocaleText("Joystick_UP_movement"))

        self.label_slider_UP_min = ctk.CTkLabel(self.box_slider_UP_main, text="10", font=self.app_header_font)
        self.label_slider_UP_min.pack(side="top")

        self.box_joystick_center2 = ctk.CTkFrame(self.box_joystick_center)
        self.box_joystick_center2.pack(padx=5, side="top", expand=True, fill="both")

        # Slider LEFT
        self.box_slider_LEFT = ctk.CTkFrame(self.box_joystick_center2)
        self.box_slider_LEFT.pack(padx=(0, 15), side="left", expand=True, fill="y")

        self.box_slider_LEFT_value = ctk.CTkFrame(self.box_slider_LEFT)
        self.box_slider_LEFT_value.pack(side="top")

        self.label_slider_LEFT_value = ctk.CTkLabel(self.box_slider_LEFT_value, text="0", font=self.app_header_font)
        self.label_slider_LEFT_value.pack(side="top")

        self.box_slider_LEFT_main = ctk.CTkFrame(self.box_slider_LEFT)
        self.box_slider_LEFT_main.pack(side="top")

        self.label_slider_LEFT_max = ctk.CTkLabel(self.box_slider_LEFT_main, text="50", font=self.app_header_font)
        self.label_slider_LEFT_max.pack(side="left")

        self.slider_LEFT = ctk.CTkSlider(self.box_slider_LEFT_main, from_=50, to=10, command=self.joystick_slider_left_event, orientation="horizontal", height=20, width=140)
        self.slider_LEFT.pack(padx= 5, side="left", expand=True, fill="x")

        self.tooltip_slider_LEFT = CTkToolTip(self.slider_LEFT, GetLocaleText("Joystick_LEFT_movement"))

        self.label_slider_LEFT_min = ctk.CTkLabel(self.box_slider_LEFT_main, text="10", font=self.app_header_font)
        self.label_slider_LEFT_min.pack(side="left")

        # Slider NEUTRAL
        self.box_slider_NEUTRAL = ctk.CTkFrame(self.box_joystick_center2)
        self.box_slider_NEUTRAL.pack(pady=(0, 5), side="left", expand=True, fill="y")

        self.box_slider_NEUTRAL_value = ctk.CTkFrame(self.box_slider_NEUTRAL)
        self.box_slider_NEUTRAL_value.pack(side="top")

        self.label_slider_NEUTRAL_value = ctk.CTkLabel(self.box_slider_NEUTRAL_value, text="0", font=self.app_header_font)
        self.label_slider_NEUTRAL_value.pack(side="top")

        self.box_slider_NEUTRAL_main = ctk.CTkFrame(self.box_slider_NEUTRAL)
        self.box_slider_NEUTRAL_main.pack(side="left")

        self.label_slider_NEUTRAL_min = ctk.CTkLabel(self.box_slider_NEUTRAL_main, text="0", font=self.app_header_font)
        self.label_slider_NEUTRAL_min.pack(side="left")

        self.slider_NEUTRAL = ctk.CTkSlider(self.box_slider_NEUTRAL_main, from_=0, to=20, command=self.joystick_slider_neutral_event, orientation="horizontal", height=20, width=140)
        self.slider_NEUTRAL.pack(padx=5, side="left", expand=True, fill="x")

        self.tooltip_slider_NEUTRAL = CTkToolTip(self.slider_NEUTRAL, GetLocaleText("Joystick_NEUTRAL_movement"))

        self.label_slider_NEUTRAL_max = ctk.CTkLabel(self.box_slider_NEUTRAL_main, text="20", font=self.app_header_font)
        self.label_slider_NEUTRAL_max.pack(side="left")

        # Slider RIGHT
        self.box_slider_RIGHT = ctk.CTkFrame(self.box_joystick_center2)
        self.box_slider_RIGHT.pack(padx=(15, 0), side="left", expand=True, fill="y")

        self.box_slider_RIGHT_value = ctk.CTkFrame(self.box_slider_RIGHT)
        self.box_slider_RIGHT_value.pack(side="top")

        self.label_slider_RIGHT_value = ctk.CTkLabel(self.box_slider_RIGHT_value, text="0", font=self.app_header_font)
        self.label_slider_RIGHT_value.pack(side="top")

        self.box_slider_RIGHT_main = ctk.CTkFrame(self.box_slider_RIGHT)
        self.box_slider_RIGHT_main.pack(side="top")

        self.label_slider_RIGHT_min = ctk.CTkLabel(self.box_slider_RIGHT_main, text="10", font=self.app_header_font)
        self.label_slider_RIGHT_min.pack(side="left")

        self.slider_RIGHT = ctk.CTkSlider(self.box_slider_RIGHT_main, from_=10, to=50, command=self.joystick_slider_right_event, orientation="horizontal", height=20, width=140)
        self.slider_RIGHT.pack(padx= 5, side="left", expand=True, fill="x")

        self.tooltip_slider_RIGHT = CTkToolTip(self.slider_RIGHT, GetLocaleText("Joystick_RIGHT_movement"))

        self.label_slider_RIGHT_max = ctk.CTkLabel(self.box_slider_RIGHT_main, text="50", font=self.app_header_font)
        self.label_slider_RIGHT_max.pack(side="left")

        # Slider DOWN
        self.box_slider_DOWN = ctk.CTkFrame(self.box_joystick_center)
        self.box_slider_DOWN.pack(side="top", expand=True, fill="y")

        self.box_slider_DOWN_value = ctk.CTkFrame(self.box_slider_DOWN)
        self.box_slider_DOWN_value.pack(padx=(0, 5), side="left")

        self.label_slider_DOWN_value = ctk.CTkLabel(self.box_slider_DOWN_value, text="0", font=self.app_header_font)
        self.label_slider_DOWN_value.pack(side="left")

        self.box_slider_DOWN_main = ctk.CTkFrame(self.box_slider_DOWN)
        self.box_slider_DOWN_main.pack(side="left")

        self.label_slider_DOWN_min = ctk.CTkLabel(self.box_slider_DOWN_main, text="10", font=self.app_header_font)
        self.label_slider_DOWN_min.pack(side="top")

        self.slider_DOWN = ctk.CTkSlider(self.box_slider_DOWN_main, from_=50, to=10, command=self.joystick_slider_down_event, orientation="vertical", height=140, width=20)
        self.slider_DOWN.pack(side="top", expand=True, fill="y")

        self.tooltip_slider_DOWN = CTkToolTip(self.slider_DOWN, GetLocaleText("Joystick_DOWN_movement"))

        self.label_slider_DOWN_max = ctk.CTkLabel(self.box_slider_DOWN_main, text="50", font=self.app_header_font)
        self.label_slider_DOWN_max.pack(side="top")

        # Joystick Tab right

        self.box_joystick_right = ctk.CTkFrame(self.box_joystick, height=400)
        self.box_joystick_right.pack(side="left", expand=True, fill="both")

        # Slider linking
        self.box_slider_linking = ctk.CTkFrame(self.box_joystick_right, height=400)
        self.box_slider_linking.pack(side="top", expand=True, fill="both")

        self.label_slider_linking = ctk.CTkLabel(self.box_slider_linking, text=GetLocaleText("Link_Sliders"), font=self.app_header_font)
        self.label_slider_linking.pack(side="top", anchor="w")

        self.slider_linking_var = ctk.IntVar(value=0)
        self.slider_linking_all = ctk.CTkRadioButton(self.box_slider_linking, text=GetLocaleText("all"), font=self.app_font, variable= self.slider_linking_var, value=0)
        self.slider_linking_all.pack(anchor="w")

        self.slider_linking_horizontal = ctk.CTkRadioButton(self.box_slider_linking, text=GetLocaleText("horizontal"), font=self.app_font, variable= self.slider_linking_var, value=1)
        self.slider_linking_horizontal.pack(anchor="w")

        self.slider_linking_vertical = ctk.CTkRadioButton(self.box_slider_linking, text=GetLocaleText("vertical"), font=self.app_font, variable= self.slider_linking_var, value=2)
        self.slider_linking_vertical.pack(anchor="w")
        
        self.slider_linking_none = ctk.CTkRadioButton(self.box_slider_linking, text=GetLocaleText("none"), font=self.app_font, variable= self.slider_linking_var, value=3)
        self.slider_linking_none.pack(anchor="w")

        # Joystick Preference
        self.box_joystick_preference = ctk.CTkFrame(self.box_joystick_right)
        self.box_joystick_preference.pack(side="top", fill="x")

        self.label_joystick_preference = ctk.CTkLabel(self.box_joystick_preference, text=GetLocaleText("Calculated_joystick_preference_values"), font=self.app_header_font)
        self.label_joystick_preference.pack(side="top", anchor="w")

        self.cols_joystick_preference = (GetLocaleText("Preference"), GetLocaleText("Value"))
        self.list_joystick_preference = ttk.Treeview(self.box_joystick_preference, show="headings", columns=self.cols_joystick_preference, height=10)
        self.list_joystick_preference.heading("#1", text=self.cols_joystick_preference[0], anchor="w")
        self.list_joystick_preference.heading("#2", text=self.cols_joystick_preference[1], anchor="w")
        self.list_joystick_preference.column("#1", width=150)
        self.list_joystick_preference.column("#2", width=10)
        self.list_joystick_preference.pack(expand=True, fill="both")

    def build_misc_tab(self, parent_frame):
        # Sliders Box #
        self.box_misc_sliders = ctk.CTkFrame(parent_frame)
        self.box_misc_sliders.pack(padx=2, side="left", expand=True, fill="both")

        # Sliders top

        self.box_misc_sliders_top = ctk.CTkFrame(self.box_misc_sliders)
        self.box_misc_sliders_top.pack(side="top", expand=True, fill="both")

        # Mouse Speed
        self.box_mouse_speed = ctk.CTkFrame(self.box_misc_sliders_top, height=150)
        self.box_mouse_speed.pack(padx=1, pady=1, expand=True, side="left", fill="both")

        self.label_mouse_speed = ctk.CTkLabel(self.box_mouse_speed, text=GetLocaleText("Mouse_Speed"), font=self.app_header_font)
        self.label_mouse_speed.pack(side="top")

        self.label_mouse_speed_100 = ctk.CTkLabel(self.box_mouse_speed, text="100", font=self.app_header_font)
        self.label_mouse_speed_100.pack(side="top")

        self.label_mouse_speed_value = ctk.CTkLabel(self.box_mouse_speed, text="0", font=self.app_header_font)
        self.label_mouse_speed_value.pack(side="left")

        self.slider_mouse_speed = ctk.CTkSlider(self.box_mouse_speed, from_=0, to=100, command=self.slider_mouse_speed_event, orientation="vertical", height=150, width=20)
        self.slider_mouse_speed.pack(side="top", expand=True, fill="y")

        self.tooltip_mouse_speed = CTkToolTip(self.slider_mouse_speed, GetLocaleText("Mouse_Speed_Description"))

        self.label_mouse_speed_0 = ctk.CTkLabel(self.box_mouse_speed, text="0", font=self.app_header_font)
        self.label_mouse_speed_0.pack(side="top")

        # Brightness
        self.box_brightness = ctk.CTkFrame(self.box_misc_sliders_top, height=150)
        self.box_brightness.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_brightness = ctk.CTkLabel(self.box_brightness, text=GetLocaleText("Brightness"), font=self.app_header_font)
        self.label_brightness.pack(side="top")

        self.label_brightness_100 = ctk.CTkLabel(self.box_brightness, text="100", font=self.app_header_font)
        self.label_brightness_100.pack(side="top")

        self.label_brightness_value = ctk.CTkLabel(self.box_brightness, text="0", font=self.app_header_font)
        self.label_brightness_value.pack(side="left")

        self.slider_brightness = ctk.CTkSlider(self.box_brightness, from_=0, to=100, command=self.slider_brightness_event, orientation="vertical", height=150, width=20)
        self.slider_brightness.pack(side="top", expand=True, fill="y")

        self.tooltip_brightness = CTkToolTip(self.slider_brightness, GetLocaleText("Brightness_Description"))

        self.label_brightness_0 = ctk.CTkLabel(self.box_brightness, text="0", font=self.app_header_font)
        self.label_brightness_0.pack(side="top")

        # Volume
        self.box_volume = ctk.CTkFrame(self.box_misc_sliders_top, height=150)
        self.box_volume.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_volume = ctk.CTkLabel(self.box_volume, text=GetLocaleText("Volume"), font=self.app_header_font)
        self.label_volume.pack(side="top")

        self.label_volume_100 = ctk.CTkLabel(self.box_volume, text="100", font=self.app_header_font)
        self.label_volume_100.pack(side="top")

        self.label_volume_value = ctk.CTkLabel(self.box_volume, text="0", font=self.app_header_font)
        self.label_volume_value.pack(side="left")

        self.slider_volume = ctk.CTkSlider(self.box_volume, from_=0, to=100, command=self.slider_volume_event, orientation="vertical", height=150, width=20)
        self.slider_volume.pack(side="top", expand=True, fill="y")

        self.tooltip_volume = CTkToolTip(self.slider_volume, GetLocaleText("Volume_Description"))

        self.label_volume_0 = ctk.CTkLabel(self.box_volume, text="0", font=self.app_header_font)
        self.label_volume_0.pack(side="top")

        # D-Pad outer ring
        self.box_D_Pad_outer_ring = ctk.CTkFrame(self.box_misc_sliders_top, height=150)
        self.box_D_Pad_outer_ring.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_D_Pad_outer_ring = ctk.CTkLabel(self.box_D_Pad_outer_ring, text=GetLocaleText("D_Pad_outer_ring"), font=self.app_header_font)
        self.label_D_Pad_outer_ring.pack(side="top")

        self.label_D_Pad_outer_ring_100 = ctk.CTkLabel(self.box_D_Pad_outer_ring, text="100", font=self.app_header_font)
        self.label_D_Pad_outer_ring_100.pack(side="top")

        self.label_D_Pad_outer_ring_value = ctk.CTkLabel(self.box_D_Pad_outer_ring, text="0", font=self.app_header_font)
        self.label_D_Pad_outer_ring_value.pack(side="left")

        self.slider_D_Pad_outer_ring = ctk.CTkSlider(self.box_D_Pad_outer_ring, from_=0, to=100, command=self.slider_D_Pad_outer_ring_event, orientation="vertical", height=150, width=20)
        self.slider_D_Pad_outer_ring.pack(side="top", expand=True, fill="y")

        self.tooltip_D_Pad_outer_ring = CTkToolTip(self.slider_D_Pad_outer_ring, GetLocaleText("D_Pad_outer_ring_Description"))

        self.label_D_Pad_outer_ring_0 = ctk.CTkLabel(self.box_D_Pad_outer_ring, text="0", font=self.app_header_font)
        self.label_D_Pad_outer_ring_0.pack(side="top")

        # D-Pad inner ring
        self.box_D_Pad_inner_ring = ctk.CTkFrame(self.box_misc_sliders_top, height=150)
        self.box_D_Pad_inner_ring.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_D_Pad_inner_ring = ctk.CTkLabel(self.box_D_Pad_inner_ring, text=GetLocaleText("D_Pad_inner_ring"), font=self.app_header_font)
        self.label_D_Pad_inner_ring.pack(side="top")

        self.label_D_Pad_inner_ring_100 = ctk.CTkLabel(self.box_D_Pad_inner_ring, text="100", font=self.app_header_font)
        self.label_D_Pad_inner_ring_100.pack(side="top")

        self.label_D_Pad_inner_ring_value = ctk.CTkLabel(self.box_D_Pad_inner_ring, text="0", font=self.app_header_font)
        self.label_D_Pad_inner_ring_value.pack(side="left")

        self.slider_D_Pad_inner_ring = ctk.CTkSlider(self.box_D_Pad_inner_ring, from_=0, to=100, command=self.slider_D_Pad_inner_ring_event, orientation="vertical", height=150, width=20)
        self.slider_D_Pad_inner_ring.pack(side="top", expand=True, fill="y")

        self.tooltip_D_Pad_inner_ring = CTkToolTip(self.slider_D_Pad_inner_ring, GetLocaleText("D_Pad_inner_ring_Description"))

        self.label_D_Pad_inner_ring_0 = ctk.CTkLabel(self.box_D_Pad_inner_ring, text="0", font=self.app_header_font)
        self.label_D_Pad_inner_ring_0.pack(side="top")

        # Sliders bottom

        self.box_misc_sliders_bottom = ctk.CTkFrame(self.box_misc_sliders)
        self.box_misc_sliders_bottom.pack(side="top", expand=True, fill="both")

        # Sip and Puff max pressure
        self.box_SP_max = ctk.CTkFrame(self.box_misc_sliders_bottom, height=150)
        self.box_SP_max.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_SP_max = ctk.CTkLabel(self.box_SP_max, text=GetLocaleText("Sip_and_Puff_max_pressure"), font=self.app_header_font)
        self.label_SP_max.pack(padx=10, side="top")

        self.label_SP_max_100 = ctk.CTkLabel(self.box_SP_max, text="100", font=self.app_header_font)
        self.label_SP_max_100.pack(side="top")

        self.label_SP_max_value = ctk.CTkLabel(self.box_SP_max, text="0", font=self.app_header_font)
        self.label_SP_max_value.pack(side="left")

        self.slider_SP_max = ctk.CTkSlider(self.box_SP_max, from_=0, to=100, command=self.slider_SP_max_event, orientation="vertical", height=150, width=20)
        self.slider_SP_max.pack(side="top", expand=True, fill="y")

        self.tooltip_SP_max = CTkToolTip(self.slider_SP_max, GetLocaleText("Sip_and_Puff_max_pressure_Description"))

        self.label_SP_max_0 = ctk.CTkLabel(self.box_SP_max, text="0", font=self.app_header_font)
        self.label_SP_max_0.pack(side="top")

        # Sip and Puff high threshold
        self.box_SP_high = ctk.CTkFrame(self.box_misc_sliders_bottom, height=150)
        self.box_SP_high.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_SP_high = ctk.CTkLabel(self.box_SP_high, text=GetLocaleText("Sip_and_Puff_high_threshold"), font=self.app_header_font)
        self.label_SP_high.pack(padx=10, side="top")

        self.label_SP_high_100 = ctk.CTkLabel(self.box_SP_high, text="100", font=self.app_header_font)
        self.label_SP_high_100.pack(side="top")

        self.label_SP_high_value = ctk.CTkLabel(self.box_SP_high, text="0", font=self.app_header_font)
        self.label_SP_high_value.pack(side="left")

        self.slider_SP_high = ctk.CTkSlider(self.box_SP_high, from_=0, to=100, command=self.slider_SP_high_event, orientation="vertical", height=150, width=20)
        self.slider_SP_high.pack(side="top", expand=True, fill="y")

        self.tooltip_SP_high = CTkToolTip(self.slider_SP_high, GetLocaleText("Sip_and_Puff_high_threshold_Description"))

        self.label_SP_high_0 = ctk.CTkLabel(self.box_SP_high, text="0", font=self.app_header_font)
        self.label_SP_high_0.pack(side="top")

        # Sip and Puff low threshold
        self.box_SP_low = ctk.CTkFrame(self.box_misc_sliders_bottom, height=150)
        self.box_SP_low.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_SP_low = ctk.CTkLabel(self.box_SP_low, text=GetLocaleText("Sip_and_Puff_low_threshold"), font=self.app_header_font)
        self.label_SP_low.pack(padx=10, side="top")

        self.label_SP_low_100 = ctk.CTkLabel(self.box_SP_low, text="100", font=self.app_header_font)
        self.label_SP_low_100.pack(side="top")

        self.label_SP_low_value = ctk.CTkLabel(self.box_SP_low, text="0", font=self.app_header_font)
        self.label_SP_low_value.pack(side="left")

        self.slider_SP_low = ctk.CTkSlider(self.box_SP_low, from_=0, to=100, command=self.slider_SP_low_event, orientation="vertical", height=150, width=20)
        self.slider_SP_low.pack(side="top", expand=True, fill="y")

        self.tooltip_SP_low = CTkToolTip(self.slider_SP_low, GetLocaleText("Sip_and_Puff_low_threshold_Description"))

        self.label_SP_low_0 = ctk.CTkLabel(self.box_SP_low, text="0", font=self.app_header_font)
        self.label_SP_low_0.pack(side="top")

        # Lip maximum
        self.box_Lip_max = ctk.CTkFrame(self.box_misc_sliders_bottom, height=150)
        self.box_Lip_max.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_Lip_max = ctk.CTkLabel(self.box_Lip_max, text=GetLocaleText("Lip_maximum"), font=self.app_header_font)
        self.label_Lip_max.pack(padx=10, side="top")

        self.label_Lip_max_100 = ctk.CTkLabel(self.box_Lip_max, text="100", font=self.app_header_font)
        self.label_Lip_max_100.pack(side="top")

        self.label_Lip_max_value = ctk.CTkLabel(self.box_Lip_max, text="0", font=self.app_header_font)
        self.label_Lip_max_value.pack(side="left")

        self.slider_Lip_max = ctk.CTkSlider(self.box_Lip_max, from_=0, to=100, command=self.slider_Lip_max_event, orientation="vertical", height=150, width=20)
        self.slider_Lip_max.pack(side="top", expand=True, fill="y")

        self.tooltip_Lip_max = CTkToolTip(self.slider_Lip_max, GetLocaleText("Lip_maximum_Description"))

        self.label_Lip_max_0 = ctk.CTkLabel(self.box_Lip_max, text="0", font=self.app_header_font)
        self.label_Lip_max_0.pack(side="top")

        # Lip minimum
        self.box_Lip_min = ctk.CTkFrame(self.box_misc_sliders_bottom, height=150)
        self.box_Lip_min.pack(padx=1, pady=1, side="left", expand=True, fill="both")

        self.label_Lip_min = ctk.CTkLabel(self.box_Lip_min, text=GetLocaleText("Lip_minimum"), font=self.app_header_font)
        self.label_Lip_min.pack(padx=10, side="top")

        self.label_Lip_min_100 = ctk.CTkLabel(self.box_Lip_min, text="100", font=self.app_header_font)
        self.label_Lip_min_100.pack(side="top")

        self.label_Lip_min_value = ctk.CTkLabel(self.box_Lip_min, text="0", font=self.app_header_font)
        self.label_Lip_min_value.pack(side="left")

        self.slider_Lip_min = ctk.CTkSlider(self.box_Lip_min, from_=0, to=100, command=self.slider_Lip_min_event, orientation="vertical", height=150, width=20)
        self.slider_Lip_min.pack(side="top", expand=True, fill="y")

        self.tooltip_Lip_min = CTkToolTip(self.slider_Lip_min, GetLocaleText("Lip_minimum_Description"))

        self.label_Lip_min_0 = ctk.CTkLabel(self.box_Lip_min, text="0", font=self.app_header_font)
        self.label_Lip_min_0.pack(side="top")

        # Settings Box #
        self.box_misc_settings = ctk.CTkScrollableFrame(parent_frame, width=400)
        self.box_misc_settings.pack(padx=1, side="left", fill="both")

        self.label_misc_settings = ctk.CTkLabel(self.box_misc_settings, text=GetLocaleText("Other_settings"), font=self.app_header_font)
        self.label_misc_settings.pack(side="top")

        self.tabs_misc_settings1 = ctk.CTkTabview(self.box_misc_settings, height=130, corner_radius=15)
        self.tabs_misc_settings1.pack(side="top", fill="both")
        self.tabs_misc_settings1._segmented_button.configure(font=self.app_font)

        self.tab_misc_do = self.tabs_misc_settings1.add(GetLocaleText("Digital_Outputs"))
        self.tab_misc_mouse = self.tabs_misc_settings1.add(GetLocaleText("Mouse"))
        self.tab_misc_bluetooth = self.tabs_misc_settings1.add(GetLocaleText("Bluetooth"))

        # Digital Outputs
        self.checkbox_do_1_state = ctk.IntVar(value=0)
        self.checkbox_do_1 = ctk.CTkSwitch(self.tab_misc_do, text=GetLocaleText("Digital_output_1"), variable=self.checkbox_do_1_state)
        self.checkbox_do_1.pack(side="top", anchor="w")

        self.tooltip_do_1 = CTkToolTip(self.checkbox_do_1, GetLocaleText("Digital_output_1_Description"))

        self.checkbox_do_2_state = ctk.IntVar(value=0)
        self.checkbox_do_2 = ctk.CTkSwitch(self.tab_misc_do, text=GetLocaleText("Digital_output_2"), variable=self.checkbox_do_2_state)
        self.checkbox_do_2.pack(side="top", anchor="w")

        self.tooltip_do_2 = CTkToolTip(self.checkbox_do_2, GetLocaleText("Digital_output_2_Description"))

        # Mouse
        self.checkbox_circular_deadzone_state = ctk.IntVar(value=0)
        self.checkbox_circular_deadzone = ctk.CTkSwitch(self.tab_misc_mouse, text=GetLocaleText("Enable_Circular_Dead_Zone"), variable=self.checkbox_circular_deadzone_state)
        self.checkbox_circular_deadzone.pack(side="top", anchor="w")

        self.tooltip_circular_deadzone = CTkToolTip(self.checkbox_circular_deadzone, GetLocaleText("Enable_Circular_Dead_Zone_Description"))

        self.box_mouse_response = ctk.CTkFrame(self.tab_misc_mouse)
        self.box_mouse_response.pack(side="top", anchor="w")

        self.label_mouse_response = ctk.CTkLabel(self.box_mouse_response, text=GetLocaleText("Mouse_Response_Curve"), font=self.app_font)
        self.label_mouse_response.pack(side="left")

        self.mouse_responses = [GetLocaleText("Linear"), GetLocaleText("Mixed"), GetLocaleText("Parabolic")]
        self.choice_mouse_response = ctk.CTkOptionMenu(self.box_mouse_response, values=self.mouse_responses)
        self.choice_mouse_response.pack(padx=(5, 0), side="left")

        self.tooltip_mouse_response = CTkToolTip(self.choice_mouse_response, GetLocaleText("Mouse_Response_Curve_Description"))

        # Bluetooth
        self.checkbox_enable_serial_port_state = ctk.IntVar(value=0)
        self.checkbox_enable_serial_port = ctk.CTkSwitch(self.tab_misc_bluetooth, text=GetLocaleText("Enable_file_management_over_serial_port"), variable=self.checkbox_enable_serial_port_state, command=self.EnableSerialPortEvent)
        self.checkbox_enable_serial_port.pack(side="top", anchor="w")

        self.tooltip_enable_serial_port = CTkToolTip(self.checkbox_enable_serial_port, GetLocaleText("Enable_file_management_over_serial_port_Description"))

        self.box_BT_device_mode = ctk.CTkFrame(self.tab_misc_bluetooth)
        self.box_BT_device_mode.pack(side="top", anchor="w")

        self.label_BT_device_mode = ctk.CTkLabel(self.box_BT_device_mode, text=GetLocaleText("Device_Mode"), font=self.app_header_font)
        self.label_BT_device_mode.pack(side="left")

        self.BT_device_modes = [GetLocaleText("Device_None"), GetLocaleText("Keyboard"), GetLocaleText("Game_Pad"), GetLocaleText("Mouse"), GetLocaleText("Combo"), GetLocaleText("Joystick"), GetLocaleText("SSP")]
        self.choice_BT_device_mode = ctk.CTkOptionMenu(self.box_BT_device_mode, values=self.BT_device_modes)
        self.choice_BT_device_mode.pack(padx=(5, 0), side="left")

        self.tooltip_BT_device_mode = CTkToolTip(self.choice_BT_device_mode, GetLocaleText("Device_Mode_Description"))

        self.box_BT_auth_mode = ctk.CTkFrame(self.tab_misc_bluetooth)
        self.box_BT_auth_mode.pack(pady=2, side="top", anchor="w")

        self.label_BT_auth_mode = ctk.CTkLabel(self.box_BT_auth_mode, text=GetLocaleText("Auth"), font=self.app_header_font)
        self.label_BT_auth_mode.pack(side="left")

        self.BT_auth_modes = ["0", "1", "2", "4"]
        self.choice_BT_auth_mode = ctk.CTkOptionMenu(self.box_BT_auth_mode, values=self.BT_auth_modes)
        self.choice_BT_auth_mode.pack(padx=(5, 0), side="left")

        self.tooltip_BT_auth_mode = CTkToolTip(self.choice_BT_auth_mode, GetLocaleText("Auth_Description"))

        # Misc tabs 2

        self.tabs_misc_settings2 = ctk.CTkTabview(self.box_misc_settings, height=130, corner_radius=15)
        self.tabs_misc_settings2.pack(side="top", fill="both")
        self.tabs_misc_settings2._segmented_button.configure(font=self.app_font)

        self.tab_misc_SP = self.tabs_misc_settings2.add(GetLocaleText("Sip_and_Puff_misc"))
        self.tab_misc_usb = self.tabs_misc_settings2.add(GetLocaleText("USB_settings"))
        self.tab_misc_vg = self.tabs_misc_settings2.add(GetLocaleText("Virtual_gamepad_emulator"))

        # Sip and Puff misc
        self.checkbox_select_files_state = ctk.IntVar(value=0)
        self.checkbox_select_files = ctk.CTkSwitch(self.tab_misc_SP, text=GetLocaleText("Enable_select_file_with_side_tube"), variable=self.checkbox_select_files_state)
        self.checkbox_select_files.pack(side="top", anchor="w")

        self.tooltip_select_files = CTkToolTip(self.checkbox_select_files, GetLocaleText("Enable_select_file_with_side_tube_Description"))

        self.checkbox_swap_state = ctk.IntVar(value=0)
        self.checkbox_swap = ctk.CTkSwitch(self.tab_misc_SP, text=GetLocaleText("Enable_swap_inputs_with_side_tube"), variable=self.checkbox_swap_state)
        self.checkbox_swap.pack(side="top", anchor="w")

        self.tooltip_swap = CTkToolTip(self.checkbox_swap, GetLocaleText("Enable_swap_inputs_with_side_tube_Description"))

        self.box_SP_low_delay = ctk.CTkFrame(self.tab_misc_SP, width=500)
        self.box_SP_low_delay.pack(side="top", expand=True, anchor="w")

        self.label_SP_low_delay = ctk.CTkLabel(self.box_SP_low_delay, text=GetLocaleText("Low_threshold_delay"), font=self.app_font)
        self.label_SP_low_delay.pack(padx=(0, 5), side="left")

        self.SP_low_delay_value = ctk.IntVar(value=1200)

        self.btn_SP_low_delay_decrement = ctk.CTkButton(self.box_SP_low_delay, text="-", width=20, font=self.app_header_font, command=lambda: self.decrement(self.SP_low_delay_value, 100))
        self.btn_SP_low_delay_decrement.pack(side="left")

        self.entry_SP_low_delay = ctk.CTkEntry(self.box_SP_low_delay, textvariable=self.SP_low_delay_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_SP_low_delay.pack(side="left")

        self.tooltip_SP_low_delay = CTkToolTip(self.entry_SP_low_delay, GetLocaleText("Low_threshold_delay_Description"))

        self.btn_SP_low_delay_increment = ctk.CTkButton(self.box_SP_low_delay, text="+", width=20, font=self.app_header_font, command=lambda: self.increment(self.SP_low_delay_value, 3000))
        self.btn_SP_low_delay_increment.pack(side="left")

        self.box_SP_high_delay = ctk.CTkFrame(self.tab_misc_SP)
        self.box_SP_high_delay.pack(side="top", expand=True, anchor="w")

        self.label_SP_high_delay = ctk.CTkLabel(self.box_SP_high_delay, text=GetLocaleText("High_threshold_delay"), font=self.app_font)
        self.label_SP_high_delay.pack(padx=(0, 5), side="left")

        self.SP_high_delay_value = ctk.IntVar(value=2000)

        self.btn_SP_high_delay_decrement = ctk.CTkButton(self.box_SP_high_delay, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.SP_high_delay_value, 1000))
        self.btn_SP_high_delay_decrement.pack(side="left")

        self.entry_SP_high_delay = ctk.CTkEntry(self.box_SP_high_delay, textvariable=self.SP_high_delay_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_SP_high_delay.pack(side="left")

        # TODO self.SP_high_delay_value.trace_add("write", self.on_entry_change)

        self.tooltip_SP_high_delay = CTkToolTip(self.entry_SP_high_delay, GetLocaleText("High_threshold_delay_Description"))

        self.btn_SP_high_delay_increment = ctk.CTkButton(self.box_SP_high_delay, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.SP_high_delay_value, 3000))
        self.btn_SP_high_delay_increment.pack(side="left")

        # USB Settings
        self.checkbox_ps4_boot_mode_state = ctk.IntVar(value=0)
        self.checkbox_ps4_boot_mode = ctk.CTkSwitch(self.tab_misc_usb, text=GetLocaleText("Enable_boot_in_PS4_USB_mode"), variable=self.checkbox_ps4_boot_mode_state)
        self.checkbox_ps4_boot_mode.pack(side="top", anchor="w")

        self.tooltip_ps4_boot_mode = CTkToolTip(self.checkbox_ps4_boot_mode, GetLocaleText("Enable_boot_in_PS4_USB_mode_Description"))

        self.checkbox_Titan2_state = ctk.IntVar(value=0)
        self.checkbox_Titan2 = ctk.CTkSwitch(self.tab_misc_usb, text=GetLocaleText("Titan_Two_PS4_flag"), variable=self.checkbox_Titan2_state)
        self.checkbox_Titan2.pack(side="top", anchor="w")

        self.tooltip_titan2 = CTkToolTip(self.checkbox_Titan2, GetLocaleText("Titan_Two_PS4_flag_Description"))

        self.checkbox_usb_comm_state = ctk.IntVar(value=0)
        self.checkbox_enable_usb_comm = ctk.CTkSwitch(self.tab_misc_usb, text=GetLocaleText("Enable_Voice_Commands_over_USB"), variable=self.checkbox_usb_comm_state)
        self.checkbox_enable_usb_comm.pack(side="top", anchor="w")

        self.tooltip_enable_usb_comm = CTkToolTip(self.checkbox_enable_usb_comm, GetLocaleText("Enable_Voice_Commands_over_USB_Description"))

        self.checkbox_usb_A_host_mode_state = ctk.IntVar(value=0)
        self.checkbox_usb_A_host_mode = ctk.CTkSwitch(self.tab_misc_usb, text=GetLocaleText("Enable_USB_A_Host_mode"), variable=self.checkbox_usb_A_host_mode_state)
        self.checkbox_usb_A_host_mode.pack(side="top", anchor="w")

        self.tooltip_usb_A_host_mode = CTkToolTip(self.checkbox_usb_A_host_mode, GetLocaleText("Enable_USB_A_Host_mode_Description"))

        # Virtual gamepad emulator
        self.checkbox_enable_vg_X360_state = ctk.IntVar(value=0)
        self.checkbox_enable_vg_X360 = ctk.CTkSwitch(self.tab_misc_vg, text=GetLocaleText("Enable_virtual_XBox_controller_emulation"), variable=self.checkbox_enable_vg_X360_state)
        self.checkbox_enable_vg_X360.pack(side="top", anchor="w")

        self.tooltip_enable_vg_X360 = CTkToolTip(self.checkbox_enable_vg_X360, GetLocaleText("Enable_virtual_XBox_controller_emulation_Description"))

        self.checkbox_enable_vg_DS4_state = ctk.IntVar(value=0)
        self.checkbox_enable_vg_DS4 = ctk.CTkSwitch(self.tab_misc_vg, text=GetLocaleText("Enable_virtual_Dualshock_4_emulation"), variable=self.checkbox_enable_vg_DS4_state)
        self.checkbox_enable_vg_DS4.pack(side="top", anchor="w")

        self.tooltip_enable_vg_DS4 = CTkToolTip(self.checkbox_enable_vg_DS4, GetLocaleText("Enable_virtual_Dualshock_4_emulation_Description"))

        self.checkbox_enable_HIDHide_state = ctk.IntVar(value=0)
        self.checkbox_enable_HIDHide = ctk.CTkSwitch(self.tab_misc_vg, text=GetLocaleText("Enable_HIDHide"), variable=self.checkbox_enable_HIDHide_state)
        self.checkbox_enable_HIDHide.pack(side="top", anchor="w")

        self.tooltip_enable_HIDHide = CTkToolTip(self.checkbox_enable_HIDHide, GetLocaleText("Enable_HIDHide_Description"))

        # Program Settings
        self.box_program_settings = ctk.CTkFrame(self.box_misc_settings)
        self.box_program_settings.pack(side="top", expand=True, fill="both")

        self.label_program_settings = ctk.CTkLabel(self.box_program_settings, text=GetLocaleText("Program_Settings"), font=self.app_header_font)
        self.label_program_settings.pack(side="top")

        self.box_locale = ctk.CTkFrame(self.box_program_settings)
        self.box_locale.pack(side="top", anchor="w")

        self.label_locale = ctk.CTkLabel(self.box_locale, text=GetLocaleText("Language"), font=self.app_font)
        self.label_locale.pack(side="left")

        self.locales = self.GetLocaleNames()
        self.choice_locale = ctk.CTkOptionMenu(self.box_locale, values=self.locales, command=self.change_locale)
        self.choice_locale.pack(padx=(5, 0), side="left")

        self.tooltip_change_locale = CTkToolTip(self.choice_locale, GetLocaleText("Language_Description"))

        self.box_appearance_mode = ctk.CTkFrame(self.box_program_settings)
        self.box_appearance_mode.pack(pady=2, side="top", anchor="w")

        self.label_appearance_mode = ctk.CTkLabel(self.box_appearance_mode, text=GetLocaleText("Appearance_Mode"), anchor="w", font=self.app_font)
        self.label_appearance_mode.pack(side="left")

        self.appearance_modes = [GetLocaleText("Light"), GetLocaleText("Dark"), GetLocaleText("System")]
        self.choice_appearance_mode = ctk.CTkOptionMenu(self.box_appearance_mode, values=self.appearance_modes, command=self.change_appearance_mode)
        self.choice_appearance_mode.pack(padx=(5, 0), side="left")

        self.tooltip_appearance_mode = CTkToolTip(self.choice_appearance_mode, GetLocaleText("Appearance_Mode_Description"))

        self.checkbox_minimize_to_tray_state = ctk.BooleanVar(value=False)
        self.checkbox_minimize_to_tray = ctk.CTkSwitch(self.box_program_settings, text=GetLocaleText("Minimize_to_tray"), variable=self.checkbox_minimize_to_tray_state, command=self.checkbox_minimize_to_tray_event)
        self.checkbox_minimize_to_tray.pack(side="top", anchor="w")

        self.tooltip_minimize_to_tray = CTkToolTip(self.checkbox_minimize_to_tray, GetLocaleText("Minimize_to_tray_Description"))

        self.checkbox_start_minimized_state = ctk.BooleanVar(value=False)
        self.checkbox_start_minimized = ctk.CTkSwitch(self.box_program_settings, text=GetLocaleText("Start_minimized"), variable=self.checkbox_start_minimized_state)
        self.checkbox_start_minimized.pack(side="top", anchor="w")

        self.tooltip_start_minimized = CTkToolTip(self.checkbox_start_minimized, GetLocaleText("Start_minimized_Description"))

        self.checkbox_log_state = ctk.BooleanVar(value=False)
        self.checkbox_enable_log = ctk.CTkSwitch(self.box_program_settings, text=GetLocaleText("Enable_Log"), variable=self.checkbox_log_state, command=self.checkbox_enable_log_event)
        self.checkbox_enable_log.pack(side="top", anchor="w")

        self.tooltip_enable_log = CTkToolTip(self.checkbox_enable_log, GetLocaleText("Enable_Log_Description"))

        self.btn_import_data = ctk.CTkButton(self.box_program_settings, text=GetLocaleText("Import_Data"), height=60, cursor="hand2", font=self.app_header_font, command=self.import_data)
        self.btn_import_data.pack(side="top", fill="x", anchor="w")

        self.tooltip_import_data = CTkToolTip(self.btn_import_data, GetLocaleText("Import_Data_Description"))

        # def test():
        #     dlg = ctk.CTkInputDialog(text=GetLocaleText('Enter_Your_Email_Address'), title=GetLocaleText('Google_Spreadsheets_Account'))
        #     dlg_input = dlg.get_input() # waits for input
        #     if dlg_input is not None and dlg_input != "":
        #         print(dlg_input)

        # self.btn_test = ctk.CTkButton(self.box_program_settings, text="Test", height=50, cursor="hand2", font=self.app_header_font, command=test)
        # self.btn_test.pack(side="top", anchor="w")

        self.slider_mouse_speed.set(0)
        self.slider_brightness.set(0)
        self.slider_volume.set(0)
        self.slider_D_Pad_outer_ring.set(0)
        self.slider_D_Pad_inner_ring.set(0)
        self.slider_SP_max.set(0)
        self.slider_SP_high.set(0)
        self.slider_SP_low.set(0)
        self.slider_Lip_max.set(0)  
        self.slider_Lip_min.set(0)

        self.choice_locale.set(GetLocaleText(settings.get("current_locale")))
        self.choice_appearance_mode.set(GetLocaleText(settings.get("appearance_mode")))

        if settings.get("minimize_to_tray") == True:
            self.checkbox_minimize_to_tray.toggle()

        if settings.get("start_mimimized") == True:
            self.checkbox_start_minimized.toggle()

        if settings.get("enable_log") == True:
            self.checkbox_enable_log.toggle()

    def build_firmware_tab(self, parent_frame):
        # Current firmware Box
        self.box_current_firmware = ctk.CTkFrame(parent_frame)
        self.box_current_firmware.pack(side="left", expand=True, fill="both")

        self.label_current_firmware = ctk.CTkLabel(self.box_current_firmware, text=GetLocaleText("Current_Firmware"), font=self.app_header_font)
        self.label_current_firmware.pack(side="top")

        self.label_new_version = ctk.CTkLabel(self.box_current_firmware, text=GetLocaleText("New_Firmware_Version_Info"), font=self.app_font)
        self.label_new_version.pack(side="top")

        self.box_build_number = ctk.CTkFrame(self.box_current_firmware)
        self.box_build_number.pack(side="top", expand=True)

        self.label_Build_number = ctk.CTkLabel(self.box_build_number, text=GetLocaleText("Build_number"), font=self.app_header_font)
        self.label_Build_number.pack(side="top")

        self.build_number = int(0)
        self.build_number_text = ctk.CTkTextbox(self.box_build_number, width=100, height=20, font=self.app_font)
        self.build_number_text.configure(state="disabled")
        self.build_number_text._textbox.tag_configure("text", justify='center') # set text position using tag_configure.
        self.build_number_text.pack(pady=5, side="top", expand=True)

        self.label_firmware_update = ctk.CTkLabel(self.box_current_firmware, text=GetLocaleText("Firmware_Update_Info"), font=self.app_font)
        self.label_firmware_update.pack(side="top")

        self.box_available_firmware = ctk.CTkFrame(parent_frame)
        self.box_available_firmware.pack(side="left", expand=True, fill="both")

        self.label_available_firmware = ctk.CTkLabel(self.box_available_firmware, text=GetLocaleText("Available_Firmware"), font=self.app_header_font)
        self.label_available_firmware.pack(side="top")

        self.box_list_firmware = ctk.CTkFrame(self.box_available_firmware)
        self.box_list_firmware.pack(side="top", expand=True, fill="both")

        self.xscrollbar_list_firmware = ctk.CTkScrollbar(self.box_list_firmware, width=20, orientation="horizontal")
        self.xscrollbar_list_firmware.pack(side="bottom", fill="x")

        self.yscrollbar_list_firmware = ctk.CTkScrollbar(self.box_list_firmware)
        self.yscrollbar_list_firmware.pack(side="right", fill="y")

        self.cols_list_ctrl_firmware = (GetLocaleText("Build"), GetLocaleText("Remark"))
        self.list_ctrl_firmware = ttk.Treeview(self.box_list_firmware, show="headings", columns=self.cols_list_ctrl_firmware, height=15, yscrollcommand=self.yscrollbar_list_firmware.set, xscrollcommand=self.xscrollbar_list_firmware.set)
        self.list_ctrl_firmware.heading("#1", text=self.cols_list_ctrl_firmware[0], anchor="w")
        self.list_ctrl_firmware.heading("#2", text=self.cols_list_ctrl_firmware[1], anchor="w")
        self.list_ctrl_firmware.column("#1", width=30)
        self.list_ctrl_firmware.column("#2", width=310)
        self.list_ctrl_firmware.pack(side="top", expand=True, fill="both")

        self.xscrollbar_list_firmware.configure(command=self.list_ctrl_firmware.xview)
        self.yscrollbar_list_firmware.configure(command=self.list_ctrl_firmware.yview)

        self.btn_download_selected_build = ctk.CTkButton(self.box_available_firmware, text=GetLocaleText("Download_Firmware"), width=200, height=50, cursor="hand2", font=self.app_header_font, command=self.DownloadFirmwareEvent)
        self.btn_download_selected_build.pack(padx=2, pady=2, side="top", expand=True, fill="both")

        self.update_build_number(quadstick_drive_serial_number(self))
        self.update_firmware_list()

    def build_voice_control_tab(self, parent_frame):
        # Voice Transcript Box
        self.box_voice_cmd_transcript = ctk.CTkFrame(parent_frame)
        self.box_voice_cmd_transcript.pack(padx= 5, pady=5, side="left", expand=True, fill="both")

        self.label_voice_cmd_transcript = ctk.CTkLabel(self.box_voice_cmd_transcript, text=GetLocaleText("Voice_Command_Transcript"), font=self.app_header_font)
        self.label_voice_cmd_transcript.pack(side="top")

        self.voice_transcript = ctk.CTkTextbox(self.box_voice_cmd_transcript, font=self.app_font)
        self.voice_transcript.pack(side="left", expand=True, fill="both")

        # Currently active voice commands Box
        self.box_Currently_active_voice_commands = ctk.CTkFrame(parent_frame)
        self.box_Currently_active_voice_commands.pack(padx= 5, pady=5, side="left", expand=True, fill="both")

        self.label_Currently_active_voice_commands = ctk.CTkLabel(self.box_Currently_active_voice_commands, text=GetLocaleText("Currently_active_voice_commands"), font=self.app_header_font)
        self.label_Currently_active_voice_commands.pack(side="top")

        self.list_Currently_active_voice_commands = ctk.CTkTextbox(self.box_Currently_active_voice_commands, font=self.app_font)
        self.list_Currently_active_voice_commands.configure(state="disabled")
        self.list_Currently_active_voice_commands.pack(side="left", expand=True, fill="both")

    def build_voice_files_tab(self, parent_frame):
        # Voice files Box
        self.box_voice_files = ctk.CTkFrame(parent_frame)
        self.box_voice_files.pack(padx= 5, pady=5, side="left", expand=True, fill="both")

        self.box_voice_files_main = ctk.CTkFrame(self.box_voice_files, height=400)
        self.box_voice_files_main.pack(side="top", expand=True, fill="both")

        self.label_voice_files = ctk.CTkLabel(self.box_voice_files_main, text=GetLocaleText("In_Vocola_folder"), font=self.app_header_font)
        self.label_voice_files.pack(side="top")

        self.xscrollbar_box_voice_files = ctk.CTkScrollbar(self.box_voice_files_main, width=20, orientation="horizontal")
        self.xscrollbar_box_voice_files.pack(side="bottom", fill="x")

        self.yscrollbar_box_voice_files = ctk.CTkScrollbar(self.box_voice_files_main)
        self.yscrollbar_box_voice_files.pack(side="right", fill="y")

        self.cols_box_voice_files = (GetLocaleText("ID"), GetLocaleText("Filename"))
        self.list_voice_files = ttk.Treeview(self.box_voice_files_main, show="headings", columns=self.cols_box_voice_files, height=15, xscrollcommand=self.xscrollbar_box_voice_files.set, yscrollcommand=self.yscrollbar_box_voice_files.set)
        self.list_voice_files.heading("#1", text=self.cols_box_voice_files[0], anchor="w")
        self.list_voice_files.heading("#2", text=self.cols_box_voice_files[1], anchor="w")
        self.list_voice_files.column("#1", stretch=False, width=100)
        self.list_voice_files.column("#2", width=250)
        self.list_voice_files.pack(expand=True, fill="both")

        self.xscrollbar_box_voice_files.configure(command=self.list_voice_files.xview)
        self.yscrollbar_box_voice_files.configure(command=self.list_voice_files.yview)

        self.box_voice_files_btns = ctk.CTkFrame(self.box_voice_files)
        self.box_voice_files_btns.pack(side="top", expand=True, fill="both")

        self.btn_edit_voice_file = ctk.CTkButton(self.box_voice_files_btns, text=GetLocaleText("Edit_Voice_File"), width=200, height=50, font=self.app_header_font, state="disabled")
        self.btn_edit_voice_file.pack(padx=2, pady=2, side="top", expand=True, fill="both")

        self.btn_delete_voice_file = ctk.CTkButton(self.box_voice_files_btns, text=GetLocaleText("Delete_voice_file"), width=200, height=50, font=self.app_header_font, state="disabled")
        self.btn_delete_voice_file.pack(padx=2, pady=2, side="top", expand=True, fill="both")

        self.tooltip_delete_voice_file = CTkToolTip(self.btn_delete_voice_file, GetLocaleText("Delete_voice_file_Description"))

        # Voice cmd files Box
        self.box_voice_cmd_files = ctk.CTkFrame(parent_frame)
        self.box_voice_cmd_files.pack(padx= 5, pady=5, side="left", expand=True, fill="both")

        self.box_voice_cmd_files_main = ctk.CTkFrame(self.box_voice_cmd_files, height=400)
        self.box_voice_cmd_files_main.pack(side="top", expand=True, fill="both")

        self.label_voice_cmd_files = ctk.CTkLabel(self.box_voice_cmd_files_main, text=GetLocaleText("Vocola_Voice_Command_Language_Files"), font=self.app_header_font)
        self.label_voice_cmd_files.pack(side="top")

        self.xscrollbar_voice_cmd_files = ctk.CTkScrollbar(self.box_voice_cmd_files_main, width=20, orientation="horizontal")
        self.xscrollbar_voice_cmd_files.pack(side="bottom", fill="x")

        self.yscrollbar_voice_cmd_files = ctk.CTkScrollbar(self.box_voice_cmd_files_main)
        self.yscrollbar_voice_cmd_files.pack(side="right", fill="y")

        self.cols_voice_cmd_files = (GetLocaleText("ID"), GetLocaleText("Filename"))
        self.list_voice_cmd_files = ttk.Treeview(self.box_voice_cmd_files_main, show="headings", columns=self.cols_voice_cmd_files, height=15, xscrollcommand=self.xscrollbar_voice_cmd_files.set, yscrollcommand=self.yscrollbar_voice_cmd_files.set)
        self.list_voice_cmd_files.heading("#1", text=self.cols_voice_cmd_files[0], anchor="w")
        self.list_voice_cmd_files.heading("#2", text=self.cols_voice_cmd_files[1], anchor="w")
        self.list_voice_cmd_files.column("#1", stretch=False, width=100)
        self.list_voice_cmd_files.column("#2", width=210)
        self.list_voice_cmd_files.pack(expand=True, fill="both")

        self.xscrollbar_voice_cmd_files.configure(command=self.list_voice_cmd_files.xview)
        self.yscrollbar_voice_cmd_files.configure(command=self.list_voice_cmd_files.yview)

        self.box_voice_cmd_files_btns = ctk.CTkFrame(self.box_voice_cmd_files)
        self.box_voice_cmd_files_btns.pack(side="top", expand=True, fill="both")

        self.btn_download_voice_file = ctk.CTkButton(self.box_voice_cmd_files_btns, text=GetLocaleText("Download_to_Vocola_folder"), width=200, height=50, font=self.app_header_font, state="disabled")
        self.btn_download_voice_file.pack(padx=2, pady=2, side="top", expand=True, fill="both")

        self.tooltip_download_voice_file = CTkToolTip(self.btn_download_voice_file, GetLocaleText("Download_to_Vocola_folder_Description"))

    def build_external_pointers_tab(self, parent_frame):
        # External Pointer Box left
        self.box_external_pointers_left = ctk.CTkFrame(parent_frame)
        self.box_external_pointers_left.pack(padx= 5, pady=5, side="left", expand=True, fill="both")

        self.box_TIR_Left = ctk.CTkFrame(self.box_external_pointers_left)
        self.box_TIR_Left.pack(side="left", fill="x")

        self.TIR_LeftUp = ctk.CTkProgressBar(self.box_TIR_Left, orientation="vertical", width=30, corner_radius=0)
        self.TIR_LeftUp.pack()
        self.TIR_LeftUp.set(0)

        self.TIR_LeftLeft = ctk.CTkProgressBar(self.box_TIR_Left, orientation="horizontal", width=350, height=30, corner_radius=0)
        self.TIR_LeftLeft.pack()
        self.TIR_LeftLeft.set(0.544)

        self.TIR_LeftDown = ctk.CTkProgressBar(self.box_TIR_Left, orientation="vertical", width=30, corner_radius=0)
        self.TIR_LeftDown.pack()
        self.TIR_LeftDown.set(1)

        self.box_TIR_right = ctk.CTkFrame(self.box_external_pointers_left)
        self.box_TIR_right.pack(side="left", fill="x")

        self.TIR_RightUp = ctk.CTkProgressBar(self.box_TIR_right, orientation="vertical", width=30, corner_radius=0)
        self.TIR_RightUp.pack()
        self.TIR_RightUp.set(0)

        self.TIR_RightRight = ctk.CTkProgressBar(self.box_TIR_right, orientation="horizontal", width=350, height=30, corner_radius=0)
        self.TIR_RightRight.pack()
        self.TIR_RightRight.set(0.544)

        self.TIR_RightDown = ctk.CTkProgressBar(self.box_TIR_right, orientation="vertical", width=30, corner_radius=0)
        self.TIR_RightDown.pack()
        self.TIR_RightDown.set(1)

        # External Pointer Box right
        self.box_external_pointers_right = ctk.CTkFrame(parent_frame, width=400)
        self.box_external_pointers_right.pack(side="right", fill="y")

        self.label_external_pointers_settings = ctk.CTkLabel(self.box_external_pointers_right, text=GetLocaleText("External_Pointers_Settings"), font=self.app_header_font)
        self.label_external_pointers_settings.pack(side="top")

        # External Pointer Center Dead Zone 
        self.box_TIR_DeadZone = ctk.CTkFrame(self.box_external_pointers_right)
        self.box_TIR_DeadZone.pack(side="top", fill="x")

        self.label_TIR_DeadZone = ctk.CTkLabel(self.box_TIR_DeadZone, text=GetLocaleText("External_Pointer_Center_Dead_Zone"), font=self.app_font)
        self.label_TIR_DeadZone.pack(padx=(0, 5), side="left")

        self.TIR_DeadZone_value = ctk.IntVar(0)

        self.btn_TIR_DeadZone_decrement = ctk.CTkButton(self.box_TIR_DeadZone, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.TIR_DeadZone_value))
        self.btn_TIR_DeadZone_decrement.pack(side="left")

        self.entry_TIR_DeadZone = ctk.CTkEntry(self.box_TIR_DeadZone, textvariable=self.TIR_DeadZone_value, validate="key", validatecommand=(self.vcmd, "%P"), width=40)
        self.entry_TIR_DeadZone.pack(side="left")

        self.tooltip_TIR_DeadZone = CTkToolTip(self.entry_TIR_DeadZone, GetLocaleText("External_Pointer_Center_Dead_Zone_Description"))

        self.btn_TIR_DeadZone_increment = ctk.CTkButton(self.box_TIR_DeadZone, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.TIR_DeadZone_value, 100))
        self.btn_TIR_DeadZone_increment.pack(side="left")

        # Mouse Capture Settings

        self.label_mouse_capture_settings = ctk.CTkLabel(self.box_external_pointers_right, text=GetLocaleText("Mouse_Capture_Settings"), font=self.app_header_font)
        self.label_mouse_capture_settings.pack(pady=(15, 0), side="top")

        self.box_capture_mode = ctk.CTkFrame(self.box_external_pointers_right)
        self.box_capture_mode.pack(pady=5, side="top", anchor="w", fill="x")

        self.label_capture_mode = ctk.CTkLabel(self.box_capture_mode, text=GetLocaleText("Mouse_Capture_Mode"), anchor="w", font=self.app_font)
        self.label_capture_mode.pack(padx=(0, 5), side="left")

        self.capture_modes = [GetLocaleText("Off"), GetLocaleText("Position"), GetLocaleText("Motion")]
        self.choice_capture_mode = ctk.CTkOptionMenu(self.box_capture_mode, values=self.capture_modes, command=self.change_capture_mode)
        self.choice_capture_mode.pack(side="left")

        self.tooltip_capture_mode = CTkToolTip(self.choice_capture_mode, GetLocaleText("Mouse_Capture_Settings_Description"))

        # Mouse Capture Center

        self.box_mouse_capture_center = ctk.CTkFrame(self.box_external_pointers_right)
        self.box_mouse_capture_center.pack(pady=2, side="top", anchor="w", fill="x")

        # Mouse Capture Center X
        self.box_mouse_capture_center_x = ctk.CTkFrame(self.box_mouse_capture_center)
        self.box_mouse_capture_center_x.pack(padx=(0, 10), side="left", fill="x")

        self.label_mouse_capture_center_x = ctk.CTkLabel(self.box_mouse_capture_center_x, text=GetLocaleText("Mouse_Capture_Center_X"), font=self.app_header_font)
        self.label_mouse_capture_center_x.pack(side="top")

        self.mouse_capture_center_x_value = ctk.IntVar(value=0)

        self.btn_mouse_capture_center_x_decrement = ctk.CTkButton(self.box_mouse_capture_center_x, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.mouse_capture_center_x_value))
        self.btn_mouse_capture_center_x_decrement.pack(side="left")

        self.entry_mouse_capture_center_x = ctk.CTkEntry(self.box_mouse_capture_center_x, textvariable=self.mouse_capture_center_x_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_mouse_capture_center_x.pack(side="left")

        self.tooltip_mouse_capture_center_x = CTkToolTip(self.entry_mouse_capture_center_x, GetLocaleText("Mouse_Capture_Center_X_Description"))

        self.btn_mouse_capture_center_x_increment = ctk.CTkButton(self.box_mouse_capture_center_x, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.mouse_capture_center_x_value, 10000))
        self.btn_mouse_capture_center_x_increment.pack(side="left")

        # Mouse Capture Center Y
        self.box_mouse_capture_center_y = ctk.CTkFrame(self.box_mouse_capture_center)
        self.box_mouse_capture_center_y.pack(side="left", anchor="w", fill="x")

        self.label_mouse_capture_center_y = ctk.CTkLabel(self.box_mouse_capture_center_y, text=GetLocaleText("Mouse_Capture_Center_y"), width=30, font=self.app_header_font)
        self.label_mouse_capture_center_y.pack(side="top")

        self.mouse_capture_center_y_value = ctk.IntVar(value=0)

        self.btn_mouse_capture_center_y_decrement = ctk.CTkButton(self.box_mouse_capture_center_y, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.mouse_capture_center_y_value))
        self.btn_mouse_capture_center_y_decrement.pack(side="left")

        self.entry_mouse_capture_center_y = ctk.CTkEntry(self.box_mouse_capture_center_y, textvariable=self.mouse_capture_center_y_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_mouse_capture_center_y.pack(side="left")

        self.tooltip_mouse_capture_center_y = CTkToolTip(self.entry_mouse_capture_center_y, GetLocaleText("Mouse_Capture_Center_y_Description"))

        self.btn_mouse_capture_center_y_increment = ctk.CTkButton(self.box_mouse_capture_center_y, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.mouse_capture_center_y_value, 10000))
        self.btn_mouse_capture_center_y_increment.pack(side="left")

        # Mouse Capture Width and Height

        self.box_mouse_capture_width_height = ctk.CTkFrame(self.box_external_pointers_right)
        self.box_mouse_capture_width_height.pack(pady=2, side="top", anchor="w", fill="x")

        # Mouse Capture Width
        self.box_mouse_capture_width = ctk.CTkFrame(self.box_mouse_capture_width_height)
        self.box_mouse_capture_width.pack(padx=(0, 10), side="left", anchor="w", fill="x")

        self.label_mouse_capture_width = ctk.CTkLabel(self.box_mouse_capture_width, text=GetLocaleText("Mouse_Capture_Width"), font=self.app_header_font)
        self.label_mouse_capture_width.pack(side="top")

        self.mouse_capture_width_value = ctk.IntVar(value=0)

        self.btn_mouse_capture_width_decrement = ctk.CTkButton(self.box_mouse_capture_width, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.mouse_capture_width_value))
        self.btn_mouse_capture_width_decrement.pack(side="left")

        self.entry_mouse_capture_width = ctk.CTkEntry(self.box_mouse_capture_width, textvariable=self.mouse_capture_width_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_mouse_capture_width.pack(side="left")

        self.tooltip_mouse_capture_width = CTkToolTip(self.entry_mouse_capture_width, GetLocaleText("Mouse_Capture_Width_Description"))

        self.btn_mouse_capture_width_increment = ctk.CTkButton(self.box_mouse_capture_width, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.mouse_capture_width_value, 10000))
        self.btn_mouse_capture_width_increment.pack(side="left")

        # Mouse Capture Height
        self.box_mouse_capture_height = ctk.CTkFrame(self.box_mouse_capture_width_height)
        self.box_mouse_capture_height.pack(side="left", anchor="w", fill="x")

        self.label_mouse_capture_height = ctk.CTkLabel(self.box_mouse_capture_height, text=GetLocaleText("Mouse_Capture_Height"), font=self.app_header_font)
        self.label_mouse_capture_height.pack(side="top")

        self.mouse_capture_height_value = ctk.IntVar(value=0)

        self.btn_mouse_capture_height_decrement = ctk.CTkButton(self.box_mouse_capture_height, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.mouse_capture_height_value))
        self.btn_mouse_capture_height_decrement.pack(side="left")

        self.entry_mouse_capture_height = ctk.CTkEntry(self.box_mouse_capture_height, textvariable=self.mouse_capture_height_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_mouse_capture_height.pack(side="left")

        self.tooltip_mouse_capture_height = CTkToolTip(self.entry_mouse_capture_height, GetLocaleText("Mouse_Capture_Height_Description"))

        self.btn_mouse_capture_height_increment = ctk.CTkButton(self.box_mouse_capture_height, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.mouse_capture_height_value, 10000))
        self.btn_mouse_capture_height_increment.pack(side="left")

        # Mouse Capture Gain

        self.box_mouse_capture_gain = ctk.CTkFrame(self.box_external_pointers_right)
        self.box_mouse_capture_gain.pack(pady=2, side="top", anchor="w", fill="x")

        # Mouse Capture Gain X
        self.box_mouse_capture_gain_x = ctk.CTkFrame(self.box_mouse_capture_gain)
        self.box_mouse_capture_gain_x.pack(padx=(0, 10), side="left", anchor="w", fill="x")

        self.label_mouse_capture_gain_x = ctk.CTkLabel(self.box_mouse_capture_gain_x, text=GetLocaleText("Mouse_Capture_Gain_X"), font=self.app_header_font)
        self.label_mouse_capture_gain_x.pack(side="top")

        self.mouse_capture_gain_x_value = ctk.IntVar(value=0)

        self.btn_mouse_capture_gain_x_decrement = ctk.CTkButton(self.box_mouse_capture_gain_x, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.mouse_capture_gain_x_value))
        self.btn_mouse_capture_gain_x_decrement.pack(side="left")

        self.entry_mouse_capture_gain_x = ctk.CTkEntry(self.box_mouse_capture_gain_x, textvariable=self.mouse_capture_gain_x_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_mouse_capture_gain_x.pack(side="left")

        self.tooltip_mouse_capture_gain_x = CTkToolTip(self.entry_mouse_capture_gain_x, GetLocaleText("Mouse_Capture_Gain_X_Description"))

        self.btn_mouse_capture_gain_x_increment = ctk.CTkButton(self.box_mouse_capture_gain_x, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.mouse_capture_gain_x_value, 1000))
        self.btn_mouse_capture_gain_x_increment.pack(side="left")

        # Mouse Capture Gain Y
        self.box_mouse_capture_gain_y = ctk.CTkFrame(self.box_mouse_capture_gain)
        self.box_mouse_capture_gain_y.pack(padx=(0, 10), side="left", anchor="w", fill="x")

        self.label_mouse_capture_gain_y = ctk.CTkLabel(self.box_mouse_capture_gain_y, text=GetLocaleText("Mouse_Capture_Gain_Y"), font=self.app_header_font)
        self.label_mouse_capture_gain_y.pack(side="top")

        self.mouse_capture_gain_y_value = ctk.IntVar(value=0)

        self.btn_mouse_capture_gain_y_decrement = ctk.CTkButton(self.box_mouse_capture_gain_y, text="-", width=30, font=self.app_header_font, command=lambda: self.decrement(self.mouse_capture_gain_y_value))
        self.btn_mouse_capture_gain_y_decrement.pack(side="left")

        self.entry_mouse_capture_gain_y = ctk.CTkEntry(self.box_mouse_capture_gain_y, textvariable=self.mouse_capture_gain_y_value, validate="key", validatecommand=(self.vcmd, "%P"), width=50)
        self.entry_mouse_capture_gain_y.pack(side="left")

        self.tooltip_mouse_capture_gain_y = CTkToolTip(self.entry_mouse_capture_gain_y, GetLocaleText("Mouse_Capture_Gain_Y_Description"))

        self.btn_mouse_capture_gain_y_increment = ctk.CTkButton(self.box_mouse_capture_gain_y, text="+", width=30, font=self.app_header_font, command=lambda: self.increment(self.mouse_capture_gain_y_value, 1000))
        self.btn_mouse_capture_gain_y_increment.pack(side="left")

        self.box_mouse_capture_btn= ctk.CTkFrame(self.box_external_pointers_right)
        self.box_mouse_capture_btn.pack(pady=(15, 0), side="top", anchor="w", fill="x")

        self.btn_start_mouse_capture = ctk.CTkButton(self.box_mouse_capture_btn, text=GetLocaleText("Capture_Mouse_Location"), height=80, cursor="hand2", font=self.app_header_font, command=self.StartMouseCaptureEvent)
        self.btn_start_mouse_capture.pack(side="top", expand=True, fill="x")

        self.tooltip_start_mouse_capture = CTkToolTip(self.btn_start_mouse_capture, GetLocaleText("Capture_Mouse_Location_Description"))

    def recreate_tabs(self, tab_definitions, selected_index=0):
        # Clear old tabs
        for name in self.tabs._name_list[:]:
            self.tabs.delete(name)

        # Create new tabs
        for name, builder in tab_definitions:
            frame = self.tabs.add(name)
            builder(frame)

        # Select and force redraw
        # if 0 <= selected_index < len(tab_definitions):
        #     self.tabs.set(tab_definitions[selected_index][0])

    def increment(self, parent_value, max_value=None):
        try:
            current = parent_value.get()
            new = current + 1
            if max_value is None or new <= max_value:
                parent_value.set(new)
        except ValueError:
            return

    def decrement(self, parent_value, min_value=0):
        try:
            current = parent_value.get()
            new = current - 1
            if new >= min_value:
                parent_value.set(new)
        except ValueError:
            return

    def load_initial_values(self):
        global settings
        global preferences

        self.SendConsoleMessage("Version: " + VERSION)

        # TODO Vocola
        # vocola_installed = os.path.isdir(VocolaPath)
        # if not vocola_installed:
        #     self.notebook_voice_files.Disable()

        # print("global settings: ", repr(settings))

        # Initialize preferences with the last ones saved before
        # Then try to read them from the quadstick
        prefs = settings.get("preferences")
        # print("previous preferences: ", repr(prefs))
        if prefs:
            preferences.clear()
            preferences.update(prefs)
        d = find_quadstick_drive(True)
        if d is None: # QuadStick not plugged into pc
            self.SendConsoleMessage("QuadStick flash drive not found.")
            # self.button_reload.Disable()
        else:
            self.SendConsoleMessage("QuadStick drive letter: " + d[:2])
        if load_preferences_file(self) is not None: # Try both flash and ssp access to prefs files
            # Update status box
            self.SendConsoleMessage("Loaded preferences OK")
            # Send telemetry for QMP settings
            settings['preferences'] = preferences
        else:
            # Disable any tabs that need the quadstick
            self.SendConsoleMessage("Using previously saved preference values")
        # if d is None:
            # self.download_selected_build.Disable()

        # Make sure mouse capture settings have initial default values
        settings['mouse_capture_mode'] = settings.get('mouse_capture_mode', "Off").strip()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        settings['mouse_center_x'] = int(settings.get('mouse_center_x', screen_width/2))
        settings['mouse_center_y'] = int(settings.get('mouse_center_y', screen_height/2))
        settings['mouse_width']    = int(settings.get('mouse_width', screen_width))
        settings['mouse_height']   = int(settings.get('mouse_height', screen_height))
        settings['mouse_gain_x']   = int(settings.get('mouse_gain_x', 100))
        settings['mouse_gain_y']   = int(settings.get('mouse_gain_y', 100))

        # mouse capture settings
        self.choice_capture_mode.set(GetLocaleText(settings['mouse_capture_mode']))
        self.mouse_capture_center_x_value.set(settings['mouse_center_x'])
        self.mouse_capture_center_y_value.set(settings['mouse_center_y'])
        self.mouse_capture_width_value.set(settings['mouse_width'])
        self.mouse_capture_height_value.set(settings['mouse_height'])
        self.mouse_capture_gain_x_value.set(settings['mouse_gain_x'])
        self.mouse_capture_gain_y_value.set(settings['mouse_gain_y'])
        # clear and force update of quadstick factory game and voice files
        # self._game_profiles = [] #settings.get('game_profiles', [])
        # self._voice_files = [] #settings.get('voice_files', [])
        self._read_online_files_flag = True # tells widget to read files the first time program runs the Notebook page with the csv file list
        # self._csv_files = []

        print("update controls")
        self.updateControls(True)
        print("controls updated")
        return True

    def updateControls(self, fast=False):
        # self.button_delete_csv.Disable()
        # self.button_load_and_run.Disable()
        # self.button_remove_user_game.Disable()

        # Set up sliders for joystick
        joystick_deflection_minimum = int(preferences.get('joystick_deflection_minimum', defaults['joystick_deflection_minimum']))
        joystick_deflection_maximum = int(preferences.get('joystick_deflection_maximum', defaults['joystick_deflection_maximum']))
        deflection_multiplier_up    = int(preferences.get('deflection_multiplier_up', defaults['deflection_multiplier_up']))
        deflection_multiplier_down  = int(preferences.get('deflection_multiplier_down', defaults['deflection_multiplier_down']))
        deflection_multiplier_left  = int(preferences.get('deflection_multiplier_left', defaults['deflection_multiplier_left']))
        deflection_multiplier_right = int(preferences.get('deflection_multiplier_right', defaults['deflection_multiplier_right']))

        self.slider_NEUTRAL.set(joystick_deflection_minimum)
        self.label_slider_NEUTRAL_value.configure(text=int(joystick_deflection_minimum))

        self.slider_UP.set(int((joystick_deflection_maximum * 100) / deflection_multiplier_up))
        self.label_slider_UP_value.configure(text=int((joystick_deflection_maximum * 100) / deflection_multiplier_up))

        self.slider_DOWN.set(int((joystick_deflection_maximum * 100) / deflection_multiplier_down))
        self.label_slider_DOWN_value.configure(text=int((joystick_deflection_maximum * 100) / deflection_multiplier_down))
    
        self.slider_LEFT.set(int((joystick_deflection_maximum * 100) / deflection_multiplier_left))
        self.label_slider_LEFT_value.configure(text=int((joystick_deflection_maximum * 100) / deflection_multiplier_left))

        self.slider_RIGHT.set(int((joystick_deflection_maximum * 100) / deflection_multiplier_right))
        self.label_slider_RIGHT_value.configure(text=int((joystick_deflection_maximum * 100) / deflection_multiplier_right))

        self.update_joystick_preference_grid()

        # Set up sliders for D_Pad
        self.slider_D_Pad_outer_ring.set(int(preferences.get('joystick_D_Pad_outer', defaults['joystick_D_Pad_outer'])))
        self.slider_D_Pad_outer_ring_event(int(preferences.get('joystick_D_Pad_outer', defaults['joystick_D_Pad_outer'])))

        self.slider_D_Pad_inner_ring.set(int(preferences.get('joystick_D_Pad_inner', defaults['joystick_D_Pad_inner'])))
        self.slider_D_Pad_inner_ring_event(int(preferences.get('joystick_D_Pad_inner', defaults['joystick_D_Pad_inner'])))

        # Set up sliders and spinners for Sip/Puff
        self.slider_SP_low.set(int(preferences.get('sip_puff_threshold_soft', defaults['sip_puff_threshold_soft'])))
        self.slider_SP_low_event(int(preferences.get('sip_puff_threshold_soft', defaults['sip_puff_threshold_soft'])))

        self.slider_SP_high.set(int(preferences.get('sip_puff_threshold', defaults['sip_puff_threshold'])))
        self.slider_SP_high_event(int(preferences.get('sip_puff_threshold', defaults['sip_puff_threshold'])))

        self.slider_SP_max.set(int(preferences.get('sip_puff_maximum', defaults['sip_puff_maximum'])))
        self.slider_SP_max_event(int(preferences.get('sip_puff_maximum', defaults['sip_puff_maximum'])))

        self.SP_low_delay_value.set(preferences.get('sip_puff_delay_soft', defaults['sip_puff_delay_soft']))
        self.SP_high_delay_value.set(preferences.get('sip_puff_delay_hard', defaults['sip_puff_delay_hard']))

        # Set up sliders for Lip
        self.slider_Lip_max.set(int(preferences.get('lip_position_maximum', defaults['lip_position_maximum'])))
        self.slider_Lip_max_event(int(preferences.get('lip_position_maximum', defaults['lip_position_maximum'])))

        self.slider_Lip_min.set(int(preferences.get('lip_position_minimum', defaults['lip_position_minimum'])))
        self.slider_Lip_min_event(int(preferences.get('lip_position_minimum', defaults['lip_position_minimum'])))

        # Set up mouse, volume and brightness
        self.slider_mouse_speed.set(int(preferences.get('mouse_speed', defaults['mouse_speed'])))
        self.slider_mouse_speed_event(int(preferences.get('mouse_speed', defaults['mouse_speed'])))

        self.slider_brightness.set(int(preferences.get('brightness', defaults['brightness'])))
        self.slider_brightness_event(int(preferences.get('brightness', defaults['brightness'])))

        self.slider_volume.set(int(preferences.get('volume', defaults['volume'])))
        self.slider_volume_event(int(preferences.get('volume', defaults['volume'])))

        # Set up digital outputs
        self.checkbox_do_1_state.set(int(preferences.get('digital_out_1', defaults['digital_out_1'])) > 0)
        self.checkbox_do_2_state.set(int(preferences.get('digital_out_2', defaults['digital_out_2'])) > 0)

        # Set up bluetooth
        device_mode = self.GetDeviceModeName(preferences.get('bluetooth_device_mode', defaults['bluetooth_device_mode']).strip())
        self.choice_BT_device_mode.set(device_mode)

        choices = self.BT_auth_modes
        choice_index = choices.index(preferences.get('bluetooth_authentication_mode', defaults['bluetooth_authentication_mode']).strip())
        self.choice_BT_auth_mode.set(choice_index)

        # choices = self.choice_BT_connection_modes
        # choice_index = choices.index(preferences.get('bluetooth_connection_mode', defaults['bluetooth_connection_mode']).strip())
        # self.choice_BT_connection_mode.Select(choice_index)

        # remote_BTA = preferences.get('bluetooth_remote_address', "").strip()
        # self.text_ctrl_BTA_remote_address.insert(0, remote_BTA)

        # if (preferences.get('bluetooth_connection_mode') == 'auto'):
            # self.text_ctrl_BTA_remote_address.Enable()
            # self.BTA_label.Enable()
            # self.text_ctrl_BTA_remote_address.Show()
            # self.BTA_label.Show()
        # else:
            # self.text_ctrl_BTA_remote_address.Disable()
            # self.BTA_label.Disable()
            # self.text_ctrl_BTA_remote_address.Hide()
            # self.BTA_label.Hide()

        # Set up misc preferences
        self.checkbox_select_files_state.set(int(preferences.get('enable_select_files', defaults['enable_select_files'])) > 0)
        self.checkbox_swap_state.set(int(preferences.get('enable_swap_inputs', defaults['enable_swap_inputs'])) > 0)
        self.checkbox_circular_deadzone_state.set(int(preferences.get('joystick_dead_zone_shape', defaults['joystick_dead_zone_shape'])) > 0)
        self.checkbox_ps4_boot_mode_state.set(int(preferences.get('enable_DS3_emulation',defaults['enable_DS3_emulation'])) > 0)
        self.checkbox_usb_A_host_mode_state.set(int(preferences.get('enable_usb_a_host',defaults['enable_usb_a_host'])) > 0)
        self.checkbox_Titan2_state.set(int(preferences.get('titan_two',defaults['titan_two'])) > 0)

        choices = self.mouse_responses
        choice_index = int(preferences.get('mouse_response_curve', defaults['mouse_response_curve']))
        choice = choices[choice_index] # Get item name by index
        self.choice_mouse_response.set(choice)

        usb_comm = int(preferences.get('enable_usb_comm', defaults['enable_usb_comm'])) > 0
        self.checkbox_usb_comm_state.set(usb_comm)

        # Get list of csv files
        # if not self.list_csv_files.GetColumnCount(): # Prevent second call here from addingmore columns
        #     self.list_csv_files.InsertColumn(0, "#")
        #     self.list_csv_files.InsertColumn(1, GetLocaleText("Filename"))
        #     self.list_csv_files.InsertColumn(2, GetLocaleText("Spreadsheet"))

        # self.update_quadstick_flash_files_items(fast)

        # vocola_installed = os.path.isdir(VocolaPath)
        # if vocola_installed:
        #     # get list of voice files
        #     x = list_voice_files()
        #     self.list_box_voice_files.Clear()
        #     self.list_box_voice_files.InsertItems(x, 0)

        # Prepare list of online factory game files widget
        # if not self.online_game_files_list.GetColumnCount(): # prevent second call here from addingmore columns
        #     self.online_game_files_list.InsertColumn(0, GetLocaleText("Filename"))
        #     self.online_game_files_list.InsertColumn(1, GetLocaleText("Spreadsheet"))

        # if vocola_installed:
        #     # prepare list of online voice files widget
        #     if not self.online_voice_files_list.GetColumnCount():
        #         self.online_voice_files_list.InsertColumn(0, GetLocaleText("Filename"))
        #         self.online_voice_files_list.InsertColumn(1, "game name")
        # self.update_online_game_files_list_items() # updates widget, not actual list

        # if vocola_installed:
        #     self.update_online_voice_files_list_items() # updates widget, not actual list

        # Initialize any user confuration files
        # if not self.user_game_files_list.GetColumnCount(): # prevent second call here from addingmore columns
        #     self.user_game_files_list.InsertColumn(0, GetLocaleText("Filename"))
        #     self.user_game_files_list.InsertColumn(1, GetLocaleText("Spreadsheet"))
        # self.update_user_game_files_list_items() # updates widget from settings["user_game_profiles"]

        # Init virtual gamebus settings
        self.checkbox_enable_vg_DS4_state.set(settings.get('enable_VG4', 0))
        self.checkbox_enable_vg_X360_state.set(settings.get('enable_VGX', 0))

        # Set up locale   TODO Remove?
        # choices = self.locales
        # choice_index = choices.index(GetLocaleText(settings.get('current_locale')))
        # choice = choices[choice_index] # Get item name by index
        # self.choice_locale.set(choice)

        self.checkbox_log_state.set(settings.get('enable_log', 0))

        self.checkbox_minimize_to_tray_state.set(settings.get('minimize_to_tray', 0))
        self.checkbox_start_minimized_state.set(settings.get('start_mimimized', 0))

        # Init serial connection enable
        self.checkbox_enable_serial_port_state.set(settings.get('enable_serial_port', 1))

        # Set up external pointers tab
        self.TIR_DeadZone_value.set(settings.get('TIR_DeadZone', 0))
        # self.checkbox_trackir_start_state.set(int(settings.get('TIR_Window', 0)))

        # Firmware tab widgets
        build_number = quadstick_drive_serial_number(self)
        # self.build_number_text.SetValue(str(build_number))
        self._available_firmware_list = None

        # if not self.list_ctrl_firmware.GetColumnCount():
        #     self.list_ctrl_firmware.InsertColumn(0, GetLocaleText("Build"))
        #     self.list_ctrl_firmware.InsertColumn(1, GetLocaleText("Remark"))

        if build_number is None or build_number < 1215:
            preferences["enable_DS3_emulation"] = "0" # Make sure DS4 mode off
        #     self.checkbox_ps4_boot_mode.Disable()
        # else:
        #     self.checkbox_ps4_boot_mode.Enable()

        # if build_number is not None and build_number < 1301:
        #     if US1:
        #         self.SendConsoleMessage("You will need to update the Firmware to use the UltraStik")
                
        # if vocola_installed:
        #     try:
        #         generate_includes_vch_file()
        #         self.InitializeWordList('')
        #     except Exception as e:
        #         print("generate_includes_vch_file exception in update controls: ", repr(e))

        # Determine if load_and_run should be enabled or disabled
        # If a com port previously used, usb comm enabled and the qs is plugged in
        # or if bluetooth ssp is enabled, then allow button to run
        # d = find_quadstick_drive()        
        # if not ( d or (( preferences.get('bluetooth_device_mode') == 'ssp' ) or has_serial_ports() or usb_comm )):
        #     self.button_load_and_run.Disable()
        #     self.button_download_csv.Disable()
        #     self.button_save.Disable()

        check_for_newer_version(self)
        return True # Indicate good read

    def ToggleHIDHideStatus(self, event):
        print("Event handler 'ToggleHIDHideStatus' ")
        widget = event.widget
        flag = widget.get()
        try:
            if flag:
                H.hide_quadstick(QS)
            else:
                H.unhide_quadstick(QS)
        except:
            pass

    def on_USB_status_timer(self): # periodically checks the USB status of the Quadstick
        # print ("Check USB status")
        try:
            if QS._qs is None:
                QS.open()
                if VG:
                    VG.reset()
            else:
                if not QS._qs.is_plugged():
                    QS.close()
                    self.SendConsoleMessage("Quadstick disconnected")
        except Exception as e:
            # Comment out to prevent spamming the log to the point it is too long to be sent by a debug reportAdd commentMore actions
            # print ('USB status exception: ', repr(e))
            # print (traceback.format_exc())
            pass
        self.after(3000, self.on_USB_status_timer)

def initialize():
    read_repr_file() # Load global settings

    # Initialize default locale if not exists in settings
    if not settings.get('current_locale'):
        settings['current_locale'] = "en"

    # Initialize default appearance mode if not exists in settings
    if not settings.get('appearance_mode'):
        settings['appearance_mode'] = "System"

    # Set up locale
    global current_locale
    current_locale = settings.get('current_locale')

    # Set fallback locale if other locale not found.
    if not cparser.read('./locales/' + current_locale + '.INI', 'utf-8'):
        cparser.read('./locales/en.INI', 'utf-8')

    # Set up appearance mode
    global current_appearance_mode
    current_appearance_mode = settings.get('appearance_mode')

    # Set up log
    global enable_log

    # Initialize default log state if not exists in settings
    if not settings.get('enable_log'):
        settings['enable_log'] = enable_log

    enable_log = settings.get('enable_log')

if '-debug' in sys.argv:
    DEBUG = True

if __name__ == "__main__":
    initialize()
    app = QuadStickConfigurator()
    app.mainloop()
