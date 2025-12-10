import ctypes
from ctypes import wintypes
import json
import os
import sys

try:
    import pystray
    from pystray import MenuItem as item
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "pystray", "pillow", "-q"])
    import pystray
    from pystray import MenuItem as item
    from PIL import Image, ImageDraw, ImageFont

try:
    from pywinusb import hid
except ImportError:
    import subprocess
    subprocess.run([sys.executable, "-m", "pip", "install", "pywinusb", "-q"])
    from pywinusb import hid

VENDOR_ID = 0x3554
PRODUCT_ID = 0xF510

POLLING_RATES = {
    1000: bytes.fromhex("08070000000201540000000000000000ef"),
    4000: bytes.fromhex("08070000000220350000000000000000ef"),
    8000: bytes.fromhex("08070000000240150000000000000000ef"),
}

GENERIC_READ = 0x80000000
GENERIC_WRITE = 0x40000000
FILE_SHARE_READ = 0x01
FILE_SHARE_WRITE = 0x02
OPEN_EXISTING = 3

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "polling_rate.json")


class PollingTray:
    def __init__(self):
        self.current_rate = self.load_rate()
        self.device_path = self.find_device()
        self.icon = None
    
    def load_rate(self):
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r") as f:
                    return json.load(f).get("rate", 1000)
        except:
            pass
        return 1000
    
    def save_rate(self, rate):
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump({"rate": rate}, f)
        except:
            pass
    
    def find_device(self):
        devices = hid.HidDeviceFilter(vendor_id=VENDOR_ID, product_id=PRODUCT_ID).get_devices()
        for device in devices:
            try:
                device.open()
                out_reports = [r.report_id for r in device.find_output_reports()]
                device.close()
                if 0x08 in out_reports:
                    return device.device_path
            except:
                pass
        return None
    
    def send_command(self, data):
        if not self.device_path:
            self.device_path = self.find_device()
        if not self.device_path:
            return False
        
        kernel32 = ctypes.windll.kernel32
        hid_dll = ctypes.windll.hid
        
        handle = kernel32.CreateFileW(
            self.device_path,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None, OPEN_EXISTING, 0, None
        )
        
        if handle == -1:
            return False
        
        try:
            report = (ctypes.c_ubyte * len(data))(*data)
            return bool(hid_dll.HidD_SetOutputReport(handle, ctypes.byref(report), len(data)))
        finally:
            kernel32.CloseHandle(handle)
    
    def set_rate(self, rate):
        packet = POLLING_RATES.get(rate)
        if packet and self.send_command(packet):
            self.current_rate = rate
            self.save_rate(rate)
            self.update_icon()
            return True
        return False
    
    def create_icon_image(self):
        size = 64
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        if self.current_rate == 8000:
            bg_color = (200, 60, 60)
        elif self.current_rate == 4000:
            bg_color = (0, 180, 100)
        else:
            bg_color = (80, 80, 180)
        
        draw.ellipse([2, 2, size-2, size-2], fill=bg_color)
        
        text = "8K" if self.current_rate == 8000 else ("4K" if self.current_rate == 4000 else "1K")
        
        try:
            font = ImageFont.truetype("arial.ttf", 24)
        except:
            font = ImageFont.load_default()
        
        bbox = draw.textbbox((0, 0), text, font=font)
        text_w = bbox[2] - bbox[0]
        text_h = bbox[3] - bbox[1]
        x = (size - text_w) // 2
        y = (size - text_h) // 2 - 2
        
        draw.text((x, y), text, fill="white", font=font)
        
        return img
    
    def update_icon(self):
        if self.icon:
            self.icon.icon = self.create_icon_image()
            self.icon.title = f"LAMZU: {self.current_rate} Hz"
    
    def on_1000(self, icon, item):
        self.set_rate(1000)
    
    def on_4000(self, icon, item):
        self.set_rate(4000)
    
    def on_8000(self, icon, item):
        self.set_rate(8000)
    
    def on_quit(self, icon, item):
        icon.stop()
    
    def run(self):
        menu = pystray.Menu(
            item('1000 Hz', self.on_1000, checked=lambda item: self.current_rate == 1000),
            item('4000 Hz', self.on_4000, checked=lambda item: self.current_rate == 4000),
            item('8000 Hz', self.on_8000, checked=lambda item: self.current_rate == 8000),
            pystray.Menu.SEPARATOR,
            item('Exit', self.on_quit)
        )
        
        self.icon = pystray.Icon(
            "lamzu_polling",
            self.create_icon_image(),
            f"LAMZU: {self.current_rate} Hz",
            menu
        )
        
        self.icon.run()


if __name__ == "__main__":
    app = PollingTray()
    app.run()
