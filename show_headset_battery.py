# Author: TheRealRazbi (https://github.com/TheRealRazbi)
# License: MPL-2.0
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

import os
import threading
import time
import traceback

import requests
import win32con
import win32gui
import winsound
from PIL import Image, ImageDraw, ImageFont
from dotenv import load_dotenv
from pystray import Icon, MenuItem

from razbi_utils.core import toggle_visibility, show_window

visible = False
WINDOW_TITLE = 'Show Battery As An Icon'
BATTERY_VERY_LOW_AT = 15
BATTERY_LOW_AT = 30
BATTERY_CHARGED_AT = 85


def hide_window(in_title: str, actually_show_it_instead=False):
    program_hwnd = None

    def win_enum_handler(hwnd, ctx) -> [None, int]:
        title = win32gui.GetWindowText(hwnd)
        if in_title in title:
            nonlocal program_hwnd
            program_hwnd = hwnd
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
            if actually_show_it_instead:
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

    win32gui.EnumWindows(win_enum_handler, None)
    return program_hwnd


def create_image(charge_level, width=64, height=64):
    text_color = (255, 255, 255)
    if isinstance(charge_level, int):
        if charge_level <= BATTERY_VERY_LOW_AT:
            text_color = (255, 66, 88)  # red
        elif charge_level <= BATTERY_LOW_AT:
            text_color = (255, 174, 0)  # orange
        elif charge_level >= BATTERY_CHARGED_AT:
            text_color = (102, 255, 132)  # green

    image = Image.new('RGBA', (width, height), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    font_size = height
    font_path = "arial.ttf"
    text = str(charge_level)

    text_width, text_height = 0, 0

    # reduce the font size until the text fits the image
    while font_size > 0:
        try:
            font = ImageFont.truetype(font_path, font_size)
        except IOError:
            font = ImageFont.load_default()
            font_size -= 1
            continue

        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        if text_width <= width and text_height <= height:
            break
        font_size -= 1

    x = (width - text_width) / 2
    y = (height - text_height) / 12

    draw.text((x, y), text, font=font, fill=text_color)

    return image


def update_icon(icon, charge_level):
    if charge_level == -1:
        icon.visible = False
        return
    if icon.visible is False:
        icon.visible = True
    icon.icon = create_image(charge_level)
    icon.title = f'Headset Battery: {charge_level}%'


def setup(icon):
    icon.running = True
    host = os.getenv('HEADSET_CLIENT_HOST', 'localhost')

    def update():
        last_charge_level = -1
        while icon.running:
            charge_level = get_battery_level(host)
            if charge_level != last_charge_level:
                print(f"Updating icon with charge level: {charge_level}")
                if (0 < last_charge_level < charge_level and charge_level >= BATTERY_CHARGED_AT) \
                        or (0 < last_charge_level > charge_level and charge_level <= BATTERY_VERY_LOW_AT):
                    # ^ bigger than one to avoid playing the sfx when the charge level is -1 aka None
                    winsound.MessageBeep(
                        winsound.MB_OK)  # TODO: change it to a more distinct sound or make the sfx work
                last_charge_level = charge_level
                update_icon(icon, charge_level)
            time.sleep(60)

    update_thread = threading.Thread(target=update, daemon=True)
    update_thread.start()


def get_battery_level(host="localhost"):
    try:
        res = requests.get(f"http://{host}:9833/battery_status")

        battery_level = res.json()['battery_status']
        return int(battery_level)  # raises ValueError if status is not a number
    except (requests.exceptions.ConnectionError, ValueError):
        print("Failed to request battery status")
        return -1


def main():
    print("Running Headset Battery Icon")
    load_dotenv()
    icon_global = Icon('HeadsetBattery',
                       icon=create_image(-1),
                       menu=[MenuItem('Show/Hide', lambda: toggle_visibility(WINDOW_TITLE)),
                             MenuItem('Exit',
                                      lambda icon: icon.stop())]
                       ).run(setup)


if __name__ == '__main__':
    try:
        hide_window(WINDOW_TITLE)
        main()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        traceback.print_exc()
        show_window(WINDOW_TITLE)
        input("Press Enter to close...")
