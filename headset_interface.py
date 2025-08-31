# Author: TheRealRazbi (https://github.com/TheRealRazbi)
# License: MPL-2.0
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

import logging
import os
import threading
import time
import traceback
from enum import Enum

import hid  # requires installing specifically https://pypi.org/project/hidapi/
import win32con
import win32gui
from dotenv import load_dotenv
from pyee.asyncio import AsyncIOEventEmitter
from flask import Flask, jsonify
from flask_cors import CORS
from waitress import serve

from razbi_utils.core import show_window

VENDOR_ID = 2385
PRODUCT_ID = 5866
USAGE_PAGE = 65299
WINDOW_TITLE = 'Check Battery Server'

load_dotenv()
PORT = 9833
HOST = os.getenv('HEADSET_SERVER_HOST', 'localhost')

app = Flask(__name__)
flask_logger = logging.getLogger('werkzeug')
flask_logger.setLevel(logging.ERROR)
CORS(app)


@app.route('/battery_status')
def battery_status():
    if not hasattr(app, 'battery_status') or app.battery_status is None:
        print("Was requested for battery status, but it's not available.")
        return jsonify({'battery_status': 'unknown'})
    print(f"Returning battery status: {app.battery_status}")
    return jsonify({'battery_status': app.battery_status})


# todo: add a websocket endpoint for the app to announce to the listener

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


class CommandMeaning(Enum):
    """
    The positions in the received data have the following meaning.
    """
    position_info = 3  # represents the type of information, such as battery status, volume, etc.
    position_battery = 7  # represents the battery status
    type_battery_status = 2


class HyperXCloudFlightS(AsyncIOEventEmitter):
    def __init__(self, flask_app, debug=False, update_delay=300, invalidate_after=1800):
        super().__init__()
        self.debug = debug
        self.update_delay = update_delay
        self.invalidate_after = invalidate_after
        self.devices = [d for d in hid.enumerate(VENDOR_ID, PRODUCT_ID)]
        if not self.devices:
            print("HyperX Cloud Flight S was not found. Searching for it...")
            # raise Exception('HyperX Cloud Flight S was not found')
        self.bootstrap_device = None
        self.interval = None
        self.battery_status = None
        self.app = flask_app
        self.flask_thread = None
        self.last_battery_update_at = None
        self.device_ready = threading.Event()
        self.bootstrap()

    def start_flask_server(self):
        def run_server():
            print(f"Starting server at http://{HOST}:{PORT}")
            serve(self.app, host=HOST, port=PORT)

        self.flask_thread = threading.Thread(target=run_server, daemon=True)
        self.flask_thread.start()

    def bootstrap(self):
        try:
            if self.last_battery_update_at is not None and time.time() - self.last_battery_update_at > self.invalidate_after:
                print(f"Invalidating battery status at {time.strftime('%H:%M:%S')}")
                self.app.battery_status = None
                self.device_ready.clear()
                self.bootstrap_device = None  # invalidating the device to re-find it if it got unplugged

            if self.bootstrap_device is None:
                if not self.devices:
                    self.devices = [d for d in hid.enumerate(VENDOR_ID, PRODUCT_ID)]

                for device in self.devices:
                    if device['usage_page'] == USAGE_PAGE and device['usage'] == 1:
                        self.bootstrap_device = hid.device()
                        try:
                            self.bootstrap_device.open_path(device['path'])
                        except OSError:
                            self.bootstrap_device = None
                            continue
                        break

            if self.bootstrap_device is None:
                print(f"Searched for headset at {time.strftime('%H:%M:%S')}. Not found.")
                return

            if self.device_ready.is_set():
                return  # prevent re-running the check and prompting the headset unnecessarily
            try:
                buffer = [0x21] + [0x00] * 19
                self.bootstrap_device.write(buffer)
                print(f"Searched for headset at {time.strftime('%H:%M:%S')}. Found.")
                self.device_ready.set()
            except OSError:
                print("Had an OSError, was the pc was set to sleep just then, or the device got unplugged?")
            except Exception as e:
                self.emit('error', e)
        finally:
            self.interval = threading.Timer(self.update_delay, self.bootstrap)
            # ^ re-define this to have it check over and over. Don't check if it's None
            self.interval.start()

    def run(self):
        print(f"Started at {time.strftime('%X')}")
        self.device_ready.wait()
        while True:
            if self.bootstrap_device is None:
                self.bootstrap()
                self.device_ready.wait()
            try:
                data = self.bootstrap_device.read(20)
            except OSError:
                print("Had an OSError, was the pc was set to sleep just then, or the device got unplugged?")
                self.bootstrap_device = None
                self.device_ready.clear()
                continue
            if self.debug:
                print(f"{data} length: {len(data)}")
            if data:
                self.process_data(data)

    def process_data(self, data):
        if data[CommandMeaning.position_info.value] == CommandMeaning.type_battery_status.value:
            previous_battery_status = self.battery_status
            self.battery_status = data[CommandMeaning.position_battery.value]
            if previous_battery_status == self.battery_status:
                self.last_battery_update_at = time.time()  # prevent invalidating the status if the battery didn't decrease
                return
            self.emit('battery_status', self.battery_status)
            # if self.debug:
            self.app.battery_status = self.battery_status
            print(f"Battery status: {self.battery_status}")
            self.last_battery_update_at = time.time()


if __name__ == "__main__":
    hide_window(WINDOW_TITLE)
    headset = None
    try:
        headset = HyperXCloudFlightS(flask_app=app, debug=False, update_delay=300)
        headset.start_flask_server()
        headset.run()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        traceback.print_exc()
        show_window(WINDOW_TITLE)
        # if headset.flask_thread is not None:
        #     headset.flask_thread._stop()
        input("Press Enter to close...")
