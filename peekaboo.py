import win32gui
import win32ui
import win32con
import win32api
import win32com.client
from PIL import Image
import io
import requests
import time
import argparse
import subprocess
import win32api
import win32con

def get_desktop_dimensions():
    monitor_info = win32api.EnumDisplayMonitors(None, None)
    primary_monitor = monitor_info[0][0]  # Assuming the first monitor is the primary one
    monitor_info_ex = win32api.GetMonitorInfo(primary_monitor)
    rect = monitor_info_ex['Monitor']
    width = rect[2] - rect[0]
    height = rect[3] - rect[1]
    # print(width, height)
    return width, height

def main(host, key):
    r = requests.post(host+'/new_session', json={'_key': key})
    if r.status_code != 200:
        print('Server not available.')
        return

    shell = win32com.client.Dispatch('WScript.Shell')
    PREV_IMG = None
    while True:
        width, height = get_desktop_dimensions()

        # device context
        desktop_dc = win32gui.GetWindowDC(win32gui.GetDesktopWindow())
        img_dc = win32ui.CreateDCFromHandle(desktop_dc)

        # memory context
        mem_dc = img_dc.CreateCompatibleDC()

        screenshot = win32ui.CreateBitmap()
        screenshot.CreateCompatibleBitmap(img_dc, width, height)
        mem_dc.SelectObject(screenshot)

        bmpinfo = screenshot.GetInfo()

        # copy into memory
        mem_dc.BitBlt((0, 0), (width, height), img_dc, (0, 0), win32con.SRCCOPY)

        bmpstr = screenshot.GetBitmapBits(True)

        pillow_img = Image.frombytes('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX')

        with io.BytesIO() as image_data:
            pillow_img.save(image_data, 'PNG')
            image_data_content = image_data.getvalue()

        if image_data_content != PREV_IMG:
            files = {}
            filename = str(round(time.time()*1000))+'_'+key
            files[filename] = ('img.png', image_data_content, 'multipart/form-data')

            try:
                r = requests.post(host+'/capture_post', files=files)
            except Exception as e:
                pass

            PREV_IMG = image_data_content
        else:
            # print('no desktop change')
            pass

        # events
        try:
            r = requests.post(host+'/events_get', json={'_key': key})
            data = r.json()
            for e in data['events']:
                print(e)

                if e['type'] == 'click':
                    win32api.SetCursorPos((e['x'], e['y']))
                    time.sleep(0.1)
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, e['x'], e['y'], 0, 0)
                    time.sleep(0.02)
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, e['x'], e['y'], 0, 0)

                if e['type'] == 'keydown':
                    cmd = ''

                    if e['shiftKey']:
                        cmd += '+'

                    if e['ctrlKey']:
                        cmd += '^'

                    if e['altKey']:
                        cmd += '%'

                    if len(e['key']) == 1:
                        cmd += e['key'].lower()
                    else:
                        cmd += '{'+e['key'].upper()+'}'

                    print(cmd)
                    shell.SendKeys(cmd)

        except Exception as err:
            print(err)
            pass

        # free
        mem_dc.DeleteDC()
        win32gui.DeleteObject(screenshot.GetHandle())
        time.sleep(0.2)

print('hello')
parser = argparse.ArgumentParser(description='pyRD')
import getpass
username = getpass.getuser()
print("Username:", username)
parser.add_argument('addr', help='server address', type=str)
# parser.add_argument('key', help='access key', type=str)


import requests
import threading
import time

base_url = '<http://server_ip>:<server_port>'

def send_request():
    while True:
        try:
            # Define the data to be sent
            data = {
                'key1': username,
            }

            # Send the GET request with data
            response = requests.get(base_url + '/checkUsernames', params=data)
            print(response.text)
        except Exception as e:
            print("Error:", e)
        
        # Wait for 10 seconds before sending the next request
        time.sleep(10)

# Start the thread
thread = threading.Thread(target=send_request)
thread.daemon = True  # Set the thread as daemon so it exits when the main program exits
thread.start()

# Main program continues...
# You can put your main program logic here
# It will run concurrently with the thread making GET requestsf
main(base_url, username)
