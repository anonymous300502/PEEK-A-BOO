import pygetwindow as gw
import pyautogui
import time
import requests
import win32com.client
import argparse
import io

def capture_screen():
    screen = pyautogui.screenshot()
    return screen

def get_screen_size():
    return pyautogui.size()

def main(host, key):
    r = requests.post(host+'/new_session', json={'_key': key})
    if r.status_code != 200:    
        print('Server not available.')
        return

    shell = win32com.client.Dispatch('WScript.Shell')
    PREV_IMG = None
    while True:
        desktop = capture_screen()
        width, height = get_screen_size()

        pillow_img = desktop

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
            #print('no desktop change')
            pass

        # events
        try:
            r = requests.post(host+'/events_get', json={'_key': key})
            data = r.json()
            for e in data['events']:
                print(e)

                if e['type'] == 'click':
                    pyautogui.moveTo(e['x'], e['y'], duration=0.2)  # Move to the position first
                    pyautogui.click()  # Then click

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

        time.sleep(0.2)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='pyRD')
    parser.add_argument('addr', help='server address', type=str)
    parser.add_argument('key', help='access key', type=str)
    args = parser.parse_args()

    main(args.addr, args.key)
