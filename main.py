# coding : utf-8
import os
import ctypes
import requests
import webbrowser
from wox import Wox
import json
from subprocess import Popen, PIPE

URL = "http://fanyi.youdao.com/openapi.do?keyfrom=longcwang&key=131895274&type=data&doctype=json&version=1.1&q="
VOICE_PALY = "D:\Program Files (x86)\Clementine\clementine.exe"
voice_url = 'http://dict.youdao.com/dictvoice?type=2&audio={word}'
VOICE_DIR = os.path.join(os.getcwd(), 'voice')
recode_txt = os.path.join(os.getcwd(), "recode.txt")
recode_json = os.path.join(os.getcwd(), "recode.json")
wcscpy = ctypes.cdll.msvcrt.wcscpy
OpenClipboard = ctypes.windll.user32.OpenClipboard
EmptyClipboard = ctypes.windll.user32.EmptyClipboard
GetClipboardData = ctypes.windll.user32.GetClipboardData
SetClipboardData = ctypes.windll.user32.SetClipboardData
CloseClipboard = ctypes.windll.user32.CloseClipboard
CF_UNICODETEXT = 13
GlobalAlloc = ctypes.windll.kernel32.GlobalAlloc
GlobalLock = ctypes.windll.kernel32.GlobalLock
GlobalUnlock = ctypes.windll.kernel32.GlobalUnlock
GMEM_DDESHARE = 0x2000

if not os.path.isdir(VOICE_DIR):
    os.mkdir(VOICE_DIR)

word_dict = dict()


class YoudaoDict(Wox):

    def query(self, key):
        if "-" in key:
            word_dict.setdefault('isplay', 1)
            key = " ".join(key.split()[1:])
            word_dict.setdefault('word', key)
        else:
            word_dict.setdefault('word', key)
        with open(recode_json, 'w') as f:
            json.dump(word_dict, f)
        results = []
        URL = "http://fanyi.youdao.com/openapi.do"
        params = {
                    "keyfrom": "f2ec-org",
                    "key": "1787962561",
                    "doctype": "json",
                    "version": "1.1",
                    "type": "data",
                    "q": "hello",
                }
        params['q'] = key
        r = requests.get(URL, params=params)
        a = 1
        results.append({
                "Title": ' '.join(r.json()['translation']),
                "SubTitle": "简明释义",
                "IcoPath": "Images/zxc.ico",
                "JsonRPCAction": {
                    "method": "put",
                    "parameters": [' '.join(r.json()['translation'])],
                    "dontHideAfterAction": False
                }
        })
        try:
            try:
                for item in r.json()['basic']['explains']:
                    results.append({
                        "Title": item.split('.')[1],
                        "SubTitle": item.split('.')[0]+". 基本解释",
                        "IcoPath": "Images/zxc.ico",
                        "JsonRPCAction": {
                            "method": "put",
                            "parameters": [item.split('.')[1]],
                            "dontHideAfterAction": False
                        }
                    })
            except IndexError:
                results.append({
                    "Title": " ;".join(r.json()['basic']['explains']),
                    "SubTitle": "基本解释",
                    "IcoPath": "Images/zxc.ico",
                    "JsonRPCAction": {
                        "method": "put",
                        "parameters": [" ;".join(r.json()['basic']['explains'])],
                        "dontHideAfterAction": False
                    }
                })
        except KeyError:
            pass
        try:
            for item in r.json()['web']:
                results.append(
                    {
                        "Title": ' ;'.join(item['value']),
                        "SubTitle": item['key'],
                        "IcoPath": "Images/zxc.ico",
                        "JsonRPCAction": {
                            "method": "put",
                            "parameters": [' ;'.join(item['value'])],
                            "dontHideAfterAction": False
                        }
                    }
                )
        except KeyError:
            pass
        return results

    def detail(self, url):
        webbrowser.open(url)

    def put(self, data):
        to_dict = None
        with open(recode_json, "r") as f:
            to_dict = json.load(f)
        if to_dict.get("isplay", 0):
            word = to_dict.get("word")
            if ord(word[0]) < 127:
                self.play(word)
            else:
                self.play(data)

        if isinstance(data, str):
            OpenClipboard(None)
            EmptyClipboard()
            hCd = GlobalAlloc(GMEM_DDESHARE, 2 * (len(data) + 1))
            pchData = GlobalLock(hCd)
            wcscpy(ctypes.c_wchar_p(pchData), data)
            GlobalUnlock(hCd)
            SetClipboardData(CF_UNICODETEXT, hCd)
            CloseClipboard()

    def get_voice(self, word):
        voice_file = os.path.join(VOICE_DIR, word+'.mp3')
        if not os.path.isfile(voice_file):
            r = requests.get(voice_url.format(word=word))
            with open(voice_file, 'wb') as f:
                f.write(r.content)
        return voice_file

    def play(self, word):
        voice_file = self.get_voice(word)
        p = Popen([VOICE_PALY, voice_file], stderr=PIPE, stdin=PIPE, stdout=PIPE)
        p.communicate()


if __name__ == "__main__":
    YoudaoDict()
