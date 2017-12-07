# coding : utf-8
import os
import ctypes
import requests
from wox import Wox
import json
from subprocess import call, PIPE


URL = "http://fanyi.youdao.com/openapi.do?keyfrom=longcwang&key=131895274&type=data&doctype=json&version=1.1&q="
VOICE_PALY = "D:\Program Files (x86)\Clementine\clementine.exe"
VOICE_URL = 'http://dict.youdao.com/dictvoice?type=2&audio={word}'
VOICE_DIR = os.path.join(os.getcwd(), 'voice')
RECODE_TXT = os.path.join(os.getcwd(), "recode.txt")
RECODE_JSON = os.path.join(os.getcwd(), "recode.json")


class YoudaoDict(Wox):
    def __init__(self):
        super().__init__()
        self.init_env()
        self.word_dict = dict()

    def query(self, key):
        results = []
        if "-" in key:
            self.word_dict.setdefault('isplay', 1)
            key = " ".join(key.split()[1:])
            self.word_dict.setdefault('word', key)
        else:
            self.word_dict.setdefault('word', key)
        with open(RECODE_JSON, 'w') as f:
            json.dump(self.word_dict, f)
        url = "http://fanyi.youdao.com/openapi.do"
        params = {
                    "keyfrom": "f2ec-org",
                    "key": "1787962561",
                    "doctype": "json",
                    "version": "1.1",
                    "type": "data",
                    "q": "hello",
                }
        params['q'] = key
        r = requests.get(url, params=params)
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

    def put(self, data):
        to_dict = None
        with open(RECODE_JSON, "r") as f:
            to_dict = json.load(f)
        if to_dict.get("isplay", 0):
            word = to_dict.get("word")
            if ord(word[0]) < 127:
                self.play(word)
            else:
                self.play(data)

        if isinstance(data, str):
            CF_UNICODETEXT = 13
            GMEM_DDESHARE = 0x2000
            wcscpy = ctypes.cdll.msvcrt.wcscpy
            OpenClipboard = ctypes.windll.user32.OpenClipboard
            EmptyClipboard = ctypes.windll.user32.EmptyClipboard
            SetClipboardData = ctypes.windll.user32.SetClipboardData
            CloseClipboard = ctypes.windll.user32.CloseClipboard
            GlobalAlloc = ctypes.windll.kernel32.GlobalAlloc
            GlobalLock = ctypes.windll.kernel32.GlobalLock
            GlobalUnlock = ctypes.windll.kernel32.GlobalUnlock
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
            r = requests.get(VOICE_URL.format(word=word))
            with open(voice_file, 'wb') as f:
                f.write(r.content)
        return voice_file

    def play(self, word):
        voice_file = self.get_voice(word)
        call([VOICE_PALY, voice_file], shell=False ,stderr=PIPE, stdin=PIPE, stdout=PIPE)

    def init_env(self):
        if not os.path.isdir(VOICE_DIR):
            os.mkdir(VOICE_DIR)


if __name__ == "__main__":
    YoudaoDict()