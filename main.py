# coding: utf-8
import ctypes
import requests
import webbrowser
from wox import Wox

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


class YoudaoDict(Wox):

    def query(self, key):
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
        if isinstance(data, str):
            OpenClipboard(None)
            EmptyClipboard()
            hCd = GlobalAlloc(GMEM_DDESHARE, 2 * (len(data) + 1))
            pchData = GlobalLock(hCd)
            wcscpy(ctypes.c_wchar_p(pchData), data)
            GlobalUnlock(hCd)
            SetClipboardData(CF_UNICODETEXT, hCd)
            CloseClipboard()


if __name__ == "__main__":
    YoudaoDict()
