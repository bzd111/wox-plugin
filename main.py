# coding : utf-8
import os
import json
from subprocess import call, PIPE

import requests

from wox import Wox  # noqa

URL = 'http://fanyi.youdao.com/openapi.do'
VOICE_PALY = 'D:\Program Files (x86)\Clementine\clementine.exe'
VOICE_URL = 'http://dict.youdao.com/dictvoice?type=2&audio={word}'
VOICE_DIR = os.path.join(os.getcwd(), 'voice')
RECODE_TXT = os.path.join(os.getcwd(), 'recode.txt')
RECODE_JSON = os.path.join(os.getcwd(), 'recode.json')


class YoudaoDict(Wox):
    word_dict = {}

    def __init__(self):
        super().__init__()
        self.init_env()

    def query(self, key):
        results = []
        if '-' in key:
            self.word_dict.setdefault('isplay', 1)
            key = ' '.join(key.split()[1:])
            self.word_dict.setdefault('word', key)
        else:
            self.word_dict.setdefault('word', key)
        with open(RECODE_JSON, 'w') as f:
            json.dump(self.word_dict, f)
        params = {
            'keyfrom': 'f2ec-org',
            'key': '1787962561',
            'doctype': 'json',
            'version': '1.1',
            'type': 'data',
            'q': 'hello',
        }
        params['q'] = key
        r = requests.get(URL, params=params)
        results.append({
            'Title': ' '.join(r.json()['translation']),
            'SubTitle': '简明释义',
            'IcoPath': 'Images/zxc.ico',
            'JsonRPCAction': {
                'method': 'copy',
                'parameters': [' '.join(r.json()['translation'])],
                'dontHideAfterAction': False
            }
        })
        try:
            try:
                for item in r.json()['basic']['explains']:
                    results.append({
                        'Title': item.split('.')[1],
                        'SubTitle': item.split('.')[0] + '. 基本解释',
                        'IcoPath': 'Images/zxc.ico',
                        'JsonRPCAction': {
                            'method': 'copy',
                            'parameters': [item.split('.')[1]],
                            'dontHideAfterAction': False
                        }
                    })
            except IndexError:
                results.append({
                    'Title': ' ;'.join(r.json()['basic']['explains']),
                    'SubTitle': '基本解释',
                    'IcoPath': 'Images/zxc.ico',
                    'JsonRPCAction': {
                        'method': 'copy',
                        'parameters': [' ;'.join(r.json()['basic']['explains'])],
                        'dontHideAfterAction': False
                    }
                })
        except KeyError:
            pass
        try:
            for item in r.json()['web']:
                results.append(
                    {
                        'Title': ' ;'.join(item['value']),
                        'SubTitle': item['key'],
                        'IcoPath': 'Images/zxc.ico',
                        'JsonRPCAction': {
                            'method': 'copy',
                            'parameters': [' ;'.join(item['value'])],
                            'dontHideAfterAction': False
                        }
                    }
                )
        except KeyError:
            pass
        return results

    def copy(self, data):
        call('echo ' + data + '| clip', shell=True)

    def get_voice(self, word):
        voice_file = os.path.join(VOICE_DIR, word + '.mp3')
        if not os.path.isfile(voice_file):
            r = requests.get(VOICE_URL.format(word=word))
            with open(voice_file, 'wb') as f:
                f.write(r.content)
        return voice_file

    def play(self, word):
        voice_file = self.get_voice(word)
        call([VOICE_PALY, voice_file], shell=False, stderr=PIPE, stdin=PIPE, stdout=PIPE)

    def init_env(self):
        if not os.path.isdir(VOICE_DIR):
            os.mkdir(VOICE_DIR)


if __name__ == '__main__':
    YoudaoDict()
