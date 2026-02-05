---
title: APIを使わずにPythonでGmailを作成
tags:
  - Python
private: false
updated_at: '2020-10-17T15:47:18+09:00'
id: 83e511f09237ea6a3397
organization_url_name: null
slide: false
ignorePublish: false
---
# 概要
* [APIを使わずにVBAでGmailを作成](https://qiita.com/neruru_y/items/a5d0a3f7ef30f5a36962) のPython版になります。
* Excelでメール本文書くなんてやりづらいと思う人(自分もそう思いました)の為の別の提案

# PythonでGmailのメール作成画面表示
今回のやり方としては、テキストファイルに本文を書き、
そのテキストファイルを取得して、URLを作成する、というやり方になります。
最近日報ではもうこのやり方でやってます。
本記事では、日報を想定とした書き方の上で、前回の流れ通り書いていきます。

## URL作成
URLやパラメータは前回の記事にも書いてありますが、以下の通りです。
```https://mail.google.com/mail/?view=cm``` このURLに、パラメータを連結させます。
パラメータは以下の通り(ほぼまんまです)  

| パラメータ | 意味 |
| :----: | :----: |
| to= | To |
| cc= | Cc |
| bcc= | Bcc |
| su= | 件名 |
| body= | 本文 |

```python:URL作成
from datetime import datetime
import urllib.parse

def getUrl(body: str) -> str:
    url = "https://mail.google.com/mail/?view=cm"
    url += "&to=to@hoge.co.jp"
    url += "&cc=cc@hoge.co.jp"
    url += "&bcc=bcc@hoge.co.jp"
    today = datetime.now()
    url +=  f"&su=日報 {today.month}/{today.strftime('%d')} 日報太郎"
    url += f"&body={strenc(body)}"

    return url

def strenc(txt: str) -> str :
    lst = list("#'|^<>{};?@&$" + '"')
    for v in lst:
        txt = txt.replace(v, urllib.parse.quote(v))
    txt = urllib.parse.quote(txt)
    return txt
```

ここらへんは前回とほぼ同じです。

## URLを開く
こちらも前回と同じで、コマンドを実行していきます。
Pythonなので、VBAと違って書きやすいです。

```python:PythonからstartコマンドでURLを開く
import subprocess

def openUrl(url: str):
    subprocess.call(f'cmd /c start "" "{url}"', shell=True)
```

## メール作成
関数ができたので、テキストファイルの内容を取得して、メールを作成していきます。
デスクトップに、```メール.txt```というファイルがあったとします。

```python
import os

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop/')
fileName = "メール.txt"

def main():
    with open(f"{desktop}{fileName}", mode="r", encoding="UTF-8") as f:
        url = getUrl(f.read())
        openUrl(url)
```

### バッチ作成
これで実行すれば、ブラウザが開き、Gmailが開かれるのですが、
いちいちコマンドプロンプトで、Pythonを実行っていうのはめんどくさいので、バッチを書いていきます。
Pythonファイルは、デスクトップに置いておく必要はないので、
例として、ユーザーフォルダの中に、```/work/python/DailyReport/DailyReport.py``` としておきます。

```bat
@echo off
cd ../work/python/DailyReport
python DailyReport.py
```

このバッチを、デスクトップに保存します。
これにより、```メール.txt```に本文を書いて、バッチを実行すれば、
ブラウザで開かれ、メール作成画面が表示されたかと思います。
(今回実行結果は割愛します。内容一緒なので・・・)
outlookに関しても、前回の[おまけ](https://qiita.com/neruru_y/items/a5d0a3f7ef30f5a36962#%E3%81%8A%E3%81%BE%E3%81%91)にある、```getUrl()```を真似すればできるかと思います。

## 全体のソースコード
```python:DailyReport.py
from datetime import datetime
import os
import urllib.parse
import subprocess

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop/')
fileName = "メール.txt"

def main():
    with open(f"{desktop}{fileName}", mode="r", encoding="UTF-8") as f:
        url = getUrl(f.read())
        openUrl(url)

def getUrl(body: str) -> str:
    url = "https://mail.google.com/mail/?view=cm"
    url += "&to=to@hoge.co.jp"
    url += "&cc=cc@hoge.co.jp"
    url += "&bcc=bcc@hoge.co.jp"
    today = datetime.now()
    url +=  f"&su=日報 {today.month}/{today.strftime('%d')} 日報太郎"
    url += f"&body={strenc(body)}"

    return url

def strenc(txt: str) -> str :
    lst = list("#'|^<>{};?@&$" + '"')
    for v in lst:
        txt = txt.replace(v, urllib.parse.quote(v))
    txt = urllib.parse.quote(txt)
    return txt

def openUrl(url: str):
    subprocess.call(f'cmd /c start "" "{url}"', shell=True)

if __name__ == "__main__":
    main()

```

```bat:mail.bat
@echo off
cd ../work/python/DailyReport
python DailyReport.py
```

# 参考サイト
* [APIを使わずにVBAでGmailを作成](https://qiita.com/neruru_y/items/a5d0a3f7ef30f5a36962)
