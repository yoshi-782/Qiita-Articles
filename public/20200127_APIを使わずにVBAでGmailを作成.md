---
title: APIを使わずにVBAでGmailを作成
tags:
  - Excel
  - VBA
private: false
updated_at: '2020-01-27T14:01:03+09:00'
id: a5d0a3f7ef30f5a36962
organization_url_name: null
slide: false
ignorePublish: false
---
# 概要
VBAでGmailのメールを作成する方法を記載していきますが、このやり方は、以下に該当する方におすすめです。

* ちょっとしたメールを送信するのにわざわざAPIを使えるようにするのはめんどくさい。
* 自分で送信ボタンを押して確実に送信されたことを確認したい。

いや、自分で送信ボタンを押す自体めんどくさい、マクロ実行したらもう勝手に送ってくれという方は、このやり方はおすすめしません。

またVBAでメールを作成と書きましたが、実際に作るのはURLです。
VBAでURLを作成し、
そのURLをブラウザで表示させるという仕組みになります。
# VBAでGmailのメール作成画面表示
宛先、件名、本文を元にメールを作成していきます。
手順としては、URL作成→URLを開く→送信ボタンをユーザがクリックする、という流れです。
## URL作成
```https://mail.google.com/mail/?view=cm``` このURLに、パラメータを連結させます。
パラメータは以下の通り(ほぼまんまです)

| パラメータ | 意味 |  
| :----: | :----: |
| &to= | To |
| &cc= | Cc |
| &bcc= | Bcc |
| &su= | 件名 |
| &body= | 本文 |
本文、件名、To、Cc、Bccを渡して、URLを返すプロシージャを作成します。

```vb:URL作成
Private Function getUrl(body As String, Optional subj As String = "", Optional addr As String = "", Optional cc As String = "", Optional bcc As String = "") As String
    Dim url As String: url = "https://mail.google.com/mail/?view=cm"
    Dim prams(4) As String
    prams(0) = IIf(Len(addr) > 0, "&to=" & addr, "")
    prams(1) = IIf(Len(cc) > 0, "&cc=" & cc, "")
    prams(2) = IIf(Len(bcc) > 0, "&bcc=" & bcc, "")
    prams(3) = IIf(Len(subj) > 0, "&su=" & subj, "")
    prams(4) = "&body=" & encodeText(body)
    getUrl = url & Join(prams, "")
End Function

'文字列エンコード
Private Function encodeText(text As String) As String
    Dim enc As Variant: enc = Split("[-]-%-\-#-'-|-`-^-""-<->-{-}-;-?-:-@-&-=-+-$-,", "-")
    Dim e As Variant
    For Each e In enc
        text = Replace(text, e, Application.WorksheetFunction.EncodeURL(e))
    Next e
    encodeText = Replace(text, vbLf, "%0D%0A")
End Function
```

文字列エンコードは、URLで使用できないもののみをエンコードし、
最後に```vbLf```を```%0D%0A```に置換しています。
件名は、なぜかエンコードした文字がそのまま表示されるのでエンコードしてません。
(エンコードしないとエラーになる文字もあるかも)

## URLを開く
VBAからコマンドラインを実行し、URLを開きます。

``` vb:VBAからstartコマンドでURLを開く
Private Sub openUrl(url As String)
    CreateObject("WScript.shell").Run "cmd /c start" & " " & String(2, Chr(34)) & " " & Chr(34) & url & Chr(34), 0, True
End Sub
```

ちなみにコマンドは結果的に以下になります。

```text:コマンド

cmd /c start "" https://mail.google.com/mail/?view=cm...
```

## メール作成
関数ができたので、早速実行します。
![メール](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/176021/469e90d5-d7f2-aa41-9091-78947a451b45.png)

```vb:main
Public Sub main()
    Dim addr As String: addr = Range("C2").Value
    Dim cc As String: cc = Range("C3").Value
    Dim subj As String: subj = Range("C5").Value
    Dim body As String: body = Range("C6").Value
    
    Dim url As String: url = getUrl(body, subj, addr, cc)
    Call openUrl(url)
End Sub

```

実行結果
![実行結果.gif](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/176021/60ce8505-6260-4178-11fb-45cece6b1517.gif)

# おまけ
このURLを作成してメール送信画面を表示させる方法は、Web版のOutlookでもできます。  
ただ、Outlookでは、件名と本文をURLに連結させることはできましたが、宛先はToのみしか連結できませんでした。
CC、BCCは対応していないみたいです。
```URL:https://outlook.office365.com/mail/deeplink/compose?```

| パラメータ | 意味 |  
| :----: | :----: |
| subject= | 件名 |
| body= | 本文 |
| to= | To |

```vb:outlook版getUrl関数
Private Function getUrl(body As String, Optional subj As String, Optional addr As String) As String
    Dim url As String: url = "https://outlook.office365.com/mail/deeplink/compose?"
    Dim prams(2) As String
    prams(0) = "&body=" & encodeText(body)
    prams(1) = IIf(Len(subj) > 0, "&subject=" & subj, "")
    prams(2) = IIf(Len(addr) > 0, "&to=" & addr, "")
    getUrl = url & Join(prams, "")
end Function
```

# 参考サイト
* [【URLリンク】のクリックでGmail作成画面を出す方法（宛先、件名の自動挿入）](https://kitaney-google.blogspot.com/2013/12/urlgmail.html)
* [URLで使用可能な文字、使用できない文字](https://www.ipentec.com/document/web-url-invalid-char)
* [Outlook on the Web の新規メール作成画面を開くハイパーリンク](https://idea.tostring.jp/?p=2826)
