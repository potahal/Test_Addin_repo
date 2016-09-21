

---
ms.Toctitle:SenderEmailAddress プロパティ
title:SenderEmailAddress プロパティ
ms.ContentId:a157894c-adf2-1cef-ec7c-8516dbef2b7f
---
# SenderEmailAddress プロパティ




Outlook アイテムの送信者の電子メール アドレスを表す文字列型 (**String**) の値を取得します。値の取得のみ可能です。

## 構文

        UNRESOLVED_TOKEN_VAL(offexpression).**SenderEmailAddress**




        UNRESOLVED_TOKEN_VAL(offexpression) **MailItem** オブジェクトを表す変数を指定します。



## 注釈
このプロパティは、MAPI プロパティの **PidTagSenderEmailAddress** に対応しています。



## 例
次に示す Microsoft Visual Basic for Applications (VBA) のコードは、受信トレイにある "テスト" という名前のフォルダー内のすべてのアイテムを反復処理して、"someone@example.com" が送信したアイテムに黄色いフラグを設定します。エラーを発生させずにこのコードを実行するには、既定の受信トレイ フォルダーに "テスト" フォルダーが存在することを確認し、"someone@example.com" を "テスト" フォルダーの実際の送信者の電子メール アドレスに置き換えます。

```vba
Sub SetFlagIcon()
    Dim mpfInbox As Outlook.Folder
    Dim obj As Outlook.MailItem
    Dim i As Integer
    
    Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Folders("Test")
    ' Loop all items in the Inbox\Test Folder
    For i = 1 To mpfInbox.Items.Count
        If mpfInbox.Items(i).Class = olMail Then  
            Set obj = mpfInbox.Items.Item(i)
            If obj.SenderEmailAddress = "someone@example.com" Then
                'Set the yellow flag icon
                obj.FlagIcon = olYellowFlagIcon
                obj.Save
            End If
         End If
    Next
End Sub
```





