

---
ms.Toctitle:RemoteItem.MarkForDownload プロパティ (Outlook)(機械翻訳)
title:RemoteItem.MarkForDownload プロパティ (Outlook)(機械翻訳)
ms.ContentId:1edfec8a-511f-6e0a-df6c-f6602c1d3d0a
---
# RemoteItem.MarkForDownload プロパティ (Outlook)(機械翻訳)




リモート ユーザーがそれを受信した後、アイテムのステータスを決定する**OlRemoteStatus**の定数を設定または返します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**MarkForDownload**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **RemoteItem** オブジェクトを表す変数を指定します。



## 注釈
このプロパティによって、メッセージングの柔軟性が向上したデータ転送機能がリモート ユーザーに提供されます。



## 例
次の例は、ユーザーの**受信トレイ**内で、完全にダウンロードされていないアイテムを検索します。該当するアイテムが見つかった場合はユーザーにメッセージを表示し、アイテムにダウンロードのマークを付けます。

```vba
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To mpfInbox.Items.Count 
 
 Set obj = mpfInbox.Items.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox ("This item has not been fully downloaded.") 
 
 'Mark the item to be downloaded. 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 End If 
 
 Next 
 
End Sub
```




## Related Topics

[RemoteItem オブジェクトのメンバー](15c0872e-88cc-9b9b-c31e-c15d6971e6e0.md)

[RemoteItem オブジェクト](6302aaff-cdcf-4d86-60f1-4bed15540d9f.md)




