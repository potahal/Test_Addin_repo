

---
ms.Toctitle:TaskRequestItem.DownloadState プロパティ (Outlook)(機械翻訳)
title:TaskRequestItem.DownloadState プロパティ (Outlook)(機械翻訳)
ms.ContentId:3ddd05d1-fa91-6d5e-cb02-5a9df90ad2af
---
# TaskRequestItem.DownloadState プロパティ (Outlook)(機械翻訳)




アイテムのダウンロード状況を示す **OlDownloadState** 列挙に属している定数を取得します。値の取得のみ可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**DownloadState**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskRequestItem** オブジェクトを表す変数を指定します。



## 例
次の Microsoft Visual Basic for Applications (VBA) の例は、ユーザーの**受信トレイ**内でまだ完全にダウンロードされていないアイテムを検索します。該当するアイテムが見つかった場合は、ユーザーにメッセージを表示し、アイテムにダウンロードのマークを付けます。

```vba
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim objItems As Outlook.Items 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 Dim iCount As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 Set objItems = mpfInbox.Items 
 
 iCount = objItems.Count 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To iCount 
 
 Set obj = objItems.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox "This item has not been fully downloaded." 
 
 'Mark the item to be downloaded 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 obj.Save 
 
 End If 
 
 Next 
 
End Sub
```




## Related Topics

[オブジェクトのメンバー](d43114ee-be91-ff02-3424-525da2cf3a50.md)

[オブジェクト](2908a28a-634c-e786-aa53-f3e32038b727.md)




