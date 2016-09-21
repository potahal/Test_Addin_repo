

---
ms.Toctitle:AppointmentItem.Copy メソッド (Outlook)(機械翻訳)
title:AppointmentItem.Copy メソッド (Outlook)(機械翻訳)
ms.ContentId:947f1cfd-f60c-a47e-ba4d-3ffde8c13c91
---
# AppointmentItem.Copy メソッド (Outlook)(機械翻訳)




オブジェクトの別のインスタンスを作成します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Copy**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **AppointmentItem** オブジェクトを表す変数を指定します。



## 例
例には、この Visual Basic for Applications 電子メール メッセージを作成、**件名**に「スピーチ」を設定をコピーするには、 **Copy**メソッドを使用し、受信トレイ フォルダー内の [メールの保存] をという名前の新しく作成された電子メール フォルダーにコピーを移動します。

```vba
Sub CopyItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myNewFolder As Outlook.Folder 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myCopiedItem As Outlook.MailItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set myNewFolder = myFolder.Folders.Add("Saved Mail", olFolderDrafts) 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 myItem.Subject = "Speeches" 
 
 Set myCopiedItem = myItem.Copy 
 
 myCopiedItem.Move myNewFolder 
 
End Sub
```




## Related Topics

[AppointmentItem オブジェクトのメンバー](c72c459d-6d3c-7a05-aa4a-b1b767ddc0b2.md)

[AppointmentItem オブジェクト](204a409d-654e-27aa-643a-8344c631b82d.md)




