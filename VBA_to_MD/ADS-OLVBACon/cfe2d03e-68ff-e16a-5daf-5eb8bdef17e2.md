

---
ms.Toctitle:フォルダー内の既存のアイテムが使用するフォームの変更
title:フォルダー内の既存のアイテムが使用するフォームの変更
ms.ContentId:cfe2d03e-68ff-e16a-5daf-5eb8bdef17e2
---
# フォルダー内の既存のアイテムが使用するフォームの変更




フォルダー内の既存のアイテムに関連付けられているフォームの変更が必要な場合もあります。通常、このフォームの変更は、アイテムをインポートした場合、または標準の Outlook フォームに基づいて既にアイテムを作成した後にユーザー設定フォームを作成する場合に必要になります。



"メッセージ クラス" フィールドは、Outlook のユーザー インターフェイスを使って直接変更することはできませんが、VBScript、Visual Basic、または VBA を使用すると変更できます。



次のオートメーション コードは、独自のソリューションを開発する場合に基本として使用できます。このコードは、新しいフォームの名前が MyForm であることを前提としています。このコードを実行すると、既定の連絡先フォルダー内のすべての連絡先が MyForm を使用するように変更されます。

```sourcecode
Sub ChangeMessageClass() 
Set olNS = Application.GetNameSpace("MAPI") 
Set ContactsFolder = _ 
 olNS.GetDefaultFolder(olFolderContacts) 
Set ContactItems = ContactsFolder.Items 
 
For Each Itm in ContactItems 
 If Itm.MessageClass <> "IPM.Contact.MyForm" Then 
 Itm.MessageClass = "IPM.Contact.MyForm" 
 Itm.Save 
 End If 
Next 
End Sub
```


>[!NOTE]
>既定のフォルダー以外のフォルダーを使用する場合は、フォルダー一覧で利用可能な任意のフォルダーを参照する **Folders** コレクション オブジェクトを使用します。




