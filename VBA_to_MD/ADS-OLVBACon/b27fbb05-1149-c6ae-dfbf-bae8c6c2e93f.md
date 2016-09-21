

---
ms.Toctitle:EntryID および StoreID の使用
title:EntryID および StoreID の使用
ms.ContentId:b27fbb05-1149-c6ae-dfbf-bae8c6c2e93f
---
# EntryID および StoreID の使用




このトピックでは、アイテムのエントリ ID とストアのストア ID を使用して、**NameSpace** 内の特定のアイテムにアクセスする方法について説明します。



リンクまたはフォルダー内の項目を相互参照に関連するUNRESOLVED_TOKEN_VAL(outlooknv1)を使用してより複雑なソリューションを作成する場合、1 つは、各アイテムの MAPI ベースの識別子 (Id) を使用します。アイテムとフォルダーに格納されている Id がわかっている場合は、 **NameSpace.GetItemFromID**メソッドを使用してアイテムを直接参照できます。



Outlook の各アイテムには、**EntryID** と呼ばれるフィールドがあります。これはメッセージング格納システムによって生成される一意の ID フィールドであり、アイテムを格納する MAPI フォルダーで使用されます。フォルダー内にアイテムが作成されるたびに、アイテムに新しい **EntryID** が割り当てられることに注意する必要があります。つまり、アイテムを別のフォルダーに移動したり、エクスポートしてからインポートしたりすると (同じフォルダー内での操作だとしても)、**EntryID** フィールドが変更されます。



各フォルダーには、**Folder.StoreID** と呼ばれる ID フィールドがあります。この ID フィールドの値は、特定のメッセージ ストアのすべてのフォルダーに対して同じです。また、各フォルダーは一意のエントリ ID フィールドも持ちます。



**GetItemFromID** メソッドを使用して ID に基づいてアイテムを取得する場合は、アイテムの **EntryID** とフォルダーの **StoreID** の両方を指定する必要があります。**StoreID** を指定しないと、**GetItemFromID** は既定のメッセージ ストアを検索します。



次の Microsoft Visual Basic for Applications (VBA) の例は、**GetItemFromID** メソッドの使い方を示しています。このコードは、既定の連絡先フォルダーから **StoreID** を取得し、そのフォルダー内のすべての連絡先のエントリ ID を配列 (`MyEntryID`) に代入し、最後に特定の連絡先アイテムを取得します。

```sourcecode
Sub OutlookEntryID() 
 ' If there are more than 500 contacts, change the following line: 
 Dim MyEntryID(500) As String 
 Dim StoreID As String 
 Dim EntryID As String 
 
 Set olns = Application.GetNamespace("MAPI") 
 Set objFolder = olns.GetDefaultFolder(olFolderContacts) 
 ' Get the StoreID, which is a property of the folder. 
 StoreID = objFolder.StoreID 
 ' Set objAllContacts equal to the collection of all contacts. 
 Set AllContacts = objFolder.Items 
 I = 0 
 ' Loop to get all of the EntryIDs for the contacts. 
 For Each Item In AllContacts 
 I = I + 1 
 MyEntryID(I) = Item.EntryID 
 Next 
 ' Randomly choose the 2nd contact to retrieve. 
 Set Item = olns.GetItemFromID(MyEntryID(2), StoreID) 
 Item.Display 
End Sub
```



