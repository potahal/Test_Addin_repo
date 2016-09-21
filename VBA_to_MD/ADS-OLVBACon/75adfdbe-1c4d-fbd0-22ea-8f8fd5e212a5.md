

---
ms.Toctitle:ソリューションのストレージ アイテムにデータを格納します。
title:ソリューションのストレージ アイテムにデータを格納します。
ms.ContentId:75adfdbe-1c4d-fbd0-22ea-8f8fd5e212a5
---
# ソリューションのストレージ アイテムにデータを格納します。




このトピックでは、Outlook オブジェクト モデルに備わったソリューション ストレージに対して、プライベートなアプリケーション データを格納する方法について説明します。

1. アプリケーション データを格納するフォルダーを決めます。

>[!NOTE]
>ソリューション ストレージはフォルダー内の非表示のアイテムとして作成されるため、ソリューション データを格納できるのは、ストア プロバイダーが非表示のアイテムをサポートし、クライアントがそのフォルダーへの書き込みの権限を持つ場合に限られます。


2. **Folder.GetStorage**を使用して、存在しない場合、既存の**StorageItem**オブジェクトまたは新しい**StorageItem**オブジェクトのいずれかを取得します。
3. **StorageItem.Size**を使用して、 **StorageItem**が新しいかどうか判断します。場合は、 **StorageItem.UserProperties**の**Add**メソッドを使用してカスタム プロパティ**の順序番号**を作成します。
4. **注文番号**のプロパティを設定します。これは、既存の**StorageItem**が既にカスタム プロパティ定義されている**注文番号**を持っていると仮定します。
5. **StorageItem.Save**を使用して、フォルダー内の非表示のアイテムとしての**StorageItem**オブジェクトを保存します。


```sourcecode
Sub StoreData() 
 Dim oInbox As Folder 
 Dim myStorage As StorageItem 
 Dim myPrivateProperty As UserProperty 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 ' Get an existing instance of StorageItem by subject, or create new if it doesn't exist 
 Set myStorage = oInbox.GetStorage("My Private Storage", olIdentifyBySubject) 
 
 If myStorage.Size = 0 Then 
 'There was no existing StorageItem by this subject, so created a new one 
 'Create a custom property for Order Number 
 Set myPrivateProperty = myStorage.UserProperties.Add("Order Number", olNumber) 
 Else 
 'Assume that existing storage has the Order Number property already 
 Set myPrivateProperty = myStorage.UserProperties("Order Number") 
 End If 
 myPrivateProperty.Value = lngOrderNumber 
 myStorage.Save 
End Sub
```



