

---
ms.Toctitle:Conversation.GetAlwaysMoveToFolder メソッド (Outlook)(機械翻訳)
title:Conversation.GetAlwaysMoveToFolder メソッド (Outlook)(機械翻訳)
ms.ContentId:ecad049d-338b-d5e0-f241-a9dddaeae316
---
# Conversation.GetAlwaysMoveToFolder メソッド (Outlook)(機械翻訳)




スレッドで受信された新規アイテムを常に移動する先として指定された配信ストア内のフォルダーを示す、**Folder** オブジェクトを返します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetAlwaysMoveToFolder**(**Store**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Conversation** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Store*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Store**|スレッド アイテムの移動先フォルダーが存在するストアです。|



### 戻り値
移動先となる会話で受信されたすべての新しいアイテムは、常に指定されたストア内の**Folder**オブジェクト。





## 注釈
*格納*パラメーターは、アーカイブ .pst ストアなど、配信ストアを表している場合、 **GetAlwaysMoveToFolder**メソッドは、既定の配信ストアに会話の項目に適用する**フォルダー**オブジェクトを返します。



常に会話の項目を移動する、**削除済みアイテム**フォルダー以外のフォルダーが指定されていない場合、 **GetAlwaysMoveToFolder**メソッドは**Null** (**Nothing**で Visual Basic) を返します。



## 例
アプリケーション (VBA) の例を次の Microsoft Visual Basic では、移動先となる最初のメール アイテムを閲覧ウィンドウに表示されるの会話に到着した新しいアイテムは、常にフォルダーを検索する方法を示します。コード例では、 `DemoGetAlwaysMoveToFolder`では、会話が選択されているメール アイテムのストアで有効になって、会話が存在するフォルダーを取得する**GetAlwaysMoveToFolder**を使用して、フォルダー名が表示される場合は、メール アイテムの会話のオブジェクトを取得するを確認します。

```vba
Sub DemoGetAlwaysMoveToFolder() 
 
 Dim oMail As Outlook.MailItem 
 
 Dim oConv As Outlook.Conversation 
 
 Dim oStore As Outlook.Store 
 
 
 
 ' Get Item displayed in Reading Pane. 
 
 Set oMail = ActiveExplorer.Selection(1) 
 
 Set oStore = oMail.Parent.Store 
 
 If oStore.IsConversationEnabled Then 
 
 Set oConv = oMail.GetConversation 
 
 If Not (oConv Is Nothing) Then 
 
 Dim oFolder As Outlook.folder 
 
 Set oFolder = _ 
 
 oConv.GetAlwaysMoveToFolder(oStore) 
 
 If Not (oFolder Is Nothing) Then 
 
 Debug.Print "MoveToFolder: " & oFolder.name 
 
 Else 
 
 Debug.Print "MoveToFolder action not set" 
 
 End If 
 
 End If 
 
 End If 
 
End Sub
```




## Related Topics

[オブジェクトのメンバーを会話](09ff1e8e-7c5a-0b1e-e8e2-e259f66f71c8.md)

[会話オブジェクト](2705d38a-ebc0-e5a7-208b-ffe1f5446b1b.md)




