

---
ms.Toctitle:TaskRequestItem.CustomPropertyChange イベント (Outlook)(機械翻訳)
title:TaskRequestItem.CustomPropertyChange イベント (Outlook)(機械翻訳)
ms.ContentId:5fcaaba2-706b-76e3-cd6d-be435bca584b
---
# TaskRequestItem.CustomPropertyChange イベント (Outlook)(機械翻訳)




アイテム (親オブジェクトのインスタンス) のカスタム プロパティが変更されると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CustomPropertyChange**(**Name**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskRequestItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Name*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|変更されたカスタム プロパティの名前を指定します。|





## 注釈
プロパティ名がプロシージャに渡されるため、どのプロパティが変更されたかを確認できます。



## 例
この Microsoft Visual Basic Scripting Edition (VBScript) の例では、 **CustomPropertyChange**イベントを使用して、ブール型のフィールドが**True**に設定すると、コントロールを有効にします。



この例では、フォームの 2 ページ目の 2 つのカスタム フィールドを作成します。1 つ目の**ブール値**フィールドの場合は、"RespondBy"の名前です。2 番目のフィールドは"DateToRespond"という名前です。

```sourcecode
Sub Item_CustomPropertyChange(ByVal myPropName) 
 
 Select Case myPropName 
 
 Case "RespondBy" 
 
 Set myPages = Item.GetInspector.ModifiedFormPages 
 
 Set myCtrl = myPages("P.2").Controls("DateToRespond") 
 
 If Item.UserProperties("RespondBy").Value Then 
 
 myCtrl.Enabled = True 
 
 myCtrl.Backcolor = 65535 'Yellow 
 
 Else 
 
 myCtrl.Enabled = False 
 
 myCtrl.Backcolor = 0 'Black 
 
 End If 
 
 Case Else 
 
 End Select 
 
End Sub
```




## Related Topics

[オブジェクトのメンバー](d43114ee-be91-ff02-3424-525da2cf3a50.md)

[オブジェクト](2908a28a-634c-e786-aa53-f3e32038b727.md)




