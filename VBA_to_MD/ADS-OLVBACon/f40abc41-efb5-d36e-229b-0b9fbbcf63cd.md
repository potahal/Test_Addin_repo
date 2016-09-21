

---
ms.Toctitle:AppointmentItem.CustomPropertyChange イベント (Outlook)(機械翻訳)
title:AppointmentItem.CustomPropertyChange イベント (Outlook)(機械翻訳)
ms.ContentId:f40abc41-efb5-d36e-229b-0b9fbbcf63cd
---
# AppointmentItem.CustomPropertyChange イベント (Outlook)(機械翻訳)




アイテム (親オブジェクトのインスタンス) のカスタム プロパティが変更されると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CustomPropertyChange**(**Name**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **AppointmentItem** オブジェクトを表す変数を指定します。

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

[AppointmentItem オブジェクトのメンバー](c72c459d-6d3c-7a05-aa4a-b1b767ddc0b2.md)

[AppointmentItem オブジェクト](204a409d-654e-27aa-643a-8344c631b82d.md)




