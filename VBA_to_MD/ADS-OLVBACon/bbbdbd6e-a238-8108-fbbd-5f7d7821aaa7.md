

---
ms.Toctitle:Application.Explorers プロパティ (Outlook)(機械翻訳)
title:Application.Explorers プロパティ (Outlook)(機械翻訳)
ms.ContentId:bbbdbd6e-a238-8108-fbbd-5f7d7821aaa7
---
# Application.Explorers プロパティ (Outlook)(機械翻訳)




開いているすべてのエクスプ ローラーを表す**Explorer**オブジェクトを格納している**エクスプ ローラー**コレクション オブジェクトを返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Explorers**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数を指定します。



## 例
次の Microsoft Visual Basic for Applications (VBA) の例は、開いているエクスプローラー ウィンドウの数を表示します。

```vba
Private Sub CountExplorers() 
 
 MsgBox "There are " & _ 
 
 Application.Explorers.Count & " Explorers." 
 
End Sub
```




次の VBA の例は、**受信トレイ**を表示するエクスプ ローラーで選択されているすべてのメール アイテムの送信者を表示するのには、 **Count**プロパティと**Selection**プロパティによって返される**選択範囲**のコレクションの**Item**メソッドを使用します。次の使用例を実行するが少なくとも 1 つのメール アイテムを受信トレイを表示するエクスプ ローラーで選択する必要があります。******依頼**が存在しないために、仕事の依頼など、メール以外のアイテムを選択する場合は、エラーが表示される場合があります。

```vba
Sub GetSelectedItems() 
 
 Dim myOlExp As Outlook.Explorer 
 
 Dim myOlSel As Outlook.Selection 
 
 Dim MsgTxt As String 
 
 Dim x As Integer 
 
 
 
 MsgTxt = "You have selected items from: " 
 
 Set myOlExp = Application.Explorers.Item(1) 
 
 If myOlExp = "Inbox" Then 
 
 Set myOlSel = myOlExp.Selection 
 
 For x = 1 To myOlSel.Count 
 
 MsgTxt = MsgTxt & myOlSel.Item(x).SenderName & ";" 
 
 Next x 
 
 MsgBox MsgTxt 
 
End If 
 
End Sub
```




## Related Topics

[Application オブジェクト メンバー](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)

[Application オブジェクト](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)




