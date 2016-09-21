

---
ms.Toctitle:FormDescription.ScriptText プロパティ (Outlook)(機械翻訳)
title:FormDescription.ScriptText プロパティ (Outlook)(機械翻訳)
ms.ContentId:56ea4cd6-a9f0-cd0c-a378-dab6399bd1ca
---
# FormDescription.ScriptText プロパティ (Outlook)(機械翻訳)




フォームのスクリプト エディターのすべての VBScript コードを含みます。文字列型 (**String**) の値を返します。値の取得のみ可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ScriptText**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **FormDescription** オブジェクトを表す変数を指定します。



## 例
次の Microsoft Visual Basic Scripting Edition (VBScript) の例は、**Open** 
 イベントを使用して、**MailItem** の **HTMLBody** プロパティにアクセスします。これにより、**MailItem** の **Inspector** の **EditorType** プロパティが **olEditorHTML** に設定されます。**MailItem** の **Body** プロパティを設定すると、**EditorType** プロパティが既定値に変更されます。たとえば、既定の電子メール エディターが RTF に設定されている場合、**EditorType** は **olEditorRTF** に設定されます。デザイン モードでフォームの Script Editor に次のコードを入力した場合、実行時にフォームの本文を変更すると、**EditorType** プロパティの変化がメッセージ ボックスに表示されます。最後のメッセージ ボックスでは、Script Editor にすべての VBScript コードを表示するために **Script Text** プロパティを使用しています。

```sourcecode
Function Item_Open() 
 
 'Set the HTMLBody of the item. 
 
 Item.HTMLBody = "<HTML><H2>My HTML page.</H2><BODY>My body.</BODY></HTML>" 
 
 'Item displays HTML message. 
 
 Item.Display 
 
 'MsgBox shows EditorType is 2. 
 
 MsgBox "HTMLBody EditorType is " & Item.GetInspector.EditorType 
 
 'Access the Body and show 
 
 'the text of the Body. 
 
 MsgBox "This is the Body: " & Item.Body 
 
 'After accessing, EditorType 
 
 'is still 2. 
 
 MsgBox "After accessing, the EditorType is " & Item.GetInspector.EditorType 
 
 'Set the item's Body property. 
 
 Item.Body = "Back to default body." 
 
 'After setting, EditorType is 
 
 'now back to the default. 
 
 MsgBox "After setting, the EditorType is " & Item.GetInspector.EditorType 
 
 'Access the items's 
 
 'FormDescription object. 
 
 Set myForm = Item.FormDescription 
 
 'Display all the code 
 
 'in the Script Editor. 
 
 MsgBox myForm.ScriptText 
 
End Function
```




## Related Topics

[FormDescription Object](c88f92c4-4cac-84b3-6118-1150d42d7cff.md)

[FormDescription Object Members](664724e9-e74b-32ad-93e4-8d4cb27b3082.md)




