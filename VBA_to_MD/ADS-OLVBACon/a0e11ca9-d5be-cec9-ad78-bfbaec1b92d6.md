

---
ms.Toctitle:Inspector.SetCurrentFormPage メソッド (Outlook)(機械翻訳)
title:Inspector.SetCurrentFormPage メソッド (Outlook)(機械翻訳)
ms.ContentId:a0e11ca9-d5be-cec9-ad78-bfbaec1b92d6
---
# Inspector.SetCurrentFormPage メソッド (Outlook)(機械翻訳)




インスペクターに指定したフォーム ページまたはフォーム領域を表示します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SetCurrentFormPage**(**PageName**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Inspector** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*PageName*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|フォーム ページの表示名、またはフォーム領域の内部名。|





## 注釈
フォーム領域が個別の置換またはすべて置換フォーム領域である場合、フォーム領域の **InternalName** プロパティを指定して **SetCurrentFormPage** を使用し、フォーム領域を表示します。



## 例
次の Visual Basic for Applications (VBA) の例は、**SetCurrentFormPage** メソッドを使用して、現在開いているアイテムの [**すべてのフィールド**] ページを表示します。エラーが発生した場合は、ユーザーにメッセージを表示します。

```vba
Sub ShowAllFieldsPage() 
 
 On Error GoTo ErrorHandler 
 
 Dim myInspector As Inspector 
 
 Dim myItem As Object 
 
 
 
 Set myInspector = Application.ActiveInspector 
 
 myInspector.SetCurrentFormPage ("All Fields") 
 
 Set myItem = myInspector.CurrentItem 
 
 myItem.Display 
 
Exit Sub 
 
 
 
ErrorHandler: 
 
 MsgBox Err.Description, vbInformation 
 
End Sub
```




## Related Topics

[Inspector Object Members](acd3e13f-4727-7966-d2a5-a95e4528425c.md)

[Inspector Object](d7384756-669c-0549-1032-c3b864187994.md)




