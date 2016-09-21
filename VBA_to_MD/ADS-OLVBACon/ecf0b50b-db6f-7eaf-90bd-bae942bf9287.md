

---
ms.Toctitle:Application.Quit イベント (Outlook)(機械翻訳)
title:Application.Quit イベント (Outlook)(機械翻訳)
ms.ContentId:ecf0b50b-db6f-7eaf-90bd-bae942bf9287
---
# Application.Quit イベント (Outlook)(機械翻訳)





          UNRESOLVED_TOKEN_VAL(outlooknv1) が終了するときに発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Quit**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを返すオブジェクト式を指定します。



## 注釈
このイベントは、Microsoft Visual  Basic Scripting Edition (VBScript) では使用できません。



## 例
次の Microsoft Visual Basic for Applications (VBA) の例は、Outlook が終了するときにお別れのあいさつを表示します。このサンプル コードはクラス モジュールに置いてください。



```vba
Private Sub Application_Quit() 
 
 MsgBox "Goodbye, " & Application.GetNamespace("MAPI").CurrentUser 
 
End Sub
```




## Related Topics

[Application オブジェクト](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)

[Application オブジェクト メンバー](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)




