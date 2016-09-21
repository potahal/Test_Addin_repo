

---
ms.Toctitle:CommandBars.DisableCustomize プロパティ (Office)
title:CommandBars.DisableCustomize プロパティ (Office)
ms.ContentId:cbebdaa7-2e8d-af73-fd18-03b3b11f98ac
---
# CommandBars.DisableCustomize プロパティ (Office)




ツールバーのカスタマイズが無効になっている場合は**True**です。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**DisableCustomize**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **CommandBars** オブジェクトを表す変数を指定します。



## 例
次の使用例は、オンまたはオフに、 **DisableCustomize**プロパティを切り替えます。

```sourcecode
Sub ToggleCustomize() 
 With Application.CommandBars 
 If .DisableCustomize = True Then 
 .DisableCustomize = False 
 Else 
 .DisableCustomize = True 
 End If 
 End With 
End Sub
```




>[!NOTE]
>
              UNRESOLVED_TOKEN_VAL(osdepreccommandbars)
            





## Related Topics

[CommandBars オブジェクト](0e312e21-14ee-5055-d604-b66e61c53b47.md)

[CommandBars オブジェクトのメンバー](c11db22d-b7bb-20a2-a455-e441cb8d5bc0.md)




