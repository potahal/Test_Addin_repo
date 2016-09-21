

---
ms.Toctitle:Column.Session プロパティ (Outlook)(機械翻訳)
title:Column.Session プロパティ (Outlook)(機械翻訳)
ms.ContentId:d0bc26d3-cb93-cc0d-ed87-9b51a2d35bcc
---
# Column.Session プロパティ (Outlook)(機械翻訳)




現在のセッションの **NameSpace** オブジェクトを取得します。値の取得のみ可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Session**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Column** オブジェクトを表す変数です。



## 注釈
**セッション**のプロパティと**Application.GetNamespace**メソッドは、現在のセッションの**名前空間**のオブジェクトを取得するのには同じ意味で使用できます。両方のメンバーでは、同じ目的を果たします。たとえば、次のステートメントは、同じ機能を実行します。

```vba
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vba
Set objSession = Application.Session
```




## Related Topics

[列オブジェクト](b7eb6916-2d80-57c3-2077-47a2a4c73185.md)

[列オブジェクトのメンバー](c9b724b2-49e3-8cd5-95c7-0e4ea423df46.md)




