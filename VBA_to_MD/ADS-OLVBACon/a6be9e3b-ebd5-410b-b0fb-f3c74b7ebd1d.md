

---
ms.Toctitle:ViewField.Session プロパティ (Outlook)(機械翻訳)
title:ViewField.Session プロパティ (Outlook)(機械翻訳)
ms.ContentId:a6be9e3b-ebd5-410b-b0fb-f3c74b7ebd1d
---
# ViewField.Session プロパティ (Outlook)(機械翻訳)




現在のセッションの**名前空間**のオブジェクトを返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Session**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ViewField** オブジェクトを表す変数を指定します。



## 注釈
**セッション**のプロパティは、 **GetNamespace**メソッドは、現在のセッションの**名前空間**のオブジェクトを取得するのには同じ意味で使用できます。両方のメンバーでは、同じ目的を果たします。たとえば、次のステートメントは、同じ機能を実行します。

```vba
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vba
Set objSession = Application.Session
```




## Related Topics

[ViewField オブジェクトのメンバー](7269ccc0-7dca-f0ce-2aed-b6cc7b435cf7.md)

[ViewField オブジェクト](997319f0-7ff3-a712-8484-2e442965e187.md)




