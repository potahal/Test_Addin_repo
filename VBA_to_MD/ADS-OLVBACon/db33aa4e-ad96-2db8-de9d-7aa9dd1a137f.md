

---
ms.Toctitle:PropertyAccessor.Session プロパティ (Outlook)(機械翻訳)
title:PropertyAccessor.Session プロパティ (Outlook)(機械翻訳)
ms.ContentId:db33aa4e-ad96-2db8-de9d-7aa9dd1a137f
---
# PropertyAccessor.Session プロパティ (Outlook)(機械翻訳)




現在のセッションの**名前空間**のオブジェクトを返します。







読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Session**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **PropertyAccessor** オブジェクトを表す変数を指定します。



## 注釈
**セッション**のプロパティと**Application.GetNamespace**メソッドは、現在のセッションの**名前空間**のオブジェクトを取得するのには同じ意味で使用できます。両方のメンバーでは、同じ目的を果たします。たとえば、次のステートメントは、同じ機能を実行します。

```vba
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vba
Set objSession = Application.Session
```




## Related Topics

[PropertyAccessor オブジェクトのメンバー](3356e345-8878-0ed7-6783-1e49ddecc066.md)

[PropertyAccessor オブジェクト](2fc91e13-703c-3ec9-9066-ffee7144306c.md)




