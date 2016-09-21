

---
ms.Toctitle:ContactItem.Session プロパティ (Outlook)(機械翻訳)
title:ContactItem.Session プロパティ (Outlook)(機械翻訳)
ms.ContentId:b67eb0d4-9b97-2be7-fc24-ecdd58fb01ca
---
# ContactItem.Session プロパティ (Outlook)(機械翻訳)




現在のセッションの**名前空間**のオブジェクトを返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Session**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ContactItem** オブジェクトを表す変数を指定します。



## 注釈
**セッション**のプロパティは、 **GetNamespace**メソッドは、現在のセッションの**名前空間**のオブジェクトを取得するのには同じ意味で使用できます。両方のメンバーでは、同じ目的を果たします。などの次のステートメントは、同じ機能を実行します。

```vba
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vba
Set objSession = Application.Session
```




## Related Topics

[ContactItem オブジェクトのメンバー](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)

[ContactItem オブジェクト](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)




