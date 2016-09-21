

---
ms.Toctitle:DistListItem.Session プロパティ (Outlook)(機械翻訳)
title:DistListItem.Session プロパティ (Outlook)(機械翻訳)
ms.ContentId:c36e7ef0-baf0-daa3-2e9a-c8d9d78bb6d5
---
# DistListItem.Session プロパティ (Outlook)(機械翻訳)




現在のセッションの**名前空間**のオブジェクトを返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Session**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **DistListItem** オブジェクトを表す変数を指定します。



## 注釈
**セッション**のプロパティは、 **GetNamespace**メソッドは、現在のセッションの**名前空間**のオブジェクトを取得するのには同じ意味で使用できます。両方のメンバーでは、同じ目的を果たします。などの次のステートメントは、同じ機能を実行します。

```vba
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vba
Set objSession = Application.Session
```




## Related Topics

[配布リスト オブジェクト](027c3986-abff-d9b1-ecc2-26d60805e952.md)

[配布リスト オブジェクトのメンバー](3ba4af84-ce84-61d9-1bc9-fab41bf6f125.md)




