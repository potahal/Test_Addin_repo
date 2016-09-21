

---
ms.Toctitle:フィールドを参照する
title:フィールドを参照する
ms.ContentId:2ddf0f1d-f889-d631-caf2-af5d80c6b9ef
---
# フィールドを参照する




アイテム内のフィールドにアクセスする必要がある場合、そのフィールドが Outlook の標準の組み込みフィールドであるか、またはユーザー定義フィールドであるかによって、使用するメソッドが異なります。



どちらの場合も、フィールドには直接アクセスしないでください。	直接アクセスするのではなく、対象となるアイテムのプロパティとしてフィールドを参照します。



たとえば、メール メッセージの "件名" フィールドからテキストを取得するには、次の VBScript の例に示すように、そのアイテムの Subject**Subject** プロパティを使用します。

```sourcecode
mySubject = Item.Subject
```




フィールドがユーザー定義のフィールドである場合、次の VBScript の例に示すように、そのアイテムの UserProperties**UserProperties** プロパティを使用してアクセスします。この例は、アイテムに ReferredBy という名前のユーザー定義フィールドが既に存在することを前提としています。

```sourcecode
MyReferral = Item.UserProperties("ReferredBy")
```



