

---
ms.Toctitle:コントロールの名前を変更する方法
title:コントロールの名前を変更する方法
ms.ContentId:cc3adf8b-526b-9eed-aabd-34991be2a85e
---
# コントロールの名前を変更する方法




次のコード例」を選択します「に"Test"をという名前のページで、Microsoft Forms 2.0**名前**のプロパティ**] チェック ボックス**を設定するのには、現在**Inspector**オブジェクトの**ModifiedFormPages**プロパティを使用します。

```sourcecode
Item.GetInspector.ModifiedFormPages("Test").Checkbox1.Name = "Selection"
```


>[!NOTE]
>コントロールには、固有の名前を付けるようにしてください。




