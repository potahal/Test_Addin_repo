

---
ms.Toctitle:DistListItem.GetInspector プロパティ (Outlook)(機械翻訳)
title:DistListItem.GetInspector プロパティ (Outlook)(機械翻訳)
ms.ContentId:2ffab19b-17a3-0de0-f9dd-3a8fbfea8efd
---
# DistListItem.GetInspector プロパティ (Outlook)(機械翻訳)




指定した項目を含むインスペクターを表す**Inspector**オブジェクトを取得します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetInspector**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **DistListItem** オブジェクトを表す変数を指定します。



## 注釈
このプロパティは、 **Application.ActiveInspector**メソッドを使用して、 **Inspector.CurrentItem**プロパティを設定するのではなく、アイテムを表示する**インスペクター**オブジェクトを取得するのに便利です。アイテムの**Inspector**オブジェクトが既に存在する場合、 **GetInspector**プロパティは新規に作成するのではなく場合は、その**Inspector**オブジェクトを返します。



## Related Topics

[配布リスト オブジェクトのメンバー](3ba4af84-ce84-61d9-1bc9-fab41bf6f125.md)

[配布リスト オブジェクト](027c3986-abff-d9b1-ecc2-26d60805e952.md)




