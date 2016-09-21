

---
ms.Toctitle:RemoteItem.Delete メソッド (Outlook)(機械翻訳)
title:RemoteItem.Delete メソッド (Outlook)(機械翻訳)
ms.ContentId:062afc1f-4c3f-bd6e-ad88-e8fa394a5edd
---
# RemoteItem.Delete メソッド (Outlook)(機械翻訳)




アイテムを含むフォルダーからアイテムを削除します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Delete**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **RemoteItem** オブジェクトを表す変数。



## 注釈
**Delete** メソッドは、コレクション内の単一のアイテムを削除します。フォルダーの **Items** コレクション内のすべてのアイテムを削除するには、フォルダー内の最新のアイテムから順に 1 つずつ削除する必要があります。たとえば、フォルダーの items コレクションである `AllItems` で `n` 個のアイテムがフォルダーにある場合、`AllItems.Item(n)` から削除を開始し、`AllItems.Item(1)` にたどり着くまでインデックスを 1 つずつ減少しながら削除を繰り返します。



**Delete** メソッドは、アイテムをそのフォルダーから [**削除済みアイテム**] フォルダーに移動します。アイテムを含むフォルダーが  [**削除済みアイテム**] フォルダーである場合は、**Delete** メソッドは、アイテムを完全に削除します。



## Related Topics

[RemoteItemObject](6302aaff-cdcf-4d86-60f1-4bed15540d9f.md)

[RemoteItemMembers](15c0872e-88cc-9b9b-c31e-c15d6971e6e0.md)

[すべてのアイテムおよび削除済みアイテム フォルダー内のサブフォルダーを削除します。](359a416b-43d4-396e-e348-5624c4ca3599.md)




