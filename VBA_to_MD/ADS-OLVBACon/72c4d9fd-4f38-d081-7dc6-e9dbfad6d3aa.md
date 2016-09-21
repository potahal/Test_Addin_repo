

---
ms.Toctitle:Conversation.GetRootItems メソッド (Outlook)(機械翻訳)
title:Conversation.GetRootItems メソッド (Outlook)(機械翻訳)
ms.ContentId:72c4d9fd-4f38-d081-7dc6-e9dbfad6d3aa
---
# Conversation.GetRootItems メソッド (Outlook)(機械翻訳)




スレッドのすべてのルート アイテムを含む **SimpleItems** コレクションを返します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetRootItems**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Conversation** オブジェクトを表す変数を指定します。

### 戻り値
ルート アイテムまたは会話のすべてのルート アイテムを含む**SimpleItems**コレクションです。





## 注釈
1 つのスレッドには、1 つ以上のルート アイテムが存在できます。たとえば、スレッドのルート アイテムに 3 つの子アイテムがあり、このルート アイテムが完全に削除された場合、3 つの子アイテムすべてがルート アイテムになります。



**会話**のオブジェクトが取得された後、会話からすべての項目は削除すると、 **GetRootItems**は、オブジェクトの存在の**SimpleItems**コレクションを返します。この例では、 **SimpleItems**コレクションの**Count**プロパティは 0 を返します。



## Related Topics

[オブジェクトのメンバーを会話](09ff1e8e-7c5a-0b1e-e8e2-e259f66f71c8.md)

[会話オブジェクト](2705d38a-ebc0-e5a7-208b-ffe1f5446b1b.md)




