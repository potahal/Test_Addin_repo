

---
ms.Toctitle:TaskRequestDeclineItem.GetConversation メソッド (Outlook)(機械翻訳)
title:TaskRequestDeclineItem.GetConversation メソッド (Outlook)(機械翻訳)
ms.ContentId:2c6cdc44-3fb0-5cbc-dae4-a14ae2ed1fda
---
# TaskRequestDeclineItem.GetConversation メソッド (Outlook)(機械翻訳)




現在のアイテムが属しているスレッドを表す **Conversation** オブジェクトを取得します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetConversation**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskRequestDeclineItem** オブジェクトを表す変数を指定します。

### 戻り値
**会話**を表すオブジェクトは、この項目が所属する会話です。





## 注釈
**GetConversation**は、項目の会話が存在しない場合は**Null** (**Nothing**で Visual Basic) を返します。次のシナリオ内の項目の会話は存在しません。

- アイテムが保存されていません。



アイテムは、自動保存、ユーザーの操作によって、プログラムを使用して、保存できます。
- 送信可能なアイテム (メール アイテム、予定アイテム、連絡先アイテムなど) が送信されていない。
- Windows レジストリによって、スレッドが無効になっている。
- ストアでスレッド ビューがサポートされていない (たとえば、UNRESOLVED_TOKEN_VAL(ex14long) より前のバージョンの Microsoft Exchange に対して、Outlook が従来のオンライン モードで実行されている)。ストアでスレッド ビューがサポートされているかどうかを判断するには、**Store** オブジェクトの **IsConversationEnabled** プロパティを使用します。








## Related Topics

[TaskRequestDeclineItem オブジェクト](e842c7c0-7943-9219-329b-30b892ab99b0.md)

[TaskRequestDeclineItem オブジェクトのメンバー](3de31d0d-2444-876c-5d4d-1192851301af.md)




