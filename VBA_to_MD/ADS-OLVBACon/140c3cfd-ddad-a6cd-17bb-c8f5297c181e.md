

---
ms.Toctitle:OlkListBox.AfterUpdate イベント (Outlook)(機械翻訳)
title:OlkListBox.AfterUpdate イベント (Outlook)(機械翻訳)
ms.ContentId:140c3cfd-ddad-a6cd-17bb-c8f5297c181e
---
# OlkListBox.AfterUpdate イベント (Outlook)(機械翻訳)




ユーザー インターフェイスを介してコントロールのデータが変更された後に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AfterUpdate**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **OlkListBox** オブジェクトを表す変数を指定します。



## 注釈
**BeforeUpdate**および**AfterUpdate**は、コントロール内のデータ アイテムに保存されている任意の時間を発生します。



このコントロールに対して**AfterUpdate**を含むイベントの一般的な順序は次のとおりです。

1. ユーザーがコントロールにフォーカスを移動する
2. **BeforeUpdate**
3. コントロールのデータが更新される
4. **AfterUpdate**
5. **終了**: ユーザー コントロールからフォーカスを移動します。








## Related Topics

[OlkListBox オブジェクト](373d2a00-97e5-2ed3-f15f-577d97b32334.md)

[OlkListBox オブジェクトのメンバー](b8bed0b5-6994-1492-055e-4067b232f9c4.md)




