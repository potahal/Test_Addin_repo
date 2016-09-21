

---
ms.Toctitle:OlkCheckBox.AfterUpdate イベント (Outlook)(機械翻訳)
title:OlkCheckBox.AfterUpdate イベント (Outlook)(機械翻訳)
ms.ContentId:a207e36b-9afe-b7e3-9dd4-9af2ae16cf7d
---
# OlkCheckBox.AfterUpdate イベント (Outlook)(機械翻訳)




ユーザー インターフェイスを介してコントロールのデータが変更された後に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AfterUpdate**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **OlkCheckBox** オブジェクトを表す変数を指定します。



## 注釈
**BeforeUpdate**および**AfterUpdate**は、コントロール内のデータ アイテムに保存されている任意の時間を発生します。



このコントロールに対して**AfterUpdate**を含むイベントの一般的な順序は次のとおりです。

1. ユーザーがコントロールにフォーカスを移動する
2. **BeforeUpdate**
3. コントロールのデータが更新される
4. **AfterUpdate**
5. **終了**: ユーザー コントロールからフォーカスを移動します。








## Related Topics

[OlkCheckBox オブジェクトのメンバー](acf62b06-215d-6b2b-57b0-ccbfd0c92aed.md)

[OlkCheckBox オブジェクト](79460205-a604-7011-a9b3-14e651807f09.md)




