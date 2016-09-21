

---
ms.Toctitle:OlkOptionButton.KeyUp イベント (Outlook)(機械翻訳)
title:OlkOptionButton.KeyUp イベント (Outlook)(機械翻訳)
ms.ContentId:abca8eca-b1a2-a050-0418-daa10cc4cabc
---
# OlkOptionButton.KeyUp イベント (Outlook)(機械翻訳)




ユーザーがキーを離したときに発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**KeyUp**(**KeyCode**, **Shift**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **OlkOptionButton** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*KeyCode*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**長整数型**|押されていたキーを表す数値です。|
|*Shift*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**整数型 (Integer)**|**Shift キーを押し**、 **ctrl キー**、または**ALT**キーが押されたかどうかを指定する**OlShiftState**列挙の定数のビットごとの OR マスクです。|





## 注釈
**KeyUp**イベント中に押された修飾子キー (**shift キーを押し**、 **ctrl キー**、または**alt キーを押し**) の状態は、*シフト*パラメーターを通じてアクセスします。



## Related Topics

[OlkOptionButton オブジェクト](a7aab427-a2f0-a153-f558-c13559610c99.md)

[OlkOptionButton オブジェクトのメンバー](e5d545e6-496f-6a11-af73-faa3eb20647c.md)




