

---
ms.Toctitle:OlkTextBox.MouseMove イベント (Outlook)(機械翻訳)
title:OlkTextBox.MouseMove イベント (Outlook)(機械翻訳)
ms.ContentId:431bf2ee-6c9f-6dd9-5c9a-dde84acd87db
---
# OlkTextBox.MouseMove イベント (Outlook)(機械翻訳)




マウス ポインターがコントロール上を通過した後に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**MouseMove**(**Button**, **Shift**, **X**, **Y**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **OlkTextBox** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Button*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**整数型 (Integer)**|押されたマウス ボタンを示す **OlMouseButton** の定数です。|
|*Shift*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**整数型 (Integer)**|**Shift キーを押し**、 **ctrl キー**、または**ALT**キーが押されたかどうかを指定する**OlShiftState**列挙の定数のビットごとの OR マスクです。|
|*X*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**[OLE_XPOS_CONTAINER]**|マウス ポインターの X 軸上の位置を、フォームからの相対位置で示します。|
|*Y*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**[OLE_YPOS_CONTAINER]**|マウス ポインターの Y 軸上の位置を、フォームからの相対位置で示します。|





## 注釈
**ALT**キーを押すと、 **MouseMove**イベントが発生します。



## Related Topics

[ようにオブジェクト](8c9438bf-e20a-2f70-90ac-097cf09594ca.md)

[ようにオブジェクトのメンバー](f4a5f9ea-15f7-164e-d7ca-77a0842105c8.md)




