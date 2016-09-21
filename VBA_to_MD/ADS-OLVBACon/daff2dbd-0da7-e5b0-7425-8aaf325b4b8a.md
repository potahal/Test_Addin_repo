

---
ms.Toctitle:OlkInfoBar.MouseUp イベント (Outlook)(機械翻訳)
title:OlkInfoBar.MouseUp イベント (Outlook)(機械翻訳)
ms.ContentId:daff2dbd-0da7-e5b0-7425-8aaf325b4b8a
---
# OlkInfoBar.MouseUp イベント (Outlook)(機械翻訳)




ユーザーがコントロール上で押していたマウス ボタンを離した後に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**MouseUp**(**Button**, **Shift**, **X**, **Y**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **OlkInfoBar** オブジェクトを表す変数です。

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





## Related Topics

[OlkInfoBar オブジェクト](1aec19db-d28b-ef9b-3227-45aa4a296de6.md)

[OlkInfoBar オブジェクトのメンバー](e7675cde-b1f0-153a-f4a9-b2d3bf5a0aff.md)




