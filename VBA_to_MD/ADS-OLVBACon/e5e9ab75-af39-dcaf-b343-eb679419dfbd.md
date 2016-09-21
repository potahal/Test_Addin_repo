

---
ms.Toctitle:PostItem.Forward イベント (Outlook)(機械翻訳)
title:PostItem.Forward イベント (Outlook)(機械翻訳)
ms.ContentId:e5e9ab75-af39-dcaf-b343-eb679419dfbd
---
# PostItem.Forward イベント (Outlook)(機械翻訳)




親オブジェクトのインスタンスであるアイテムに対し、ユーザーが "**転送**" アクションを選択するか、または **Forward** メソッドが呼び出されると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Forward**(**Forward**, **Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **PostItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Forward*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**オブジェクト型 (Object)**|転送される新しいアイテムです。|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合は、転送操作は完了せずと、新しいアイテムは表示されません。|





## 注釈
Vbscript の場合、この関数の戻り値を**False**に設定して、転送アクションは完了せず、新しいアイテムは表示されません。



## Related Topics

[PostItem オブジェクトのメンバー](5b150db1-c96d-0721-ec36-d5b5ebc20fd8.md)

[PostItem オブジェクト](de44065d-4e93-315a-279f-7b92f09c0465.md)




