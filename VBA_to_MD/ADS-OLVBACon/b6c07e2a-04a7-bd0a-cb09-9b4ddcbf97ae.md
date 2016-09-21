

---
ms.Toctitle:TaskRequestUpdateItem.Reply イベント (Outlook)(機械翻訳)
title:TaskRequestUpdateItem.Reply イベント (Outlook)(機械翻訳)
ms.ContentId:b6c07e2a-04a7-bd0a-cb09-9b4ddcbf97ae
---
# TaskRequestUpdateItem.Reply イベント (Outlook)(機械翻訳)




ユーザーがアイテム (親オブジェクトのインスタンス) に対して [**返信**] アクションを選択すると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Reply**(**Response**, **Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskRequestUpdateItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Response*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**オブジェクト型 (Object)**|元のメッセージへの返信として送信される新しいアイテムです。|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合返信操作は完了せず、新しいアイテムは表示されません。|





## 注釈
返信されるアイテムを **MailItem** オブジェクトとして返します。



で Microsoft Visual Basic スクリプト版 (VBScript)、この関数の戻り値を**False**に設定する場合は、返信アクションは完了せず、新しいアイテムは表示されません。



## Related Topics

[TaskRequestUpdateItem オブジェクトのメンバー](f4a396b3-c2f7-68a7-efa7-877328a7fc21.md)

[TaskRequestUpdateItem オブジェクト](5bc407fe-b3f6-3e46-8b91-e2ed96292cec.md)




