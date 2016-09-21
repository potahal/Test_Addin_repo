

---
ms.Toctitle:ReportItem.Send イベント (Outlook)(機械翻訳)
title:ReportItem.Send イベント (Outlook)(機械翻訳)
ms.ContentId:aab0b0f3-8e33-f1fa-cc74-d914effcb833
---
# ReportItem.Send イベント (Outlook)(機械翻訳)




ユーザーがアイテム (親オブジェクトのインスタンス) に対して [**送信**] アクションを選択すると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Send**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ReportItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合は、送信操作は完了せずと、インスペクターが開いたままです。|





## 注釈
Microsoft Visual Basic Scripting Edition (VBScript)、この関数の戻り値を**False**に設定した場合、アイテムは送信されません。



## Related Topics

[ReportItem オブジェクトのメンバー](5a5662dd-e969-bbd5-129b-44609ba1cf9f.md)

[ReportItem オブジェクト](16ebe336-72e0-42f6-99d3-edecc3ea284d.md)




