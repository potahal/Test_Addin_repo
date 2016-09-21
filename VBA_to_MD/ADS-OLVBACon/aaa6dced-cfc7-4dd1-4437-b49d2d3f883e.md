

---
ms.Toctitle:DistListItem.Close イベント (Outlook)(機械翻訳)
title:DistListItem.Close イベント (Outlook)(機械翻訳)
ms.ContentId:aaa6dced-cfc7-4dd1-4437-b49d2d3f883e
---
# DistListItem.Close イベント (Outlook)(機械翻訳)




アイテム (親オブジェクトのインスタンス) に関連付けられたインスペクターが閉じるときに発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Close**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **DistListItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合は、閉じる操作は完了せずと、インスペクターは開いたまま。|





## 注釈
で Microsoft Visual Basic スクリプト版 (VBScript)、この関数の戻り値を**False**に設定する場合は、閉じる操作は完了せず、インスペクターは開いたままです。



**Close**メソッドを使用して、このイベントが発生する場合、取り消すことができます**Close**メソッドが**呼び出すことにより**使用されている場合。



## Related Topics

[配布リスト オブジェクトのメンバー](3ba4af84-ce84-61d9-1bc9-fab41bf6f125.md)

[配布リスト オブジェクト](027c3986-abff-d9b1-ecc2-26d60805e952.md)




