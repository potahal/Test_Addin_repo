

---
ms.Toctitle:ContactItem.Close イベント (Outlook)(機械翻訳)
title:ContactItem.Close イベント (Outlook)(機械翻訳)
ms.ContentId:beeeb53c-94fe-ae1b-7870-87bd37b3debf
---
# ContactItem.Close イベント (Outlook)(機械翻訳)




アイテム (親オブジェクトのインスタンス) に関連付けられたインスペクターが閉じるときに発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Close**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ContactItem** オブジェクトを表す変数を指定します。

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

[ContactItem オブジェクトのメンバー](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)

[ContactItem オブジェクト](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)




