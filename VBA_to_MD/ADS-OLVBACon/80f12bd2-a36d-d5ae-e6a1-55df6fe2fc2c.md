

---
ms.Toctitle:ContactItem.Open イベント (Outlook)(機械翻訳)
title:ContactItem.Open イベント (Outlook)(機械翻訳)
ms.ContentId:80f12bd2-a36d-d5ae-e6a1-55df6fe2fc2c
---
# ContactItem.Open イベント (Outlook)(機械翻訳)




親オブジェクトのインスタンスを **Inspector** で開こうとすると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Open**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ContactItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合ファイルを開く操作は完了せず、インスペクターが表示されていません。|





## 注釈
このイベントが発生すると、 **Inspector**オブジェクトが初期化されていますが表示されていません。**ユーザーが直接対応しているで、インスペクターでアイテムが開かれるとき、編集ビューでアイテムを選択するときにも発生**、 **Open**イベントは**Read**イベントとは異なります。



で Microsoft Visual Basic スクリプト版 (VBScript)、この関数の戻り値を**False**に設定する場合は、ファイルを開く操作は完了せず、インスペクターは表示されません。



## Related Topics

[ContactItem オブジェクト](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)

[ContactItem オブジェクトのメンバー](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)




