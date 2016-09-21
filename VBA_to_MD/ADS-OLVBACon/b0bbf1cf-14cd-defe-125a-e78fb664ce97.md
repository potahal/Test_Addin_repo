

---
ms.Toctitle:PostItem.Open イベント (Outlook)(機械翻訳)
title:PostItem.Open イベント (Outlook)(機械翻訳)
ms.ContentId:b0bbf1cf-14cd-defe-125a-e78fb664ce97
---
# PostItem.Open イベント (Outlook)(機械翻訳)




親オブジェクトのインスタンスを **Inspector** で開こうとすると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Open**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **PostItem** オブジェクトを表す変数を指定します。

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

[PostItem オブジェクトのメンバー](5b150db1-c96d-0721-ec36-d5b5ebc20fd8.md)

[PostItem オブジェクト](de44065d-4e93-315a-279f-7b92f09c0465.md)




