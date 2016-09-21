

---
ms.Toctitle:SharingItem.Open イベント (Outlook)(機械翻訳)
title:SharingItem.Open イベント (Outlook)(機械翻訳)
ms.ContentId:b795dbfa-2d47-0ee4-98ef-0c44bb6a0bec
---
# SharingItem.Open イベント (Outlook)(機械翻訳)




親オブジェクトのインスタンスを **Inspector** で開こうとすると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Open**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharingItem** オブジェクトを返すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合ファイルを開く操作は完了せず、インスペクターが表示されていません。|





## 注釈
このイベントが発生すると、 **Inspector**オブジェクトが初期化されていますが表示されていません。**ユーザーが直接対応しているで、インスペクターでアイテムが開かれるとき、編集ビューでアイテムを選択するときにも発生**、 **Open**イベントは**Read**イベントとは異なります。



で Microsoft Visual Basic スクリプト版 (VBScript)、この関数の戻り値を**False**に設定する場合は、ファイルを開く操作は完了せず、インスペクターは表示されません。



## Related Topics

[SharingItem オブジェクトのメンバー](719ad60e-2242-2c54-778f-006b61690389.md)

[SharingItem オブジェクト](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)




