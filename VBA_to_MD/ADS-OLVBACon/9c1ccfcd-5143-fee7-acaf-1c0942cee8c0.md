

---
ms.Toctitle:TaskRequestAcceptItem.BeforeAttachmentPreview イベント (Outlook)(機械翻訳)
title:TaskRequestAcceptItem.BeforeAttachmentPreview イベント (Outlook)(機械翻訳)
ms.ContentId:9c1ccfcd-5143-fee7-acaf-1c0942cee8c0
---
# TaskRequestAcceptItem.BeforeAttachmentPreview イベント (Outlook)(機械翻訳)




親オブジェクトのインスタンスに関連付けられた添付ファイルがプレビューされる前に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**BeforeAttachmentPreview**(**Attachment**, **Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskRequestAcceptItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Attachment*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Attachment**|**添付ファイル**のプレビューを表示します。|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|**True**に設定すると、操作を取り消すそれ以外の場合、**添付ファイル**のプレビューを表示するを許可するを**False**に設定します。|





## 注釈
このイベントは、添付ファイルをアクティブなエクスプローラーの閲覧ウィンドウの添付ファイル ストップ、またはアクティブなインスペクターからプレビューする前に発生します。



## Related Topics

[TaskRequestAcceptItem オブジェクトのメンバー](fe91c4cc-f505-11d8-0d0a-84fc4d355651.md)

[TaskRequestAcceptItem オブジェクト](a2905f72-0a67-b07d-7f85-84fe4de17c25.md)




