

---
ms.Toctitle:ContactItem.BeforeAttachmentSave イベント (Outlook)(機械翻訳)
title:ContactItem.BeforeAttachmentSave イベント (Outlook)(機械翻訳)
ms.ContentId:c4c33ade-25db-f9d9-69fb-97dcce76bf45
---
# ContactItem.BeforeAttachmentSave イベント (Outlook)(機械翻訳)




添付ファイルが保存される直前に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**BeforeAttachmentSave**(**Attachment**, **Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ContactItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Attachment*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Attachment**|**添付ファイル**を保存します。|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。場合は、イベント プロシージャでは、この引数を設定する**場合は True**、保存操作は完了せず、添付ファイルは変更されません。|





## 注釈
このイベントは、添付ファイルは、メッセージ ストアに保存するときに対応します。**BeforeAttachmentSave**イベントは、アイテムが保存されるとき、添付ファイルが保存される直前に発生します。ユーザは、添付ファイルを編集し、それらの変更を保存し場合、 **BeforeAttachmentSave**イベントはその時点では発生しません代わりにアイテム自体を後で保存するときに発生します。行われなかった、 **SaveAsFile**メソッドを使用してハード ディスクに添付ファイルを保存するとします。



Vbscript の場合、 **False**を保存するこの関数の戻り値を設定する操作は取り消され、添付ファイルは変更されません。



## Related Topics

[ContactItem オブジェクト](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)

[ContactItem オブジェクトのメンバー](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)




