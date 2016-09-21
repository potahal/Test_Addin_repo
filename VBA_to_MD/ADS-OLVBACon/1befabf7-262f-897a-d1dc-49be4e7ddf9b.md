

---
ms.Toctitle:TaskItem.Respond メソッド (Outlook)(機械翻訳)
title:TaskItem.Respond メソッド (Outlook)(機械翻訳)
ms.ContentId:1befabf7-262f-897a-d1dc-49be4e7ddf9b
---
# TaskItem.Respond メソッド (Outlook)(機械翻訳)




タスクの依頼に返信します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Respond**(**Response**, **fNoUI**, **fAdditionalTextDialog**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Response*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**OlTaskResponse**|
            依頼への返信を指定します。|
|*fNoUI*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**バリアント型 (Variant)**|**True を指定**するダイアログ ボックスは表示されません。応答が自動的に送信されます。**False**応答のダイアログ ボックスを表示します。|
|*fAdditionalTextDialog*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**バリアント型 (Variant)**|**False**の入力をユーザーにプロンプトを表示しません。応答を編集するためのインスペクターに表示されます。**True の**送信] または [コメントの送信をユーザーに確認します。この引数は、 *fNoUI*が**False**の場合にのみ有効です。|



### 戻り値
タスクの依頼への返信を表す **TaskItem**。





## 注釈
**OlTaskAccept**パラメーターを指定して**Respond**メソッドを呼び出すと、Outlook は、仕事の依頼アイテムを複製する新しい**TaskItem**を作成します。新しいアイテムが別のエントリ ID があります。Outlook は、元のアイテムを削除します。



次の表では、親オブジェクト、および*fNoUI*および*fAdditionalTextDialog*パラメーターによって**応答**のメソッドの動作について説明します。

|**fNoUI と fAdditionalTextDialog**|**結果**|
|---|---|
|**True、True**|ユーザー インターフェイスなしの返信アイテムが返されます。応答を送信するには、 **Send**メソッドを呼び出す必要があります。|
|**True、False**|**True**、**True** のときと同じです。|
|**False、True**|**Display**メソッドが呼び出された場合、ユーザー プロンプトが表示されます。それ以外の場合、メッセージを表示せず、アイテムが送信され、結果のアイテムはありません。|
|**False、False**|何も起こりません。|



## Related Topics

[TaskItem オブジェクトの場合](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)

[TaskItem オブジェクトのメンバー](97234a76-2fc5-bbe4-2e14-25ae18694fc9.md)




