

---
ms.Toctitle:RemoteItem.Write イベント (Outlook)(機械翻訳)
title:RemoteItem.Write イベント (Outlook)(機械翻訳)
ms.ContentId:a38eef6b-23da-ba10-ad94-cc63e2bf60c2
---
# RemoteItem.Write イベント (Outlook)(機械翻訳)




親オブジェクトのインスタンスが保存されると発生します。**Save** メソッドや **SaveAs** メソッドを使用した場合のような明示的な保存、またはアイテムのインスペクターを閉じるときに表示されるメッセージへの対応のような暗黙的な保存の両方で発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Write**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **RemoteItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|(未使用の vbscript)。**False**イベントが発生します。場合は、イベント プロシージャでは、この引数を設定する**場合は True**、保存操作は完了しません。|





## 注釈
Microsoft Visual Basic Scripting Edition (VBScript の) 場合は**False**を保存するこの関数の戻り値を設定する操作は完了しません。



## Related Topics

[RemoteItem オブジェクトのメンバー](15c0872e-88cc-9b9b-c31e-c15d6971e6e0.md)

[RemoteItem オブジェクト](6302aaff-cdcf-4d86-60f1-4bed15540d9f.md)




