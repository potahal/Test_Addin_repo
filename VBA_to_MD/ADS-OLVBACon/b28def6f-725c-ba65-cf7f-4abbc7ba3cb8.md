

---
ms.Toctitle:SelectNamesDialog.CcLabel プロパティ (Outlook)(機械翻訳)
title:SelectNamesDialog.CcLabel プロパティ (Outlook)(機械翻訳)
ms.ContentId:b28def6f-725c-ba65-cf7f-4abbc7ba3cb8
---
# SelectNamesDialog.CcLabel プロパティ (Outlook)(機械翻訳)




**[名前の選択**] ダイアログ ボックスの [ **Cc** ] コマンド ボタンに表示されるテキストの**文字列**を設定または返します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CcLabel**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SelectNamesDialog** オブジェクトを表す変数を指定します。



## 注釈
受信者の編集ボックスのアクセラレータ キーを提供するには、アクセス キーとして機能する文字の直前にアンパサンド (&)、ラベル引数の文字列内の文字を含めます。たとえば、 **CcLabel**が"ローカル & 出席者"の文字列の場合は、ユーザーが最初の受信者の編集ボックスにフォーカスを移動するのには**alt キーを押しながら A**を押します。



**CcLabel**を設定しない場合、既定値になります"Cc"のローカライズされた文字列。**CcLabel**を空の文字列に設定すると、する場合、[ **Cc** ] コマンド ボタンが表示されます**->**。**CcLabel**プロパティに複数の 32 文字 (アンパサンド (&) アクセス キーを含む) が含まれている場合の最初の 32 文字だけがコマンド ボタンに表示します。



## Related Topics

[SelectNamesDialog オブジェクト](1522736a-3cad-9f1c-4da9-b52a3a01731c.md)

[SelectNamesDialog オブジェクトのメンバー](0f5546af-f89a-8a8b-ced9-a2d646bf9634.md)




