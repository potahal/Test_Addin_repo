

---
ms.Toctitle:TaskItem.Categories プロパティ (Outlook)(機械翻訳)
title:TaskItem.Categories プロパティ (Outlook)(機械翻訳)
ms.ContentId:c4099fe0-23af-a4cb-dfef-92cbe0c6e600
---
# TaskItem.Categories プロパティ (Outlook)(機械翻訳)




Outlook アイテムに割り当てられているカテゴリを表す**文字列**を設定または返します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Categories**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskItem** オブジェクトを表す変数を指定します。



## 注釈
**カテゴリ**は、Outlook アイテムに割り当てられているカテゴリの名前の区切り記号付きの文字列です。このプロパティは、複数の分類項目の区切り記号として、値の名前、 **sList**、Windows レジストリに**HKEY_CURRENT_USER\Control Panel\International**の下に指定された文字を使用します。カテゴリ名の文字列を項目名の配列に変換するには、Microsoft Visual Basic 関数**Split**を使用します。



## Related Topics

[TaskItem オブジェクトの場合](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)

[TaskItem オブジェクトのメンバー](97234a76-2fc5-bbe4-2e14-25ae18694fc9.md)




