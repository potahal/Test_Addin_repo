

---
ms.Toctitle:Table.TableFields プロパティ (Project)
title:Table.TableFields プロパティ (Project)
ms.ContentId:2db4b5fd-6238-b4ab-dc9f-5de991eaad8e
---
# Table.TableFields プロパティ (Project)




テーブル内のフィールドを表す**テーブル**のコレクションを取得します。読み取り専用**テーブル**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**TableFields**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Table** オブジェクトを表す変数です。



## 例
次の例は、入力テーブルの列の文字列配置を変更します。このマクロは、ユーザーに中央揃えにする列を指定するよう求めるメッセージを表示し、次に表示を更新してビューを再表示します。

```vba
Sub AutoWrap() 
 Dim fieldNumber As Integer 
 
 fieldNumber = InputBox$(Prompt:="Enter the number of the " _ 
 & "column you want to center in the Entry table." _ 
 & Chr(13) & "For example, Column 1 is the Indicators " _ 
 & "column.") 
 
 ActiveProject.TaskTables("Entry").TableFields(fieldNumber _ 
 + 1).AlignData = pjCenter 
 
 TableApply Name:="&Entry" 
End Sub
```





