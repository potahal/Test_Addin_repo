

---
ms.Toctitle:Project.TaskTableList プロパティ (Project)
title:Project.TaskTableList プロパティ (Project)
ms.ContentId:a36abbcb-db7d-f593-7e5c-df00fd96f010
---
# Project.TaskTableList プロパティ (Project)




プロジェクト内のすべてのタスク テーブルを表す**List**オブジェクトを取得します。読み取り専用**リスト**。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**TaskTableList**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Project** オブジェクトを表す変数です。



## 例
次の使用例は、作業中のプロジェクトでタスク テーブルの一覧を表示します。

```vba
Sub SeeAllTables() 
 
 Dim Temp As Variant 
 Dim TaskTableNames As String 
 
 For Each Temp In ActiveProject.TaskTableList 
 TaskTableNames = TaskTableNames & vbCrLf & Temp 
 Next Temp 
 
 MsgBox TaskTableNames 
 
End Sub
```





