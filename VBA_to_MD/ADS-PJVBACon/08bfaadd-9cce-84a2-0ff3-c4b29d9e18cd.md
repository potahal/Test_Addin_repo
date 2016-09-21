

---
ms.Toctitle:Project.Tasks プロパティ (Project)
title:Project.Tasks プロパティ (Project)
ms.ContentId:08bfaadd-9cce-84a2-0ff3-c4b29d9e18cd
---
# Project.Tasks プロパティ (Project)




プロジェクト内のタスクを表す**Tasks**コレクションを取得します。読み取り専用**タスク**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Tasks**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Project** オブジェクトを表す変数です。



## 例
次の使用例は、作業中のプロジェクト内のすべてのタスクの名前を表示します。

```vba
Sub TaskNames() 
 
 Dim T As Task, Names As String 
 
 For Each T In ActiveProject.Tasks 
 Names = Names & T.Name & vbCrLf 
 Next T 
 
 MsgBox Names 
 
End Sub
```





