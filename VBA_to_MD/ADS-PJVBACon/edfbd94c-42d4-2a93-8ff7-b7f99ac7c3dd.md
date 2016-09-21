

---
ms.Toctitle:Task.Flag7 プロパティ (Project)
title:Task.Flag7 プロパティ (Project)
ms.ContentId:edfbd94c-42d4-2a93-8ff7-b7f99ac7c3dd
---
# Task.Flag7 プロパティ (Project)




取得またはタスクのフラグのカスタム フィールドの値を設定します。読み取り/書き込み**バリアント**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Flag7**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Task** オブジェクトを表す変数です。



## 例
次の使用例は、 **Flag1**が**True**に設定をされているすべてのタスクを削除します。

```vba
Sub DeleteNonEssentialTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 
 ' Delete nonessential tasks in the active project. 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 If T.Flag1 = True Then T.Delete 
 End If 
 Next T 
 
End Sub
```





