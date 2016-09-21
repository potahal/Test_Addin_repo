

---
ms.Toctitle:Assignment.Flag17 プロパティ (Project)
title:Assignment.Flag17 プロパティ (Project)
ms.ContentId:cda8dbba-c35c-86a8-348b-ed0ac4a15db5
---
# Assignment.Flag17 プロパティ (Project)




**True の****割り当て**に関連付けられたフラグが設定されている場合です。読み取り/書き込み**バリアント**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Flag17**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Assignment** オブジェクトを表す変数です。



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





