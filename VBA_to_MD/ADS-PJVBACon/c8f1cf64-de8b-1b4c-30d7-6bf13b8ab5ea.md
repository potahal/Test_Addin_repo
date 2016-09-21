

---
ms.Toctitle:Resource.Flag18 プロパティ (Project)
title:Resource.Flag18 プロパティ (Project)
ms.ContentId:c8f1cf64-de8b-1b4c-30d7-6bf13b8ab5ea
---
# Resource.Flag18 プロパティ (Project)




**該当**の**リソース**に関連付けられたフラグが設定されている場合です。読み取り/書き込み**バリアント**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Flag18**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Resource** オブジェクトを表す変数です。



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





