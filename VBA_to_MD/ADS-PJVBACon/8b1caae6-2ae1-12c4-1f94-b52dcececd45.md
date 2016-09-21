

---
ms.Toctitle:Period.Count プロパティ (Project)
title:Period.Count プロパティ (Project)
ms.ContentId:8b1caae6-2ae1-12c4-1f94-b52dcececd45
---
# Period.Count プロパティ (Project)




**期間**オブジェクト内の日数を取得します。 読み取り専用**の整数**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Count**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Period** オブジェクトを表す変数です。



## 例
次の使用例は、ユーザーにリソースの名前を入力するように求めるメッセージを表示し、そのリソースをリソースの割り当てがないタスクに割り当てます。

```vba
Sub AssignResource() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim R As Resource ' Resource object used in For Each loop 
 Dim Rname As String ' Resource name 
 Dim RID As Long ' Resource ID 
 
 RID = 0 
 RName = InputBox$("Enter the name of a resource: ") 
 
 For Each R in ActiveProject.Resources 
 If R.Name = RName Then 
 RID = R.ID 
 Exit For 
 End If 
 Next R 
 
 If RID <> 0 Then 
 ' Assign the resource to tasks without any resources. 
 For Each T In ActiveProject.Tasks 
 If T.Assignments.Count = 0 Then 
 T.Assignments.Add ResourceID:=RID 
 End If 
 Next T 
 Else 
 MsgBox Prompt:=RName & " is not a resource in this project.", buttons:=vbExclamation 
 End If 
 
End Sub
```





