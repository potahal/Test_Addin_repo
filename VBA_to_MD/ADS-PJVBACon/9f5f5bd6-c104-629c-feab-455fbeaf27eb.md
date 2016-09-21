

---
ms.Toctitle:Resource.Group プロパティ (Project)
title:Resource.Group プロパティ (Project)
ms.ContentId:9f5f5bd6-c104-629c-feab-455fbeaf27eb
---
# Resource.Group プロパティ (Project)




リソースが属しているグループを設定します。値の取得および設定が可能です。文字列型 (**String**) の値を使用します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Group**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Resource** オブジェクトを表す変数です。



## 例
次の使用例は、作業中のプロジェクトのリソースのうち、指定されたグループに属しているリソースを削除します。

```vba
Sub DeleteResourcesInGroup() 
 
 Dim Entry As String ' The group specified by the user 
 Dim Deletions As Integer ' The number of deleted resources 
 Dim R As Resource ' The resource object used in loop 
 
 ' Prompt user for the name of a group. 
 Entry = InputBox$("Enter a group name:") 
 
 ' Cycle through the resources of the active project. 
 For Each R in ActiveProject.Resources 
 ' Delete a resource if its group name matches the user's request. 
 If R.Group = Entry Then 
 R.Delete 
 Deletions = Deletions + 1 
 End If 
 Next R 
 
 ' Display the number of resources that were deleted. 
 MsgBox(Deletions & " resources were deleted.") 
 
End Sub
```





