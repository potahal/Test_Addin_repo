

---
ms.Toctitle:Project.ResourceGroupList プロパティ (Project)
title:Project.ResourceGroupList プロパティ (Project)
ms.ContentId:a64fe8c8-e75f-3cab-e77a-54fc6c1bf0a5
---
# Project.ResourceGroupList プロパティ (Project)




作業中のプロジェクトのリソース グループを表す**List**オブジェクトを取得します。読み取り専用**リスト**。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ResourceGroupList**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Project** オブジェクトを表す変数です。



## 例
次の使用例は、作業中のプロジェクトでリソース フィルターの一覧を表示します。

```vba
Sub SeeAllResGroups() 
 
 Dim Temp As Variant 
 Dim ResGroupNames As String 
 
 For Each Temp In ActiveProject.ResourceGroupList 
 ResGroupNames = ResGroupNames & vbCrLf & Temp 
 Next Temp 
 
 MsgBox ResGroupNames 
 
End Sub
```





