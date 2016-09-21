

---
ms.Toctitle:GroupCriteria2 オブジェクト (Project)
title:GroupCriteria2 オブジェクト (Project)
ms.ContentId:ac785cc4-dbe3-0b1d-d1f1-6d45c93bfb1d
---
# GroupCriteria2 オブジェクト (Project)




グループ階層構造を維持し、セルの色を 16 進値で指定できる **GroupCriterion2** オブジェクトのコレクションを格納します。

## 例
**GroupCriterion2 オブジェクトの使い方**



単一の **GroupCriterion2** オブジェクトを取得するには、**GroupCriteria2(***Index***)** を使用します (*Index* には条件のインデックスを指定します)。次の例は、標準単価リソース グループの第 1 条件のセルの色を青に設定します。

```vba
ActiveProject.ResourceGroups2("Standard Rate").GroupCriteria2(1).CellColor = &HFF0000
```




**GroupCriteria2 コレクションの使い方**



**GroupCriteria2** コレクションを取得するには、**GroupCriteria** プロパティを使用します。次の使用例は、指定されたタスク グループの条件として使用されるフィールドの一覧と、昇順と降順のどちらで並べ替えるかを表示します。

```vba
Dim GC2 As GroupCriterion2  
Dim Fields As String  
  
For Each GC2 In ActiveProject.TaskGroups2("Priority Keeping Outline Structure").GroupCriteria  
    If GC2.Ascending = True Then  
       Fields = Fields & GC2.Index & ". " & GC2.FieldName & " is sorted in ascending order." _
           & vbCrLf  
    Else  
        Fields = Fields & GC2.Index & ". " & GC2.FieldName & " is sorted in descending order." _
           & vbCrLf  
    End If  
Next GC2  

MsgBox Fields
```




**GroupCriterion2** オブジェクトを **GroupCriteria2** コレクションに追加するには、**AddEx** メソッドを使用します (**CellColor** は 16 進値で指定できます)。次の例は、指定したリソース グループに新しいグループ化条件、つまり、作業時間の達成率 25% ごとにグループ化し、昇順に並べ替えるという条件を追加します。

```vba
ActiveProject.ResourceGroups2("Response Pending").GroupCriteria2.AddEx "% Work Complete", True, _  
    CellColor:=&H0101FF, GroupOn:=pjGroupOnPct1_25
```




## Related Topics

[プロジェクト オブジェクト モデル](900b167b-88ec-ea88-15b7-27bb90c22ac6.md)

[GroupCriteria2 オブジェクトのメンバー](b52e84f3-4332-9c5a-cd2c-c4b57cfc40ea.md)




