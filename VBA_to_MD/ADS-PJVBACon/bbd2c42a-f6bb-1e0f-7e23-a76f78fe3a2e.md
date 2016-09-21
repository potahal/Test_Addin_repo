

---
ms.Toctitle:Task.ActualOvertimeWork プロパティ (Project)
title:Task.ActualOvertimeWork プロパティ (Project)
ms.ContentId:bbd2c42a-f6bb-1e0f-7e23-a76f78fe3a2e
---
# Task.ActualOvertimeWork プロパティ (Project)




取得実績超過作業時間 (分) タスクの。読み取り専用**バリアント**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ActualOvertimeWork**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Task** オブジェクトを表す変数です。



## 例
次の使用例は、超過作業時間が設定されているタスクの総コストを計算して、超過時間のコストを示します。また、タスクごとのコスト詳細も表示します。

```vba
Sub PriceOfOvertime() 
 Dim T As Task 
 Dim Price As Variant 
 Dim Breakdown As String 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 If T.ActualOvertimeWork <> 0 Then 
 Price = Price + T.ActualOvertimeCost 
 Breakdown = Breakdown & T.Name & ": " & _ 
 ActiveProject.CurrencySymbol & _ 
 T.ActualOvertimeCost & vbCrLf 
 End If 
 End If 
 Next T 
 
 If Breakdown <> "" Then 
 MsgBox Breakdown & vbCrLf & "Total: " & _ 
 ActiveProject.CurrencySymbol & Price 
 End If 
 
End Sub
```





