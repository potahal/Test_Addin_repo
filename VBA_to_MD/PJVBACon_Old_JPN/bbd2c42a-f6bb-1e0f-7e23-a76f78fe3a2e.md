
# Task.ActualOvertimeWork プロパティ (Project)

取得実績超過作業時間 (分) タスクの。読み取り専用 **バリアント** です。


## 構文

 _式_. **ActualOvertimeWork**

 _式_ **Task** オブジェクトを表す変数です。


## 例

次の使用例は、超過作業時間が設定されているタスクの総コストを計算して、超過時間のコストを示します。また、タスクごとのコスト詳細も表示します。


```
Sub PriceOfOvertime() 
 Dim T As Task 
 Dim Price As Variant 
 Dim Breakdown As String 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 If T.ActualOvertimeWork <> 0 Then 
 Price = Price + T.ActualOvertimeCost 
 Breakdown = Breakdown &amp; T.Name &amp; ": " &amp; _ 
 ActiveProject.CurrencySymbol &amp; _ 
 T.ActualOvertimeCost &amp; vbCrLf 
 End If 
 End If 
 Next T 
 
 If Breakdown <> "" Then 
 MsgBox Breakdown &amp; vbCrLf &amp; "Total: " &amp; _ 
 ActiveProject.CurrencySymbol &amp; Price 
 End If 
 
End Sub
```

