
# Task.PercentWorkComplete プロパティ (Project)

取得またはタスクの完了作業時間の割合を設定します。 サマリー タスクに対しては読み取り専用です。読み取り/書き込み **バリアント** です。


## 構文

 _式_. **PercentWorkComplete**

 _式_ **Task** オブジェクトを表す変数を指定します。


## 例

次の使用例は、作業中のプロジェクトの作業時間の達成をユーザーが指定した割合を超える割合で各タスクの **true を指定** する **マーク** のプロパティを設定します。


```
Sub MarkTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 Dim Entry As String ' Percentage entered by user 
 
 ' Prompt user for a percentage. 
 Entry = InputBox$("Mark tasks that exceed what percentage of work complete? (0-100)") 
 
 If Not IsNumeric(Entry) Then 
 MsgBox ("Please enter a number only.") 
 Exit Sub 
 ElseIf Entry < 0 Or Entry > 100 Then 
 MsgBox ("You did not enter a percentage from 0 to 100.") 
 Exit Sub 
 End If 
 
 ' Mark tasks with percentage of work complete greater than user entry. 
 For Each T In ActiveProject.Tasks 
 If T.PercentWorkComplete > Val(Entry) Then 
 T.Marked = True 
 Else 
 T.Marked = False 
 End If 
 Next T 
 
End Sub
```

