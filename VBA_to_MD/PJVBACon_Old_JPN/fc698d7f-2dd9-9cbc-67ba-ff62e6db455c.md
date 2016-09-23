
# Task.PercentComplete プロパティ (Project)

取得またはタスクの達成率を設定します。読み取り/書き込み **バリアント** です。


## 構文

 _式_. **PercentComplete**

 _式_ **Task** オブジェクトを表す変数です。


## 例

次の使用例は、達成率が 85% 以上で、2 つ以上のリソースが割り当てられているタスクを検索し、そのタスクに割り当てられているリソースのうち 1 つを削除します。


```
Sub ReallocateResource() 
 
 Dim Entry As String ' The name of the resource to remove 
 Dim T As Task ' The task object used in For loop 
 Dim RA As Assignment ' The resource assignment object to the task 
 
 Entry = InputBox$("Enter a resource name:") 
 
 ' Remove the resource from 85 percent complete tasks with 2+ resources. 
 For Each T In ActiveProject.Tasks 
 If T.PercentComplete >= 85 And T.Resources.Count >= 2 Then 
 For Each RA In T.Assignments 
 If UCase(Entry) = UCase(RA.ResourceName) Then 
 RA.Delete 
 End If 
 Next 
 End If 
 Next T 
 
End Sub
```

