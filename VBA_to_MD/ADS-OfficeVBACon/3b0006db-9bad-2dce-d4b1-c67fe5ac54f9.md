

---
ms.Toctitle:WorkflowTasks オブジェクト (Office)
title:WorkflowTasks オブジェクト (Office)
ms.ContentId:3b0006db-9bad-2dce-d4b1-c67fe5ac54f9
---
# WorkflowTasks オブジェクト (Office)




**WorkflowTask**オブジェクトのコレクションを表します。

## 例
次の例では、現在のドキュメント内の各ワークフロー タスクの名前を表示し、特定のタスクのワークフロー タスク編集ユーザー インターフェイスを表示します。**GetWorkflowTasks**メソッドを呼び出すことによって、サーバーへのラウンドト リップが含まれることに注意してください。

```vba
Sub DisplayWorkTask() 
Dim objWorkflowTasks As WorkflowTasks 
Dim objWorkflowTask As WorkflowTask 
Dim cnt As Integer 
 
Set objWorkflowTasks = Document.GetWorkflowTasks() 
 
For cnt = 1 To objWorkflowTasks.Count 
 Debug.Print objWorkflowTask(cnt).Name 
Next 
 
Set objWorkflowTask = objWorkflowTasks(1) 
objWorkflowTask.Show 
 
End Sub 

```




## Related Topics

[オブジェクト モデル リファレンス](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)

[WorkflowTasks オブジェクトのメンバー](a627f77c-fd47-ef66-edbd-9b4c4fcd9920.md)




