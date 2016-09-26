
# WorkflowTasks-Objekt (Office)

Stellt eine Auflistung von  **WorkflowTask** -Objekten dar.


## Beispiel

Im folgenden Beispiel wird der Name jeder Workflowaufgabe im aktuellen Dokument angezeigt und dann die Workflow bearbeiten Benutzeroberfläche für eine bestimmte Aufgabe angezeigt. Beachten Sie, dass die  **GetWorkflowTasks** -Methode aufrufen einen Roundtrip zum Server beinhaltet.


```
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


## Siehe auch


#### Konzepte


[-Objektmodellreferenz](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Weitere Ressourcen


[Elemente des WorkflowTasks-Objekts](http://msdn.microsoft.com/library/a627f77c-fd47-ef66-edbd-9b4c4fcd9920%28Office.15%29.aspx)