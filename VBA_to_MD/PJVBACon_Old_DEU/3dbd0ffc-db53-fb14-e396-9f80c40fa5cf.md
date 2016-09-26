
# Resource.Flag20 Property (Project)

 **True,** Wenn das Flag einer **Ressource** zugeordnet festgelegt ist. Lese-/Schreibzugriff **Variant**.


## Syntax

 _Ausdruck_. **Flag20**

 _Ausdruck_ Eine Variable, die ein **Resource** -Objekt darstellt.


## Beispiel

Das folgende Beispiel löscht alle Aufgaben, die die  **Attribut1** auf **True** festgelegt haben.


```
Sub DeleteNonEssentialTasks() 
 
 Dim T As Task ' Task object used in For Each loop 
 
 ' Delete nonessential tasks in the active project. 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 If T.Flag1 = True Then T.Delete 
 End If 
 Next T 
 
End Sub
```

