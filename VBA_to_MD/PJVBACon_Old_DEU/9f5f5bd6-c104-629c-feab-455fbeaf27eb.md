
# Resource.Group Property (Project)

Ruft die Gruppe ab, zu der eine Ressource gehört, oder legt sie fest.  **String** -Wert mit Lese-/Schreibzugriff.


## Syntax

 _Ausdruck_. **Group**

 _Ausdruck_ Eine Variable, die ein **Resource** -Objekt darstellt.


## Beispiel

Im folgenden Beispiel werden die Ressourcen des aktiven Projekts gelöscht, die zu einer vom Benutzer angegebenen Gruppe gehören.


```
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
 MsgBox(Deletions &amp; " resources were deleted.") 
 
End Sub
```

