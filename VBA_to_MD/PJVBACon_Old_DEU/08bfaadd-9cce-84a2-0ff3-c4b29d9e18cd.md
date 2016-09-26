
# Project.Tasks Property (Project)

Ruft eine  **[Tasks](bc6bb4a5-95a6-9d1f-3e28-92b9548a544a.md)** -Auflistung zurück, die Vorgänge im Projekt darstellt. Read-only **Aufgaben**.


## Syntax

 _Ausdruck_. **Tasks**

 _Ausdruck_ Eine Variable, die ein **Project** -Objekt darstellt.


## Beispiel

Im folgenden Beispiel werden die Namen aller Vorgänge im aktiven Projekt angezeigt.


```
Sub TaskNames() 
 
 Dim T As Task, Names As String 
 
 For Each T In ActiveProject.Tasks 
 Names = Names &amp; T.Name &amp; vbCrLf 
 Next T 
 
 MsgBox Names 
 
End Sub
```

