
# SharedWorkspaceTask.Title-Eigenschaft (Office)

Legt fest oder ruft den Titel der ein  **SharedWorkspaceTask** -Objekt ab. Lese-/Schreibzugriff.


 **Hinweis**  Ab Microsoft Office 2010 ist dieses Objekt oder Element veraltet und sollte nicht verwendet werden.


## Syntax

 _Ausdruck_. **Title**

 _Ausdruck_ Eine Variable, die ein **SharedWorkspaceTask** -Objekt darstellt.


### Rückgabewert

String


## Bemerkungen

Die  **Title** -Eigenschaft ist die einzige erforderliche Eigenschaft einer freigegebenen Arbeitsbereichsaufgabe an. Verwenden Sie die optionale **Description** -Eigenschaft, um zusätzliche Informationen über die Aufgabe bereitzustellen oder zurückzugeben.


## Beispiel

Im folgenden Beispiel wird eine Liste der Titel aller Aufgaben im aktuellen freigegebenen Arbeitsbereich angezeigt.


```
 Dim swsTask As Office.SharedWorkspaceTask 
    Dim strTasks As String 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        strTasks = strTasks &amp; swsTask.Title &amp; vbCrLf 
    Next 
    MsgBox strTasks, vbInformation + vbOKOnly, _ 
        "Tasks in Shared Workspace" 
    Set swsTask = Nothing 
 

```


## Siehe auch


#### Konzepte


[SharedWorkspaceTask-Objekt](fbd82b03-53fa-12ff-9fb2-07bef012dde8.md)
#### Weitere Ressourcen


[Elemente des SharedWorkspaceTask-Objekts](http://msdn.microsoft.com/library/5b5589d1-f907-7357-f930-eede569d2021%28Office.15%29.aspx)