
# SharedWorkspace.Tasks-Eigenschaft (Office)

Ruft eine  **[SharedWorkspaceTasks](de26341f-44d1-131e-1dbe-e31f3f68e312.md)** -Auflistung ab, die die Aufgabenliste des aktuellen freigegebenen Arbeitsbereichs darstellt. Schreibgeschützt.


 **Hinweis**  Ab Microsoft Office 2010 ist dieses Objekt oder Element veraltet und sollte nicht verwendet werden.


## Syntax

 _Ausdruck_. **Tasks**

 _Ausdruck_ Eine Variable, die ein **SharedWorkspace** -Objekt darstellt.


## Beispiel

Das folgende Beispiel listet die Aufgaben des aktuellen freigegebenen Arbeitsbereichs auf.


```
   Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " &amp; swsTasks.Count &amp; _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```


## Siehe auch


#### Konzepte


[SharedWorkspace-Objekts](7512f0ff-382d-d344-9424-aa10549d14f9.md)
#### Weitere Ressourcen


[Elemente des SharedWorkspace-Objekts](http://msdn.microsoft.com/library/e4c2b518-d955-27e1-3e73-173d3c4f961d%28Office.15%29.aspx)