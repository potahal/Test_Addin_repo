
# Task.LinkPredecessors Method (Project)

Fügt dem Vorgang mindestens einen Vorgänger hinzu.


## Syntax

 _Ausdruck_. **LinkPredecessors**( ** _Tasks_**, ** _Link_**, ** _Lag_** )

 _Ausdruck_ Eine Variable, die ein **Task** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Tasks_|Erforderlich|**Object**|Das angegebene  **Task** oder **Tasks** -Objekt wird zum Vorgänger des durch **Expression** angegebenen Vorgangs.|
| _Link_|Optional|**Long**|Eine Konstante, die die Beziehung zwischen Aufgaben gibt an, die verknüpft werden. Dies kann eine der [PjTaskLinkType](141a1145-0eb5-3664-4755-394584aec8ac.md) -Konstanten sein. Der Standardwert ist **PjFinishToStart**.|
| _Lag_|Optional|**Variant**|Eine Zeichenfolge, die Dauer der Zeitabstand zwischen Vorgängen angibt. Um Zeitabstand zwischen Vorgängen anzugeben, verwenden Sie einen Ausdruck für die  **Verzögerung**, die auf einen negativen Wert ausgewertet wird.|

## Beispiel

Im folgenden Beispiel wird der Benutzer aufgefordert, den Namen eines Vorgangs einzugeben. Dieser Vorgang wird dann als Vorgänger der ausgewählten Vorgänge definiert.


```
Sub LinkTasksFromPredecessor() 
    Dim Entry As String   ' Task name entered by user 
    Dim T As Task         ' Task object used in For Each loop 
    Dim I As Long         ' Used in For loop 
    Dim Exists As Boolean ' Whether or not the task exists 
 
    Entry = InputBox$("Enter the name of a task:") 
 
    Exists = False ' Assume task doesn't exist. 
 
    ' Search active project for the specified task. 
    For Each T In ActiveProject.Tasks 
        If T.Name = Entry Then 
            Exists = True 
            ' Make the task a predecessor of the selected tasks. 
            For I = 1 To ActiveSelection.Tasks.Count 
                ActiveSelection.Tasks(I).LinkPredecessors Tasks:=T 
            Next I 
        End If 
    Next T 
 
    ' If task doesn't exist, display an error and quit the procedure. 
    If Not Exists Then 
        MsgBox ("Task not found.") 
        Exit Sub 
    End If 
End Sub
```

