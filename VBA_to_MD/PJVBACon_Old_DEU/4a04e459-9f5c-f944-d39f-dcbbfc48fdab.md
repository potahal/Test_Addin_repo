
# Application.UpdateTasks Method (Project)

Aktualisiert die ausgewählten Vorgänge.


## Syntax

 _Ausdruck_. **UpdateTasks**( ** _PercentComplete_**, ** _ActualDuration_**, ** _RemainingDuration_**, ** _ActualStart_**, ** _ActualFinish_**, ** _Notes_** )

 _Ausdruck_ Eine Variable, die ein **Application** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _PercentComplete_|Optional|**Variant**|Der abgeschlossene Prozentsatz der aktiven Vorgänge.|
| _ActualDuration_|Optional|**Variant**|Die aktuelle Dauer der ausgewählten Vorgänge.|
| _RemainingDuration_|Optional|**Variant**|Die verbleibende Dauer der ausgewählten Vorgänge.|
| _ActualStart_|Optional|**Variant**|Der aktuelle Anfangstermin der ausgewählten Vorgänge.|
| _ActualFinish_|Optional|**Variant**|Der aktuelle Endtermin der ausgewählten Vorgänge.|
| _Notes_|Optional|**String**|Anmerkungen im Feld  **Notizen** für die ausgewählten Vorgänge. Der Wert kann nur im reinen Textformat und nicht im RTF-Format, wie im Dialogfeld **Notizen**, angegeben werden.|

### Rückgabewert

 **Boolean**


## Bemerkungen

Verwendung der  **UpdateTasks** -Methode ohne Angabe von Argumenten wird das Dialogfeld **Vorgänge aktualisieren** angezeigt.


## Beispiel

Im folgenden Beispiel wird ein neuer Vorgang namens "TestTask-1" erstellt und so aktualisiert, dass er zu 50% abgeschlossen ist. Anschließend wird der Vorgang gelöscht.


```
Sub Update_Tasks() 
 
 'Activate Gantt Chart 
 ViewApply Name:="Gantt Chart" 
 
 'Create a task 
 RowInsert 
 SetTaskField Field:="Name", Value:="TestTask-1" 
 SetTaskField Field:="Duration", Value:="2" 
 
 'Update the percent complete of the new task. 
 UpdateTasks PercentComplete:="50" 
 
 'Delete the new task 
 ActiveProject.Tasks("TestTask-1").Delete 
End Sub
```

