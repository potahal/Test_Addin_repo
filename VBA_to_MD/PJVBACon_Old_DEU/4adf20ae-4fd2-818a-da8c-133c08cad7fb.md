
# Application.SelectBeginning Method (Project)

Markiert die erste Zelle der aktiven Tabelle oder Ansicht.


## Syntax

 _Ausdruck_. **SelectBeginning**( ** _Extend_** )

 _Ausdruck_ Eine Variable, die ein **Application** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Extend_|Optional|**Boolean**|**True,** Wenn die aktuelle Auswahl auf die erste Zelle erweitert wird. Wenn die aktive Ansicht der Netzplandiagramm oder Ressource: Grafik ist, wird die Extend ignoriert. Der Standardwert ist **False**.|

### Rückgabewert

 **Boolean**


## Bemerkungen

In der Ressource: Grafik wird mit  **SelectBeginning** die Ressource mit der niedrigsten Identifikationsnummer ausgewählt. In der Netzplandiagramm aktiviert **SelectBeginning** das am nächsten an der linken oberen Ecke der Ansicht.


## Beispiel

Im folgenden Beispiel wird das Feld  **Name** in Zeile 4 als Anfangsfeld im Balkendiagramm markiert.


```
Sub Select_Beginning() 
 
 ViewApply Name:="&amp;Gantt Chart" 
 SelectTaskField Row:=4, Column:="Name", RowRelative:=False 
 
 SelectBeginning Extend:=True 
End Sub
```

