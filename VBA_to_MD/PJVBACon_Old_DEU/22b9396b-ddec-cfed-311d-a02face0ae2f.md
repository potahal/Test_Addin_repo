
# Application.SelectResourceColumn Method (Project)

Markiert eine Spalte, die Ressourceninformationen enthält.


## Syntax

 _Ausdruck_. **SelectResourceColumn**( ** _Column_**, ** _Additional_**, ** _Extend_**, ** _Add_** )

 _Ausdruck_ Eine Variable, die ein **Application** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Column_|Optional|**String**|Der Feldname der zu markierenden Spalte. Der Standardwert ist die Spalte, die die aktive Zelle enthält.|
| _Additional_|Optional|**Integer**|Die Anzahl der zusätzlichen Spalten rechts neben der  **Spalte** auswählen. Wenn **Extend** **True** ist, wird **Additional** ignoriert. Der Standardwert ist 0.|
| _Extend_|Optional|**Boolean**|**True,** Wenn alle Spalten zwischen der aktuellen Markierung und **Column** ausgewählt sind. Der Standardwert ist **False**.|
| _Add_|Optional|**Boolean**|**True,** Wenn die aktuelle Spalte in der Auswahl enthalten ist. Der Standardwert ist **False**.|

### Rückgabewert

 **Boolean**


## Bemerkungen

Die  **SelectResourceColumn** -Methode ist nur verfügbar, wenn die Ansicht Ressource: Tabelle oder Ressource: Einsatz die aktive Ansicht ist.


## Beispiel

Im folgenden Beispiel werden die Spalte  **Indicators** und die nächsten zwei Spalten markiert.


```
Sub Select_ResourceColumn() 
 
 'Activate Resource Sheet 
 ViewApply Name:="&amp;Resource Sheet" 
 SelectResourceColumn Column:="Indicators", Additional:=2 
End Sub
```

