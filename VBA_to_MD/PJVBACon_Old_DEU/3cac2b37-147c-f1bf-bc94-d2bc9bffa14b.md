
# Application.GanttBarStyleDelete Method (Project)

Löscht eine Vorgangsbalkenart aus dem aktiven Balkendiagramm.


## Syntax

 _Ausdruck_. **GanttBarStyleDelete**( ** _Item_** )

 _Ausdruck_ Eine Variable, die ein **Application** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Item_|Erforderlich|**String**|**Zeichenfolge**. Der Name oder die Zeilennummer Anzahl von Vorgangsbalkens, der aus dem Dialogfeld **Balkenarten** gelöscht.|

### Rückgabewert

 **Boolean**


## Bemerkungen

Zum manuellen Anzeigen des Dialogfelds  **Balkenarten** klicken Sie auf die Registerkarte **Format** unter der Registerkarte **Gantt-Diagrammtools**. Klicken Sie in der Gruppe  **Balkenarten** auf **Balkenarten** in der Dropdownliste **Format**. Das Dialogfeld  **Balkenarten** kann bis zu 200 Arten enthalten.


## Beispiel

Mit den folgenden Befehl wird die Balkenartnummer 41 im Dialogfeld  **Balkenarten** gelöscht.


```
GanttBarStyleDelete Item:="41"
```

