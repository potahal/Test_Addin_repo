
# TextRange2.InsertChartField-Methode (Office)

Fügt ein Feld in den Textkörper einer Bezeichnung Daten in einem Diagramm.

Diese Methode gilt nur für datenbeschriftungen in einem Diagramm. Durch Aufrufen dieser Methode auf einer beliebigen anderen [TextRange2](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md) -Objekts löst einen Laufzeitfehler zurück.

## Syntax

 _Ausdruck_. **InsertChartField** _(ChartFieldType,_ _Formula,_ _Position)_

 _Ausdruck_ Eine Variable, die ein **TextRange2** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _ChartFieldType_|Erforderlich|[MsoChartFieldType](ce6b367d-d09f-4345-33e3-f181b1a9a41d.md)|Gibt den Typ des Diagramms Felds zum Einfügen in eine Beschriftung zugewiesen.|
| _Formula_|Optional|**string**|Gibt an, eine Zelle (oder einen Bereich), wenn die Konstante  **MsoChartFieldFormula** für den _ChartFieldType_ -Parameter übergeben wird.|
| _Position_|Optional|**integer**|Gibt die Position des Zeichens, an das Diagramm-Feld eingefügt wird. Standardmäßig wird das Feld an das Ende des Texts anfügen. Wenn der Positionswert außerhalb des Bereichs ist, wird der Standardwert verwendet.|
| _ChartFieldType_|Erforderlich|MSOCHARTFIELDTYPE||
| _Formula_|Optional|STRING||
| _Position_|Optional|INT||
|Name|Erforderlich/Optional|Datentyp|Beschreibung|

### Rückgabewert

[TextRange2](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md)

