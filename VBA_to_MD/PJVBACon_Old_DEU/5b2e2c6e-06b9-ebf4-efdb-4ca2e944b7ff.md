
# StartDriver.EffectiveDateAdd Property (Project)

Ruft Datum und Uhrzeit, die durch eine angegebene Dauer ein anderes Datum folgt mithilfe effektiven Kalender für einen manuell geplanten Vorgang. Read-only  **Variant**.


## Syntax

 _Ausdruck_. **EffectiveDateAdd**( ** _Date_**, ** _Duration_** )

 _Ausdruck_ Ein Ausdruck, der ein **StartDriver** -Objekt zurückgibt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Date_|Erforderlich|**Variant**|Beliebiges Datum und beliebige Uhrzeit, z. B. "10.07.2010" oder "10.07.2010 14:00:00".|
| _Duration_|Erforderlich|**Variant**|Hinzuzufügender zeitlicher Abstand, z. B. "3d" (3 Tage) oder "2w" (2 Wochen).|

## Hinweise

Die  **EffectiveDateAdd** -Eigenschaft verwendet effektiven Kalender für manuell geplante Vorgänge, wodurch Aufgaben beginnen oder enden auf arbeitsfreie Zeiten. Die Eigenschaft und die Argumente haben keine Auswirkung auf tatsächlichen Vorgangstermine.

Die Eigenschaften  **[EffectiveDateSubtract](14529bd1-9029-d1bc-60a0-b7863cba4d6d.md)**, **EffectiveDateAdd** und **[EffectiveDateDifference](9b825839-31de-71f8-9804-015dfd5a293c.md)** können Anfangs-und Endtermine für manuell geplante Vorgänge.

Verwenden Sie zur Berechnung von Terminen für automatisch geplante Vorgänge, bei denen Sie auch den Kalender angeben können, die  **[DateAdd](df0da054-495c-c224-ebc8-b47acb78e2af.md)** -Methode.


## Beispiel

Mit der folgenden Anweisung wird der Wert "7/9/2009 5:00:00 PM", zurückgegeben, ein Termin mit sechs Tagen Abstand vom angegebenen Datum.


```
Debug.Print ActiveProject.Tasks(3).StartDriver.EffectiveDateAdd("7/2/2009", "6d")
```

