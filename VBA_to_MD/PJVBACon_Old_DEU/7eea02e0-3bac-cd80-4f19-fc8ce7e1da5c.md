
# PjBoxSet Enumeration (Project)

Enthält Konstanten, mit denen die Erstellung, die Auswahl oder das Verschieben eines Vorgangs in der Ansicht  **Netzplandiagramm** angegeben wird.



|**Name**|**Wert**|**Beschreibung**|
|:-----|:-----|:-----|
|**pjBoxAddToSelection**|0|Wählt den Vorgang aus und behält ggf. die bereits vorhandene Auswahl bei.|
|**pjBoxCreate**|1|Erstellt einen neuen Vorgang und hebt ggf. die bereits vorhandene Auswahl auf.|
|**pjBoxMoveAbsolute**|2|Positioniert den Vorgang relativ zur linken oberen Ecke der Ansicht. Wenn mehrere Vorgänge ausgewählt sind und TaskID nicht angegeben wird, werden alle ausgewählten Vorgänge verschoben. Wenn TaskID angegeben wird, wird die Auswahl aufgehoben und nur dieser Vorgang verschoben.|
|**pjBoxMoveRelative**|3|Positioniert den Vorgang relativ zur aktuellen Position. Wenn mehrere Vorgänge ausgewählt sind und TaskID nicht angegeben wird, werden alle ausgewählten Vorgänge verschoben. Wenn TaskID angegeben wird, wird die Auswahl aufgehoben und nur dieser Vorgang verschoben.|
|**pjBoxSelect**|4|Wählt den Vorgang aus und hebt ggf. die bereits vorhandene Auswahl auf.|
|**pjBoxUnselect**|5|Entfernt den Vorgang aus der Auswahl. Wenn mehrere Vorgänge ausgewählt sind und TaskID nicht angegeben wird, wird der Knoten mit dem Fokus aus der Auswahl entfernt. Wenn TaskID angegeben wird, wird nur dieser Vorgang aus der Auswahl entfernt.|
