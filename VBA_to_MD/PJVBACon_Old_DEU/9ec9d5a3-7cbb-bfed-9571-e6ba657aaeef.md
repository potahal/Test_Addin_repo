
# Application.GanttBarFormatEx Method (Project)

Ändert die Formatierung von Vorgangsbalken gegenüber den Standardarten. Dabei können die Farben als RGB-Hexadezimalwert dargestellt werden.


## Syntax

 _Ausdruck_. **GanttBarFormatEx**( ** _TaskID_**, ** _GanttStyle_**, ** _StartShape_**, ** _StartType_**, ** _StartColor_**, ** _MiddleShape_**, ** _MiddlePattern_**, ** _MiddleColor_**, ** _EndShape_**, ** _EndType_**, ** _EndColor_**, ** _LeftText_**, ** _RightText_**, ** _TopText_**, ** _BottomText_**, ** _InsideText_**, ** _Reset_**, ** _ProjectName_** )

 _Ausdruck_ Ein Ausdruck, der ein **Application** -Objekt zurückgibt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _TaskID_|Optional|**Long**|Die Identifikationsnummer des Vorgangs, der im Gantt-Diagramm geändert werden soll. In der Standardeinstellung werden die Vorgangsbalken der ausgewählten Vorgänge geändert.|
| _GanttStyle_|Optional|**Integer**|Die Formatvorlage, die von Vorgangsbalkens, der formatiert werden soll. Der Wert für GanttStyle basiert auf der Position des Balkens Formatvorlage im Dialogfeld  **Balkenarten**. Beispielsweise gibt der Wert 3 die dritte Balkenart in der Liste.|
| _StartShape_|Optional|**Integer**|Die Form für den Anfang des Vorgangsbalkens. Dies kann eine der  **[PjBarEndShape](0598711b-46ad-1940-103b-12345f32efd8.md)** -Konstanten sein.|
| _StartType_|Optional|**Integer**|Die Art für den Anfang des Vorgangsbalkens. Dies kann eine der  **[PjBarType](abc6a0b2-90bd-48d4-283a-a53618856692.md)** -Konstanten sein.|
| _StartColor_|Optional|**Long**|Die Farbe für den Anfang des Vorgangsbalkens. Dies kann ein RGB-Hexadezimalwert sein, dabei enthält das letzte Byte den Wert für Rot. Z. B. entspricht der Wert &amp;H00FFFF Gelb.|
| _MiddleShape_|Optional|**Integer**|Die Form für die Mitte des Vorgangsbalkens. Dies kann eine der  **[PjBarShape](057356dc-9cab-fbdc-563e-f81cc54a2c33.md)** -Konstanten sein.|
| _MiddlePattern_|Optional|**Integer**|Das Muster für die Mitte des Vorgangsbalkens. Dies kann eine der  **[PjFillPattern](4f6af32c-5efd-42b6-4017-20a1497c1b6d.md)** -Konstanten sein.|
| _MiddleColor_|Optional|**Long**|Die Farbe für den Mittelabschnitt des Vorgangsbalkens. Dies kann ein RGB-Hexadezimalwert sein, dabei enthält das letzte Byte den Wert für Rot. Z. B. entspricht der Wert &amp;H00FFFF Gelb.|
| _EndShape_|Optional|**Integer**|Das Shape Ende des Vorgangsbalkens. Dies kann eine der  **folgenden PjBarEndShape** -Konstanten sein.|
| _EndType_|Optional|**Integer**|Der Typ Ende des Vorgangsbalkens. Kann eine der folgenden  **PjBarType** -Konstanten sein: **PjDashed**, **PjFramed** oder **PjSolid**.|
| _EndColor_|Optional|**Long**|Die Farbe für das Ende des Vorgangsbalkens. Dies kann ein RGB-Hexadezimalwert sein, dabei enthält das letzte Byte den Wert für Rot. Z. B. entspricht der Wert &amp;HFFFF00 Blaugrün.|
| _LeftText_|Optional|**String**|Das auf der linken Seite des Vorgangsbalkens anzuzeigende Vorgangsfeld.|
| _RightText_|Optional|**String**|Das auf der rechten Seite des Vorgangsbalkens anzuzeigende Vorgangsfeld.|
| _TopText_|Optional|**String**|Das oberhalb des Vorgangsbalkens anzuzeigende Vorgangsfeld.|
| _BottomText_|Optional|**String**|Das unterhalb des Vorgangsbalkens anzuzeigende Vorgangsfeld.|
| _InsideText_|Optional|**String**|Das innerhalb des Vorgangsbalkens anzuzeigende Vorgangsfeld.|
| _Reset_|Optional|**Boolean**|**True,** Wenn die Leiste Formatierung zurückgesetzt wird, auf die Standard-Formatierung der Formatvorlage im Dialogfeld **Balkenarten**; anderenfalls  **False**.|
| _ProjectName_|Optional|**String**|Der Name des Projekts, das  **TaskID** enthält. Der Standardwert ist der Name des aktiven Projekts.|

### Rückgabewert

 **Boolean**


## Hinweise

Verwenden die  **GanttBarFormatEx** -Methode ohne Angabe von Argumenten wird das Dialogfeld **Balken formatieren** angezeigt.

Zum Definieren der Standardarten, bei denen die Farben durch RGB-Hexadezimalwerte dargestellt werden können, verwenden Sie die  **[GanttBarEditEx](b574b975-a869-31ba-e525-df8775330b0a.md)** -Methode.


## Beispiel

Im folgenden Beispiel wird für den Anfang des Vorgangs mit der Vorgangsnummer 3 eine mittelgroße rote Raute angezeigt.


```
Sub GanttBar_Format() 
 
    'Activate Gantt Chart view 
    ViewApply Name:="&amp;Gantt Chart" 
    GanttBarFormatEx TaskID:=3, StartShape:=pjDiamond, StartType:=pjSolid, StartColor:=&amp;H8888FF
End Sub
```

