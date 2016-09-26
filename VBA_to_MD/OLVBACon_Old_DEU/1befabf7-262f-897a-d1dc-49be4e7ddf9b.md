
# TaskItem.Respond Method (Outlook)

Antwortet auf eine Aufgabenanfrage.


## Syntax

 _Ausdruck_. **Respond**( ** _Response_**, ** _fNoUI_**, ** _fAdditionalTextDialog_** )

 _Ausdruck_ Eine Variable, die ein **TaskItem** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Response_|Erforderlich|**[OlTaskResponse](7616cbdc-fc9c-abbe-fd07-ebdadc13ede2.md)**|Die Antwort auf die Anfrage.|
| _fNoUI_|Erforderlich|**Variant**|**True**, wenn ein Dialogfeld nicht angezeigt; die Antwort wird automatisch gesendet. **False**, wenn das Dialogfeld für die Reaktion anzuzeigen.|
| _fAdditionalTextDialog_|Erforderlich|**Variant**|**False**, wenn keine Aufforderung den Benutzer zur Eingabe; die Antwort wird zur Bearbeitung im Inspektor angezeigt. **True**, wenn fordert den Benutzer auf Senden oder mit Kommentaren senden. Dieses Argument ist nur gültig, wenn _fNoUI_ auf **false festgelegt** ist.|

### Rückgabewert

Ein  **[TaskItem](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)** -Objekt, das die Antwort auf die Aufgabenanfrage darstellt.


## Bemerkungen

Beim Aufruf der Methode  **reagieren**, mit dem Parameter **als OlTaskAccept** erstellt Outlook eine neue **TaskItem**, die die aufgabenanfrageelement dupliziert. Das neue Element hat eine andere Eintrags-ID Outlook entfernt das Originalelement.

In der folgenden Tabelle wird das Verhalten der  **reagieren** -Methode abhängig von der übergeordnete Objekt und die Parameter _fNoUI_ und _fAdditionalTextDialog_ beschrieben.



|** _fNoUI, fAdditionalTextDialog_**|** _Ergebnis_**|
|:-----|:-----|
|**True, True**|Antwortelement wird ohne Benutzeroberfläche zurückgegeben. Um die Antwort zu senden, müssen Sie die Methode  **[Send](54f751fc-cff1-5d17-f635-f688cd8ad6f8.md)** aufrufen.|
|**True, False**|Gleiches Ergebnis wie bei  **True, True**.|
|**False, True**|Wenn die  **[Display](fea0619d-06dc-df44-fe93-5756eefb1be0.md)** -Methode aufgerufen wurde, wird die Aufforderung des Benutzers angezeigt. Andernfalls wird das Element ohne Bestätigung gesendet und das resultierende Element wird Nothing zurück.|
|**False, False**|Keine Aktion.|

## Siehe auch


#### Konzepte


[TaskItem-Objekt](5df8cfa5-5460-a5a1-a130-ba5bca1a0091.md)
#### Weitere Ressourcen


[Elemente des TaskItem-Objekts](http://msdn.microsoft.com/library/97234a76-2fc5-bbe4-2e14-25ae18694fc9%28Office.15%29.aspx)