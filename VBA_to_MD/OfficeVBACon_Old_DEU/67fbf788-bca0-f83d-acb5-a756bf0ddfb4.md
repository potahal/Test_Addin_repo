
# SharedWorkspace.CreateNew-Methode (Office)

Erstellt einen Dokumentarbeitsbereich-Website auf dem Server und fügt das aktive Dokument der neuen freigegebenen Arbeitsbereichwebsite hinzu.


 **Hinweis**  Ab Microsoft Office 2010 ist dieses Objekt oder Element veraltet und sollte nicht verwendet werden.


## Syntax

 _Ausdruck_. **CreateNew**( ** _URL_**, ** _Name_** )

 _Ausdruck_ Eine Variable, die ein **SharedWorkspace** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _URL_|Optional|**Variant**|Die URL für den übergeordneten Ordner, in dem der neue freigegebene Arbeitsbereich erstellt werden soll. Wenn keine URL angegeben wird, wird die Website im Standardserverpfad des Benutzers erstellt.|
| _Name_|Optional|**Variant**|Der Name der neuen freigegebenen Arbeitsbereichwebsite. Standardmäßig ist dies der Name des aktiven Dokuments ohne dessen Dateinamenerweiterung. Wenn Sie z. B. eine Arbeitsbereichwebsite für die Datei  **Budget.xls** erstellen, ist der Name der neuen Arbeitsbereichwebsite **Budget**.|

## Bemerkungen

Verwenden Sie die  **CreateNew** -Methode, um einer freigegebenen Arbeitsbereichwebsite für das aktive Dokument zu erstellen. Ausgelassen werden Sie, die beiden optionalen Argumente zum Erstellen der Website mit dem Namen des aktiven Dokuments im Server-Standardspeicherort des Benutzers.

Die  **CreateNew** -Methode erzeugt einen Fehler, wenn das aktive Dokument geändert hat, die nicht gespeichert wurden. Das Dokument muss gespeichert werden, bevor Sie es einer freigegebenen Arbeitsbereichwebsite hinzufügen können.


 **Hinweis**  Unmittelbar nach einer freigegebenen Dokumentarbeitsbereich-Website erstellen, und klicken Sie dann im aktive Dokument auf der Website erstellen wird das aktive Dokument geschlossen und dann erneut geöffnet, damit die Kopie des aktiven Dokuments, die der Benutzer erhält befindet sich auf der Website ist. Wenn das aktive Dokument vor dem Aufrufen der  **CreateNew** -Methode gespeichert wurde, steht diese Kopie des Dokuments für den Zeitraum während die neue Kopie erstellt wird. Daraufhin wird eine Ausnahme für Code, der versucht, auf die gespeicherte Kopie während der Erstellung Zeitraum zuzugreifen. Eine umgangen werden, die eine kurze Verzögerung (vorgeschlagenen 15 Sekunden oder mehr) zugrunde liegen, bevor Sie versuchen, das aktive Dokument von einem Skript zugreifen. Darüber hinaus werden ein zwischengespeichertes Objekt, das auf das lokale Dokument verweist zeigt so verweisen Sie auf das Dokument in der freigegebenen Arbeitsbereichwebsite aktualisiert.


## Beispiel

Im folgenden Beispiel wird eine freigegebenen Dokumentarbeitsbereich-Website erstellt, im URL http://Server/Sites/MySite/, nennt den Arbeitsbereich "My Shared Budget Document" und der Website im aktive Dokument hinzugefügt. Die  **URL** -Eigenschaft der neuen freigegebenen Arbeitsbereichwebsite gibt http://server/sites/mysite/My%20Shared%20Budget%20Document/, gibt die **Name** -Eigenschaft "Mein Shared Budget Document und zeigt die **Count** -Eigenschaft der **SharedWorkspaceFiles** -Auflistung eine einzelne Datei.


```
   Dim sws As Office.SharedWorkspace 
    Dim strSWSInfo As String 
    Set sws = ActiveWorkbook.SharedWorkspace 
    sws.CreateNew "http://server/sites/mysite/", "My Shared Budget Document" 
    strSWSInfo = "Name: " &amp; sws.Name &amp; vbCrLf &amp; _ 
        "URL: " &amp; sws.URL &amp; vbCrLf &amp; _ 
        "File(s): " &amp; sws.Files.Count 
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _ 
        "New Shared Workspace Information" 
    Set sws = Nothing 

```


## Siehe auch


#### Konzepte


[SharedWorkspace-Objekts](7512f0ff-382d-d344-9424-aa10549d14f9.md)
#### Weitere Ressourcen


[Elemente des SharedWorkspace-Objekts](http://msdn.microsoft.com/library/e4c2b518-d955-27e1-3e73-173d3c4f961d%28Office.15%29.aspx)