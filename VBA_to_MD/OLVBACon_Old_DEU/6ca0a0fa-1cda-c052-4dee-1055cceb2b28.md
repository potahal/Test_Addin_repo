
# Verwenden von Ereignissen mit Automation

Zum Erstellen eines Ereignishandlers für Microsoft Outlook-Objekte in Microsoft Visual Basic oder Microsoft Visual Basic für Applikationen (VBA) in einer anderen Anwendung müssen Sie die folgenden vier Schritte abschließen:


1. Legen Sie einen Verweis auf die Outlook-Objektbibliothek fest.
    
2. Deklarieren Sie eine Objektvariable, um auf die Ereignisse zu reagieren.
    
3. Verfassen Sie die spezifischen Ereignisprozeduren.
    
4. Initialisieren Sie das deklarierte Objekt.
    

Informieren Sie sich über das [Arbeiten mit Ereignissen in Visual Basic für Applikationen](560bb264-05d0-dbc6-39c2-b95b12f50ed9.md).


## Festlegen eines Verweises auf die Outlook-Objektbibliothek

Bevor Sie ein Outlook-Objekt in Code für Visual Basic oder Visual Basic für Applikationen verwenden können, müssen Sie zuerst im Dialogfeld  **Verweise** einen Verweis auf das Outlook-Objektmodell festlegen. Weitere Informationen über die Verwendung dieses Dialogfelds erhalten Sie in der Onlinehilfe Ihrer Programmierumgebung.


## Deklarieren der Objektvariable

Nachdem Sie einen Verweis auf die Objektmodell-Bibliothek festgelegt haben, müssen Sie die Variablen deklarieren, die auf die zu verwendenden Objekte verweisen. Sie können die Variable in dem Modul deklarieren, in dem das Objekt verwendet wird (d. h. in dem Modul, das die Ereignishandler-Prozedur enthält), im Allgemeinen deklarieren Sie es jedoch in einem Klassenmodul, sodass es von jedem Modul in Ihrem Programm verwendet werden kann.

Um zum Beispiel eine Objektvariable für das  **[Application](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)** -Objekt in einem Klassenmodul zu deklarieren, verwenden Sie Code wie den folgenden.




```
Public WithEvents myOlApp As Outlook.Application
```

Sie müssen mit dem  `WithEvents`-Schlüsselwort festzulegen, dass die Objektvariable verwendet wird, um auf Ereignisse zu reagieren, die vom Objekt ausgelöst werden.


## Schreiben der Ereignisprozedur

Nachdem ein neues Objekt mit Ereignissen deklariert worden ist, wird es in der  **Objekt**-Liste im Codefenster des Klassenmoduls angezeigt, und Sie können die Ereignisprozedur des Objekts aus der Liste  **Prozeduren/Ereignisse** auswählen. Wenn Sie z. B. das **[ItemSend](54f506ea-87a2-29b9-2b33-67bc87167933.md)** -Ereignis für ein als `myOlApp` deklariertes **Application** -Objekt auswählen, wird im Codefenster die folgende leere Prozedur angezeigt.


```
Private Sub myOlApp_ItemSend(Item as Object, Cancel as Boolean) 
 
End Sub
```


## Initialisieren des deklarierten Objekts

Bevor die Prozedur ausgeführt wird, müssen Sie das deklarierte Objekt (in diesem Beispiel  `myOlApp`) mit dem  **Application** -Objekt verbinden. Wenn Sie das Objekt in einem Klassenmodul mit dem Namen `EventClassModule` deklariert haben, dann können Sie den folgenden Code in jedem Modul verwenden.


```
Dim myClass as New EventClassModule  
Sub Register_Event_Handler()  
    Set myClass.myOlApp = "Outlook.Application"  
End Sub
```

Wenn die




```
Register_Event_Handler
```

-Prozedur ausgeführt wird, zeigt das  `myOlApp`-Objekt im Formular oder Klassenmodul auf das Outlook- **Application** -Objekt, und die Ereignisprozedur wird ausgeführt, wenn das Ereignis auftritt.

