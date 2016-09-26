
# Verweisen auf Felder

Beim Zugriff auf Felder in einem Element hängt die Wahl der zu verwendenden Methode davon ab, ob das Feld ein Standardfeld, ein integriertes Outlook-Feld oder ein benutzerdefiniertes Feld ist.

Sie greifen jedoch in keinem Fall direkt auf das Feld zu. Stattdessen verweisen Sie auf das Feld als eine Eigenschaft des Elements, mit dem Sie arbeiten.

Um beispielsweise den Text vom Feld Betreff einer E-Mail-Nachricht zu erhalten, verwenden Sie die Subject **Subject**-Eigenschaft des Steuerelements, wie im folgenden Beispiel für VBScript gezeigt.




```
mySubject = Item.Subject
```

Auf ein benutzerdefiniertes Feld greifen Sie mithilfe der UserProperties **UserProperties**-Eigenschaft des Steuerelements zu, wie im folgenden Beispiel für VBScript gezeigt. Dieses Beispiel setzt voraus, dass das Element bereits ein benutzerdefiniertes Feld mit dem Namen  **ReferredBy** enthält.



```
MyReferral = Item.UserProperties("ReferredBy")
```

