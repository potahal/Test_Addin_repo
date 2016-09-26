
# RemoteItem.MarkForDownload Property (Outlook)

Zurückgeben oder festlegen eine  **[OlRemoteStatus](2df0404c-26c9-87d4-6916-d75aff8e3fbc.md)** -Konstante, die den Status eines Elements bestimmt, nachdem es von einem Remotebenutzer empfangen wurde. Lese-/Schreibzugriff.


## Syntax

 _Ausdruck_. **MarkForDownload**

 _Ausdruck_ Eine Variable, die ein **RemoteItem** -Objekt darstellt.


## Bemerkungen

Diese Eigenschaft verleiht Remotebenutzern mit nicht idealen Fähigkeiten zur Datenübertragung eine erhöhte Messagingflexibilität.


## Beispiel

Im folgenden Beispiel wird der Posteingang des Benutzers nach Objekten durchsucht, für die noch kein vollständiger Download durchgeführt wurde. Wenn solche Objekte gefunden werden, wird dem Benutzer eine Nachricht angezeigt und das Objekt für den Download markiert.


```
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To mpfInbox.Items.Count 
 
 Set obj = mpfInbox.Items.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox ("This item has not been fully downloaded.") 
 
 'Mark the item to be downloaded. 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 End If 
 
 Next 
 
End Sub
```


## Siehe auch


#### Konzepte


[RemoteItem-Objekt](6302aaff-cdcf-4d86-60f1-4bed15540d9f.md)
#### Weitere Ressourcen


[Elemente des RemoteItem-Objekts](http://msdn.microsoft.com/library/15c0872e-88cc-9b9b-c31e-c15d6971e6e0%28Office.15%29.aspx)