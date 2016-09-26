
# TaskRequestItem.DownloadState Property (Outlook)

Gibt eine zur  **[OlDownloadState](ff5e00db-ad06-ddf1-6e3a-536c0ae4ef34.md)** -Aufzählung gehörende Konstante zurück, die den Downloadstatus des Elements angibt. Schreibgeschützt.


## Syntax

 _Ausdruck_. **DownloadState**

 _Ausdruck_ Eine Variable, die ein **TaskRequestItem** -Objekt darstellt.


## Beispiel

Im folgenden Beispiel für Microsoft Basic für Applikationen (VBA) wird der  **Posteingang** des Benutzers nach Objekten durchsucht, für die noch kein vollständiger Download durchgeführt wurde. Wenn solche Objekte gefunden werden, wird dem Benutzer eine Nachricht angezeigt und das Objekt für den Download markiert.


```
Sub DownloadItems() 
 
 Dim mpfInbox As Outlook.Folder 
 
 Dim objItems As Outlook.Items 
 
 Dim obj As Object 
 
 Dim i As Integer 
 
 Dim iCount As Integer 
 
 
 
 Set mpfInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox) 
 
 Set objItems = mpfInbox.Items 
 
 iCount = objItems.Count 
 
 'Loop all items in the Inbox folder 
 
 For i = 1 To iCount 
 
 Set obj = objItems.Item(i) 
 
 'Verify if the state of the item is olHeaderOnly 
 
 If obj.DownloadState = olHeaderOnly Then 
 
 MsgBox "This item has not been fully downloaded." 
 
 'Mark the item to be downloaded 
 
 obj.MarkForDownload = olMarkedForDownload 
 
 obj.Save 
 
 End If 
 
 Next 
 
End Sub
```


## Siehe auch


#### Konzepte


[TaskRequestItem-Objekt](2908a28a-634c-e786-aa53-f3e32038b727.md)
#### Weitere Ressourcen


[Elemente des TaskRequestItem-Objekts](http://msdn.microsoft.com/library/d43114ee-be91-ff02-3424-525da2cf3a50%28Office.15%29.aspx)