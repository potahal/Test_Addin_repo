
# TaskRequestItem.CustomPropertyChange Event (Outlook)

Tritt auf, wenn eine benutzerdefinierte Aktion eines Elements (bei dem es sich um eine Instanz des übergeordneten Objekts handelt) geändert wird.


## Syntax

 _Ausdruck_. **CustomPropertyChange**( ** _Name_** )

 _Ausdruck_ Eine Variable, die ein **TaskRequestItem** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Name_|Erforderlich|**String**|Der Name der benutzerdefinierten Eigenschaft, die geändert wurde.|

## Bemerkungen

Der Name der Eigenschaft wird an die Prozedur übergeben, damit Sie ermitteln können, welche benutzerdefinierte Eigenschaft geändert wurde.


## Beispiel

In diesem Beispiel für Microsoft Visual Basic Scripting Edition (VBScript) wird das  **CustomPropertyChange** -Ereignis verwendet, um ein Steuerelement zu aktivieren, wenn ein Feld vom Typ Boolean auf **True** festgelegt ist.

In diesem Beispiel erstellen Sie zwei benutzerdefinierte Felder auf der zweiten Seite eines Formulars. Das erste ein Feld  **vom Typ Boolean** ist "RespondBy". Das zweite Feld heißt "DateToRespond".




```
Sub Item_CustomPropertyChange(ByVal myPropName) 
 
 Select Case myPropName 
 
 Case "RespondBy" 
 
 Set myPages = Item.GetInspector.ModifiedFormPages 
 
 Set myCtrl = myPages("P.2").Controls("DateToRespond") 
 
 If Item.UserProperties("RespondBy").Value Then 
 
 myCtrl.Enabled = True 
 
 myCtrl.Backcolor = 65535 'Yellow 
 
 Else 
 
 myCtrl.Enabled = False 
 
 myCtrl.Backcolor = 0 'Black 
 
 End If 
 
 Case Else 
 
 End Select 
 
End Sub
```


## Siehe auch


#### Konzepte


[TaskRequestItem-Objekt](2908a28a-634c-e786-aa53-f3e32038b727.md)
#### Weitere Ressourcen


[Elemente des TaskRequestItem-Objekts](http://msdn.microsoft.com/library/d43114ee-be91-ff02-3424-525da2cf3a50%28Office.15%29.aspx)