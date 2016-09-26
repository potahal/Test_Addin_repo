
# CardView.Filter Property (Outlook)

Gibt zurück oder legt einen  **String** -Wert, der den Filter für eine Ansicht darstellt. Lese-/Schreibzugriff.


## Syntax

 _Ausdruck_. **Filter**

 _Ausdruck_ Eine Variable, die ein **CardView** -Objekt darstellt.


## Hinweise

Der Wert dieser Eigenschaft ist eine Zeichenfolge in der DASL-Syntax, die den aktuellen Filter für die Ansicht darstellt. Weitere Informationen zur Verwendung der DASL-Syntax zum Filtern in einer Ansicht finden Sie unter [Filtern von Elementen](4038e042-1b07-5d18-18b0-c2b58c9c42da.md).


## Beispiel

Im folgende Visual Basic für Applikationen (VBA) wird ein  **[View](41c8d149-9912-1685-4c8b-3c849cc6f1ed.md)** -Objekt abgerufen, mithilfe der **[CurrentView](177e6387-9ccb-cb71-bbe5-332c25485848.md)** -Eigenschaft des **[Explorer](026591e5-049f-503a-4166-34e6dbc225fb.md)** -Objekt ab, und legt die **[Filter](9a4b4b27-d543-df82-3058-e0a6ad2f51a1.md)** -Eigenschaft des **View** -Objekt wird an nur die Outlook-Elemente anzuzeigen, die letzte Woche empfangen wurden.


```
Private Sub FilterViewToLastWeek() 
 
 Dim objView As View 
 
 
 
 ' Obtain a View object reference to the current view. 
 
 Set objView = Application.ActiveExplorer.CurrentView 
 
 
 
 ' Set a DASL filter string, using a DASL macro, to show 
 
 ' only those items that were received last week. 
 
 objView.Filter = "%lastweek(""urn:schemas:httpmail:datereceived"")%" 
 
 
 
 ' Save and apply the view. 
 
 objView.Save 
 
 objView.Apply 
 
End Sub 
 

```


## Siehe auch


#### Konzepte


[CardView-Objekt](cdac229b-f2b6-9ecb-e1a7-b53509426570.md)
#### Weitere Ressourcen


[Elemente des CardView-Objekts](http://msdn.microsoft.com/library/8b9eda10-1ece-c961-e432-3fca6dfb4f07%28Office.15%29.aspx)