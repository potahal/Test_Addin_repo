
# DocumentProperty.LinkToContent-Eigenschaft (Office)

Ist  **True**, wenn der Wert der benutzerdefinierten Dokumenteigenschaft mit dem Inhalt des Containerdokuments verknüpft ist. **False,** Wenn der Wert statische ist. Lese-/Schreibzugriff.


## Syntax

 _Ausdruck_. **LinkToContent**( ** _pfLinkRetVal_** )

 _Ausdruck_ Eine Variable, die ein **DocumentProperty** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _pfLinkRetVal_|Erforderlich|**Boolean**|Gibt an, ob die Dokumenteigenschaft mit dem Containerdokument verknüpft ist.|

## Bemerkungen

Diese Eigenschaft gilt nur für benutzerdefinierte Dokumenteigenschaften. Für integrierte Dokumenteigenschaften ist der Wert dieser Eigenschaft  **False**.

Verwenden Sie die  **LinkSource** -Eigenschaft, um die Quelle für die angegebene verknüpfte Eigenschaft festzulegen. Die **LinkSource** -Eigenschaft festlegen, wird die **LinkToContent** -Eigenschaft auf **True** festgelegt.

Für Excel Wenn LinkToContent auf  **True** festgelegt ist, müssen Sie einen Adresse oder einen Bereich Namen für die[LinkSource](http://msdn.microsoft.com/library/3e3a6ebc-615a-298e-c40f-cbb6d5cf63e3.md %28Office.15%29.aspx) aus der Arbeitsmappe angeben. Wenn der Name-Adresse oder der Bereich mehr als eine Zelle abdeckt, nimmt die benutzerdefinierten Document-Eigenschaft den Wert von der linken oberen Zelle des Bereichs.


## Beispiel

In diesem Beispiel wird der Verknüpfungsstatus der benutzerdefinierten Dokumenteigenschaft angezeigt. Für das Beispiel funktioniert muss  **dp** ein gültiges **DocumentProperty** -Objekt sein.


```
Sub DisplayLinkStatus(dp As DocumentProperty) 
 Dim stat As String, tf As String 
 If dp.LinkToContent Then 
 tf = "" 
 Else 
 tf = "not " 
 End If 
 stat = "This property is " &amp; tf &amp; "linked" 
 If dp.LinkToContent Then 
 stat = stat + Chr(13) &amp; "The link source is " &amp; dp.LinkSource 
 End If 
 MsgBox stat 
End Sub
```


## Siehe auch


#### Konzepte


[DocumentProperty-Objekt](dd54ca3c-e0e2-4816-539a-17c5b4a928b1.md)
[Sync-Objekt](1cb049a0-a803-969a-7923-15ddb8da8f3b.md)
#### Weitere Ressourcen


[Elemente des DocumentProperty-Objekts](http://msdn.microsoft.com/library/568da0ff-fa90-150a-06ec-611de886334e%28Office.15%29.aspx)
[Elemente des Sync-Objekts](http://msdn.microsoft.com/library/748726bd-83de-425a-5af8-177c34e3a013%28Office.15%29.aspx)