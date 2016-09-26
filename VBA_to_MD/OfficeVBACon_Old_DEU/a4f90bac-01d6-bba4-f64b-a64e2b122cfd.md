
# CustomXMLPart-Objekt (Office)

Stellt ein  **CustomXMLPart** in einer **CustomXMLParts** -Auflistung dar.


## Beispiel

Im folgenden Beispiel wird einem  **CustomXMLPart** -Objekt eine Komponente hinzugefügt.


```
Sub AddPartToCollection() 
    Dim myPart As CustomXMLPart 
 
    Set myPart = ActiveDocument.CustomXMLParts.Add("<author>Mark Twain</author>") 
     
End Sub
```


## Siehe auch


#### Konzepte


[-Objektmodellreferenz](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)
#### Weitere Ressourcen


[Elemente des CustomXMLPart-Objekts](http://msdn.microsoft.com/library/76fe85f4-5a35-7d12-2989-6f17a094dcdf%28Office.15%29.aspx)