
# Categories.Item Method (Outlook)

Gibt ein  **[Category](143ef095-54b0-cbe2-e356-632029061ac2.md)** -Objekt aus der Auflistung zurück.


## Syntax

 _Ausdruck_. **Item**( ** _Index_** )

 _Ausdruck_ Eine Variable, die ein **Categories** -Objekt darstellt.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Index_|Erforderlich|**Variant**|Entweder ein  **Long** -Wert, der die Indexnummer des Objekts darstellt, oder ein **String** -Wert, der **[den](b9a711e9-f79d-f4f7-88bb-eaeb61d64089.md)** Name- oder **[CategoryID](e75ed17a-940f-2325-8739-1367329854d2.md)** -Eigenschaftswert eines Objekts in der Auflistung darstellt.|

### Return Value

Ein  **Category** -Objekt, das das angegebene Objekt darstellt.


## Hinweise

Wenn der Name einer Kategorie in  _Index_angegeben ist, gibt diese Methode das erste  **Category** -Objekt, das mit den angegebenen Wert übereinstimmt. Wenn keine Übereinstimmung gefunden werden kann, gibt die Methode **Null** ( **Nothing** in Visual Basic.)


## Siehe auch


#### Konzepte


[Categories-Objekt](319efa26-269d-9f2f-c8ec-33082e80a9e2.md)
#### Weitere Ressourcen


[Elemente des Categories-Objekts](http://msdn.microsoft.com/library/36fd8906-69fa-5aa8-b026-a2de208ccd56%28Office.15%29.aspx)