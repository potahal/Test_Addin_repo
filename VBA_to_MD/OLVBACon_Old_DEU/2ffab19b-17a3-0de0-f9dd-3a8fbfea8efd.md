
# DistListItem.GetInspector Property (Outlook)

Gibt ein  **[Inspector](d7384756-669c-0549-1032-c3b864187994.md)** -Objekt, das einen Inspektor initialisiert wird, um das angegebene Element enthalten darstellt. Schreibgeschützt.


## Syntax

 _Ausdruck_. **GetInspector**

 _Ausdruck_ Eine Variable, die ein **DistListItem** -Objekt darstellt.


## Bemerkungen

Diese Eigenschaft ist nützlich, um die in der Anzeige des Elements, im Gegensatz zur Verwendung der  **[Application.ActiveInspector](3f2b6491-7b4b-8165-327e-b319711d5656.md)** -Methode und Festlegen der **[Inspector.CurrentItem](eaaf0192-a169-c107-95a6-b8e759a3b873.md)** -Eigenschaft ein **Inspector** -Objekt zurückzugeben. Wenn bereits ein **Inspector** -Objekt für das Element vorhanden ist, gibt die **GetInspector** -Eigenschaft, anstatt einen neuen Anwendungspool erstellen **Inspector** -Objekt zurück.


## Siehe auch


#### Konzepte


[DistListItem-Objekt](027c3986-abff-d9b1-ecc2-26d60805e952.md)
#### Weitere Ressourcen


[Elemente des DistListItem-Objekts](http://msdn.microsoft.com/library/3ba4af84-ce84-61d9-1bc9-fab41bf6f125%28Office.15%29.aspx)