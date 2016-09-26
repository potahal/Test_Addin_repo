
# Conflicts.GetLast Method (Outlook)

Gibt das letzte Objekt in der  **[Conflicts](c4e1c060-519a-a6d1-8fb2-c7dfa1e3e66f.md)** -Auflistung zurück.


## Syntax

 _Ausdruck_. **GetLast**

 _Ausdruck_ Eine Variable, die ein **Conflicts** -Objekt darstellt.


### Rückgabewert

Ein  **[Conflict](a7c8f12a-08ba-9fff-60b8-a02d1c7f6f33.md)** -Objekt, das das letzte in der Auflistung enthaltene Objekt darstellt.


## Bemerkungen

Es gibt  **Nothing** zurück, wenn kein letztes Objekt vorhanden, beispielsweise, ist wenn die Auflistung leer ist. Um die **[GetFirst](f257a9f1-d9ec-c13a-62f7-0228d55342da.md)**, **GetNext**, **[GetNext](2e21ea88-c732-17ee-cd87-698fee992269.md)** und **[GetPrevious](23b5d75a-e1eb-7164-df92-71e37a1ec79f.md)** Methoden in einer großen Auflistung sicherzustellen, rufen Sie **GetFirst**, bevor Sie **GetNext für diese Auflistung** und **GetLast, bevor Sie  **GetPrevious** aufrufen**. Um sicherzustellen, dass Sie die Aufrufe immer auf die gleiche Auflistung ausführen, erstellen Sie eine explizite Variable, die auf diese vor dem Durchführen einer Schleife.


## Siehe auch


#### Konzepte


[Conflicts-Objekt](c4e1c060-519a-a6d1-8fb2-c7dfa1e3e66f.md)
#### Weitere Ressourcen


[Elemente des Conflicts-Objekts](http://msdn.microsoft.com/library/dcc61922-d119-1bb9-c175-a80a73599559%28Office.15%29.aspx)