
# Table.TableFields Property (Project)

Dient zum Abrufen einer  **[TableFields](7f749404-0723-7a17-b83f-f43725c45fc5.md)** -Auflistung zurück, die Felder in der Tabelle darstellt. Read-only **TableFields**.


## Syntax

 _Ausdruck_. **TableFields**

 _Ausdruck_ Eine Variable, die ein **Table** -Objekt darstellt.


## Beispiel

Im folgenden Beispiel wird die Ausrichtung einer Spalte in einer Eingabetabelle geändert. Das Makro fordert zu der Eingabe auf, welche Spalte zentriert ausgerichtet werden soll, ändert die Anzeige und aktualisiert die Ansicht.


```
Sub AutoWrap() 
 Dim fieldNumber As Integer 
 
 fieldNumber = InputBox$(Prompt:="Enter the number of the " _ 
 &amp; "column you want to center in the Entry table." _ 
 &amp; Chr(13) &amp; "For example, Column 1 is the Indicators " _ 
 &amp; "column.") 
 
 ActiveProject.TaskTables("Entry").TableFields(fieldNumber _ 
 + 1).AlignData = pjCenter 
 
 TableApply Name:="&amp;Entry" 
End Sub
```

