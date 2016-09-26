
# Project.ProjectNotes Property (Project)

Ruft ab oder legt die Notizen für das Projekt. Lese-/Schreibzugriff  **Zeichenfolge**.


## Syntax

 _Ausdruck_. **ProjectNotes**

 _Ausdruck_ Eine Variable, die ein **Project** -Objekt darstellt.


## Hinweise

Um das Dialogfeld  **Eigenschaften** von Project in Project angezeigt wird, wählen Sie die Registerkarte **Datei** auf dem Menüband die **Backstage**-Ansicht anzuzeigen, wählen die Registerkarte  **Info**, und wählen Sie dann im Dropdown-Menü  **Projektinformationen** **Erweiterte Eigenschaften**.


## Beispiel

Im folgenden Beispiel werden dem Feld  **Comments** im Dialogfeld **Eigenschaften** für das Projekt Uhrzeit und Datum hinzugefügt. Das Projekt wird anschließend gespeichert.


```
Sub SaveAndNoteTime() 
    Projects(1).ProjectNotes = Projects(1).ProjectNotes &amp; vbCrLf _ 
        &amp; "This project was last saved on " _ 
        &amp; Date$ &amp; " at " &amp; Time$ &amp; "." 
    FileSave 
End Sub
```

