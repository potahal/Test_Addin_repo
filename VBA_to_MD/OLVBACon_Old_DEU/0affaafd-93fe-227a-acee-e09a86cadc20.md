
# Table-Objekt (Outlook)

Stellt einen Satz Elementdaten aus einem  **[Folder](3cf6cda8-6d70-666e-2643-9d9c5b9cacfc.md)** - oder **[Search](226a5d49-3caf-90dd-725c-265404d1939f.md)** -Objekt dar, wobei die Elemente die Zeilen der Tabelle und die Eigenschaften die Spalten der Tabelle darstellen.


## Bemerkungen

Die  **Tabelle** stellt ein Read-only dynamisches Rowset von Daten in einem **Folder-** oder **Search** -Objekt. Sie können die **[Folder.GetTable](08d184cb-0c41-01b1-abc5-305476380f8b.md)** oder **[Search.GetTable](3aba6b77-73a3-9620-9c18-b2e03c7b63bc.md)** verwenden, um ein **Table** -Objekt abzurufen, die eine Gruppe von Elementen in einem Ordner oder Suchordner darstellt. Wenn das **Table** -Objekt von **Folder.GetTable** abgerufen wird, können Sie einen Filter (in **[Table.Restrict](ecdd30f6-e12c-8025-3ded-592d2fad2bb8.md)** ) weiter angeben, um eine Teilmenge der Elemente in den Ordner abzurufen. Wenn Sie alle Filter nicht angeben, erhalten Sie alle Elemente in den Ordner.

Standardmäßig enthält jedes Element in die zurückgegebene  **Tabelle** nur einen Standard-Teil seiner Eigenschaften. Sie können jede Zeile einer **Tabelle** als Element im Ordner, jeder Spalte als eine Eigenschaft des Elements und der **Tabelle** als eine in-Memory-Rowset, das schnelle Enumeration ermöglicht und Filtern von Elementen im Ordner betrachten. Obwohl Hinzufügungen und Löschvorgänge des zugrunde liegenden Ordners durch die Zeilen der **Tabelle** wiedergegeben werden, unterstützt der **Tabelle** keine Ereignisse für das Hinzufügen, ändern und Löschen von Zeilen. Wenn Sie ein beschreibbaren-Objekt aus **der Tabellenzeile** benötigen, erhalten die Eintrags-ID für die Zeile aus der Standard-Eintrags-ID-Spalte in der **Tabelle** und dann die **[GetItemFromID](f2abff80-4c04-998b-654b-28600424a16f.md)** -Methode des **[NameSpace](f0dcaa19-07f5-5d42-a3bf-2e42b7885644.md)** -Objekts verwenden, um ein vollständiges Element abzurufen, wie ein **[MailItem-Objekt](14197346-05d2-0250-fa4c-4a6b07daf25f.md)** oder ein **[ContactItem-Objekt](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)**, unterstützt, die Lese-/ Schreibvorgänge. Weitere Informationen zu standardmäßigen Spalten in einer **Tabelle** finden Sie unter[In einem Table-Objekt angezeigte Standardeigenschaften](649c64f3-2d1e-23f1-bf13-3368da79e62b.md).

Weitere Informationen auf das  **Table** -Objekt finden Sie unter[aufzählen, suchen und Filtern von Elementen in einem Ordner](d786d292-7a0e-0e1a-e132-affbfde37744.md).


## Beispiel

Im folgenden Codebeispiel wird veranschaulicht, wie das  **Table** -Objekt eine gefilterte Auswahl von Elementen basierend auf der **LastModificationTime** -Eigenschaft zurückgibt. Es wird gezeigt, wie die Standardeigenschaften sowie bestimmte Eigenschaften für die Elemente aufgelistet.


```
Sub DemoTable() 
 
 'Declarations 
 
 Dim Filter As String 
 
 Dim oRow As Outlook.Row 
 
 Dim oTable As Outlook.Table 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 'Get a Folder object for the Inbox 
 
 Set oFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 'Define Filter to obtain items last modified after May 1, 2005 
 
 Filter = "[LastModificationTime] > '5/1/2005'" 
 
 'Restrict with Filter 
 
 Set oTable = oFolder.GetTable(Filter) 
 
 
 
 'Remove all columns in the default column set 
 
 oTable.Columns.RemoveAll 
 
 'Specify desired properties 
 
 With oTable.Columns 
 
 .Add ("Subject") 
 
 .Add ("LastModificationTime") 
 
 'PR_ATTR_HIDDEN referenced by the MAPI proptag namespace 
 
 .Add ("http://schemas.microsoft.com/mapi/proptag/0x10F4000B") 
 
 End With 
 
 
 
 'Enumerate the table using test for EndOfTable 
 
 Do Until (oTable.EndOfTable) 
 
 Set oRow = oTable.GetNextRow() 
 
 Debug.Print (oRow("Subject")) 
 
 Debug.Print (oRow("LastModificationTime")) 
 
 Debug.Print (oRow("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) 
 
 Loop 
 
End Sub
```


## Methoden



|**Name**|
|:-----|
|[FindNextRow](e09019ca-e4bb-2597-7b9e-a56c1b5fce6c.md)|
|[FindRow](5722cf58-d026-007a-558f-90b73bad920d.md)|
|[GetArray](2594bb2e-290f-8e88-52d1-cd2b2191bbe3.md)|
|[GetNextRow](e01ddaa0-a869-2f52-5e46-84d4d4090e61.md)|
|[GetRowCount](06014c43-700a-8502-bad7-b3f93a22e870.md)|
|[MoveToStart](af499471-dd21-9374-7399-3ce977368015.md)|
|[Einschränken](ecdd30f6-e12c-8025-3ded-592d2fad2bb8.md)|
|[Sortieren](4e4867c2-27b8-f920-59ce-b60116d22054.md)|

## Eigenschaften



|**Name**|
|:-----|
|[Anwendung](10e7611e-e3b3-a07c-da85-f8c270a37212.md)|
|[Klasse](bea314b0-9db9-ac67-a897-49e619da1066.md)|
|[Spalten](57005ab1-ad49-296d-5b34-24dfd8f0987f.md)|
|[EndOfTable](8c185230-65ce-1b66-7b63-8de3533dea86.md)|
|[Das übergeordnete](1c6a54ac-ba4d-72a2-0871-a3522582dbde.md)|
|[Sitzung](8a17876d-6637-f30b-6c0f-32cfc8b77d51.md)|

## Siehe auch


#### Konzepte


[Outlook-Objektmodellreferenz](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)
#### Weitere Ressourcen


[Elemente des Tabelle-Objekts](http://msdn.microsoft.com/library/bd9db35d-0738-22cf-a936-425d5a0ead87%28Office.15%29.aspx)