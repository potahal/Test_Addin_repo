
# Durchsuchen des Posteingangs nach Elementen, deren Betreff 'Office' enthält

In diesem Thema werden zwei Codebeispiele gezeigt, in denen mithilfe von DASL-Abfragen nach Elementen im Posteingang gesucht wird, deren Betreffzeile "Office" enthält. Im ersten Codebeispiel wird  **[Folder.GetTable](08d184cb-0c41-01b1-abc5-305476380f8b.md)** verwendet, und im zweiten wird **[Application.AdvancedSearch](7b433d8b-08b9-dff1-b854-287d76b47a90.md)** verwendet, um die DASL-Abfrage anzuwenden.

In jedem Codebeispiel wird das Schlüsselwort  **ci_phrasematch** für die Inhaltsindizierung in einem DASL-Filter für die Eigenschaft **http://schemas.microsoft.com/mapi/proptag/0x0037001E** (die **Subject** -Eigenschaft, auf die durch den MAPI-ID-Namespace verwiesen wird) verwendet, um das Wort "office" im Betreff zu suchen. Der Filter wird (mithilfe von **Folder.GetTable** oder **Application.AdvancedSearch** ) auf Elemente im Posteingang angewendet, und die Betreffzeilen der einzelnen durch die Suche zurückgegebenen Elemente werden gedruckt.

 **Hinweis**  Da bei der Übereinstimmung die Groß-/Kleinschreibung nicht beachtet wird, wird durch  **Folder.GetTable** oder **Application.AdvancedSearch** jedes Element, dessen Betreff "Office" oder "office" enthält, zurückgegeben. Beachten Sie, dass in beiden Beispielen der Betreff jeder Zeile im sich ergebenden **[Table](0affaafd-93fe-227a-acee-e09a86cadc20.md)** -Objekt gedruckt wird. Das kleinere **Table** -Objekt wird anstelle des **[Search.Results](405166fa-d0bc-33d2-f4aa-908fb821edd6.md)** -Objekts verwendet, um eine bessere Leistung zu erzielen. Die **Subject** -Eigenschaft ist in einem **Table** -Objekt enthalten, das durch eine Suche für einen beliebigen Ordner zurückgegeben wird. Der Posteingang kann jedoch wie jeder Ordner in Outlook heterogene Elemente enthalten und ist nicht auf E-Mail-Elemente beschränkt. Wenn Sie auf eine für einen bestimmten Elementtyp im Posteingang spezifische Eigenschaft zugreifen möchten, verwenden Sie **[Columns.Add](d438cfeb-629f-4234-6f4f-ffa086ef9a41.md)**, um diese Eigenschaft einzuschließen und das **Table** -Objekt zu aktualisieren. Überprüfen Sie für jede im **Table** -Objekt zurückgegebene Zeile den Nachrichtentyp des Elements, bevor Sie auf die Eigenschaft zugreifen.

In diesem Codebeispiel wird zum Ausführen der Suche  **Folder.GetTable** verwendet:



```
Sub RestrictTableForInbox() 
    Dim oT As Outlook.Table 
    Dim strFilter As String 
    Dim oRow As Outlook.Row 
     
    'Construct filter for Subject containing 'Office' 
    Const PropTag  As String = "http://schemas.microsoft.com/mapi/proptag/" 
    strFilter = "@SQL=" &amp; Chr(34) &amp; PropTag  _ 
        &amp; "0x0037001E" &amp; Chr(34) &amp; " ci_phrasematch 'Office'" 
     
    'Do search and obtain Table on Inbox 
    Set oT = Application.Session.GetDefaultFolder(olFolderInbox).GetTable(strFilter) 
     
    'Print Subject of each returned item 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        Debug.Print oRow("Subject") 
    Loop 
End Sub
```

In diesem Codebeispiel wird zum Ausführen der Suche  **Application.AdvancedSearch** verwendet:



```
Public blnSearchComp As Boolean 
Private Sub Application_AdvancedSearchComplete(ByVal SearchObject As Search) 
    MsgBox "The AdvancedSearchComplete Event fired" 
    blnSearchComp = True 
End Sub 
 
Sub TestSearchWithTable() 
    Dim oSearch As Search 
    Dim oTable As Table 
    Dim strQuery As String 
    Dim oRow As Row 
         
    blnSearchComp = False 
     
    'Construct filter. 0x0037001E represents Subject 
    strQuery = _ 
        "http://schemas.microsoft.com/mapi/proptag/0x0037001E" &amp; _ 
        " ci_phrasematch 'Office'" 
     
    'Do search 
    Set oSearch = _ 
        Application.AdvancedSearch("Inbox", strQuery, False, "Test") 
    While blnSearchComp = False 
        DoEvents 
    Wend 
 
    'Obtain Table 
    Set oTable = oSearch.GetTable 
     
    'Print Subject of each returned item 
    Do Until oTable.EndOfTable 
        Set oRow = oTable.GetNextRow 
        Debug.Print oRow("Subject") 
    Loop 
End Sub

```

