
# ????????????????????????????

????????????????  **[Table](0affaafd-93fe-227a-acee-e09a86cadc20.md)** ????????????????????????????????????????????????????????????????? **Table** ?????????????????????????????????????????????????

????? ?????????????  **Categories** ????????????? **urn:schemas-microsoft-com:office:office#Keywords** ????????????? **Table** ??????? **Table** ???????? **Categories** ???????????



```
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```

variant??????????????????????????????????????????????????????variant ??????????????????????????????????????????????????????????????



```
oRow("urn:schemas-microsoft-com:office:office#Keywords")
```

? Empty ???????



```
Sub TableCategories() 
    Dim oT As Outlook.Table 
    Dim oRow As Outlook.Row 
    Dim varCat 
    Dim j As Integer 
    Dim strCategories As String 
 
    Set oT = Application.ActiveExplorer.CurrentFolder.GetTable() 
    oT.Columns.Add ("urn:schemas-microsoft-com:office:office#Keywords") 
    oT.Sort "LastModificationTime", True 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        'Obtain any values of the Categories property 
        varCat = oRow("urn:schemas-microsoft-com:office:office#Keywords") 
        If Not (IsEmpty(varCat)) Then 
            'Form a string out of the item's categories 
            For j = 0 To UBound(varCat) 
                strCategories = strCategories &amp; (varCat(j)) &amp; ", " 
            Next 
            'Remove last trailing ", " 
            strCategories = Left(strCategories, Len(strCategories) - 2) 
        Else 
            'The item does not have any categories 
            strCategories = "" 
        End If 
        Debug.Print ("Subject: " _ 
           &amp; oRow("Subject") &amp; vbCrLf &amp; "Categories: ") &amp; strCategories &amp; vbCrLf 
    Loop 
End Sub
```

