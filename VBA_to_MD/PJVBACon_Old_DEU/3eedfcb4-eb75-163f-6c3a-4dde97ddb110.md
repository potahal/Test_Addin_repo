
# Application.AlignTableCellBottom Method (Project)
Richtet Text am unteren Rand der Zelle, für die ausgewählten Zellen in einer Berichtstabelle aus.

## Syntax

 _Ausdruck_. **AlignTableCellBottom**

 _Ausdruck_ Eine Variable, die ein Objekt Application **Application** repräsentiert.


### Rückgabewert

 **Boolean**


## Beispiel

Im folgenden Beispiel richtet das Makro  **AlignTableCells** den Text für alle Tabellen des angegebenen Berichts.


```
Sub TestAlignReportTables()
    Dim reportName As String
    Dim alignment As String   ' The value can be "top", "center", or "bottom".
    
    reportName = "Align Table Cells Report"
    alignment = "top"
    
    AlignTableCells reportName, alignment
End Sub

' Align the text for all tables in a specified report.
Sub AlignTableCells(reportName As String, alignment As String)
    Dim theReport As Report
    Dim shp As Shape
    
    Set theReport = ActiveProject.Reports(reportName)
    
    ' Activate the report. If the report is already active,
    ' ignore the run-time error 1004 from the Apply method.
    On Error Resume Next
    theReport.Apply
    On Error GoTo 0
    
    For Each shp In theReport.Shapes
        Debug.Print "Shape: " &amp; shp.Type &amp; ", " &amp; shp.Name
        
        If shp.HasTable Then
            shp.Select
            
            Select Case alignment
                Case "top"
                    AlignTableCellTop
                Case "center"
                    AlignTableCellVerticalCenter
                Case "bottom"
                    AlignTableCellBottom
                Case Else
                    Debug.Print "AlignTableCells error: " &amp; vbCrLf _
                        &amp; "alignment must be top, center, or bottom."
                End Select
        End If
    Next shp
End Sub
```


## Siehe auch


#### Konzepte


[Application-Objekt](8eb91712-7784-a102-38c0-19bb056c27e9.md)
#### Weitere Ressourcen


[Report-Objekt](38ef993e-e5cd-b451-06aa-41eb0e93450e.md)
[AlignTableCellTop-Methode](51eca157-64c4-f114-243e-895d97adf45a.md)
[AlignTableCellVerticalCenter-Methode](c790d8f7-e792-0718-3166-312640ff3f73.md)