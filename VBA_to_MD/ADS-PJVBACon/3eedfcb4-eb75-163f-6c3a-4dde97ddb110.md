
---
ms.Toctitle:Application.AlignTableCellBottom メソッド (プロジェクト)
title:Application.AlignTableCellBottom メソッド (プロジェクト)
ms.ContentId:3eedfcb4-eb75-163f-6c3a-4dde97ddb110
---
# Application.AlignTableCellBottom メソッド (プロジェクト)





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AlignTableCellBottom**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数。

### 戻り値
**Boolean**





## 例
次の例では、 **AlignTableCells**マクロは、指定したレポートのすべてのテーブルのテキストを配置します。

```vba
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
        Debug.Print "Shape: " & shp.Type & ", " & shp.Name
        
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
                    Debug.Print "AlignTableCells error: " & vbCrLf _
                        & "alignment must be top, center, or bottom."
                End Select
        End If
    Next shp
End Sub
```




## Related Topics

[アプリケーション オブジェクト](8eb91712-7784-a102-38c0-19bb056c27e9.md)

[レポート オブジェクト](38ef993e-e5cd-b451-06aa-41eb0e93450e.md)

[AlignTableCellTop メソッド](51eca157-64c4-f114-243e-895d97adf45a.md)

[AlignTableCellVerticalCenter メソッド](c790d8f7-e792-0718-3166-312640ff3f73.md)




