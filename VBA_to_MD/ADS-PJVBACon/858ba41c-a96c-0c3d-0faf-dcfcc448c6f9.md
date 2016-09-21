
---
ms.Toctitle:Chart.DataTable プロパティ (プロジェクト)
title:Chart.DataTable プロパティ (プロジェクト)
ms.ContentId:858ba41c-a96c-0c3d-0faf-dcfcc448c6f9
---
# Chart.DataTable プロパティ (プロジェクト)





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**DataTable**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Chart** オブジェクトを表す変数。



## 注釈
**IMsoDataTable**オブジェクトを表示するには、オブジェクト ブラウザーで右クリックし、**隠しメンバーの表示]**を選択し。



## 例
次の使用例は、アクティブなレポートのグラフに外枠の罫線が付いたデータ テーブルを追加します。

```vba
Sub ShowDataTable()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    With chartShape.Chart
        .HasDataTable = True
        .DataTable.HasBorderOutline = True
    End With
End Sub
```




## プロパティ値
**IMSODATATABLE**



## Related Topics

[Chart オブジェクト](810d4ec1-69d2-c432-b9da-57042b783b85.md)




