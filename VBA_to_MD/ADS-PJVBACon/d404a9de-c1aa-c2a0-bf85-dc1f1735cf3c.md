
---
ms.Toctitle:メソッド (プロジェクト)
title:メソッド (プロジェクト)
ms.ContentId:d404a9de-c1aa-c2a0-bf85-dc1f1735cf3c
---
# メソッド (プロジェクト)





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AddChart***(Style,**Type,**Left,**Top,**Width,**Height,**NewLayout)*




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Shapes** オブジェクトを表す変数。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Style*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**Integer**|グラフの色のスタイルを指定します。値は**、グラフのスタイル**] で [**デザイン**] タブの [**グラフ ツール**リボンの [**色の変更**ボックスの一覧に対応 (ただし、値が同じ順序ではありません)。|
|*Type*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**XlChartType**|縦棒グラフや円グラフなどを追加するグラフの種類。|
|*Left*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**Single**|グラフの左端からポイント単位で位置を指定します。|
|*Top*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**Single**|グラフの上端からポイント単位で位置を指定します。|
|*Width*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**Single**|グラフの幅は、ポイント単位で指定します。|
|*Height*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**Single**|グラフの高さは、ポイント単位で指定します。|
|*NewLayout*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**Boolean**|*NewLayout*は、 UNRESOLVED_TOKEN_VAL(pjgenericshort)では使用されません。|
|*Style*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |INT||
|*Type*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |可能||
|*Left*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |FLOAT||
|*Top*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |FLOAT||
|*Width*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |FLOAT||
|*Height*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |FLOAT||
|*NewLayout*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |bool||
|名前|必須/オプション|データ型|説明|



### 戻り値
**Shape**





## 例
次の例では、バーのオレンジ色のバーを表示したグラフの種類の既定値を持つレポートを作成します。

```vba
Sub AddDefaultChart()
    Dim chartReport As Report
    Dim reportName As String
    
    ' Add a report.
    reportName = "Test chart report"
    Set chartReport = ActiveProject.Reports.Add(reportName)

    ' Add a chart.
    Dim chartShape As shape
    Set chartShape = ActiveProject.Reports(reportName).Shapes.AddChart(Style:=12)
    
    With chartShape
        .Chart.SetElement msoElementChartTitleAboveChart
        .Chart.ChartTitle.Text = "Test Chart"
    End With
End Sub
```




## Related Topics

[図形オブジェクト](6e42040c-dd5a-de4c-afa8-f9e33d1e5054.md)

[Shape オブジェクト](d2b32bcd-5595-a4a7-9772-feb25fd0103a.md)

[グラフ オブジェクト](810d4ec1-69d2-c432-b9da-57042b783b85.md)




