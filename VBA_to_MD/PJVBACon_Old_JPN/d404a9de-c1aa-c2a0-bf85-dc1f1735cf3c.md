
# メソッド (プロジェクト)
アクティブなレポートの指定した位置にグラフを作成します。グラフを表す **Shape**オブジェクトを返します。

## 構文

 _式_. **AddChart** _(Style,_ _Type,_ _Left,_ _Top,_ _Width,_ _Height,_ _NewLayout)_

 _式_ **Shapes** オブジェクトを表す変数。


### パラメーター



|**名前**|**必須/オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _Style_|省略可能|**Integer**|グラフの色のスタイルを指定します。値は **、グラフのスタイル**] で [ **デザイン**] タブの [ **グラフ ツール**リボンの [ **色の変更**ボックスの一覧に対応 (ただし、値が同じ順序ではありません)。|
| _Type_|省略可能|**XlChartType**|縦棒グラフや円グラフなどを追加するグラフの種類。|
| _Left_|省略可能|**Single**|グラフの左端からポイント単位で位置を指定します。|
| _Top_|省略可能|**Single**|グラフの上端からポイント単位で位置を指定します。|
| _Width_|省略可能|**Single**|グラフの幅は、ポイント単位で指定します。|
| _Height_|省略可能|**Single**|グラフの高さは、ポイント単位で指定します。|
| _NewLayout_|省略可能|**Boolean**| _NewLayout_は、 Projectでは使用されません。|
| _Style_|省略可能|INT||
| _Type_|省略可能|可能||
| _Left_|省略可能|FLOAT||
| _Top_|省略可能|FLOAT||
| _Width_|省略可能|FLOAT||
| _Height_|省略可能|FLOAT||
| _NewLayout_|省略可能|bool||
|名前|必須/オプション|データ型|説明|

### 戻り値

 **Shape**


## 例

次の例では、バーのオレンジ色のバーを表示したグラフの種類の既定値を持つレポートを作成します。


```
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


## 関連項目


#### その他の技術情報


[図形オブジェクト](6e42040c-dd5a-de4c-afa8-f9e33d1e5054.md)
[Shape オブジェクト](d2b32bcd-5595-a4a7-9772-feb25fd0103a.md)
[グラフ オブジェクト](810d4ec1-69d2-c432-b9da-57042b783b85.md)