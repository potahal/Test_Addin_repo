
# Application.VisualReportsSaveDatabase メソッド (Project)

既定のディレクトリ、または指定したディレクトリにビジュアル レポート データベースを保存します。


## 構文

 _式_. **VisualReportsSaveDatabase**( ** _strNamePath_**, ** _PjVisualReportsDataLevel_** )

 _式_ **Application** オブジェクトを表す変数です。


### パラメーター



|**名前**|**必須 / オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _strNamePath_|省略可能|**文字列型 (String)**|データベース ファイル (.mbd) の保存先にする場所の名前と完全パスを指定します。|
| _PjVisualReportsDataLevel_|省略可能|**長整数型 (Long)**|データ レベルを保存します。 **[PjVisualReportsDataLevel](56792ea8-6459-38ef-e994-95024e6d8fe9.md)** 定数のいずれかをすることができます。既定では **pjLevelAutomatic** です。|

### 戻り値

 **ブール型 (Boolean)**


## 注釈

PjVisualReportsDataLevel パラメーターでは、タイム スケール領域のデータをアクセスできるレベルを指定します。などの場合は **pjLevelMonths** (月数) が指定されている、 **pjLevelDays** にアクセスすることはできません (日) です。


## 例

 **VisualReportsSaveDatabase** メソッドを使用する例を次に示します。


```
Sub a() 
 Dim tf As Boolean 
 tf = Application.VisualReportsSaveDatabase("C:\mydb.mdb", pjLevelAutomatic) 
 If tf = True Then 
 MsgBox ("Database saved successfully") 
 Else 
 MsgBox ("Database wasn't saved successfully") 
 End If 
End Sub 

```

