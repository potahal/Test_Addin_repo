

---
ms.Toctitle:Application.VisualReportsSaveDatabase メソッド (Project)
title:Application.VisualReportsSaveDatabase メソッド (Project)
ms.ContentId:edcbaff5-beb1-ba11-fb65-ec26a24ab23d
---
# Application.VisualReportsSaveDatabase メソッド (Project)




既定のディレクトリ、または指定したディレクトリにビジュアル レポート データベースを保存します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**VisualReportsSaveDatabase**(**strNamePath**, **PjVisualReportsDataLevel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*strNamePath*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|データベース ファイル (.mbd) の保存先にする場所の名前と完全パスを指定します。|
|*PjVisualReportsDataLevel*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|データ レベルを保存します。**PjVisualReportsDataLevel**定数のいずれかをすることができます。既定では**pjLevelAutomatic**です。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
PjVisualReportsDataLevel パラメーターでは、タイム スケール領域のデータをアクセスできるレベルを指定します。などの場合は**pjLevelMonths** (月数) が指定されている、 **pjLevelDays**にアクセスすることはできません (日) です。



## 例
**VisualReportsSaveDatabase**メソッドを使用する例を次に示します。

```vba
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





