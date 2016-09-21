

---
ms.Toctitle:OfficeDataSourceObject.ApplyFilter メソッド (Office)
title:OfficeDataSourceObject.ApplyFilter メソッド (Office)
ms.ContentId:9ce3ed9b-3d84-9ebd-68ae-be958ad1a99c
---
# OfficeDataSourceObject.ApplyFilter メソッド (Office)




差し込み印刷データ ファイルにフィルターを適用し、指定したレコードから指定した条件を満たすレコードのみを抽出します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ApplyFilter**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **OfficeDataSourceObject** オブジェクトを表す変数を指定します。



## 例
次の使用例は、Region フィールドが空白のレコードをすべて削除する新しいフィルターを追加し、作業中の文書に適用します。

```sourcecode
Sub OfficeFilters() 
 Dim appOffice As OfficeDataSourceObject 
 Dim appFilters As ODSOFilters 
 
 Set appOffice = Application.OfficeDataSourceObject 
 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" & _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 Set appFilters = appOffice.Filters 
 
 MsgBox appOffice.RowCount 
 
 appFilters.Add Column:="Region", Comparison:=msoFilterComparisonEqual, _ 
 Conjunction:=msoFilterConjunctionAnd, bstrCompareTo:="WA" 
 appOffice.ApplyFilter 
 
 MsgBox appOffice.RowCount 
 
End Sub
```




## Related Topics

[OfficeDataSourceObject オブジェクト](d5e5401b-643e-c12c-2648-f281af481f45.md)

[OfficeDataSourceObject オブジェクトのメンバー](57ba0dc6-80e7-04a9-a619-2a3e6aa2cdff.md)




