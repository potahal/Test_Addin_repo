
# OfficeDataSourceObject.ApplyFilter ???? (Office)

????????? ????????????????????????????????????????????????


## ??

 _?_. **ApplyFilter**

 _?_ **OfficeDataSourceObject** ??????????????????


## ?

???????Region ???????????????????????????????????????????????


```
Sub OfficeFilters() 
 Dim appOffice As OfficeDataSourceObject 
 Dim appFilters As ODSOFilters 
 
 Set appOffice = Application.OfficeDataSourceObject 
 
 appOffice.Open bstrConnect:="DRIVER=SQL Server;SERVER=ServerName;" &amp; _ 
 "UID=user;PWD=;DATABASE=Northwind", bstrTable:="Employees" 
 
 Set appFilters = appOffice.Filters 
 
 MsgBox appOffice.RowCount 
 
 appFilters.Add Column:="Region", Comparison:=msoFilterComparisonEqual, _ 
 Conjunction:=msoFilterConjunctionAnd, bstrCompareTo:="WA" 
 appOffice.ApplyFilter 
 
 MsgBox appOffice.RowCount 
 
End Sub
```


## ????


#### ??


[OfficeDataSourceObject ??????](d5e5401b-643e-c12c-2648-f281af481f45.md)
#### ????????


[OfficeDataSourceObject ???????????](http://msdn.microsoft.com/library/57ba0dc6-80e7-04a9-a619-2a3e6aa2cdff%28Office.15%29.aspx)