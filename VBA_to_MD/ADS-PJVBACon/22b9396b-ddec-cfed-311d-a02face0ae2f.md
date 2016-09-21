

---
ms.Toctitle:Application.SelectResourceColumn メソッド (Project)
title:Application.SelectResourceColumn メソッド (Project)
ms.ContentId:22b9396b-ddec-cfed-311d-a02face0ae2f
---
# Application.SelectResourceColumn メソッド (Project)




リソースの情報を含む列を選択します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SelectResourceColumn**(**Column**, **Additional**, **Extend**, **Add**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Column*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|選択する列のフィールド名を指定します。既定値は、アクティブ セルが含まれている列です。|
|*Additional*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**整数型 (Integer)**|**列**の右側を選択する追加の列の数です。**Extend**が**True**の場合は、**その他**は無視されます。既定値は 0 です。|
|*Extend*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True の**現在の選択範囲と**列**の間のすべての列が選択されている場合です。既定値は、 **false を指定**します。|
|*Add*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True の**現在の列が選択範囲に含まれている場合です。既定値は、 **false を指定**します。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
**SelectResourceColumn**メソッドを使用可能なは、リソース シート] または [リソース配分状況] ビューがアクティブなビューのみです。



## 例
次の使用例は、[**状況説明マーク**] 列とその隣の 2 列を選択します。

```vba
Sub Select_ResourceColumn() 
 
 'Activate Resource Sheet 
 ViewApply Name:="&Resource Sheet" 
 SelectResourceColumn Column:="Indicators", Additional:=2 
End Sub
```





