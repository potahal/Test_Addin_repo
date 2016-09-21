

---
ms.Toctitle:Application.SelectBeginning メソッド (Project)
title:Application.SelectBeginning メソッド (Project)
ms.ContentId:4adf20ae-4fd2-818a-da8c-133c08cad7fb
---
# Application.SelectBeginning メソッド (Project)




作業中のテーブルまたはビューの最初のセルを選択します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SelectBeginning**(**Extend**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Extend*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**と、最初のセルに現在の選択範囲を拡張します。作業中のビューが、[ネットワーク ダイアグラム] または [リソース グラフの場合は、拡張は無視されます。既定値は、 **false を指定**します。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
リソースのグラフでは、 **SelectBeginning**は、最小の id 番号を持つリソースを選択します。ネットワーク ダイアグラム] ビューでは、 **SelectBeginning**は、ビューの左上隅に最も近いボックスを選択します。



## 例
次の使用例は、行 4 の [名前] フィールドをガント チャートの開始フィールドとして選択します。

```vba
Sub Select_Beginning() 
 
 ViewApply Name:="&Gantt Chart" 
 SelectTaskField Row:=4, Column:="Name", RowRelative:=False 
 
 SelectBeginning Extend:=True 
End Sub
```





