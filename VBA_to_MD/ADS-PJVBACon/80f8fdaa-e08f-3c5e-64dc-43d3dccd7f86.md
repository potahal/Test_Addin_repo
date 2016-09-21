

---
ms.Toctitle:Application.GanttBarLinks メソッド (Project)
title:Application.GanttBarLinks メソッド (Project)
ms.ContentId:80f8fdaa-e08f-3c5e-64dc-43d3dccd7f86
---
# Application.GanttBarLinks メソッド (Project)




[ガント チャート] ビューのタスクのリンク表示形式を設定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GanttBarLinks**(**Display**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Display*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|先行タスクのリンクの端からのリンクが描画されます。**PjGanttBarLink**定数のいずれかをすることができます。既定値は**PjNoGanttBarLinks**です。|



### 戻り値
**ブール型 (Boolean)**





## 例
次の例では、最初にリンクをクリアし、次にそのリンクを 1 つのタスク バーの終わりから次のタスク バーの先頭に表示します。

```vba
Sub GanttBar_Links() 
'First clear links, than links from end to top of the next bar 
 'Activate Gantt Chart view 
 ViewApply Name:="&Gantt Chart" 
 GanttBarLinks pjNoGanttBarLinks 
 GanttBarLinks pjToTop 
End Sub
```





