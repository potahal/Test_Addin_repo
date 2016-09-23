
# Project.SetObjectMatchingID メソッド (Project)

[ **構成内容変更**] ダイアログ ボックス内のオブジェクトの照合 ID 値を設定し、たとえば、"Gantt Chart" で指定されたビューを変更します。


## 構文

 _式_. **SetObjectMatchingID**( ** _ObjectType_**, ** _ObjectName_**, ** _MatchingID_** )

 _式_ **Project** オブジェクトを表す変数です。


### パラメーター



|**名前**|**必須 / オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|必須|**長整数型 (Long)**|オブジェクトの種類を  **[pjOrganizer](d176be88-4df9-3826-c806-f7f650fffb39.md)** クラスの定数で指定します。|
| _ObjectName_|必須|**文字列型 (String)**|オブジェクトの表示名を指定します。|
| _MatchingID_|必須|**文字列型 (String)**|設定する照合 ID を示す文字列を指定します。|

## 例

次の例では、表示名が "Gantt Chart" でオブジェクトの種類が  **pjView** の照合 ID を "Gantt Chart 1" に設定します。


```
ActiveProject.SetObjectMatchingID ObjectType:=pjView, ObjectName:="Gantt Chart", MatchingID:="Gantt Chart 1"
```

