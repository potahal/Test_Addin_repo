

---
ms.Toctitle:Project.SetObjectMatchingID メソッド (Project)
title:Project.SetObjectMatchingID メソッド (Project)
ms.ContentId:d0d79e0a-bfec-9882-bfe9-72f7c51f0baf
---
# Project.SetObjectMatchingID メソッド (Project)




[**構成内容変更**] ダイアログ ボックス内のオブジェクトの照合 ID 値を設定し、たとえば、"Gantt Chart" で指定されたビューを変更します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SetObjectMatchingID**(**ObjectType**, **ObjectName**, **MatchingID**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Project** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*ObjectType*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**長整数型 (Long)**|オブジェクトの種類を **pjOrganizer** クラスの定数で指定します。|
|*ObjectName*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|オブジェクトの表示名を指定します。|
|*MatchingID*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|設定する照合 ID を示す文字列を指定します。|





## 例
次の例では、表示名が "Gantt Chart" でオブジェクトの種類が **pjView** の照合 ID を "Gantt Chart 1" に設定します。

```vba
ActiveProject.SetObjectMatchingID ObjectType:=pjView, ObjectName:="Gantt Chart", MatchingID:="Gantt Chart 1"
```





