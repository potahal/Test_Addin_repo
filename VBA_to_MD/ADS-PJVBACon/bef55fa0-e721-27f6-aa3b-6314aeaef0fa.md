

---
ms.Toctitle:StartDriver.OverAllocatedAssignments プロパティ (Project)
title:StartDriver.OverAllocatedAssignments プロパティ (Project)
ms.ContentId:bef55fa0-e721-27f6-aa3b-6314aeaef0fa
---
# StartDriver.OverAllocatedAssignments プロパティ (Project)




タスク開始ドライバーの割り当てを超過する取得します。読み取り専用**OverAllocatedAssignments**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**OverAllocatedAssignments**(**fOverPeak**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **StartDriver** オブジェクトを返すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*overallocationType*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**PjOverallocationType**|割り当て超過の種類を決定する **PjOverallocationType** クラスの定数のいずれかを使用できます。|





## 注釈
割り当て超過は、マイルストーン、プレースホルダー タスク、または割り当てのないタスクでは発生しません。



## 例
次のコマンドは、リソースが他のタスクで作業中である割り当て超過の割り当ての数を返します。

```vba
Debug.Print ActiveProject.Tasks(2).StartDriver.OverAllocatedAssignments(pjOverallocationTypeWorkingOnOtherTasks).Count
```




## Related Topics

[StartDriver オブジェクト](4df2c386-a31e-faea-e286-d510f11cca57.md)




