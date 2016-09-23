
# StartDriver.OverAllocatedAssignments プロパティ (Project)

タスク開始ドライバーの割り当てを超過する取得します。読み取り専用 **OverAllocatedAssignments** です。


## 構文

 _式_. **OverAllocatedAssignments**( ** _fOverPeak_** )

 _式_ **StartDriver** オブジェクトを返すオブジェクト式を指定します。


### パラメーター



|**名前**|**必須/オプション**|**データ型**|**説明**|
|:-----|:-----|:-----|:-----|
| _overallocationType_|必須|**PjOverallocationType**|割り当て超過の種類を決定する  **[PjOverallocationType](b2eaea51-6884-194c-9a68-75669fcc8283.md)** クラスの定数のいずれかを使用できます。|

## 注釈

割り当て超過は、マイルストーン、プレースホルダー タスク、または割り当てのないタスクでは発生しません。


## 例

次のコマンドは、リソースが他のタスクで作業中である割り当て超過の割り当ての数を返します。


```
Debug.Print ActiveProject.Tasks(2).StartDriver.OverAllocatedAssignments(pjOverallocationTypeWorkingOnOtherTasks).Count
```


## 関連項目


#### 概念


[StartDriver オブジェクト](4df2c386-a31e-faea-e286-d510f11cca57.md)