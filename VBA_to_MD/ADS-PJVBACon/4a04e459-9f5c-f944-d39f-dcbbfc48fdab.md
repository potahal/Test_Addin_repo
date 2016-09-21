

---
ms.Toctitle:Application.UpdateTasks メソッド (Project)
title:Application.UpdateTasks メソッド (Project)
ms.ContentId:4a04e459-9f5c-f944-d39f-dcbbfc48fdab
---
# Application.UpdateTasks メソッド (Project)




選択したタスクを更新します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**UpdateTasks**(**PercentComplete**, **ActualDuration**, **RemainingDuration**, **ActualStart**, **ActualFinish**, **Notes**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*PercentComplete*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|アクティブ タスクの達成率を指定します。|
|*ActualDuration*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|選択したタスクの実績期間を指定します。|
|*RemainingDuration*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|選択したタスクの残存期間を指定します。|
|*ActualStart*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|選択したタスクの実績開始日を指定します。|
|*ActualFinish*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|選択したタスクの実績終了日を指定します。|
|*Notes*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|選択したタスクの [メモ] フィールドのコメントを指定します。使用できる値はテキストのみです。[**メモ**] ダイアログ ボックスのようにリッチ テキスト形式 (RTF) を使用することはできません。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
**UpdateTasks**メソッドを使用して引数を指定せずには、**タスクの更新**] ダイアログ ボックスが表示されます。



## 例
次の使用例は、「TestTask-1」という名前のタスクを作成し、タスクの達成率を 50% に更新します。次に、タスクを削除します。

```vba
Sub Update_Tasks() 
 
 'Activate Gantt Chart 
 ViewApply Name:="Gantt Chart" 
 
 'Create a task 
 RowInsert 
 SetTaskField Field:="Name", Value:="TestTask-1" 
 SetTaskField Field:="Duration", Value:="2" 
 
 'Update the percent complete of the new task. 
 UpdateTasks PercentComplete:="50" 
 
 'Delete the new task 
 ActiveProject.Tasks("TestTask-1").Delete 
End Sub
```





