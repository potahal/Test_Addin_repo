

---
ms.Toctitle:Task.LinkPredecessors メソッド (Project)
title:Task.LinkPredecessors メソッド (Project)
ms.ContentId:6aaf3dfc-3f8c-a7a7-9f7f-59bd1d5a50b3
---
# Task.LinkPredecessors メソッド (Project)




タスクに 1 つ以上の先行タスクを追加します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**LinkPredecessors**(**Tasks**, **Link**, **Lag**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Task** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Tasks*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**オブジェクト型 (Object)**|指定された**タスク**または**タスク**オブジェクトでは、**式**で指定されたタスクの先行タスクになります。|
|*Link*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|リンクを設定するタスク間の関係を指定する定数です。[PjTaskLinkType](141a1145-0eb5-3664-4755-394584aec8ac.md)定数のいずれかをすることができます。既定値は、 **pjFinishToStart**です。|
|*Lag*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|リンクされたタスク間のラグ タイムの期間を指定する文字列です。タスク間にリード タイムを指定するには、 **Lag**に負の値として評価される式を使用します。|





## 例
次の使用例は、タスク名の入力を求めるメッセージを表示し、ユーザーが指定したタスクを現在選択されているタスクの先行タスクとして設定します。

```vba
Sub LinkTasksFromPredecessor() 
    Dim Entry As String   ' Task name entered by user 
    Dim T As Task         ' Task object used in For Each loop 
    Dim I As Long         ' Used in For loop 
    Dim Exists As Boolean ' Whether or not the task exists 
 
    Entry = InputBox$("Enter the name of a task:") 
 
    Exists = False ' Assume task doesn't exist. 
 
    ' Search active project for the specified task. 
    For Each T In ActiveProject.Tasks 
        If T.Name = Entry Then 
            Exists = True 
            ' Make the task a predecessor of the selected tasks. 
            For I = 1 To ActiveSelection.Tasks.Count 
                ActiveSelection.Tasks(I).LinkPredecessors Tasks:=T 
            Next I 
        End If 
    Next T 
 
    ' If task doesn't exist, display an error and quit the procedure. 
    If Not Exists Then 
        MsgBox ("Task not found.") 
        Exit Sub 
    End If 
End Sub
```





