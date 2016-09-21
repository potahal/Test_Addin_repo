
---
ms.Toctitle:Application.AddSiteColumn メソッド (プロジェクト)
title:Application.AddSiteColumn メソッド (プロジェクト)
ms.ContentId:0ec78b0b-b4bf-3dea-0ed6-af78798bd7cd
---
# Application.AddSiteColumn メソッド (プロジェクト)





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AddSiteColumn***(ProjectField*, *SharePointName)*




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*ProjectField*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**PjField**|列挙では、 **PjField** 、新しい列に表示するのには、プロジェクトのフィールドを指定する定数のサブセットのいずれか(「解説」を参照してください)、禁止されているフィールドのいずれかをすることはできません。|
|*SharePointName*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**Variant**|新しい列の名前です。|
|*ProjectField*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |PJFIELD||
|*SharePointName*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |VARIANT||
|名前|必須/オプション|データ型|説明|



### 戻り値
**Boolean**



**True**場合は、列を追加します。





## 注釈
**AddSiteColumn**メソッドは、次の場合に実行時エラー 1004、「アプリケーション定義またはオブジェクト定義エラー、」を提供します。

- 作業中のプロジェクトは、同期された SharePoint タスク リストに関連付けられているではありません。プロジェクトがローカルの場合は、同期された SharePoint タスク リストを作成するのには、 **LinkToTaskList**メソッドを使用できます。
- SharePoint タスク リスト内の列名は既に存在します。列名の一覧を表示するには、SharePoint で、[タスク] リストを開くし、[**リスト**] タブで**[ビューの変更**を選択します。**の設定ですか。ビューを編集する**ページにすべてのタスク] ボックスの一覧で利用可能な列名が表示されます。
- *ProjectField*値は、 **pjResourceActualCost**など、タスクではないフィールドです。SharePoint タスク リストは、タスク フィールド、リソース フィールドではなくを示しています。
- *ProjectField*値は、 **pjTaskEnterpriseProjectText1**などのエンタープライズ ユーザー設定フィールドまたは参照テーブルのフィールドを**pjTaskResourceEnterpriseRBS**のようです。ローカル タスク ユーザー設定フィールド、 **pjTaskText1**などが有効です。
- *ProjectField*値は、表 1 で禁止されているフィールドのいずれかです。これらのフィールドは、禁止されているリソースのフィールドおよびエンタープライズ ユーザー設定フィールド以外は。フィールド、禁止されている他のフィールドに関連しているために禁止されていることも、既定の SharePoint タスク リストでサポートされていない値型。表 1 およびその他の禁止されているフィールドに**PjField**、1,338 定数の長いリストにするように見えますが、 **AddSiteColumn**メソッドを使用して 357?including ローカル タスクのユーザー設定の fields?that を使用することができます。表 1 です。禁止されているフィールドを追加pjTaskActivepjTaskActualOvertimeWorkpjTaskACWPpjTaskAssignmentDelaypjTaskAssignmentPeakUnitspjTaskAssignmentUnitspjTaskBaseline[1-10]BudgetCostpjTaskBaseline[1-10]BudgetWorkpjTaskBaseline[1-10]FixedCostAccrualpjTaskBaselineBudgetCostpjTaskBaselineBudgetWorkpjTaskBaselineFixedCostAccrualpjTaskBudgetCostpjTaskBudgetWorkpjTaskCalendarGuidpjTaskConstraintDatepjTaskConstraintTypepjTaskCostRateTablepjTaskDeliverableGuidpjTaskDeliverableTypepjTaskDemandedRequestedpjTaskEarnedValueMethodpjTaskEnterpriseOutlineCode[1-30]pjTaskExternalTaskpjTaskFinishSlackpjTaskFixedCostAccrualpjTaskFreeSlackpjTaskGuidpjTaskHideBarpjTaskHyperlinkpjTaskHyperlinkAddresspjTaskHyperlinkHrefpjTaskHyperlinkScreenTippjTaskHyperlinkSubAddresspjTaskIDpjTaskIgnoreWarningspjTaskIndicatorspjTaskIsAssignmentpjTaskLevelAssignmentspjTaskLevelDelaypjTaskLinkedFieldspjTaskManualpjTaskMilestonepjTaskNotespjTaskObjectspjTaskOutlineCode[1-10]pjTaskOutlineLevelpjTaskOutlineNumberpjTaskPathDrivenSuccessorpjTaskPathDrivingPredecessorpjTaskPathPredecessorpjTaskPathSuccessorpjTaskPreleveledFinishpjTaskPreleveledStartpjTaskPrioritypjTaskResourceTypepjTaskStartSlackpjTaskStatuspjTaskStatusIndicatorpjTaskSubprojectpjTaskSubprojectReadOnlypjTaskTotalSlackpjTaskTypepjTaskWarningpjTaskWorkContour




既に [タスク] リスト内に存在する、実行の値*SharePointName* parameter?although の一意の名前を使用する場合、疑問があるフィールドを追加することはできます。



## 例
**AddDurationColumns**マクロを使用して SharePoint サイトのタスク リストを作成、Project Professional でプロジェクトを作成、タスク リストをインポートするのには、 **LinkToTaskList**メソッドを使用します。、リボンの [**プロジェクト**] タブで、[**基準計画の設定**] コマンドを使用して、作業中のプロジェクトの基準計画を設定し、いくつかのタスクの期間を変更します。



**AddDurationColumns**マクロ SharePoint タスクで使用可能な列の一覧にタスクの期間と基準期間を追加する (図 1 参照) を一覧表示します。

>[!NOTE]
>**AddDurationColumns**マクロを実行した後、SharePoint タスク リストで、変更を同期するのには Project Professional でプロジェクトを保存する必要があります。



```vba
Sub AddDurationColumns()
    Dim success As Boolean
    Dim results As String
    Dim columnName As String
    Dim fieldName As PjField
    results = ""
    
    ' Add the first column.
    fieldName = pjTaskBaselineDurationText
    columnName = "Baseline duration"
    
    ' If the field name exists in the SharePoint tasks list, or fieldName
    ' is one of the prohibited fields, the AddSiteColumn method
    ' returns error 1100.
    On Error Resume Next
    
    success = AddSiteColumn(fieldName, columnName)
    
    If success Then
        results = "Added site column: " & columnName
    Else
        results = "Error in AddSiteColumn: " & columnName
    End If
    
    ' Add the second column.
    fieldName = pjTaskDurationText
    columnName = "Current duration"
    
    success = AddSiteColumn(fieldName, columnName)
    
    If success Then
        results = results & vbCrLf & "Added site column: " & columnName
    Else
        results = results & vbCrLf & "Error in AddSiteColumn: " & columnName
    End If
    
    Debug.Print results
End Sub
```




プロジェクトを保存した後は、sharepoint タスク リストに移動します。[**リスト**] タブで、[**ビューの変更**] コマンドを選択します。設定の編集ページを表示、 **AddDurationColumns**マクロを追加する、**現在の期間**フィールド**[基準期間]**フィールドを選択します。図 1 は、2 つの新しいフィールドを持つタスクの一覧を示します。

![図 1 です。同期 SharePoint タスク リストにフィールドを追加します。](eb9352be-2047-4fc8-8899-a902a71a6b11.md)




## Related Topics

[アプリケーション オブジェクト](8eb91712-7784-a102-38c0-19bb056c27e9.md)

[LinkToTaskList メソッド](65ae7bd0-446f-74dd-15fc-0a260342be90.md)

[PjField 列挙](f0df0929-921c-1f33-ab42-192efdaeb64d.md)




