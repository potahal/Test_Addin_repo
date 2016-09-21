

---
ms.Toctitle:SharedWorkspaceTask.Title プロパティ (Office)
title:SharedWorkspaceTask.Title プロパティ (Office)
ms.ContentId:038d24fe-5afa-c61d-16e7-7a8c8fca2ccf
---
# SharedWorkspaceTask.Title プロパティ (Office)




**SharedWorkspaceTask**オブジェクトのタイトルを取得または設定します。読み取り/書き込み。

>[!NOTE]
>Microsoft Office 2010 以降、このオブジェクトまたはメンバーは推奨されていないため、使用しないでください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Title**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharedWorkspaceTask** オブジェクトを表す変数です。

### 戻り値
文字列型 (String)





## 注釈
**Title**プロパティは、共有ワークスペースのタスクの 1 つの必須プロパティです。提供したり、タスクに関する追加情報を取得するオプションの**説明**のプロパティを使用します。



## 例
次の例は、現在の共有ワークスペースのすべてのタスクのタイトル一覧を表示します。



```sourcecode
 Dim swsTask As Office.SharedWorkspaceTask 
    Dim strTasks As String 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        strTasks = strTasks & swsTask.Title & vbCrLf 
    Next 
    MsgBox strTasks, vbInformation + vbOKOnly, _ 
        "Tasks in Shared Workspace" 
    Set swsTask = Nothing 
 

```




## Related Topics

[SharedWorkspaceTask オブジェクト](fbd82b03-53fa-12ff-9fb2-07bef012dde8.md)

[SharedWorkspaceTask オブジェクトのメンバー](5b5589d1-f907-7357-f930-eede569d2021.md)




