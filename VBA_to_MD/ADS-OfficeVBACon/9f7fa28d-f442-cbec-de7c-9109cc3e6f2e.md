

---
ms.Toctitle:SharedWorkspace.Tasks プロパティ (Office)
title:SharedWorkspace.Tasks プロパティ (Office)
ms.ContentId:9f7fa28d-f442-cbec-de7c-9109cc3e6f2e
---
# SharedWorkspace.Tasks プロパティ (Office)




現在の共有ワークスペースのタスクの一覧を表す **SharedWorkspaceTasks** コレクションを取得します。値の取得のみ可能です。

>[!NOTE]
>Microsoft Office 2010 以降、このオブジェクトまたはメンバーは推奨されていないため、使用しないでください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Tasks**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharedWorkspace** オブジェクトを表す変数を指定します。



## 例
次の使用例は、現在の共有ワークスペースのタスクの一覧を表示します。

```sourcecode
   Dim swsTasks As Office.SharedWorkspaceTasks 
    Set swsTasks = ActiveWorkbook.SharedWorkspace.Tasks 
    MsgBox "There are " & swsTasks.Count & _ 
        " task(s) in the current shared workspace.", _ 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsTasks = Nothing 

```




## Related Topics

[SharedWorkspace オブジェクト](7512f0ff-382d-d344-9424-aa10549d14f9.md)

[SharedWorkspace オブジェクトのメンバー](e4c2b518-d955-27e1-3e73-173d3c4f961d.md)




