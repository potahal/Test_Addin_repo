

---
ms.Toctitle:SharedWorkspace.Files プロパティ (Office)
title:SharedWorkspace.Files プロパティ (Office)
ms.ContentId:e4a2f80e-5cb7-8ff2-3ab7-2b8c2d9d3cfb
---
# SharedWorkspace.Files プロパティ (Office)




**ワークスペース**内の**場合、スペース**のオブジェクトへのアクセスを提供します。読み取り専用です。

>[!NOTE]
>Microsoft Office 2010 以降、このオブジェクトまたはメンバーは推奨されていないため、使用しないでください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Files**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharedWorkspace** オブジェクトを表す変数を指定します。



## 例
次の使用例は、現在の共有ワークスペースに保存されているファイルの数を表示します。

```vba
    Dim swsFiles As Office.SharedWorkspaceFiles 
    Set swsFiles = ActiveWorkbook.SharedWorkspace.Files 
    MsgBox "There are " & swsFiles.Count & _ 
        " file(s) 
        vbInformation + vbOKOnly, _ 
        "Collection Information" 
    Set swsFiles = Nothing 

```




## Related Topics

[SharedWorkspace オブジェクトのメンバー](e4c2b518-d955-27e1-3e73-173d3c4f961d.md)

[SharedWorkspace オブジェクト](7512f0ff-382d-d344-9424-aa10549d14f9.md)




