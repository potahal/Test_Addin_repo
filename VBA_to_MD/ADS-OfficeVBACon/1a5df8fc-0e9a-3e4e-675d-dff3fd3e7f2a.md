

---
ms.Toctitle:SharedWorkspaceFolder.FolderName プロパティ (Office)
title:SharedWorkspaceFolder.FolderName プロパティ (Office)
ms.ContentId:1a5df8fc-0e9a-3e4e-675d-dff3fd3e7f2a
---
# SharedWorkspaceFolder.FolderName プロパティ (Office)




共有ワークスペースのメイン ドキュメントのライブラリ フォルダー内のサブフォルダーの名前を取得します。値の取得のみ可能です。

>[!NOTE]
>Microsoft Office 2010 以降、このオブジェクトまたはメンバーは推奨されていないため、使用しないでください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**FolderName**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharedWorkspaceFolder** オブジェクトを表す変数を指定します。



## 注釈
**FolderName**プロパティは、形式の parentfolder/サブフォルダーのサブフォルダー名を返します。たとえば、共有ワークスペースには、[サポート ドキュメント] をという名前のフォルダーが含まれている、**フォルダー名**のプロパティは共有ドキュメントとサポート ドキュメントを返します。



## 例
次の使用例は、共有ワークスペースの中にあるサブフォルダーの数と各サブフォルダーの名前を表示します。

```sourcecode
    Dim swsFolder As Office.SharedWorkspaceFolder 
    Dim strFolderInfo As String 
    strFolderInfo = "The shared workspace contains " & _ 
        ActiveWorkbook.SharedWorkspace.Folders.Count & " folder(s)." & vbCrLf 
    If ActiveWorkbook.SharedWorkspace.Folders.Count > 0 Then 
        For Each swsFolder In ActiveWorkbook.SharedWorkspace.Folders 
            strFolderInfo = strFolderInfo & swsFolder.FolderName & vbCrLf 
        Next 
    End If 
    MsgBox strFolderInfo, vbInformation + vbOKOnly, _ 
        "Folders in Shared Workspace" 
    Set swsFolder = Nothing 

```




## Related Topics

[SharedWorkspaceFolder オブジェクト](297c4ed7-2232-5240-ca34-d374038c66a2.md)

[SharedWorkspaceFolder オブジェクトのメンバー](e7e0a32a-ce01-e08f-f251-27d93273110e.md)




