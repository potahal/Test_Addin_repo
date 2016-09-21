

---
ms.Toctitle:Project.ProjectNotes プロパティ (Project)
title:Project.ProjectNotes プロパティ (Project)
ms.ContentId:2a9dcdbe-50f2-544a-8aba-c2db0d6762bc
---
# Project.ProjectNotes プロパティ (Project)




取得またはプロジェクトのメモを設定します。読み取りまたは書き込み**文字列**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ProjectNotes**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Project** オブジェクトを表す変数です。



## 注釈

            UNRESOLVED_TOKEN_VAL(pjgenericshort)で、プロジェクト**のプロパティ**] ダイアログ ボックスを表示するには、 **Backstage**ビューを表示します、[**情報**] タブを選択し、**プロジェクト情報**] ドロップダウン メニューで**[詳細プロパティ**を選択して、リボンの [**ファイル**] タブを選択します。



## 例
次の使用例は、プロジェクトの [**プロパティ**] ダイアログ ボックスの [**コメント**] フィールドに日時を追加して、そのプロジェクトを保存します。

```vba
Sub SaveAndNoteTime() 
    Projects(1).ProjectNotes = Projects(1).ProjectNotes & vbCrLf _ 
        & "This project was last saved on " _ 
        & Date$ & " at " & Time$ & "." 
    FileSave 
End Sub
```





