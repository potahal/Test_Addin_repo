

---
ms.Toctitle:Application.FileBuildID プロパティ (Project)
title:Application.FileBuildID プロパティ (Project)
ms.ContentId:6fae0673-614d-6cb2-31c2-bff9eabeecc9
---
# Application.FileBuildID プロパティ (Project)




指定したプロジェクトのファイル ビルド id 番号 (ID) を取得します。ビルド ID は、バージョンとファイルを作成したプロジェクトのアプリケーションのビルドで構成されます。読み取り専用**文字列**です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**FileBuildID**(**Name**, **UserID**, **DatabasePassWord**)





            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Name*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|
                        プロジェクト ファイル名、ソース ファイル名、またはデータ ソース名を指定します。
|
|*UserID*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**文字列型 (String)**|
                        データベースにアクセスするときに使用するユーザー ID を指定します。*Name* がデータベース以外の場合、引数 *UserID* は無視されます。
|
|*DatabasePassWord*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|
                        データベースにアクセスするときに使用するパスワードを指定します。*Name* がデータベース以外の場合、引数 *DatabasePassWord* は無視されます。
|





## 注釈
**FileBuildID**プロパティは、実際に開かずに、プロジェクト ファイルのファイル ビルド ID を取得できます。



## 例
次の例では、Test.mpp プロジェクトのビルド ID を取得します。UNRESOLVED_TOKEN_VAL(pjgenericshort)ビルド ファイルを作成したが 15.0.4027.1000 の場合は、 **FileBuildID**の値は「15,0,4027,1000」です。

```vba
Sub File_BuildID()
    Dim ProjID As String

    ProjID = Application.FileBuildID("C:\Project\VBA\Samples\Test.mpp")
    Debug.Print ProjID
End Sub
```





