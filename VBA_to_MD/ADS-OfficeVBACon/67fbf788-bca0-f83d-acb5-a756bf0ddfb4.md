

---
ms.Toctitle:SharedWorkspace.CreateNew メソッド (Office)
title:SharedWorkspace.CreateNew メソッド (Office)
ms.ContentId:67fbf788-bca0-f83d-acb5-a756bf0ddfb4
---
# SharedWorkspace.CreateNew メソッド (Office)




サーバーにドキュメント ワークスペース サイトを作成し、新しい共有ワークスペース サイトにアクティブ ドキュメントを追加します。

>[!NOTE]
>Microsoft Office 2010 以降、このオブジェクトまたはメンバーは推奨されていないため、使用しないでください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CreateNew**(**URL**, **Name**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharedWorkspace** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*URL*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|新しい共有ワークスペースの作成先とする親フォルダーの URL を指定します。URL の指定を省略した場合、サイトは、ユーザーの既定のサーバー フォルダーに作成されます。|
|*Name*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|新しい共有ワークスペース サイトの名前を指定します。既定値は、アクティブ ドキュメントの名前からファイル名拡張子が除外された名前です。たとえば、"Budget.xls" のワークスペース サイトを作成すると、新しいサイトの名前は "Budget" になります。|





## 注釈
作業中の文書の共有ワークスペース サイトを作成するのにには、 **CreateNew**メソッドを使用します。ユーザーの既定のサーバーの場所で作業中の文書の名前を使用してサイトを作成する 2 つの省略可能な引数を省略します。



**CreateNew**メソッドでは、作業中の文書に保存されていない変更がある場合にエラーが発生します。共有ワークスペース サイトに追加する前に、ドキュメントを保存する必要があります。

>[!NOTE]
>共有ワークスペース サイトを作成して、サイトで作業中の文書を作成し、直後に作業中の文書がひんぱんに開閉し、ユーザーに表示される作業中の文書のコピーは、サイトにあるようにします。場合**CreateNew**メソッドを呼び出す前に作業中の文書を保存すると、そのドキュメントのコピーは使用できません時間の期間の新しいコピーの作成中にします。これにより、期間の作成中に保存されたコピーにアクセスしようとする任意のコードの例外が発生します。1 つの回避策では、任意のスクリプトからアクティブ ドキュメントにアクセスする前に、わずかな遅延 (推奨される 15 秒以上) を適用します。さらに、ローカル ドキュメントを指すすべてのキャッシュされたオブジェクトは、共有ワークスペース サイトのドキュメントを指すように更新します。





## 例
次の使用例 URL http://server/sites/mysite/に共有ワークスペース サイトを作成するには、ワークスペースの名前を「マイ共有予算のドキュメント」、およびサイトにアクティブ ドキュメントを追加します。新しい共有ワークスペース サイトの**URL**プロパティを返します http://server/sites/mysite/My%20Shared%20Budget%20Document/、 **Name**プロパティ「共有予算文書、 **SharedWorkspaceFiles**コレクションの**Count**プロパティが 1 つファイルを表示し、。

```vba
   Dim sws As Office.SharedWorkspace 
    Dim strSWSInfo As String 
    Set sws = ActiveWorkbook.SharedWorkspace 
    sws.CreateNew "http://server/sites/mysite/", "My Shared Budget Document" 
    strSWSInfo = "Name: " & sws.Name & vbCrLf & _ 
        "URL: " & sws.URL & vbCrLf & _ 
        "File(s): " & sws.Files.Count 
    MsgBox strSWSInfo, vbInformation + vbOKOnly, _ 
        "New Shared Workspace Information" 
    Set sws = Nothing 

```




## Related Topics

[SharedWorkspace オブジェクト](7512f0ff-382d-d344-9424-aa10549d14f9.md)

[SharedWorkspace オブジェクトのメンバー](e4c2b518-d955-27e1-3e73-173d3c4f961d.md)




