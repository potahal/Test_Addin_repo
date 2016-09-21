

---
ms.Toctitle:SharedWorkspaceLink オブジェクト (Office)
title:SharedWorkspaceLink オブジェクト (Office)
ms.ContentId:eb36dbed-fc41-08df-3cbc-affbaf5f9784
---
# SharedWorkspaceLink オブジェクト (Office)




共有ドキュメント ワークスペース サイトに保存されている URL リンクを表します。

>[!NOTE]
>Microsoft Office 2010 以降、このオブジェクトまたはメンバーは推奨されていないため、使用しないでください。





## 注釈
**SharedWorkspaceLink**オブジェクトを使用すると、追加のドキュメントとドキュメント ワークスペース サイトのドキュメントで共同作業しているメンバーの関心のある情報へのリンクを管理できます。



特定の**SharedWorkspaceLink**オブジェクトを取得するのにには、 **SharedWorkspaceLinks**コレクションの**項目**(*インデックス*) のプロパティを使用します。



[**共有ワークスペース**] ウィンドウの [**リンク**] タブおよびワークスペースの Web ページ上に表示されるリンクの説明を設定するのには、[**説明**] プロパティを使用します。リンクの宛先アドレスを設定するのにには、 **URL**プロパティを使用します。**「メモ**」プロパティを使用して、リンクに関する追加情報を提供します。



**SharedWorkspaceLink**オブジェクトのプロパティを変更した後、変更をサーバーにアップロードするのにには、 **Save**メソッドを使用します。



**CreatedBy**、 **CreatedDate**、**こうした**、 **ModifiedDate**プロパティを使用して、各リンクの履歴に関する情報を返します。



## 例
次の使用例は、共有ワークスペース サイトの最初のリンク先を Microsoft Developer Network のホーム ページに変更し、この変更内容をサーバーにアップロードします。

```sourcecode
    Dim swsLink As Office.SharedWorkspaceLink 
    Set swsLink = ActiveWorkbook.SharedWorkspace.Links(1) 
    With swsLink 
        .Description = "MSDN Home Page" 
        .URL = "http://msdn.microsoft.com/" 
        .Notes = "My favorite site for developers!" 
        .Save 
    End With 
    Set swsLink = Nothing 

```




## Related Topics

[オブジェクト モデル リファレンス](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)

[SharedWorkspaceLink オブジェクトのメンバー](fa8d7312-77cc-77b7-14ca-a6aa7f63fa7b.md)




