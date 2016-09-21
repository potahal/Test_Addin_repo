

---
ms.Toctitle:SharedWorkspaceFiles.Creator プロパティ (Office)
title:SharedWorkspaceFiles.Creator プロパティ (Office)
ms.ContentId:e5d5aff6-e658-daff-da59-c7e74b727eba
---
# SharedWorkspaceFiles.Creator プロパティ (Office)




**SharedWorkspaceFiles**オブジェクトが作成されたアプリケーションを示す 32 ビット整数を取得します。読み取り専用です。

>[!NOTE]
>Microsoft Office 2010 以降、このオブジェクトまたはメンバーは推奨されていないため、使用しないでください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Creator**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharedWorkspaceFiles** オブジェクトを表す変数を指定します。

### 戻り値
長整数型 (Long)





## 注釈
例として、オブジェクトが Microsoft Word で作成された場合を返します 1297307460 を表す文字列"MSWD"です。Microsoft excel では 1480803660 を返します。この値は、Word では、定数 wdCreatorCode または Excel で xlCreatorCode によっても表すことができます。**Creator**プロパティは、各アプリケーションが 4 文字のクリエーター コードを持つ Macintosh で使用するために設計されました。たとえば、Word には作成者コード MSWD です。このプロパティの詳細については、Microsoft Office Macintosh Edition に含まれているヘルプの言語リファレンスを参照してください。



**Application**プロパティが常にアクティブなアプリケーションの名前を返します文字列形式のと同様、 **Creator**プロパティは常にアクティブなアプリケーションの数値識別子を返します。オブジェクトを作成したユーザーの名前を取得するのに**場合、スペース**のオブジェクトの**CreatedBy**プロパティを使用します。Office ドキュメントの作成者についての情報を取得するのにには、ドキュメントのプロパティを使用します。



## Related Topics

[SharedWorkspaceFiles オブジェクト](5e2937f7-f794-dffb-a1ec-69ea9a9e3546.md)

[SharedWorkspaceFiles オブジェクトのメンバー](30e841ce-c8f1-249a-3bc7-6f204be64536.md)




