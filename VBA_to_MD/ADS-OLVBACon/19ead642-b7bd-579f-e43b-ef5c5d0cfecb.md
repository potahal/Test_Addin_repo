

---
ms.Toctitle:MailItem.Display メソッド (Outlook)
title:MailItem.Display メソッド (Outlook)
ms.ContentId:19ead642-b7bd-579f-e43b-ef5c5d0cfecb
---
# MailItem.Display メソッド (Outlook)




現在のアイテムの新しい **Inspector** オブジェクトを表示します。

## 構文
UNRESOLVED_TOKEN_VAL(offexpression).**Display**(**Modal**)



UNRESOLVED_TOKEN_VAL(offexpression)**MailItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Modal*|UNRESOLVED_TOKEN_VAL(offoptional)|**バリアント型 (Variant)**|**True** を指定すると、ウィンドウがモーダルになります。既定値は **False** です。|





## 注釈
**Display** メソッドは、以前のバージョンとの互換性を保つために、エクスプローラー ウィンドウおよびインスペクター ウィンドウをサポートしています。エクスプローラー ウィンドウまたはインスペクター ウィンドウをアクティブにするには、**Activate** メソッドを使います。



UNRESOLVED_TOKEN_VAL(outlooknv1) オブジェクト モデルを使用して "安全でない" ファイル システムのオブジェクト (フォルダーに直接投稿されたファイル) を開こうとすると、プログラムが C または C++ で書かれている場合、リターン コード **E_FAIL** が返されます。Outlook 2000 およびそれ以前のバージョンでは、**Display** メソッドを使用して、"安全でない" ファイル システムのオブジェクトを開くことができます。



## 例
次の Visual Basic for Applications の例は、**受信トレイ** フォルダーの先頭のアイテムを表示します。この例では、アイテムを特定しているため、**受信トレイ**にアイテムが存在しないとエラーが発生します。フォルダーにアイテムがない場合は、メッセージ ボックスが表示されます。

>[!NOTE]
>**Items** コレクション オブジェクトのアイテムが特定の順番で並べられているとは限りません。



```vba
Sub DisplayFirstItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 On Error GoTo ErrorHandler 
 
 myFolder.Items(1).Display 
 
 Exit Sub 
 
ErrorHandler: 
 
 MsgBox "There are no items to display." 
 
End Sub
```




## Related Topics

[MailItem オブジェクト メンバー](1094d7df-ee80-a4b0-5a21-db2979506e6b.md)

[MailItem オブジェクト](14197346-05d2-0250-fa4c-4a6b07daf25f.md)




