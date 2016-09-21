

---
ms.Toctitle:MailItem.GetInspector プロパティ (Outlook)(機械翻訳)
title:MailItem.GetInspector プロパティ (Outlook)(機械翻訳)
ms.ContentId:9ba8bdbf-1dd5-eaff-3889-33433e3cb3fa
---
# MailItem.GetInspector プロパティ (Outlook)(機械翻訳)




指定したアイテムを格納するように初期化されたインスペクターを表す **Inspector** オブジェクトを取得します。値の取得のみ可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetInspector**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **MailItem** オブジェクトを表す変数。



## 注釈
このプロパティでは、アイテムを表示する **Inspector** オブジェクトを取得できます。**Application.ActiveInspector** メソッドを使用して **Inspector.CurrentItem** プロパティを設定する代わりに使用できます。アイテムの **Inspector** オブジェクトが既に存在する場合、**GetInspector** プロパティは新しいオブジェクトを作成せずに、その **Inspector** オブジェクトを取得します。





## 例
次の Visual Basic for Applications (VBA) の例では、メール アイテムを作成し、そのアイテムにタイトルを割り当て、本文のテキストを追加する関数 `InsertBodyTextInWordEditor` を示します。まず、**Subject** プロパティを使用して、タイトル "Testing..." を割り当てます。次に、**Display** メソッドを呼び出して、インスペクターでメール アイテムを開きます。Word エディターでテキストをメール アイテムの本文として挿入するために、この関数では Word オブジェクト モデルの **Document** オブジェクトと **Range** オブジェクトを使用します。アイテムの **GetInspector** プロパティを使用して既存の **Inspector** オブジェクトを取得し、**Inspector.WordEditor** プロパティを使用してアイテムの **Word.Document** オブジェクトを取得します。さらに **Word.Document** オブジェクトを使用して **Word.Range** オブジェクトにアクセスし、アイテムの本文にテキストを挿入します。



この例では Word オブジェクト モデルにアクセスするため、この例を正しくコンパイルするには、あらかじめ Microsoft Word のオブジェクト ライブラリへの参照を追加しておく必要があります。

```vba
Sub InsertBodyTextInWordEditor() 
 Dim myItem As Outlook.MailItem 
 Dim myInspector As Outlook.Inspector 
 'You must add a reference to the Microsoft Word Object Library 
 'before this sample will compile 
 Dim wdDoc As Word.Document 
 Dim wdRange As Word.Range 
 
 On Error Resume Next 
 Set myItem = Application.CreateItem(olMailItem) 
 myItem.Subject = "Testing..." 
 myItem.Display 
 'GetInspector property returns Inspector 
 Set myInspector = myItem.GetInspector 
 'Obtain the Word.Document for the Inspector 
 Set wdDoc = myInspector.WordEditor 
 If Not (wdDoc Is Nothing) Then 
 'Use the Range object to insert text 
 Set wdRange = wdDoc.Range(0, wdDoc.Characters.Count) 
 wdRange.InsertAfter ("Hello world!") 
 End If 
End Sub
```




## Related Topics

[MailItem Object Members](1094d7df-ee80-a4b0-5a21-db2979506e6b.md)

[MailItem Object](14197346-05d2-0250-fa4c-4a6b07daf25f.md)




