

---
ms.Toctitle:ContactItem.AddPicture メソッド (Outlook)(機械翻訳)
title:ContactItem.AddPicture メソッド (Outlook)(機械翻訳)
ms.ContentId:aa02c3b2-bb4c-fde9-3dbf-f871cbc200b1
---
# ContactItem.AddPicture メソッド (Outlook)(機械翻訳)




連絡先アイテムに画像を追加します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AddPicture**(**Path**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ContactItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Path*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|連絡先アイテムに追加する画像の絶対パスとファイル名を含む文字列です。|





## 注釈
連絡先アイテムに既に画像が追加されている場合は、このメソッドによって既存の画像が上書きされます。



使用できる画像は、アイコン、GIF、JPEG、BMP、TIFF、WMF、EMF、PNG の各ファイルです。UNRESOLVED_TOKEN_VAL(outlooknv1) では、必要に応じて画像のサイズ変更が自動的に実行されます。



## 例
次の Microsoft Visual Basic for Applications (VBA) の例は、連絡先名と連絡先の画像を含むファイル名の入力をユーザーに求め、連絡先アイテムに画像を追加します。連絡先アイテムの画像が既に存在する場合は、既存の画像を新しいファイルで上書きするかどうかを確認するメッセージをユーザーに表示します。

```vba
Sub AddPictureToAContact() 
 
 Dim myNms As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myContactItem As Outlook.ContactItem 
 
 Dim strName As String 
 
 Dim strPath As String 
 
 Dim strPrompt As String 
 
 
 
 Set myNms = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNms.GetDefaultFolder(olFolderContacts) 
 
 strName = InputBox("Type the name of the contact: ") 
 
 Set myContactItem = myFolder.Items(strName) 
 
 If myContactItem.HasPicture = True Then 
 
 strPrompt = MsgBox("The contact already has a picture associated with it. Do you want to overwrite the existing picture?", vbYesNo) 
 
 If strPrompt = vbNo Then 
 
 Exit Sub 
 
 End If 
 
 End If 
 
 strPath = InputBox("Type the file name for the contact: ") 
 
 myContactItem.AddPicture (strPath) 
 
 myContactItem.Save 
 
 myContactItem.Display 
 
 End Sub
```




## Related Topics

[ContactItem オブジェクトのメンバー](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)

[ContactItem オブジェクト](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)




