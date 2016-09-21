

---
ms.Toctitle:Store.GetSearchFolders メソッド (Outlook)(機械翻訳)
title:Store.GetSearchFolders メソッド (Outlook)(機械翻訳)
ms.ContentId:aed6ba0b-5e20-adb9-6f62-d030a0de2e0b
---
# Store.GetSearchFolders メソッド (Outlook)(機械翻訳)




**Store** オブジェクトに定義されている検索フォルダーを表す **Folders** コレクション オブジェクトを返します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetSearchFolders**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Store** オブジェクトを表す変数。

### 戻り値
**ストア**オブジェクトのすべての検索フォルダーを表す**Folders**コレクション オブジェクトです。





## 注釈
**GetSearchFolders**は、**ストア**のすべての表示されているアクティブな検索フォルダーを取得します。初期化されていないか、古い検索フォルダーは返しません。



**GetSearchFolders****フォルダー**コレクションのオブジェクトを返します**Folders.Count**等しいゼロ (0)**ストア**の検索フォルダーが定義されていない場合。



検索フォルダーのコレクションを表す**フォルダー**コレクション オブジェクトの場合は、 **Folders.Parent**は、 **Store.GetRootFolder**と同じオブジェクトを返します。**Folder.Folders**は、 **Null** (**Nothing**で Visual Basic) を返します。



## 例
次に示す Microsoft Visual Basic for Applications (VBA) のコードは、現在のセッションのすべてのストアの検索フォルダーを列挙するコード例です。


```vba
Sub EnumerateSearchFoldersInStores() 
 
 Dim colStores As Outlook.Stores 
 
 Dim oStore As Outlook.Store 
 
 Dim oSearchFolders As Outlook.folders 
 
 Dim oFolder As Outlook.Folder 
 
 
 
 On Error Resume Next 
 
 Set colStores = Application.Session.Stores 
 
 For Each oStore In colStores 
 
 Set oSearchFolders = oStore.GetSearchFolders 
 
 For Each oFolder In oSearchFolders 
 
 Debug.Print (oFolder.FolderPath) 
 
 Next 
 
 Next 
 
End Sub
```




## Related Topics

[ストア オブジェクト](1eb22fe9-8849-7476-5388-2515b48591b9.md)

[ストア オブジェクトのメンバー](84c1d423-e507-0b3b-6570-33829b94be04.md)




