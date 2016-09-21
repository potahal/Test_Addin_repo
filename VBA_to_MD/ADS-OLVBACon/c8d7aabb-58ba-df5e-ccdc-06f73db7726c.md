

---
ms.Toctitle:NavigationFolder オブジェクト (Outlook)(機械翻訳)
title:NavigationFolder オブジェクト (Outlook)(機械翻訳)
ms.ContentId:c8d7aabb-58ba-df5e-ccdc-06f73db7726c
---
# NavigationFolder オブジェクト (Outlook)(機械翻訳)




ナビゲーション ウィンドウのナビゲーション モジュールのナビゲーション グループに表示されるナビゲーション フォルダーを表します。

## 注釈
親**NavigationGroup**オブジェクトの**NavigationFolders**コレクションから**NavigationFolder**オブジェクトを取得するのにには、 **Item**メソッドを使用します。**NavigationFolders**コレクションの**Add**メソッドを使用すると、既存の**Folder**オブジェクトに基づいて新しい**NavigationFolder**オブジェクトを作成できます。



**NavigationFolder**オブジェクトの基になる**Folder**オブジェクトを設定するには、**フォルダー**のメソッドを使用します。



ナビゲーション フォルダーが選択されているかどうかを確認するには、**IsSelected** プロパティを使用し、ナビゲーション ウィンドウ内のナビゲーション フォルダーの表示位置を取得または設定するには、**Position** プロパティを使用します。**DisplayName** プロパティを使用して、ナビゲーション ウィンドウ内のナビゲーション フォルダーの表示名を取得することもできます。



**NavigationFolders**コレクションから、 **IsSideBySide**プロパティを**CalendarModule**オブジェクトに関連付けられているナビゲーション フォルダーの表示モードを設定するのにナビゲーション フォルダーを削除することができるかどうかを決定するのにには、 **IsRemovable**プロパティを使用します。



## Related Topics

[NavigationFolder オブジェクトのメンバー](1ec2e16d-c7ca-86b1-9283-839a2b9aca05.md)

[Outlook オブジェクト モデル リファレンス](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)




