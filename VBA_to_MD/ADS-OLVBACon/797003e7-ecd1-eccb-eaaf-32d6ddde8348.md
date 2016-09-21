

---
ms.Toctitle:Application オブジェクト (Outlook)
title:Application オブジェクト (Outlook)
ms.ContentId:797003e7-ecd1-eccb-eaaf-32d6ddde8348
---
# Application オブジェクト (Outlook)





          UNRESOLVED_TOKEN_VAL(outlooknv1) アプリケーション全体を表します。

## 備考
これは、**CreateObject** メソッドまたは組み込みの Visual Basic **GetObject** 関数を使用して返される階層内のオブジェクトのみです。



Outlook **Application** オブジェクトには次の用途があります。

- ルート オブジェクトとして、Outlook 階層内の他のオブジェクトにアクセスできるようにする。
- オブジェクト階層をスキャンすることなく、**CreateItem** を使用して、作成された新しいアイテムに直接アクセスできるようにする。
- アクティブなインターフェイス オブジェクト (エクスプローラーとインスペクター) にアクセスできるようにする。




別のアプリケーションから Outlook を操作するためにオートメーションを使用する場合、**CreateObject** メソッドを使用して Outlook **Application** オブジェクトを作成します。



## 例
次の Visual Basic for Applications (VBA) の例では、Outlook を起動して (既に実行されていない場合)、既定の受信トレイ フォルダーを開きます。

```vba
Set myNameSpace = Application.GetNameSpace("MAPI") 
 
Set myFolder= _ 
 
 myNameSpace.GetDefaultFolder(olFolderInbox) 
 
myFolder.Display
```




次の Visual Basic for Applications (VBA) の例は、**Application** オブジェクトを使用し、新しい連絡先を作成して開きます。

```vba
Set myItem = Application.CreateItem(olContactItem) 
 
myItem.Display
```




## イベント

|**名前**|
|---|
|[AdvancedSearchComplete](4f33ad44-20a3-62cd-aa1b-db74581ebb3c.md)|
|[AdvancedSearchStopped](a1a4ec9f-c0e3-6acd-b63c-89194ed70efd.md)|
|[BeforeFolderSharingDialog](e06257eb-f2d9-63cf-1220-dda55ee0ea14.md)|
|[ItemLoad](aed0656d-4e5a-550a-1116-76773215a897.md)|
|[ItemSend](54f506ea-87a2-29b9-2b33-67bc87167933.md)|
|[MAPILogonComplete](db6f7cf8-2a45-560f-f592-613de86e08e2.md)|
|[NewMail](cfc848e8-98b1-163a-c177-53993c20bb14.md)|
|[NewMailEx](3b6873a3-0ccf-0e46-1cac-0eeabb3a896b.md)|
|[OptionsPagesAdd](aa13cd97-de96-00f8-a532-ca8ee9b00343.md)|
|[Quit](ecf0b50b-db6f-7eaf-90bd-bae942bf9287.md)|
|[Reminder](f8c9fa87-3daa-58e1-7b8d-3c819cd4cab2.md)|
|[Startup](d4724d96-2572-b1e3-e202-0bfffb5cf7d5.md)|



## メソッド

|**名前**|
|---|
|[ActiveExplorer](f6dd27c0-4319-c7fc-191f-8b3b2ea319d3.md)|
|[ActiveInspector](3f2b6491-7b4b-8165-327e-b319711d5656.md)|
|[ActiveWindow](5f5b4e8b-61e4-417b-6b0c-14d1ccb41594.md)|
|[AdvancedSearch](7b433d8b-08b9-dff1-b854-287d76b47a90.md)|
|[CopyFile](dc848d48-23e0-d0a9-049d-b2ae414151d5.md)|
|[CreateItem](e5fbf367-db16-5042-823e-68e6b805e612.md)|
|[CreateItemFromTemplate](5e6c0ec4-779d-3743-afdb-606ad512ba95.md)|
|[CreateObject](09b6ff5b-a750-c07d-7499-c1f8a00214fe.md)|
|[GetNamespace](6175d0d9-5a61-ce45-35c0-b70895d757b3.md)|
|[GetObjectReference](426ade68-155b-9076-b3f8-4108f44688b0.md)|
|[IsSearchSynchronous](cd757b43-5e3f-1504-9944-7431bda6f004.md)|
|[Quit](664bc8ba-ad97-8d4f-02f9-7f9bdd04beea.md)|
|[RefreshFormRegionDefinition](35183f18-7c59-80c5-e281-af15afe39198.md)|



## プロパティ

|**名前**|
|---|
|[Application](c49cfea1-d126-75eb-fb3d-6f040526cef0.md)|
|[Assistance](14d6eb82-82ab-ea67-6a0b-103a535b8d41.md)|
|[Class](5bfb1d90-8c16-fdbe-374f-0b10d64915c3.md)|
|[COMAddIns](f911199d-dc2e-9b88-d807-a5737a39f29e.md)|
|[DefaultProfileName](53c6a189-9337-6413-72e5-bf6ea8794361.md)|
|[Explorers](bbbdbd6e-a238-8108-fbbd-5f7d7821aaa7.md)|
|[Inspectors](c2dde847-d033-90e3-30d2-62ff375d6843.md)|
|[IsTrusted](4caeb41a-9cc3-1195-22a9-ad8eae12ce53.md)|
|[LanguageSettings](8367a51a-629f-3349-fe0b-a978b2bbc9a5.md)|
|[Name](a0ac022e-4d46-fffb-aa13-f95249e30bdb.md)|
|[Parent](d83e85a0-f3d4-bf95-0568-0411a5d09350.md)|
|[PickerDialog](14acc98b-c234-d59b-d089-d6782ffb08a0.md)|
|[ProductCode](cdb4678a-fa6b-7d4f-b0b1-b34811749bf5.md)|
|[Reminders](1f5428f0-6362-a691-2fad-c80e48dce3f5.md)|
|[Session](720b2849-fe01-afb3-363c-f3bf0cd7d872.md)|
|[TimeZones](920e55d1-9914-fa74-101a-921083328d23.md)|
|[Version](08a74ab8-7e02-3956-1827-4b6690acdec1.md)|



## Related Topics

[アプリケーション オブジェクトのメンバー](3519c89c-2353-85ee-7ddc-62e5dd85a8e7.md)

[Outlook オブジェクト モデル リファレンス](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)




