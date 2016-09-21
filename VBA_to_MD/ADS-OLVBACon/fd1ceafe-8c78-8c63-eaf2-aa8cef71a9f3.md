

---
ms.Toctitle:SharingItem.Permission プロパティ (Outlook)(機械翻訳)
title:SharingItem.Permission プロパティ (Outlook)(機械翻訳)
ms.ContentId:fd1ceafe-8c78-8c63-eaf2-aa8cef71a9f3
---
# SharingItem.Permission プロパティ (Outlook)(機械翻訳)




受信者が **SharingItem** に対して与えられるアクセス権を決定する **OlPermission** クラスの定数を取得または設定します。値の取得および設定が可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Permission**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **SharingItem** オブジェクトを表す変数です。



## 注釈
**アクセス許可**プロパティは、 **SharingItem**のアクセス許可の状態を正確に反映するように**PermissionTemplateGuid**プロパティを使用して同期する必要があります。**PermissionTemplateGuid**プロパティを有効な GUID に設定する必要があります増える可能性があります**OlPermission.olPermissionTemplate**へ**のアクセス許可**プロパティを設定します。



情報権利管理 (IRM) が設定されていません (この場合、**アクセス許可**プロパティは、 **OlPermission.olUnrestricted**)、または (である場合、**アクセス許可**プロパティは、 **OlPermission.olDoNotForward**)、 **SharingItem**を転送しないように制限では、 **PermissionTemplateGuid**プロパティの値は空白にする必要があります。



IRM で保護されているコンテンツは、2007 Microsoft Office system 以降を実行中の任意のコンピューターで閲覧できますが、IRM で保護された電子メールを作成または送信するには、Microsoft Office Professional Edition 2003、Microsoft Office Outlook 2007、またはそれ以降のバージョンの Outlook が必要です。



## Related Topics

[SharingItem オブジェクトのメンバー](719ad60e-2242-2c54-778f-006b61690389.md)

[SharingItem オブジェクト](63dd3451-44f3-7cc4-c6e2-7dad5835a7d2.md)




