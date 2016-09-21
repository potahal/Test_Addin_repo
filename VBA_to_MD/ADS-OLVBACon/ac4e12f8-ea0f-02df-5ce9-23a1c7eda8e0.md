

---
ms.Toctitle:表示名を電子メール アドレスにマップします。
title:表示名を電子メール アドレスにマップします。
ms.ContentId:ac4e12f8-ea0f-02df-5ce9-23a1c7eda8e0
---
# 表示名を電子メール アドレスにマップします。




このトピックでは、表示名を取得し、現在のセッションのメッセージング システムで認識される電子メール アドレスにマップする Visual Basic for Applications (VBA) の例を示します。



各 Outlook セッションでは、トランスポート プロバイダーによって、メッセージング システムがメッセージを配信できるアドレス帳コンテナーのセットが定義されます。各アドレス帳コンテナーは、Outlook のアドレス一覧に対応します。アドレス帳コンテナーのセットに表示名が定義されている場合、この表示名は現在のセッションで解決することができ、アドレス一覧にはこの表示名にマップするエントリがあります。アドレス一覧のエントリにはさまざまな種類があり、Exchange ユーザーおよび Exchange 配布リストはその一例です。



このサンプル コードでは、`ResolveDisplayNameToSMTP` 関数で、例として表示名 "Dan Wilson" を使用します。まず、この表示名に基づいて **Recipient** オブジェクトを作成し、**Recipient.Resolve** を呼び出すことにより、この表示名がアドレス一覧に定義されているかどうかを確認します。名前が解決されると、`ResolveDisplayNameToSMTP` は、次に **Recipient** オブジェクトにマップされている **AddressEntry** オブジェクトを使用して、受信者の種類および (可能な場合は) 電子メール アドレスを取得します。

- **AddressEntry** オブジェクトの種類が Exchange ユーザーである場合、`ResolveDisplayNameToSMTP` は **AddressEntry.GetExchangeUser** を呼び出して、対応する **ExchangeUser** オブジェクトを取得します。**ExchangeUser.PrimarySmtpAddress** によって、表示名にマップされている電子メール アドレスが示されます。
- **AddressEntry** オブジェクトが Exchange 配布リストである場合、`ResolveDisplayNameToSMTP` は **AddressEntry.GetExchangeDistributionList** を呼び出して、**ExchangeDistributionList** オブジェクトを取得します。**ExchangeDistributionList.PrimarySmtpAddress** によって、表示名にマップされている電子メール アドレスが示されます。






```vba
Sub ResolveDisplayNameToSMTP() 
 Dim oRecip As Outlook.Recipient 
 Dim oEU As Outlook.ExchangeUser 
 Dim oEDL As Outlook.ExchangeDistributionList 
 
 Set oRecip = Application.Session.CreateRecipient("Dan Wilson") 
 oRecip.Resolve 
 If oRecip.Resolved Then 
 Select Case oRecip.AddressEntry.AddressEntryUserType 
 Case OlAddressEntryUserType.olExchangeUserAddressEntry 
 Set oEU = oRecip.AddressEntry.GetExchangeUser 
 If Not (oEU Is Nothing) Then 
 Debug.Print oEU.PrimarySmtpAddress 
 End If 
 Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry 
 Set oEDL = oRecip.AddressEntry.GetExchangeDistributionList 
 If Not (oEDL Is Nothing) Then 
 Debug.Print oEDL.PrimarySmtpAddress 
 End If 
 End Select 
 End If 
End Sub
```



