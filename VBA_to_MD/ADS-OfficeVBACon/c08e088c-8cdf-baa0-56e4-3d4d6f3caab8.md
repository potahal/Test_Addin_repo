

---
ms.Toctitle:Permission.StoreLicenses プロパティ (Office)
title:Permission.StoreLicenses プロパティ (Office)
ms.ContentId:c08e088c-8cdf-baa0-56e4-3d4d6f3caab8
---
# Permission.StoreLicenses プロパティ (Office)




取得または、作業中の文書を表示するのにはユーザーのライセンスをユーザーがアクセス権の管理サーバーに接続できないときにオフラインで表示できるようにするのにはキャッシュかどうかを示す**ブール**値を設定します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**StoreLicenses**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Permission** オブジェクトを表す変数です。



## 注釈
**StoreLicenses**プロパティに対応 (とその値は、反対の) アクセス許可のユーザー インターフェイスで**ユーザーの権限を確認するのには接続を必要とする**オプションです。**StoreLicenses**の**偽**のユーザー以外のドキュメントの所有者する必要があります、アクセス権管理サーバーへの接続し、使用可能なドキュメントを開くたびにコンテンツが保護されている Microsoft Office で提供される情報の権限の管理サービスを使用してライセンスを取得します。



## 例
次の使用例は、 **StoreLicenses**の設定など、作業中の文書のアクセス権の設定に関する情報を表示します。

```sourcecode
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." & vbCrLf 
 strIRMInfo = strIRMInfo & " View in trusted browser: " & _ 
 irmPermission.EnableTrustedBrowser & vbCrLf & _ 
 " Document author: " & irmPermission.DocumentAuthor & vbCrLf & _ 
 " Users with permissions: " & irmPermission.Count & vbCrLf & _ 
 " Cache licenses locally: " & irmPermission.StoreLicenses & vbCrLf & _ 
 " Request permission URL: " & irmPermission.RequestPermissionURL & vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo & " Permissions applied from policy:" & vbCrLf & _ 
 " Policy name: " & irmPermission.PolicyName & vbCrLf & _ 
 " Policy description: " & irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo & " Custom permissions applied." 
 End If 
 Else 
 strIRMInfo = "Permissions are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```




## Related Topics

[アクセス許可オブジェクトのメンバー](75614d24-cd47-ef9b-aba5-112206daa358.md)

[アクセス許可オブジェクト](4bdf7058-d4ba-0bd4-c5cd-141d67245ced.md)




