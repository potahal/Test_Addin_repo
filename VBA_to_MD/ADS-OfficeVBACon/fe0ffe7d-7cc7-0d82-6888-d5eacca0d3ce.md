

---
ms.Toctitle:SignatureInfo オブジェクト (Office)
title:SignatureInfo オブジェクト (Office)
ms.ContentId:fe0ffe7d-7cc7-0d82-6888-d5eacca0d3ce
---
# SignatureInfo オブジェクト (Office)




デジタル署名やドキュメント内署名を作成するための情報を表します。

## 例
次の例では、デジタル証明書の有効期限の日付を取得するのには、 **SignatureInfo**オブジェクトの**GetCertificationDetails**メソッドを使用します。

```vba
Sub GetCertDetails() 
Dim objSignatureInfo As SignatureInfo 
Dim varDetail As Variant 
 
strDetail = objSignatureInfo.GetCertificationDetail(certdetExpirationDate) 
 
End Sub 

```




## Related Topics

[オブジェクト モデル リファレンス](499c789a-aba2-0fad-649a-0ea964cd3b5e.md)

[SignatureInfo オブジェクトのメンバー](52c19097-8afb-d35c-a9f7-eae81e91c05d.md)




