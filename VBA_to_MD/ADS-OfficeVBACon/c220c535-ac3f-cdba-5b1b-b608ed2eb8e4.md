

---
ms.Toctitle:CustomXMLPart.SelectNodes メソッド (Office)
title:CustomXMLPart.SelectNodes メソッド (Office)
ms.ContentId:c220c535-ac3f-cdba-5b1b-b608ed2eb8e4
---
# CustomXMLPart.SelectNodes メソッド (Office)




カスタム XML 部分からノードのコレクションを選択します。

>[!NOTE]
>
              UNRESOLVED_TOKEN_VAL(osoffdevnodtdsincustomxml)
            





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**SelectNodes**(**XPath**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **CustomXMLPart** オブジェクトを表すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*XPath*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|XPath 式を指定します。|



### 戻り値
CustomXMLNodes





## 例
次の例では、カスタム XML 部分を追加し、名前空間 URI に一致する部分を選択し、その部分から XPath 式に一致するノードを選択します。

```vba
Dim cxp1 As CustomXMLPart 
Dim cxn As CustomXMLNode 
 
' Add a custom xml part. 
ActiveDocument.CustomXMLParts.Add "<supplier>" 
 
' Return the first custom xml part with the given namespace. 
Set cxp1 = ActiveDocument.CustomXMLParts("urn:invoice:namespace")  
 
' Get all of the nodes matching an XPath expression. 
 Set cxns = cxp1.SelectNodes("//*[@unitPrice > 20]") 

```




## Related Topics

[CustomXMLPart オブジェクト](a4f90bac-01d6-bba4-f64b-a64e2b122cfd.md)

[CustomXMLPart オブジェクトのメンバー](76fe85f4-5a35-7d12-2989-6f17a094dcdf.md)




