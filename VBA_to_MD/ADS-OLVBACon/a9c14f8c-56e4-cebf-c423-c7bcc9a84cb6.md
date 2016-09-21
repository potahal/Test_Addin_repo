

---
ms.Toctitle:Page.CanPaste プロパティ (Outlook フォーム スクリプト)
title:Page.CanPaste プロパティ (Outlook フォーム スクリプト)
ms.ContentId:a9c14f8c-56e4-cebf-c423-c7bcc9a84cb6
---
# Page.CanPaste プロパティ (Outlook フォーム スクリプト)




**Boolean**オブジェクトがサポートするデータがクリップボードに含まれているかどうかを指定する値を返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CanPaste**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Page** オブジェクトを表す変数です。



## 注釈
**True の**場合オブジェクトは、オブジェクトがクリップボードから貼り付けられる情報を受信できない場合は、 **False** 、クリップボードから貼り付けられる情報を受け取ることができます。



**Height** プロパティは、値の取得のみ行うことができます。



オブジェクトがサポートされていない形式でクリップボードのデータがある場合、 **CanPaste**プロパティが**False**にします。たとえば、テキストのみをサポートするオブジェクトにビットマップを貼り付けようとすると、 **canpaste プロパティの値**は**False**になります。




