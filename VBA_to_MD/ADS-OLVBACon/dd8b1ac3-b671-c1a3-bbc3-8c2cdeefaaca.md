

---
ms.Toctitle:ContactItem.CompanyLastFirstNoSpace プロパティ (Outlook)(機械翻訳)
title:ContactItem.CompanyLastFirstNoSpace プロパティ (Outlook)(機械翻訳)
ms.ContentId:dd8b1ac3-b671-c1a3-bbc3-8c2cdeefaaca
---
# ContactItem.CompanyLastFirstNoSpace プロパティ (Outlook)(機械翻訳)




後に姓、名、およびスペースを入れずに、ミドル ネーム、姓と名の間で連絡先の会社名を表す**文字列**を返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CompanyLastFirstNoSpace**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **ContactItem** オブジェクトを表す変数を指定します。



## 注釈
このプロパティは、 **[得意先名]**、 **[氏名]****[部署名]**、および**ミドル ネーム**のプロパティから解析されます。**姓****姓**、および**ミドル ネーム**のプロパティは、 **FullName**プロパティから解析自体です。場合のみこのプロパティの値で入力されて、関連付けられているプロパティ (**姓**、**姓**、**ミドル ネーム****[得意先名]**、および**サフィックス**) には、アジア言語の (DBCS) 文字が含まれています。対応するフィールドにアジア言語の文字が含まれていない場合、プロパティは空になります。



## Related Topics

[ContactItem オブジェクト](8e32093c-a678-f1fd-3f35-c2d8994d166f.md)

[ContactItem オブジェクトのメンバー](a8b13369-4c87-02aa-e62a-1f3067e559fa.md)




