

---
ms.Toctitle:FromRssFeedRuleCondition.FromRssFeed プロパティ (Outlook)(機械翻訳)
title:FromRssFeedRuleCondition.FromRssFeed プロパティ (Outlook)(機械翻訳)
ms.ContentId:f4138eaf-084c-bc18-af2a-cdbceb69e05d
---
# FromRssFeedRuleCondition.FromRssFeed プロパティ (Outlook)(機械翻訳)




ルールの条件で評価する RSS 購読を表す**文字列**の要素の配列を設定または返します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**FromRssFeed**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **FromRssFeedRuleCondition** オブジェクトを表す変数です。



## 注釈
配列の各要素は、1 つの RSS 購読となります。複数の RSS 購読は、論理 OR 条件として評価されます。



Outlook オブジェクト モデルを通じて、有効な RSS 購読名の一覧を取得することはできません。XML ファイル、フォルダーの [ドライブ] の \Documents と Settings\ の [ユーザー名] \Local Settings\Application Data\Microsoft\Outlook\ にある Outlook.Sharing.xml.obi から有効な RSS 購読名の一覧を取得することができます。< ローカル > タグの`name`属性には、 **FromRssFeed**の文字列の配列で指定する必要がある RSS 購読名が含まれています。すべての RSS 購読を列挙するには、< バインディング > タグを調べて、 `<binding prov="{0006F0AF-0000-0000-C000-000000000046}">`。



配列内の 1 つまたは複数の要素に長さ 0 の文字列 ("") または無効な RSS 購読が含まれている場合、エラーが発生します。



## Related Topics

[FromRssFeedRuleCondition オブジェクトのメンバー](0c0a949a-d654-6701-f70d-9a5bb908fed8.md)

[FromRssFeedRuleCondition オブジェクト](8de6e629-7e3d-b4df-d758-a5bff3abd6a1.md)




