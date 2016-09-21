

---
ms.Toctitle:TextRange2.Replace メソッド (Office)
title:TextRange2.Replace メソッド (Office)
ms.ContentId:e14f0ad0-3b9c-d9f5-a13d-d3bbdcae50e1
---
# TextRange2.Replace メソッド (Office)




テキスト範囲内の特定のテキストを検索します。 し、テキストが、指定した文字列では、最初に見つかったテキストを表す**TextRange2**オブジェクトを取得します。一致が見つからなかった場合は**Nothing**を返します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Replace**(**FindWhat**, **ReplaceWhat**, **After**, **MatchCase**, **WholeWords**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TextRange2** オブジェクトを表すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*FindWhat*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|検索するテキストを指定します。|
|*ReplaceWhat*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|見つけたテキストを置換するテキストを指定します。|
|*After*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|まで検索する**FindWhat**の次の出現箇所を指定したテキスト範囲内の文字の位置を指定します。たとえば、テキスト範囲の 5 番目の文字から検索する場合は**後**の「4」と指定します。この引数を省略すると、テキスト範囲の最初の文字は、検索のため、開始点として使用されます。|
|*MatchCase*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**MsoTriState**|大文字小文字を区別するかどうかを指定します。|
|*WholeWords*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**MsoTriState**|単語全体として検索するかどうかを指定します。|



### 戻り値
TextRange2





## Related Topics

[TextRange2 オブジェクト](a6a59c9b-9b64-c1e2-2e98-a1f99025c877.md)

[TextRange2 オブジェクトのメンバー](26daffff-b9ef-fd94-f5b7-ed3a09840cb6.md)




