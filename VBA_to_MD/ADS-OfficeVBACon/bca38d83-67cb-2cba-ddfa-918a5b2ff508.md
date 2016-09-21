

---
ms.Toctitle:CommandBars.Item プロパティ (Office)
title:CommandBars.Item プロパティ (Office)
ms.ContentId:bca38d83-67cb-2cba-ddfa-918a5b2ff508
---
# CommandBars.Item プロパティ (Office)




**CommandBars**コレクションから**CommandBar**オブジェクトを取得します。読み取り専用です。

## 

>[!NOTE]
>一部の Microsoft Office アプリケーションにおける CommandBars の使用方法が、Microsoft Office Fluent ユーザー インターフェイスの新しいリボン コンポーネントによって置き換えられました。詳細については、ヘルプでキーワード "リボン" を検索してください。





## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Item**(**Index**)




            UNRESOLVED_TOKEN_VAL(offexpression) 必ず指定します。**CommandBars** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Index*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**バリアント型 (Variant)**|取得するオブジェクトの名前またはインデックス番号を指定します。|





## 例
Item はオブジェクトまたはコレクションの既定のメンバーです。たとえば、次の 2 つのステートメントの実行結果は同じで、CommandBar オブジェクトが cmdBar に割り当てられます。

```vba
Set cmdBar = CommandBars.Item("Standard") 
Set cmdBar = CommandBars("Standard")
```




>[!NOTE]
>
              UNRESOLVED_TOKEN_VAL(osdepreccommandbars)
            





## Related Topics

[CommandBars オブジェクト](0e312e21-14ee-5055-d604-b66e61c53b47.md)

[CommandBars オブジェクトのメンバー](c11db22d-b7bb-20a2-a455-e441cb8d5bc0.md)




