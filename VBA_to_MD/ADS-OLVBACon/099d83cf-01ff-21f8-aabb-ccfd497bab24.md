

---
ms.Toctitle:AddressEntry.Update メソッド (Outlook)(機械翻訳)
title:AddressEntry.Update メソッド (Outlook)(機械翻訳)
ms.ContentId:099d83cf-01ff-21f8-aabb-ccfd497bab24
---
# AddressEntry.Update メソッド (Outlook)(機械翻訳)




メッセージング システムの **AddressEntry** オブジェクトに対する変更を反映します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Update**(**MakePermanent**, **Refresh**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **AddressEntry** オブジェクトを返すオブジェクト式を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*MakePermanent*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|**True** に設定すると、元のアドレス帳に対するすべての変更内容が反映され、プロパティのキャッシュがクリアされます。**False** に設定すると、変更内容は反映されずに、プロパティのキャッシュがクリアされます。既定値は **True** です。
|
|*Refresh*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**バリアント型 (Variant)**|**True** に設定すると、元のアドレス帳の値をプロパティのキャッシュに読み込みます。**False** に設定すると、プロパティのキャッシュに読み込みません。既定値は **False** です。
|





## 注釈
新しいアドレス項目を作成したり、既存のアドレス項目を変更したりしても、*MakePermanent* パラメーターを **True** に設定して **Update** メソッドを実行するまでは、変更内容は有効になりません。





キャッシュに格納されている内容をクリアし、アドレス帳の値をもう一度読み込むには、*MakePermanent* パラメーターを **False** に設定し、*Refresh* パラメーターを **True** に設定して **Update** メソッドを実行します。





## Related Topics

[AddressEntry Object Members](74c88069-aec4-952b-556f-03873fbb488b.md)

[AddressEntry Object](d4a0a85e-8bab-bc56-57bc-d70c3c570c8e.md)




