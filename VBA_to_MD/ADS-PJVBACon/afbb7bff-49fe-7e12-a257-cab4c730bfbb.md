

---
ms.Toctitle:Application.CustomFieldIndicators メソッド (Project)
title:Application.CustomFieldIndicators メソッド (Project)
ms.ContentId:afbb7bff-49fe-7e12-a257-cab4c730bfbb
---
# Application.CustomFieldIndicators メソッド (Project)




ユーザー設定フィールドのマークを設定します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CustomFieldIndicators**(**FieldID**, **SummaryInheritsNonsummary**, **ProjectInheritsSummary**, **ShowToolTips**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*FieldID*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**長整数型 (Long)**|ユーザー設定のフィールドを指定します。使用できる定数は、**PjCustomField** クラスの定数のいずれかです。|
|*SummaryInheritsNonsummary*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|の**true の**場合、集計行を使用して同じグラフィカル インジケーターを表示するための条件をテストし、非サマリー行と同じイメージを使用します。**False**場合は、サマリー行のマークに基づく、さまざまな条件、値の設定し、非サマリー行と異なる画像を使用します。|
|*ProjectInheritsSummary*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True**プロジェクト サマリー行は、同じテスト条件を使用して、グラフィカルなインジケーターを表示するため、同じを使用している場合に、サマリー行とをイメージします。**False**プロジェクト サマリー行のマークは、別に基づいている場合条件、値の設定し、その他のサマリー行と異なる画像を使用します。|
|*ShowToolTips*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True の**場合、ユーザー設定フィールドのマークの上でマウスを一時停止するには、ユーザー設定フィールドの実際のデータとツール ヒントが表示されます。|



### 戻り値
**ブール型 (Boolean)**






