

---
ms.Toctitle:Application.CustomFieldIndicatorDelete メソッド (Project)
title:Application.CustomFieldIndicatorDelete メソッド (Project)
ms.ContentId:729eafe9-4d1a-07a6-efbc-ab0c94e3af59
---
# Application.CustomFieldIndicatorDelete メソッド (Project)




ユーザー設定フィールドのマークの条件一覧から条件を消去します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**CustomFieldIndicatorDelete**(**FieldID**, **Index**, **CriteriaList**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*FieldID*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**長整数型 (Long)**|ユーザー設定のフィールドを指定します。使用できる定数は、**PjCustomField** クラスの定数のいずれかです。|
|*Index*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**整数型 (Integer)**|**CriteriaList**で指定されたリストから削除するのにはテスト条件の位置を指定します。|
|*CriteriaList*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**長整数型 (Long)**|削除するテスト条件を含む条件一覧をします。**PjCriteriaList**定数は、次のいずれか: **pjCriteriaNonSummary****pjCriteriaSummary**、 **pjCriteriaProjectSummary**。既定値は**pjCriteriaNonSummary**です。|



### 戻り値
**ブール型 (Boolean)**





## 注釈
**CustomFieldIndicatorDelete**メソッドは、別のリストから値を継承するように設定されているため*CriteriaList*で指定されたリストが読み取り専用の場合、トラップ可能なエラー (エラー コード 1004年) を返します。




