

---
ms.Toctitle:Application.DrawingMove メソッド (Project)
title:Application.DrawingMove メソッド (Project)
ms.ContentId:0d6e2b43-a9ab-1e9d-ad89-afa01afddb50
---
# Application.DrawingMove メソッド (Project)




アクティブな図形の積み重ね順序を前方または後方に移動します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**DrawingMove**(**Forward**, **Full**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Forward*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|**True の**場合は、アクティブな描画オブジェクトは描画レイヤーで前方に移動します。既定値は、 **false を指定**します。|
|*Full*|
                        UNRESOLVED_TOKEN_VAL(offoptional)
                      |**ブール型 (Boolean)**|の**場合は true**アクティブな描画オブジェクトが**転送用**に指定した方向の範囲全体を移動する場合。場合は**false**オブジェクトが 1 つのレイヤーを移動します。既定値は、 **false を指定**します。|



### 戻り値
**ブール型 (Boolean)**






