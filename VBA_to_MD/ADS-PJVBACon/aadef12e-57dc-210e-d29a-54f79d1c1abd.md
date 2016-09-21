

---
ms.Toctitle:Application.ProjectBeforeResourceDelete イベント (Project)
title:Application.ProjectBeforeResourceDelete イベント (Project)
ms.ContentId:aadef12e-57dc-210e-d29a-54f79d1c1abd
---
# Application.ProjectBeforeResourceDelete イベント (Project)




リソースが削除される前に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ProjectBeforeResourceDelete**(**res**, **Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*res*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Resource**|削除されるリソースです。|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合、リソースは削除されません。|



### 戻り値
なし





## 注釈
Project のイベントは、プロジェクトが他のドキュメントまたはアプリケーションに埋め込まれている場合は発生しません。



**ProjectBeforeResourceDelete** イベントは、ユーザー設定のフォームで変更を行ったときには発生しません。




