

---
ms.Toctitle:Application.LoadWebPane イベント (Project)
title:Application.LoadWebPane イベント (Project)
ms.ContentId:b9fefabb-3d0b-9aa7-6d3b-b8fd8000571d
---
# Application.LoadWebPane イベント (Project)




プロジェクトが **[タスク影響要素]**、**[成果物]**、または **[プロジェクト/リソースのインポート ウィザード]** の Web ウィンドウをロードしたときに発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**LoadWebPane**(**Window**, **TargetPage**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Application** オブジェクトを表す変数です。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Window*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ウィンドウ**|**LoadWebBrowserControl**メソッドが呼び出された位置からウィンドウです。|
|*TargetPage*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**文字列型 (String)**|**LoadWebBrowserControl**メソッドを呼び出すために使用された同じ TargetPage パラメーター。|



### 戻り値
なし





## 注釈
Project のイベントは、プロジェクトが他のドキュメントまたはアプリケーションに埋め込まれている場合は発生しません。




