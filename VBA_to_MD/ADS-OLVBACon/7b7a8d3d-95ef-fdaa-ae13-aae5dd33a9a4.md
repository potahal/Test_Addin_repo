
---
ms.Toctitle:PostItem.ReadComplete イベント (Outlook)
title:PostItem.ReadComplete イベント (Outlook)
ms.ContentId:7b7a8d3d-95ef-fdaa-ae13-aae5dd33a9a4
---
# PostItem.ReadComplete イベント (Outlook)





## バージョン情報

            UNRESOLVED_TOKEN_VAL(ol15versionadded)
          



## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**ReadComplete***(Cancel)*




            UNRESOLVED_TOKEN_VAL(offexpression)PostItem**PostItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|||||
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Boolean**|(未使用の vbscript)。**False**イベントが発生します。イベント プロシージャでこの引数を**True**に設定すると、その読み取り操作は完了せず、アイテムは閲覧ウィンドウまたはインスペクターに表示されません。|





## 注釈
[BeforeRead](26a64e4e-a48e-84e8-4fea-70913a8f170f)イベントの後、アイテムの[読み取り](404c9b17-c5b6-a802-420a-f8fd279b5f9b.md)イベントの前に、 **ReadComplete**イベントが発生します。



いつアイテムをメモリからアンロードするかを決定するには、[Unload](42dea931-c3dd-b8ff-5ace-0744b17650e0.md) イベントを使用します。



**ReadComplete**イベントは、Exchange クライアント拡張機能 (ECE) イベントの**IExchExtMessageEvents::OnReadComplete**に対応します。



## Related Topics

[PostItem オブジェクト](de44065d-4e93-315a-279f-7b92f09c0465.md)

[PostItem オブジェクトのメンバー](5b150db1-c96d-0721-ec36-d5b5ebc20fd8.md)




