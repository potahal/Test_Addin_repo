

---
ms.Toctitle:Stores.StoreAdd イベント (Outlook)(機械翻訳)
title:Stores.StoreAdd イベント (Outlook)(機械翻訳)
ms.ContentId:26e7eddc-9c5a-ffff-d574-afa48e5953d8
---
# Stores.StoreAdd イベント (Outlook)(機械翻訳)




プログラムまたはユーザーの操作によって、現在のセッションに **Store** が追加されると発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**StoreAdd**(**Store**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Stores** オブジェクトを表す変数。

### パラメーター

|**名前**|**必須/オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Store*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**Store**|現在のセッションに追加する**ストア**。|





## 注釈
このイベントを発生させるには、Outlook が実行されている必要があります。このイベントは、次のいずれかに該当する場合に発生します。



- [**ファイル**] メニューの [**開く**] をポイントし、[**Outlook データ ファイル**] をクリックして、[**Outlook データ ファイルを開く**] ダイアログ ボックスを通じてストアを追加したとき。
- [**アカウント マネージャー**] ダイアログ ボックスの [**データ ファイル**] タブを通じてストアを追加したとき。
- **Namespace.AddStore** メソッドの呼び出しによってストアが正常に追加されたとき。








このイベントは、次のいずれかに該当する場合は発生しません。


- Outlook が起動し、プライマリ ストアまたは代理ストアが開かれるとき。
- Outlook が実行中でない場合に、Microsoft Windows のコントロール パネルの [**メール**] アプレットを通じてストアが追加されたとき。
- [**Microsoft Exchange Server**] ダイアログ ボックスの [**詳細設定**] タブを通じて代理ストアが追加されたとき。









このイベントを使用すると、ストアが追加されたことを検出し、そのストア内のアイテムに対して適切な操作を実行できます。イベントを使用しない場合は、**Stores** コレクションをポーリングする必要があります。



## Related Topics

[ストア オブジェクト](8915a8e4-9c22-21d5-c492-051d393ce5f7.md)

[ストア オブジェクトのメンバー](f3fec99a-54b2-c13e-d96a-c8c5e2429f99.md)




