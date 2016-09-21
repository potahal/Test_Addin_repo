

---
ms.Toctitle:TaskRequestDeclineItem.BeforeCheckNames イベント (Outlook)(機械翻訳)
title:TaskRequestDeclineItem.BeforeCheckNames イベント (Outlook)(機械翻訳)
ms.ContentId:dd8b01bc-1368-b0ef-d0eb-b6bc955cf98f
---
# TaskRequestDeclineItem.BeforeCheckNames イベント (Outlook)(機械翻訳)





          UNRESOLVED_TOKEN_VAL(outlooknv1) がアイテム (親オブジェクトのインスタンス) の受信者コレクションの名前解決を開始する直前に発生します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**BeforeCheckNames**(**Cancel**)




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TaskRequestDeclineItem** オブジェクトを表す変数を指定します。

### パラメーター

|**名前**|**必須 / オプション**|**データ型**|**説明**|
|---|---|---|---|
|*Cancel*|
                        UNRESOLVED_TOKEN_VAL(offrequired)
                      |**ブール型 (Boolean)**|**False**イベントが発生します。イベント プロシージャでこの引数に**True**を設定する場合、名前解決プロセスは完了しません。|





## 注釈
VBScript で**BeforeCheckNames**イベントを使用するが、フォームの電子メール名が解決されると、イベントは発生しません。



このイベントは、次のような状況下では発生しません。 


- 履歴項目の書式をカスタマイズした後、[**連絡先**] フィールドで連絡先を解決した場合。

- 連絡先の書式をカスタマイズした後、[**連絡先**] フィールドで連絡先を解決した場合。

- なんらかの書式をカスタマイズした後、Outlook によってバックグラウンドで自動的に名前が解決された場合。

- プログラムを通じて受信者を作成し、解決した場合。








## Related Topics

[TaskRequestDeclineItem オブジェクト](e842c7c0-7943-9219-329b-30b892ab99b0.md)

[TaskRequestDeclineItem オブジェクトのメンバー](3de31d0d-2444-876c-5d4d-1192851301af.md)




