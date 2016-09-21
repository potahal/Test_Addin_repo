

---
ms.Toctitle:IRibbonControl.Tag プロパティ (Office)
title:IRibbonControl.Tag プロパティ (Office)
ms.ContentId:d0f041c0-d7bc-7a4f-df9b-ba62fa08f1ca
---
# IRibbonControl.Tag プロパティ (Office)




このプロパティを使って任意の文字列を格納し、実行時にその文字列を取り出します。値の取得のみ可能です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**Tag**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **IRibbonControl** オブジェクトを返す式。

### 戻り値
文字列型 (String)





## 注釈
通常、 **Id**プロパティを使用してリボン ユーザー インターフェイス XML カスタマイズ ファイル内のコントロールの間で区別できます。ただし、Id に含めることができます上の制限がある (英数字以外の文字を含まないし、それらすべて一意でなければなりません)。



**Tag**プロパティは、これらの制限を持っていないし、ID が動作しない、次の状況で使用するため。

- ファイル名などのコントロールに特殊文字を格納する必要がある場合。例 : tag=”C:\path\file.xlsm”
- 場合、コールバック プロシージャで同じ方法で処理する複数のコントロールが、Id (一意である必要があります) のすべてのリストを保持したくないです。 などの可能性があります、リボンの異なるタブにボタンがある、すべてタグ ="blue"、および**ID**プロパティではなく**Tag**プロパティをチェックするだけと、追加いくつかの一般的な操作です。








## 例
リボン ユーザー インターフェイスをカスタマイズするために使用する XML を次のようにタグを設定できます。MyFunction 処理が呼び出されると、「何らかの文字列」なりますが、 **Tag**プロパティを読み取ることができます。

```xml
<button id=”mybutton” tag=”some string” onAction=”MyFunction”/>
```




## Related Topics

[IRibbonControl オブジェクト](63aef709-e1d3-b1a6-76af-b568ad0e69ae.md)

[IRibbonControl オブジェクトのメンバー](396d85dc-ddd5-8985-0830-22ee5b1579dc.md)




