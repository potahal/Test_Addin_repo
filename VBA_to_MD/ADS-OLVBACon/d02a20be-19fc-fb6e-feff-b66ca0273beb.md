

---
ms.Toctitle:Items.GetLast メソッド (Outlook)(機械翻訳)
title:Items.GetLast メソッド (Outlook)(機械翻訳)
ms.ContentId:d02a20be-19fc-fb6e-feff-b66ca0273beb
---
# Items.GetLast メソッド (Outlook)(機械翻訳)




コレクションの末尾のオブジェクトを返します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetLast**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Items** オブジェクトを表す変数。

### 戻り値
コレクションに含まれている最後のオブジェクトを表す文字列型 ( **Object** ) の値を指定します。





## 注釈
**Nothing**が返されます最後のオブジェクトが存在しない場合、たとえば、コレクションが空の場合。**GetFirst**、**末尾**、 **GetNext**、および大規模なコレクションの**1 つ**のメソッドの動作が正しいことを確認、そのコレクションに**GetNext**を呼び出す前に**GetFirst**を呼び出すし、**末尾**の**1 つ**を呼び出す前に呼び出し。コレクションの呼び出しを常に行っていることを確認するには、ループに入る前に、そのコレクションを参照する明示的な変数を作成します。



## Related Topics

[アイテム オブジェクトのメンバー](bcc2cf6c-b6fb-e1a2-1d5c-d7e2bdf6b7dc.md)

[Items オブジェクト](3a99730b-e62a-5ca6-f6ec-911c95173242.md)




