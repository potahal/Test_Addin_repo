

---
ms.Toctitle:Conflicts.GetLast メソッド (Outlook)(機械翻訳)
title:Conflicts.GetLast メソッド (Outlook)(機械翻訳)
ms.ContentId:2f82fcab-7c8e-3df7-adc1-8f701d3bf9cb
---
# Conflicts.GetLast メソッド (Outlook)(機械翻訳)




**Conflicts** コレクションの末尾のオブジェクトを返します。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**GetLast**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **Conflicts** オブジェクトを表す変数。

### 戻り値
コレクションに格納されている末尾のオブジェクトを表す **Conflict** オブジェクト。





## 注釈
**Nothing**が返されます最後のオブジェクトが存在しない場合、たとえば、コレクションが空の場合。**GetFirst**、**末尾**、 **GetNext**、および大規模なコレクションの**1 つ**のメソッドの動作が正しいことを確認、そのコレクションに**GetNext**を呼び出す前に**GetFirst**を呼び出すし、**末尾**の**1 つ**を呼び出す前に呼び出し。コレクションの呼び出しを常に行っていることを確認するには、ループに入る前に、そのコレクションを参照する明示的な変数を作成します。



## Related Topics

[オブジェクトのメンバーの競合](dcc61922-d119-1bb9-c175-a80a73599559.md)

[オブジェクトの競合](c4e1c060-519a-a6d1-8fb2-c7dfa1e3e66f.md)




