

---
ms.Toctitle:JournalItem.AutoResolvedWinner プロパティ (Outlook)(機械翻訳)
title:JournalItem.AutoResolvedWinner プロパティ (Outlook)(機械翻訳)
ms.ContentId:9a0fa7a4-aeeb-c616-47cc-cc28dfdaa35c
---
# JournalItem.AutoResolvedWinner プロパティ (Outlook)(機械翻訳)




**ブール値**アイテムが自動競合解決の勝者であるかどうかを返します。読み取り専用です。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**AutoResolvedWinner**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **JournalItem** オブジェクトを表す変数を指定します。



## 注釈
値が**False**の場合、必ずしもアイテムの競合が自動的に解決されなかったであります。アイテムは、別のアイテムと競合していない可能性があります。



アイテムの**JournalItem.Conflicts**プロパティを 0 より大きいの**Conflicts.Count**の場合**が必要**] プロパティが**True**の場合は、競合の自動解決の勝者を勧めします。その一方で、アイテムが競合している**が必要**プロパティは**False**として、自動競合が解決で優先されなかったデータです。



## Related Topics

[JournalItem オブジェクト](6e850295-39f9-47b8-e866-9622e9958c69.md)

[JournalItem オブジェクトのメンバー](13a0cd10-44bc-a167-c613-93985f698d95.md)




