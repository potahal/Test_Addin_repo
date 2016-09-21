

---
ms.Toctitle:取得した AddressRuleCondition オブジェクト (Outlook)(機械翻訳)
title:取得した AddressRuleCondition オブジェクト (Outlook)(機械翻訳)
ms.ContentId:8cf897ad-a8f9-67ea-c0fa-d7f4bb917bd4
---
# 取得した AddressRuleCondition オブジェクト (Outlook)(機械翻訳)




メッセージの受信者または送信者のアドレスが **AddressRuleCondition.Address** で指定されたアドレスに含まれているかどうかを評価するルールの条件を表します。

## 注釈
**取得した AddressRuleCondition**は、**取得した RuleCondition**オブジェクトから派生します。各ルールは、 **RecipientAddress**プロパティと**SenderAddress**を持つ**RuleConditions**オブジェクトに関連付けられています。これらの各プロパティは常に**取得した AddressRuleCondition**オブジェクトを返します。**AddressRuleCondition.ConditionType**は、これらのルールの条件の間で区別します。ルールは、ルールの条件が有効になっているこれらのいずれかがある場合、 **AddressRuleCondition.Enabled**は**真**でしょう。



ルールの処理を指定する方法の詳細については、「[ルールの条件を指定する](812c131a-fe23-1b8b-5e2d-9459d7102630.md)」を参照してください。



## Related Topics

[Outlook オブジェクト モデル リファレンス](73221b13-d8d8-99b8-3394-b95dbbfd5ddc.md)

[取得した AddressRuleCondition オブジェクトのメンバー](d15b0554-6b47-b201-fd41-744ea056d3f6.md)




