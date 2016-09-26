
# RuleConditions.From Property (Outlook)

Gibt ein  **[ToOrFromRuleCondition](ec5cae2a-cde8-5681-6a49-74e2f0226a4f.md)** -Objekt mit einer **[ToOrFromRuleCondition.ConditionType](a5c6e08c-643e-965d-cd3e-b434f20579a0.md)** der **OlConditionFrom** zurück. Schreibgeschützt.


## Syntax

 _Ausdruck_. **From**

 _Ausdruck_ Eine Variable, die ein **RuleConditions** -Objekt darstellt.


## Hinweise

Verwenden Sie das zurückgegebene  **ToOrFromRuleCondition** -Objekt beim Aufzählen der regelbedingungen oder Ausnahmebedingungen einer vorhandenen Regel oder beim Erstellen einer neuen Regel, gibt die Bedingung oder eine Ausnahmebedingung, die der Absender der Nachricht in der angegebenen Liste der **[Empfänger](774f56b7-4de8-9584-60cd-4fbf361f4c85.md)** ist.

Diese Eigenschaft immer der  **[RuleConditions](e8e9a05a-b36b-add2-b294-8cdc5a97e119.md)** -Auflistung gibt ein **ToOrFromRuleCondition** -Objekt unabhängig davon, ob die Regel für diese **RuleConditions** -Auflistung eine solche regelbedingung definiert wurde. Wenn die Regel definiert und eine solche regelbedingung aktiviert wurde, klicken Sie dann wird **[ToOrFromRuleCondition.Enabled](31e43906-b47a-95e3-d51b-3fa6af553fad.md)** auf **true festgelegt**.


## Siehe auch


#### Konzepte


[RuleConditions-Objekt](e8e9a05a-b36b-add2-b294-8cdc5a97e119.md)
#### Weitere Ressourcen


[Elemente des RuleConditions-Objekts](http://msdn.microsoft.com/library/b2af6ebf-f9f8-8106-20a3-1725c3b78174%28Office.15%29.aspx)