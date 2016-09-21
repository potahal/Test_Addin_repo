

---
ms.Toctitle:TextBox.MultiLine プロパティ (Outlook フォーム スクリプト)
title:TextBox.MultiLine プロパティ (Outlook フォーム スクリプト)
ms.ContentId:f42aadc5-ecd9-090b-cdf0-aba0a1a024b2
---
# TextBox.MultiLine プロパティ (Outlook フォーム スクリプト)




取得または設定する**ブール値**コントロールが受け入れるし、複数行のテキストを表示するかどうかを指定します。読み取り/書き込み。

## 構文

            UNRESOLVED_TOKEN_VAL(offexpression).**MultiLine**




            UNRESOLVED_TOKEN_VAL(offexpression)
            **TextBox** オブジェクトを表す変数です。



## 注釈
True の場合、複数行の文字列の取得と表示を許可します (既定値)。Falase の場合、複数行の文字列の取得と表示を許可しません。



複数行の文字列の取得と表示が設定されているテキスト ボックス (**TextBox**) コントロールでは、強制改行を使用でき、文字列の量に合わせて行数が調整されます。必要に応じて垂直スクロール バーを表示することもできます。



単一行の **TextBox** では、改行は使用できず、垂直方向のスクロール バーは使用されません。



**MultiLine**プロパティと**WordWrap**プロパティをサポートするコントロールの場合は、 **MultiLine**が**False**の場合、**ワード ラップ**は無視されます。



単一行のコントロールでは、**WordWrap** プロパティの値は無視されます。



複数行の **TextBox** で **MultiLine** を **False** に変更すると、非印字文字 (キャリッジ リターンや改行など) も含め、**TextBox** のすべての文字が 1 行に収められます。



**実行される**と、**複数行**のプロパティは、密接に関連します。**True**および**False**の**実行される**値は、 **MultiLine**が**True**の場合にのみ適用されます。**MultiLine**が**False**の場合は、 **enter キーを常に****実行される**の値に関係なく、タブ オーダーで次のコントロールにフォーカスに移動します。



**Ctrl + Enter** を押した場合の効果は、**MultiLine** の値によっても異なります。**MultiLine** が **True** に設定されている場合は、**EnterKeyBehavior** の値に関係なく、**Ctrl + Enter** を押すと、新しい行が作成されます。**MultiLine** が **False** に設定されている場合は、**Ctrl + Enter** を押しても何も生じません。



**TabKeyBehavior**プロパティと**MultiLine**プロパティは、密接に関連します。上記の値は、 **MultiLine**が**True**の場合にのみ適用されます。**MultiLine**が**False**の場合は、 **TabKeyBehavior**の値に関係なく、タブ オーダーで次のコントロールにフォーカスを移動する**tab キーを常に**。



**Ctrl + Tab** を押した場合の効果は、**MultiLine** の値によっても異なります。**MultiLine** が **True** に設定されている場合は、**TabKeyBehavior** の値に関係なく、**Ctrl + Tab** を押すと、新しい行が作成されます。**MultiLine** が **False** に設定されている場合は、**Ctrl + Tab** を押しても何も生じません。




