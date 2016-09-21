

---
ms.Toctitle:マネージ COM アドインを使用して Office Fluent リボンをカスタマイズする
title:マネージ COM アドインを使用して Office Fluent リボンをカスタマイズする
ms.ContentId:7926e6bc-c7ae-cc6f-faa5-28e2e6de664c
---
# マネージ COM アドインを使用して Office Fluent リボンをカスタマイズする




UNRESOLVED_TOKEN_VAL(officenv) スイートで Microsoft Office Fluent ユーザー インターフェイスのリボン コンポーネントを使用すると、ユーザーは Office アプリケーションを柔軟に操作できます。リボン拡張機能 (RibbonX) では、単純なテキストベースの宣言型 XML マークアップを使用して、リボンを作成およびカスタマイズします。



このトピックのコード例は、開かれるドキュメントに関係なく、Office アプリケーションのリボンをカスタマイズする方法を示しています。以下の手順では、管理された COM アドインを使用してアプリケーション レベルのカスタマイズを作成し、Microsoft Visual C# を使用して UNRESOLVED_TOKEN_VAL(vsdev11long) のアドインを作成します。このプロジェクトでは、カスタム タブ、カスタム グループ、およびカスタム ボタンをリボンに追加します。手順を完了するには、以下のタスクを実行します。

1. XML カスタマイズ ファイルを作成します。
2. C# を使用して、マネージ COM アドイン プロジェクトを UNRESOLVED_TOKEN_VAL(vsdev11long) に作成します。
3. XML カスタマイズ ファイルを埋め込みリソースとしてプロジェクトに追加します。
4. **IRibbonExtensibility** インターフェイスを実装します。
5. ボタンがクリックされたときに起動するコールバック メソッドを作成します。
6. プロジェクトのビルド、インストール、およびテストを行います。




**XML カスタマイズ ファイルを作成する**



この手順では、カスタム コンポーネントをリボンに追加するファイルを作成します。 

1. テキスト エディターで、以下の XML マークアップを追加します。 

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"> 
  <ribbon> 
    <tabs> 
      <tab id="CustomTab" label="My Tab"> 
        <group id="SampleGroup" label="Sample Group"> 
          <button id="Button" label="Insert Company Name" size="large" onAction="InsertCompanyName" /> 
        </group > 
      </tab> 
    </tabs> 
  </ribbon> 
</customUI> 

```

2. ファイルを閉じて、「**customUI.xml**」という名前を付けて保存します。








**マネージ COM アドイン プロジェクトを作成する**



この手順では、COM アドイン C# プロジェクトを UNRESOLVED_TOKEN_VAL(vsdev11long) に作成します。

1. UNRESOLVED_TOKEN_VAL(vsdev11long) を起動します。
2. **[ファイル]** メニューの **[新しいプロジェクト]** をクリックします。
3. **[新しいプロジェクト]** ダイアログ ボックスの **[プロジェクトの種類]** の下で、**[その他のプロジェクト]** を展開し、**[機能拡張プロジェクト]** をクリックして、**[共有アドイン]** をダブルクリックします。
4. プロジェクトに名前を付けます。この例では、「**RibbonXSampleCS**」と入力します。
5. **共有アドイン ウィザード**の最初の画面で、**[次へ]** をクリックします。
6. **[Visual C# を使用してアドインを作成]** を選択し、**[次へ]** をクリックします。
7. **[Microsoft Word]** 以外のすべての選択を解除して、**[次へ]** をクリックします。
8. アドインの名前と説明を入力し、**[次へ]** をクリックします。
9. **[アドイン オプションを選択します]** 画面で、**[ホスト アプリケーションの読み込み時にアドインを読み込む]** をオンにして、**[次へ]** をクリックします。
10. **[完了]** をクリックしてウィザードを終了します。 









**プロジェクトへの外部参照を追加する**



この手順では、Word プライマリ相互運用機能アセンブリおよびタイプ ライブラリへの参照を追加します。 

1. ソリューション エクスプローラーで **[参照]** を右クリックし、**[参照の追加]** をクリックします。

>[!NOTE]
>**[参照]** フォルダーが表示されない場合は、**[プロジェクト]** メニューをクリックし、**[すべてのファイルを表示]** をクリックします。


2. **[.NET]** タブを下にスクロールし、**Ctrl** キーを押して **[Microsoft.Office.Interop.Word]** を選択します。
3. **[COM]** タブを下にスクロールし、**[Microsoft Office 15.0 Object Library]** (または使用している Office のバージョンに適したライブラリ) を選択して、**[OK]** をクリックします。
4. 以下の名前空間参照がまだない場合、これらをプロジェクト (**namespace** 行の下) に追加します。 

```csharp
using System.Reflection; 
using Microsoft.Office.Core; 
using System.IO; 
using System.Xml; 
using Extensibility; 
using System.Runtime.InteropServices; 
using MSword = Microsoft.Office.Interop.Word; 

```









**カスタマイズ ファイルを埋め込みリソースとして追加する**



この手順では、XML カスタマイズ ファイルをプロジェクトの埋め込みリソースとして追加します。

1. ソリューション エクスプローラーで **[RibbonXSampleCS]** を右クリックし、**[追加]** をポイントして、**[既存の項目]** をクリックします。
2. 作成した **customUI.xml** ファイルに移動し、ファイルを選択して、**[追加]** をクリックします。
3. ソリューション エクスプローラーで **[customUI.xml]** を右クリックして、**[プロパティ]** を選択します。
4. **[プロパティ]** ウィンドウで **[ビルド アクション]** を選択し、**[埋め込まれたリソース]** まで下にスクロールします。








**IRibbonExtensibility インターフェイスを実装する**



この手順では、実行時に Word アプリケーションへの参照を作成するためのコードを Extensibility.IDTExtensibility2::OnConnection に追加します。また、**IRibbonExtensibility** インターフェイスの唯一のメンバーである **GetCustomUI** を実装します。

1. ソリューション エクスプローラーで、**[Connect.cs]** を右クリックし、**[コードの表示]** をクリックします。
2. **Connect** メソッドの後に以下の宣言を追加します。この宣言により、**Word アプリケーション** オブジェクトへの参照を作成します。 `private MSword.Application applicationObject;`
3. 以下の行を **OnConnection** メソッドに追加します。このステートメントにより、**Word アプリケーション** オブジェクトのインスタンスを作成します。 `applicationObject =(MSword.Application)application;`
4. パブリック クラスの Connect ステートメントの末尾にコンマを追加し、「**IRibbonExtensibility**」と入力します。

>[!NOTE]
>Microsoft IntelliSense を使用して、独自のインターフェイス メソッドを挿入できます。たとえば、パブリック クラスの Connect: ステートメントの末尾に 「**IRibbonExtensibility**」と入力し、右クリックして **[インターフェイスの実装]** をポイントし、**[インターフェイスの明示的な実装]** をクリックします。これにより、**GetCustomUI** メソッドのスタブが追加されます。この実装は以下のようなコードになります。 



```csharp
string IRibbonExtensibility.GetCustomUI(string RibbonID) 
{ 
}
```

5. 次のステートメントを **GetCustomUI** メソッドに挿入し、既存のコードを上書きします。`return GetResource("customUI.xml");`
6. 以下のメソッドを **GetCustommUI** メソッドの下に挿入します。 

```csharp
private string GetResource(string resourceName) 
        { 
            Assembly asm = Assembly.GetExecutingAssembly(); 
            foreach (string name in asm.GetManifestResourceNames()) 
            { 
                if (name.EndsWith(resourceName)) 
                { 
                    System.IO.TextReader tr = new System.IO.StreamReader(asm.GetManifestResourceStream(name)); 
                    //Debug.Assert(tr != null); 
                    string resource = tr.ReadToEnd(); 
 
                    tr.Close(); 
                    return resource; 
                } 
            } 
            return null; 
        } 

```
**GetCustomUI** メソッドは **GetResource** メソッドを呼び出します。**GetResource** メソッドは、実行時にこのアセンブリへの参照を設定した後、customUI.xml という名前のリソースを見つけるまで埋め込みリソースをループ処理します。次に、XML マークアップが含まれる埋め込みファイルを読み取る **StreamReader** オブジェクトのインスタンスを作成します。プロシージャは XML を **GetCustomUI** メソッドに渡し、このメソッドは XML をリボンに戻します。または、XML マークアップが含まれる文字列を構築し、その文字列を **GetCustomUI** メソッドに直接読み込むこともできます。
7. **GetResource** メソッドの後に、このメソッドを追加します。このメソッドは、ドキュメントのページの先頭に会社名を挿入します。 

```csharp
public void InsertCompanyName(IRibbonControl control) 
        { 
        // Inserts the specified text at the beginning of a range or selection. 
            string MyText; 
            MyText = "Microsoft Corporation"; 
 
            MSword.Document doc = applicationObject.ActiveDocument; 
 
            //Inserts text at the beginning of the active document. 
            object startPosition = 0; 
            object endPosition = 0; 
            MSword.Range r = (MSword.Range)doc.Range( 
                   ref startPosition, ref endPosition); 
            r.InsertAfter(MyText); 
        } 

```









**プロジェクトをビルドしてインストールする**



この手順では、アドインおよびそのセットアップ プロジェクトをビルドします。続行する前に、Word が終了していることを確認します。

1. **[プロジェクト]** メニューの **[ソリューションのビルド]** をクリックします。ビルドが完了すると、ウィンドウの左下に通知が表示されます。
2. ソリューション エクスプローラーで、**[RibbonXSampleCSSetup]** を右クリックし、**[ビルド]** をクリックします。
3. **[RibbonXSampleCSSetup]** を再度右クリックし、**[インストール]** をクリックして、**RibbonXSampleCSSetup セットアップ ウィザード**を開始します。
4. 各画面で **[次へ]** をクリックし、最後の画面で **[閉じる]** をクリックします。
5. Word を起動します。他のタブの右に **[My Tab]** タブが表示されます。








**プロジェクトをテストする**



**[My Tab]** タブをクリックし、**[Insert Company Name]** をクリックして、会社名をドキュメントのカーソルの位置に挿入します。カスタマイズされたリボンが表示されない場合、以下の手順を実行して、Windows レジストリにエントリを追加しなければならないことがあります。 

>[!CAUTION]
>以下の手順では、レジストリの変更方法について説明します。レジストリを変更する前に、レジストリのバックアップを必ず作成し、問題が発生した場合にレジストリを復元する方法について確実に理解しておいてください。レジストリのバックアップ、復元、および編集方法については、Microsoft サポート技術情報で記事「**256986 Microsoft Windows レジストリの説明**」を検索してください。



1. ソリューション エクスプローラーで、セットアップ プロジェクトの **[RibbonXSampleCSSetup]** を右クリックし、**[表示]** をポイントしてから **[レジストリ]** をクリックします。
2. **[レジストリ]** タブから、アドイン用の次のレジストリ キーに移動します。HKCU\Software\Microsoft\Office\Word\AddIns\RibbonXSampleCS.Connect

>[!NOTE]
>**[RibbonXSampleCS.Connect]** キーが存在しない場合は、そのキーを作成できます。キーを作成するには、**[アドイン]** フォルダーを右クリックし、**[新規]** をポイントして、**[キー]** をクリックします。キーに「**RibbonXSampleCS.Connect**」という名前を付けます。「**LoadBehavior**」という名前の **DWord** を追加して、その値を「**3**」に設定します。 










