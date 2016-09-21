

---
ms.Toctitle:オートメーションと共にイベントを使用する
title:オートメーションと共にイベントを使用する
ms.ContentId:6ca0a0fa-1cda-c052-4dee-1055cceb2b28
---
# オートメーションと共にイベントを使用する





          UNRESOLVED_TOKEN_VAL(outlooknv1) オブジェクトのイベント ハンドラーを、別のアプリケーションの Microsoft Visual Basic または Microsoft Visual Basic for Applications (VBA) で作成するには、次の 4 つの手順を完了する必要があります。

1. Outlook オブジェクト ライブラリへの参照を設定します。
2. イベントに応答するオブジェクト変数を宣言します。
3. 個別のイベント プロシージャを記述します。
4. 宣言したオブジェクトを初期化します。




「[working with events in Outlook Visual Basic for Applications](560bb264-05d0-dbc6-39c2-b95b12f50ed9.md)」の詳細情報。

## Outlook オブジェクト ライブラリへの参照を設定します。
Outlook オブジェクトを Visual Basic または Visual Basic for Applications のコードで使用する前に、最初に Outlook オブジェクト モデルへの参照を [**参照**] ダイアログ ボックスで設定する必要があります。このダイアログ ボックスの使用方法の詳細は、お使いのプログラミング環境のオンライン ヘルプを参照してください。



## オブジェクト変数を宣言する
オブジェクト モデル ライブラリを参照した後は、使用するオブジェクトを参照する変数を宣言する必要があります。変数はオブジェクトを使用するモジュールの中で宣言できます (つまり、イベント ハンドラー プロシージャを含むモジュール)。ただし、より一般的にはクラス モジュールの中で宣言して、プログラム内のすべてのモジュールで使用できるようにします。



たとえば、クラス モジュールで **Application** オブジェクトのオブジェクト変数を宣言するには、次のようなコードを使用します。

```sourcecode
Public WithEvents myOlApp As Outlook.Application
```




`WithEvents` キーワードを使用して、そのオブジェクト変数がオブジェクトによってトリガーされたイベントへの応答に使用されることを指定する必要があります。



## イベント プロシージャを記述する
新しいオブジェクトをイベントで宣言した後は、クラス モジュールのコード ウィンドウの [**オブジェクト**] 一覧に表示されます。そして、オブジェクトのイベント プロシージャを [**プロシージャ/イベント**] 一覧から選択できます。たとえば、`myOlApp` として宣言された **Application** オブジェクトに **ItemSend** イベントを選択すると、コード ウィンドウには次の空のプロシージャが表示されます。

```sourcecode
Private Sub myOlApp_ItemSend(Item as Object, Cancel as Boolean) 
 
End Sub
```




## 宣言したオブジェクトを初期化する
プロシージャを実行する前に、**Application** オブジェクトで宣言したオブジェクト (この例では、`myOlApp`) に接続する必要があります。オブジェクトを `EventClassModule` と呼ばれるクラス モジュールで宣言した場合は、すべてのモジュールで次のコードが使用できます。

```vba
Dim myClass as New EventClassModule  
Sub Register_Event_Handler()  
    Set myClass.myOlApp = "Outlook.Application"  
End Sub
```




プロシージャが

```sourcecode
Register_Event_Handler
```




実行中の場合、フォームまたはクラス モジュールの `myOlApp` オブジェクトは、Outlook **Application** オブジェクトを指し示し、イベントが発生するとイベント プロシージャが実行されます。




