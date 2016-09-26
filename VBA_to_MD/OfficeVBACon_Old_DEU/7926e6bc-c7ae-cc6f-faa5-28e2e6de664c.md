
# Anpassen der Office Fluent-Menübands mithilfe eines verwalteten COM-Add-Ins

Die Menüband-Komponente der Microsoft Office Fluent-Benutzeroberfläche in Microsoft Office-Suites verleiht Benutzern Flexibilität bei der Arbeit mit Office-Anwendungen. Die Menübanderweiterung (RibbonX) verwendet einfaches, textbasiertes, deklaratives XML-Markup, um das Menüband zu erstellen und anzupassen.

Dieses Codebeispiel zeigt, wie Sie das Menüband in einer Office-Anwendung anpassen, ganz gleich, welches Dokument gerade geöffnet ist. Anhand der folgenden Schritte erstellen Sie Anpassungen auf Anwendungsebene, wobei Sie ein verwaltetes COM-Add-In verwenden. Sie erstellen das Add-In in Microsoft Visual Studio 2012 unter Verwendung von Microsoft Visual C#. Das Projekt fügt dem Menüband eine benutzerdefinierte Registerkarte, eine benutzerdefinierte Gruppe und eine benutzerdefinierte Schaltfläche hinzu. Führen Sie hierfür die nachfolgenden Schritte durch.

1. Erstellen Sie die XML-Anpassungsdatei.
    
2. Erstellen Sie ein verwaltetes COM-Add-In-Projekt in Microsoft Visual Studio 2012 mit C#.
    
3. Fügen Sie die XML-Anpassungsdatei dem Projekt als eingebettete Ressource hinzu.
    
4. Implementieren Sie die  **IRibbonExtensibility** -Benutzeroberfläche.
    
5. Erstellen Sie eine Rückrufmethode, die durch Klicken auf die Schaltfläche ausgelöst wird.
    
6. Erstellen, installieren und testen Sie das Projekt.
    
 **Erstellen der XML-Anpassungsdatei**
In diesem Schritt erstellen Sie die Datei, die dem Menüband die benutzerdefinierten Komponenten hinzufügt.

1. Fügen Sie das folgende XML-Markup in einem Text-Editor hinzu.
    
  ```XML
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

2. Schließen und speichern Sie die Datei unter dem Namen  **customUI.xml**.
    

 **Erstellen des verwalteten COM-Add-In-Projekts**
In diesem Schritt erstellen Sie das COM-Add-In C#-Projekt in Microsoft Visual Studio 2012.

1. Starten Sie Microsoft Visual Studio 2012.
    
2. Klicken Sie im Menü  **Datei** auf **Neues Projekt**.
    
3. Erweitern Sie im Dialogfeld  **Neues Projekt** unter **Projekttypen** die Option **Andere Projekttypen**, klicken Sie dann auf  **Erweiterungsprojekte**, und doppelklicken Sie auf  **Gemeinsames Add-In**.
    
4. Geben Sie einen Namen für das Projekt ein. Verwenden Sie für dieses Beispiel  **RibbonXSampleCS**.
    
5. Klicken Sie auf der ersten Seite des  **Assistenten für gemeinsames Add-In** auf **Weiter**.
    
6. Wählen Sie  **Ein Add-In mit Visual C# erstellen**, und klicken Sie dann auf  **Weiter**.
    
7. Lassen Sie nur  **Microsoft Word** ausgewählt, und klicken Sie auf **Weiter**.
    
8. Geben Sie einen Namen und eine Beschreibung für das Add-In ein, und klicken Sie auf  **Weiter**.
    
9. Klicken Sie auf der Seite  **Wählen Sie die Add-In-Optionen aus** auf **Das Add-In laden, wenn die Hostanwendung geladen wird**, und klicken Sie auf  **Weiter**.
    
10. Klicken Sie auf  **Fertigstellen**, um den Assistenten abzuschließen.
    

 **Hinzufügen von externen Verweisen zum Projekt**
In diesem Schritt fügen Sie Verweise zum primären Interop-Assemblys von Word zur Typbibliothek hinzu.

1. Klicken Sie im Projektmappen-Explorer mit der rechten Maustaste auf  **Verweise**, und klicken Sie dann auf  **Verweis hinzufügen**.
    
     **Hinweis**  Falls der Ordner  **Verweise** nicht angezeigt wird, klicken Sie auf das Menü **Projekt**, und klicken Sie dann auf  **Alle Dateien anzeigen**.
2. Führen Sie auf der Registerkarte  **.NET** einen Bildlauf nach unten durch, drücken Sie die **STRG**-TASTE, und wählen Sie  **Microsoft.Office.Interop.Word**.
    
3. Führen Sie auf der Registerkarte  **COM** einen Bildlauf nach unten durch, wählen Sie **Microsoft Office 15.0-Objektbibliothek** (oder die Bibliothek, die für Ihre Version von Office passend ist), und klicken Sie auf **OK**.
    
4. Fügen Sie dem Projekt unterhalb der Zeile  **namespace** die folgenden Namespace-Verweise hinzu, sofern diese noch nicht existieren.
    
  ```C#
  using System.Reflection; 
using Microsoft.Office.Core; 
using System.IO; 
using System.Xml; 
using Extensibility; 
using System.Runtime.InteropServices; 
using MSword = Microsoft.Office.Interop.Word; 

  ```


 **Hinzufügen der Anpassungsdatei als eingebettete Ressource**
In diesem Schritt fügen Sie die XML-Anpassungsdatei als eingebettete Ressource in das Projekt ein.

1. Klicken Sie im Projektmappen-Explorer mit der rechten Maustaste auf  **RibbonXSampleCS**, zeigen Sie auf  **Hinzufügen**, und klicken Sie dann auf  **Vorhandenes Element**.
    
2. Navigieren Sie zu der von Ihnen erstellten Datei  **customUI.xml**, markieren Sie die Datei, und klicken Sie auf  **Hinzufügen**.
    
3. Klicken Sie im Projektmappen-Explorer mit der rechten Maustaste auf  **customUI.xml**, und klicken Sie dann auf  **Eigenschaften**.
    
4. Wählen Sie im Fenster  **Eigenschaften** die Option **Buildvorgang** aus, und führen Sie dann einen Bildlauf nach unten zu **Eingebettete Ressource** aus.
    

 **Implementieren der IRibbonExtensibility-Benutzeroberfläche**
In diesem Schritt können Sie Code zu "Extensibility.IDTExtensibility2::OnConnection" hinzufügen, um einen Verweis auf die Word-Anwendung zur Laufzeit zu erstellen. Außerdem implementieren Sie das einzige Mitglied der  **IRibbonExtensibility** -Benutzeroberfläche, und zwar **GetCustomUI**.

1. Klicken Sie im Projektmappen-Explorer mit der rechten Maustaste auf  **Connect.cs**, und klicken Sie dann auf  **Code anzeigen**.
    
2. Nach der  **Connect** -Methode fügen Sie die folgende Deklaration hinzu, wobei ein Verweis auf das **Word-Anwendungs** -Objekt erstellt wird:
    
     `private MSword.Application applicationObject;`
    
3. Fügen Sie der  **OnConnection** -Methode den folgenden Code hinzu: Diese Anweisung erstellt eine Instanz des **Word-Anwendungs** -Objekts:
    
     `applicationObject =(MSword.Application)application;`
    
4. Fügen Sie am Ende der Connect-Anweisung für die öffentliche Klasse ein Komma ein, und geben Sie  **IRibbonExtensibility** ein.
    
     **Hinweis**  Sie können Microsoft IntelliSense verwenden, um Schnittstellenmethoden für Sie einzufügen. Geben Sie beispielsweise am Ende der Connect-Anweisung für die öffentliche Klasse  **IRibbonExtensibility** ein, klicken Sie mit der rechten Maustaste auf **Schnittstelle implementieren**, und klicken Sie anschließend auf  **Schnittstelle explizit implementieren**. Auf diese Weise wird ein Stub für die  **GetCustomUI** -Methode hinzugefügt. Die Implementierung sieht ähnlich wie der folgende Code aus.

  ```C#
  string IRibbonExtensibility.GetCustomUI(string RibbonID) 
{ 
}
  ```

5. Fügen Sie die folgende Anweisung in die  **GetCustomUI** -Methode ein, und überschreiben Sie den vorhandenen Code. `return GetResource("customUI.xml");`
    
6. Fügen Sie die folgende Methode unterhalb der  **GetCustommUI** -Methode ein:
    
  ```C#
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


    Die  **GetCustomUI** -Methode ruft die **GetResource** -Methode auf. Die **GetResource** -Methode setzt während der Laufzeit einen Verweis auf diese Assembly und durchläuft dann die eingebettete Ressource, bis es die eine benannte Datei "customUI.xml" findet. Anschließend erstellt sie eine Instanz des **StreamReader** -Objekts, das die eingebettete Datei mit dem XML-Markup liest. Die Prozedur reicht die XML-Daten an die **GetCustomUI** -Methode weiter, die die XML-Daten an das Menüband zurückgibt. Alternativ können Sie eine Zeichenfolge erstellen, die das XML-Markup enthält und diese direkt in die **GetCustomUI** -Methode einlesen.
    
7. Fügen Sie nach der  **GetResource** -Methode diese Methode ein. Diese Methode fügt den Firmennamen in das Dokument am Seitenbeginn ein.
    
  ```C#
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


 **Erstellen und installieren des Projekts**
In diesem Schritt erstellen Sie das Add-In und sein Setup-Projekt. Bevor Sie fortfahren, stellen Sie sicher, dass Sie Word beendet haben.

1. Klicken Sie im Menü  **Projekt** auf **Projektmappe erstellen**. Wenn das Projekt erstellt wurde, wird unten links im Fenster eine Benachrichtigung angezeigt.
    
2. Klicken Sie im Projektmappen-Explorer mit der rechten Maustaste auf  **RibbonXSampleCSSetup**, und klicken Sie dann auf  **Erstellen**.
    
3. Klicken Sie erneut mit der rechten Maustaste auf  **RibbonXSampleCSSetup**, und klicken Sie auf  **Installieren**, um den  **RibbonXSampleCSSetup-Setup-Assistenten** zu starten.
    
4. Klicken Sie auf jeder Seite auf  **Weiter**, und klicken Sie anschließend auf der letzten Seite auf  **Schließen**.
    
5. Starten Sie Word. Rechts neben den anderen Registerkarten sollte die Registerkarte  **My Tab** angezeigt werden.
    

 **Testen des Projekts**
Klicken Sie auf die Registerkarte  **My Tab**, und klicken Sie anschließend auf  **Firmennamen einfügen**, um an der Position des Cursors im Dokument den Firmennamen einzufügen. Wird kein angepasstes Menüband angezeigt, müssen Sie möglicherweise einen Eintrag in der Windows-Registrierung einfügen, indem Sie die folgenden Schritte ausführen.

 **Vorsicht**  Die nächsten Schritte enthalten Informationen dazu, wie Sie die Windows-Registrierung bearbeiten. Bevor Sie Änderungen an der Registrierung vornehmen, erstellen Sie eine Sicherungskopie, und stellen Sie sicher, dass Sie genau wissen, wie die Registrierung im Falle eines Problems wiederhergestellt wird. Weitere Informationen zum Sichern, Wiederherstellen und Bearbeiten der Registrierung finden Sie im folgenden Artikel der Microsoft Knowledge Base:  **256986 Beschreibung der Microsoft Windows-Registrierung**.


1. Klicken Sie im Projektmappen-Explorer mit der rechten Maustaste auf das Setup-Projekt,  **RibbonXSampleCSSetup**, zeigen Sie auf  **Anzeigen**, und klicken Sie dann auf  **Registrierung**.
    
2. Navigieren Sie über die Registerkarte  **Registrierung** zum folgenden Registrierungsschlüssel für das Add-In: HKCU\Software\Microsoft\Office\Word\AddIns\RibbonXSampleCS.Connect
    
     **Hinweis**  Wenn der Schlüssel  **RibbonXSampleCS.Connect** nicht existiert, können Sie ihn erstellen. Klicken Sie hierfür mit der rechten Maustaste auf den Ordner **Addins**, zeigen Sie auf  **Neu**, und klicken Sie anschließend auf  **Schlüssel**. Benennen Sie den Schlüssel  **RibbonXSampleCS.Connect**. Fügen Sie  **LoadBehavior** **DWord** ein, und setzen Sie seinen Wert auf **3**.

