
# SharedWorkspaceFolder.FolderName-Eigenschaft (Office)

Ruft den Namen eines Unterordners innerhalb des Hauptordners der Dokumentbibliothek eines freigegebenen Arbeitsbereichs ab. Schreibgeschützt.


 **Hinweis**  Ab Microsoft Office 2010 ist dieses Objekt oder Element veraltet und sollte nicht verwendet werden.


## Syntax

 _Ausdruck_. **FolderName**

 _Ausdruck_ Eine Variable, die ein **SharedWorkspaceFolder** -Objekt darstellt.


## Hinweise

Die  **FolderName** -Eigenschaft gibt den Namen des Unterordners im Format Parentfolder/Subfolder zurück. Wenn des freigegebenen Arbeitsbereichs einen Ordner namens "Unterstützende Dokumente" enthält, gibt die **FolderName** -Eigenschaft beispielsweise freigegebene Dokumente/unterstützen Dokumente zurück.


## Beispiel

Das folgende Beispiel zeigt die Anzahl der Unterordner im freigegebenen Arbeitsbereich und deren Namen an.


```
    Dim swsFolder As Office.SharedWorkspaceFolder 
    Dim strFolderInfo As String 
    strFolderInfo = "The shared workspace contains " &amp; _ 
        ActiveWorkbook.SharedWorkspace.Folders.Count &amp; " folder(s)." &amp; vbCrLf 
    If ActiveWorkbook.SharedWorkspace.Folders.Count > 0 Then 
        For Each swsFolder In ActiveWorkbook.SharedWorkspace.Folders 
            strFolderInfo = strFolderInfo &amp; swsFolder.FolderName &amp; vbCrLf 
        Next 
    End If 
    MsgBox strFolderInfo, vbInformation + vbOKOnly, _ 
        "Folders in Shared Workspace" 
    Set swsFolder = Nothing 

```


## Siehe auch


#### Konzepte


[SharedWorkspaceFolder-Objekt](297c4ed7-2232-5240-ca34-d374038c66a2.md)
#### Weitere Ressourcen


[Elemente des SharedWorkspaceFolder-Objekts](http://msdn.microsoft.com/library/e7e0a32a-ce01-e08f-f251-27d93273110e%28Office.15%29.aspx)