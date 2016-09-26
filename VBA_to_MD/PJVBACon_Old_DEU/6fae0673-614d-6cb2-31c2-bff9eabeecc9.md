
# Application.FileBuildID Property (Project)

Ruft die Datei Build-ID (ID) des angegebenen Projekts ab. Die Dateibuild-ID besteht aus der Version und Build von Project-Anwendung, die die Datei erstellt. Read-only  **Zeichenfolge**.


## Syntax

 _Ausdruck_. **FileBuildID**( ** _Name_**, ** _UserID_**, ** _DatabasePassWord_** )

 _Ausdruck_ Eine Variable, die ein Objekt **Application** repräsentiert.


### Parameter



|**Name**|**Erforderlich/Optional**|**Datentyp**|**Beschreibung**|
|:-----|:-----|:-----|:-----|
| _Name_|Erforderlich|**String**|Der Name einer Projektdatei, Quelldatei oder Datenquelle.|
| _UserID_|Optional|**String**|Eine Benutzer-ID für den Zugriff auf eine Datenbank. Wird durch  _Name_ keine Datenbank angegeben, wird _UserID_ ignoriert.|
| _DatabasePassWord_|Optional|**Variant**|Ein Kennwort für den Zugriff auf eine Datenbank. Wird durch  _Name_ keine Datenbank angegeben, wird _DatabasePassWord_ ignoriert.|

## Bemerkungen

Die  **FileBuildID** -Eigenschaft kann die Dateibuild-ID einer Projektdatei abrufen, ohne versehentlich zu öffnen.


## Beispiel

Im folgenden Beispiel wird die Dateibuild-ID für das Projekt Test.mpp. Wenn Project erstellen, die die Datei erstellt 15.0.4027.1000 ist, ist der  **FileBuildID** -Wert "15,0,4027,1000".


```
Sub File_BuildID()
    Dim ProjID As String

    ProjID = Application.FileBuildID("C:\Project\VBA\Samples\Test.mpp")
    Debug.Print ProjID
End Sub
```

