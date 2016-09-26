
# Security Behavior of the Outlook Object Model

## 

Outlook-Objektmodell enthält Einstiegspunkte für Outlook-Daten an bestimmten Speicherorten Daten speichern und Senden von E-mails zugreifen. Diese Einstiegspunkte sind für legitime und bösartige Anwendungsentwickler gleichermaßen verfügbar. Versionen von Outlook 98 und Outlook 2000 mit dem Outlook-e-Mail-Sicherheitsupdate angewendet, und alle nachfolgende Versionen ab Outlook 2000 SP2 den Objektmodellschutz verwenden, um Benutzer zu schützen. Object Model Guard warnt Benutzer vor und fordert den Benutzer zur Bestätigung beim nicht vertrauenswürdige Anwendungen das Objektmodell verwenden, um e-Mail-Adressinformationen zu erhalten, werden Daten außerhalb von Outlook gespeichert, bestimmte Aktionen ausführen und e-Mail-Nachrichten senden. Zwar Object Model Guard erfolgreich identifiziert und schützen diese Einstiegspunkte, sind zwei verschiedene Hauptaspekte, die den Objektmodellschutz Hauptproblemen vorhanden:


- Bestätigung für legitime Applikationen übermäßige Sicherheit können die standardmäßigen Umständen Applikationen den Objektmodellschutz in früheren Versionen von Outlook Aufrufen führen.
    
- Die Nachteile von COM und Windows Identifizierung der jeweiligen Anwendung, die den Objektmodellschutz aufruft haben es für Benutzer schwierig zu Sicherheit mit Sicherheit Antworten vorgenommen.
    
Weitere Informationen über die verschiedenen Sicherheitshinweise von den Objektmodellschutz finden Sie unter [Sicherheitswarnungen für Outlook Object Model](7e0cd805-5104-73af-d74f-b00480db91c4.md). Weitere Informationen zu den geschützten Einstiegspunkte des Objektmodells finden Sie unter [geschützte Eigenschaften und Methoden](8522d350-a257-2924-2260-3cc02b6ebbca.md).


## Standardverhalten

Versionen von Outlook vor Outlook 2007 haben verlassen, auf den Objektmodellschutz Outlook Adressbuchdaten zu schützen und Vermeiden von nicht vertrauenswürdigen Anwendungen e-Mail-Nachrichten senden. Zwar weiterhin Outlook den Objektmodellschutz verwenden, um ähnliche Schutz bereitzustellen, hat es neue Standard Umständen definiert, wenn der Objektmodellschutz Warnungen, reduzieren übermäßig viele Sicherheitswarnungen angemessenen generiert Beibehaltung einen angemessenen Grad an Sicherheit für Outlook-Clients.

 **In-Process-Add-ins**

In-Process-Outlook-add-ins führen gerade die Host-Outlook-Anwendung aus. In-Process-COM-add-ins in Outlook sind standardmäßig als vertrauenswürdig. Diese COM-Add-ins werden in der Liste der vertrauenswürdigen Anwendungen vom Administrator des Clientcomputers registriert, und verwenden, muss das  **[Application](797003e7-ecd1-eccb-eaaf-32d6ddde8348.md)** -Objekt, das an das **OnConnection** -Ereignis der Notiz Add-Ins übergeben wird, wenn Sie ein neues **Application** -Objekt mithilfe der **[CreateObject](09b6ff5b-a750-c07d-7499-c1f8a00214fe.md)** -Methode erstellen, dieses Objekt und alle untergeordneten Objekte, Eigenschaften und Methoden des nicht vertrauenswürdig sind.

Finden Sie weitere Informationen über das  **OnConnection** -Ereignis der[IDTExtensibility2](https://msdn.microsoft.com/en-us/library/extensibility.idtextensibility2.aspx) -Dokumentation auf MSDN.

 **Prozessübergreifenden-Add-ins**

Standardmäßig nutzt die Outlook auf das Vorhandensein und den Status der eine geeignete Antivirussoftware auf dem Clientcomputer prozessübergreifenden Applikationen vertraut: Wenn Outlook erkennt, dass mit dem akzeptable Status Antivirensoftware ausgeführt wird, wird Outlook Sicherheitswarnungen für den Endbenutzer deaktiviert. Alle prozessübergreifenden COM-Aufrufer und add-ins werden ohne Sicherheitswarnungen ausgeführt, wenn alle der folgenden Bedingungen zu speichern:


- Der Clientcomputer wird ausgeführt, Windows XP Service Pack 2 (SP2), Windows Vista oder eine höhere Version von Windows und Windows Security Center (Windows-Skriptkomponente) gibt an, dass die Antivirussoftware auf dem Computer in einer "Gut" Integritätsstatus ist.
    
- Die Antivirussoftware auf dem Clientcomputer installiert ist darauf ausgelegt, für Windows Vista, Windows XP SP2 oder höher.
    
- Outlook wird auf dem Clientcomputer in eine der folgenden Arten konfiguriert:
    
      - Verwendet die standardmäßige Outlook-Sicherheitseinstellungen (d. h., keine Gruppenrichtlinien einrichten)
    
  - Sicherheitseinstellungen mithilfe von Gruppenrichtlinien definiert verwendet jedoch keinen programmgesteuerten Zugriffsrichtlinie angewendet
    
  - Verwendet von Sicherheitseinstellungen mithilfe von Gruppenrichtlinien die festgelegt ist, eine Warnung, wenn die Antivirensoftware inaktiv oder nicht mehr aktuell ist definiert
    
Weitere Informationen finden Sie unter der "Code Security Changes in Microsoft Office Outlook 2007" im MSDN-Artikel.


## Sicherheitsoptionen

 **Windows-Gruppenrichtlinien**

Administratoren können Vertrauensstellungscenter in Outlook verwenden, um das Standardverhalten zu ändern. Wählen Sie den Zugriff auf das Sicherheitscenter  **Tools** und anschließend auf **Trust Center** aus. Klicken Sie im Sicherheitscenter auf **Den programmgesteuerten Zugriff**. Im Dialogfeld  **Sicherheit für den programmgesteuerten Zugriff** enthält Optionen, jedoch nicht das Standardverhalten.

Die drei Einstellungen im Dialogfeld  **Sicherheit für den programmgesteuerten Zugriff** sind:


-  **Warnen verdächtigen Aktivitäten, wenn mein Antivirusprogramm inaktiv oder veraltet (empfohlen)** Diese Einstellung ist die Standardeinstellung, und das oben beschriebene Verhalten implementiert. Dies ist die empfohlene Einstellung für alle Benutzer.
    
-  **Immer warnen verdächtigen Aktivitäten** Diese Einstellung wird von Outlook zum verhält sich wie Outlook 2003, denen prozessübergreifenden COM-Aufrufer und nicht vertrauenswürdigen Add-Ins Sicherheitswarnungen aufrufen zurückgesetzt.
    
-  **Nie warnen, verdächtigen Aktivitäten (nicht empfohlen)** Diese Einstellung wird nie Sicherheitshinweise und Object Model Guard wird deaktiviert. Diese Einstellung sollte nur in einer kontrollierten Umgebung verwendet werden, in dem das Risiko von bösartigem Code auf dem Computer niedrig ist.
    
Diese Einstellungen sind nur verfügbar, wenn der aktuelle Benutzer auf dem Computer als Administrator angemeldet ist. Benutzer ohne Administratorrechte sehen die aktuelle Einstellung jedoch ist nicht möglich, diese zu ändern. Einstellungen für den programmgesteuerten Zugriff können auch über die Gruppenrichtlinie gesteuert werden. Weitere Informationen zum Konfigurieren von Outlook-Einstellungen mithilfe von Gruppenrichtlinien finden Sie in der Office Resource Kit-Websites.

 **Sicherheitsformular in öffentlichen Exchange-Ordner**

Administratoren können das Outlook-Sicherheitsformular in einem öffentlichen Ordner befindet, um Outlook konfigurieren. In diesem Fall Outlook nicht den Status der Antivirussoftware nutzen und wird standardmäßig nur Trust-add-ins im Sicherheitsformular aufgeführt. Es werden nur drei Verhaltensweisen: Benutzer wird aufgefordert, nie prompt und automatisch zulassen und verweigern nie prompt und automatisch.

Um den neuen Code Sicherheitsverhalten basierend auf den Status der Antivirussoftware nutzen zu können, müssen Administratoren entweder über die standardmäßige Outlook-Sicherheitseinstellungen verwenden oder Konfigurieren von Outlook zum Gruppenrichtlinien verwenden, um dieses Verhalten zu überschreiben.

