
# TextBox.BorderColor Property (Outlook Forms Script)

Zurückgeben oder Festlegen einer  **Long**, der die Rahmen die Farbe eines Objekts angibt. Lese-/Schreibzugriff.


## Syntax

 _Ausdruck_. **BorderColor**

 _Ausdruck_ Eine Variable, die ein **TextBox** -Objekt darstellt.


## Bemerkungen

Sie können eine beliebige ganze Zahl, die eine gültige Farbe darstellt. Sie können auch eine Farbe angeben, mithilfe der Visual Basic  **RGB** -Funktion mit Rot, Grün und Blau-Komponenten. Der Wert jeder Farbkomponente ist eine ganze Zahl, die von 0 bis 255. Beispielsweise können Sie Blaugrün wie die ganzzahligen Wert 4966415 oder als Rot-, Grün- und Blau Farbe Komponenten 15, 200, 75, angeben, wie im folgenden Beispiel dargestellt.


```
RGB(15,200,75)
```

Um die  **BorderColor** -Eigenschaft verwenden zu können, muss die **[BorderStyle](c71b8117-a731-d0ab-89a7-84dd9aa089c4.md)** -Eigenschaft auf einen anderen Wert als 0 festgelegt sein.

 **BorderStyle** verwendet **BorderColor-Eigenschaft**, um die Rahmenfarben zu definieren. Die **[SpecialEffect](b7365d4e-c25d-9fa6-c088-0cc5bb6bb200.md)** -Eigenschaft verwendet die Systemfarben ausschließlich auf die um entsprechenden Rahmenfarben zu definieren. Für Windows-Betriebssystemen werden Systemeinstellungen Farbe mit dem Symbol **Anzeigen** in der Systemsteuerung festgelegt.

