
# ComboBox.Text Property (Outlook Forms Script)

Zurückgeben oder Festlegen einer  **Zeichenfolge**, die Text in einem **[ComboBox-Steuerelement](31e7c1de-ee4e-b3d9-4579-7fc6b215bad3.md)**, ändern die ausgewählte Zeile im Steuerelement angibt. Lese-/Schreibzugriff.


## Syntax

 _Ausdruck_. **Text**

 _Ausdruck_ Eine Variable, die ein **ComboBox** -Objekt darstellt.


## Bemerkungen

Als Standard gilt eine Nullzeichenfolge ("").

 **Text** können Sie um den Wert des Steuerelements zu aktualisieren. Wenn der Wert von **Text** einem vorhandenen Listeneintrag entspricht, wird der Wert der **[ListIndex](2c4e473b-15e1-dce2-8748-30953b00a60f.md)** -Eigenschaft (der Index der aktuellen Zeile) auf die Zeile festgelegt, die **Text** übereinstimmt. Wenn der Wert der **Text** eine Zeile nicht übereinstimmt, wird **ListIndex** auf - 1 festgelegt.

Wenn die  **Text** -Eigenschaft eines **ComboBox-Steuerelements** ändert (beispielsweise wenn ein Benutzer eine Eingabe in ein Steuerelement macht), wird der neue Text mit der von **[TextColumn](5ebf37ef-4cec-ec42-d42f-ab886b86e913.md)** angegebenen Datenspalte verglichen.

 **Text** können Sie um den Wert eines Eintrags in einem **ComboBox-Steuerelements** zu ändern. Verwenden Sie die **[Spalte](f00c388f-fe1f-5458-281f-4bfa549291d5.md)** oder **[Liste](687f44e8-7b4b-eab5-93b8-022cd4d1c302.md)** -Eigenschaft für diesen Zweck.

Die Farbe des Texts wird von der  **[ForeColor](256d695a-df00-d22c-b2aa-e21036beea35.md)** -Eigenschaft bestimmt.

