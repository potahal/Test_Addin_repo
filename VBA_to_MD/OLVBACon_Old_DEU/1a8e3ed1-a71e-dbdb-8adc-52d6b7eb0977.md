
# Page.PictureTiling Property (Outlook Forms Script)

Zurückgeben oder Festlegen einer  **vom Typ Boolean**, der angibt, ob ein Bild auf dem Hintergrund des Objekts wiederholt wird. Lese-/Schreibzugriff.


## Syntax

 _Ausdruck_. **PictureTiling**

 _Ausdruck_ Eine Variable, die ein **Page** -Objekt darstellt.


## Bemerkungen

 **True,** Wenn die Grafik nebeneinander, auf dem Hintergrund, **False,** andernfalls (Standard angeordnet wird).

Wenn eine Grafik kleiner ist als das Formular oder die Seite, auf dem bzw. der sie sich befindet, können Sie die Grafik wiederholen und kachelartig auf einem Formular oder einer Seite anordnen.

Das Muster Nebeneinanderanordnen hängt von der aktuellen Einstellung der  **[PictureAlignment](c52f0b5b-c703-d9d6-1bae-e4fe9b696cf8.md)** und **[PictureSizeMode](24a0415a-f89a-c0fb-9c44-b33484c8cd49.md)** -Eigenschaft ab. Beispielsweise wenn **PictureAlignment** auf 0 festgelegt ist, das Nebeneinanderanordnen Muster beginnt an der linken oberen Ecke und wiederholt das Bild über das Formular oder die Seite und Höhe des Formulars oder der Seite. Wenn **PictureSizeMode** auf 0 festgelegt ist, schneidet das Nebeneinanderanordnen Muster die letzte Kachel, wenn es nicht vollständig auf das Formular oder die Seite passt.

