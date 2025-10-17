Sub PrzygotujDaneDoCopilota()
    ' Zaznacza zakres danych i kopiuje do schowka
    Range("A1:D100").Select
    Selection.Copy
    MsgBox "Dane skopiowane. Wklej je teraz do Copilota w Excelu i poproś o analizę."
End Sub
