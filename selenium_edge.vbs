Sub PobierzDaneZEgdeDoExcela()
    Dim driver As New Selenium.EdgeDriver
    Dim tabela As Object
    Dim wiersz As Object
    Dim komorka As Object
    Dim ws As Worksheet
    Dim i As Long, j As Long
    
    ' Utwórz nowy arkusz na dane
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "DaneZEdge"
    
    ' Uruchom Edge i wejdź na stronę
    driver.Start "edge"
    driver.Get "https://www.example.com"
    
    ' Przykład: znajdź pierwszą tabelę na stronie
    Set tabela = driver.FindElementByTag("table")
    
    ' Przejdź po wierszach tabeli
    i = 1
    For Each wiersz In tabela.FindElementsByTag("tr")
        j = 1
        For Each komorka In wiersz.FindElementsByTag("td")
            ws.Cells(i, j).Value = komorka.Text
            j = j + 1
        Next komorka
        i = i + 1
    Next wiersz
    
    driver.Quit
    MsgBox "Dane pobrane do arkusza '" & ws.Name & "'"
End Sub
