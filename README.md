# SnakesAndLadders-VBA
Popular board game "Snakes and Ladders" translated into VBA language with Excel interface.

Here is the whole code bellow, so you don't have to download Excel file.

--------------------------------------------------------------------------------------------------

Public indexP1, indexP2, indexP3, indexP4 As Integer

Public plansza As Variant

Public gracz1Nazwa, gracz2Nazwa, gracz3Nazwa, gracz4Nazwa As String
Public liczbaGraczy As Integer
Public kolejnosc As Integer

Public Function RzutKostka() As Integer
    
    Dim wynik As Integer
    wynik = Round(Rnd() * 5 + 1, 0)
    
    Cells(4, 14).Value = wynik
    RzutKostka = wynik
    
End Function


Sub Reset()
    
    Range("M6:N9").ClearContents
    
    SprawdzanieLiczbyGraczy
    
    Range("N2").Value = gracz1Nazwa
    Range("N2").Interior.Color = vbRed
    
    Set p1 = Arkusz1.Shapes("Player1")
    Dim p1Pozycja As Range
    Set p1Pozycja = Arkusz1.Range("A10")
    p1.Left = p1Pozycja.Left
    p1.Top = p1Pozycja.Top
    
    Set p2 = Arkusz1.Shapes("Player2")
    Dim p2Pozycja As Range
    Set p2Pozycja = Arkusz1.Range("A10")
    p2.Left = p2Pozycja.Left
    p2.Top = p2Pozycja.Top
    
    Set p3 = Arkusz1.Shapes("Player3")
    Dim p3Pozycja As Range
    Set p3Pozycja = Arkusz1.Range("A10")
    p3.Left = p3Pozycja.Left
    p3.Top = p3Pozycja.Top
    
    Set p4 = Arkusz1.Shapes("Player4")
    Dim p4Pozycja As Range
    Set p4Pozycja = Arkusz1.Range("A10")
    p4.Left = p4Pozycja.Left
    p4.Top = p4Pozycja.Top
    
    indexP1 = 0
    indexP2 = 0
    indexP3 = 0
    indexP4 = 0
    
    Range("N4").Value = ""
    Range("N7").Value = 0
    Range("N6").Value = 0
    Range("N8").Value = 0
    Range("N9").Value = 0
    
    kolejnosc = 0
    
End Sub

Public Function PulapkaLubDrabina(ByVal wynik As Integer) As Integer
    
    Select Case wynik
        Case Is = 2
            wynik = wynik + 19
            MsgBox "Trafiles na drabinke :). Przechodzisz na poziom 22."
            PulapkaLubDrabina = wynik
        Case Is = 8
            wynik = wynik + 19
            MsgBox "Trafiles na drabinke :). Przechodzisz na poziom 28."
            PulapkaLubDrabina = wynik
        Case Is = 13
            wynik = wynik - 10
            MsgBox "Ups! Trafiles na weza. Cofasz sie na poziom 4."
            PulapkaLubDrabina = wynik
        Case Is = 24
            wynik = wynik + 19
            MsgBox "Trafiles na drabinke :). Przechodzisz na poziom 44."
            PulapkaLubDrabina = wynik
        Case Is = 32
            wynik = wynik - 6
            MsgBox "Ups! Trafiles na weza. Cofasz sie na poziom 27."
            PulapkaLubDrabina = wynik
        Case Is = 40
            wynik = wynik + 21
            MsgBox "Trafiles na drabinke :). Przechodzisz na poziom 62."
            PulapkaLubDrabina = wynik
        Case Is = 47
            wynik = wynik + 21
            MsgBox "Trafiles na drabinke :). Przechodzisz na poziom 69."
            PulapkaLubDrabina = wynik
        Case Is = 50
            wynik = wynik - 2
            MsgBox "Ups! Trafiles na weza. Cofasz sie na poziom 49."
            PulapkaLubDrabina = wynik
        Case Is = 65
            wynik = wynik + 8
            MsgBox "Trafiles na drabinke :). Przechodzisz na poziom 74."
            PulapkaLubDrabina = wynik
        Case Is = 75
            wynik = wynik - 30
            MsgBox "Ups! Trafiles na weza. Cofasz sie na poziom 46."
            PulapkaLubDrabina = wynik
        Case Is = 77
            wynik = wynik + 19
            MsgBox "Trafiles na drabinke :). Przechodzisz na poziom 97."
            PulapkaLubDrabina = wynik
        Case Is = 81
            wynik = wynik - 24
            MsgBox "Ups! Trafiles na weza. Cofasz sie na poziom 58."
            PulapkaLubDrabina = wynik
        Case Is = 91
            wynik = wynik - 24
            MsgBox "Ups! Trafiles na weza. Cofasz sie na poziom 68."
            PulapkaLubDrabina = wynik
        Case Else
            PulapkaLubDrabina = wynik
    End Select
    
End Function

Sub PrzebiegGry()
    
    Select Case kolejnosc
        Case Is = 0
            gracz1
        Case Is = 1
            gracz2
        Case Is = 2
            gracz3
        Case Is = 3
            gracz4
        End Select
    
End Sub

Sub SprawdzanieLiczbyGraczy()
     
     liczbaGraczy = Application.InputBox("Podaj liczb« graczy od 2 do 4", Type:=1)
     Do While liczbaGraczy < 2 Or liczbaGraczy > 4
        liczbaGraczy = Application.InputBox("Podaj liczb« graczy od 2 do 4", Type:=1)
        Loop
    
     Select Case liczbaGraczy
        Case Is = 2
        gracz1Nazwa = InputBox("Podaj nazwe pierwszego gracza")
        gracz2Nazwa = InputBox("Podaj nazwe drugiego gracza")
        Range("M6").Value = "Punkty " & gracz1Nazwa
        Range("M7").Value = "Punkty " & gracz2Nazwa
        Case Is = 3
        gracz1Nazwa = InputBox("Podaj nazwe pierwszego gracza")
        gracz2Nazwa = InputBox("Podaj nazwe drugiego gracza")
        gracz3Nazwa = InputBox("Podaj nazwe trzeciego gracza")
        Range("M6").Value = "Punkty " & gracz1Nazwa
        Range("M7").Value = "Punkty " & gracz2Nazwa
        Range("M8").Value = "Punkty " & gracz3Nazwa
        Case Is = 4
        gracz1Nazwa = InputBox("Podaj nazwe pierwszego gracza")
        gracz2Nazwa = InputBox("Podaj nazwe drugiego gracza")
        gracz3Nazwa = InputBox("Podaj nazwe trzeciego gracza")
        gracz4Nazwa = InputBox("Podaj nazwe czwartego gracza")
        Range("M6").Value = "Punkty " & gracz1Nazwa
        Range("M7").Value = "Punkty " & gracz2Nazwa
        Range("M8").Value = "Punkty " & gracz3Nazwa
        Range("M9").Value = "Punkty " & gracz4Nazwa
        End Select
     
End Sub


Sub gracz1()
    
    Dim wynikKostki As Integer
    wynikKostki = RzutKostka()
    indexP1 = indexP1 + wynikKostki

    plansza = Array("A10", "B10", "C10", "D10", "E10", "F10", "G10", "H10", "I10", "J10", _
                        "J9", "I9", "H9", "G9", "F9", "E9", "D9", "C9", "B9", "A9", _
                        "A8", "B8", "C8", "D8", "E8", "F8", "G8", "H8", "I8", "J8", _
                        "J7", "I7", "H7", "G7", "F7", "E7", "D7", "C7", "B7", "A7", _
                        "A6", "B6", "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6", _
                        "J5", "I5", "H5", "G5", "F5", "E5", "D5", "C5", "B5", "A5", _
                        "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", "I4", "J4", _
                        "J3", "I3", "H3", "G3", "F3", "E3", "D3", "C3", "B3", "A3", _
                        "A2", "B2", "C2", "D2", "E2", "F2", "G2", "H2", "I2", "J2", _
                        "J1", "I1", "H1", "G1", "F1", "E1", "D1", "C1", "B1", "A1")
                        
    indexP1 = PulapkaLubDrabina(indexP1)
    
    If indexP1 > 99 Then
        indexP1 = indexP1 - wynikKostki
    ElseIf indexP1 = 99 Then
        MsgBox "Wygral gracz " & gracz1Nazwa & "!"
    End If
    
    Dim p1Pozycja As Range
    Set p1Pozycja = Arkusz1.Range(plansza(indexP1))
    
    Set p1 = Arkusz1.Shapes("Player1")
    p1.Left = p1Pozycja.Left
    p1.Top = p1Pozycja.Top
    
    Range("N6").Value = indexP1 + 1
    Range("N2").Value = gracz2Nazwa
    Range("N2").Interior.ColorIndex = 13
    
    kolejnosc = kolejnosc + 1
End Sub
Sub gracz2()
    
    Dim wynikKostki As Integer
    wynikKostki = RzutKostka()
    indexP2 = indexP2 + wynikKostki
    
    plansza = Array("A10", "B10", "C10", "D10", "E10", "F10", "G10", "H10", "I10", "J10", _
                        "J9", "I9", "H9", "G9", "F9", "E9", "D9", "C9", "B9", "A9", _
                        "A8", "B8", "C8", "D8", "E8", "F8", "G8", "H8", "I8", "J8", _
                        "J7", "I7", "H7", "G7", "F7", "E7", "D7", "C7", "B7", "A7", _
                        "A6", "B6", "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6", _
                        "J5", "I5", "H5", "G5", "F5", "E5", "D5", "C5", "B5", "A5", _
                        "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", "I4", "J4", _
                        "J3", "I3", "H3", "G3", "F3", "E3", "D3", "C3", "B3", "A3", _
                        "A2", "B2", "C2", "D2", "E2", "F2", "G2", "H2", "I2", "J2", _
                        "J1", "I1", "H1", "G1", "F1", "E1", "D1", "C1", "B1", "A1")
                        
    indexP2 = PulapkaLubDrabina(indexP2)
    
    If indexP2 > 99 Then
        indexP2 = indexP2 - wynikKostki
    ElseIf indexP2 = 99 Then
        MsgBox "Wygral gracz " & gracz2Nazwa & "!"
    End If
    
    Dim p2Pozycja As Range
    Set p2Pozycja = Arkusz1.Range(plansza(indexP2))
    
    Set p2 = Arkusz1.Shapes("Player2")
    p2.Left = p2Pozycja.Left
    p2.Top = p2Pozycja.Top
    
    Range("N7").Value = indexP2 + 1
    
    If liczbaGraczy = 2 Then
        kolejnosc = kolejnosc - 1
        Range("N2").Value = gracz1Nazwa
        Range("N2").Interior.Color = vbRed
    Else
        kolejnosc = kolejnosc + 1
        Range("N2").Value = gracz3Nazwa
        Range("N2").Interior.Color = vbYellow
    End If
End Sub
Sub gracz3()
    
    Dim wynikKostki As Integer
    wynikKostki = RzutKostka()
    indexP3 = indexP3 + wynikKostki
    
    plansza = Array("A10", "B10", "C10", "D10", "E10", "F10", "G10", "H10", "I10", "J10", _
                        "J9", "I9", "H9", "G9", "F9", "E9", "D9", "C9", "B9", "A9", _
                        "A8", "B8", "C8", "D8", "E8", "F8", "G8", "H8", "I8", "J8", _
                        "J7", "I7", "H7", "G7", "F7", "E7", "D7", "C7", "B7", "A7", _
                        "A6", "B6", "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6", _
                        "J5", "I5", "H5", "G5", "F5", "E5", "D5", "C5", "B5", "A5", _
                        "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", "I4", "J4", _
                        "J3", "I3", "H3", "G3", "F3", "E3", "D3", "C3", "B3", "A3", _
                        "A2", "B2", "C2", "D2", "E2", "F2", "G2", "H2", "I2", "J2", _
                        "J1", "I1", "H1", "G1", "F1", "E1", "D1", "C1", "B1", "A1")
                        
    indexP3 = PulapkaLubDrabina(indexP3)
    
    If indexP3 > 99 Then
        indexP3 = indexP3 - wynikKostki
    ElseIf indexP3 = 99 Then
        MsgBox "Wygral gracz " & gracz3Nazwa & "!"
    End If
    
    Dim p3Pozycja As Range
    Set p3Pozycja = Arkusz1.Range(plansza(indexP3))
    
    Set p3 = Arkusz1.Shapes("Player3")
    p3.Left = p3Pozycja.Left
    p3.Top = p3Pozycja.Top
    
    Range("N8").Value = indexP3 + 1
    
    If liczbaGraczy = 3 Then
        kolejnosc = kolejnosc - 2
        Range("N2").Value = gracz1Nazwa
        Range("N2").Interior.Color = vbRed
    Else
        kolejnosc = kolejnosc + 1
        Range("N2").Value = gracz4Nazwa
        Range("N2").Interior.Color = vbBlue
    End If
End Sub

Sub gracz4()
    
    Dim wynikKostki As Integer
    wynikKostki = RzutKostka()
    indexP4 = indexP4 + wynikKostki
    
    plansza = Array("A10", "B10", "C10", "D10", "E10", "F10", "G10", "H10", "I10", "J10", _
                        "J9", "I9", "H9", "G9", "F9", "E9", "D9", "C9", "B9", "A9", _
                        "A8", "B8", "C8", "D8", "E8", "F8", "G8", "H8", "I8", "J8", _
                        "J7", "I7", "H7", "G7", "F7", "E7", "D7", "C7", "B7", "A7", _
                        "A6", "B6", "C6", "D6", "E6", "F6", "G6", "H6", "I6", "J6", _
                        "J5", "I5", "H5", "G5", "F5", "E5", "D5", "C5", "B5", "A5", _
                        "A4", "B4", "C4", "D4", "E4", "F4", "G4", "H4", "I4", "J4", _
                        "J3", "I3", "H3", "G3", "F3", "E3", "D3", "C3", "B3", "A3", _
                        "A2", "B2", "C2", "D2", "E2", "F2", "G2", "H2", "I2", "J2", _
                        "J1", "I1", "H1", "G1", "F1", "E1", "D1", "C1", "B1", "A1")
                        
    indexP4 = PulapkaLubDrabina(indexP4)
    
    If indexP4 > 99 Then
        indexP4 = indexP4 - wynikKostki
    ElseIf indexP4 = 99 Then
        MsgBox "Wygral gracz " & gracz4Nazwa & "!"
    End If
    
    Dim p4Pozycja As Range
    Set p4Pozycja = Arkusz1.Range(plansza(indexP4))
    
    Set p4 = Arkusz1.Shapes("Player4")
    p4.Left = p4Pozycja.Left
    p4.Top = p4Pozycja.Top
    
    Range("N9").Value = indexP4 + 1
    Range("N2").Value = gracz1Nazwa
    Range("N2").Interior.Color = vbRed
    
    
    kolejnosc = kolejnosc - 3
    
End Sub

