Sub FindNamesPESELsAndPriceWithInput_Finalll()          'this code works flawlessly
Application.ScreenUpdating = False                      'GM6 tested and approved
    
    Dim wordApp As Word.Application              'it is assigned to button #2
    Dim wordDoc As Word.Document                 '[0-9]{1;5}[ ]{1;2}/[0-9]{4}
    Dim excelApp As Excel.Application
    Dim mySheet As Excel.Worksheet
    Dim Para As Word.Paragraph
    Dim rng As Word.Range
    Dim fullName As Word.Range
    
    Dim pStart As Long
    Dim pEnd As Long
    Dim Length As Long
    Dim textToFind1 As String
    Dim textToFind2 As String
    Dim textToFind3 As String
    Dim textToFind4 As String
    Dim textToFind5 As String
    Dim textToFind6 As String
    Dim name As String
    Dim pesel As String
    Dim price As Double      'Dim price As Single    'Dim price As Long
    
    Dim startPos As Long
    Dim endPos As Long
    Dim parNmbr As Long
    Dim x As Long
    Dim flag As Boolean
    Dim scndRng As Range
    
    
        
    'Assigning object variables and values
    Set wordApp = GetObject(, "Word.Application")       'At its simplest, CreateObject creates an instance of an object,
    Set excelApp = GetObject(, "Excel.Application")     'whereas GetObject gets an existing instance of an object.
    Set wordDoc = wordApp.ActiveDocument
    Set mySheet = Application.ActiveWorkbook.ActiveSheet
    'Set MySheet = ExcelApp.ActiveWorkbook.ActiveSheet
    Set rng = wordApp.ActiveDocument.Content
    Set scndRng = ActiveSheet.Range("A10:J40").Find("cena", , xlValues)
    textToFind1 = "KRS 0000511671, REGON: 147269372, NIP: 5252588142," ' "KRS 0000609737, REGON 364061169, NIP 951-24-09-783,"    ' "REGON 364061169, NIP 951-24-09-783,"
    textToFind2 = "- ad."        'w umowach deweloperskich FW2 było "- ad."
    textToFind3 = "Tożsamość stawających"
    textToFind4 = "PESEL"
    textToFind5 = "cenę brutto w kwocie łącznej"
    textToFind6 = "cenę brutto w kwocie "
    x = 11
    
    'InStr function returns a Variant (Long) specifying the position of the first occurrence of one string within another.
    startPos = InStr(1, rng, textToFind1) - 1    'here we get 1410 or 1439 or 1555, we're looking 4 "TextToFind1"
    endPos = InStr(1, rng, textToFind2) - 1      'here we get 2315 or 2595 or 2207, we're looking 4 "- ad."
    
    parNmbr = rng.Paragraphs.Count
    Debug.Print "Total # of paragraphs = " & parNmbr
    
    'Calibrating the range from which the names will be pulled
        If startPos > 1 And endPos > 1 Then                                ' Exit Sub
            rng.SetRange Start:=startPos, End:=endPos                      ' startPos = 0 Or endPos = 0 Or endPos = -1
            Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
            Debug.Print rng
            rng.MoveStart wdParagraph, 1                                   ' Moves the start position of the specified range.
            Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
            Debug.Print rng
        Else
            endPos = InStr(1, rng, textToFind3) - 1 'here we get 2595 lub 2207, we're looking 4 "Tożsamość stawających"
            rng.SetRange Start:=startPos, End:=endPos                                  'startPos = 0 Or endPos = 0 Or endPos = -1
            Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
            Debug.Print rng
            rng.MoveStart wdParagraph, 1
            rng.MoveEnd wdParagraph, -1
            Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
            Debug.Print rng
        End If
        
    
    If startPos <= 0 Or endPos <= 0 Then
        MsgBox ("Client's names were not found!")
    'Client's names input
    Else
        For Each Para In rng.Paragraphs
            'get NAME
            name = Trim$(Para.Range.Words(3))    'Trim$ is the string version. Use this if you are using it on a string.
            Debug.Print name
            pStart = InStr(1, Para, ".") + 1      'here we get 3     'we should get 3
            Length = InStr(1, Para, ",") - pStart  'here we get 22/29/27/39 - 3
            Debug.Print Trim$(Mid(Para, pStart, Length))
            name = Trim$(Mid(Para, pStart, Length))
            'get PESEL
            pStart = InStr(1, Para, textToFind4) + Len(textToFind4) + 1       'textToFind4 = "PESEL"
            Length = InStr(pStart, Para, ",") - pStart  '51-pStart = 11
            Debug.Print Trim$(Mid(Para, pStart, Length))
            pesel = Trim$(Mid(Para, pStart, Length))
            x = x + 1
            'Cells(x, 1).Value = Trim(Mid(Para, pStart, Length))
            ActiveSheet.Cells(x, 1).Value = name
            ActiveSheet.Cells(x, 4).Value = pesel
        Next Para
    'End of client's names input
        
        
        'Extract the app price
        Set rng = wordApp.ActiveDocument.Content
        
        With rng.Find
            .Text = textToFind5           'TextToFind5 = "cenę brutto w kwocie łącznej"
            .MatchWildcards = False
            .MatchCase = False
            .Forward = True
            .Execute
            'Debug.Print rng
            If .Found = True Then
               Set rng = wordApp.ActiveDocument.Content
               'InStr returns a Long specifying the position of the first occurrence of one string within another.
               startPos = InStr(1, rng, textToFind5)      'here we get 47310, we're looking 4 "cenę brutto w kwocie łącznej"
               endPos = InStr(startPos, rng, ",00zł")     'here we get 47380, we're looking 4 ",00zł"
                    If endPos > startPos + 1000 Or endPos < startPos Then
                           rng.SetRange Start:=startPos - 1, End:=ActiveDocument.Range.End    'Resetting the rng.   'Sets the starting and ending character positions for an existing range.
                           Debug.Print rng
                              With rng.Find
                                   .Text = "[0-9]{3},[0-9]{2}zł"
                                   .MatchWildcards = True
                                   .MatchCase = False
                                   .Forward = True
                                   .Execute
                                   If .Found = True Then price = rng.Duplicate
                                       Debug.Print price
                                       Set rng = wordApp.ActiveDocument.Content
                                       startPos = InStr(1, rng, textToFind5)      'we're looking for "cenę brutto w kwocie "
                                       endPos = InStr(startPos, rng, price) + 6   'price.Text
                                       startPos = startPos + Len(textToFind5)     'len = 28
                                       Debug.Print Replace(Mid(rng, startPos, endPos - startPos), ".", "")
                                       'Set rng = wordApp.ActiveDocument.Content
                                       price = Replace(Mid(rng, startPos, endPos - startPos), ".", "")
                                       'price = Trim(price)
                                       Debug.Print price
                              End With
                     Else
                         startPos = startPos + Len(textToFind5)     'now start position is reassigned at 47331
                         'Debug.Print rng
                         Debug.Print Replace(Mid(rng, startPos, endPos - startPos), ".", "")             'up to this moment macro works fine while customers buy appartment in shares, like 1/2 or 1/3, in the next line it crashes.
                         price = Replace(Mid(rng, startPos, endPos - startPos), ".", "")
                         price = Trim(price)
                         Debug.Print price
                     End If
                     If Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 4 Then
                        mySheet.Range("F12:F13") = Range("AE12").Value      'Range("AE12") = "współwłasn"
                     ElseIf Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 6 Then
                        mySheet.Range("F12:F14") = Range("AE12").Value
                     Else
                        mySheet.Range("F12:F15") = Range("AE12").Value
                     End If
            Else
               Set rng = wordApp.ActiveDocument.Content     'this command without "set" will destroy the document's formating;
               With rng.Find
                .Text = textToFind6             '="cenę brutto w kwocie "
                .MatchWildcards = False
                .MatchCase = False
                .Forward = True
                .Execute
                  If .Found = True Then
                     Set rng = wordApp.ActiveDocument.Content
                     startPos = InStr(1, rng, textToFind6)      'we're looking 4 "cenę brutto w kwocie "
                     endPos = InStr(startPos, rng, ",00zł")     'here we get 47380, we're looking 4 ",00zł"
                        If endPos > startPos + 1000 Or endPos < startPos Then
                           rng.SetRange Start:=startPos - 1, End:=ActiveDocument.Range.End    'Resetting the rng.   'Sets the starting and ending character positions for an existing range.
                           Debug.Print rng
                              With rng.Find
                                   .Text = "[0-9]{3},[0-9]{2}zł"
                                   .MatchWildcards = True
                                   .MatchCase = False
                                   .Forward = True
                                   .Execute
                                   If .Found = True Then price = rng.Duplicate
                                       Debug.Print price
                                       Set rng = wordApp.ActiveDocument.Content
                                       startPos = InStr(1, rng, textToFind6)      'we're looking for "cenę brutto w kwocie "
                                       endPos = InStr(startPos, rng, price) + 6   'price.Text
                                       startPos = startPos + Len(textToFind6)     'len = 21
                                       Debug.Print Replace(Mid(rng, startPos, endPos - startPos), ".", "")
                                       price = Replace(Mid(rng, startPos, endPos - startPos), ".", "")
                                       price = Trim(price)
                                       Debug.Print price
                              End With
                         Else
                            startPos = startPos + Len(textToFind6)     ' + 21 characters     'now start position is reassigned at 47331.
                            Debug.Print Replace(Mid(rng, startPos, endPos - startPos), ".", "")
                            price = Replace(Mid(rng, startPos, endPos - startPos), ".", "")
                            price = Trim(price)
                            Debug.Print price
                         End If
                  End If
                End With
            End If
        End With
        'ActiveSheet.Cells(28, 5).Value = price
        Debug.Print scndRng.Address
        scndRng.Offset(0, 1) = price
    End If
    
    Application.ScreenUpdating = True
    'Replace(Mid(rng, startPos, endPos - startPos), ".", "").Find
End Sub
