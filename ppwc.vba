' Thanks to QuickSortMultiDim by Dieter Otter (http://www.vbarchiv.net/tipps/tipp_1881.html)


Option Explicit

    Public pptApp As PowerPoint.Application
    Public pptSlide As Slide
    
    Public ppWidth As Long
    Public ppHeight As Long
    
    ' max distance to center of slide    
    Public maxDiag As Long
    
    Dim coord() As Long

    Public shapeX As Long
    Public shapeY As Long
    Public shapeWidth As Double
    Public shapeHeight As Double
    Public factorHeight As Double
    
    ' factorPerformance - if you want to speed up the creation process and dont care about un-clean results, decrease this number 
    Public factorPerformance As Double


Sub StartEngine()

    Dim i, j, k As Long
    Dim wordList As Variant
    Dim fontSize As Double
    Dim duration As Long
    Dim startTime As Date
    Dim wordSum As Long
    Dim topWords As Long
    
    startTime = Now()
    wordSum = Application.WorksheetFunction.Sum(Range("wordCount"))
    
    If wordSum = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
        
    ppWidth = Range("ppWidth")
    ppHeight = Range("ppHeight")
    
    fontSize = Range("fontSize")
    factorHeight = Range("factorHeight")
    factorPerformance = Range("factorPerformance")
        
    '-------------------------------#1 get words from source and count them
    wordList = getWordList
     

    '-------------------------------#2 new powerpoint instance 
    PowerPoint "open"
    maxDiag = ((ppWidth ^ 2 + ppHeight ^ 2) ^ 0.5) * 0.5
    
    On Error Resume Next
    
    ReDim coord(ppHeight * factorPerformance, ppWidth * factorPerformance)  
                 
    createShape CStr(wordList(UBound(wordList), 0)), fontSize * CDbl(wordList(UBound(wordList), 3)) / 100
        shapeX = (ppWidth * 0.5 - shapeWidth * 0.5)
        shapeY = (ppHeight * 0.5 - shapeHeight * 0.5)
    positionShape
    SaveShapePosition CLng(wordList(UBound(wordList), 4))
    
    '-------------------------------#3 put words to slide
    topWords = Range("topWords")
    If topWords > UBound(wordList, 1) Then topWords = UBound(wordList, 1)
    For i = UBound(wordList, 1) - 1 To UBound(wordList, 1) - topWords + 1 Step -1
        createShape CStr(wordList(i, 0)), fontSize * CDbl(wordList(i, 3)) / 100
        
        GetNextShapePosition CLng(wordList(i, 4))
        
        positionShape
        
        SaveShapePosition CLng(wordList(i, 4))
    Next i



    Dim y As Long
    Dim x As Long

    With Worksheets("Tabelle3")
    
        For y = 0 To ppHeight
            For x = 0 To ppWidth
                If coord(y, x) <> 0 Then .Cells(y, x) = coord(y, x)
            Next x
        Next y
    End With
    
    Application.ScreenUpdating = True
    
    duration = DateDiff("s", startTime, Now())
    Range("speed") = duration & " sec " & Chr(13) & Chr(10) & "(" & Format(duration / wordSum, "0.0") & " sec/Word)"
  
End Sub

Sub createShape(word As String, fontSize As Double)
    With pptSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, 10, 10)
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        
        With .TextFrame
            .TextRange.Text = word
            .TextRange.Font.Color = RGB(0, 0, 0)
            .TextRange.Paragraphs.ParagraphFormat.Alignment = ppAlignCenter
        End With
        
        With .TextFrame2
            .WordWrap = False
            .MarginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .marginTop = 0
            .TextRange.Font.Size = fontSize
            .AutoSize = msoAutoSizeShapeToFitText
        End With
        
        shapeHeight = .height
        shapeWidth = .width
    End With
End Sub


Sub GetNextShapePosition(group As Long)
    Dim x As Long
    Dim y As Long
    Dim k As Long
    Dim j As Long
    Dim n As Long
    
    Dim diag As Long
  
    Dim marginX As Long
    Dim marginY As Long
    
    marginX = (shapeWidth * 0.5) * factorPerformance
    marginY = (shapeHeight * 0.5) * factorHeight * factorPerformance
    
    On Error Resume Next
    
    '-------------------------------#1 get possible position from borders
    For y = 0 To ppHeight * factorPerformance
        For x = 0 To ppWidth * factorPerformance
            If coord(y, x) <= 1 Then coord(y, x) = 0
        Next x
    Next y

    n = 0   'counter to identify next step in the array
    For y = 0 To ppHeight * factorPerformance
        For x = 0 To ppWidth * factorPerformance
            If coord(y, x) > 1 Then
                For k = -1 * (marginY + 1) To marginY + 1
                    For j = -1 * (marginX + 1) To marginX + 1
                        If coord(y + k, x + j) = 0 Then
                            coord(y + k, x + j) = -1 * coord(y, x)
                            n = n + 1
                        End If
                    Next j
                Next k
                
                For k = -1 * marginY To marginY
                    For j = -1 * marginX To marginX
                        If coord(y + k, x + j) < 0 Then
                            coord(y + k, x + j) = 1
                            n = n - 1
                        End If
                    Next j
                Next k
            End If
        Next x
    Next y

    
    ReDim firstslot(n, 3) As Variant
    
    n = 0
    '-------------------------------#4 keep in mind possible position
    For y = 0 To ppHeight * factorPerformance
        For x = 0 To ppWidth * factorPerformance
            If coord(y, x) < 0 Then
            
                'distance to center of slide, decrease it to create a more compact cloud                
                diag = (((ppWidth * 0.5 * factorPerformance) - x) ^ 2 + ((ppHeight * 0.5 * factorPerformance) - y) ^ 2) ^ 0.5
                
                firstslot(n, 0) = x
                firstslot(n, 1) = y
                firstslot(n, 2) = diag
                n = n + 1
            End If
        Next x
    Next y

    'n = (UBound(firstSlot, 1) * Rnd) + 1  this was for testing purposes: randomize recocknition of next position
    
    'next position by next distance to center of the slide
    QuickSortMultiDim firstslot, 2
    
    shapeX = firstslot(1, 0) / factorPerformance
    shapeX = shapeX - shapeWidth * 0.5
    shapeY = firstslot(1, 1) / factorPerformance
    shapeY = shapeY - shapeHeight * 0.5

End Sub

Sub positionShape()
    With pptSlide.Shapes(pptSlide.Shapes.Count)
        .left = shapeX
        .top = shapeY
    End With
End Sub


Sub SaveShapePosition(group As Long)
    Dim x As Long
    Dim y As Long
    
    On Error Resume Next
    '-------------------------------#1 mark last word in array
     For y = shapeY * factorPerformance To (shapeY + shapeHeight) * factorPerformance
        For x = shapeX * factorPerformance To (shapeX + shapeWidth) * factorPerformance
            coord(y, x) = 1 + group
        Next x
    Next y
End Sub



Sub PowerPoint(action As String)
    If action = "open" Then
    '===========================================================================================================
        Set pptApp = New PowerPoint.Application
        With pptApp
            .Visible = msoTrue
            .Presentations.Add
            With .ActivePresentation
                Set pptSlide = .Slides.Add(index:=.Slides.Count + 1, Layout:=ppLayoutBlank)
                    If ppWidth = 0 Then ppWidth = .PageSetup.slideWidth Else .PageSetup.slideWidth = ppWidth
                    If ppHeight = 0 Then ppHeight = .PageSetup.slideHeight Else .PageSetup.slideHeight = ppHeight
            End With
        End With
        
    ElseIf action = "fill" Then
    '===========================================================================================================
    
        
    ElseIf action = "close" Then
    '===========================================================================================================
    
    End If
End Sub

Sub worksheet_change(ByVal Target As Range)
    Set Target = Intersect(Target, Range("sourceText"))
    If Target Is Nothing Then
        Exit Sub
    Else
        CreateWordList (Range("sourceText"))
    End If
End Sub

Sub Worksheet_SelectionChange(ByVal Target As Range)
  If Target.Address = Range("startEngine").Address Then StartEngine
End Sub

Function CreateWordList(words As String) As Variant
    Dim arr_words() As String
    ReDim arr_list(0)
    ReDim arr_count(0)
    Dim result()

    Dim i As Long
    Dim j As Long
    Dim inList As Boolean
    
    words = Replace(words, ".", " ")
    words = Replace(words, ",", " ")
    words = Replace(words, "  ", " ")
    words = Replace(words, Chr(10), "")
    words = UCase(Replace(words, Chr(13), ""))
    arr_words = Split(words, " ")

    
    For i = LBound(arr_words) To UBound(arr_words)
        inList = False
        For j = LBound(arr_list) To UBound(arr_list)
            If arr_words(i) = arr_list(j) Then
                inList = True
                arr_count(j) = CLng(arr_count(j)) + 1
            End If
        Next j
        If inList = False And arr_words(i) <> "" Then
            arr_list(UBound(arr_list)) = arr_words(i)
            arr_count(UBound(arr_count)) = 1
            ReDim Preserve arr_list(UBound(arr_list) + 1)
            ReDim Preserve arr_count(UBound(arr_count) + 1)
        End If
    Next i
       
    ReDim result(UBound(arr_list), 4)

    For i = LBound(arr_list) To UBound(arr_list) - 1
        result(i, 0) = arr_list(i)
        result(i, 1) = arr_count(i)
    Next i
    
    QuickSortMultiDim result, 1
    
    Range("wordList").ClearContents
    
    j = 1
    For i = UBound(result) To LBound(result) + 1 Step -1
        Range("wordList").Cells(j, 1) = result(i, 0)
        Range("wordList").Cells(j, 2) = result(i, 1)
        j = j + 1
    Next i

End Function


Function getWordList()
    Dim i As Integer
    Dim j As Integer
    Dim result()
    Dim arr_words() As String
    Dim wordCount As Long
    
    wordCount = Application.WorksheetFunction.CountA(Range("wordCount"))
    ReDim result(wordCount, 4)
    
    ActiveWorkbook.Worksheets("start").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("start").Sort.SortFields.Add Key:=Range("wordList").Columns(2), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("start").Sort
        .SetRange Range("wordList")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    For i = 1 To wordCount
        result(i, 0) = Range("wordList").Cells(i, 1)
        result(i, 1) = Range("wordList").Cells(i, 2)
        result(i, 2) = Round((result(i, 1) / wordCount) * 100, 2)
        result(i, 3) = Round(result(i, 1) / Range("wordList").Cells(1, 2) * 100, 0)
        
        If i > 0 Then
            If result(i - 1, 3) <> result(i, 3) Then
                result(i, 4) = result(i - 1, 4) + 1
            Else
                result(i, 4) = result(i - 1, 4)
            End If
        Else
            result(i, 4) = 1
        End If
    
        Range("wordList").Cells(i, 3) = result(i, 2)
        Range("wordList").Cells(i, 4) = result(i, 3)
        Range("wordList").Cells(i, 5) = result(i, 4)
    
    Next i
    
    QuickSortMultiDim result, 1
        
    getWordList = result
End Function


' http://www.vbarchiv.net/tipps/tipp_1881.html
' vSort: 2-dimensionales Array
' index: Spalte, nach der sortiert werden soll (1, 2, 3, ...)
Public Sub QuickSortMultiDim(vSort As Variant, _
  Optional ByVal index As Integer = 1, _
  Optional ByVal lngStart As Variant, _
  Optional ByVal lngEnd As Variant)
 
  ' Wird die Bereichsgrenze nicht angegeben,
  ' so wird das gesamte Array sortiert
 
  If IsMissing(lngStart) Then lngStart = LBound(vSort)
  If IsMissing(lngEnd) Then lngEnd = UBound(vSort)
 
  Dim i As Long
  Dim j As Long
  Dim h As Variant
  Dim x As Variant
  Dim u As Long
  Dim lb_dim As Integer
  Dim ub_dim As Integer

  ' Anzahl Elemente pro Datenzeile
  lb_dim = LBound(vSort, 2)
  ub_dim = UBound(vSort, 2)
 
  i = lngStart: j = lngEnd
  x = vSort((lngStart + lngEnd) / 2, index)
 
  ' Array aufteilen
  Do
 
    While (vSort(i, index) < x): i = i + 1: Wend
    While (vSort(j, index) > x): j = j - 1: Wend
 
    If (i <= j) Then
      ' Wertepaare miteinander tauschen
      For u = lb_dim To ub_dim
        h = vSort(i, u)
        vSort(i, u) = vSort(j, u)
        vSort(j, u) = h
      Next u
      i = i + 1: j = j - 1
    End If
  Loop Until (i > j)
 
  ' Rekursion (Funktion ruft sich selbst auf)
  If (lngStart < j) Then QuickSortMultiDim vSort, index, lngStart, j
  If (i < lngEnd) Then QuickSortMultiDim vSort, index, i, lngEnd
End Sub

Sub ClosePPT()
    Dim pptApp As PowerPoint.Application
    Set pptApp = New PowerPoint.Application
    pptApp.Quit
End Sub

