Option Explicit
Option Private Module

Function CustomMsgBox(Optional Msg = "Selecione o csv a importar", Optional Style = vbOKCancel)
    Dim Title, Help, Ctxt, MyString
    
    Title = "Manpower - Tiago Carvalho"
    CustomMsgBox = MsgBox(Msg, Style, Title)
    
End Function

Function TurnFilterOFF(Optional sh)
    sh.Activate
    If IsMissing(sh) Then
        With ThisWorkbook
            .Activate
            Set sh = .Sheets(1)
        End With
    End If
    
    If sh.AutoFilterMode Then
        sh.AutoFilterMode = False
    End If
End Function

Function SortTable(sh, tableName, Optional ColName = "Nº Caso")
    Dim tbl As ListObject
    Dim rng As Range
    
    sh.Activate
    
    Set tbl = sh.ListObjects(tableName)
    Set rng = Range(tableName & "[" & ColName & "]")
    
    With tbl.Sort
       .SortFields.Clear
       .SortFields.Add Key:=rng, SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
    
End Function

Function SelectFile(Optional FileType = ".csv")
    Dim FileNamePath
    
    FileNamePath = Application.GetOpenFilename
    
    If FileNamePath = False Then
        Exit Function
    End If
    
    If Right(FileNamePath, Len(FileType)) = FileType Then
        SelectFile = FileNamePath
    End If
    
End Function

Function CSV2Array(FileNamePath, Optional OtherActions) As Variant()
    Dim Response, WbDados
    Dim NumbRows As Long, NumbCols As Long
    
    Set WbDados = Workbooks.Open(FileNamePath, , , , , , , , , , , , , True)
    
    If Not IsMissing(OtherActions) Then
        Application.Run OtherActions
    End If
    
    NumbRows = WbDados.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
    NumbCols = WbDados.Sheets(1).Cells(1, Columns.Count).End(xlToLeft).Column
    
    CSV2Array = WbDados.Sheets(1).Cells(1, 1).Resize(NumbRows, NumbCols).Value
    
    WbDados.Close SaveChanges:=False
End Function

Function Sheet2Array(FileNamePath, Optional sh = 1, Optional OtherActions, Optional CloseFile = True)
    Dim Response, WbDados
    Dim NumbRows As Long, NumbCols As Long
    
    Set WbDados = Workbooks.Open(FileNamePath, , , , , , , , , , , , , True)
    
    If Not IsMissing(OtherActions) Then
        Application.Run OtherActions, sh
    End If
    
    NumbRows = WbDados.Sheets(sh).Cells(Rows.Count, 1).End(xlUp).Row
    NumbCols = WbDados.Sheets(sh).Cells(1, Columns.Count).End(xlToLeft).Column
    
    Sheet2Array = WbDados.Sheets(sh).Cells(1, 1).Resize(NumbRows, NumbCols).Value
    
    If CloseFile = True Then
        WbDados.Close SaveChanges:=False
    End If
End Function

Function Data2Array(sh)
    Dim NumbRows As Long, NumbCols As Long
    
    NumbRows = sh.Cells(Rows.Count, 1).End(xlUp).Row
    NumbCols = sh.Cells(1, Columns.Count).End(xlToLeft).Column
    
    Data2Array = sh.Cells(1, 1).Resize(NumbRows, NumbCols).Value
End Function

Function ArrayBinarySearch(SearchArray, ValueSearch, Optional SearchArrayCol = 1)
    Dim tempMinIndex As Long
    Dim tempMaxIndex As Long
    Dim tempMiddleIndex As Long
    
    If ValueSearch < SearchArray(LBound(SearchArray, SearchArrayCol), SearchArrayCol) Or ValueSearch > SearchArray(UBound(SearchArray, SearchArrayCol), SearchArrayCol) Then
        ArrayBinarySearch = False
        Exit Function
    End If
    
    
    tempMinIndex = LBound(SearchArray, SearchArrayCol)
    tempMaxIndex = UBound(SearchArray, SearchArrayCol)
    
    Do While tempMinIndex <= tempMaxIndex
        
        tempMiddleIndex = (tempMinIndex + tempMaxIndex) / 2
        
        If ValueSearch = SearchArray(tempMiddleIndex, SearchArrayCol) Then
            ArrayBinarySearch = True
            Exit Function
        
        ElseIf ValueSearch < SearchArray(tempMiddleIndex, SearchArrayCol) Then
            tempMaxIndex = tempMiddleIndex - 1
            
        Else
            tempMinIndex = tempMiddleIndex + 1
        End If
        
    Loop
    
    ArrayBinarySearch = False
    
End Function

Function SLA24h(StartDatetime, EndDatetime)
    Dim Start_Date As Date, End_Date As Date
    Dim FirstDayDuration As Double, DaysDurationInHours As Double
    Dim LastDayDuration As Double, Start_Time As Double, End_Time As Double

    If EndDatetime = Empty Or IsNull(EndDatetime) Then EndDatetime = Now()
    
    Start_Date = Int(StartDatetime)
    End_Date = Int(EndDatetime)
    Start_Time = StartDatetime - Start_Date
    End_Time = EndDatetime - End_Date
    
    If Start_Time < 8 / 24 Then Start_Time = 8 / 24
    If End_Time < 8 / 24 Then End_Time = 8 / 24
    
    FirstDayDuration = (24 - Start_Time * 24)
    LastDayDuration = (End_Time * 24 - 8)
    
    If End_Date - Start_Date > 1 Then
        DaysDurationInHours = (End_Date - Start_Date - 1) * 16
        SLA24h = DaysDurationInHours + FirstDayDuration + LastDayDuration
    ElseIf End_Date - Start_Date = 1 Then
        DaysDurationInHours = 0
        SLA24h = FirstDayDuration + LastDayDuration
    Else
        DaysDurationInHours = 0
        SLA24h = DaysDurationInHours + ((End_Time - Start_Time) * 24)
    End If
    
End Function

Function Sla8_21_Seg_Sex(StartDatetime, EndDatetime)
    Dim DiaStartDate As Date, DiaEndDate As Date, i As Date
    Dim HoraStartDate As Double, HoraEndDate As Double, tempHoras As Double
    Dim tempDiasDecorridos As Integer, diasEmHoras As Integer, tempDiaSemana As Integer
    
    If EndDatetime = Empty Or IsNull(EndDatetime) Then EndDatetime = Now()
    
    DiaStartDate = Int(StartDatetime)
    DiaEndDate = Int(EndDatetime)
    HoraStartDate = StartDatetime - Int(StartDatetime)
    HoraEndDate = EndDatetime - Int(EndDatetime)
    
'Caso a hora de inicio seja fora de horas, depois das 21h, passa para o dia seguinte
    If HoraStartDate > 21 / 24 Then
        DiaStartDate = DiaStartDate + 1
        HoraStartDate = 8 / 24
    End If
    
'Caso a data de inicio seja Sábado
'Adiciona 2 dias para começar na Segunda-Feira
'Caso a data de inicio seja Domingo
'Adiciona 1 dia para começar na Segunda-Feira
    Select Case Weekday(DiaStartDate, vbMonday)
        Case 6 'Sábado
            DiaStartDate = DiaStartDate + 2
            HoraStartDate = 8 / 24
        Case 7 'Domingo
            DiaStartDate = DiaStartDate + 1
            HoraStartDate = 8 / 24
    End Select
    
'Caso a hora de inicio seja antes da hora de inicio
'acerta para as 8h
    If HoraStartDate < 8 / 24 Then
        HoraStartDate = 8 / 24
    End If
    
'Contar dias decorridos
    For i = DiaStartDate To DiaEndDate
        If i = DiaEndDate Then Exit For
        tempDiaSemana = Weekday(i, vbMonday)
        If Not tempDiaSemana = 6 And Not tempDiaSemana = 7 Then
                tempDiasDecorridos = tempDiasDecorridos + 1
        End If
    Next i
    
    diasEmHoras = tempDiasDecorridos * 13 ' por serem 13 horas por dia
    tempHoras = (HoraEndDate - HoraStartDate) * 24
    
    Sla8_21_Seg_Sex = diasEmHoras + tempHoras
End Function
