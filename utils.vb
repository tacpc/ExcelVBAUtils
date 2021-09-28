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
