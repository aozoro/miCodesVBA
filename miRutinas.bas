Sub listarComboBox(ByVal rng As Range, ByVal ComboBox As Object, Optional ByVal limpiar As Boolean = True)
'-------------------------------------------------------------------------------------------------------------------------
'Lista un Combobox desde un rango

'Variables
'rng            =   rango  de elementos que se llenaran en el combobox
'Combobox       =   objeto Combobox que se llenará
'limpiar        =   TRUE: resetea combobox, FALSE: se agregan sin borrar los anteriores
'-------------------------------------------------------------------------------------------------------------------------    
    Dim celda As Range
    
    If limpiar Then ComboBox.Clear
    
    For Each celda In rng
        ComboBox.AddItem celda.Value
    Next celda
End Sub

Sub listarComboBox_Criterio( _
    ByVal rng As Range, _
    ByVal ComboBox As Object, _
    ByVal rngCriterio As Range, _
    Optional ByVal criterio As String = "(Sin Filtrar)", _
    Optional ByVal igualdad As Boolean = True, _
    Optional ByVal limpiar As Boolean = True)
'-------------------------------------------------------------------------------------------------------------------------
'Lista un Combobox desde un rango según una condición

'Variables
'rng            =   rango  de elementos que se llenaran en el combobox
'Combobox       =   objeto Combobox que se llenará
'rngCriterio    =   rango de elementos que verificaran si van o no en el combobox
'criterio       =   criterio con el que se compara con igualdad o diferencia, el valor default quita cualquier criterio
'igualdad       =   TRUE    :   se incluyen los elementos de rng donde su respectivo rngCriterio sea igual al criterio
'                   FALSE   :   se incluyen los elemetnos de rng donde su repsectivo rngCriterio se diferente al criterio
'limpiar        =   TRUE: resetea combobox, FALSE: se agregan sin borrar los anteriores

'Rutinas auxioliares: listarCombobox()
'-------------------------------------------------------------------------------------------------------------------------
    Dim celda As Range
    Dim sw As Boolean
    Dim jj As Long
    
    If criterio = "(Sin Criterio)" Then
        Call listarComboBox(rng, ComboBox,limpiar)
        Exit Sub
    End If
    
    If limpiar Then ComboBox.Clear
    
    For Each celda In rngCriterio
        jj = jj + 1
        sw = celda.Value = criterio
        
        If Not igualdad Then sw = Not sw
        
        If sw Then
            ComboBox.AddItem rng.Cells(jj, 1)
        End If
    Next celda
End Sub

Function uFila(Optional ByVal Hoja As Worksheet, _
    Optional ByVal CriterioColumna = 1) As Long
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    With Hoja
        If .Cells(.Rows.Count, CriterioColumna).Value <> "" Then
            uFila = .Rows.Count
        Else
            uFila = .Cells(.Rows.Count, CriterioColumna).End(xlUp).row
        End If
    End With
End Function

Function uColumna(Optional ByVal Hoja As Worksheet, _
    Optional ByVal CriterioFila = 1) As Long
    
    If Hoja Is Nothing Then Set Hoja = ActiveSheet
    
    With Hoja
        If .Cells(CriterioFila, .Columns.Count).Value <> "" Then
            uColumna = .Columns.Count
        Else
            uColumna = .Cells(CriterioFila, .Columns.Count).End(xlToLeft).Column
        End If
    End With
End Function

Function ElegirCarpeta()
    Dim diaFolder As FileDialog

    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Show
    On Error Resume Next
    ElegirCarpeta = diaFolder.SelectedItems(1)
    On Error GoTo 0
End Function

Function Col_Num2Let(ByVal colNumber As Long) As String
    Dim ColumnNumber As Long
    Dim ColumnLetter As String

    Col_Num2Let = Split(Cells(1, colNumber).Address, "$")(1)
End Function

Function HojaExiste(ByVal nameHoja As String)
    Dim Hoja As Worksheet
    Dim aux As Boolean
    
    aux = False
    For Each Hoja In ThisWorkbook.Worksheets
        If Hoja.Name = nameHoja Then
            aux = True
        End If
    Next Hoja
    
    HojaExiste = aux
End Function

Sub ResetarHoja(ByVal nameHoja As String, Optional ByVal Libro As Workbook)
    Dim hojaNueva As Worksheet
    Dim nHojas As Long
    
    If Libro Is Nothing Then Set Libro = ThisWorkbook
    If HojaExiste(nameHoja) Then
        Application.DisplayAlerts = False
        Libro.Worksheets(nameHoja).Delete
        Application.DisplayAlerts = True
    End If
    
    nHojas = Libro.Worksheets.Count
    Set hojaNueva = Libro.Worksheets.Add(after:=Libro.Worksheets(nHojas))
    
    hojaNueva.Name = nameHoja
End Sub

Sub BorrarHoja(ByVal nameHoja As String, Optional ByVal Libro As Workbook)
    Dim hojaNueva As Worksheet
    
    If Libro Is Nothing Then Set Libro = ThisWorkbook
    If HojaExiste(nameHoja) Then
        Application.DisplayAlerts = False
        Libro.Worksheets(nameHoja).Delete
        Application.DisplayAlerts = True
    End If
End Sub

Function ArchivosCarpeta(Folder)
    Dim file As String
    Dim j As Integer
    Dim Retorno() As String
    
    file = Dir(Folder & "\" & "*.xls*")
    j = 0
    Do While file <> ""
        ReDim Preserve Retorno(j) As String
        Retorno(j) = ConcatenarFolderFile(Folder, file)
        j = j + 1
        file = Dir
    Loop
    
    ArchivosCarpeta = Retorno
End Function

Function ConcatenarFolderFile(ByVal Folder As String, ByVal file As String) As String
    If Right(Folder, 1) = "/" Then
        ConcatenarFolderFile = Folder & file
    Else
        ConcatenarFolderFile = Folder & "\" & file
    End If
End Function


Sub toValores(ByVal rng As Range)
    rng.Copy
    rng.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub

Sub RutinaPDF(ByVal Hoja As Worksheet, ByVal PathFolder As String, ByVal nombreFile As String, ByVal horizontal As Boolean)
    With Hoja
        .ResetAllPageBreaks
        
        If horizontal Then
            .PageSetup.Orientation = xlLandscape
        Else
            .PageSetup.Orientation = xlPortrait
        End If

        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .ExportAsFixedFormat Type:=xlTypePDF, Filename:=PathFolder & "\" & nombreFile & ".pdf", _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    End With
End Sub
