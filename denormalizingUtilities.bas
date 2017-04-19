Attribute VB_Name = "denormalizingUtilities1"
Sub denormalize()
Attribute denormalize.VB_Description = "Utilities for denormalizing"
Attribute denormalize.VB_ProcData.VB_Invoke_Func = " \n14"
'
' denormalize Makro
' splits content of cells in selection at comma
' into array and inserts a new row for every element

'
    Dim c As Range
    Dim r As Range
    Dim ids() As String
    Dim id As String
    Dim curr_id As String
    
    With Selection
        Set c = .Find(",", LookIn:=xlValues)
        If Not c Is Nothing Then
            ' save first_adress to prevent infinite loop
            first_row_index = Selection.Item(1).Row
            Do
                previous_row_index = c.Row
                ' continuously output the current row to the status bar
                Application.StatusBar = "Current Row is " & c.Row
                ids = Split(CStr(c.Value), ",")
                For id_count = LBound(ids) To UBound(ids)
                    ' remove trailing and leading whitespaces
                    curr_id = Trim(ids(id_count))
                    ' only insert for ids that are not empty
                    If Not curr_id = "" Then
                        c.Value = curr_id
                        ' duplicate the entire row if there are more ids
                        If id_count < UBound(ids) Then
                            Set r = c.EntireRow
                            Application.CutCopyMode = False
                            r.Copy
                            r.Insert
                        End If
                    End If
                Next
                ' try to continue search
                Set c = .FindNext(c)
                ' only try to get the adress if c is not nothing
                If c Is Nothing Then
                    next_row_index = previous_row_index
                Else
                    next_row_index = c.Row
                End If
                DoEvents
            Loop While Not c Is Nothing And next_row_index > previous_row_index
        End If
    End With
    Application.StatusBar = False
End Sub

Sub replaceLineBreak()
'
' replaceLineBreak Makro
' replaces all line break characters in
' selection with comma

'
    Dim c As Range
    Dim strng As String
    
    With Selection
        Set c = .Find(vbCr, LookIn:=xlValues)
        If Not c Is Nothing Then
            first_Adress = c.Adress
            Do
                Application.StatusBar = "First Loop, Current Row is " & c.Row
                strng = Application.Substitute(CStr(c.Value), vbCr, vbLf)
                c.Value = Application.Substitute(CStr(strng), vbLf, ",")
                Set c = .FindNext(c)
                If c Is Nothing Then Exit Do
                next_address = c.Address
                DoEvents
            Loop While next_address <> first_address
        End If
        
        Set c = .Find(vbLf, LookIn:=xlValues)
        If Not c Is Nothing Then
            first_Adress = c.Address
            Do
                Application.StatusBar = "Second Loop, Current Row is " & c.Row
                strng = Application.Substitute(CStr(c.Value), vbCr, vbLf)
                c.Value = Application.Substitute(CStr(strng), vbLf, ",")
                Set c = .FindNext(c)
                If c Is Nothing Then Exit Do
                next_address = c.Address
                DoEvents
            Loop While next_address <> first_address
        End If
    End With
    Application.StatusBar = False
End Sub

Sub setKitID()
'
' setKitId Makro
' sets the same kit_id for rows, where the column on its right
' contains the same value as the above

'
    Dim c As Range
    Dim c_kit_id As Range
    Dim above_c As Range
    Dim above_c_kit_id As Range
    Dim rng As Range
    
    Dim col_offset_to_kit_id As Integer
    col_offset_to_kit_id = 1
    
    Dim id_count As Integer
    Set rng = Selection.Columns(col_offset_to_kit_id).EntireColumn
    id_count = 0
    Set rng = Nothing
    
    l_limit = Range("A1").SpecialCells(xlCellTypeLastCell).Row
    
    For Each c In Selection.Columns(1).Cells
        If c.Row > 1 And CStr(c.Value) <> "" Then
            Set c_kit_id = c.Offset(0, col_offset_to_kit_id)
            Set above_c = c.Offset(-1, 0)
            Set above_c_kit_id = c.Offset(-1, col_offset_to_kit_id)
            Application.StatusBar = "Current Row is " & c.Row
            
            If Not c.Value <> above_c.Value Then
                c_kit_id.Value = above_c_kit_id.Value
            Else
                id_count = id_count + 1
                c_kit_id.Value = id_count
            End If
            DoEvents
        End If
    Next c
    Application.StatusBar = False
End Sub
