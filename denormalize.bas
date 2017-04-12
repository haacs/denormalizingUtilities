Attribute VB_Name = "denormalizingUtilities"
Sub denormalize()
Attribute denormalize.VB_Description = "denormalizes"
Attribute denormalize.VB_ProcData.VB_Invoke_Func = " \n14"
'
' denormalize Makro
' splits content of cells in selected column at comma
' into array and inserts a new row for every element

'
    Dim c As Range
    Dim r As Range
    Dim ids() As String
    Dim id As String
    Dim id_count As Integer
    Dim first_Adress As Variant
    
    With ActiveCell.EntireColumn
        Set c = .Find(",", LookIn:=xlValues)
        If Not c Is Nothing Then
            ' save first_adress to prevent infinite loop
            first_Adress = c.Address
            Do
                ids = Split(CStr(c.Value), ",")
                For id_count = LBound(ids) To UBound(ids)
                    ' remove trailing and leading whitespaces
                    ' before writing to cell
                    c.Value = Trim(ids(id_count))
                    ' duplicate the entire row if there are more ids
                    If id_count < UBound(ids) Then
                        Set r = c.EntireRow
                        Application.CutCopyMode = False
                        r.Copy
                        r.Insert
                    End If
                Next
                ' try to continue search
                Set c = .FindNext(c)
                ' only try to get the adress if c is not nothing
                If c Is Nothing Then
                    next_adress = first_address
                Else
                    next_address = c.Address
                End If
            Loop While Not c Is Nothing And next_address <> first_address
        End If
    End With
End Sub

Sub replaceLineBreak()
'
' replaceLineBreak Makro
' replaces all line break characters in
' selection with comma

'
    For Each oneCell In Selection
        ' Take String from cell and replace vbLf with vbCr
        strWithoutVbLf = Application.Substitute(CStr(oneCell.Value), vbLf, vbCr)
        ' Replace vbCr with "," and write back into cell
        oneCell.Value = Application.Substitute(strWithoutVbLf, vbCr, ",")
    Next oneCell
End Sub


