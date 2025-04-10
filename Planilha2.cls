VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_change(ByVal target As Range)

' Initializing variables
Dim oldVal As String
Dim newVal As String

' Checking for changes in the drop down lists

Application.EnableEvents = True

If target.Address = "$J$21" Then 'Or target.Address = "$G$3" Then
    
    If target.SpecialCells(xlCellTypeAllValidation) Is Nothing Then
        Exit Sub
' Adding the multiple selected items to the list
    Else
        If target.Value = "" Then
            Exit Sub

        Else
            Application.EnableEvents = False
            newVal = target.Value
            Application.Undo
            oldVal = target.Value
            If oldVal = "" Then
                target.Value = newVal
            Else
                If InStr(oldVal, newVal) = 0 Then
                    target.Value = oldVal & "; " & vbCrLf & newVal '
                Else
                    target.Value = oldVal
                End If
            End If
        End If
    End If
End If

Application.EnableEvents = True

End Sub
