Attribute VB_Name = "DataValidation"
Option Explicit

Private Function Validation_List(ByVal Args As Variant) As String
    Dim sVal As String
    Dim sOut As String
    Dim i As Long
    Dim j As Long
    Dim bAdd As Boolean
    
    For i = 1 To Args(0).Rows.Count
        sVal = CStr(Args(0).Rows(i).Value)

        If Len(sVal) > 0 Then
            If InStr(1, sVal, ",") > 0 Then
                sVal = Replace(sVal, ",", Chr(130))
            End If
            
            bAdd = True
            For j = LBound(Args) + 1 To UBound(Args) Step 2
                If Args(j) Is Nothing Then
                    bAdd = True
                ElseIf Args(j).Rows(i).Value = Args(j + 1) And Len(Args(j + 1)) > 0 Then
                    bAdd = True
                Else
                    bAdd = False
                    Exit For
                End If
            Next
            
            If bAdd Then
                If Len(sOut) = 0 Then
                    sOut = sVal
                Else
                    If InStr(1, sOut, sVal) = 0 Then
                        sOut = sOut & "," & sVal
                    End If
                End If
            End If
        End If
    Next
    
    Validation_List = sOut
End Function

Public Sub Validation_Add(ByRef Target_Cell As Range, ParamArray Args() As Variant)
    Dim sTmp As String
    Dim sSort As String
    
    'Sorts the List first based on first argument and requires sheet be the same table.
    Args(0).Worksheet.Range(Args(0).Worksheet.Name).Sort Key1:=Args(0), _
    Order1:=xlAscending, Header:=xlYes, OrderCustom:=1, MatchCase:=False, _
    Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
    
    sTmp = Validation_List(Args)
    
    Target_Cell.Validation.Delete
    If Len(sTmp) > 0 Then
        With Target_Cell.Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:=sTmp
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With
    Else
        Target_Cell.Validation.Delete
    End If
End Sub

