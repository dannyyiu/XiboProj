Attribute VB_Name = "Verify"
'@file Verify
'@brief Exception and validation functions
Option Private Module


'@brief Return validity of chosen email item.
'
' Currently considers valid if:
' - only one email selected
' - PPT attachments are found
'
Function EmailValid() As Boolean

    ' Selection check: make sure it's only one selected email
    If OutlookVB.SelectionCount() = 0 Then
        ' Nothing selected, invalid.
        MsgBox "No email selected."
        EmailValid = False
        
    ElseIf OutlookVB.SelectionCount() > 1 Then
        ' More than one email/item selected, invalid. May change in future
        MsgBox "More than one email (or item) selected." & _
               "Please select only one."
        EmailValid = False
        
    Else
        ' Selection check passed.
        
        ' Attachment check: make sure there are valid attachments
        If OutlookVB.AttachmentCount() = 0 Then
            ' No attachments
            MsgBox "No attachments in selected email."
            EmailValid = False
        ElseIf OutlookVB.HasPPT() = False Then
            ' Attachments, but no PPT files
            MsgBox "No PPT attachments found."
            EmailValid = False
        Else
            ' Attachments, with PPT files
            EmailValid = True
        End If
        
    End If
    
End Function


'@brief XiboForm form validation, check all fields valid.
'
' If any field is invalid, color will change to red. If a
' field is empty or valid, color will change to black.
' Return True if all values valid, False otherwise
'
Function XiboFormValid() As Boolean
    
    Dim Year, Month, Day As String
    Year = XiboForm.YearText.Text
    Month = XiboForm.MonthText.Text
    Day = XiboForm.DayText.Text
    XiboFormValid = True 'True until found invalid
    
    ' Year check
    If Not (IsNumeric(Year) And Len(Year) = 4) Then
        ' Year invalid
        
        ' Change text color to red if text exists
        If Len(Year) > 0 Then
            XiboForm.YearText.ForeColor = &HC0&
        Else
            XiboForm.YearText.ForeColor = &H0&
        End If
        
        XiboFormValid = False
    Else
        XiboForm.YearText.ForeColor = &H0&
    End If
    
    ' Month check:
    If IsNumeric(Month) Then
        'Numeric
        
        If (Not InStr(Month, ".") And _
            CInt(Month) >= 1 And _
            CInt(Month) <= 12) Then
            ' Not decimal, between 1 and 12
        
            ' Valid, change to black text
            XiboForm.MonthText.ForeColor = &H0&
        Else
            ' Numeric but invalid
            XiboForm.MonthText.ForeColor = &HC0&
            XiboFormValid = False
            
        End If
        
    Else '/numeric
        ' Invalid, change text color to red if text exists
        If Len(Month) > 0 Then
            XiboForm.MonthText.ForeColor = &HC0&
        Else
            XiboForm.MonthText.ForeColor = &H0&
        End If
        
        XiboFormValid = False
        
    End If
    
    ' Day check
    If IsNumeric(Day) Then
        ' Numeric
        
        If (Not InStr(Day, ".") And _
            CInt(Day) >= 1 And _
            CInt(Day) <= 31) Then
            'Not decimal, between 1 and 31
            
            ' Valid, change text to black
            XiboForm.DayText.ForeColor = &H0&
        Else
            ' Numeric but invalid
            XiboForm.DayText.ForeColor = &HC0&
            XiboFormValid = False
            
        End If
        
    Else '/numeric
        ' Invalid, change text color to red if text exists
        If Len(Day) > 0 Then
            XiboForm.DayText.ForeColor = &HC0&
        Else
            XiboForm.DayText.ForeColor = &H0&
        End If
        
        XiboFormValid = False
        
    End If
        
End Function
