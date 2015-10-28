VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} XiboForm 
   Caption         =   "Xibo: Insert Takedown Date"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   OleObjectBlob   =   "XiboForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "XiboForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@file XiboForm
'@brief XiboForm scripts


'@brief Initial values
Sub Initialize()
    
    'Output path display
    OutputPathLabel.Caption = Globals.OUTDIR
    
    'Default year
    YearText.Text = Year(Date)
    
    'Set focus on month textbox
    MonthText.SetFocus
    
End Sub


'@brief Event when focus moves out of year textbox
'
' Change color of text depending on if value is valid
'
Private Sub YearText_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    'Change color if invalid
    Call Verify.XiboFormValid
    
End Sub


'@brief Event when focus moves out of month textbox
'
' Pad "0" to left of text until 2 digits reached,
' then change color of text depending on if value is valid
'
Private Sub MonthText_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    'Pad 0's to the left of text until 2 digits reached
    MonthText.Text = LPadded(MonthText.Text, "0", 2)
    
    'Change color if invalid
    Call Verify.XiboFormValid
    
End Sub


'@brief Event when focus moves out of day textbox
'
' Pad "0" to left of text until 2 digits reached,
' then change color of text depending if value is valid
'
Private Sub DayText_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    
    'Pad "0" to left of text until 2 digits reached
    DayText.Text = LPadded(DayText.Text, "0", 2)
    
    'Change color if invalid
    Call Verify.XiboFormValid
    
End Sub


'@brief Generate button event
'
' If all fields are valid, save them to global variables
' and close this form. Otherwise, display a message.
'
Private Sub GenerateBtn_Click()

    'Save date. Generation is done by main script.
    If Verify.XiboFormValid() Then
        Globals.YYYY = YearText.Text
        Globals.MM = MonthText.Text
        Globals.DD = DayText.Text
        XiboForm.hide
    Else
        MsgBox "Some values still invalid!"
    End If
    
End Sub


'@brief (Helper) Padding to the left of string until certain length reached
'
'@param InputString: String to pad "0"s to
'@param PadChar: Character used for padding, length 1
'@param TotalLen:
Function LPadded(InputString As String, _
                 PadChar As String, _
                 TotalLen As Integer)
    
    'Return input string by default
    LPadded = Trim(InputString)
    
    If (Len(PadChar) = 1 And _
        TotalLen > Len(InputString)) Then
        'Requires padding
        
        Dim i As Integer
        'Loop through and pad 0 to left
        For i = 1 To (TotalLen - Len(InputString))
            LPadded = "0" & LPadded
        Next
    
    End If
    
End Function



