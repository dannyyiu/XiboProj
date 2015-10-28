Attribute VB_Name = "OutlookVB"
'@file OutlookVB
'@brief Helper functions to reference Outlook VBA
Option Private Module


'@brief Return current explorer instance
Public Function CurrentExplorer() As Outlook.Explorer

    Set CurrentExplorer = Application.ActiveExplorer
    
End Function


'@brief Return list of current selected items
Function CurrentSelection() As Outlook.Selection

    Set CurrentSelection = CurrentExplorer().Selection

End Function


'@brief Return the number of current selected items
Function SelectionCount() As Integer
    
    SelectionCount = CurrentSelection.Count
    
End Function


' =============== Single Selection helper functions ================


'@brief Return Attachments of current selection
Function Attachments()
    
    Set Attachments = CurrentSelection().item(1).Attachments
    
End Function


'@brief Return the number of attachments in current selection
Function AttachmentCount() As Integer

    AttachmentCount = Attachments().Count

End Function


'@brief Return True if email attachments include .ppt files
Function HasPPT() As Boolean

    Dim Attachment As Outlook.Attachment
    HasPPT = False
    ' Loop through each attachment
    For Each Attachment In Attachments()
        ' Check for .ppt extention. Will work for .pptx
        If (InStr(Attachment.DisplayName, ".ppt") Or _
            InStr(Attachment.DisplayName, ".PPT")) Then
            HasPPT = True
        End If
    Next
    
End Function


'@brief Save all attachments without modifications
Sub SaveAttachments()
    
    ' Keep track of saved files to ensure uniqueness
    Dim SavedFiles As Object
    Set SavedFiles = CreateObject("Scripting.Dictionary")
    
    'Download attachments
    Dim Attachment As Outlook.Attachment
    For Each Attachment In Attachments()
        If Not SavedFiles.Exists(Attachment.DisplayName) Then
            Attachment.SaveAsFile (Globals.SAVEDIR & Attachment.DisplayName)
        End If
        
    Next
    
End Sub

