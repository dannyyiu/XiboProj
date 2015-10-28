Attribute VB_Name = "PowerpointVB"
'@file PowerpointVB
'@brief Helper functions to reference PowerPoint VBA
'
Option Private Module


'@brief Return a Powerpoint application object
Public Function ApplicationObj() As Object

    Set ApplicationObj = CreateObject("PowerPoint.Application")
    
End Function


'@brief Return an active presentation object given an open application
Public Function ActivePresentation(ByVal Application As Object) As Object

    Set ActivePresentation = Application.ActivePresentation
    
End Function


'@brief Return a presentations object given an open application
Public Function Presentations(ByVal Application As Object) As Object

    Set Presentations = Application.Presentations
    
End Function
