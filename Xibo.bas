Attribute VB_Name = "Xibo"
'@file Xibo
'@brief Generate files for Xibo uploading based on attachments
'
' Download attachments in current email if it contains PPT files,
' and use Powerpoint VBA to export it accordingly. Saves finished
' items in OUTDIR specified in global constants below. All
' configurations can be set in Globals.
'
'@notes
' Attachment PPT file extensions flexible, as long as it starts
' with either "ppt" or "PPT"
'
' PPT filenames from attachments cannot have "." in them besides
' ".ppt", otherwise output file will be saved with the first
' section. For example, "this.file.ppt" will be converted to
' "this.png". See subroutine Generate() remove extension part.
'


'@brief Main script
Public Sub XiboMain()
    
    If EmailValid() Then
        
        ' Save attachments to Globals.SAVEDIR
        Call ClearDir(Globals.SAVEDIR)  'Delete existing contents
        OutlookVB.SaveAttachments 'Save attachments
        
        ' Initialize XiboForm with some initial values
        XiboForm.Initialize
        ' Show XiboForm
        XiboForm.Show 'Date values will be saved to globals
        
        ' Folder of saved attachments
        Dim File As Variant
        File = Dir(Globals.SAVEDIR) 'First filename
            
        ' Validate:
        ' XiboForm saved values (in case window closed)
        ' If attachments were saved
        If (Verify.XiboFormValid() And _
            Len(Globals.YYYY) > 0 And _
            Len(Globals.MM) > 0 And _
            Len(Globals.DD) > 0 And _
            File <> "") Then
            ' All values valid, ready for file conversion.
            
            ' Delete existing content from output folder
            Call ClearDir(Globals.OUTDIR)
            
            ' Loop through saved attachments and process
            Dim PPTApplication As Object 'Holder for the ppt instance
            While (File <> "")
            
                'Generate PNG and save to OUTDIR
                Set PPTApplication = Generate(File)
                File = Dir ' Iterate to next file
                
            Wend
            'Quitting only after loop for better performance
            PPTApplication.Quit
            
            ' Create text file with embedded HTML
            Call CreateTextFile
            
            ' Complete
            If Globals.OPEN_OUTDIR Then
                ' Open directory containing the generated files
                Shell "explorer """ & _
                      Globals.OUTDIR & "", vbNormalFocus
            Else
                ' If set to False in Globals, just display a message.
                MsgBox "Task Complete"
            End If
            
            
        End If '/xiboform valid check
        
    End If '/email valid check
    
    
End Sub


' ======================== Helper functions =======================


'@breif Generate PNG and txt files to OUTDIR
'
' Open a PPT file from SAVEDIR with Powerpoint, export to PNG
' with correct size (set in Globals), rename with date values
' (set in Globals), and save them to OUTDIR. Generate a text
' file (filename set in Globals) with embedded HTML, also save
' to OUTDIR. Return PPT Application instance.
'
'@params File: filename of attachment in SAVEDIR
'
Private Function Generate(ByVal File As String)
    
    ' Declare PPT objects
    Dim PPTApplication, _
        Presentations, _
        CurrentPPT As Object
        
    ' Set Application object
    Set PPTApplication = _
        PowerpointVB.ApplicationObj()
    ' Set Presentations object
    Set Presentations = _
        PowerpointVB.Presentations(PPTApplication)
        
    ' Open presentation
    Presentations.Open _
        FileName:=Globals.SAVEDIR & File
        
    ' Set current active presentation
    Set CurrentPPT = _
        PowerpointVB.ActivePresentation(PPTApplication)
        
    ' Set slide dimensions
    With CurrentPPT.PageSetup
        .SlideWidth = Globals.SLIDE_WIDTH
        .SlideHeight = Globals.SLIDE_HEIGHT
    End With
    
    'Save slides as PNG files in OUTDIR
    Dim SlideCount As Integer
    Dim OutputFile As String 'Filename for output file
    SlideCount = CurrentPPT.Slides.Count ' count slides
    OutputFile = Split(File, ".")(0) 'remove extension
    If SlideCount = 1 Then
        'Single Slide, save without slide number
        ' ie. 2010-03-12_powerpointname.png
        
        ' Export
        With CurrentPPT.Slides(1)
            .Export Globals.OUTDIR & _
                Globals.YYYY & "-" & _
                Globals.MM & "-" & _
                Globals.DD & "_" & _
                OutputFile & _
                "." & LCase(Globals.PPT_EXPORT), _
                Globals.PPT_EXPORT
        End With
        
    ElseIf SlideCount > 1 Then
        'Multiple slides, save with slide number
        ' ie. 2014-03-12_powerpointname(3).png
        
        ' Loop through slides
        For i = 1 To SlideCount
        
            'Export
            CurrentPPT.Slides(i).Export _
                Globals.OUTDIR & _
                Globals.YYYY & "-" & _
                Globals.MM & "-" & _
                Globals.DD & "_" & _
                OutputFile & _
                "(" & i & ")." & _
                LCase(Globals.PPT_EXPORT), _
                Globals.PPT_EXPORT
        Next
    
    End If '/slidecount check
    
    ' Close PPT file
    'PPTApplication.Quit
    CurrentPPT.Close
    Set Generate = PPTApplication 'Returning this for quitting
    
End Function


'@brief Delete all files in a directory
Private Sub ClearDir(DirName As String)
    
    On Error Resume Next
    Kill DirName & "*.*"
    
End Sub


'@brief Generate a text file formatted with HTML (from original version)
'
' Follwing sub will create a textfile that has embed codes
' for each files created so that the user can simply copy and
' paste the text into the Xibo window for easier upload.
'
Private Sub CreateTextFile()

    'Object for creating text file
    Dim fso As Object '10/27: Previously undeclared variable fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim TextFile As Variant
    'Quotation mark
    Dim Q As Variant
    Q = """" '10/27: Changed from Chr$(34)

    Dim FileName As Variant

    'Search png files at the following directory
    FileName = Dir(Globals.OUTDIR & _
                   "*" & LCase(Globals.PPT_EXPORT) & "*")

    'Create text file
    Set TextFile = fso.CreateTextFile(Globals.OUTDIR & _
                                      Globals.TEXT_FILE, True)

    'Write this at the beginning of the file
    TextFile.writeline ("For each file being uploaded, copy and paste" & _
                        "the following texts onto the section " & _
                        "for embedded code:")
    TextFile.writeline ("")
    TextFile.writeline ("")

    'For each searched item in the folder write the embed code
    Do While Len(FileName) > 0

        TextFile.writeline ("")
        TextFile.writeline ("")
        TextFile.writeline (FileName)
        TextFile.writeline ("")
        TextFile.writeline ("")
        TextFile.writeline ("<!DOCTYPE HTML PUBLIC " & _
                            Q & "-//W3C//DTD HTML 3.2//EN" & _
                            Q & ">")
        TextFile.writeline ("<html>")
        TextFile.writeline ("<body> <body style=" & _
                            Q & "color: rgb(124, 112, 218);" & _
                            " background-color: rgb(255, 255, 255)" & _
                            Q & ">")
        TextFile.writeline ("<body leftmargin=" & Q & "0" & Q & " " & _
                            "topmargin=" & Q & "0" & Q & " " & _
                            "marginwidth=" & Q & "0" & Q & " " & _
                            "marginheight=" & Q & "0" & Q & ">")
        TextFile.writeline ("<center>")
        TextFile.writeline ("<img src=" & Q & "G:\Content\" & _
                            FileName & Q & ">")
        TextFile.writeline ("</center>")
        TextFile.writeline ("</body>")
        TextFile.writeline ("</html>")
        TextFile.writeline ("")
        TextFile.writeline ("")
        FileName = Dir
   
    Loop

    'Close file
    TextFile.Close

End Sub

