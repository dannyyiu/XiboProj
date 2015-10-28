Attribute VB_Name = "Globals"
'@file Globals
'@brief Global variables and constants.
'
' All project configurations can be set here by changing
' the constants.
'
Option Private Module


'@brief Global Constants

Public Const SAVEDIR = "C:\Xibo\" 'Attachment save folder
Public Const OUTDIR = "C:\XiboPrepared\" 'Output folder
Public Const SLIDE_WIDTH = 975.18 'Slide width
Public Const SLIDE_HEIGHT = 525.18 'Slide height
Public Const TEXT_FILE = "embedded-code.txt" 'text filename
Public Const PPT_EXPORT = "PNG" 'must be valid PPT export type
Public Const OPEN_OUTDIR = True 'open directory after finished


'@brief Global Vars
Public YYYY As String 'Takedown year
Public MM As String 'Takedown month
Public DD As String 'Takedown day
