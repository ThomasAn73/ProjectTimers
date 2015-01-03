VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call microsoftWindowResizeCode 'this is not critical. If removed be sure to delete the mdlFormResizeLimit module as well
End Sub

Private Sub microsoftWindowResizeCode()
    'Save handle to the form.
    gHW = Me.hwnd

    'Begin subclassing.
    Hook
End Sub

Private Sub microsoftResizecodeUnload()
    'Stop subclassing. (this the microsoft code)
    Unhook
End Sub
