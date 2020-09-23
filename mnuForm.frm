VERSION 5.00
Begin VB.Form mnuForm 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuResize 
      Caption         =   "Resize"
      Begin VB.Menu mnuMin 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu mnuMax 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu mnuRest 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuBr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "mnuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is nessesary for the menu only.

Private Sub mnuClose_Click()
Unload frmMain
Unload mnuForm

End Sub

Private Sub mnuMax_Click()
frmMain.I3_Click (1)
End Sub

Private Sub mnuMin_Click()
frmMain.I3_Click (0)
End Sub

Private Sub mnuRest_Click()
frmMain.I3_Click (1)
End Sub
