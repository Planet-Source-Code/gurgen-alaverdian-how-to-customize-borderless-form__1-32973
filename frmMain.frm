VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "Sysinfo.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Blue 'n Rusty"
   ClientHeight    =   4305
   ClientLeft      =   -90
   ClientTop       =   -660
   ClientWidth     =   6000
   DrawMode        =   15  'Merge Pen Not
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":08CA
   ScaleHeight     =   4300
   ScaleMode       =   0  'User
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   1080
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList IL 
      Left            =   390
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1922
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D75
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox P3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   195
      Picture         =   "frmMain.frx":21C0
      ScaleHeight     =   225
      ScaleWidth      =   5610
      TabIndex        =   0
      Top             =   4110
      Width           =   5610
   End
   Begin VB.PictureBox P2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmMain.frx":2C3A
      ScaleHeight     =   555
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   6000
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rusty"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00466C8E&
         Height          =   225
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   180
         Width           =   705
      End
      Begin VB.Image I3 
         Height          =   300
         Index           =   2
         Left            =   5610
         Picture         =   "frmMain.frx":4037
         Stretch         =   -1  'True
         Top             =   120
         Width           =   300
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Blue 'n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A46B2D&
         Height          =   225
         Index           =   0
         Left            =   390
         TabIndex        =   2
         Top             =   180
         Width           =   705
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   90
         Picture         =   "frmMain.frx":44AF
         Top             =   150
         Width           =   240
      End
      Begin VB.Image I3 
         Height          =   300
         Index           =   1
         Left            =   5340
         Picture         =   "frmMain.frx":48AF
         Stretch         =   -1  'True
         Top             =   120
         Width           =   300
      End
      Begin VB.Image I3 
         Height          =   300
         Index           =   0
         Left            =   5070
         Picture         =   "frmMain.frx":4CEA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   300
      End
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4305
      Index           =   1
      Left            =   5805
      Picture         =   "frmMain.frx":50EE
      ScaleHeight     =   4300
      ScaleMode       =   0  'User
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox P1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4300
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":58EB
      ScaleHeight     =   4300
      ScaleMode       =   0  'User
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const hotSpot = 200
Private setMaxWin   As Boolean
Private winTop      As Long
Private winLeft     As Long
Private winWidth    As Long
Private winHeight   As Long
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault     'Restore default pointer as the mouse is over the form.
End Sub
Private Sub Form_Paint()
TileBackground      'Call a sub to tile background picture.
End Sub
Private Sub Form_Resize()
With Me

    If .WindowState = 1 Then Exit Sub     'Prevents resizing in form min. state.
    If .Width < 4000 Then .Width = 4000   'Restricts form width.
    If .Height < 4000 Then .Height = 4000 'Restricts form height.

    TileBackground                        'Tiles background picture on the form.
    
    'Routine to resize frame bars and form state icons.
    
    P2.Width = .Width
    P1(0).Height = .Height
    P1(1).Height = .Height
    P1(1).Left = .Width - 195 '195 is width of a picture itself.
    P3.Width = .Width - 390   'bottom frame width minus width of right and left bars.
    P3.Top = .Height - 195
    I3(0).Left = .Width - 930   'sets state icons position.
    I3(1).Left = .Width - 660
    I3(2).Left = .Width - 390
End With
End Sub
Sub TileBackground()
Dim bgdImage    As Picture
Dim X           As Integer
Dim Y           As Integer

'Used to tile bakground picture

Set bgdImage = Me.Picture
Y = 0
While Y < Me.Height
    X = 0
    While X < Me.Width
        PaintPicture bgdImage, X, Y
        X = X + bgdImage.Width \ 2
    Wend
        Y = Y + bgdImage.Height \ 2
Wend

End Sub
Public Sub I3_Click(Index As Integer)
'Sets Form State
With Me
    Select Case Index

        Case 0
            'minimize window
            .WindowState = 1
        Case 2
            'Close program
            Unload Me
            Unload mnuForm
        Case 1
    
        'Maximize and Restore a form.
        'setMaxWin is declared globally and toggle between normal and maximized state of the form.
        'I could've simply use Me.WindowState = 0 or 2, however when maximazing a borderless form
        'task bar is getting covered with the form.

            Select Case setMaxWin
                Case False
                    setMaxWin = True       'Form is in a normal state. Going to switch to max.
                    winTop = .Top          'Keep current form position in memory for the restore.
                    winLeft = .Left
                    winWidth = .Width
                    winHeight = .Height
                    I3(1).Picture = IL.ListImages(1).Picture 'Switch form state icon to restore.
            
                    'Resize form to fit available screen space. The easiest way to do it, is with SysInfo control
                    .Move SysInfo1.WorkAreaLeft, SysInfo1.WorkAreaTop, SysInfo1.WorkAreaWidth, SysInfo1.WorkAreaHeight
                Case True
                    setMaxWin = False       'Form is in a max. state. Going to switch to normal.
                    .Move winLeft, winTop, winWidth, winHeight      'Retrieve position from memory.
                    I3(1).Picture = IL.ListImages(2).Picture           'Switch form state icon to max.
            End Select
    End Select
End With
End Sub
Private Sub I3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
I3(Index).BorderStyle = 1 'To get a push button effect
End Sub
Private Sub I3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault 'Restore default pointer as the mouse is over the buttons.
End Sub
Private Sub I3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
I3(Index).BorderStyle = 0 'to get a push button effect
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    ConfigMenu                  'Configure Menu
    PopupMenu mnuForm.mnuResize, vbPopupMenuRightButton, , , mnuForm.mnuClose  'Show menu
End If
End Sub
Private Sub P2_DblClick()
I3_Click (1) 'imitates double click on the title bar
End Sub
Private Sub P3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbSizeNS 'Sets mouse pointer
End Sub
Private Sub P1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
CaptureHit  'Since vertical frames serve 4 resize directions, call a sub to figure out correct mouse pointer.
End Sub
Private Sub P2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CaptureHit  'Identical to the above only for title bar.
End Sub
Private Sub P3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If setMaxWin Then Exit Sub
If Button = vbLeftButton Then ReleaseForm  'Call a sub to initialize form resize.
End Sub
Private Sub P1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If setMaxWin Then Exit Sub
If Button = vbLeftButton Then ReleaseForm   'Call a sub to initialize form resize.
End Sub
Private Sub P2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Moves and resize form
  
If Button = vbLeftButton Then
    If setMaxWin Then Exit Sub 'Do not move maximazed form.
    If Y > hotSpot Then        'If you below the resize hit spot then set parameter to move form
        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    Else: ReleaseForm            'If you hit the resize spot then initialize resize
    End If
Else:
    ConfigMenu                  'Configure Menu
    PopupMenu mnuForm.mnuResize, vbPopupMenuRightButton, , , mnuForm.mnuClose  'Show menu
End If
End Sub
Private Sub ReleaseForm()
Dim retParam As Long

'This Sub will initialized form resize.

retParam = CaptureHit       'Call the sub to get the current mouse position and corner/side parameter.

ReleaseCapture              'Free up the form
    
'Set parameter according to the returned value from CaptureHit
'Depending on the parameter, resize from appropriate corner/side.
    
SendMessage Me.hWnd, WM_NCLBUTTONDOWN, retParam, ByVal 0&

End Sub
Private Sub ConfigMenu()
'Hide menu options depending on windows state.

If setMaxWin Then
    mnuForm.mnuRest.Enabled = True
    mnuForm.mnuMin.Enabled = True
    mnuForm.mnuMax.Enabled = False
Else:
    mnuForm.mnuRest.Enabled = False
    mnuForm.mnuMin.Enabled = True
    mnuForm.mnuMax.Enabled = True
End If
End Sub
Private Function CaptureHit()
'This Sub is used to capture mouse cursor position on the form and set the side/corner parameter.
Dim curPos  As POINTAPI
Dim FormX       As Long
Dim FormY       As Long

If setMaxWin Then Exit Function 'Do not capture if form is maximazed.


GetCursorPos curPos  'Captures screen mouse cursor position.

'I need to get a current mouse position on the form regardless at which control the mouse is over
'(a picture box in this case). So I get the global cursor position in pix., then multiply it by TwipsPerPixel
'converting to twips and finally deducting the distance between form and the top and left of the screen.

FormX = (curPos.curX * Screen.TwipsPerPixelX) - Me.Left
FormY = (curPos.curY * Screen.TwipsPerPixelY) - Me.Top

'Depending on the hotSpot setting and mouse cursor position sets mouse pointer and returns
'one of eight possible directions.

    Select Case FormX
        Case Is < hotSpot
            If FormY < hotSpot Then
                CaptureHit = HTTOPLEFT
                Me.MousePointer = vbSizeNWSE
            ElseIf FormY >= (ScaleHeight - hotSpot) Then
                CaptureHit = HTBOTTOMLEFT
                Me.MousePointer = vbSizeNESW
            Else
                CaptureHit = HTLEFT
                Me.MousePointer = vbSizeWE
            End If
        Case Is >= (ScaleWidth - hotSpot)
            If FormY < hotSpot Then
                CaptureHit = HTTOPRIGHT
                Me.MousePointer = vbSizeNESW
            ElseIf FormY >= (ScaleHeight - hotSpot) Then
                CaptureHit = HTBOTTOMRIGHT
                Me.MousePointer = vbSizeNWSE
            Else
                CaptureHit = HTRIGHT
                Me.MousePointer = vbSizeWE
            End If
        Case Else
            If FormY < hotSpot Then
                CaptureHit = HTTOP
                Me.MousePointer = vbSizeNS
            ElseIf FormY >= (ScaleHeight - hotSpot) Then
                CaptureHit = HTBOTTOM
                Me.MousePointer = vbSizeNS
            Else
                CaptureHit = HTCAPTION
                Me.MousePointer = vbDefault
            End If
    End Select
End Function
