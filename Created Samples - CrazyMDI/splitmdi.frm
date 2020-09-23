VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDI Form with resizers everywhere"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8400
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture7 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1095
      Left            =   0
      MouseIcon       =   "splitmdi.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   1095
      ScaleWidth      =   8400
      TabIndex        =   6
      Top             =   5070
      Width           =   8400
      Begin VB.TextBox Text2 
         Height          =   855
         Left            =   120
         MousePointer    =   3  'I-Beam
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3735
      Left            =   6345
      MouseIcon       =   "splitmdi.frx":0152
      MousePointer    =   99  'Custom
      ScaleHeight     =   3735
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   1335
      Width           =   2055
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   120
         MousePointer    =   1  'Arrow
         ScaleHeight     =   3375
         ScaleWidth      =   1935
         TabIndex        =   4
         Top             =   0
         Width           =   1935
         Begin VB.Image Image1 
            Height          =   7200
            Left            =   0
            Picture         =   "splitmdi.frx":02A4
            Top             =   0
            Width           =   9600
         End
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   3735
      Left            =   0
      MouseIcon       =   "splitmdi.frx":4B3EA
      MousePointer    =   99  'Custom
      ScaleHeight     =   3735
      ScaleWidth      =   1800
      TabIndex        =   1
      Top             =   1335
      Width           =   1800
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   3375
         Left            =   0
         MousePointer    =   1  'Arrow
         ScaleHeight     =   3375
         ScaleWidth      =   1695
         TabIndex        =   3
         Top             =   0
         Width           =   1695
         Begin VB.DirListBox Dir1 
            Height          =   3015
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1335
      Left            =   0
      MouseIcon       =   "splitmdi.frx":4B53C
      MousePointer    =   99  'Custom
      ScaleHeight     =   1335
      ScaleWidth      =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   8400
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   0
         MousePointer    =   1  'Arrow
         ScaleHeight     =   1215
         ScaleWidth      =   6135
         TabIndex        =   5
         Top             =   0
         Width           =   6135
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   735
            Left            =   720
            TabIndex        =   9
            Top             =   120
            Width           =   2895
         End
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
   End
   Begin VB.Menu mnuedit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' SplitMDI
' 1998-2000 Edward Blake
' Email: blakee@cyanwerks.com
'        blakee@rovoscape.com
' This project Shows how to create a
' multiple-splitted MDI parent window.
' This project in particular manages 4 panes.
' There are only 2 API functions used too.
'
' http://www.cyanwerks.com/
'

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private BMove As Boolean
Private SplitCoord As Single
Private OldX As Single
Private OldY As Single

Private Sub MDIForm_Load()
    Form1.Show
End Sub

Private Sub MDIForm_Resize()
    Picture6.Refresh
End Sub

Private Sub Picture1_Resize()
    Picture6.Move 0!, 0!, Picture1.ScaleWidth, Picture1.ScaleHeight - 60!
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    BMove = True
    RetVal = SetCapture(Picture2.hwnd)
    OldX = -32
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BMove Then
        Picture1.DrawMode = 6
        Picture2.DrawMode = 6
        If OldX <> -32 Then
            Picture1.Line (OldX - 15!, 0!)-(OldX + 15!, Height), , BF
            Picture2.Line (OldX - 15!, 0!)-(OldX + 15!, Height), , BF
        End If
        Picture1.Line (X - 15!, 0!)-(X + 15!, Height), , BF
        Picture2.Line (X - 15!, 0!)-(X + 15!, Height), , BF
        OldX = X
    End If
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    If BMove Then
        RetVal = ReleaseCapture()
        BMove = False
        Picture1.Cls
        Picture2.Cls
        SplitCoord = X
        If SplitCoord <= 500! Then SplitCoord = 500!
        'If SplitCoord >= (ScaleWidth - 500!) Then SplitCoord = (ScaleWidth - 500!)
        Picture2.Width = SplitCoord
    End If
End Sub

Private Sub Picture2_Resize()
    Picture4.Move 0!, 0!, Picture2.ScaleWidth - 60!, Picture2.ScaleHeight
    Picture6.Refresh
End Sub

Private Sub Picture3_Resize()
    Picture6.Refresh
    Frame1.Refresh
    Picture5.Move 60!, 0!, Picture3.ScaleWidth - 60!, Picture3.ScaleHeight
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    BMove = True
    RetVal = SetCapture(Picture1.hwnd)
    OldY = -32

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BMove Then
        Picture1.DrawMode = 6
        If OldX <> -32 Then
            Picture1.Line (0!, OldY - 30)-(Width, OldY + 30), , BF
        End If
        Picture1.Line (0!, Y - 30)-(Width, Y + 30), , BF
        OldY = Y
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    If BMove Then
        RetVal = ReleaseCapture()
        BMove = False
        Picture1.Cls
        Picture2.Cls
        SplitCoord = Y
        If SplitCoord <= 500! Then SplitCoord = 500!
        'If SplitCoord >= (ScaleWidth - 500!) Then SplitCoord = (ScaleWidth - 500!)
        Picture1.Height = SplitCoord
    End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    BMove = True
    RetVal = SetCapture(Picture3.hwnd)
    OldX = -32

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BMove Then
        Picture1.DrawMode = 6
        Picture3.DrawMode = 6
        If OldX <> -32 Then
            Picture1.Line (OldX - 15! + Picture3.Left, 0!)-(OldX + 15! + Picture3.Left, Height), , BF
            Picture3.Line (OldX - 15!, 0!)-(OldX + 15!, Height), , BF
        End If
        Picture1.Line (X - 15! + Picture3.Left, 0!)-(X + 15! + Picture3.Left, Height), , BF
        Picture3.Line (X - 15!, 0!)-(X + 15!, Height), , BF
        OldX = X
    End If
End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    If BMove Then
        RetVal = ReleaseCapture()
        BMove = False
        Picture1.Cls
        Picture3.Cls
        SplitCoord = X
        
        If SplitCoord > Picture3.ScaleWidth Then SplitCoord = Picture3.ScaleWidth - 100!
        Picture3.Width = Picture3.Width - SplitCoord
    End If
End Sub

Private Sub Picture4_Resize()
    Dir1.Move 0!, 0!, Picture4.ScaleWidth, Picture4.ScaleHeight
    Frame1.Refresh
    Picture4.Refresh
End Sub

Private Sub Picture6_Resize()
    Frame1.Move 0!, 0!, Picture6.ScaleWidth, Picture6.ScaleHeight
    Picture6.Refresh
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    BMove = True
    RetVal = SetCapture(Picture7.hwnd)
    OldY = -32

End Sub

Private Sub Picture7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If BMove Then
        Picture7.DrawMode = 6
        If OldX <> -32 Then
            Picture7.Line (0!, OldY - 30)-(Width, OldY + 30), , BF
        End If
        Picture7.Line (0!, Y - 30)-(Width, Y + 30), , BF
        OldY = Y
    End If
End Sub

Private Sub Picture7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim RetVal As Long
    If BMove Then
        RetVal = ReleaseCapture()
        BMove = False
        Picture7.Cls
        SplitCoord = Y
        If SplitCoord >= Picture7.Height - 100! Then SplitCoord = Picture7.Height - 100!
        Picture7.Height = Picture7.Height - SplitCoord
    End If
End Sub

Private Sub Picture7_Resize()
    Text2.Move 0!, 60!, Picture7.ScaleWidth, Picture7.ScaleHeight - 60!
End Sub
