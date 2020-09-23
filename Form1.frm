VERSION 5.00
Object = "{4A49E33C-EE47-11D1-AE0A-00A0C92A54B0}#1.0#0"; "QCVIDEOX.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Sauri Soft's Quick Cam Viewer v1.0"
   ClientHeight    =   6870
   ClientLeft      =   -45
   ClientTop       =   -330
   ClientWidth     =   9300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnHelp 
      BackColor       =   &H0000C000&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton btnFreeze 
      BackColor       =   &H00FFC0C0&
      Caption         =   "||"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Pause"
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton btnPaused 
      BackColor       =   &H0000C000&
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton btnZoomOut 
      BackColor       =   &H00FFC0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Zoom Out"
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton btnZoomin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Zoom In"
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton btnCaptura 
      BackColor       =   &H0000C000&
      Caption         =   "Capture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   3480
   End
   Begin VB.CommandButton btnVideo 
      BackColor       =   &H0000C000&
      Caption         =   "Camera"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton btnFormat 
      BackColor       =   &H0000C000&
      Caption         =   "Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin QCVIDEOXLib.QCVideoX QCVideoX1 
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   1296
      _StockProps     =   0
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   8880
      MousePointer    =   2  'Cross
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      ToolTipText     =   "Good byeee"
      Top             =   0
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":09CC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9375
   End
   Begin VB.Image Image2 
      Height          =   1785
      Left            =   7680
      Picture         =   "Form1.frx":4A88
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1440
   End
   Begin VB.Image imgScreen 
      Height          =   1920
      Left            =   1320
      MouseIcon       =   "Form1.frx":8B30
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":8C82
      Top             =   720
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Driver As Integer
Dim Freeze As Boolean


Private Sub btnCaptura_Click()
    SaveAs
End Sub

Private Sub btnFormat_Click()
QCVideoX1.DlgVideoFormat
End Sub

Private Sub btnSalir_Click()
End
End Sub

Private Sub btnFreeze_Click()
    Freeze = Not Freeze
End Sub

Private Sub btnHelp_Click()
    MsgBox "C'mon, just right click on the image box to open a popup menu", vbExclamation, "Help? haha"
End Sub

Private Sub btnPaused_Click()
    Freeze = Not Freeze
End Sub

Private Sub btnVideo_Click()
QCVideoX1.DlgVideoSource
End Sub

Private Sub btnZoomin_Click()
    ZoomIn
End Sub

Private Sub btnZoomOut_Click()
    ZoomOut
End Sub

Private Sub Form_Load()
Dim aux As Boolean
  
ChDir App.Path
  
Image1.Top = 0
Image1.Left = 0
Image1.Width = Form1.Width
Image3.Left = Form1.Width - Image3.Width

MoveZoomButtons

QCVideoX1.Initialization Driver
QCVideoX1.SetColorDepth 24
QCVideoX1.SetFrameSize 320, 240
QCVideoX1.EnablePictureSmart True
QCVideoX1.SetColorDepth 16
QCVideoX1.SetVideoDisplay True

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Form1.Left = Form1.Left + X
    Form1.Top = Form1.Top + Y
End If
End Sub

Private Sub Image3_Click()
 End
End Sub

Private Sub imgScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    frmMenu.PopupMenu frmMenu.menu
End If
End Sub

Private Sub imgScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    imgScreen.Left = imgScreen.Left + X
    imgScreen.Top = imgScreen.Top + Y
    MoveZoomButtons
End If

End Sub

Private Sub Timer1_Timer()
Static band As Boolean
If band = False Then
    band = True
Else
    imgScreen.Stretch = True
End If

If Freeze = False Then
    QCVideoX1.SaveSingleFrameToFile "temp.bmp"
    imgScreen.Picture = LoadPicture("temp.bmp")
    imgScreen.Refresh
End If

End Sub

Private Sub MoveZoomButtons()
    btnZoomin.Left = imgScreen.Left
    btnZoomin.Top = imgScreen.Top
    btnZoomOut.Left = btnZoomin.Left + btnZoomin.Width
    btnZoomOut.Top = imgScreen.Top
    btnFreeze.Left = btnZoomOut.Left + btnZoomOut.Width
    btnFreeze.Top = imgScreen.Top
End Sub



Private Sub ZoomIn()
    imgScreen.Width = imgScreen.Width + (100 * Screen.TwipsPerPixelX)
    imgScreen.Height = imgScreen.Height + (100 * Screen.TwipsPerPixelY)
End Sub

Private Sub ZoomOut()
    imgScreen.Width = imgScreen.Width - (100 * Screen.TwipsPerPixelX)
    imgScreen.Height = imgScreen.Height - (100 * Screen.TwipsPerPixelY)
End Sub

