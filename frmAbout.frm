VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About QuickCam Viewer v1.0"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "Stop looking at me!"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim str As String
str = "Sauri Soft's Quick Cam Viewer v1.0." + vbCrLf
str = str + "By: Jorge Sauri Creus." + vbCrLf + vbCrLf
str = str + "Sauri Soft Inc. 2000. No rights reserved. :P" + vbCrLf
Label1.Caption = str
End Sub

