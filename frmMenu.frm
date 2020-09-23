VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu menu 
      Caption         =   " "
      Begin VB.Menu mnuSave 
         Caption         =   "Save as..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub menuExit_Click()
    End
End Sub

Private Sub mnuAbout_Click()
    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuSave_Click()
    SaveAs
End Sub
