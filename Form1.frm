VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Handler"
   ClientHeight    =   2190
   ClientLeft      =   2895
   ClientTop       =   2280
   ClientWidth     =   4110
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   274
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   117
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Stereogram generator
' Created By Andy Nova*
' andy@highsupport.com
' http://www.highsupport.com
Public copyable As Boolean
Private Sub Form_Activate()
MDIForm1.mnuGenerate = True
If copyable = True Then MDIForm1.mnuCopy = True: MDIForm1.mnuSave = True
End Sub

Private Sub Form_Resize()
Picture1.Left = 0
Picture1.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIForm1.mnuGenerate = False
MDIForm1.mnuCopy = False
MDIForm1.mnuSave = False
stopper
End Sub
