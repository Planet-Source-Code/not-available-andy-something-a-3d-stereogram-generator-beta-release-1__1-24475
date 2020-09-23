VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8235
   ClientLeft      =   1650
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "MDIForm1"
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub
Private Sub mnuExit_Click()
End
End Sub
Private Sub mnuCopy_Click()
copier MDIForm1.ActiveForm.Picture1.Image
End Sub
Private Sub MDIForm_Load()
Form2.Command2_Click
mnuSave.Enabled = False
mnuCopy.Enabled = False
mnuGenerate.Enabled = False
End Sub
Private Sub mnuOptions_Click()
Form2.Show vbModal, Me
End Sub
Private Sub mnuOpen_Click()
opener Me
End Sub
Private Sub mnuSave_Click()
saver MDIForm1.ActiveForm.Picture1.Image
End Sub
Private Sub mnuPaste_Click()
paster
End Sub
Private Sub mnuGenerate_Click()
MDIForm1.ActiveForm.copyable = True
MDIForm1.mnuCopy = True
MDIForm1.mnuSave = True
generate MDIForm1.ActiveForm.Picture1, Form2.Picture1
End Sub
