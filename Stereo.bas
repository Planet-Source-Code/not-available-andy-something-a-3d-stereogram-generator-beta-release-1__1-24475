Attribute VB_Name = "Stereo"
' Stereogram generator
' Created By Andy Nova*
' andy@highsupport.com
' http://www.highsupport.com

Global stopgen As Boolean                                       'Define a global stopper variable
Function stopper()
stopgen = True                                                  'Set Value of stopper to true
End Function
Function generate(activepix As PictureBox, patternpix As PictureBox)
On Error GoTo Error                                             'If an error goto error: handler
    stopgen = False                                             'Reset Value of stopper to false
    Dim n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer, n5 As Integer 'define variables
    n5 = 0                                                      'set bring current loc to 0 for later compare
    For n4 = 0 To activepix.ScaleHeight                         'loop the length of initial pattern
        n5 = n5 + 1                                             'bring current loc to var n5 for later compare
        For n3 = 0 To patternpix.ScaleWidth                     'Loop the width of initial pattern
            If stopgen = True Then GoTo Ender                   'If Stop Called then goto Stop handel
            If n5 >= patternpix.ScaleHeight - 1 Then n5 = 0     'If the Height is equavalent to the current end then repeat pattern
            activepix.PSet (n3, n4), patternpix.Point(n3, n5)   'Set The Initial Pattern Pixel by Pixel
        Next n3                                                 'Continue repeat the width loop
    Next n4                                                     'continue repeat the legnth loop
    For n1 = patternpix.ScaleWidth To activepix.Width           'loop the length of the activepix starting after initial pattern set
        Debug.Print Int((n1 / activepix.Width) * 100)
        DoEvents                                                'Allow the user to see the change while they wait. (Slows down dramatically)
        activepix.Refresh                                       'Refreshes the window for user vision and pixel lookback comparison
        For n2 = 0 To activepix.Height                          'loop the length of activepix height
            If stopgen = True Then GoTo Ender                   'If Stop Called then goto Stop handel
                                                                'Heres where the actual magic works
                                                                'Basically it looks at the current positions color
                                                                'compares it with the pixel that is eyes width apart (approx 2.5 in)(on a 72 dpi monitor that is about 30 pixels)
                                                                'and get distance that the color is at and subtranting value from 16.7 billion colors values
                                                                'subtranting value from 16.7 billion colors values and placing it at the new pixel location
            activepix.PSet (n1, n2), activepix.Point(n1 - 30 - activepix.Point(n1, n2) / 1677721, n2)  'wow eye magic
        Next n2                                                 'continue repeat the 2nd height loop
    Next n1                                                     'continue repeat the 2nd width loop
    Exit Function                                               'leave the function when done so err handler has its place
Error:                                                          'if an error go here
MsgBox "An Error Has Occured in Loading the Picture" & vbCrLf & "Error #" & Err.Number & vbCrLf & Err.Description 'Display a error msg with num and description
stopper                                                         'stop the generation if an error
Exit Function                                                   'exit the function for ender
Ender:                                                          'ender location (not really used anymore, just had a return origionally, but decided not necessary ne more
End Function
Function copier(picturecopy As PictureBox)
Clipboard.Clear                                                 'Clear the clipboard
Clipboard.SetData picturecopy.Image                             'Put the picture into the clipboard
End Function
Function saver(activepix As PictureBox, cdholder As Form)
cdholder.CommonDialog1.FileName = ""                                     'Set filename to null for Save Dialog
cdholder.CommonDialog1.Filter = "Bitmap|*.bmp"                           'Set filter to bitmap for Save Dialog
cdholder.CommonDialog1.ShowSave                                          'Show Save dialog
If cdholder.CommonDialog1.FileName = "" Then Exit Function               'If Save dialog is null just quit function
SavePicture activepix.Image, cdholder.CommonDialog1.FileName             'Save the picture using Save Dialog Filename
End Function
Function paster()
If Clipboard.GetData = 0 Then Exit Function                          'If there is nothing in empty then just quit
load Clipboard.GetData, "Pasted Image"                          'call loader and load image in clipboard
End Function
Function opener(cdholder)
On Error GoTo Error                                            'If an error goto error: handler
cdholder.CommonDialog1.FileName = ""                           'Set filename to null for open Dialog
cdholder.CommonDialog1.Filter = "Pictures|*.bmp;*.dib;*.jpg;*.gif|Everything|*.*" 'Set filter to pictures and all for Open Dialog
cdholder.CommonDialog1.ShowOpen                                'show open dialog
If cdholder.CommonDialog1.FileName = "" Then Exit Function     'if open dialog is null then just quit
load LoadPicture(cdholder.CommonDialog1.FileName), cdholder.CommonDialog1.FileTitle 'Load the picture with the name stated
Exit Function                                                  'leave the function when done so err handler has its place
Error:                                                         'If an error go here
MsgBox "An Error Has Occured in Loading the Picture" & vbCrLf & "Error #" & Err.Number & vbCrLf & Err.Description 'Display a error msg with num and description
End Function
Function load(ByVal pictureimport As IPictureDisp, captionimport As String)
On Error GoTo Error                                            'If an error goto error: handler
    Dim Form As New Form1                                      'Make a new form using preset template for
    Form.Picture1.Picture = pictureimport                      'Put loaded picture into form template
    Form.Caption = captionimport                               'put picture name into form template caption
    Form.Width = Form.Picture1.Picture.Width / 1.7 + 80        'size template width to the picture
    Form.Height = Form.Picture1.Picture.Height / 1.7 + 400     'size template height to the picture
    Exit Function                                              'leave the function when done so err handler has its place
Error:                                                         'If an error go here
MsgBox "An Error Has Occured in Loading the Picture" & vbCrLf & "Error #" & Err.Number & vbCrLf & Err.Description 'Display a error msg with num and description
End Function


