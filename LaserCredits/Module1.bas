Attribute VB_Name = "Module1"
'******************************************************************************************
'LaserShow Credits V3 Writer project by Peter Hebels, Website "www.phsoft.cjb.net"           *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

'Some API calls
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const LF_FACESIZE = 32

'Types used by the fontcreator
Public Type LOGFONT
     lfHeight As Long
     lfWidth As Long
     lfEscapement As Long
     lfOrientation As Long
     lfWeight As Long
     lfItalic As Byte
     lfUnderline As Byte
     lfStrikeOut As Byte
     lfCharSet As Byte
     lfOutPrecision As Byte
     lfClipPrecision As Byte
     lfQuality As Byte
     lfPitchAndFamily As Byte
     lfFaceName As String * LF_FACESIZE
End Type


'Declares used for drawing the text
Public X, Y, I As Integer
Public PicArray As Integer
Public NumPics As Integer
Public DrawSpeed As Integer
Public DownSpeed As Integer

'Constants for the BitBlt API call
Public Const MERGEPAINT = &HBB0226
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020

'Store values
Public PicTop As Long
Public PicHalf As Long
Public TimerEnabled As Boolean
Public ItemText1, ItemText2, ItemText3 As String
Public DrawSequence As Integer
Public NumDrawSequences As Integer

'This is the main text-draw sub
Public Function StartTextDraw()
'The color variable stored from GetPixel
Dim PixCol As Long
'For the color values
Dim r, g, b As Integer
'Resume after every error
On Error Resume Next
SetTextToPic
'are we at the end of our show?

If DrawSequence >= NumDrawSequences Then
 MsgBox "This was the show, thanx for trying this app!", vbInformation, "See ya next time!"
  'Stop the loop, otherwise it will keep trying to this Function
  TimerEnabled = False
  'End the app
  Unload Form1
  End
 Exit Function
End If

'I use ControlArrays in this app for easy coding!
'Every call to this Function will go to the next picturebox in the array
If PicArray = 2 Then PicArray = -1
PicArray = PicArray + 1

'Start some loops, so the text-draw routine will begin
'Calculate the X value
For X = Form1.Picture1(PicArray).Width - 1 To 0 Step -1
'Calculate the Y value
For Y = 0 To Form1.Picture1(PicArray).Height - 1
    
    'Move the Laser together with X
    Form1.Picture2.Top = X
    'Get the pixel color from the picture
    PixCol = GetPixel(Form1.Picture1(PicArray).hdc, X, Y)
    
    'Make the image 16Bits color
    r = PixCol Mod 256
    b = Int(PixCol / 65536)
    g = (PixCol - (b * 65536) - r) / 256
     
    'Draw the lines (Laser Beams :))
    PicTop = Form1.Picture2.Top
    PicHalf = Form1.Picture2.Height / 2
    Form1.Line (35, PicTop + PicHalf)-(X + 130, Y + 50), RGB(r, g, b)

Next Y
'Let windows do its events
DoEvents
'This changes the speed, with can be set in FormLoad
'This also reduces the CPU usage a lot ~90%
Sleep DrawSpeed
Next X

'Clear the picture
Form1.Picture = Form1.Picture
'Enable the timer, so the picture goes down and disapears
Form1.Timer1_Timer 'Start loop


End Function

'Sub for printing fonts to pictureboxes, found on the msdn cd
Public Function SetTextToPic()
Dim font As LOGFONT
Dim prevFont As Long, hFont As Long, ret As Long

 Const FONTSIZE = 15 'Size of the text to be drawn
     Form1.Picture1(0).Cls
     Form1.Picture1(1).Cls
     Form1.Picture1(2).Cls
     
     font.lfEscapement = 0
     font.lfFaceName = "Arial" & Chr$(0)

     font.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
     hFont = CreateFontIndirect(font)
     
     prevFont = SelectObject(Form1.Picture1(0).hdc, hFont)
     prevFont = SelectObject(Form1.Picture1(1).hdc, hFont)
     prevFont = SelectObject(Form1.Picture1(2).hdc, hFont)
     
     Form1.Picture1(0).CurrentX = Form1.Picture1(0).Left + Form1.Picture1(0).Width / 9
     Form1.Picture1(0).CurrentY = Form1.Picture1(0).ScaleHeight / 9
     
     Form1.Picture1(1).CurrentX = Form1.Picture1(1).Left + Form1.Picture1(1).Width / 9
     Form1.Picture1(1).CurrentY = Form1.Picture1(1).ScaleHeight / 9
     
     Form1.Picture1(2).CurrentX = Form1.Picture1(2).Left + Form1.Picture1(2).Width / 9
     Form1.Picture1(2).CurrentY = Form1.Picture1(2).ScaleHeight / 9
     
     'Print the font to the picturebox
     DrawSequence = DrawSequence + 1
     
     'IMPORTANT.......................
     'This may looks difficult but is actualy easy!
     'The DrawSequence variable will increase with one every time this function is called
     'This function is called by the drawing function with contains 3 picture arrays.
     'So the DrawSequence variable will increase with 3 if all the texts are shown
     'It begins at 1, so 1 + The 3 already shown = 4
     
     '\|/ It begins with 4 \|/
     If DrawSequence = 4 Then
     ItemText1 = "This is test string 4"
     ItemText2 = "This is test string 5"
     ItemText3 = "This is test string 6"
     End If
     
     'Now the above 3 are shown increase with another 3, with will be 7
     '\|/ It begins with 7 \|/
     If DrawSequence = 7 Then
     ItemText1 = "This is test string 7"
     ItemText2 = "This is test string 8"
     ItemText3 = "This is test string 9"
     End If
     
     'The next DrawSequence will be 7 + 3 = 10
          
     'Here the texts will be printed to the pictureboxes
     Form1.Picture1(0).Print ItemText1
     Form1.Picture1(1).Print ItemText2
     Form1.Picture1(2).Print ItemText3

End Function

