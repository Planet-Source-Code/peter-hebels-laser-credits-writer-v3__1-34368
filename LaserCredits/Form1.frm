VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LaserShow credits writer V-III"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12705
   DrawWidth       =   3
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   847
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      ScaleHeight     =   5535
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   0
      Width           =   105
      Begin VB.Image Image2 
         Height          =   5490
         Left            =   0
         Picture         =   "Form1.frx":030A
         Top             =   0
         Width           =   105
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   2
      Left            =   7920
      ScaleHeight     =   720
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   0
      Width           =   4635
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   1
      Left            =   7920
      ScaleHeight     =   720
      ScaleWidth      =   4635
      TabIndex        =   2
      Top             =   720
      Width           =   4635
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   450
      TabIndex        =   1
      Top             =   600
      Width           =   450
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         Picture         =   "Form1.frx":12BC
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   0
      Left            =   7920
      ScaleHeight     =   720
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   1440
      Width           =   4635
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   8
      Y1              =   0
      Y2              =   368
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'LaserShow Credits Writer V-III project by Peter Hebels, Website "www.phsoft.cjb.net"      *
'Iam not responsible for any damages may caused by this project                           *
'******************************************************************************************

'******************************************************************************************
'This simple app is used to draw strings, with can even be changed at runtime
'With an laserlike effect to a form.
'After drawing the text will go down and waits at the end of the form for the other
'text to arrive and then dissapears smoothly.
'
'To change the texts shown you only have to change the NumDrawSequences variable
'and the ItemText variables
'The NumDrawSequences variable has to be set to the number of ItemTexts +1
'So 9 ItemText will be 10 NumDrawSequences
'Simple huh!!
'
'Note there is a small bug:
'If the third text is at the end and the fourth arrives the third will dissapear
'into nowhere :(
'Can you fix it, if you can pleaz let me know.
'my email "hebels13@zonnet.nl"
'Thanks!!

'I hope this app is usefull to you
'Also let me know if you have updated it, because I use this too in some programs

Dim I As Integer 'Used for controlling the picture's movement
Dim I2 As Integer 'Used for controlling the movingdown picture

Private Sub Form_Load()
 
 'Form control code:
 Form1.Show 'Show the form
 Me.ScaleMode = 3 'Set the scalemode to pixels because we going to use bitblt
 Me.Width = 7140 'Don't show the pictureboxes
  
 'set the values:
 NumPics = 3 'How many pict's have we used in one array
 
 DrawSpeed = 10 'How fast are the letters drawn, Higher value = Slower speed
 DownSpeed = 2  'How fast the text goes down, Higher value = Slower speed
 PicArray = -1  'We begin at 0
 NumDrawSequences = 10 'Number of DrawSequences, you have to change this if you add more strings
                       'read the info at the begin of form1 for more info about this.
 
 Picture1(0).BackColor = Form1.BackColor 'Use the same backcolor as the form
 Picture1(1).BackColor = Form1.BackColor 'Use the same backcolor as the form
 Picture1(2).BackColor = Form1.BackColor 'Use the same backcolor as the form
 
 'This are the begin strings, you can change these to your own texts.
 'See module1.SetTextToPic for more info
 ItemText1 = "This test is string 1"
 ItemText2 = "This test is string 2"
 ItemText3 = "This test is string 3"
  
 Form1.AutoRedraw = True 'Set AutoRedraw to true, otherwise nothing will appear.
  
 StartTextDraw 'Start the text-Draw function
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
TimerEnabled = False 'Make sure the timer is stopped
Unload Form1 'Unload the form
End 'End the program
End Sub

Public Sub Timer1_Timer()
'This timersub does the downgoing text job
On Error Resume Next 'Resume on errors

TimerEnabled = True 'set this back to true otherwise the loop won't start

'Create a loop
Do

If I > Form1.ScaleHeight - Picture1(PicArray).Height - 100 Then
  I2 = I 'Keep the I value for the waiting picture
  
  I = 0 'Reset I to 0 because we have written a new text
  StartTextDraw 'Restart the drawing function
  Exit Do 'Stop the loop so the moving picture stops
 Exit Sub
End If

I = I + 2 'For the image going down
I2 = I2 + 1 'For the waiting image

Picture2.Top = I 'Move the laser down

'BitBlt the image using I as the Y value, so its going down
BitBlt Form1.hdc, X + 132, Y - 2 + I, Form1.Picture1(PicArray).Width, Form1.Picture1(PicArray).Height, Form1.Picture1(PicArray).hdc, 0, 0, SRCCOPY
BitBlt Form1.hdc, X + 132, Y - 2 + I2, Form1.Picture1(PicArray - 1).Width, Form1.Picture1(PicArray - 1).Height, Form1.Picture1(PicArray - 1).hdc, 0, 0, SRCCOPY
'Refresh the form every single frame.
Form1.Refresh

'Sleep is used fo controlling speed
Sleep DownSpeed

DoEvents

'This will stop the loop if needed
If TimerEnabled = False Then Exit Do

Loop

End Sub
