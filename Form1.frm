VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Capture PayBar"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   2475
   End
   Begin VB.Label lblsecs 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      Height          =   300
      Left            =   1845
      TabIndex        =   2
      Top             =   1830
      Width           =   930
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "You Have:"
      Height          =   300
      Left            =   1508
      TabIndex        =   1
      Top             =   1335
      Width           =   1605
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form1.frx":0000
      Height          =   1125
      Left            =   68
      TabIndex        =   0
      Top             =   45
      Width           =   4545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) = "j" Or Chr(KeyAscii) = "J" Then 'if the User pressed a j or J then
   'run this code
   
   Dim p As Where 'declare the Mouse Point Type
   
   Call GetCursorPos(p) 'get the Mouse Position in X and Y coordinates
   win = WindowFromPoint(p.Pointa, p.Pointb) 'find the window that the mouse is
   'over
   win = GetParent(win) 'get the Parent of whatever window you are on
   'this is very important.  If they pressed 'J' over a text box then you would move the
   'text box into our prison but leave the main Window (or Parent) on the desktop
   'not pretty. But definitely Funny LOL
      
   Call SetParent(win, MDIForm1.hwnd) 'change the Parent of the Window from the
   'desktop to the Main Window.  This is the heart of the program
   
   Call MoveWindow(win, (50 * wincount), 200, 300, 50, True) 'change the
   'window shape a location to make it more managable
   wincount = wincount + 1 'increment the number of windows for the next time the
   'user adds a window
   Unload Me
   
End If
End Sub

Private Sub Timer1_Timer()
If lblsecs.Caption = 0 Then 'if the seconds label has reached 0
   Unload Me 'then unload this form
Else 'else
   lblsecs.Caption = lblsecs.Caption - 1 'deduct a second from the time left
End If

End Sub
