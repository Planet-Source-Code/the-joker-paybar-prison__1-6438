VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "PayBar Prison"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   930
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   1640
      ButtonWidth     =   1640
      ButtonHeight    =   1482
      Appearance      =   1
      ImageList       =   "img"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "capture"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   "exit"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList img 
      Left            =   615
      Top             =   2325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   55
      ImageHeight     =   50
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":12D4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
wincount = 1 'set the count of windows to one this will be used later to decide the
'position of a newly added window

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'if the user pressed the 'X' on the form then make sure to release all of the
'PayBars before we shut down.  Most will shut down on their own.  however some have errors doing
'this so it is better to release them using the windows shut down API
For i = 0 To 9999 'set the range of the search
   win = GetParent(i)
   If win = MDIForm1.hwnd Then
     'this is one of the captured windows
     Call WindowHandle(win, 0) 'call the shutdown window sub
   End If
Next

'end the program
Unload Me
End

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'see whats up.
'see what button was pressed
If Button.Key = "capture" Then 'capture key was pressed.  That means
' you should load the capture windows form

Form1.Show 1 'Make the form a modal form.  That way you don't have to
'worry about them capturing a PayBar that is already in Prison


Else
For i = 0 To 9999 'set the range of the search
   win = GetParent(i)
   If win = MDIForm1.hwnd Then
     'this is one of the captured windows
     Call WindowHandle(win, 0) 'call the shutdown window sub
   End If
Next

'end the program
Unload Me
End

End If
End Sub
