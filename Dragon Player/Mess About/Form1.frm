VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   660
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4230
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   0
      Width           =   9945
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   2820
         TabIndex        =   2
         Top             =   2040
         Width           =   495
      End
      Begin Project1.LCDDisplay LCDDisplay1 
         Height          =   600
         Left            =   3960
         TabIndex        =   1
         Top             =   2160
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   1058
         ElementColour   =   1018284
         BeginProperty DisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisplayPosition =   91
         Text            =   "Welcome To GFX's Dragon Player..."
         DisplayErrors   =   -1  'True
         ElementHeight   =   3
         DisplayErrors   =   -1  'True
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim A As Double

Private Sub Command1_Click()
    
    LCDDisplay1.ReleaseMemory
    End
    
End Sub

Private Sub Form_Load()

    SkinForm Me, Picture1
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&

End Sub

Private Sub Timer1_Timer()
    
    If Timer - A > 0.05 Or Timer - A < 0 Then
        LCDDisplay1.UpdateDisplay True
        A = Timer
    End If
    
End Sub
