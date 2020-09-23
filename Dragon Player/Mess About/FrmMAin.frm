VERSION 5.00
Begin VB.Form FrmPlayer 
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
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   0
      MousePointer    =   1  'Arrow
      Picture         =   "FrmMAin.frx":0000
      ScaleHeight     =   282
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   663
      TabIndex        =   0
      ToolTipText     =   "Click Ying And Yang To Expose Navigation Buttons"
      Top             =   0
      Width           =   9945
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000000&
         Caption         =   "Command1"
         Height          =   135
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1620
         Width           =   195
      End
      Begin VB.PictureBox MoveBar 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   135
         Left            =   4200
         MousePointer    =   5  'Size
         ScaleHeight     =   105
         ScaleWidth      =   5265
         TabIndex        =   4
         Top             =   1620
         Width           =   5295
      End
      Begin Project1.LCDDisplay LCDDisplay1 
         Height          =   735
         Left            =   4560
         TabIndex        =   1
         Top             =   2040
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   1296
         ElementColour   =   1018284
         BeginProperty DisplayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "BankGothic Lt BT"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         YOffSet         =   1
         Text            =   "#0 Welcome To GFX's Dragon Player #0"
         ImagePlaceHolder=   "#"
         ImageCollection =   "FrmMAin.frx":89294
         ImageCellWidth  =   47
         ImageCellHeight =   11
         ElementHeight   =   3
         StartVisible    =   0   'False
      End
   End
   Begin VB.PictureBox MaskImage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   0
      Picture         =   "FrmMAin.frx":89916
      ScaleHeight     =   282
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   663
      TabIndex        =   3
      Top             =   0
      Width           =   9945
   End
End
Attribute VB_Name = "FrmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StdTmr As Double
Dim FadTmr As Double

Private ButtonDC As Long
Private ButtonBMP As Long

Private BufferDC As Long
Private BufferBMP As Long

Private NavGuardTrans As Integer
Private NavToGuardSteps As Integer
Private DrawingNav As Boolean

Private Sub Command1_Click()
    
    LCDDisplay1.ReleaseMemory
    End
    
End Sub

Private Sub Form_Load()

    NavControls.Height = 61
    NavControls.Width = 61
    NavControls.X = 169
    NavControls.Y = 111
    NavControls.Visible = False
    NavToGuardSteps = 10
    
    NavGuardTrans = 0
    
    BufferDC = CreateCompatibleDC(GetDC(0))
    BufferBMP = CreateCompatibleBitmap(GetDC(0), NavControls.Width, NavControls.Height)
    SelectObject BufferDC, BufferBMP
    
    ButtonDC = LoadGraphicDC(App.Path & "\Images\Dragon Buttons.Bmp")
    SelectObject ButtonDC, ButtonBMP
    
    MaskImage.Visible = False
    SkinForm Me, MaskImage
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    DeleteDC ButtonDC
    DeleteObject ButtonBMP
    DeleteDC BufferDC
    DeleteObject BufferBMP
    
End Sub

Private Sub MoveBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 1 Then
        ReleaseCapture
        SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    NavControls.Visible = Not NavControls.Visible
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.Caption = X & " " & Y
    
End Sub

Private Sub Timer1_Timer()
    
    'Standard Display Refresh Routine
    If Timer - StdTmr > 0.05 Or Timer - StdTmr < 0 Then
        LCDDisplay1.UpdateDisplay True

        StdTmr = Timer
        
    End If
    
    'Nav Panel Fade Routine
    If Timer - FadTmr > 0.02 Or Timer - FadTmr < 0 Then
        If NavControls.Visible = False Then
            'Hide Controls
            If NavGuardTrans > 0 Then
                NavGuardTrans = NavGuardTrans - 1
                UpdateNavGuardTrans
            End If
        Else
            'Expose Controls
            If NavGuardTrans < NavToGuardSteps Then
                NavGuardTrans = NavGuardTrans + 1
                UpdateNavGuardTrans
            End If
        End If
        
        FadTmr = Timer
        
    End If
    
End Sub

Private Sub UpdateNavGuardTrans()
    
    'ButtonDC Is The Memory Image HDC Created At Load
    
    Dim X As Integer, Y As Integer
    Dim Col1 As Long, Col2 As Long
    Dim BlendedColour As Long
    
    'Iterate Through Every Pixel
    For Y = 0 To NavControls.Width - 1
        For X = 0 To NavControls.Width - 1
            
            'Colour Point Of Nav Guard
            Col1 = GetPixel(ButtonDC, X, Y)
            
            'Colour Point Of Nav Controls
            Col2 = GetPixel(ButtonDC, X + NavControls.Width, Y)
               
            If Col1 <> RGB(255, 0, 255) Then '<> Mask Colour
                    'Get The Point Colour Between The Two Images Dependent
                    'On Fade Phase
                BlendedColour = Blender(Col1, Col2, NavToGuardSteps, NavGuardTrans)
                    'Draw The Point To The Buffer (Memory Image)
                SetPixelV BufferDC, X, Y, BlendedColour
            Else
                    'If Mask Colout Just Plot It
                SetPixelV BufferDC, X, Y, Col1
            End If
        Next X
    Next Y
    
    'Blit The Buffer Onto The Visible DC Applying Transparency
    'To The Masked Pixels
    TransparentBlt Picture1.hdc, NavControls.X, NavControls.Y _
        , NavControls.Width, NavControls.Height, BufferDC, 0, 0 _
        , NavControls.Width, NavControls.Height, RGB(255, 0, 255)
    
    Picture1.Refresh
    
End Sub
