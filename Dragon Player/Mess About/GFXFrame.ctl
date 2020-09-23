VERSION 5.00
Begin VB.UserControl GFXFrame 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "GFXFrame.ctx":0000
   Begin VB.Label LblFrameCaption 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "GFXFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'SIMPLE FRAME BY S.ATKISS

Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private vFrameCaption As String
Private vFrameFont As Font
Private vFrameOpacity As Integer
Private vFontColour As OLE_COLOR

Private Sub ReDrawFrame()
    
    Cls
    
    LblFrameCaption.Top = 0
    
    LblFrameCaption.BackColor = Ambient.BackColor
    LblFrameCaption.ForeColor = FontColour
    
    LblFrameCaption.AutoSize = True
    Set LblFrameCaption.Font = FrameFont
    LblFrameCaption.Caption = FrameCaption
    
    LblFrameCaption.Width = LblFrameCaption.Width + 1
    
    LblFrameCaption.Refresh
    
    BackColor = Ambient.BackColor
    
    If FrameCaption = "" Then
        LblFrameCaption.Visible = False
    Else
        LblFrameCaption.Visible = True
    End If
    
    ForeColor = GetOpacityColourEx(vbWhite, Ambient.BackColor, FrameOpacity)
    DrawWidth = 2
    RoundRect hdc, 2, Int(LblFrameCaption.Height \ 2) + 2, ScaleWidth - 1, ScaleHeight - 1, 6, 6
    
    ForeColor = GetOpacityColourEx(vbBlack, Ambient.BackColor, FrameOpacity)
    DrawWidth = 1
    RoundRect hdc, 0, Int(LblFrameCaption.Height \ 2), ScaleWidth, ScaleHeight, 8, 8
    
    Refresh
    
End Sub

Private Function GetOpacityColourEx(Colour1 As Long, Colour2 As Long, Opacity As Integer) As Long
    
    Dim SR As Double
    Dim SG As Double
    Dim SB As Double
    
    Dim SpriteColour As Long
    
    Dim R1 As Byte, G1 As Byte, B1 As Byte
    CRGB Colour1, R1, G1, B1
    
    Dim R2 As Byte, G2 As Byte, B2 As Byte
    CRGB Colour2, R2, G2, B2
    
    SR = CDbl(R1) - R2
    SG = CDbl(G1) - G2
    SB = CDbl(B1) - B2
    
    SR = (SR + 1) / 100
    SG = (SG + 1) / 100
    SB = (SB + 1) / 100
        
    GetOpacityColourEx = RGB(R2 + (SR * Opacity), G2 + (SG * Opacity), B2 + (SB * Opacity))

End Function


Private Function CRGB(LongColour As Long, Optional Red As Byte, Optional Green As Byte, Optional Blue As Byte)

        Red = LongColour And 255
        Green = (LongColour \ 256) And 255
        Blue = (LongColour \ 65536) And 255

End Function


Private Sub UserControl_AmbientChanged(PropertyName As String)
    
    ReDrawFrame
    
End Sub

Private Sub UserControl_InitProperties()
    
    vFrameCaption = Name
    Set vFrameFont = Ambient.Font
    vFrameOpacity = 50
    vFontColour = vbBlack
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        vFrameCaption = .ReadProperty("FrameCaption", Name)
        Set vFrameFont = .ReadProperty("FrameFont", Ambient.Font)
        vFrameOpacity = .ReadProperty("FrameOpacity", 50)
        vFontColour = .ReadProperty("FontColour", vbBlack)
    End With
    
    ReDrawFrame
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        .WriteProperty "FrameCaption", vFrameCaption, Name
        .WriteProperty "FrameFont", vFrameFont, Ambient.Font
        .WriteProperty "FrameOpacity", vFrameOpacity, 50
        .WriteProperty "FontColour", vFontColour, vbBlack
    End With
    
End Sub

Private Sub UserControl_Resize()
    
    ReDrawFrame
    
End Sub

Public Property Get FrameCaption() As String
    
    FrameCaption = vFrameCaption
    
End Property

Public Property Let FrameCaption(ByVal vNewCaption As String)
    
    vFrameCaption = vNewCaption
    
    PropertyChanged "FrameCaption"
    
    ReDrawFrame
    
End Property



Public Property Get FrameFont() As Font
    
    Set FrameFont = vFrameFont
    
End Property

Public Property Set FrameFont(ByVal vNewFont As Font)
    
    Set vFrameFont = vNewFont
    
    PropertyChanged "FrameFont"
    
    ReDrawFrame
    
End Property

Public Property Get FrameOpacity() As Integer
    
    FrameOpacity = vFrameOpacity
    
End Property

Public Property Let FrameOpacity(ByVal vNewOpacity As Integer)
    
    vFrameOpacity = vNewOpacity
    
    PropertyChanged "FrameOpacity"
    
    ReDrawFrame
    
End Property

Public Property Get FontColour() As OLE_COLOR
    
    FontColour = vFontColour
    
End Property

Public Property Let FontColour(ByVal vNewColour As OLE_COLOR)
    
    vFontColour = vNewColour
    
    PropertyChanged "FontColour"
    
    ReDrawFrame
    
End Property
