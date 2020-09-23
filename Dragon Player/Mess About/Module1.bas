Attribute VB_Name = "Module1"
Option Explicit

'Imaging
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


'Drag
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

'Drag And Skin
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Public Const RGN_OR = 2
    
Public Type NavControls
    x As Integer
    y As Integer
    Width As Integer
    Height As Integer
    Visible As Boolean
    End Type
    Public NavControls As NavControls

Public Sub SkinForm(Frm As Form, MaskPic As PictureBox, Optional TransColor As Long)
    
    Dim RetR As Long
    Dim RgnFinal As Long
    Dim RgnTmp As Long
    
    Dim hHeight As Long
    Dim wWidth As Long
    
    Dim Col As Long
    Dim Start As Long
    Dim RowR As Long
    
    MaskPic.AutoSize = True
    MaskPic.AutoRedraw = True

    With Frm
        .Height = MaskPic.Height
        .Width = MaskPic.Width
    End With

    If TransColor < 1 Then
        TransColor = GetPixel(MaskPic.hdc, 0, 0)
    End If
    
    hHeight = MaskPic.Height / Screen.TwipsPerPixelY
    wWidth = MaskPic.Width / Screen.TwipsPerPixelX
    
    RgnFinal = CreateRectRgn(0, 0, 0, 0)

    For RowR = 0 To hHeight - 1
        
        Col = 0

        Do While Col < wWidth

            Do While Col < wWidth And GetPixel(MaskPic.hdc, Col, RowR) = TransColor
                
                Col = Col + 1
            Loop

            If Col < wWidth Then
                Start = Col

                Do While Col < wWidth And GetPixel(MaskPic.hdc, Col, RowR) <> TransColor
                    Col = Col + 1
                Loop
                
                If Col > wWidth Then Col = wWidth
                
                RgnTmp = CreateRectRgn(Start, RowR, Col, RowR + 1)
                RetR = CombineRgn(RgnFinal, RgnFinal, RgnTmp, RGN_OR)
                DeleteObject (RgnTmp)
                
            End If
        Loop
        
    Next RowR

    RetR = SetWindowRgn(Frm.hwnd, RgnFinal, True)

End Sub

Public Function LoadGraphicDC(sFileName As String) As Long
'cheap error handling
'On Error Resume Next

'temp variable to hold our DC address
    Dim LoadGraphicDCTEMP As Long

'create the DC address compatible with
'the DC of the screen
    LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))

'load the graphic file into the DC...
    SelectObject LoadGraphicDCTEMP, LoadPicture(sFileName)

'return the address of the file
    LoadGraphicDC = LoadGraphicDCTEMP
    
End Function

Public Function P2TX(Value As Integer) As Integer
    
    T2PX = Value * Screen.TwipsPerPixelX
    
End Function

Public Function P2TY(Value As Integer) As Integer
    
    T2PY = Value * Screen.TwipsPerPixelY
    
End Function

Public Function Blender(FromColour As Long, ToColour As Long, ColourSteps As Integer, CurrentStep As Integer) As Long
    
    Dim RF As Byte, GF As Byte, BF As Byte
    Dim RT As Byte, GT As Byte, BT As Byte
    Dim DR As Double, DG As Double, DB As Double
    
    CRGB FromColour, RF, GF, BF
    CRGB ToColour, RT, GT, BT
    
    DR = ((CDbl(RT) - RF) / ColourSteps) + 0.5
    DG = ((CDbl(GT) - GF) / ColourSteps) + 0.5
    DB = ((CDbl(BT) - BF) / ColourSteps) + 0.5
    
    Blender = RGB(RF + Int(DR * CurrentStep), GF + Int(DG * CurrentStep), BF + Int(DB * CurrentStep))
    
End Function

Public Function CRGB(LongColour As Long, Optional Red As Byte, Optional Green As Byte, Optional Blue As Byte)

        Red = LongColour And 255
        Green = (LongColour \ 256) And 255
        Blue = (LongColour \ 65536) And 255

End Function


