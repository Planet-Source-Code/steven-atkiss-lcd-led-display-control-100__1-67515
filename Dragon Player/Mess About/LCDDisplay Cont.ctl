VERSION 5.00
Begin VB.UserControl LCDDisplay 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillStyle       =   0  'Solid
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox Canvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   405
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   37
      TabIndex        =   0
      Top             =   2340
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Character 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   2115
      Visible         =   0   'False
      Width           =   90
   End
End
Attribute VB_Name = "LCDDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'LCD CONTROL BY S.ATKISS

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long


'loading sprites
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'end of copy-paste here...

'our Buffer's DC
Private myBufferBMP As Long
Private myBackBuffer As Long

Private myImageDC As Long

Private Type Alpha
    Width As Integer
    Height As Integer
    Pnt() As Integer
    End Type
    Private Alpha() As Alpha
    Private MtxImage() As Alpha
    
Private Type POINTAPI
    X As Long
    Y As Long
    End Type

Private vStartVisible           As Boolean
Private vUseFontAntiAliasing     As Boolean
Private vBackColour             As OLE_COLOR
Private vElementColour          As OLE_COLOR
Private vElementOpacity         As Integer
Private vDisplayFont            As Font
Private vYOffSet                As Integer
Private vCharacterSpaceOffSet   As Integer
Private vDisplayPosition        As Integer
Private vText                   As String
Private vImagePlaceHolder       As String
Private vImageCollection        As Picture
Private vImageRows              As Integer
Private vImageCols              As Integer
Private vImageCellWidth         As Integer
Private vImageCellHeight        As Integer
Private vBounce                 As Boolean
Private vDisplayErrors          As Boolean
Private vElementHeight          As Integer
Private vElementWidth           As Integer
Private vElementGlow            As Boolean
Private vGlowIntencity          As Integer
Private vLCDScaleWidth          As Integer
Private vLCDScaleHeight         As Integer
Private vSnapToElement          As Boolean
Private ElementGlowColour       As Long
Private vFadeLeft               As Boolean
Private vFadeRight              As Boolean

Private Const Def_GlowIntencity = 20
Private Const Def_ElementHeight = 4
Private Const Def_ElementWidth = 4

Public Enum ElementStyle
    [Square] = 0
    '[Oval] = 1
    End Enum
Private vElementType            As ElementStyle

Public Enum LCDScaleModeAs
    [Twips] = 0
    [Pixel] = 1
    [Elements] = 2
    End Enum
Private vLCDScaleMode              As LCDScaleModeAs

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)

Private GenHeight As Integer
Private WidthCount As Integer 'Width Of Matrix String
Private MatrixString() As Integer
Private MaxLead As Integer
Private ScrollLeft As Boolean

Private Resizeing As Boolean
Private Initiated As Boolean


Private Sub UserControl_Click()
    
    RaiseEvent Click
    
End Sub

Private Sub UserControl_DblClick()
    
    RaiseEvent DblClick
    
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
        
    RaiseEvent KeyPress(KeyAscii)
        
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub


Private Sub UserControl_InitProperties()

    vBackColour = vbBlack
    vElementColour = vbWhite
    vElementOpacity = 100
    Set vDisplayFont = Ambient.Font
    vCharacterSpaceOffSet = 1
    vYOffSet = 0
    vDisplayPosition = 0
    vText = UserControl.Name
    vImagePlaceHolder = ""
    Set vImageCollection = Nothing
    vImageRows = 1
    vImageCols = 1
    vImageCellWidth = 0
    vImageCellHeight = 0
    vBounce = False
    vDisplayErrors = False
    vElementType = [Square]
    vElementHeight = Def_ElementHeight
    vElementWidth = Def_ElementWidth
    vElementGlow = True
    vGlowIntencity = Def_GlowIntencity
    vLCDScaleMode = [Elements]
    vSnapToElement = True
    vUseFontAntiAliasing = True
    vStartVisible = True
    vFadeLeft = True
    vFadeRight = True
    
    InitiateControl
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        vBackColour = .ReadProperty("BackColour", vbBlack)
        vElementColour = .ReadProperty("ElementColour", vbWhite)
        vElementOpacity = .ReadProperty("ElementOpacity", 100)
        Set vDisplayFont = .ReadProperty("DisplayFont", Ambient.Font)
        vYOffSet = .ReadProperty("YOffSet", 0)
        vCharacterSpaceOffSet = .ReadProperty("CharacterSpaceOffSet", 1)
        vDisplayPosition = .ReadProperty("DisplayPosition", 0)
        vText = .ReadProperty("Text", UserControl.Name)
        vImagePlaceHolder = .ReadProperty("ImagePlaceHolder", "")
        Set vImageCollection = .ReadProperty("ImageCollection", Nothing)
        vImageRows = .ReadProperty("ImageRows", 1)
        vImageCols = .ReadProperty("ImageCols", 1)
        vImageCellWidth = .ReadProperty("ImageCellWidth", 0)
        vImageCellHeight = .ReadProperty("ImageCellHeight", 0)
        vBounce = .ReadProperty("Bounce", False)
        vDisplayErrors = .ReadProperty("DisplayErrors", False)
        vElementType = .ReadProperty("ElementType", [Square])
        vElementHeight = .ReadProperty("ElementHeight", Def_ElementHeight)
        vElementWidth = .ReadProperty("ElementWidth", Def_ElementWidth)
        vElementGlow = .ReadProperty("ElementGlow", True)
        vGlowIntencity = .ReadProperty("GlowIntencity", Def_GlowIntencity)
        vLCDScaleMode = .ReadProperty("LCDScaleMode", [Elements])
        vLCDScaleWidth = .ReadProperty("LCDScaleWidth", 0)
        vLCDScaleHeight = .ReadProperty("LCDScaleHeight", 0)
        vSnapToElement = .ReadProperty("SnapToElement", True)
        vUseFontAntiAliasing = .ReadProperty("UseFontAntiAliasing", True)
        vDisplayErrors = .ReadProperty("DisplayErrors", False)
        vStartVisible = .ReadProperty("StartVisible", True)
        vFadeLeft = .ReadProperty("FadeLeft", True)
        vFadeRight = .ReadProperty("FadeRight", True)
    End With
    
    
    InitiateControl
    UpdateDisplay
    
End Sub

Private Sub UserControl_Resize()
    
    If SnapToElement = True Then
        If Resizeing = False Then
            Resizeing = True
            Height = ((Int(ScaleHeight / ElementHeight) * ElementHeight) + 1) * Screen.TwipsPerPixelY
            Width = ((Int(ScaleWidth / ElementWidth) * ElementWidth) + 1) * Screen.TwipsPerPixelX
        Else
            Exit Sub
        End If
        
        Resizeing = False
        
    End If
    
    If Initiated = True Then DrawElements

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    With PropBag
        .WriteProperty "BackColour", vBackColour, vbBlack
        .WriteProperty "ElementColour", vElementColour, vbWhite
        .WriteProperty "ElementOpacity", vElementOpacity, 100
        .WriteProperty "DisplayFont", vDisplayFont, Ambient.Font
        .WriteProperty "YOffSet", vYOffSet, 0
        .WriteProperty "CharacterSpaceOffSet", vCharacterSpaceOffSet, 1
        .WriteProperty "DisplayPosition", vDisplayPosition, 0
        .WriteProperty "Text", vText, UserControl.Name
        .WriteProperty "ImagePlaceHolder", vImagePlaceHolder, ""
        .WriteProperty "ImageCollection", vImageCollection, Nothing
        .WriteProperty "ImageRows", vImageRows, 1
        .WriteProperty "ImageCols", vImageCols, 1
        .WriteProperty "ImageCellWidth", vImageCellWidth, 0
        .WriteProperty "ImageCellHeight", vImageCellHeight, 0
        .WriteProperty "Bounce", vBounce, False
        .WriteProperty "DisplayErrors", vDisplayErrors, False
        .WriteProperty "ElementType", vElementType, [Square]
        .WriteProperty "ElementHeight", vElementHeight, Def_ElementHeight
        .WriteProperty "ElementWidth", vElementWidth, Def_ElementWidth
        .WriteProperty "ElementGlow", vElementGlow, True
        .WriteProperty "GlowIntencity", vGlowIntencity, Def_GlowIntencity
        .WriteProperty "LCDScaleMode", vLCDScaleMode, [Elements]
        .WriteProperty "LCDScaleWidth", vLCDScaleWidth, 0
        .WriteProperty "LCDScaleHeight", vLCDScaleHeight, 0
        .WriteProperty "SnapToElement", vSnapToElement, True
        .WriteProperty "UseFontAntiAliasing", vUseFontAntiAliasing, True
        .WriteProperty "DisplayErrors", vDisplayErrors, False
        .WriteProperty "StartVisible", vStartVisible, True
        .WriteProperty "FadeLeft", vFadeLeft, True
        .WriteProperty "FadeRight", vFadeRight, True
    End With
    
End Sub

Public Property Get BackColour() As OLE_COLOR
    
    BackColour = vBackColour
    
End Property

Public Property Let BackColour(ByVal vNewColour As OLE_COLOR)
    
    vBackColour = vNewColour
    
    PropertyChanged "BackColour"
    
    DrawElements
    
End Property


Public Property Get ElementColour() As OLE_COLOR
    
    ElementColour = vElementColour
    
End Property

Public Property Let ElementColour(ByVal vNewColour As OLE_COLOR)
    
    vElementColour = vNewColour
    
    PropertyChanged "ElementColour"
    
    DrawElements
    
End Property


Public Property Get ElementOpacity() As Integer
    
    ElementOpacity = vElementOpacity
    
End Property

Public Property Let ElementOpacity(ByVal vNewOpacity As Integer)
        
    If vNewOpacity < 0 Then vNewOpacity = 0
    If vNewOpacity > 100 Then vNewOpacity = 100
    
    vElementOpacity = vNewOpacity
    
    PropertyChanged "ElementOpacity"
    
    DrawElements
    
End Property


Public Property Get DisplayFont() As Font
    
    Set DisplayFont = vDisplayFont
    
End Property

Public Property Set DisplayFont(ByVal vNewFont As Font)
    
    Set vDisplayFont = vNewFont
    
    PropertyChanged "DisplayFont"

    InitiateControl
    DrawElements
    
End Property

Public Property Get YOffSet() As Integer
    
    YOffSet = vYOffSet
    
End Property

Public Property Let YOffSet(ByVal vNewYOffSet As Integer)
    
    If vNewYOffSet < -100 Then vNewYOffSet = -100
    If vNewYOffSet > 100 Then vNewYOffSet = 100
    
    vYOffSet = vNewYOffSet
        
    PropertyChanged "YOffSet"
    
    DrawElements
    
End Property

Public Property Get CharacterSpaceOffSet() As Integer
    
    CharacterSpaceOffSet = vCharacterSpaceOffSet
    
End Property

Public Property Let CharacterSpaceOffSet(ByVal vNewCharacterSpaceOffSet As Integer)
    
    If vNewCharacterSpaceOffSet < -5 Then vNewCharacterSpaceOffSet = -5
    If vNewCharacterSpaceOffSet > 5 Then vNewCharacterSpaceOffSet = 5
    
    vCharacterSpaceOffSet = vNewCharacterSpaceOffSet
    
    PropertyChanged "CharacterSpaceOffSet"
    
    BuildString
    DrawElements
    
End Property



Public Property Get DisplayPosition() As Integer
    
    DisplayPosition = vDisplayPosition
    
End Property

Private Property Let DisplayPosition(ByVal vNewPosition As Integer)
    
    vDisplayPosition = vNewPosition
    
    PropertyChanged "DisplayPosition"
    
End Property

Public Property Get Text() As String
    
    Text = vText
    
End Property

Public Property Let Text(ByVal vNewText As String)
    
    vText = vNewText
    
    PropertyChanged "Text"
    
    BuildMatrixString vNewText
    
    DrawElements
    
    Bounce = Bounce
    
End Property

Public Property Get ImagePlaceHolder() As String
    
    ImagePlaceHolder = vImagePlaceHolder
    
End Property

Public Property Let ImagePlaceHolder(ByVal vNewPlaceHolder As String)
    
    If IsNumeric(vNewPlaceHolder) = True Then
        If DisplayErrors = True Then
            MsgBox "Cannot Use Numbers as Image Place Holders.", vbOKOnly + vbInformation
        End If
        Exit Property
    End If
    
    vImagePlaceHolder = vNewPlaceHolder
    
    PropertyChanged "ImagePlaceHolder"
    
    BuildImageMatrix
    BuildMatrixString Text
    
End Property

Public Property Get ImageCollection() As Picture
    
    Set ImageCollection = vImageCollection
    
End Property
Public Property Set ImageCollection(ByVal vNewImage As Picture)
    
    Set vImageCollection = vNewImage
    
    PropertyChanged "ImageCollection"
    
    vImagePlaceHolder = ""
    vImageCols = 1
    vImageRows = 1
    PropertyChanged "ImagePlaceHolder"
    PropertyChanged "ImageCols"
    PropertyChanged "ImageRows"

End Property



Public Property Get ImageRows() As Integer
    
    ImageRows = vImageRows
    
End Property

Public Property Let ImageRows(ByVal vNewRows As Integer)
    
    If vNewRows < 1 Then vNewRows = 1
    If vNewRows > 1000 Then vNewRows = 1000
    
    If vNewRows > CollectionImageHeight Then
        If DisplayErrors = True Then
            MsgBox "Image Not High Enough To Allow For This Number Of Rows", vbOKOnly + vbInformation
        End If
        Exit Property
    End If
    
    vImageRows = vNewRows
    
    PropertyChanged "vImageRows"
    
   LoadCollectionImage
    
End Property

Public Property Get ImageCols() As Integer
    
    ImageCols = vImageCols
    
End Property

Public Property Let ImageCols(ByVal vNewCols As Integer)
    
    
    
    If vNewCols < 1 Then vNewCols = 1
    If vNewCols > 100 Then vNewCols = 100
    
    If vNewCols > CollectionImageWidth Then
        If DisplayErrors = True Then
            MsgBox "Image Not Wide Enough To Allow For This Number Of Columns", vbOKOnly + vbInformation
        End If
        Exit Property
    End If
    
    vImageCols = vNewCols
    
    PropertyChanged "ImageCols"
    
    LoadCollectionImage
    
End Property

Public Property Get ImageCellWidth() As Integer
    
    ImageCellWidth = vImageCellWidth
    
End Property

Private Property Let ImageCellWidth(ByVal vNewCellWidth As Integer)
    
    vImageCellWidth = vNewCellWidth
    
    PropertyChanged "ImageCellWidth"
    
End Property

Public Property Get ImageCellHeight() As Integer
    
    ImageCellHeight = vImageCellHeight
    
End Property

Private Property Let ImageCellHeight(ByVal vNewCellHeight As Integer)
    
    vImageCellHeight = vNewCellHeight
    
    PropertyChanged "ImageCellHeight"
    
End Property

Public Property Get Bounce() As Boolean
    
    Bounce = vBounce
    
End Property

Public Property Let Bounce(ByVal vNewBounce As Boolean)
    
    If WidthCount - (MaxLead * 2) < ScaleWidth / ElementWidth And vNewBounce = True Then
        If DisplayErrors = True Then
            MsgBox "Text Is Too Short To Bounce", vbOKOnly + vbInformation
        End If
        vNewBounce = False
    End If
    
    vBounce = vNewBounce
    
    PropertyChanged "Bounce"
    
End Property

Public Property Get DisplayErrors() As Boolean
    
    DisplayErrors = vDisplayErrors
    
End Property

Public Property Let DisplayErrors(ByVal vNewDisplayErrors As Boolean)
    
    vDisplayErrors = vNewDisplayErrors
    
    PropertyChanged "DisplayErrors"
    
End Property

Public Property Get ElementType() As ElementStyle
    
    ElementType = vElementType
    
End Property

Public Property Let ElementType(ByVal vNewElementType As ElementStyle)
    
    vElementType = vNewElementType
    
    PropertyChanged "ElementType"
    
    DrawElements
    
End Property

Public Property Get ElementHeight() As Integer
    
    ElementHeight = vElementHeight
    
End Property

Public Property Let ElementHeight(ByVal vNewElementHeight As Integer)
    
    If vNewElementHeight < 1 Then vNewElementHeight = 1
    If vNewElementHeight > 10 Then vNewElementHeight = 10
    
    vElementHeight = vNewElementHeight
    
    PropertyChanged "ElementHeight"
    
    UserControl_Resize
    
End Property

Public Property Get ElementWidth() As Integer
    
    ElementWidth = vElementWidth
    
End Property

Public Property Let ElementWidth(ByVal vNewElementWidth As Integer)
    
    If vNewElementWidth < 1 Then vNewElementWidth = 1
    If vNewElementWidth > 10 Then vNewElementWidth = 10
    
    vElementWidth = vNewElementWidth
    
    PropertyChanged "ElementWidth"
    
    UserControl_Resize
    
End Property



Public Property Get ElementGlow() As Boolean
    
    ElementGlow = vElementGlow
    
End Property

Public Property Let ElementGlow(ByVal vNewElementGlow As Boolean)
    
    vElementGlow = vNewElementGlow
    
    PropertyChanged "ElementGlow"
    
    DrawElements
    
End Property

Public Property Get GlowIntencity() As Integer
    
    GlowIntencity = vGlowIntencity
    
End Property

Public Property Let GlowIntencity(ByVal vNewGlowIntencity As Integer)
    
    If vNewGlowIntencity < 0 Then vNewGlowIntencity = 0
    If vNewGlowIntencity > 100 Then vGlowIntencity = 100
    
    vGlowIntencity = vNewGlowIntencity
    
    PropertyChanged "GlowIntencity"
    
    DrawElements
    
End Property



Public Property Get LCDScaleMode() As LCDScaleModeAs
    
    LCDScaleMode = vLCDScaleMode
    
End Property

Public Property Let LCDScaleMode(ByVal vNewLCDScaleMode As LCDScaleModeAs)
    
    vLCDScaleMode = vNewLCDScaleMode
    
    PropertyChanged "LCDScaleMode"
    
End Property

Public Property Get LCDScaleWidth() As Integer
    
    Select Case LCDScaleMode
        Case [Twips]
            LCDScaleWidth = UserControl.Width
            
        Case [Pixel]
            LCDScaleWidth = UserControl.ScaleWidth
            
        Case [Elements]
            LCDScaleWidth = UserControl.ScaleWidth / ElementWidth
            
    End Select
    
End Property

Public Property Let LCDScaleWidth(ByVal vNewLCDScaleWidth As Integer)
    
    MsgBox "This Property Is Read Only.", vbOKOnly + vbInformation
    
End Property

Public Property Get LCDScaleHeight() As Integer
    
    Select Case LCDScaleMode
        Case [Twips]
            LCDScaleHeight = UserControl.Height
            
        Case [Pixel]
            LCDScaleHeight = UserControl.ScaleHeight
            
        Case [Elements]
            LCDScaleHeight = UserControl.ScaleHeight / ElementHeight
            
    End Select
    
End Property

Public Property Let LCDScaleHeight(ByVal vNewLCDScaleHeight As Integer)
    
    If DisplayErrors = True Then MsgBox "This Property Is Read Only.", vbOKOnly + vbInformation
    
End Property

Public Property Get SnapToElement() As Boolean
    
    If DisplayErrors = True Then SnapToElement = vSnapToElement
    
End Property

Public Property Let SnapToElement(ByVal vNewSnap As Boolean)
    
    vSnapToElement = vNewSnap
    
    PropertyChanged "SnapToElement"
    
    UserControl_Resize
    
End Property

Private Sub DrawElements()
    
    If MxCharCount = -1 Or WidthCount = 0 Then InitiateControl
    
    UserControl.Cls
    
    Dim FillColour As Long
    Dim X As Integer
    Dim Y As Integer
    Dim ActualElementColour As Long
    Dim ActualWidth As Integer
    Dim ActualHeight As Integer
    Dim pPoint As POINTAPI
    Dim EdgeColour As Long
    
    ActualWidth = Int(ScaleWidth / ElementWidth)
    ActualHeight = Int(ScaleHeight / ElementHeight)
    
    ElementGlowColour = GetOpacityColourEx(ElementColour, BackColour, GlowIntencity)
    
    If ElementGlow = True Then
        ActualElementColour = GetOpacityColourEx(ElementColour, ElementGlowColour, ElementOpacity)
    Else
        ActualElementColour = GetOpacityColourEx(ElementColour, BackColour, ElementOpacity)
    End If

    ForeColor = BackColour
    BackColor = BackColour
    
    Dim AX As Integer, AY As Integer

    For Y = 0 To ActualHeight + Abs(YOffSet)
        For X = 0 To ActualWidth

            AX = X + DisplayPosition
            AY = Y
            
            If vElementGlow = True Then
                FillColor = ElementGlowColour
                FillColour = ElementGlowColour
            Else
                FillColor = BackColour
                FillColour = BackColour
            End If

            If AY <= GenHeight And AX <= WidthCount Then
                If MatrixString(AX, AY) <> 0 Then
                    If UseFontAntiAliasing = True Then
                        If vElementGlow = True Then
                            FillColour = GetOpacityColourEx(ActualElementColour, ElementGlowColour, CInt(MatrixString(AX, AY)))
                        Else
                            FillColour = GetOpacityColourEx(ActualElementColour, BackColour, CInt(MatrixString(AX, AY)))
                        End If
                    Else
                        FillColour = ActualElementColour
                    End If
                    
                    If FadeLeft = True Then
                        If X < 10 Then
                            If vElementGlow = True Then
                                FillColour = GetOpacityColourEx(FillColour, ElementGlowColour, X * 10)
                            Else
                                FillColour = GetOpacityColourEx(FillColour, BackColour, X * 10)
                            End If
                            
                        End If
                    End If
                    
                    If FadeRight = True Then
                        If X > ActualWidth - 10 Then
                            If vElementGlow = True Then
                                FillColour = GetOpacityColourEx(FillColour, ElementGlowColour, (ActualWidth - X) * 10)
                            Else
                                FillColour = GetOpacityColourEx(FillColour, BackColour, (ActualWidth - X) * 10)
                            End If
                        End If
                    End If
                    
                    SetPixelV myBackBuffer, X, Y + (YOffSet), FillColour
                Else
                    
                    SetPixelV myBackBuffer, X, Y + (YOffSet), FillColour
                End If
            Else
                
                SetPixelV myBackBuffer, X, Y + (YOffSet), FillColour
            End If
            
        Next X
    Next Y
    
    For Y = 0 To YOffSet
        For X = 0 To ActualWidth
        
        If vElementGlow = True Then
                FillColor = ElementGlowColour
                FillColour = ElementGlowColour
            Else
                FillColor = BackColour
                FillColour = BackColour
            End If
            
            SetPixelV myBackBuffer, X, Y, FillColour
        Next X
    Next Y
    
    StretchBlt UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight, myBackBuffer, 0, 0, ActualWidth, ActualHeight, vbSrcCopy
    
    If ElementHeight > 1 And ElementWidth > 1 Then
        For X = 0 To ScaleWidth Step ElementWidth
            MoveToEx UserControl.hdc, X, 0, pPoint
            LineTo UserControl.hdc, X, ScaleHeight
        Next X
        
        For Y = 0 To ScaleHeight Step ElementHeight
            MoveToEx UserControl.hdc, 0, Y, pPoint
            LineTo UserControl.hdc, ScaleWidth, Y
        Next Y
    End If
    
End Sub

Public Sub UpdateDisplay(Optional Scroll As Boolean = False)
        
    If Scroll = True Then
        If Bounce = False Then
            vDisplayPosition = DisplayPosition + 1
            If vDisplayPosition >= WidthCount - (MaxLead) Then vDisplayPosition = 0
        Else
            If ScrollLeft = True Then
            
                vDisplayPosition = DisplayPosition + 1
                If vDisplayPosition > WidthCount - (MaxLead * 2) Then ScrollLeft = False
                
            Else
                vDisplayPosition = DisplayPosition - 1
                If vDisplayPosition < MaxLead Then ScrollLeft = True
            End If
            
        End If
        PropertyChanged "Displayposition"
    End If
    DrawElements
    
End Sub

Private Sub InitiateControl()
    
    Erase Alpha
    Erase MatrixString
    
    DeleteObject myBufferBMP
    DeleteDC myBackBuffer
    
    myBackBuffer = CreateCompatibleDC(GetDC(0))
    myBufferBMP = CreateCompatibleBitmap(GetDC(0), UserControl.ScaleWidth, UserControl.ScaleHeight)
    
    SelectObject myBackBuffer, myBufferBMP
    
    LoadCollectionImage
    
    BuildAlphaMatrix
    
    If StartVisible = True Then
        vDisplayPosition = MaxLead
    Else
        vDisplayPosition = 0
    End If
    
    ScrollLeft = True
    PropertyChanged "DisplayPosition"
    
    DrawElements
    
End Sub

Private Sub LoadCollectionImage()
    
    Erase MtxImage
    
    If isImageLoaded = False Then Exit Sub
    
    DeleteDC myImageDC
    
    myImageDC = CreateCompatibleDC(GetDC(0))
    SelectObject myImageDC, vImageCollection
    
    ImageCellWidth = CollectionImageWidth \ ImageCols
    ImageCellHeight = CollectionImageHeight \ ImageRows
    
    ConvertToGreyScale
    BuildImageMatrix
    BuildMatrixString Text
    
End Sub

Private Sub ConvertToGreyScale()

    Dim X As Long
    Dim Y As Long
    Dim Red As Byte
    Dim Green As Byte
    Dim Blue As Byte
    Dim LongColour As Long
    Dim AmbientColour As Long
    
    For X = 0 To CollectionImageWidth
        For Y = 0 To CollectionImageHeight
            LongColour = GetPixel(myImageDC, X, Y)
            CRGB LongColour, Red, Green, Blue
            
            AmbientColour = (222 * CLng(Red) + 707 * CLng(Green) + 71 * CLng(Blue)) / 1000
            
            SetPixelV myImageDC, X, Y, RGB(AmbientColour, AmbientColour, AmbientColour)
            
        Next Y
    Next X
    
End Sub

Private Function isImageLoaded() As Boolean
On Error GoTo LoadedError
    
    If ImageCollection.Width <> 0 Then isImageLoaded = True

Exit Function
LoadedError:
    isImageLoaded = False
End Function

Public Function CollectionImageWidth() As Integer
    
    CollectionImageWidth = ScaleX(ImageCollection.Width)
    
End Function

Public Function CollectionImageHeight() As Integer
    
    CollectionImageHeight = ScaleY(ImageCollection.Height)
    
End Function


Public Sub BuildString()
    
    BuildMatrixString Text

End Sub

Private Function CRGB(LongColour As Long, Optional Red As Byte, Optional Green As Byte, Optional Blue As Byte)

        Red = LongColour And 255
        Green = (LongColour \ 256) And 255
        Blue = (LongColour \ 65536) And 255

End Function

Private Function GetOpacityColourEx(Colour1 As Long, Colour2 As Long, Opacity As Integer) As Long
    
    If Opacity >= 100 Then GetOpacityColourEx = Colour1: Exit Function
    If Opacity <= 0 Then GetOpacityColourEx = Colour2: Exit Function
    
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

Private Function MxCharCount() As Integer
On Error GoTo MatrixError
    
    MxCharCount = UBound(Alpha)
    
Exit Function
MatrixError:
    MxCharCount = -1
End Function

'Read Character Set Into Matrix Format
Private Sub BuildAlphaMatrix()
    
    
    Initiated = False
    
    Set Canvas.Font = DisplayFont
    Set Character.Font = DisplayFont
    
    Erase Alpha
    
    Dim LP As Integer
    Dim X As Integer, Y As Integer
     
    For LP = 33 To 126
        Character.Caption = Chr(LP)
        Canvas.Width = Character.Width
        Canvas.Height = Character.Height
        Canvas.Cls
        Canvas.Picture = Nothing
        Canvas.Print Chr(LP)
        Canvas.Picture = Canvas.Image
        
        ReDim Preserve Alpha(LP)
        Alpha(LP).Width = Canvas.ScaleWidth
        Alpha(LP).Height = Canvas.ScaleHeight
        
        ReDim Preserve Alpha(LP).Pnt(Alpha(LP).Width, Alpha(LP).Height)
        
        Dim Offx As Integer
        Dim RetR As Byte
        
        Offx = 0
        
        For Y = 0 To Canvas.ScaleHeight
            For X = Offx To Canvas.ScaleWidth - Offx
                DoEvents
                If Canvas.Point(X, Y) <> vbWhite Then
                    CRGB Canvas.Point(X, Y), , RetR
                    Alpha(LP).Pnt(X - Offx, Y) = 100 - CPerc(RetR)
                Else
                    Alpha(LP).Pnt(X - Offx, Y) = 0
                End If
            Next X
        Next Y
    Next LP
    
    Alpha(32).Width = 4 + CharacterSpaceOffSet
    Alpha(32).Height = Canvas.ScaleHeight
    ReDim Preserve Alpha(32).Pnt(Alpha(32).Width, Alpha(32).Height)
    
    GenHeight = Canvas.ScaleHeight
    
    If isImageLoaded = True Then BuildImageMatrix
    
    BuildMatrixString Text
    
End Sub

Private Sub BuildImageMatrix()

    Erase MtxImage
    
    Dim RetR As Byte
    Dim X As Integer, Y As Integer, ImageIndex As Integer
    Dim Column As Integer, Row As Integer
    
    'ImageIndex = ImageCols * ImageRows
    
    For Row = 0 To ImageRows - 1
        For Column = 0 To ImageCols - 1
            ReDim Preserve MtxImage(ImageIndex)
            MtxImage(ImageIndex).Height = ImageCellHeight
            MtxImage(ImageIndex).Width = ImageCellWidth
            
            ReDim Preserve MtxImage(ImageIndex).Pnt(ImageCellWidth, ImageCellHeight)
           
            For Y = 0 To MtxImage(ImageIndex).Height
                For X = 0 To MtxImage(ImageIndex).Width
                    CRGB GetPixel(myImageDC, X + (Column * MtxImage(ImageIndex).Width), Y + (Row * MtxImage(ImageIndex).Width)), RetR
                    MtxImage(ImageIndex).Pnt(X, Y) = CPerc(RetR)
                Next X
            Next Y
            
            ImageIndex = ImageIndex + 1
            
        Next Column
    Next Row
    
End Sub

Private Function CPerc(Colour As Byte) As Byte

    CPerc = Int((Colour / 255) * 100)
    
End Function

'Build Full String In Matrix Format
Private Sub BuildMatrixString(MyString As String)
    
    If MxCharCount = -1 Then BuildAlphaMatrix
    
    Dim LP As Integer
    Dim X As Integer, Y As Integer
    Dim MyAscii As Integer
    Dim MxPos As Integer
    Dim StartIndex As Integer 'Holds String Position Where IPH Index Starts
    Dim ImageIndex As Integer
    Dim ImagesWidth As Integer
    Dim LoopTo As Long
    
    Dim SubLp As Integer
    
    MaxLead = UserControl.ScaleWidth / ElementWidth
    WidthCount = 0
    
    Initiated = False
    Erase MatrixString
    
    For LP = 1 To Len(MyString)
        MyAscii = Asc(Mid$(MyString, LP, 1))
        WidthCount = WidthCount + Alpha(MyAscii).Width
    Next LP
    
    ImagesWidth = GetTotalImagesWidth(MyString)
    
    WidthCount = WidthCount + (CharacterSpaceOffSet - 1) * Len(MyString) + (ImagesWidth)
    
    ReDim MatrixString((WidthCount + (MaxLead * 2)), Alpha(103).Height)
    
    WidthCount = WidthCount + (MaxLead * 2)
    
    Dim IndexWidth As Integer
    
    For LP = 1 To Len(MyString)
        
        If ImagePlaceHolder <> "" And Mid$(MyString, LP, Len(ImagePlaceHolder)) = ImagePlaceHolder Then
            
            LP = LP + Len(ImagePlaceHolder) - 1
            
            StartIndex = LP + 1
            
            For SubLp = StartIndex To Len(MyString)
                If IsNumeric(Mid$(MyString, SubLp, 1)) = False Or SubLp = Len(MyString) Then
                    
                    IndexWidth = SubLp - StartIndex
                    
                    'If The Search Is At The End Of The String Then IndexWidth Will
                    'Return 0 When It Should Be One.
                    If IndexWidth = 0 Then IndexWidth = 1: SubLp = SubLp + 1
                    
                    If IsNumeric(Mid$(MyString, StartIndex, (IndexWidth))) = True Then
                       
                        ImageIndex = CInt(Mid$(MyString, StartIndex, (IndexWidth)))

                        If ImageCount > 0 And ImageIndex <= ImageCount - 1 Then
                            For X = 0 To MtxImage(ImageIndex).Width - 1
                            
                                If GenHeight > MtxImage(ImageIndex).Height Then
                                    LoopTo = MtxImage(ImageIndex).Height - 1
                                Else
                                    LoopTo = GenHeight - 1
                                End If
                                
                                For Y = 0 To LoopTo
                                
                                    If MtxImage(ImageIndex).Pnt(X, Y) <> 0 Then
                                        MatrixString(MxPos + MaxLead, Y) = MtxImage(ImageIndex).Pnt(X, Y)
                                    ElseIf MtxImage(ImageIndex).Pnt(X, Y) = 0 Then
                                        MatrixString(MxPos + MaxLead, Y) = 0
                                    End If
                                    
                                Next Y
                                MxPos = MxPos + 1
                            Next X
                        Else
                            If DisplayErrors = True Then
                                MsgBox "Image Index Doesn't Exist", vbOKOnly + vbInformation
                            End If
                        End If
                        
                    End If
                    LP = SubLp - 1: Exit For
                End If
            Next SubLp
        Else
            MyAscii = Asc(Mid$(MyString, LP, 1))
            
            For X = 0 To Alpha(MyAscii).Width - 1
            
                For Y = 0 To Alpha(MyAscii).Height - 1
                    
                    If Alpha(MyAscii).Pnt(X, Y) <> 0 Then
                        MatrixString(MxPos + MaxLead, Y) = Alpha(MyAscii).Pnt(X, Y)
                    ElseIf MatrixString(MxPos + MaxLead, Y) = 0 Then
                        MatrixString(MxPos + MaxLead, Y) = 0
                    End If
                    
                Next Y
                
                MxPos = MxPos + 1
                
            Next X
            
            MxPos = MxPos + (CharacterSpaceOffSet - 1)
        End If
    Next LP
    
    Initiated = True

    DrawElements
    
End Sub

Private Function GetTotalImagesWidth(MyString As String) As Integer
    
    If ImageCount = 0 Then Exit Function
    
    Dim LP As Integer
    Dim X As Integer, Y As Integer
    Dim MyAscii As Integer
    Dim MxPos As Integer
    Dim StartIndex As Integer 'Holds String Position Where IPH Index Starts
    Dim ImageIndex As Integer
    Dim SubLp As Integer
    Dim IndexWidth
    
    
    For LP = 1 To Len(MyString)
        
        If ImagePlaceHolder <> "" And Mid$(MyString, LP, Len(ImagePlaceHolder)) = ImagePlaceHolder Then
            
            LP = LP + Len(ImagePlaceHolder) - 1
            
            StartIndex = LP + 1
            
            For SubLp = StartIndex To Len(MyString)
                If IsNumeric(Mid$(MyString, SubLp, 1)) = False Or SubLp = Len(MyString) Then
                
                    IndexWidth = SubLp - StartIndex
                    
                    'If The Search Is At The End Of The String Then IndexWidth Will
                    'Return 0 When It Should Be One.
                    If IndexWidth = 0 Then IndexWidth = 1: SubLp = SubLp + 1
                    
                    
                    If IsNumeric(Mid$(MyString, StartIndex, (SubLp - StartIndex))) = True Then
                        
                        ImageIndex = CInt(Mid$(MyString, StartIndex, (SubLp - StartIndex)))
                        
                        If ImageCount > 0 And ImageIndex <= ImageCount - 1 Then
                        
                        GetTotalImagesWidth = GetTotalImagesWidth _
                            + MtxImage(ImageIndex).Width
                            
                        End If
                        
                        LP = SubLp - 1: Exit For
                    End If
                End If
            Next SubLp
        End If
    Next LP
    
End Function

Public Function ImageCount() As Integer
On Error GoTo CountError
    
    ImageCount = UBound(MtxImage) + 1
    
Exit Function
CountError:
    ImageCount = 0
End Function

Public Property Get UseFontAntiAliasing() As Boolean
    
    UseFontAntiAliasing = vUseFontAntiAliasing
    
End Property

Public Property Let UseFontAntiAliasing(ByVal vNewValue As Boolean)
    
    vUseFontAntiAliasing = vNewValue
    
    PropertyChanged "UseFontAntiAliasing"
    
    DrawElements
    
End Property

Public Sub ReleaseMemory()

    Erase Alpha
    Erase MatrixString
    Erase MtxImage
    
    DeleteObject myBufferBMP
    DeleteDC myBackBuffer
    DeleteDC myImageDC
    
    Set ImageCollection = Nothing
    Set vImageCollection = Nothing
    
End Sub

Public Property Get StartVisible() As Boolean
    
    StartVisible = vStartVisible
    
End Property

Public Property Let StartVisible(ByVal vNewValue As Boolean)
    
    vStartVisible = vNewValue
    
    PropertyChanged "StartVisible"
    
    If vNewValue = True Then
        vDisplayPosition = MaxLead
    Else
        vDisplayPosition = 0
    End If
    
    PropertyChanged "Displayposition"
    
    DrawElements
    
End Property

Public Property Get FadeLeft() As Boolean
    
    FadeLeft = vFadeLeft
    
End Property

Public Property Let FadeLeft(ByVal vNewValue As Boolean)
    
    vFadeLeft = vNewValue
    
    PropertyChanged "FadeLeft"
    
    DrawElements
    
End Property

Public Property Get FadeRight() As Boolean
    
    FadeRight = vFadeRight
    
End Property

Public Property Let FadeRight(ByVal vNewValue As Boolean)
    
    vFadeRight = vNewValue
    
    PropertyChanged "FadeRight"
    
    DrawElements
    
End Property
