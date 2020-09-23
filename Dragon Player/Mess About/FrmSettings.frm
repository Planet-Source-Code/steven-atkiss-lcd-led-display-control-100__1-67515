VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00D7E8EC&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin Project1.GFXFrame GFXFrame 
      Height          =   1395
      Index           =   3
      Left            =   5100
      TabIndex        =   37
      Top             =   4740
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   2461
      FrameCaption    =   "Secondery Display Options"
      BeginProperty FrameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.GFXFrame GFXFrame 
      Height          =   1575
      Index           =   0
      Left            =   420
      TabIndex        =   36
      Top             =   5220
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2778
      FrameCaption    =   "General Settings"
      BeginProperty FrameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.GFXFrame GFXFrame 
      Height          =   3915
      Index           =   2
      Left            =   420
      TabIndex        =   26
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6906
      FrameCaption    =   "Primary Display Options"
      BeginProperty FrameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox Check14 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Title"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   420
         TabIndex        =   32
         Top             =   540
         Width           =   675
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Artist"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   1380
         TabIndex        =   31
         Top             =   540
         Width           =   735
      End
      Begin VB.CheckBox Check16 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Album"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   2340
         TabIndex        =   30
         Top             =   540
         Width           =   855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Lead With Text"
         ForeColor       =   &H0096540C&
         Height          =   255
         Left            =   660
         TabIndex        =   29
         Top             =   1440
         Width           =   3015
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Lead With Images"
         ForeColor       =   &H0096540C&
         Height          =   255
         Left            =   660
         TabIndex        =   28
         Top             =   1740
         Width           =   3015
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Sync Screen To Auto Volume Shift"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   420
         TabIndex        =   27
         Top             =   3000
         Width           =   3375
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   600
         X2              =   510
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         X1              =   600
         X2              =   510
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00808080&
         X1              =   510
         X2              =   510
         Y1              =   720
         Y2              =   1875
      End
      Begin VB.Label Label2 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Display Opacity"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   420
         TabIndex        =   35
         Top             =   2100
         Width           =   3975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBEDF2&
         Caption         =   "Flash When Paused"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   720
         TabIndex        =   34
         Top             =   2400
         Width           =   1590
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00DBEDF2&
         Caption         =   "Display ""Paused"""
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   720
         TabIndex        =   33
         Top             =   2700
         Width           =   1335
      End
   End
   Begin Project1.GFXFrame GFXFrame 
      Height          =   3915
      Index           =   5
      Left            =   2100
      TabIndex        =   18
      Top             =   5280
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6906
      FrameCaption    =   "Events And Actions"
      BeginProperty FrameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.OptionButton Option4 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Only Remove From Library"
         ForeColor       =   &H0096540C&
         Height          =   255
         Left            =   600
         TabIndex        =   25
         Top             =   1020
         Width           =   3435
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Move File To Temp Delete Folder"
         ForeColor       =   &H0096540C&
         Height          =   255
         Left            =   600
         TabIndex        =   24
         Top             =   720
         Width           =   3435
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Only Add To Current Playlist"
         ForeColor       =   &H0096540C&
         Height          =   255
         Left            =   600
         TabIndex        =   23
         Top             =   1980
         Width           =   2595
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Add To Library And Playlist"
         ForeColor       =   &H0096540C&
         Height          =   255
         Left            =   600
         TabIndex        =   22
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Adding Media To Playlist On Drag And Drop"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1380
         Width           =   3855
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Clearing Unsaved Playlists"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   2340
         Width           =   3615
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Deleting Media From Library"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   420
         Width           =   2595
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         X1              =   540
         X2              =   450
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         X1              =   540
         X2              =   450
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00808080&
         X1              =   450
         X2              =   450
         Y1              =   1620
         Y2              =   2115
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         X1              =   540
         X2              =   450
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00808080&
         X1              =   540
         X2              =   450
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   450
         X2              =   450
         Y1              =   660
         Y2              =   1155
      End
   End
   Begin Project1.GFXFrame GFXFrame 
      Height          =   3855
      Index           =   1
      Left            =   6120
      TabIndex        =   12
      Top             =   4080
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   6800
      FrameCaption    =   "Adding Media"
      BeginProperty FrameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox Check1 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Expand Sub Folders"
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   2115
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Get Media Information "
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   360
         TabIndex        =   15
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Replace Deleted Files"
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   360
         TabIndex        =   14
         Top             =   2220
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Allow Duplicate Files."
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Reference File Name, Media Info, Both"
         Top             =   2520
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Reduces Search Speed However Allows Faster Library Intergration As The Media Information Will Be Readily Available."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   600
         TabIndex        =   17
         Top             =   1380
         Width           =   3375
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   360
         X2              =   3900
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   360
         X2              =   3900
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin Project1.GFXFrame GFXFrame 
      Height          =   3615
      Index           =   4
      Left            =   6240
      TabIndex        =   5
      Top             =   420
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   6376
      FrameCaption    =   "Automatic Volume Shift"
      BeginProperty FrameFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CheckBox Check13 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Fade In On Play After Pause"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   600
         TabIndex        =   11
         Top             =   1680
         Width           =   2955
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Fade In On Exit"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   600
         TabIndex        =   10
         Top             =   2580
         Width           =   2595
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Fade Out On Stop"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   600
         TabIndex        =   9
         Top             =   2280
         Width           =   2595
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Fade Out On Pause"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   600
         TabIndex        =   8
         Top             =   1980
         Width           =   2595
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Fade In On Play After Stop"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   600
         TabIndex        =   7
         Top             =   1380
         Width           =   2595
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00DBEDF2&
         Caption         =   "Cross Fade On Track Iteration"
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   600
         TabIndex        =   6
         Top             =   1080
         Width           =   2835
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3900
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   6879
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Adding Media"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Primary Display"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Secondery Display"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto Volume Shift"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Events And Actions"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "!! Full Screen Scroll Of Ind Segments"
      Height          =   255
      Left            =   420
      TabIndex        =   3
      Top             =   6480
      Width           =   3675
   End
   Begin VB.Label Label7 
      Caption         =   "!! Scroll All Selected Segments"
      Height          =   255
      Left            =   420
      TabIndex        =   2
      Top             =   6120
      Width           =   3675
   End
   Begin VB.Label Label6 
      Caption         =   "!! Fade In And Scroll Ind Segments"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   5760
      Width           =   3675
   End
   Begin VB.Label Label5 
      Caption         =   "!! When Info Not Available Display UK or Hide"
      Height          =   195
      Left            =   540
      TabIndex        =   0
      Top             =   5460
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
TabStrip1.Tabs(3).Selected = True
End Sub

Private Sub TabStrip1_Click()
    
    GFXFrame(TabStrip1.SelectedItem.Index - 1).ZOrder (0)
    
End Sub
