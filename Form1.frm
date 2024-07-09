VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   12150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20955
   BeginProperty Font 
      Name            =   "Sylfaen"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12150
   ScaleWidth      =   20955
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   17880
      TabIndex        =   47
      Top             =   11160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1191
      _Version        =   393216
      LargeChange     =   10
      Min             =   10
      Max             =   500
      SelectRange     =   -1  'True
      SelStart        =   100
      Value           =   100
   End
   Begin MSComDlg.CommonDialog cdb 
      Left            =   1560
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   8055
      Left            =   5040
      TabIndex        =   46
      Top             =   3360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   14208
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20775
      _ExtentX        =   36645
      _ExtentY        =   5741
      _Version        =   393216
      MousePointer    =   4
      Tab             =   1
      TabHeight       =   1058
      TabMaxWidth     =   3528
      BackColor       =   12615680
      TabCaption(0)   =   "&File"
      TabPicture(0)   =   "Form1.frx":007C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "&Home"
      TabPicture(1)   =   "Form1.frx":0098
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "&View"
      TabPicture(2)   =   "Form1.frx":00B4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame8 
         Caption         =   "Settings"
         Height          =   2175
         Left            =   -67530
         TabIndex        =   43
         Top             =   840
         Width           =   3375
         Begin VB.ComboBox Combo10 
            Height          =   495
            Left            =   240
            TabIndex        =   45
            Text            =   "Measurement Units"
            Top             =   1440
            Width           =   3015
         End
         Begin VB.ComboBox Combo9 
            Height          =   495
            Left            =   240
            TabIndex        =   44
            Text            =   "WordWrap"
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Zoom"
         Height          =   2175
         Left            =   -74880
         TabIndex        =   37
         Top             =   840
         Width           =   4815
         Begin VB.CommandButton Command27 
            Caption         =   "ZoomIn"
            Height          =   1215
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command26 
            Caption         =   "ZoomOut"
            Height          =   1215
            Left            =   1680
            TabIndex        =   39
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command25 
            Caption         =   "100%"
            Height          =   1215
            Left            =   3240
            TabIndex        =   38
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Show or Hide"
         Height          =   2175
         Left            =   -70065
         TabIndex        =   36
         Top             =   840
         Width           =   2535
         Begin VB.CheckBox Check2 
            Caption         =   "Status bar"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Ruler"
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   840
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Editing"
         Height          =   2175
         Left            =   17520
         TabIndex        =   32
         Top             =   840
         Width           =   2535
         Begin VB.CommandButton Command22 
            Caption         =   "Select All"
            Height          =   495
            Left            =   120
            TabIndex        =   35
            Top             =   1680
            Width           =   1575
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Replace"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Find"
            Height          =   615
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Insert"
         Height          =   2175
         Left            =   12600
         TabIndex        =   28
         Top             =   840
         Width           =   4815
         Begin VB.CommandButton Command19 
            Caption         =   "Object"
            Height          =   1215
            Left            =   3240
            TabIndex        =   31
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Date -Time"
            Height          =   1215
            Left            =   1500
            TabIndex        =   30
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox Combo8 
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Text            =   "Picture"
            Top             =   1440
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Paragraph"
         Height          =   2175
         Left            =   9000
         TabIndex        =   18
         Top             =   840
         Width           =   3495
         Begin VB.CommandButton Command18 
            Caption         =   "L"
            Height          =   495
            Left            =   240
            TabIndex        =   27
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton Command16 
            Caption         =   "P"
            Height          =   495
            Left            =   2640
            TabIndex        =   26
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton Command15 
            Caption         =   "J"
            Height          =   495
            Left            =   1920
            TabIndex        =   25
            Top             =   1380
            Width           =   600
         End
         Begin VB.CommandButton Command14 
            Caption         =   "R"
            Height          =   495
            Left            =   1365
            TabIndex        =   24
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton Command13 
            Caption         =   "E"
            Height          =   495
            Left            =   810
            TabIndex        =   23
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton Command12 
            Caption         =   "dI"
            Height          =   495
            Left            =   1080
            TabIndex        =   22
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton Command11 
            Caption         =   "iI"
            Height          =   495
            Left            =   360
            TabIndex        =   21
            Top             =   600
            Width           =   615
         End
         Begin VB.ComboBox Combo7 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   1800
            TabIndex        =   20
            Text            =   "A"
            Top             =   600
            Width           =   615
         End
         Begin VB.ComboBox Combo6 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   2520
            TabIndex        =   19
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Font"
         Height          =   2175
         Left            =   2760
         TabIndex        =   5
         Top             =   840
         Width           =   6135
         Begin VB.ComboBox Combo5 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   4680
            TabIndex        =   17
            Text            =   "PEN"
            Top             =   1380
            Width           =   1095
         End
         Begin VB.ComboBox Combo4 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3870
            TabIndex        =   16
            Text            =   "A"
            Top             =   1380
            Width           =   735
         End
         Begin VB.CommandButton Command10 
            Caption         =   "A^"
            Height          =   615
            Left            =   4120
            TabIndex        =   15
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command9 
            Caption         =   "A,"
            Height          =   615
            Left            =   5040
            TabIndex        =   14
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton cmdItalic 
            Caption         =   "I"
            Height          =   495
            Left            =   810
            TabIndex        =   13
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton cmdUnderline 
            Caption         =   "U"
            Height          =   495
            Left            =   1365
            TabIndex        =   12
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton cmdStrikethrough 
            Caption         =   "abc"
            Height          =   495
            Left            =   1935
            TabIndex        =   11
            Top             =   1380
            Width           =   735
         End
         Begin VB.CommandButton Command5 
            Caption         =   "X,,"
            Height          =   495
            Left            =   2745
            TabIndex        =   10
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton Command4 
            Caption         =   "X''"
            Height          =   495
            Left            =   3300
            TabIndex        =   9
            Top             =   1380
            Width           =   495
         End
         Begin VB.CommandButton cmdBold 
            Caption         =   "B"
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   1380
            Width           =   495
         End
         Begin VB.ComboBox comboFontSize 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3080
            TabIndex        =   7
            Text            =   "11"
            Top             =   540
            Width           =   855
         End
         Begin VB.ComboBox comboFontName 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Text            =   "Calibri"
            Top             =   540
            Width           =   2775
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Clipboard"
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2535
         Begin VB.ComboBox Combo1 
            Height          =   495
            Left            =   120
            TabIndex        =   4
            Text            =   "Paste"
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cut"
            Height          =   615
            Left            =   1560
            TabIndex        =   3
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Copy"
            Height          =   495
            Left            =   1560
            TabIndex        =   2
            Top             =   1200
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBold_Click()
RichTextBox1.SelBold = True
End Sub

Private Sub cmdItalic_Click()
RichTextBox1.SelItalic = True
End Sub

Private Sub cmdStrikethrough_Click()
RichTextBox1.SelStrikeThru = True
End Sub

Private Sub cmdUnderline_Click()
RichTextBox1.SelUnderline = True
End Sub


Private Sub comboFontName_Change()
RichTextBox1.SelFontName = comboFontName.Text
End Sub

Private Sub comboFontSize_Change()
RichTextBox1.SelFontSize = comboFontSize.Text

End Sub

Private Sub Command1_Click()
RichTextBox1.se
End Sub
