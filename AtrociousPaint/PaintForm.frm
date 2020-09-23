VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PaintForm 
   ClientHeight    =   6270
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7080
   Icon            =   "PaintForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar HScroll 
      Height          =   255
      LargeChange     =   100
      Left            =   840
      Max             =   0
      SmallChange     =   10
      TabIndex        =   52
      Top             =   6720
      Visible         =   0   'False
      Width           =   10815
   End
   Begin VB.Frame ColorFrame 
      Height          =   855
      Left            =   0
      TabIndex        =   22
      Top             =   6960
      Width           =   12015
      Begin VB.Timer StatusBarTimer 
         Interval        =   3000
         Left            =   5040
         Top             =   240
      End
      Begin MSComDlg.CommonDialog ColorDialog 
         Left            =   4440
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label ForeColorLabel 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   51
         ToolTipText     =   "Fore Color"
         Top             =   360
         Width           =   255
      End
      Begin VB.Label FillColorLabel 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   360
         TabIndex        =   50
         ToolTipText     =   "Fill Color"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   850
         TabIndex        =   49
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   850
         TabIndex        =   48
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label ColorLabel 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   555
         Left            =   150
         TabIndex        =   47
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Color 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1125
         TabIndex        =   46
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1125
         TabIndex        =   45
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   1400
         TabIndex        =   44
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   1400
         TabIndex        =   43
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00008080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   1660
         TabIndex        =   42
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   1660
         TabIndex        =   41
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   8
         Left            =   1935
         TabIndex        =   40
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   39
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   10
         Left            =   2200
         TabIndex        =   38
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   11
         Left            =   2200
         TabIndex        =   37
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   12
         Left            =   2475
         TabIndex        =   36
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   2475
         TabIndex        =   35
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00800080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   14
         Left            =   2745
         TabIndex        =   34
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   15
         Left            =   2745
         TabIndex        =   33
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00004040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   16
         Left            =   3015
         TabIndex        =   32
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   17
         Left            =   3015
         TabIndex        =   31
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   18
         Left            =   3285
         TabIndex        =   30
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   19
         Left            =   3285
         TabIndex        =   29
         ToolTipText     =   "Color"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   20
         Left            =   3555
         TabIndex        =   28
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   21
         Left            =   3555
         TabIndex        =   27
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00400040&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   22
         Left            =   3825
         TabIndex        =   26
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   23
         Left            =   3825
         TabIndex        =   25
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   24
         Left            =   4080
         TabIndex        =   24
         ToolTipText     =   "Color"
         Top             =   225
         Width           =   255
      End
      Begin VB.Label Color 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   25
         Left            =   4080
         TabIndex        =   23
         ToolTipText     =   "Color"
         Top             =   495
         Width           =   255
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   70
      Top             =   7800
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8026
            MinWidth        =   8026
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Process"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Object.ToolTipText     =   "Mouse Point"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5733
            MinWidth        =   5733
            Text            =   "VOTE ME & SEND ME COMMENTS"
            TextSave        =   "VOTE ME & SEND ME COMMENTS"
            Object.ToolTipText     =   "atrocious@hotmail.com"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   72
      Top             =   3360
      Width           =   855
      Begin VB.Frame FillFrame 
         Height          =   1095
         Left            =   75
         TabIndex        =   79
         ToolTipText     =   "FIll Type"
         Top             =   120
         Visible         =   0   'False
         WhatsThisHelpID =   10333
         Width           =   705
         Begin VB.Label FillLabel 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   60
            TabIndex        =   80
            Top             =   120
            WhatsThisHelpID =   10334
            Width           =   570
         End
         Begin VB.Shape RectangleShape 
            BackColor       =   &H00FFFFFF&
            BorderColor     =   &H00FFFFFF&
            Height          =   150
            Index           =   0
            Left            =   140
            Top             =   210
            Width           =   420
         End
         Begin VB.Shape RectangleShape 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   150
            Index           =   1
            Left            =   120
            Top             =   525
            Width           =   420
         End
         Begin VB.Shape RectangleShape 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   150
            Index           =   2
            Left            =   140
            Top             =   840
            Width           =   420
         End
      End
      Begin VB.Frame SelectionFrame 
         Height          =   1000
         Left            =   75
         TabIndex        =   75
         Top             =   120
         Visible         =   0   'False
         Width           =   735
         Begin VB.PictureBox SelectionType 
            BorderStyle     =   0  'None
            Height          =   350
            Index           =   1
            Left            =   90
            Picture         =   "PaintForm.frx":2372
            ScaleHeight     =   23
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   36
            TabIndex        =   77
            ToolTipText     =   "Cut Image"
            Top             =   575
            Width           =   535
         End
         Begin VB.PictureBox SelectionType 
            BorderStyle     =   0  'None
            Height          =   350
            Index           =   0
            Left            =   90
            Picture         =   "PaintForm.frx":2DC4
            ScaleHeight     =   23
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   36
            TabIndex        =   76
            ToolTipText     =   "Copy Image"
            Top             =   145
            Width           =   535
         End
         Begin VB.Label SelectionLabel 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   415
            Left            =   50
            TabIndex        =   78
            Top             =   115
            Width           =   635
         End
      End
      Begin VB.Frame BrushFrame 
         Height          =   1545
         Left            =   85
         TabIndex        =   73
         ToolTipText     =   "Brush Type"
         Top             =   120
         Visible         =   0   'False
         WhatsThisHelpID =   10335
         Width           =   675
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   4
            Left            =   120
            Picture         =   "PaintForm.frx":3816
            Top             =   750
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   5
            Left            =   360
            Picture         =   "PaintForm.frx":3859
            Top             =   750
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   7
            Left            =   360
            Picture         =   "PaintForm.frx":389D
            Top             =   1020
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   6
            Left            =   120
            Picture         =   "PaintForm.frx":38DF
            Top             =   1020
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   2
            Left            =   120
            Picture         =   "PaintForm.frx":3921
            Top             =   480
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   3
            Left            =   360
            Picture         =   "PaintForm.frx":3966
            Top             =   480
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   0
            Left            =   120
            Picture         =   "PaintForm.frx":39AA
            Top             =   240
            Width           =   135
         End
         Begin VB.Label BrushLabel 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   65
            TabIndex        =   74
            Top             =   165
            WhatsThisHelpID =   10336
            Width           =   255
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   1
            Left            =   360
            Picture         =   "PaintForm.frx":39EE
            Top             =   240
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   8
            Left            =   120
            Picture         =   "PaintForm.frx":3A32
            Top             =   1290
            Width           =   135
         End
         Begin VB.Image Brush 
            Appearance      =   0  'Flat
            Height          =   135
            Index           =   9
            Left            =   360
            Picture         =   "PaintForm.frx":3A71
            Top             =   1290
            Width           =   135
         End
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   360
      Left            =   0
      TabIndex        =   69
      Top             =   15
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame ToolsFrame 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   240
      WhatsThisHelpID =   10296
      Width           =   855
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   7
         Left            =   435
         Picture         =   "PaintForm.frx":3AB3
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Brush"
         Top             =   1245
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   1
         Left            =   435
         Picture         =   "PaintForm.frx":3C17
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Zoom"
         Top             =   120
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   15
         Left            =   435
         Picture         =   "PaintForm.frx":3FA5
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Curve"
         Top             =   2745
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   13
         Left            =   435
         Picture         =   "PaintForm.frx":3FFD
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Polygon"
         Top             =   2370
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   11
         Left            =   435
         Picture         =   "PaintForm.frx":4070
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Rounded Rectangle"
         Top             =   1995
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   6
         Left            =   50
         Picture         =   "PaintForm.frx":40FA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Air Brush"
         Top             =   1245
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   5
         Left            =   435
         Picture         =   "PaintForm.frx":4404
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eraser"
         Top             =   870
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10295
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   4
         Left            =   50
         Picture         =   "PaintForm.frx":4483
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Pencil"
         Top             =   870
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         WhatsThisHelpID =   10298
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   8
         Left            =   50
         Picture         =   "PaintForm.frx":4502
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Line"
         Top             =   1620
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10299
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   3
         Left            =   435
         Picture         =   "PaintForm.frx":45E7
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Fill"
         Top             =   495
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10300
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   12
         Left            =   50
         Picture         =   "PaintForm.frx":4669
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ellipse"
         Top             =   2370
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10301
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   10
         Left            =   50
         Picture         =   "PaintForm.frx":46D6
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Rectangle"
         Top             =   1995
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   10302
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   14
         Left            =   50
         Picture         =   "PaintForm.frx":4743
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Text"
         Top             =   2745
         WhatsThisHelpID =   10338
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   9
         Left            =   435
         Picture         =   "PaintForm.frx":4AC5
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Arrow"
         Top             =   1620
         WhatsThisHelpID =   10340
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   0
         Left            =   50
         Picture         =   "PaintForm.frx":4B10
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Select Area"
         Top             =   120
         WhatsThisHelpID =   10359
         Width           =   390
      End
      Begin VB.OptionButton Tools 
         Height          =   375
         Index           =   2
         Left            =   50
         Picture         =   "PaintForm.frx":4E8E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Pick Color"
         Top             =   480
         WhatsThisHelpID =   10361
         Width           =   390
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   6375
      LargeChange     =   1000
      Left            =   11640
      Max             =   0
      SmallChange     =   100
      TabIndex        =   54
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComDlg.CommonDialog SaveDialog 
      Left            =   3840
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame ScrollFrame 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   11640
      TabIndex        =   53
      Top             =   6600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame PropertyFrame 
      Height          =   495
      Left            =   0
      TabIndex        =   55
      Top             =   -120
      Width           =   12015
      Begin VB.ComboBox DWCombo 
         Height          =   315
         Left            =   7440
         TabIndex        =   62
         Top             =   145
         Width           =   1455
      End
      Begin VB.ComboBox BStyle 
         Height          =   315
         Left            =   1920
         TabIndex        =   57
         Top             =   145
         Width           =   1455
      End
      Begin VB.ComboBox FStyle 
         Height          =   315
         Left            =   4560
         TabIndex        =   56
         Top             =   145
         Width           =   1455
      End
      Begin VB.Label DWLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Draw Width"
         Height          =   255
         Left            =   6360
         TabIndex        =   63
         Top             =   210
         Width           =   975
      End
      Begin VB.Label BSLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Border Style"
         Height          =   255
         Left            =   840
         TabIndex        =   59
         Top             =   210
         Width           =   975
      End
      Begin VB.Label FSLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fill Style"
         Height          =   255
         Left            =   3600
         TabIndex        =   58
         Top             =   210
         Width           =   975
      End
   End
   Begin VB.PictureBox VResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Left            =   3600
      MousePointer    =   7  'Size N S
      ScaleHeight     =   3
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   3
      TabIndex        =   17
      Top             =   6285
      Width           =   75
   End
   Begin VB.PictureBox BResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Left            =   11280
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   19
      Top             =   6240
      Width           =   75
   End
   Begin VB.PictureBox HResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Left            =   11280
      MousePointer    =   9  'Size W E
      ScaleHeight     =   45
      ScaleWidth      =   45
      TabIndex        =   18
      Top             =   3240
      Width           =   75
   End
   Begin VB.PictureBox PaintBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5925
      Left            =   840
      ScaleHeight     =   391
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   691
      TabIndex        =   20
      Top             =   360
      Width           =   10425
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   4200
         Top             =   720
      End
      Begin VB.PictureBox FilterPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3600
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   71
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComctlLib.ImageList CursorsList 
         Left            =   4440
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":5216
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":5530
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":584A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":5B64
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":5E7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":6198
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":6A72
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":6BD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":6D36
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":7050
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":792A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PaintForm.frx":8204
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TextBox 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4080
         TabIndex        =   67
         Top             =   120
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.PictureBox BufferPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Index           =   0
         Left            =   3000
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   66
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox ClipboardPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2400
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   65
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox PicRotate 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1800
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   64
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSComDlg.CommonDialog OpenDialog 
         Left            =   3480
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog PrintDialog 
         Left            =   2400
         Top             =   105
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox ImgZoom1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1200
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox ImgZoom2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1800
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   29
         TabIndex        =   60
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.PictureBox SelectionBox 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1200
         ScaleHeight     =   33
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   33
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label TextLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   68
         Top             =   120
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Shape ImageSelector 
         BorderStyle     =   3  'Dot
         Height          =   495
         Left            =   600
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape EllipseShape 
         Height          =   495
         Left            =   0
         Shape           =   2  'Oval
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape RRectShape 
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape RectShape 
         Height          =   495
         Left            =   600
         Top             =   120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Line LineShape 
         Visible         =   0   'False
         X1              =   0
         X2              =   40
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Menu FileMnu 
      Caption         =   "&File"
      Begin VB.Menu FileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu FileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu FileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu FileSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu FileDash1 
         Caption         =   "-"
      End
      Begin VB.Menu FilePageSetup 
         Caption         =   "P&age Setup"
      End
      Begin VB.Menu FilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu FileDash2 
         Caption         =   "-"
      End
      Begin VB.Menu FileSAW 
         Caption         =   "Set As Wallpaper"
      End
      Begin VB.Menu FileDash3 
         Caption         =   "-"
      End
      Begin VB.Menu FileClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu EditMnu 
      Caption         =   "&Edit"
      Begin VB.Menu EditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu EditRedo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
      Begin VB.Menu EditDash1 
         Caption         =   "-"
      End
      Begin VB.Menu EditCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu EditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu EditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu EditDel 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu EditDash2 
         Caption         =   "-"
      End
      Begin VB.Menu EditSelect 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu EditCrop 
         Caption         =   "C&rop"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuEffect 
      Caption         =   "&Effect"
      Begin VB.Menu EffectFlip 
         Caption         =   "&Flip"
         Begin VB.Menu FlipHorizontal 
            Caption         =   "&Horizontal"
         End
         Begin VB.Menu FlipVertical 
            Caption         =   "&Vertical"
         End
         Begin VB.Menu FlipBoth 
            Caption         =   "&Both"
         End
      End
      Begin VB.Menu EffectRotate 
         Caption         =   "&Rotate"
         Begin VB.Menu Rotate45 
            Caption         =   "45°"
         End
         Begin VB.Menu Rotate90 
            Caption         =   "90°"
         End
         Begin VB.Menu Rotate135 
            Caption         =   "135°"
         End
         Begin VB.Menu Rotate180 
            Caption         =   "180°"
         End
         Begin VB.Menu Rotate225 
            Caption         =   "225°"
         End
         Begin VB.Menu Rotate270 
            Caption         =   "270°"
         End
         Begin VB.Menu Rotate315 
            Caption         =   "315"
         End
         Begin VB.Menu RotateDash1 
            Caption         =   "-"
         End
         Begin VB.Menu RotateClockwise 
            Caption         =   "&Clockwise"
            Checked         =   -1  'True
         End
         Begin VB.Menu RotateAClockwise 
            Caption         =   "&Anti-Clockwise"
         End
      End
      Begin VB.Menu EffectDash1 
         Caption         =   "-"
      End
      Begin VB.Menu EffectClear 
         Caption         =   "&Clear"
      End
   End
   Begin VB.Menu FilterMnu 
      Caption         =   "Filte&r"
      Begin VB.Menu FilterBW 
         Caption         =   "&Black && White"
      End
      Begin VB.Menu FilterBlur 
         Caption         =   "B&lur"
      End
      Begin VB.Menu FilterBright 
         Caption         =   "B&rightness"
      End
      Begin VB.Menu FilterCrease 
         Caption         =   "&Crease"
      End
      Begin VB.Menu FilterDark 
         Caption         =   "&Darkness"
      End
      Begin VB.Menu FilterDiffuse 
         Caption         =   "Di&ffuse"
      End
      Begin VB.Menu FilterEmboss 
         Caption         =   "&Emboss"
      End
      Begin VB.Menu FilterGBW 
         Caption         =   "Gra&y Black White"
      End
      Begin VB.Menu FilterGray 
         Caption         =   "&GrayScale"
      End
      Begin VB.Menu FilterInvertColor 
         Caption         =   "&InvertColors"
      End
      Begin VB.Menu FilterRepColor 
         Caption         =   "&Replace Colors"
      End
      Begin VB.Menu FilterSharp 
         Caption         =   "&Sharpen"
      End
      Begin VB.Menu FilterSolarize 
         Caption         =   "S&olarize"
      End
      Begin VB.Menu FilterSnow 
         Caption         =   "S&now"
      End
      Begin VB.Menu FilterWave 
         Caption         =   "&Wave"
      End
   End
   Begin VB.Menu HelpMnu 
      Caption         =   "&Help"
      Begin VB.Menu HelpTopics 
         Caption         =   "Help &Topics"
      End
      Begin VB.Menu HelpDash1 
         Caption         =   "-"
      End
      Begin VB.Menu PaintAbout 
         Caption         =   "&About Paint"
      End
   End
End
Attribute VB_Name = "PaintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FName
Dim SC As Boolean
Dim SI
Dim PicChanged As Boolean
Dim SMsg
Dim SErr As Boolean
Dim UnLod As Boolean
Dim FImgZoom As Boolean
Dim ImgMovX, ImgMovY
Dim rx, ry
Dim HMD, VMD, BMD As Boolean
Dim CurInd As Integer
Dim stbl As Boolean
Dim OldX, OldY
Dim mode As Boolean
Dim DPolygon As Boolean
Dim ImgMov As Boolean
Public CurBuf As Integer
Public EndBuf As Integer
Public StartBuf As Integer
Public CurrentBezierPoint As Integer
Dim PolygonLen() As PointApi
Public FFillStyle As FillStyleConstants

Public EBorderStyle As BorderStyleConstants

Public BrushShape As EBrushShape
Public Enum EBrushShape
  
FilledRect = 0
FilledCircle = 1
SRect = 2
SCircle = 3
Cross = 4
DiagonalCross = 5
UpwardDiagonal = 6
DownwardDiagonal = 7
Horizontal = 8
Vertical = 9

End Enum

Public SFillStyle As EFillStyle
Public Enum EFillStyle

BorderOnly = 0
BorderFill = 1
FillOnly = 2

End Enum

Public PSelectType As SelectType
Public Enum SelectType

TransparentSelect = 0
FilledSelect = 1

End Enum

Private Sub Brush_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

BrushLabel.Top = Brush(Index).Top - (4 * Screen.TwipsPerPixelX)
BrushLabel.Left = Brush(Index).Left - (4 * Screen.TwipsPerPixelY)
BrushShape = Index

End Sub

Private Sub BStyle_Click()

Select Case BStyle.Text
Case "Solid"
    EBorderStyle = 0
Case "Dash"
    EBorderStyle = 1
Case "Dot"
    EBorderStyle = 2
Case "Dash-Dot"
    EBorderStyle = 3
Case "Dash-Dot-Dot"
    EBorderStyle = 4
End Select

End Sub

Private Sub Color_DblClick(Index As Integer)

ColorDialog.Color = PaintForm.ForeColorLabel.BackColor
ColorDialog.Flags = &H1
ColorDialog.ShowColor
Color(Index).BackColor = ColorDialog.Color
PaintForm.ForeColorLabel.BackColor = Color(Index).BackColor

End Sub

Private Sub Color_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    ForeColorLabel.BackColor = Color(Index).BackColor
    PaintBox.ForeColor = ForeColorLabel.BackColor
ElseIf Button = vbRightButton Then
    FillColorLabel.BackColor = Color(Index).BackColor
    PaintBox.FillColor = FillColorLabel.BackColor
End If

End Sub

Private Sub DWCombo_Click()

PaintBox.DrawWidth = DWCombo.Text

End Sub

Private Sub EditCopy_Click()

ClipboardPic.Picture = SelectionBox.Image
Clipboard.SetData ClipboardPic.Image, vbCFBitmap

End Sub

Private Sub EditCrop_Click()

Dim PH, PW

PH = SelectionBox.Height
PW = SelectionBox.Width
PaintBox.Height = PH
PaintBox.Width = PW
PaintBox.Picture = SelectionBox.Image
ImageSelector.Top = 0
ImageSelector.Left = 0
ImageSelector.Height = PaintBox.ScaleHeight
ImageSelector.Width = PaintBox.ScaleWidth
SelectionBox.Width = ImageSelector.Width - 2
SelectionBox.Height = ImageSelector.Height - 2
SelectionBox.Left = ImageSelector.Left + 1
SelectionBox.Top = ImageSelector.Top + 1
SelectionBox.PaintPicture PaintBox.Image, 0, 0, SelectionBox.Width, SelectionBox.Height, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
PaintBox.PaintPicture SelectionBox.Image, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
SelectionBox = Nothing
SelectionBox.Visible = False
ImageSelector.Visible = False
EditCut.Enabled = False
EditCopy.Enabled = False
EditDel.Enabled = False
EditCrop.Enabled = False
ImgMov = False
PicChanged = True

Form_Resize
PlaceResizeBox
SetImageBuffer

End Sub

Private Sub EditCut_Click()

EditCopy_Click
EditDel_Click

End Sub

Private Sub EditDel_Click()

PaintBox.Line (SelectionBox.Left, SelectionBox.Top)-(SelectionBox.Left + SelectionBox.Width, SelectionBox.Top + SelectionBox.Height), vbWhite, BF
SelectionBox = Nothing
SelectionBox.Visible = False
ImageSelector.Visible = False
EditCut.Enabled = False
EditCopy.Enabled = False
EditCrop.Enabled = False
EditDel.Enabled = False
PicChanged = True

SetImageBuffer

End Sub

Private Sub EditPaste_Click()

If SelectionBox.Visible = True Then
    ClipboardPic.Picture = SelectionBox.Image
End If

Tools(0).Value = True
Tools_Click (0)
ImageSelector.Visible = True
ImageSelector.Top = 0
ImageSelector.Left = 0
ImageSelector.Height = ClipboardPic.Height + 2
ImageSelector.Width = ClipboardPic.Width + 2
SelectionBox.Width = ImageSelector.Width - 2
SelectionBox.Height = ImageSelector.Height - 2
SelectionBox.Left = ImageSelector.Left + 1
SelectionBox.Top = ImageSelector.Top + 1
SelectionBox.Visible = True
ClipboardPic.Picture = Clipboard.GetData(vbCFBitmap)
SelectionBox.Picture = ClipboardPic.Image
ImgMov = True
EditCut.Enabled = True
EditCopy.Enabled = True
EditDel.Enabled = True
EditCrop.Enabled = True
PicChanged = True

SetImageBuffer

End Sub

Private Sub EditRedo_Click()

On Error Resume Next

If DPolygon Then
    DrawPolygon Complete:=False
    DrawPolygon
    DPolygon = False
End If

If FImgZoom = True Then
    PaintBox.ScaleMode = 1
    FImgZoom = False
    PaintBox.Height = ImgZoom2.Height
    PaintBox.Width = ImgZoom2.Width
    PaintBox.Picture = ImgZoom2.Image
    PaintBox.ScaleMode = 3
    Form_Resize
    PlaceResizeBox
End If

If SelectionBox.Visible Then
    PaintBox.PaintPicture SelectionBox.Image, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
    EditCut.Enabled = False
    EditCopy.Enabled = False
    EditDel.Enabled = False
    EditCrop.Enabled = False
    SelectionBox.Visible = False
    SelectionBox.Cls
    ImageSelector.Visible = False
    PicChanged = True
    SetImageBuffer
End If

If CurBuf < MaxBuf Then
    CurBuf = CurBuf + 1
Else
    CurBuf = 0
End If

PaintBox.Picture = BufferPic(CurBuf).Image
PaintBox.Width = CLng(Left(BufferPic(CurBuf).Tag, Len(BufferPic(CurBuf).Tag) - 5))
PaintBox.Height = CLng(Right(BufferPic(CurBuf).Tag, 5))
EditUndo.Enabled = True

If CurBuf = EndBuf Then
    EditRedo.Enabled = False
End If

PaintBox_DblClick
PlaceResizeBox
Form_Resize

End Sub

Private Sub EditSelect_Click()

Tools(0).Value = True
Tools_Click (0)
ImageSelector.Visible = True
ImageSelector.Top = 0
ImageSelector.Left = 0
ImageSelector.Height = PaintBox.ScaleHeight
ImageSelector.Width = PaintBox.ScaleWidth
SelectionBox.Width = ImageSelector.Width - 2
SelectionBox.Height = ImageSelector.Height - 2
SelectionBox.Left = ImageSelector.Left + 1
SelectionBox.Top = ImageSelector.Top + 1
SelectionBox.Visible = True
EditCut.Enabled = True
EditCopy.Enabled = True
EditDel.Enabled = True
EditCrop.Enabled = True
SelectionBox.Picture = Nothing
SelectionBox.PaintPicture PaintBox.Image, 0, 0, SelectionBox.Width, SelectionBox.Height, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
PaintBox.Line (SelectionBox.Left, SelectionBox.Top)-(SelectionBox.Left + SelectionBox.Width, SelectionBox.Top + SelectionBox.Height), vbWhite, BF
ImgMov = True

End Sub

Private Sub EditUndo_Click()

On Error Resume Next

If DPolygon Then
    DrawPolygon Complete:=False
    DrawPolygon
    DPolygon = False
End If

If FImgZoom = True Then
    PaintBox.ScaleMode = 1
    FImgZoom = False
    PaintBox.Height = ImgZoom2.Height
    PaintBox.Width = ImgZoom2.Width
    PaintBox.Picture = ImgZoom2.Image
    PaintBox.ScaleMode = 3
    Form_Resize
    PlaceResizeBox
End If

If SelectionBox.Visible Then
    PaintBox.PaintPicture SelectionBox.Image, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
    EditCut.Enabled = False
    EditCopy.Enabled = False
    EditDel.Enabled = False
    EditCrop.Enabled = False
    SelectionBox.Visible = False
    SelectionBox.Cls
    ImageSelector.Visible = False
    PicChanged = True
    SetImageBuffer
End If

If CurBuf > 0 Then
    CurBuf = CurBuf - 1
Else
    CurBuf = MaxBuf
End If

PaintBox.Picture = BufferPic(CurBuf).Image
PaintBox.Width = CLng(Left(BufferPic(CurBuf).Tag, Len(BufferPic(CurBuf).Tag) - 5))
PaintBox.Height = CLng(Right(BufferPic(CurBuf).Tag, 5))

If CurBuf = StartBuf Then
    EditUndo.Enabled = False
End If

EditRedo.Enabled = True

PlaceResizeBox
Form_Resize
 
End Sub
  
Private Sub EffectClear_Click()

PaintBox.Picture = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub FileNew_Click()

On Error GoTo ErrHandler

If PicChanged = True Then
    SaveChanges
    Select Case SMsg
    Case vbYes
        If SC = False Then
             Call FileSaveAs_Click
        ElseIf SC = True Then
            SavePicture PaintBox.Image, OpenDialog.FileName
            PicChanged = False
        End If
        PaintBox.Picture = Nothing
        FName = "Ahsan.abmp"
        PaintForm.Caption = FName & " - AtrociousPaint"
        PicChanged = False
    Case vbNo
        PaintBox.Picture = Nothing
        FName = "Ahsan.abmp"
        PaintForm.Caption = FName & " - AtrociousPaint"
        PicChanged = False
    Case Else
        Exit Sub
    End Select
Else
    PaintBox.Picture = Nothing
    PaintForm.Caption = FName & " - AtrociousPaint"
End If

PaintForm.Tag = ""
SC = False
PaintBox.Width = 10425
PaintBox.Height = 5925

PlaceResizeBox
Form_Resize
ClearImageBuffer
Exit Sub

ErrHandler:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox "Unknown error occured.", vbCritical, "AtrociousPaint"
    End If

End Sub

Private Sub FilePageSetup_Click()

On Error GoTo PErr

PrintDialog.Flags = cdlPDPrintSetup
PrintDialog.ShowPrinter

Exit Sub

PErr:
MsgBox "Cannot print the file." & vbNewLine & vbNewLine & "Make sure the print is ready.", vbOKOnly + vbCritical, "AtrociousPaint"

End Sub

Private Sub FilePrint_Click()

On Error GoTo PErr

Printer.Print PaintBox.Image

Exit Sub

PErr:
MsgBox "Cannot print the file." & vbNewLine & vbNewLine & "Make sure the print is ready.", vbOKOnly + vbCritical, "AtrociousPaint"

End Sub

Private Sub FileSave_Click()

On Error Resume Next

If SC = False Then
    Call FileSaveAs_Click
ElseIf SC = True Then
    SavePicture PaintBox.Image, OpenDialog.FileName
    PicChanged = False
End If

End Sub

Private Sub FileSaveAs_Click()

SaveDialog.CancelError = True

On Error GoTo ErrHandler
    SaveDialog.Filter = "Ahsan's Picture Files|*.abmp| All Files|*.*"
    SaveDialog.FilterIndex = 12
    SaveDialog.ShowSave
    FName = SaveDialog.FileTitle
    OpenDialog.FileName = SaveDialog.FileName
    SavePicture PaintBox.Image, OpenDialog.FileName
    PaintForm.Caption = FName & " - AtrociousPaint"
    PicChanged = False
    SC = True
    ClearImageBuffer
Exit Sub

ErrHandler:
    If UnLod = True Then
        SErr = True
        UnLod = False
    End If

End Sub

Private Sub FileSAW_Click()

Dim s

If OpenDialog.FileName = "" Or PicChanged = True Then
    s = MsgBox("You must save file before choosing it as wallpaper.", vbOKCancel + vbExclamation, "AtrociousPaint")
    Select Case s
    Case vbOK
    If SC = False Then
        SaveDialog.CancelError = True
        On Error GoTo ErrHandler
            SaveDialog.Filter = "Ahsan's Picture Files|*.abmp| All Files|*.*"
            SaveDialog.FilterIndex = 12
            SaveDialog.ShowSave
            FName = SaveDialog.FileTitle
            OpenDialog.FileName = SaveDialog.FileName
            SavePicture PaintBox.Image, OpenDialog.FileName
            PaintForm.Caption = FName & " - AtrociousPaint"
            PicChanged = False
            SC = True
        Exit Sub
    ElseIf SC = True Then
        SavePicture PaintBox.Image, OpenDialog.FileName
        PicChanged = False
    End If
        SetWallpaper
        Exit Sub
    Case Else
        Exit Sub
    End Select
Else
    SetWallpaper
End If

Exit Sub

ErrHandler:

End Sub

Private Sub FillFrame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbLeftButton Then
    If (y >= 125) And (y < 425) Then
      RectangleShape(0).BorderColor = vbWhite
      RectangleShape(1).BorderColor = vbBlack
      RectangleShape(2).BorderColor = vbBlack
      FillLabel.Top = 150
      SFillStyle = BorderOnly
    ElseIf (y >= 450 And y < 750) Then
      RectangleShape(0).BorderColor = vbBlack
      RectangleShape(1).BorderColor = vbWhite
      RectangleShape(2).BorderColor = vbBlack
      FillLabel.Top = 465
      SFillStyle = BorderFill
    ElseIf (y >= 775 And y < 1075) Then
      RectangleShape(0).BorderColor = vbBlack
      RectangleShape(1).BorderColor = vbBlack
      RectangleShape(2).BorderColor = vbWhite
      FillLabel.Top = 780
      SFillStyle = FillOnly
    End If
End If

End Sub

Private Sub FileClose_Click()

Unload Me

End Sub

Private Sub FileOpen_Click()

If PicChanged = True Then
    SaveChanges
    Select Case SMsg
    Case vbYes
        If SC = False Then
             Call FileSaveAs_Click
        ElseIf SC = True Then
            SavePicture PaintBox.Image, OpenDialog.FileName
            PicChanged = False
        End If
        OpenFile
    Case vbNo
        OpenFile
    Case Else
        Exit Sub
    End Select
Else
    OpenFile
End If

PlaceResizeBox
Form_Resize
ClearImageBuffer

End Sub

Private Sub FilterBlur_Click()

Dim POCol
Dim PBCol1, PBCol2, PBCol3, PBCol4, PBCol5, PBCol6, PBCol7, PBCol8
Dim R, G, b
Dim R1, R2, R3, R4, R5, R6, R7, R8
Dim G1, G2, G3, G4, G5, G6, G7, G8
Dim B1, B2, B3, B4, B5, B6, B7, B8
Dim BR, BG, BB
Dim FBConst
Dim DW

FBConst = 9

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To (PaintBox.ScaleWidth - 1)
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To (PaintBox.ScaleHeight - 1)
        POCol = PaintBox.Point(x, y)
        PBCol1 = PaintBox.Point(x, y - 1)
        PBCol2 = PaintBox.Point(x - 1, y)
        PBCol3 = PaintBox.Point(x - 1, y - 1)
        PBCol4 = PaintBox.Point(x, y + 1)
        PBCol5 = PaintBox.Point(x + 1, y)
        PBCol6 = PaintBox.Point(x + 1, y + 1)
        PBCol7 = PaintBox.Point(x + 1, y - 1)
        PBCol8 = PaintBox.Point(x - 1, y + 1)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        R1 = PBCol1 Mod 256
        G1 = (PBCol1 \ 256) Mod 256
        B1 = (PBCol1 \ 256) \ 256
        R2 = PBCol2 Mod 256
        G2 = (PBCol2 \ 256) Mod 256
        B2 = (PBCol2 \ 256) \ 256
        R3 = PBCol3 Mod 256
        G3 = (PBCol3 \ 256) Mod 256
        B3 = (PBCol3 \ 256) \ 256
        R4 = PBCol4 Mod 256
        G4 = (PBCol4 \ 256) Mod 256
        B4 = (PBCol4 \ 256) \ 256
        R5 = PBCol5 Mod 256
        G5 = (PBCol5 \ 256) Mod 256
        B5 = (PBCol5 \ 256) \ 256
        R6 = PBCol6 Mod 256
        G6 = (PBCol6 \ 256) Mod 256
        B6 = (PBCol6 \ 256) \ 256
        R7 = PBCol7 Mod 256
        G7 = (PBCol7 \ 256) Mod 256
        B7 = (PBCol7 \ 256) \ 256
        R8 = PBCol8 Mod 256
        G8 = (PBCol8 \ 256) Mod 256
        B8 = (PBCol8 \ 256) \ 256
        BR = (R + R1 + R2 + R3 + R4 + R5 + R6 + R7 + R8) / FBConst
        BG = (G + G1 + G2 + G3 + G4 + G5 + G6 + G7 + G8) / FBConst
        BB = (b + B1 + B2 + B3 + B4 + B5 + B6 + B7 + B8) / FBConst
        If BR > 255 Then BR = 255
        If BR < 0 Then BR = 0
        If BG > 255 Then BG = 255
        If BG < 0 Then BG = 0
        If BB > 255 Then BB = 255
        If BB < 0 Then BB = 0
        PaintBox.PSet (x, y), RGB(BR, BG, BB)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterBright_Click()

Dim POCol
Dim R, G, b
Dim BR, BG, BB
Dim FBConst
Dim DW

FBConst = 32

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        BR = R + FBConst
        BG = G + FBConst
        BB = b + FBConst
        If BR > 255 Then BR = 255
        If BR < 0 Then BR = 0
        If BG > 255 Then BG = 255
        If BG < 0 Then BG = 0
        If BB > 255 Then BB = 255
        If BB < 0 Then BB = 0
        PaintBox.PSet (x, y), RGB(BR, BG, BB)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterBW_Click()

Dim POCol
Dim PBWCol
Dim R, G, b
Dim FBWConst
Dim DW

FBWConst = 175

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        If R >= FBWConst Then
            PBWCol = vbWhite
        Else
            PBWCol = vbBlack
        End If
        If G >= FBWConst Then
            PBWCol = vbWhite
        Else
            PBWCol = vbBlack
        End If
        If b >= FBWConst Then
            PBWCol = vbWhite
        Else
            PBWCol = vbBlack
        End If
        PaintBox.PSet (x, y), PBWCol
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterCrease_Click()

Dim POCol
Dim PCCol
Dim R, G, b
Dim FCConst
Dim DW

FCConst = 500

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = GetPixel(PaintBox.hdc, x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        If R > 255 Then R = 255
        If R < 0 Then R = 0
        If G > 255 Then G = 255
        If G < 0 Then G = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        PCCol = SetPixel(PaintBox.hdc, x, (Sin(x) * FCConst) + y, RGB(R, G, b))
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterDark_Click()

Dim POCol
Dim R, G, b
Dim DR, DG, DB
Dim FDConst
Dim DW

FDConst = -32

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        DR = R + FDConst
        DG = G + FDConst
        DB = b + FDConst
        If DR > 255 Then DR = 255
        If DR < 0 Then DR = 0
        If DG > 255 Then DG = 255
        If DG < 0 Then DG = 0
        If DB > 255 Then DB = 255
        If DB < 0 Then DB = 0
        PaintBox.PSet (x, y), RGB(DR, DG, DB)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterDiffuse_Click()

Dim POCol
Dim PDCol
Dim PDColR, PDColG, PDColB
Dim R, G, b
Dim DR, DG, DB
Dim FDConst
Dim DW

FDConst = 5

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        PDColR = PaintBox.Point(x, y + Int((Rnd * FDConst) - 2))
        DR = PDColR Mod 256
        PDColG = PaintBox.Point(x + Int((Rnd * FDConst) - 2), y)
        DG = (PDColG \ 256) Mod 256
        PDColB = PaintBox.Point(x + Int((Rnd * FDConst) - 2), y + Int((Rnd * FDConst) - 2))
        DB = (PDColG \ 256) \ 256
        If DR > 255 Then DR = 255
        If DR < 0 Then DR = 0
        If DG > 255 Then DG = 255
        If DG < 0 Then DG = 0
        If DB > 255 Then DB = 255
        If DB < 0 Then DB = 0
        PaintBox.PSet (x, y), RGB(DR, DG, DB)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterEmboss_Click()

Dim POCol
Dim PEColR, PEColG, PEColB
Dim R, G, b
Dim R1, G1, B1
Dim ER, EG, EB
Dim FEConst
Dim DW

FEConst = 125

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        PEColR = PaintBox.Point(x + 1, y + 1)
        R1 = PEColR Mod 256
        PEColG = PaintBox.Point(x + 1, y + 1)
        G1 = (PEColG \ 256) Mod 256
        PEColB = PaintBox.Point(x + 1, y + 1)
        B1 = (PEColB \ 256) \ 256
        ER = R - R1 + FEConst
        EG = G - G1 + FEConst
        EB = b - B1 + FEConst
        If ER > 255 Then ER = 255
        If ER < 0 Then ER = 0
        If EG > 255 Then EG = 255
        If EG < 0 Then EG = 0
        If EB > 255 Then EB = 255
        If EB < 0 Then EB = 0
        PaintBox.PSet (x, y), RGB(ER, EG, EB)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterGBW_Click()

Dim POCol
Dim PECol
Dim R, G, b
Dim GBWR, GBWG, GBWB
Dim FGBWConst
Dim DW

FGBWConst = 3

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        GBWR = Abs(R * (G - b + G + R)) / 256
        GBWG = Abs(R * (b - G + b + R)) / 256
        GBWB = Abs(G * (b - G + b + R)) / 256
        PECol = (GBWR + GBWG + GBWB) / FGBWConst
        If PECol > 255 Then PECol = 255
        If PECol < 0 Then PECol = 0
        PaintBox.PSet (x, y), RGB(PECol, PECol, PECol)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterGray_Click()

Dim POCol
Dim PGCol
Dim R, G, b
Dim GR, GG, GB
Dim FGConst
Dim DW

FGConst = 0.32

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        GR = R * FGConst
        GG = G * FGConst
        GB = b * FGConst
        PGCol = GR + GG + GB
        If GR > 255 Then GR = 255
        If GR < 0 Then GR = 0
        If GG > 255 Then GG = 255
        If GG < 0 Then GG = 0
        If GB > 255 Then GB = 255
        If GB < 0 Then GB = 0
        PaintBox.PSet (x, y), RGB(PGCol, PGCol, PGCol)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterInvertColor_Click()

FilterPic.Height = PaintBox.Height
FilterPic.Width = PaintBox.Width
FilterPic.Picture = PaintBox.Image
PaintBox.Picture = Nothing

PaintBox.PaintPicture FilterPic.Image, 0, 0, FilterPic.Width, FilterPic.Height, , , , , vbSrcInvert

PaintBox.Refresh
Tools(7).Value = False
PicChanged = True

SetImageBuffer

End Sub

Private Sub FilterRepColor_Click()

Dim POCol
Dim PRCol
Dim R, G, b
Dim DW

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        If R > 255 Then R = 255
        If R < 0 Then R = 0
        If G > 255 Then G = 255
        If G < 0 Then G = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        PRCol = RGB(R, G, b)
        If PRCol = ForeColorLabel.BackColor Then
            PRCol = FillLabel.BackColor
        ElseIf PRCol = FillColorLabel.BackColor Then
            PRCol = ForeColorLabel.BackColor
        End If
        PaintBox.PSet (x, y), PRCol
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterSharp_Click()

Dim POCol
Dim PSCol
Dim R, G, b
Dim R1, G1, B1
Dim SR, SG, SB
Dim FSConst
Dim DW

FSConst = 0.5

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        PSCol = PaintBox.Point(x - 1, y - 1)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        R1 = PSCol Mod 256
        G1 = (PSCol \ 256) Mod 256
        B1 = (PSCol \ 256) \ 256
        SR = R + (FSConst * (R - R1))
        SG = G + (FSConst * (G - G1))
        SB = b + (FSConst * (b - B1))
        PaintBox.PSet (x, y), RGB(Abs(SR), Abs(SG), Abs(SB))
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterSnow_Click()

Dim POCol
Dim PSCol
Dim R, G, b
Dim R1, G1, B1
Dim SR, SG, SB
Dim FSConst
Dim DW

FSConst = 0.9

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        PSCol = PaintBox.Point(x - 1, y - 1)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        R1 = PSCol Mod 256
        G1 = (PSCol \ 256) Mod 256
        B1 = (PSCol \ 256) \ 256
        SR = R + (FSConst * (R - R1))
        SG = G + (FSConst * (G - G1))
        SB = b + (FSConst * (b - B1))
        PaintBox.PSet (x, y), RGB(Abs(SR), Abs(SG), Abs(SB))
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterSolarize_Click()

Dim POCol
Dim R, G, b
Dim DW


DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = PaintBox.Point(x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        If ((R < 128) Or (R > 255)) Then R = 255 - R
        If ((G < 128) Or (G > 255)) Then G = 255 - G
        If ((b < 128) Or (b > 255)) Then b = 255 - b
        PaintBox.PSet (x, y), RGB(R, G, b)
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FilterWave_Click()

Dim POCol
Dim PWCol
Dim R, G, b
Dim FWConst
Dim DW

FWConst = 15

DW = PaintBox.DrawWidth
PaintBox.DrawWidth = 1

For x = 1 To PaintBox.ScaleWidth
    ProgressBar.Visible = True
    ProgressBar.Value = (x / PaintBox.ScaleWidth) * 100
    For y = 1 To PaintBox.ScaleHeight
        POCol = GetPixel(PaintBox.hdc, x, y)
        R = POCol Mod 256
        G = (POCol \ 256) Mod 256
        b = (POCol \ 256) \ 256
        If R > 255 Then R = 255
        If R < 0 Then R = 0
        If G > 255 Then G = 255
        If G < 0 Then G = 0
        If b > 255 Then b = 255
        If b < 0 Then b = 0
        PWCol = SetPixel(PaintBox.hdc, x, Abs((Sin(x) * FWConst) + y), RGB(R, G, b))
    Next y
    PaintBox.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Filtering....."
Next x

PaintBox.DrawWidth = DW
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
Tools(7).Value = False
PicChanged = True
PaintForm.Enabled = True

SetImageBuffer

End Sub

Private Sub FlipBoth_Click()

FlipHorizontal_Click
FlipVertical_Click

End Sub

Private Sub FlipHorizontal_Click()

FilterPic.Height = PaintBox.Height
FilterPic.Width = PaintBox.Width
FilterPic.Picture = PaintBox.Image
PaintBox.Picture = Nothing

PaintBox.PaintPicture FilterPic.Image, FilterPic.ScaleWidth, 0, -FilterPic.ScaleWidth, FilterPic.ScaleHeight, , , , , vbSrcCopy

PaintBox.Refresh
PicChanged = True

SetImageBuffer

End Sub

Private Sub FlipVertical_Click()

FilterPic.Height = PaintBox.Height
FilterPic.Width = PaintBox.Width
FilterPic.Picture = PaintBox.Image
PaintBox.Picture = Nothing

PaintBox.PaintPicture FilterPic.Image, 0, FilterPic.ScaleHeight, FilterPic.ScaleWidth, -FilterPic.ScaleHeight, , , , , vbSrcCopy

PaintBox.Refresh
PicChanged = True

SetImageBuffer

End Sub

Private Sub Form_Activate()

Dim a

FStyle.Text = "Solid"
FStyle.AddItem "Solid"
FStyle.AddItem "Transparent"
FStyle.AddItem "Horizontal Line"
FStyle.AddItem "Vertical Line"
FStyle.AddItem "Upward Diagonal"
FStyle.AddItem "Downward Diagonal"
FStyle.AddItem "Cross"
FStyle.AddItem "Diagonal Cross"

BStyle.Text = "Solid"
BStyle.AddItem "Solid"
BStyle.AddItem "Dash"
BStyle.AddItem "Dot"
BStyle.AddItem "Dash-Dot"
BStyle.AddItem "Dash-Dot-Dot"

DWCombo.Text = 1
For a = 1 To 7
    DWCombo.AddItem a
Next

stbl = False

FImgZoom = False

PaintBox.SetFocus

DrawWidth = 1

EBorderStyle = 0

Tools_Click (4)

StatusBarTimer_Timer

Form_Resize

PlaceResizeBox

ClearImageBuffer

End Sub

Private Sub Form_Load()

Dim cmd As String, StrBuff As String
Dim tFile As Long
Dim a As String
cmd = Command$

If Len(Trim(cmd)) <= 0 Then
    PaintForm.Caption = "Ahsan.abmp - AtrociousPaint"
    FName = "Ahsan.abmp"
    SC = False
    Exit Sub
Else
    a = InStrRev(cmd, "\")
    FName = Mid(cmd, a + 1)
    PaintForm.Caption = FName & " - AtrociousPaint"
    tFile = FreeFile
    Open cmd For Binary Access Read As #tFile
        StrBuff = Space(LOF(tFile))
        Get #tFile, , StrBuff
    Close #tFile
    PaintBox.Picture = LoadPicture(cmd)
    OpenDialog.FileName = cmd
    SC = True
    StrBuff = ""
    cmd = ""
End If

PicChanged = False
SErr = False
UnLod = False

End Sub

Public Sub Form_Resize()

On Error Resume Next

If PaintForm.WindowState <> vbMinimized Then
    ColorFrame.Top = Me.ScaleHeight - 1150
    StatusBar1.Top = Me.ScaleHeight - 325
    VScroll.Left = PaintForm.Width - VScroll.Width - 100
    VScroll.Height = PaintForm.Height - (ColorFrame.Height + HScroll.Height + StatusBar1.Height + PropertyFrame.Height + 325)
    HScroll.Top = ColorFrame.Top - HScroll.Height + 110
    HScroll.Width = PaintForm.Width - (ToolsFrame.Width + VScroll.Width + 100)
    ProgressBar.Width = PaintForm.Width
    
    If PaintForm.Height < 5500 Then
      PaintForm.Height = 5500
    End If
    
    If PaintForm.Width < 4550 Then
      PaintForm.Width = 4550
    End If
    
    With HScroll
      If VScroll.Visible Then
        .Max = (PaintBox.Width - (Me.Width - VScroll.Width - 1050)) / 10
      Else
        .Max = (PaintBox.Width - (Me.Width - 1050)) / 10
      End If
      .Visible = (.Max > 0)
      If .Visible Then
        .Top = ColorFrame.Top - .Height + 110
        If VScroll.Visible Then
          .Width = Me.Width - ToolsFrame.Width - VScroll.Width - 90
        Else
          .Width = Me.Width - ToolsFrame.Width - 90
        End If
      End If
    End With
    
    With VScroll
      If HScroll.Visible Then
        .Max = (PaintBox.Height - ((Me.Height - HScroll.Height) / 1.65))
      Else
        .Max = (PaintBox.Height - (Me.Height / 1.65))
      End If
      .Visible = (.Max > 0)
      If .Visible Then
        .Left = Me.Width - .Width - 110
        If HScroll.Visible Then
          .Height = Me.ScaleHeight - ColorFrame.Height - PropertyFrame.Height - HScroll.Height - 50
        Else
          .Height = Me.ScaleHeight - ColorFrame.Height - PropertyFrame.Height - 100
        End If
      End If
    End With
    
    If HScroll.Visible And VScroll.Visible Then
      ScrollFrame.Visible = True
      ScrollFrame.Left = VScroll.Left
      ScrollFrame.Top = HScroll.Top
    Else
      ScrollFrame.Visible = False
    End If
End If

PlaceResizeBox

PaintForm.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

If PicChanged = True Then
    SaveChanges
    Select Case SMsg
    Case vbYes
        UnLod = True
        If SC = False Then
            FileSaveAs_Click
        ElseIf SC = True Then
            SavePicture PaintBox.Image, OpenDialog.FileName
        End If
        If SErr = True Then
            Cancel = True
        Else
            Cancel = False
        End If
    Case vbNo
        Exit Sub
    Case vbCancel
        Cancel = True
    End Select
Else
   Exit Sub
End If

End Sub

Private Sub FStyle_Click()

Select Case FStyle.Text
Case "Solid"
    FFillStyle = vbFSSolid
Case "Transparent"
    FFillStyle = vbFSTransparent
Case "Horizontal Line"
    FFillStyle = vbHorizontalLine
Case "Vertical Line"
    FFillStyle = vbVerticalLine
Case "Upward Diagonal"
    FFillStyle = vbUpwardDiagonal
Case "Downward Diagonal"
    FFillStyle = vbDownwardDiagonal
Case "Cross"
    FFillStyle = vbCross
Case "Diagonal Cross"
    FFillStyle = vbDiagonalCross
End Select

End Sub

Private Sub HelpTopics_Click()

OpenDialog.HelpFile = App.HelpFile
OpenDialog.HelpCommand = cdlHelpContents
OpenDialog.ShowHelp

End Sub

Private Sub HResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

HMD = True
rx = Val(x)

End Sub

Private Sub HResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

If HMD = True Then

    If PaintBox.Width >= 300 Then
        PaintBox.Width = Val(PaintBox.Width) + Val(rx) + Val(x)
    End If

    If PaintBox.Width < 300 Then
        PaintBox.Width = 300
    End If

End If

Form_Resize

PlaceResizeBox

End Sub

Private Sub HResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

HMD = False

End Sub

Private Sub HScroll_Change()

Dim PaintBoxLeft As Long

PaintBoxLeft = CLng(ToolsFrame.Width) - (CLng(HScroll.Value) * 10)
PaintBox.Left = Val(PaintBoxLeft)

If TextBox.Visible = True Then
    TextBox.Visible = False
End If

PlaceResizeBox

End Sub

Private Sub PaintAbout_Click()

Load AboutPaint
AboutPaint.Show

End Sub

Private Sub PaintBox_DblClick()

On Error Resume Next

Select Case CurInd
Case 13
    If DPolygon Then
        DrawPolygon Complete:=False
        DrawPolygon
        DPolygon = False
    End If
Case 15
    PaintBox.DrawMode = 13
    PolyBezier PaintBox.hdc, BezierPoints(0), 4
    PaintBox.Refresh
    CurrentBezierPoint = 0
    PicChanged = True
    SetImageBuffer
End Select

PaintBox.DrawMode = 13
PaintBox.ForeColor = ForeColorLabel.BackColor

End Sub

Private Sub PaintBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

PaintBox.DrawMode = 13
PaintBox.ForeColor = ForeColorLabel.BackColor

If Button = 1 Then
    OldX = x
    OldY = y
    mode = True
End If

Select Case CurInd
Case 0
    If SelectionBox.Visible Then
        SelectionBox.Visible = False
        ImageSelector.Visible = False
        If SI = 0 Then
            PaintBox.PaintPicture SelectionBox.Image, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
        ElseIf SI = 1 Then
            PaintBox.Line (SelectionBox.Left, SelectionBox.Top)-(SelectionBox.Left + SelectionBox.Width, SelectionBox.Top + SelectionBox.Height), vbWhite, BF
            PaintBox.PaintPicture SelectionBox.Image, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
        End If
        EditCut.Enabled = False
        EditCopy.Enabled = False
        EditDel.Enabled = False
        EditCrop.Enabled = False
        ImgMov = False
        PicChanged = True
        SetImageBuffer
    End If
Case 1
    PaintBox.ScaleMode = 1
    If FImgZoom = False Then
        ImgZoom2.Height = PaintBox.Height
        ImgZoom2.Width = PaintBox.Width
        ImgZoom2.Picture = PaintBox.Image
        FImgZoom = True
    End If
    If Button = vbLeftButton Then
        If PaintBox.Height < 20000 Or PaintBox.Width < 20000 Then
            If ImgZoom1.Height = ImgZoom2.Height Or ImgZoom1.Width = ImgZoom2.Width Then
                ImgZoom1.Cls
                ImgZoom1.Picture = ImgZoom2.Image
                PaintBox.Height = PaintBox.Height * 1.25
                PaintBox.Width = PaintBox.Width * 1.25
                PaintBox.Cls
                PaintBox.PaintPicture ImgZoom1.Image, 0, 0, PaintBox.Width, PaintBox.Height
            Else
                ImgZoom1.Height = PaintBox.Height
                ImgZoom1.Width = PaintBox.Width
                ImgZoom1.Cls
                ImgZoom1.Picture = PaintBox.Image
                PaintBox.Height = PaintBox.Height * 1.25
                PaintBox.Width = PaintBox.Width * 1.25
                PaintBox.Cls
                PaintBox.PaintPicture ImgZoom1.Image, 0, 0, PaintBox.Width, PaintBox.Height
            End If
        Else
            MsgBox "Cannot zoom more.", vbCritical, "AtrociousPaint"
        End If
    ElseIf Button = vbRightButton Then
        If PaintBox.Height > 300 Or PaintBox.Width > 300 Then
            If ImgZoom1.Height = ImgZoom2.Height Or ImgZoom1.Width = ImgZoom2.Width Then
                ImgZoom1.Cls
                ImgZoom1.Picture = ImgZoom2.Image
                PaintBox.Height = PaintBox.Height / 1.25
                PaintBox.Width = PaintBox.Width / 1.25
                PaintBox.Cls
                PaintBox.PaintPicture ImgZoom1.Image, 0, 0, PaintBox.Width, PaintBox.Height
            Else
                ImgZoom1.Height = PaintBox.Height
                ImgZoom1.Width = PaintBox.Width
                ImgZoom1.Cls
                ImgZoom1.Picture = PaintBox.Image
                PaintBox.Height = PaintBox.Height / 1.25
                PaintBox.Width = PaintBox.Width / 1.25
                PaintBox.Cls
                PaintBox.PaintPicture ImgZoom1.Image, 0, 0, PaintBox.Width, PaintBox.Height
            End If
        End If
    End If
    PaintBox.ScaleMode = 3
    Form_Resize
    PlaceResizeBox
Case 2
    If Button = vbLeftButton Then
        ForeColorLabel.BackColor = PaintBox.Point(x, y)
        PaintBox.ForeColor = ForeColorLabel.BackColor
    ElseIf Button = vbRightButton Then
        FillColorLabel.BackColor = PaintBox.Point(x, y)
        PaintBox.FillColor = FillColorLabel.BackColor
    End If
    mode = False
Case 4
    If Button = vbLeftButton Then
        PaintBox.PSet (x, y)
    End If
Case 6
    If Button = vbLeftButton Then
        DrawSpray CInt(x), CInt(y), PaintBox.DrawWidth * 4
        PicChanged = True
        SetImageBuffer
    End If
Case 7
    If Button = vbLeftButton Then
        DrawBrush BrushShape, x, y
        PicChanged = True
        SetImageBuffer
    End If
Case 8
    If Button = vbLeftButton Then
        LineShape.Visible = True
        LineShape.X1 = x
        LineShape.Y1 = y
        LineShape.X2 = x + 5
        LineShape.Y2 = y + 5
    End If
Case 9
    If Button = vbLeftButton Then
        LineShape.Visible = True
        LineShape.X1 = x
        LineShape.Y1 = y
        LineShape.X2 = x + 5
        LineShape.Y2 = y + 5
    End If
Case 14
    If Button = vbLeftButton Then
        If TextBox.Visible = False Then
            TextBox.BackColor = PaintBox.BackColor
            TextBox.ForeColor = ForeColorLabel.BackColor
            TextBox.FontSize = PaintBox.DrawWidth * 2
            TextLabel.FontSize = PaintBox.FontSize * 2
            PaintBox.FontSize = TextBox.FontSize
            TextBox.Left = x
            TextBox.Top = y
            TextBox.Visible = True
            TextBox.SetFocus
        Else
            TextBox_DblClick
        End If
    End If
Case 13
    If Button = vbLeftButton Then
        PaintBox.DrawMode = vbXorPen
        PaintBox.ForeColor = PaintBox.BackColor Xor ForeColorLabel.BackColor
        If DPolygon = False Then
            DPolygon = True
            ReDim PolygonLen(0)
            PolygonLen(0).x = x
            PolygonLen(0).y = y
        Else
            ReDim Preserve PolygonLen(UBound(PolygonLen) + 1)
            PolygonLen(UBound(PolygonLen)).x = x
            PolygonLen(UBound(PolygonLen)).y = y
            DrawPolygon Complete:=False
        End If
        PicChanged = True
        SetImageBuffer
    End If
Case 15
    If Button = vbLeftButton Then
        PaintBox.DrawMode = 7
        PaintBox.DrawStyle = EBorderStyle
        PaintBox.ForeColor = PaintBox.BackColor Xor ForeColorLabel.BackColor
        If CurrentBezierPoint <> 0 Then
            PolyBezier PaintBox.hdc, BezierPoints(0), 4
        End If
        If CurrentBezierPoint = 0 Then
            PaintBox.DrawMode = 10
            PaintBox.CurrentX = x
            PaintBox.CurrentY = y
            For i = 0 To 3
                BezierPoints(i).x = x
                BezierPoints(i).y = y
            Next i
        End If
        If CurrentBezierPoint = 1 Then
            BezierPoints(1).x = x
            BezierPoints(1).y = y
            BezierPoints(2).x = x
            BezierPoints(2).y = y
        End If
        If CurrentBezierPoint = 2 Then
            BezierPoints(2).x = x
            BezierPoints(2).y = y
        End If
        CurrentBezierPoint = CurrentBezierPoint + 1
        PolyBezier PaintBox.hdc, BezierPoints(0), 4
        PaintBox.Refresh
        PicChanged = True
        SetImageBuffer
    End If
End Select
    
End Sub

Private Sub PaintBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

PaintBox.DrawMode = 13
PaintBox.ForeColor = ForeColorLabel.BackColor

If Button = 1 Then
     If mode = True Then
        Select Case CurInd
        Case 0
            If OldX = 0 And OldY = 0 Then OldX = x: OldY = y
                ImageSelector.Visible = True
                If x <> OldX And y <> OldY Then
                    If x > OldX And y > OldY Then
                        ImageSelector.Left = OldX
                        ImageSelector.Top = OldY
                        ImageSelector.Width = x - OldX
                        ImageSelector.Height = y - OldY
                    ElseIf x > OldX And y < OldY Then
                        ImageSelector.Left = OldX
                        ImageSelector.Top = y
                        ImageSelector.Width = x - OldX
                        ImageSelector.Height = OldY - y
                    ElseIf x < OldX And y > OldY Then
                        ImageSelector.Left = x
                        ImageSelector.Top = OldY
                        ImageSelector.Width = OldX - x
                        ImageSelector.Height = y - OldY
                    ElseIf x < OldX And y < OldY Then
                        ImageSelector.Left = x
                        ImageSelector.Top = y
                        ImageSelector.Width = OldX - x
                        ImageSelector.Height = OldY - y
                    End If
                End If
        Case 4
            PaintBox.Line (OldX, OldY)-(x, y)
            OldX = x
            OldY = y
        Case 5
            PaintBox.ForeColor = vbWhite
            PaintBox.Line (OldX, OldY)-(x, y)
            OldX = x
            OldY = y
        Case 6
            DrawSpray CInt(x), CInt(y), PaintBox.DrawWidth * 4
        Case 7
            DrawBrush BrushShape, x, y
        Case 8
            LineShape.BorderWidth = PaintBox.DrawWidth
            LineShape.BorderColor = PaintBox.ForeColor
            LineShape.BorderStyle = EBorderStyle + 1
            If OldX = 0 And OldY = 0 Then OldX = x: OldY = y
                LineShape.X1 = OldX
                LineShape.Y1 = OldY
                LineShape.X2 = x
                LineShape.Y2 = y
        Case 9
            LineShape.BorderWidth = PaintBox.DrawWidth
            LineShape.BorderColor = PaintBox.ForeColor
            LineShape.BorderStyle = EBorderStyle + 1
            If OldX = 0 And OldY = 0 Then OldX = x: OldY = y
                LineShape.X1 = OldX
                LineShape.Y1 = OldY
                LineShape.X2 = x
                LineShape.Y2 = y
        Case 10
            With RectShape
            .BorderWidth = PaintBox.DrawWidth
            Select Case SFillStyle
                Case BorderOnly
                    .FillStyle = 1
                    .BorderColor = ForeColorLabel.BackColor
                Case BorderFill
                    .FillStyle = 0
                    .BorderColor = ForeColorLabel.BackColor
                    .FillColor = FillColorLabel.BackColor
                Case FillOnly
                    .FillStyle = 0
                    .BorderColor = FillColorLabel.BackColor
                    .FillColor = FillColorLabel.BackColor
            End Select
            End With
            RectShape.BorderStyle = EBorderStyle + 1
            If OldX = 0 And OldY = 0 Then OldX = x: OldY = y
                RectShape.Visible = True
                If x <> OldX And y <> OldY Then
                    If x > OldX And y > OldY Then
                        RectShape.Left = OldX
                        RectShape.Top = OldY
                        RectShape.Width = x - OldX
                        RectShape.Height = y - OldY
                    ElseIf x > OldX And y < OldY Then
                        RectShape.Left = OldX
                        RectShape.Top = y
                        RectShape.Width = x - OldX
                        RectShape.Height = OldY - y
                    ElseIf x < OldX And y > OldY Then
                        RectShape.Left = x
                        RectShape.Top = OldY
                        RectShape.Width = OldX - x
                        RectShape.Height = y - OldY
                    ElseIf x < OldX And y < OldY Then
                        RectShape.Left = x
                        RectShape.Top = y
                        RectShape.Width = OldX - x
                        RectShape.Height = OldY - y
                    End If
                End If
        Case 11
            With RRectShape
            .BorderWidth = PaintBox.DrawWidth
            Select Case SFillStyle
                Case BorderOnly
                    .FillStyle = 1
                    .BorderColor = ForeColorLabel.BackColor
                Case BorderFill
                    .FillStyle = 0
                    .BorderColor = ForeColorLabel.BackColor
                    .FillColor = FillColorLabel.BackColor
                Case FillOnly
                    .FillStyle = 0
                    .BorderColor = FillColorLabel.BackColor
                    .FillColor = FillColorLabel.BackColor
            End Select
            End With
            RRectShape.BorderStyle = EBorderStyle + 1
            If OldX = 0 And OldY = 0 Then OldX = x: OldY = y
                RRectShape.Visible = True
                If x <> OldX And y <> OldY Then
                    If x > OldX And y > OldY Then
                        RRectShape.Left = OldX
                        RRectShape.Top = OldY
                        RRectShape.Width = x - OldX
                        RRectShape.Height = y - OldY
                    ElseIf x > OldX And y < OldY Then
                        RRectShape.Left = OldX
                        RRectShape.Top = y
                        RRectShape.Width = x - OldX
                        RRectShape.Height = OldY - y
                    ElseIf x < OldX And y > OldY Then
                        RRectShape.Left = x
                        RRectShape.Top = OldY
                        RRectShape.Width = OldX - x
                        RRectShape.Height = y - OldY
                    ElseIf x < OldX And y < OldY Then
                        RRectShape.Left = x
                        RRectShape.Top = y
                        RRectShape.Width = OldX - x
                        RRectShape.Height = OldY - y
                    End If
                End If
        Case 12
            With EllipseShape
            .BorderWidth = PaintBox.DrawWidth
            Select Case SFillStyle
                Case BorderOnly
                    .FillStyle = 1
                    .BorderColor = ForeColorLabel.BackColor
                Case BorderFill
                    .FillStyle = 0
                    .BorderColor = ForeColorLabel.BackColor
                    .FillColor = FillColorLabel.BackColor
                Case FillOnly
                    .FillStyle = 0
                    .BorderColor = FillColorLabel.BackColor
                    .FillColor = FillColorLabel.BackColor
            End Select
            End With
            EllipseShape.BorderStyle = EBorderStyle + 1
            If OldX = 0 And OldY = 0 Then OldX = x: OldY = y
                EllipseShape.Visible = True
                If x <> OldX And y <> OldY Then
                    If x > OldX And y > OldY Then
                        EllipseShape.Left = OldX
                        EllipseShape.Top = OldY
                        EllipseShape.Width = x - OldX
                        EllipseShape.Height = y - OldY
                    ElseIf x > OldX And y < OldY Then
                        EllipseShape.Left = OldX
                        EllipseShape.Top = y
                        EllipseShape.Width = x - OldX
                        EllipseShape.Height = OldY - y
                    ElseIf x < OldX And y > OldY Then
                        EllipseShape.Left = x
                        EllipseShape.Top = OldY
                        EllipseShape.Width = OldX - x
                        EllipseShape.Height = y - OldY
                    ElseIf x < OldX And y < OldY Then
                        EllipseShape.Left = x
                        EllipseShape.Top = y
                        EllipseShape.Width = OldX - x
                        EllipseShape.Height = OldY - y
                    End If
                End If
        Case 13
            PaintBox.DrawMode = 7
            If UBound(PolygonLen) = 0 Then
                ReDim Preserve PolygonLen(UBound(PolygonLen) + 1)
            Else
                DrawPolygon Complete:=False
            End If
            PolygonLen(UBound(PolygonLen)).x = x
            PolygonLen(UBound(PolygonLen)).y = y
            DrawPolygon Complete:=False
        Case 15
            PaintBox.DrawMode = 7
            PaintBox.ForeColor = PaintBox.BackColor Xor ForeColorLabel.BackColor
            PolyBezier PaintBox.hdc, BezierPoints(0), 4
            If CurrentBezierPoint = 1 Then
                BezierPoints(3).x = x
                BezierPoints(3).y = y
            End If
            If CurrentBezierPoint = 2 Then
                BezierPoints(1).x = x
                BezierPoints(1).y = y
                BezierPoints(2).x = x
                BezierPoints(2).y = y
            End If
            If CurrentBezierPoint = 3 Then
                BezierPoints(2).x = x
                BezierPoints(2).y = y
            End If
            PolyBezier PaintBox.hdc, BezierPoints(0), 4
            PaintBox.Refresh
        End Select
    End If
End If

StatusBar1.Panels(3).Text = CStr(x) & "," & CStr(y)

End Sub

Private Sub PaintBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

PaintBox.DrawMode = 13
PaintBox.ForeColor = ForeColorLabel.BackColor

If Button = 1 Then
     If mode = True Then
        Select Case CurInd
        Case 0
            If ImageSelector.Width > 2 And ImageSelector.Height > 2 Then
                SelectionBox.Width = ImageSelector.Width - 2
                SelectionBox.Height = ImageSelector.Height - 2
                SelectionBox.Left = ImageSelector.Left + 1
                SelectionBox.Top = ImageSelector.Top + 1
                SelectionBox.Visible = True
                SelectionBox.Picture = Nothing
                If SI = 0 Then
                    SelectionBox.PaintPicture PaintBox.Image, 0, 0, SelectionBox.Width, SelectionBox.Height, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
                ElseIf SI = 1 Then
                    SelectionBox.PaintPicture PaintBox.Image, 0, 0, SelectionBox.Width, SelectionBox.Height, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
                    PaintBox.Line (SelectionBox.Left, SelectionBox.Top)-(SelectionBox.Left + SelectionBox.Width, SelectionBox.Top + SelectionBox.Height), vbWhite, BF
                End If
                OldX = x
                OldY = y
                ImgMov = True
                EditCut.Enabled = True
                EditCopy.Enabled = True
                EditDel.Enabled = True
                EditCrop.Enabled = True
            Else
                ImageSelector.Visible = False
                SelectionBox.Visible = False
                EditCrop.Enabled = False
                ImgMov = False
            End If
        Case 3
            PaintBox.FillStyle = FFillStyle
            PaintBox.FillColor = ForeColorLabel.BackColor
            ExtFloodFill PaintBox.hdc, x, y, PaintBox.Point(x, y), 1
            PicChanged = True
            SetImageBuffer
        Case 4
            If Button = vbLeftButton Then
                PaintBox.PSet (x, y)
                PicChanged = True
                SetImageBuffer
            End If
        Case 5
            PaintBox.ForeColor = ForeColorLabel.BackColor
            PicChanged = True
            SetImageBuffer
        Case 8
            PaintBox.DrawStyle = EBorderStyle
            LineShape.Visible = False
            PaintBox.Line (LineShape.X1, LineShape.Y1)-(LineShape.X2, LineShape.Y2)
            OldX = 0
            OldY = 0
            PicChanged = True
            SetImageBuffer
        Case 9
            PaintBox.DrawStyle = EBorderStyle
            LineShape.Visible = False
            DrawArrow LineShape.X1, LineShape.Y1, LineShape.X2, LineShape.Y2
            OldX = 0
            OldY = 0
            PicChanged = True
            SetImageBuffer
        Case 10
            PaintBox.DrawStyle = EBorderStyle
            RectShape.Visible = False
            PaintBox.FillStyle = RectShape.FillStyle
            PaintBox.FillColor = RectShape.FillColor
            PaintBox.Line (RectShape.Left, RectShape.Top)-((RectShape.Left + RectShape.Width), (RectShape.Top + RectShape.Height)), RectShape.BorderColor, B
            OldX = x
            OldY = y
            PicChanged = True
            SetImageBuffer
        Case 11
            PaintBox.DrawStyle = EBorderStyle
            RRectShape.Visible = False
            PaintBox.FillStyle = RRectShape.FillStyle
            PaintBox.FillColor = RRectShape.FillColor
            RoundRect PaintBox.hdc, RRectShape.Left, RRectShape.Top, RRectShape.Left + RRectShape.Width, RRectShape.Top + RRectShape.Height, Int(RRectShape.Width * 0.21), Int(RRectShape.Height * 0.21)
            OldX = x
            OldY = y
            PicChanged = True
            SetImageBuffer
        Case 12
            PaintBox.DrawStyle = EBorderStyle
            EllipseShape.Visible = False
            PaintBox.FillStyle = EllipseShape.FillStyle
            PaintBox.FillColor = EllipseShape.FillColor
            Ellipse PaintBox.hdc, EllipseShape.Left, EllipseShape.Top, EllipseShape.Left + EllipseShape.Width, EllipseShape.Top + EllipseShape.Height
            OldX = x
            OldY = y
            PicChanged = True
            SetImageBuffer
        Case 15
            If CurrentBezierPoint = 3 Then
                PaintBox.DrawMode = 13
                PaintBox.DrawStyle = EBorderStyle
                PolyBezier PaintBox.hdc, BezierPoints(0), 4
                PaintBox.Refresh
                CurrentBezierPoint = 0
                PicChanged = True
                SetImageBuffer
            End If
        End Select
    End If
End If

mode = False

End Sub

Private Sub Rotate135_Click()

PicRotate.Height = PaintBox.Height
PicRotate.Width = PaintBox.Width
PicRotate.Picture = PaintBox.Image
PaintBox.Picture = Nothing

If RotateClockwise.Checked = True Then
    ImageRotate PicRotate, PaintBox, 135, True
ElseIf RotateAClockwise = True Then
    ImageRotate PicRotate, PaintBox, 135, False
End If

PicRotate = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub Rotate180_Click()

PicRotate.Height = PaintBox.Height
PicRotate.Width = PaintBox.Width
PicRotate.Picture = PaintBox.Image
PaintBox.Picture = Nothing

If RotateClockwise.Checked = True Then
    ImageRotate PicRotate, PaintBox, 180, True
ElseIf RotateAClockwise = True Then
    ImageRotate PicRotate, PaintBox, 180, False
End If

PicRotate = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub Rotate225_Click()

PicRotate.Height = PaintBox.Height
PicRotate.Width = PaintBox.Width
PicRotate.Picture = PaintBox.Image
PaintBox.Picture = Nothing

If RotateClockwise.Checked = True Then
    ImageRotate PicRotate, PaintBox, 225, True
ElseIf RotateAClockwise = True Then
    ImageRotate PicRotate, PaintBox, 225, False
End If

PicRotate = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub Rotate270_Click()

PicRotate.Height = PaintBox.Height
PicRotate.Width = PaintBox.Width
PicRotate.Picture = PaintBox.Image
PaintBox.Picture = Nothing

If RotateClockwise.Checked = True Then
    ImageRotate PicRotate, PaintBox, 270, True
ElseIf RotateAClockwise = True Then
    ImageRotate PicRotate, PaintBox, 270, False
End If

PicRotate = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub Rotate315_Click()

PicRotate.Height = PaintBox.Height
PicRotate.Width = PaintBox.Width
PicRotate.Picture = PaintBox.Image
PaintBox.Picture = Nothing

If RotateClockwise.Checked = True Then
    ImageRotate PicRotate, PaintBox, 315, True
ElseIf RotateAClockwise = True Then
    ImageRotate PicRotate, PaintBox, 315, False
End If

PicRotate = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub Rotate45_Click()

PicRotate.Height = PaintBox.Height
PicRotate.Width = PaintBox.Width
PicRotate.Picture = PaintBox.Image
PaintBox.Picture = Nothing

If RotateClockwise.Checked = True Then
    ImageRotate PicRotate, PaintBox, 45, True
ElseIf RotateAClockwise = True Then
    ImageRotate PicRotate, PaintBox, 45, False
End If

PicRotate = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub Rotate90_Click()

PicRotate.Height = PaintBox.Height
PicRotate.Width = PaintBox.Width
PicRotate.Picture = PaintBox.Image
PaintBox.Picture = Nothing

If RotateClockwise.Checked = True Then
    ImageRotate PicRotate, PaintBox, 90, True
ElseIf RotateAClockwise = True Then
    ImageRotate PicRotate, PaintBox, 90, False
End If

PicRotate = Nothing
PicChanged = True

SetImageBuffer

End Sub

Private Sub RotateAClockwise_Click()

RotateClockwise.Checked = False
RotateAClockwise.Checked = True

End Sub

Private Sub RotateClockwise_Click()

RotateClockwise.Checked = True
RotateAClockwise.Checked = False

End Sub

Private Sub SelectionBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

ImgMov = True
ImgMovX = x
ImgMovY = y

End Sub

Private Sub SelectionBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If ImgMov = True And Button = vbLeftButton Then
    SelectionBox.Move SelectionBox.Left + (x - ImgMovX), SelectionBox.Top + (y - ImgMovY)
    ImageSelector.Move ImageSelector.Left + (x - ImgMovX), ImageSelector.Top + (y - ImgMovY)
End If

End Sub

Private Sub SelectionType_Click(Index As Integer)

SI = Index

Select Case SelectionType(Index).Index
    Case 0
        SelectionLabel.Move SelectionLabel.Left, 115
        SelectionType(0).ForeColor = vbWhite
        SelectionType(1).ForeColor = vbBlack
        PSelectType = TransparentSelect
    Case 1
        SelectionType(0).ForeColor = vbBlack
        SelectionType(1).ForeColor = vbWhite
        SelectionLabel.Move SelectionLabel.Left, 535
        PSelectType = FilledSelect
End Select

End Sub

Private Sub StatusBarTimer_Timer()

If stbl = False Then
    StatusBar1.Panels(1).Text = "For Help, click Help Topics on the Help Menu."
End If

stbl = True

End Sub

Private Sub TextBox_DblClick()

PaintBox.CurrentX = TextBox.Left
PaintBox.CurrentY = TextBox.Top
PaintBox.Print TextBox.Text
TextBox.Visible = False
TextBox.Text = ""
TextBox.Height = 20
TextBox.Width = 8
PaintBox.SetFocus
PicChanged = True

SetImageBuffer

End Sub

Private Sub TextBox_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    TextBox_DblClick
Else
    TextLabel.Caption = TextBox.Text & "M"
    TextBox.Width = TextLabel.Width
End If

End Sub

Private Sub TextBox_KeyUp(KeyCode As Integer, Shift As Integer)

With TextBox
    Select Case KeyCode
    Case vbKeyEscape
        .Text = ""
        .Height = 20
        .Width = 8
        .Visible = False
    Case Else
        TextLabel.Caption = .Text & "O"
        .Width = TextLabel.Width
    End Select
End With

End Sub

Private Sub Timer1_Timer()

If Clipboard.GetFormat(vbCFBitmap) Then
    ClipboardPic.Picture = Clipboard.GetData(vbCFBitmap)
    EditPaste.Enabled = True
Else
    EditPaste.Enabled = False
End If

End Sub

Private Sub Tools_Click(Index As Integer)

On Error Resume Next

PaintBox.DrawMode = vbCopyPen

If DPolygon Then
    DrawPolygon Complete:=False
    DrawPolygon
    DPolygon = False
End If

If FImgZoom = True Then
    PaintBox.ScaleMode = 1
    FImgZoom = False
    PaintBox.Height = ImgZoom2.Height
    PaintBox.Width = ImgZoom2.Width
    PaintBox.Picture = ImgZoom2.Image
    PaintBox.ScaleMode = 3
    Form_Resize
    PlaceResizeBox
End If

If SelectionBox.Visible Then
    PaintBox.PaintPicture SelectionBox.Image, SelectionBox.Left, SelectionBox.Top, SelectionBox.Width, SelectionBox.Height
    EditCut.Enabled = False
    EditCopy.Enabled = False
    EditDel.Enabled = False
    EditCrop.Enabled = False
    SelectionBox.Visible = False
    SelectionBox.Cls
    ImageSelector.Visible = False
    PicChanged = True
    SetImageBuffer
End If

Select Case Tools(Index).Index
    Case 4, 5, 6, 8, 9, 15
        BrushFrame.Visible = False
        FillFrame.Visible = False
        SelectionFrame.Visible = False
    Case 10, 11, 12, 13
        BrushFrame.Visible = False
        FillFrame.Visible = True
        SelectionFrame.Visible = False
    Case 0, 14
        SelectionFrame.Visible = True
        BrushFrame.Visible = False
        FillFrame.Visible = False
    Case 7
        BrushFrame.Visible = True
        FillFrame.Visible = False
        SelectionFrame.Visible = False
    Case Else
        BrushFrame.Visible = False
        FillFrame.Visible = False
        SelectionFrame.Visible = False
End Select

CurInd = Tools(Index).Index

ChangeCursor

End Sub

Private Sub Tools_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

stbl = True

If stbl = True Then
    Select Case Tools(Index).Index
    Case 0
        StatusBar1.Panels(1).Text = "Selects a rectangular part of the picture to move, copy or edit."
    Case 1
        StatusBar1.Panels(1).Text = "Changes the magnification"
    Case 2
        StatusBar1.Panels(1).Text = "Picks up the color from the picture for drawing."
    Case 3
        StatusBar1.Panels(1).Text = "Fill an area with current drawing color."
    Case 4
        StatusBar1.Panels(1).Text = "Draws a free-form line one pixel wide."
    Case 5
        StatusBar1.Panels(1).Text = "Erases a portion of the picture, using the selected eraser shape."
    Case 6
        StatusBar1.Panels(1).Text = "Draws using an air-brush of the selected size."
    Case 7
        StatusBar1.Panels(1).Text = "Draws using a brush with the selected shape and size."
    Case 8
        StatusBar1.Panels(1).Text = "Draws a straight line with the selected line width."
    Case 9
        StatusBar1.Panels(1).Text = "Draws an arrow with the selected arrow width."
    Case 10
        StatusBar1.Panels(1).Text = "Draws a rectangle with the selected fill style."
    Case 11
        StatusBar1.Panels(1).Text = "Draws a rownded rectangle with the selected fill style."
    Case 12
        StatusBar1.Panels(1).Text = "Draws an ellipse with the selected fill style."
    Case 13
        StatusBar1.Panels(1).Text = "Draws a polygon with the selected fill style."
    Case 14
        StatusBar1.Panels(1).Text = "Inserts text into the picture."
    Case 15
        StatusBar1.Panels(1).Text = "Draws a curved line with the selected line width."
    End Select
stbl = False
End If

End Sub

Private Sub VScroll_Change()

Dim PaintBoxTop As Long
  
PaintBoxTop = -(CLng(VScroll.Value))
PaintBox.Top = Val(PaintBoxTop)

PlaceResizeBox

End Sub

Public Sub PlaceResizeBox()

VResize.Left = PaintBox.Left + (PaintBox.Width / 2)
VResize.Top = PaintBox.Top + PaintBox.Height

HResize.Left = PaintBox.Left + PaintBox.Width
HResize.Top = PaintBox.Top + (PaintBox.Height / 2)

BResize.Left = HResize.Left
BResize.Top = VResize.Top

End Sub

Private Sub VResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

VMD = True
ry = Val(y)

End Sub

Private Sub VResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

If VMD = True Then
    If PaintBox.Height >= 300 Then
        PaintBox.Height = PaintBox.Height + Val(ry) + Val(y)
    End If

    If PaintBox.Height < 300 Then
        PaintBox.Height = 300
    End If
End If

Form_Resize

PlaceResizeBox

End Sub

Private Sub VResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

VMD = False

End Sub

Private Sub BResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

BMD = True
rx = Val(x)
ry = Val(y)

End Sub

Private Sub BResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error Resume Next

If BMD = True Then
    If PaintBox.Height >= 300 And PaintBox.Width >= 300 Then
        PaintBox.Height = Val(PaintBox.Height) + Val(ry) + y
        PaintBox.Width = Val(PaintBox.Width) + Val(rx) + x
    End If

    If PaintBox.Height < 300 Then
        PaintBox.Height = 300
    End If
    If PaintBox.Width < 300 Then
        PaintBox.Width = 300
    End If
End If

Form_Resize

PlaceResizeBox

End Sub

Private Sub BResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

BMD = False

End Sub

Public Sub DrawArrow(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)

Dim Sign As Integer
Dim X3 As Integer
Dim Y3 As Integer
Dim X4 As Integer
Dim Y4 As Integer
Dim Beta As Single

Const ArrowTip = 45
Const TipLen = 10
Const PI = 3.14159265358979
  
PaintBox.Line (X1, Y1)-(X2, Y2)
  
If X2 - X1 <> 0 Then
    Beta = Atn((Y2 - Y1) / (X2 - X1)) * 180 / PI
Else
    Beta = 90
End If
  
If X2 > X1 Then
    Sign = 1
ElseIf X2 < X1 Then
    Sign = -1
ElseIf Y2 > Y1 Then
    Sign = 1
ElseIf Y2 < Y1 Then
    Sign = -1
End If

X3 = X2 - ((TipLen * Cos((ArrowTip + Beta) * PI / 180)) * Sign)
Y3 = Y2 - ((TipLen * Sin((ArrowTip + Beta) * PI / 180)) * Sign)
X4 = X2 - ((TipLen * Cos((ArrowTip - Beta) * PI / 180)) * Sign)
Y4 = Y2 + ((TipLen * Sin((ArrowTip - Beta) * PI / 180)) * Sign)
  
PaintBox.Line (X2, Y2)-(X3, Y3)
PaintBox.Line (X2, Y2)-(X4, Y4)

End Sub

Public Sub DrawSpray(ByVal x As Integer, ByVal y As Integer, ByVal R As Integer)

Dim DrwWidth As Integer

Const Intencity = 0.25

With PaintBox
    DrwWidth = .DrawWidth
    .DrawWidth = 1
    Randomize
    For i = 1 To ((R * R) * Intencity)
        PaintBox.PSet (x - (R / 2) + (Rnd() * R), y - (R / 2) + (Rnd() * R))
    Next
    .DrawWidth = DrwWidth
End With
  
End Sub

Public Sub DrawBrush(BrushShape As EBrushShape, ByVal x As Single, ByVal y As Single)

Dim i As Integer
Dim DrwWidth As Integer

Const BrushSize = 2
  
With PaintBox
    DrwWidth = .DrawWidth
    Select Case BrushShape
    Case FilledRect
        PaintBox.FillStyle = 0
        PaintBox.Line (x - (BrushSize * DrwWidth), y - (BrushSize * DrwWidth))-(x + (BrushSize * DrwWidth), y + (BrushSize * DrwWidth)), , BF
    Case FilledCircle
        PaintBox.FillStyle = 0
        PaintBox.FillColor = vbBlack
        PaintBox.Circle (x, y), BrushSize * DrwWidth
    Case SRect
        PaintBox.FillStyle = 1
        PaintBox.Line (x - (BrushSize * DrwWidth), y - (BrushSize * DrwWidth))-(x + (BrushSize * DrwWidth), y + (BrushSize * DrwWidth)), , B
    Case SCircle
        PaintBox.FillStyle = 1
        PaintBox.Circle (x, y), BrushSize * DrwWidth
    Case Cross
        PaintBox.Line (x - (BrushSize * DrwWidth), y)-(x + (BrushSize * DrwWidth), y)
        PaintBox.Line (x, y - (BrushSize * DrwWidth))-(x, y + (BrushSize * DrwWidth))
    Case DiagonalCross
        PaintBox.Line (x - (BrushSize * DrwWidth), y + (BrushSize * DrwWidth))-(x + (BrushSize * DrwWidth), y - (BrushSize * DrwWidth))
        PaintBox.Line (x - (BrushSize * DrwWidth), y - (BrushSize * DrwWidth))-(x + (BrushSize * DrwWidth), y + (BrushSize * DrwWidth))
    Case UpwardDiagonal
        PaintBox.Line (x - (BrushSize * DrwWidth), y + (BrushSize * DrwWidth))-(x + (BrushSize * DrwWidth), y - (BrushSize * DrwWidth))
    Case DownwardDiagonal
        PaintBox.Line (x - (BrushSize * DrwWidth), y - (BrushSize * DrwWidth))-(x + (BrushSize * DrwWidth), y + (BrushSize * DrwWidth))
    Case Horizontal
        PaintBox.Line (x - (BrushSize * DrwWidth), y)-(x + (BrushSize * DrwWidth), y)
    Case Vertical
        PaintBox.Line (x, y - (BrushSize * DrwWidth))-(x, y + (BrushSize * DrwWidth))
    End Select
End With
  
End Sub

Public Sub DrawPolygon(Optional Complete As Boolean = True, Optional OnlyDrawLastLine = True)

On Error Resume Next

Dim i As Integer
  
With PaintBox
    If Complete Then
        .DrawMode = vbCopyPen
        Select Case SFillStyle
        Case BorderOnly
            .FillStyle = 1
            .ForeColor = ForeColorLabel.BackColor
        Case BorderFill
            .FillStyle = 0
            .ForeColor = ForeColorLabel.BackColor
            .FillColor = FillColorLabel.BackColor
        Case FillOnly
            .FillStyle = 0
            .ForeColor = FillColorLabel.BackColor
            .FillColor = FillColorLabel.BackColor
        End Select
        Polygon PaintBox.hdc, PolygonLen(0), UBound(PolygonLen) + 1
        .Refresh
    Else
        If UBound(PolygonLen) > 0 Then
            If OnlyDrawLastLine Then
                PaintBox.Line (PolygonLen(UBound(PolygonLen) - 1).x, PolygonLen(UBound(PolygonLen) - 1).y)-(PolygonLen(UBound(PolygonLen)).x, PolygonLen(UBound(PolygonLen)).y)
            Else
                For i = 1 To UBound(PolygonLen)
                    PaintBox.Line (PolygonLen(i - 1).x, PolygonLen(i - 1).y)-(PolygonLen(i).x, PolygonLen(i).y)
                Next
            End If
        End If
    End If
End With

End Sub

Private Sub ChangeCursor()
                                                
With PaintBox
    .MousePointer = vbCustom
    Select Case CurInd
    Case 0
        .MousePointer = vbDefault
    Case 1
        .MouseIcon = CursorsList.ListImages(12).Picture
    Case 2
        .MouseIcon = CursorsList.ListImages(10).Picture
    Case 3
        .MouseIcon = CursorsList.ListImages(5).Picture
    Case 4
        .MouseIcon = CursorsList.ListImages(9).Picture
    Case 5
        .MouseIcon = CursorsList.ListImages(4).Picture
    Case 6
        .MouseIcon = CursorsList.ListImages(1).Picture
    Case 7
        .MouseIcon = CursorsList.ListImages(2).Picture
    Case 8, 9, 10, 11, 12, 13, 15
        .MouseIcon = CursorsList.ListImages(3).Picture
    Case 14
        .MouseIcon = CursorsList.ListImages(11).Picture
    End Select
End With

End Sub

Public Sub SaveChanges()

SMsg = MsgBox("The picture in " & FName & " has change." & Chr(13) & " Do you want to save changes?", vbYesNoCancel + vbExclamation, "AtrociousPaint")

End Sub

Public Sub OpenFile()

On Error GoTo OpenErr
    OpenDialog.CancelError = True
    OpenDialog.Filter = "Ahsan's Picture Files|*.abmp| All Files|*.*"
    OpenDialog.FilterIndex = 9
    OpenDialog.ShowOpen
    FName = OpenDialog.FileTitle
    PaintBox.Picture = LoadPicture(OpenDialog.FileName)
    PaintForm.Caption = FName & " - AtrociousPaint"
    PicChanged = False
    SC = True
Exit Sub

OpenErr:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox "Unknown error occured while opening file.", vbCritical, "AtrociousPaint"
    End If

End Sub

Public Sub SetWallpaper()

Dim keyValLen As String
Dim rtn As Long
Dim KeyName As String
Dim hKey As Long
Dim KeyValueLength As Long
Dim path As String

path = OpenDialog.FileName
keyValLen = 0
KeyName = "Desktop\TileWallpaper"
KeyValueLength = Len(keyValLen) + 1
rtn = RegOpenKey(HKEY_CURRENT_USER, "Control Panel\Desktop", hKey)
rtn = RegSetValueEx(hKey, "TileWallpaper", 0, REG_SZ, keyValLen, KeyValueLength)
rtn = RegCloseKey(hKey)
rtn = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, path, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
MsgBox "Wallpaper Applied!", vbInformation, "AtrociousPaint"

End Sub

Private Sub ImageRotate(picSource As PictureBox, picDestination As PictureBox, sngRotateAngle As Single, blnClockWise As Boolean)
  
Const conPi = 3.14159265358979
  
Dim a As Single
Dim intMaxXY As Single
Dim dXs As Long
Dim dYs As Long
Dim dXd As Long
Dim dYd As Long
Dim lngAdjustX As Long
Dim lngAdjustY As Long
Dim lngColor(3) As Long
Dim R As Integer
Dim Xs As Integer
Dim Ys As Integer
Dim Xd As Integer
Dim Yd As Integer
                              
If blnClockWise Then
    sngRotateAngle = 360 - sngRotateAngle
End If
  
Xs = picSource.ScaleWidth / 2
Ys = picSource.ScaleHeight / 2
Xd = picDestination.ScaleWidth / 2
Yd = picDestination.ScaleHeight / 2
intMaxXY = IIf(picDestination.ScaleWidth > picDestination.ScaleHeight, picDestination.ScaleWidth / 2, picDestination.ScaleHeight / 2)

If (sngRotateAngle = 90) Or (sngRotateAngle = 270) Then
    lngAdjustX = ((picDestination.ScaleHeight - _
                   picDestination.ScaleWidth) / 2) - 2
    lngAdjustY = ((picDestination.ScaleWidth - _
                   picDestination.ScaleHeight) / 2)
    picDestination.Tag = CStr(picDestination.Width)
    picDestination.Width = picDestination.Height
    picDestination.Height = CLng(picDestination.Tag)
    With PaintForm
      .PlaceResizeBox
      .Form_Resize
      .Refresh
    End With
Else
    lngAdjustX = 0
    lngAdjustY = 0
End If

sngRotateAngle = sngRotateAngle * (conPi / 180)
picDestination.DrawMode = vbCopyPen

For dXd = 0 To intMaxXY
    ProgressBar.Visible = True
    ProgressBar.Value = (dXd / (intMaxXY + 1)) * 100
    For dYd = 0 To intMaxXY
      If dXd = 0 Then
        a = conPi / 2
      Else
        a = Atn(dYd / dXd)
      End If
      R = Sqr((dXd * dXd) + (dYd * dYd))
      dXs = R * Cos(a + sngRotateAngle)
      dYs = R * Sin(a + sngRotateAngle)

      lngColor(0) = GetPixel(picSource.hdc, Xs + dXs, Ys + dYs)
      lngColor(1) = GetPixel(picSource.hdc, Xs - dXs, Ys - dYs)
      lngColor(2) = GetPixel(picSource.hdc, Xs + dYs, Ys - dXs)
      lngColor(3) = GetPixel(picSource.hdc, Xs - dYs, Ys + dXs)

      If lngColor(0) <> -1 Then
        SetPixel picDestination.hdc, Xd + dXd + lngAdjustX, _
                 Yd + dYd + lngAdjustY, lngColor(0)
      End If
      If lngColor(1) <> -1 Then
        SetPixel picDestination.hdc, Xd - dXd + lngAdjustX, _
                 Yd - dYd + lngAdjustY, lngColor(1)
      End If
      If lngColor(2) <> -1 Then
        SetPixel picDestination.hdc, Xd + dYd + lngAdjustX, _
                 Yd - dXd + lngAdjustY, lngColor(2)
      End If
      If lngColor(3) <> -1 Then
        SetPixel picDestination.hdc, Xd - dYd + lngAdjustX, _
                 Yd + dXd + lngAdjustY, lngColor(3)
      End If
    Next
    picDestination.Refresh
    PaintForm.Enabled = False
    StatusBar1.Panels(2).Text = "Rotating....."
Next
picDestination.Refresh
StatusBar1.Panels(2).Text = ""
ProgressBar.Value = 0
ProgressBar.Visible = False
PaintForm.Enabled = True
  
End Sub
  
Public Sub SetImageBuffer()

If CurBuf < MaxBuf Then
    CurBuf = CurBuf + 1
Else
    CurBuf = 0
End If

If CurBuf > BufferPic.UBound Then
    Load BufferPic(CurBuf)
End If
  
BufferPic(CurBuf).Picture = PaintBox.Image
BufferPic(CurBuf).Tag = CStr((PaintBox.Width * 100000) + PaintBox.Height)
EndBuf = CurBuf

If StartBuf = EndBuf Then
    If StartBuf < MaxBuf Then
      StartBuf = StartBuf + 1
    Else
      StartBuf = 0
    End If
End If

PicChanged = True
EditUndo.Enabled = True
EditRedo.Enabled = False

End Sub

Private Sub ClearImageBuffer()

Dim i As Integer
  
CurBuf = 0
StartBuf = 0
EndBuf = 0

For i = 1 To BufferPic.UBound
    Unload BufferPic(i)
Next

BufferPic(CurBuf).Picture = PaintBox.Image
BufferPic(CurBuf).Tag = CStr((PaintBox.Width * 100000) + PaintBox.Height)
EditUndo.Enabled = False
EditRedo.Enabled = False

End Sub
