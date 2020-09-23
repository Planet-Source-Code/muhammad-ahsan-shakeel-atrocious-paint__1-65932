VERSION 5.00
Begin VB.Form SplashForm 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4800
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SplashForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6480
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   0
         Top             =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   960
         X2              =   5640
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Copyright1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Atrocious Softwares Corporation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   1920
         TabIndex        =   5
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Image Logo 
         Height          =   825
         Left            =   960
         Picture         =   "SplashForm.frx":2372
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label CopyrightLabel 
         BackColor       =   &H00000000&
         Caption         =   "Copyright © 2005-2006"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Version 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   " Version 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   4230
         TabIndex        =   3
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label ProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "AtrociousPaint"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   675
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   3285
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: This Program is protected under Copyright Law. Unautharised usage of this Software is  strictly forbidden"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1080
         TabIndex        =   1
         Top             =   3240
         Width           =   4455
      End
   End
End
Attribute VB_Name = "SplashForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim LEVEL As Byte
LEVEL = 175

'MsgBox "This program will run in 800 x 600 resolution." & Chr(13) & " It will change resolution automatically.", vbInformation, "AtrociousPaint"

ChangeRes 800, 600

SetWindowRgn hwnd, CreateEllipticRgn(3, 3, 450, 275), True

SetWindowLong Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED

SetLayeredWindowAttributes Me.hwnd, 0, LEVEL, LWA_ALPHA

End Sub

Private Sub Timer1_Timer()

Unload Me
Load PaintForm
PaintForm.Show

End Sub



