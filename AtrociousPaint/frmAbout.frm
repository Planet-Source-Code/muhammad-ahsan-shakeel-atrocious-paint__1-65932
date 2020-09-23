VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Atrocious Paint"
   ClientHeight    =   2730
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5415
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   182
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -105
      TabIndex        =   3
      Top             =   2025
      WhatsThisHelpID =   10385
      Width           =   5550
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3945
      TabIndex        =   0
      Top             =   2250
      WhatsThisHelpID =   10379
      Width           =   1260
   End
   Begin VB.Label Label4 
      Caption         =   "Do NOT DISTRIBUTE this Software."
      Height          =   255
      Left            =   315
      TabIndex        =   6
      Top             =   1560
      Width           =   4785
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      Caption         =   "Warning:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   1155
      Width           =   4740
   End
   Begin VB.Label Label1 
      Caption         =   "Atrocious Paint"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   285
      WhatsThisHelpID =   10382
      Width           =   3885
   End
   Begin VB.Label lblCopyright 
      Caption         =   "Developed By:                Muhammad Ahsan Shakeel"
      Height          =   225
      Left            =   1440
      TabIndex        =   2
      Top             =   645
      WhatsThisHelpID =   10383
      Width           =   3885
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   255
      Picture         =   "frmAbout.frx":1042
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   255
      TabIndex        =   5
      Top             =   1065
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

