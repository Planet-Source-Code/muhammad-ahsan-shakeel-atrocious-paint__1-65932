VERSION 5.00
Begin VB.Form AboutPaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About AtrociousPaint"
   ClientHeight    =   3585
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "AboutPaint.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2474.431
   ScaleMode       =   0  'User
   ScaleWidth      =   5197.651
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   5295
      Begin VB.Label LicensedTo 
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   1200
         TabIndex        =   25
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label lblLicensedTo 
         Caption         =   "This Program is licensed to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   3360
      Top             =   120
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   2400
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4200
      TabIndex        =   1
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   13
      Left            =   2805
      TabIndex        =   22
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   120
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   12
      Left            =   2670
      TabIndex        =   21
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   150
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   11
      Left            =   2595
      TabIndex        =   20
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   60
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   10
      Left            =   2415
      TabIndex        =   19
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   150
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   9
      Left            =   2280
      TabIndex        =   18
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   150
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   8
      Left            =   2040
      TabIndex        =   17
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   135
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   7
      Left            =   1890
      TabIndex        =   16
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   150
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   6
      Left            =   1740
      TabIndex        =   15
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   150
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   5
      Left            =   1680
      TabIndex        =   14
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   60
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   4
      Left            =   1530
      TabIndex        =   13
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   150
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   3
      Left            =   1395
      TabIndex        =   12
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   150
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   2
      Left            =   1290
      TabIndex        =   11
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   135
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   1
      Left            =   1155
      TabIndex        =   10
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   120
   End
   Begin VB.Label Developer 
      Alignment       =   2  'Center
      Caption         =   "Muhammad Ahsan Shakeel"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   9
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Version 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "1.0.0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblDeveloper 
      Caption         =   "Developed By:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   7
      Top             =   960
      WhatsThisHelpID =   10383
      Width           =   1365
   End
   Begin VB.Label ProductName 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   0
      Left            =   960
      TabIndex        =   6
      Top             =   120
      WhatsThisHelpID =   10382
      Width           =   195
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "AboutPaint.frx":2372
      Stretch         =   -1  'True
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Warning 
      Caption         =   $"AboutPaint.frx":46E4
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   3795
   End
   Begin VB.Label lblWarning 
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
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1290
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5084.965
      Y1              =   911.088
      Y2              =   911.088
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   885
   End
End
Attribute VB_Name = "AboutPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim FSO As New FileSystemObject
Dim TSO As TextStream
Dim APath

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Private Sub cmdSysInfo_Click()

Call StartSysInfo

End Sub

Private Sub cmdOK_Click()

Unload Me

End Sub

Public Sub StartSysInfo()

On Error GoTo SysInfoErr
  
Dim rc As Long
Dim SysInfoPath As String

If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
    If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
    Else
        GoTo SysInfoErr
    End If
Else
    GoTo SysInfoErr
End If
    
Call Shell(SysInfoPath, vbMaximizedFocus)
    
Exit Sub

SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly

End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean

Dim i As Long
Dim rc As Long
Dim hKey As Long
Dim hDepth As Long
Dim KeyValType As Long
Dim tmpVal As String
Dim KeyValSize As Long
    
rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
tmpVal = String$(1024, 0)
KeyValSize = 1024
rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)
                        
If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError

If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
    tmpVal = Left(tmpVal, KeyValSize - 1)
Else
    tmpVal = Left(tmpVal, KeyValSize)
End If

Select Case KeyValType
Case REG_SZ
    KeyVal = tmpVal
Case REG_DWORD
    For i = Len(tmpVal) To 1 Step -1
        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
    Next
    KeyVal = Format$("&h" + KeyVal)
End Select
    
GetKeyValue = True
rc = RegCloseKey(hKey)

Exit Function
    
GetKeyError:
    KeyVal = ""
    GetKeyValue = False
    rc = RegCloseKey(hKey)

End Function

Private Sub Form_Load()

PaintForm.Enabled = False

APath = FixPath(App.path) & "APaintR.dll"

Set TSO = FSO.OpenTextFile(APath, 1, False, TristateFalse)

TSO.SkipLine

TSO.SkipLine

LicensedTo.Caption = TSO.ReadLine

End Sub

Private Sub Form_Unload(Cancel As Integer)

PaintForm.Enabled = True

End Sub

Private Sub Timer1_Timer()

If a <= 14 Then
    If a <= 13 Then
        ProductName(a).ForeColor = &HC000&
    End If
    
    If a >= 1 Then
        ProductName(a - 1).ForeColor = &HFF&
    End If
End If
a = a + 1

If a = 15 Then
    a = 0
End If

End Sub

Public Function FixPath(lzpath As String)

If Right$(lzpath, 1) = "\" Then FixPath = lzpath Else FixPath = lzpath & "\"
    
End Function

