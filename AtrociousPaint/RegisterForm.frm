VERSION 5.00
Begin VB.Form RegisterForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AtrociousPaint"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "RegisterForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Enter Serial Number"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2640
         MaxLength       =   5
         PasswordChar    =   "*"
         TabIndex        =   2
         Tag             =   "11111"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         MaxLength       =   5
         PasswordChar    =   "*"
         TabIndex        =   1
         Tag             =   "67890"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   5
         PasswordChar    =   "*"
         TabIndex        =   0
         Tag             =   "12345"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS

Public Function FixPath(lzpath As String)

If Right$(lzpath, 1) = "\" Then FixPath = lzpath Else FixPath = lzpath & "\"
    
End Function

Private Sub Command1_Click()

Dim CName
Dim ans
Dim IconPath As String, ProgPath As String
Dim FSO As New FileSystemObject
Dim TSO As TextStream
Dim APath

If WS = 2 Then
    Unload Me
End If

If Text1.Text = Text1.Tag And Text2.Text = Text2.Tag And Text3.Text = Text3.Tag Then
    Unload Me
    CName = InputBox("Please Enter your name.", "AtrociousPaint")
    APath = FixPath(App.path) & "APaintR.dll"
    Set TSO = FSO.CreateTextFile(APath)
    TSO.WriteLine " ***** Do Not Change Any Text ***** "
    TSO.WriteLine "Registered"
    TSO.WriteLine CName
    IconPath = FixPath(App.path) & "AtrociousPaint.ico"
    ProgPath = FixPath(App.path) & "AtrociousPaint.exe"
    SaveKey HKEY_CLASSES_ROOT, ".abmp"
    SaveKey HKEY_CLASSES_ROOT, ".abmp\DefaultIcon"
    SaveKey HKEY_CLASSES_ROOT, ".abmp\shell"
    SaveKey HKEY_CLASSES_ROOT, ".abmp\shell\open"
    SaveKey HKEY_CLASSES_ROOT, ".abmp\shell\open\command"
    SaveString HKEY_CLASSES_ROOT, ".abmp\DefaultIcon", "", IconPath
    SaveString HKEY_CLASSES_ROOT, ".abmp\shell\open\command", "", Chr(34) & ProgPath & Chr(34) & " %1"
    MsgBox "Apllication Registered Successfully.", vbInformation, "AtrociousPaint"
    Load SplashForm
    SplashForm.Show
Else
    WS = WS + 1
    MsgBox "Wrong Serial Number.", vbCritical, "AtrociousPaint"
    If WS <= 2 Then
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text1.SetFocus
    End If
End If

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

On Error Resume Next

Dim FSO As New FileSystemObject
Dim TSO As TextStream
Dim APath
Dim TSORead

APath = FixPath(App.path) & "APaintR.dll"

Set TSO = FSO.OpenTextFile(APath)

d = TSO.ReadLine
d = TSO.Read(10)

If d = "Registered" Then
    Unload Me
    Load SplashForm
    SplashForm.Show
End If

End Sub

Private Sub Text1_Change()

If Len(Text1.Text) = Text1.MaxLength Then
    Text2.SetFocus
End If

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

F1 KeyCode

End Sub

Private Sub Text2_Change()

If Len(Text2.Text) = Text2.MaxLength Then
    Text3.SetFocus
End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

F1 KeyCode

End Sub

Private Sub Text3_Change()

If Len(Text3.Text) = Text3.MaxLength Then
    Command1.SetFocus
End If

End Sub

Private Function F1(ByVal K As Integer)

If K = vbKeyDown Then SendKeys "{TAB}"
If K = vbKeyUp Then SendKeys "+{TAB}"
If K = vbKeyReturn Then SendKeys "{TAB}"
If K = 27 Then Unload RegisterForm

End Function

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)

F1 KeyCode

End Sub
