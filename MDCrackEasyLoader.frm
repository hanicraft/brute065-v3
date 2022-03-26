VERSION 5.00
Begin VB.Form MDCrackEasyLoader 
   Caption         =   "brute065 v3 BETA"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CrackButton 
      Caption         =   "Crack"
      Height          =   372
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   3852
   End
   Begin VB.Frame HashBox 
      Caption         =   "Hash"
      Height          =   612
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3852
      Begin VB.TextBox HashField 
         Height          =   288
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3612
      End
   End
   Begin VB.Frame OptionsBox 
      Caption         =   "Options"
      Height          =   3012
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3852
      Begin VB.OptionButton NTLM1Option 
         Caption         =   "NTLM1 Hash"
         Height          =   252
         Left            =   2280
         TabIndex        =   13
         Top             =   2640
         Width           =   1332
      End
      Begin VB.OptionButton MD5Option 
         Caption         =   "MD5 Hash"
         Height          =   252
         Left            =   1200
         TabIndex        =   12
         Top             =   2640
         Width           =   1092
      End
      Begin VB.OptionButton MD4Option 
         Caption         =   "MD4 Hash"
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1092
      End
      Begin VB.TextBox CustomCharField 
         Height          =   288
         Left            =   360
         TabIndex        =   10
         Top             =   2280
         Width           =   3372
      End
      Begin VB.CheckBox CustomCharTick 
         Caption         =   "Custom charset:"
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   3612
      End
      Begin VB.CheckBox FindAllTick 
         Caption         =   "Find all posibilities"
         Height          =   252
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   3612
      End
      Begin VB.TextBox MinimalPasswordField 
         Height          =   288
         Left            =   2160
         TabIndex        =   7
         Top             =   1560
         Width           =   372
      End
      Begin VB.CheckBox MinimalPasswordTick 
         Caption         =   "Minimal Password size:"
         Height          =   252
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   3612
      End
      Begin VB.TextBox WriteHashField 
         Height          =   288
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   3372
      End
      Begin VB.CheckBox FastWriteTick 
         Caption         =   "Fast Write (Writes computed hash file faster)"
         Height          =   252
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   3612
      End
      Begin VB.CheckBox WriteHashTick 
         Caption         =   "Write Computed Hashes to this file:"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3612
      End
      Begin VB.CheckBox VerboseExtraTick 
         Caption         =   "Verbose Extra (Shows MD5 hashes)"
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3612
      End
      Begin VB.CheckBox VerboseTick 
         Caption         =   "Verbose (Doesn't seem to work if switched off)"
         Height          =   252
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   1  'Checked
         Width           =   3612
      End
   End
   Begin VB.Label EasyLoaderText 
      Alignment       =   2  'Center
      Caption         =   "brute065 by hanicraft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      Width           =   3852
   End
End
Attribute VB_Name = "MDCrackEasyLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CrackButton_Click()
Verbose = ""
WriteHash = ""
FastWrite = ""
MinPassword = ""
FindAll = ""
CustomChar = ""
Hash = ""
If VerboseTick.Value = 1 Then Verbose = "-v"
If VerboseExtraTick.Value = 1 Then Verbose = "-V"
If WriteHashTick.Value = 1 Then WriteHash = "-w " & WriteHashField.Text
If FastWriteTick.Value = 1 Then FastWrite = "-F"
If MinimalPasswordTick.Value = 1 Then MinPassword = "-S " & MinimalPasswordField.Text
If WarnValue(MinimalPasswordField.Text) = False Then Exit Sub
If FindAllTick.Value = 1 Then FindAll = "-a"
If CustomCharTick.Value = 1 Then CustomChar = CustomCharField.Text
If MD4Option = True Then Hash = "-M MD4"
If MD5Option = True Then Hash = "-M MD5"
If NTLM1Option = True Then Hash = "-M NTLM1"
Shell CurDir$ & "\mdcrack.exe" & " " & Verbose & " " & WriteHash & " " & FastWrite & " " & MinPassword & " " & FindAll & " " & CustomChar & " " & Hash & " " & HashField.Text
End Sub

Public Function WarnValue(minpass) As Boolean
If minpass < 8 Then
    WarnValue = True
    Exit Function
End If
If minpass > 7 Then
Dim Warn As Integer
Warn = MsgBox("Warning passwords longer than 8 characters will take a long time to crack, are you sure you want to continue?", vbYesNo + vbQuestion, "Please confirm:")
    If Warn = vbYes Then
        WarnValue = True
        Exit Function
    Else
        WarnValue = False
        Exit Function
    End If
End If
End Function

