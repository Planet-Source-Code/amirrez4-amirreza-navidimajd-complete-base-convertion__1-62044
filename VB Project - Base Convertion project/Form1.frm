VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Base Converter"
   ClientHeight    =   4590
   ClientLeft      =   1980
   ClientTop       =   3270
   ClientWidth     =   4875
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOperations 
      Caption         =   "Operations"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Timer TimCheck 
      Interval        =   100
      Left            =   4440
      Top             =   240
   End
   Begin VB.TextBox TxtOutput 
      Height          =   285
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3000
      Width           =   3255
   End
   Begin VB.TextBox TxtInput 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Frame FraOutput 
      Caption         =   "Output Type :"
      Height          =   1935
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   1935
      Begin VB.OptionButton OptOutput 
         Caption         =   "HexaDecimal"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton OptOutput 
         Caption         =   "Decimal"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptOutput 
         Caption         =   "Octal"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptOutput 
         Caption         =   "Binary"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame FraInput 
      Caption         =   "Input Type :"
      Height          =   1935
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1935
      Begin VB.OptionButton OptInput 
         Caption         =   "HexaDecimal"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton OptInput 
         Caption         =   "Decimal"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton OptInput 
         Caption         =   "Octal"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton OptInput 
         Caption         =   "Binary"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Coded By Amirreza Navidi"
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   4080
      Width           =   4170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label LblInput 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Input :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   570
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private LastOutPut As Integer

Private Sub Form_Load()
Load FrmOperations
FrmOperations.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload FrmOperations
End Sub

Private Sub OptInput_Click(Index As Integer)
    TxtInput.Text = ""
    TxtInput.Text = ""
End Sub

Private Sub OptOutput_Click(Index As Integer)
    TxtInput.Text = ""
    TxtInput.Text = ""
End Sub

Private Sub TimCheck_Timer()
    Dim TInputItem As Integer
    TInputItem = InputItem
    
    If LastOutPut <> TInputItem Then
        OptOutput(LastOutPut).Enabled = True
    End If
    
    OptOutput(TInputItem).Enabled = False
    If OptOutput(TInputItem) Then
        If TInputItem <= 2 Then
            OptOutput(TInputItem + 1).Value = True
        Else
            OptOutput(0).Value = True
        End If
    End If
    
    LastOutPut = TInputItem
End Sub

Private Sub TxtInput_Change()
On Error GoTo Trap
    TxtOutput.Text = ConvertAll(TxtInput.Text, GetBase(InputItem), GetBase(OutputItem))
    Exit Sub
Trap:
    If Err.Number = 6 Then
        TxtInput.Text = Left(TxtInput, Len(TxtInput) - 1)
        TxtInput.SelStart = Len(TxtInput)
        TxtInput.MaxLength = Len(TxtInput)
    End If
    
    If TxtInput = "" Then TxtOutput = ""
End Sub

Private Sub TxtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 47 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If Not IsValidChar(GetBase(InputItem), Chr(KeyAscii)) Then KeyAscii = 0: Beep
    End If
End Sub

Private Function InputItem() As Integer
    Dim I As Integer
    For I = OptInput.LBound To OptInput.UBound - 1
        If OptInput(I).Value = True Then Exit For
    Next
    InputItem = I
End Function

Private Function OutputItem() As Integer
    Dim I As Integer
    For I = OptOutput.LBound To OptOutput.UBound - 1
        If OptOutput(I).Value = True Then Exit For
    Next
    OutputItem = I
End Function
