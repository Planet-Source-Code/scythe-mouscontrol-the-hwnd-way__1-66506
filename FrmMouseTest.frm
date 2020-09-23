VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMousTest 
   Caption         =   "MouseControl Testapp"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame FrmHit 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   840
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label Label1 
         Caption         =   "You hit the small Frame"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   5535
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1095
      Left            =   2280
      TabIndex        =   11
      Top             =   4080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1931
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Column1"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Column2"
         Object.Width           =   882
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   735
      Left            =   2280
      TabIndex        =   10
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   3720
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1 Index 0"
      Height          =   1215
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   2055
      Begin VB.Frame Frame1 
         Caption         =   "Frame1 Index 1"
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   2280
      ScaleHeight     =   1395
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   1815
   End
   Begin PrjMouseTest.MouseControl MouseControl1 
      Left            =   7560
      Top             =   4800
      _extentx        =   953
      _extenty        =   953
   End
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   4560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "FrmMousTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 MouseControl1.Init Me
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
MsgBox Button
End Sub

Private Sub MouseControl1_MouseMiddle(ControlName As String, Index As Long, ButtonDown As Boolean)
If ButtonDown Then
    Text1 = Text1 & "Middle Button Down for " & ControlName
Else
    Text1 = Text1 & "Middle Button Up for " & ControlName
End If
If Index <> -1 Then
 Text1 = Text1 & "(" & Index & ")"
End If
Text1 = Text1 & vbCrLf
End Sub

Private Sub MouseControl1_MouseOver(ControlName As String, Index As Long)
Text1 = Text1 & "MouseOver for " & ControlName
If Index <> -1 Then
 Text1 = Text1 & "(" & Index & ")"
End If
Text1 = Text1 & vbCrLf

'Here is how we use this
If ControlName = "Frame1" And Index = 1 Then
 FrmHit.Visible = True
End If
End Sub

Private Sub MouseControl1_MouseOut(ControlName As String, Index As Long)
Text1 = Text1 & "MouseOut for " & ControlName
If Index <> -1 Then
 Text1 = Text1 & "(" & Index & ")"
End If
Text1 = Text1 & vbCrLf

'Now we left the control so hide Our Frame
If ControlName = "Frame1" And Index = 1 Then
 FrmHit.Visible = False
End If
End Sub

Private Sub MouseControl1_MouseWheel(ControlName As String, Index As Long, ScrollUp As Boolean)
Dim A$

If ScrollUp Then
 A$ = "UP"
Else
 A$ = "Down"
End If

Text1 = Text1 & "MouseWheel " & A$ & " for " & ControlName
If Index <> -1 Then
 Text1 = Text1 & "(" & Index & ")"
End If
Text1 = Text1 & vbCrLf
End Sub
