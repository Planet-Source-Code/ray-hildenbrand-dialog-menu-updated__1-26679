VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{289B7027-55BE-4254-B64E-26E0841EF0FA}#1.0#0"; "DialogMenu.ocx"
Begin VB.Form frmTesting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Menu"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8745
   StartUpPosition =   2  'CenterScreen
   Begin vbpDialogMenu.DialogMenu DialogMenu1 
      Height          =   5145
      Left            =   30
      TabIndex        =   11
      Top             =   0
      Width           =   2040
      _ExtentX        =   2937
      _ExtentY        =   9075
      ShowScrollBars  =   0   'False
      BackColor       =   14737632
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Caption"
   End
   Begin VB.PictureBox picIcons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   1140
      Picture         =   "frmTest.frx":0ECA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1365
      Width           =   480
   End
   Begin VB.PictureBox picIcons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   1140
      Picture         =   "frmTest.frx":21C4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   825
      Width           =   480
   End
   Begin VB.PictureBox picIcons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   1140
      Picture         =   "frmTest.frx":2A8E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   240
      Width           =   480
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6945
      Left            =   1710
      ScaleHeight     =   6945
      ScaleWidth      =   7080
      TabIndex        =   3
      Top             =   -75
      Width           =   7080
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change Header Fore Color"
         Height          =   375
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3585
         Width           =   2235
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Toggle Header"
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3195
         Width           =   2235
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change ForeColor"
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1545
         Width           =   2235
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Change BackColor"
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1965
         Width           =   2235
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   630
         Top             =   2400
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Image From File"
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2400
         Width           =   2235
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Toggle ScrollBars"
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2805
         Width           =   2235
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   465
         X2              =   5805
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Image picHeader 
         Height          =   585
         Left            =   6060
         Top             =   225
         Width           =   645
      End
      Begin VB.Label lblHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "Organize"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   450
         TabIndex        =   4
         Top             =   240
         Width           =   2370
      End
   End
   Begin VB.PictureBox picIcons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   735
      Picture         =   "frmTest.frx":5230
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   645
      Width           =   480
   End
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CountIs As Integer

Private Sub Command1_Click()
    cd.FileName = ""
    cd.Filter = "Icon Files (*.ico)|*.ico"
    cd.ShowOpen
    If Not cd.FileName = "" Then
        CountIs = CountIs + 1
        DialogMenu1.AddButtonFromFile "Caption " & CountIs, cd.FileName, "Button " & CountIs & " tooltip"
        
    End If
End Sub



Private Sub Command2_Click()
    DialogMenu1.ShowScrollBars = Not DialogMenu1.ShowScrollBars
End Sub

Private Sub Command3_Click()
    cd.Color = DialogMenu1.BackColor
    
    cd.ShowColor
    If Not cd.Color = DialogMenu1.BackColor Then
        DialogMenu1.BackColor = cd.Color
    End If
End Sub

Private Sub Command4_Click()
    cd.Color = DialogMenu1.ForeColor
    
    cd.ShowColor
    If Not cd.Color = DialogMenu1.ForeColor Then
        DialogMenu1.ForeColor = cd.Color
    End If
End Sub

Private Sub Command5_Click()
    DialogMenu1.HeaderVisible = Not DialogMenu1.HeaderVisible
End Sub

Private Sub Command6_Click()
    cd.Color = DialogMenu1.HeaderForeColor
    
    cd.ShowColor
    If Not cd.Color = DialogMenu1.HeaderForeColor Then
        DialogMenu1.HeaderForeColor = cd.Color
    End If
End Sub

Private Sub DialogMenu1_ButtonClick(Index As Integer, KeyIs As String)
    Set picHeader.Picture = DialogMenu1.GetActiveButtonsPicture
    lblHeader.Caption = KeyIs
    
    
End Sub

Private Sub Form_Load()
    
    DialogMenu1.iList.Create
    DialogMenu1.iList.ColourDepth = ILC_COLOR24
    Dim mFont As New StdFont
    mFont.Name = "tahoma"
    mFont.Size = 8
    Set DialogMenu1.Font = mFont
    
    DialogMenu1.ForeColor = vbWhite
    DialogMenu1.AddButtonFromSTDPicture picIcons(0).Picture, "My Computer"
    DialogMenu1.AddButtonFromSTDPicture picIcons(1).Picture, "My Pictures"
    DialogMenu1.AddButtonFromSTDPicture picIcons(2).Picture, "Organize"
    DialogMenu1.AddButtonFromSTDPicture picIcons(3).Picture, "Calendar"
    CountIs = 4
    
    DialogMenu1.ForceClick 0
    
    
End Sub

