VERSION 5.00
Begin VB.UserControl DialogMenu 
   BackColor       =   &H009A9A9A&
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   2415
   ScaleWidth      =   1890
   Begin VB.PictureBox picSplit 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   1650
      TabIndex        =   8
      Top             =   2130
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H009A9A9A&
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   0
      ScaleHeight     =   150
      ScaleWidth      =   1890
      TabIndex        =   7
      Top             =   2265
      Width           =   1890
   End
   Begin VB.PictureBox picHeader 
      BackColor       =   &H009A9A9A&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   1620
      TabIndex        =   6
      Top             =   -15
      Visible         =   0   'False
      Width           =   1620
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   45
         TabIndex        =   9
         Top             =   75
         Width           =   1485
      End
   End
   Begin VB.PictureBox picDown 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1650
      MouseIcon       =   "DialogMenu.ctx":0000
      MousePointer    =   99  'Custom
      Picture         =   "DialogMenu.ctx":08CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   870
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picUp 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1650
      MouseIcon       =   "DialogMenu.ctx":0C0E
      MousePointer    =   99  'Custom
      Picture         =   "DialogMenu.ctx":14D8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1320
      Top             =   2910
   End
   Begin VB.Timer tmrMD 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2010
      Top             =   2760
   End
   Begin VB.PictureBox picScroll 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   30
      ScaleHeight     =   1455
      ScaleWidth      =   1620
      TabIndex        =   2
      Top             =   405
      Width           =   1620
      Begin VB.PictureBox picButtons 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Height          =   960
         Index           =   0
         Left            =   30
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   102
         TabIndex        =   3
         Top             =   465
         Visible         =   0   'False
         Width           =   1530
      End
   End
   Begin VB.PictureBox picPBBuffer 
      Height          =   1095
      Left            =   2100
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   1575
      Width           =   1515
   End
   Begin VB.PictureBox picCheck 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2100
      ScaleHeight     =   64
      ScaleMode       =   0  'User
      ScaleWidth      =   102
      TabIndex        =   0
      Top             =   240
      Width           =   1530
   End
End
Attribute VB_Name = "DialogMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event mousedown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute mousedown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Private mBackColor As OLE_COLOR
Private fColor As OLE_COLOR
Event ButtonClick(Index As Integer, KeyIs As String)
Private mousedown As Boolean
Private MScroll As Boolean
Private mShowHeader As Boolean
Private mHeadForeColor As OLE_COLOR
Private mCaption As String
Private AlreadyDrawing As Boolean
Private Sizing As Boolean
Private mY As Long
Private mX As Long
Public iList As dlgImgList
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    
    BackColor = mBackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    mBackColor = New_BackColor
    UserControl.BackColor() = mBackColor
    Dim i As Integer
    For i = 0 To picButtons.UBound
        picButtons(i).AutoRedraw = False
        picButtons(i).Cls
        picButtons(i).Refresh
        picButtons(i).AutoRedraw = True
        
        picButtons(i).BackColor = New_BackColor
        
    Next
    DrawImages
    picScroll.BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    fColor = New_ForeColor
    picCheck.ForeColor = New_ForeColor
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    DrawImages
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Set picCheck.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub







Private Sub picBottom_Click()
'    Dim i As Integer
'    If picBottom.MousePointer = 10 Then
'        picBottom.MousePointer = 5
'        For i = picBottom.tOp To picHeader.Height + picHeader.tOp + 15 Step -30
'                picBottom.Move 0, i
'                UserControl.Refresh
'        Next
'        Exit Sub
'    Else
'        For i = picBottom.tOp To UserControl.Height - picBottom.Height
'                picBottom.Move 0, i
'                UserControl.Refresh
'        Next
'        picBottom.MousePointer = 10
'        Exit Sub
'    End If
End Sub

Private Sub picBottom_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Sizing = True
'    mY = Y + picBottom.Top
'    picSplit.Move picBottom.Left, picBottom.Top, picBottom.Width, picBottom.Height
'    picSplit.Visible = True
'
'
End Sub

Private Sub picBottom_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
'    Sizing = False
'    picSplit.Visible = False
'    If mY > Y Then
'        If Not (mY) - Abs(Y) < picHeader.Height + picBottom.Height Then
'            UserControl.Height = (mY) - Abs(Y)
'        Else
'            picBottom.Top = picHeader.Height + picBottom.Height
'        End If
'    Else
'        UserControl.Height = Abs(Y)
'    End If
'    mY = 0
'    DrawHeader
'   ' UserControl.Height = Y
End Sub

Private Sub picButtons_Click(Index As Integer)
    'picButtons(Index).Tag = "1"
    DrawDownButton picButtons(Index), 64, 102
    Dim i As Integer
        For i = 0 To picButtons.UBound
            If picButtons(i).Tag = "2" Or picButtons(i).Tag = "1" Then
                picButtons(i).AutoRedraw = False
                picButtons(i).Cls
                picButtons(i).Refresh
                picButtons(i).AutoRedraw = True
                picButtons(i).BackColor = BackColor
                picButtons(i).Tag = "0"
               ' DrawImage i
            End If
        Next
        
    picButtons(Index).Tag = "2"
    DrawDownButton picButtons(Index), 64, 102
    RaiseEvent ButtonClick(Index, iList.ItemKey(Index + 1))
End Sub

Private Sub picButtons_GotFocus(Index As Integer)
    picButtons_Click (Index)
End Sub

Private Sub picButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If picButtons(Index).Tag = "1" Then

        Dim i As Integer
        For i = 0 To picButtons.UBound
            If picButtons(i).Tag = "1" Then
                picButtons(i).AutoRedraw = False
                picButtons(i).Cls
                picButtons(i).Refresh
                picButtons(i).AutoRedraw = True
                picButtons(i).BackColor = BackColor
                picButtons(i).Tag = "0"
                DrawImage i
            End If
        Next
        picButtons(Index).Tag = "1"
        DrawDownButton picButtons(Index), 64, 102
    Else
        For i = 0 To picButtons.UBound
            If i <> Index Then
                picButtons(i).AutoRedraw = False
                picButtons(i).Cls
                picButtons(i).Refresh
                picButtons(i).AutoRedraw = True
                picButtons(i).BackColor = BackColor
                DrawImage i
            End If
        Next
    End If


    
    
End Sub

Private Sub picButtons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If picButtons(Index).Tag = "2" Then Exit Sub
    If picButtons(Index).Tag = "1" Then Exit Sub
    picButtons(Index).Tag = "1"
    Dim i As Integer
    For i = 0 To picButtons.UBound
        If picButtons(i).Tag = "1" And Not i = Index Then
            picButtons(i).AutoRedraw = False
            picButtons(i).Cls
            picButtons(i).Refresh
            picButtons(i).AutoRedraw = True
            picButtons(i).BackColor = BackColor
            picButtons(i).Tag = "0"
            DrawImage i
        End If
    Next
    picButtons(Index).AutoRedraw = False
    'picButtons(Index).Tag = "1"
    DrawButton picButtons(Index), 64, 102
    picButtons(Index).AutoRedraw = True
End Sub

Private Sub picButtons_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
For i = 0 To picButtons.UBound
        If picButtons(i).Tag = "1" Then
            picButtons(i).AutoRedraw = False
            picButtons(i).Cls
            picButtons(i).Refresh
            picButtons(i).AutoRedraw = True
            picButtons(i).BackColor = BackColor
            picButtons(i).Tag = "0"
            DrawImage i
        End If
    Next
End Sub

Private Sub picDown_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    CheckForScrollDown
    
    
    
    tmrMD.Enabled = True
End Sub

Private Sub picDown_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove 1, 1, 1, 1
End Sub

Private Sub picDown_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrMD.Enabled = False
End Sub

Private Sub picHeader_Resize()
    DoEvents
    
   ' DrawHeader
End Sub

Private Sub picUp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   CheckForScrollUp
   tmrUp.Enabled = True
End Sub

Private Sub picUp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseMove 1, 1, 1, 1
End Sub

Private Sub picUp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrUp.Enabled = False
End Sub

Private Sub tmrMD_Timer()
    If picDown.Visible = True Then
        CheckForScrollDown
        DrawImages
    End If
End Sub

Private Sub tmrUp_Timer()
    If picUp.Visible = True Then
        CheckForScrollUp
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
    Set iList = New dlgImgList
    Set iList.ParentObject = Me
    iList.Create
    picScroll.Move 0, picHeader.Height, UserControl.Width, UserControl.Height
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent mousedown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 
    Dim i As Integer
    For i = 0 To picButtons.UBound
        If picButtons(i).Tag = "1" Then
            picButtons(i).AutoRedraw = False
            picButtons(i).Cls
            picButtons(i).Refresh
            picButtons(i).AutoRedraw = True
            picButtons(i).BackColor = BackColor
            picButtons(i).Tag = "0"
        End If
    Next
    
   
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets the color that specifies transparent areas in the MaskPicture."
    MaskColor = UserControl.MaskColor
End Property

Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Private Sub UserControl_Paint()
    
    
    ThinBorder UserControl.hwnd, True
    If HeaderVisible Then
        DrawHeader
    End If
    picBottom.AutoRedraw = False
    picBottom.Cls
    picBottom.Refresh
    
    DrawButton picBottom, picBottom.Height / Screen.TwipsPerPixelY, picBottom.Width / Screen.TwipsPerPixelX
    picBottom.AutoRedraw = True
End Sub

Private Sub UserControl_Resize()
    
    If UserControl.Height <= picHeader.Height + picBottom.Height Then
        UserControl.Height = picHeader.Height + picBottom.Height + 30
    End If
    picHeader.AutoRedraw = False
    picHeader.Cls
    'picHeader.Refresh
    'picHeader.AutoRedraw = True
    If ShowScrollBars Then
        UserControl.Width = 2040
    Else
        UserControl.Width = 1660
    End If
    
    If HeaderVisible Then

        picHeader.Move 0, 0, UserControl.Width - 30, picHeader.Height
        picScroll.Move 0, picHeader.Height, UserControl.Width, UserControl.Height - picHeader.Height

        DrawHeader
    Else
        picHeader.Move 0, 0, UserControl.Width - 30, picHeader.Height
        picScroll.Move 0, 0, UserControl.Width, UserControl.Height

    End If
    
    
    RaiseEvent Resize
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,PopupMenu
Public Sub PopupMenu(ByVal Menu As Object, Optional ByVal flags As Variant, Optional ByVal x As Variant, Optional ByVal y As Variant, Optional ByVal DefaultMenu As Variant)
Attribute PopupMenu.VB_Description = "Displays a pop-up menu on an MDIForm or Form object."
    UserControl.PopupMenu Menu, flags, x, y, DefaultMenu
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    
'    BackColor = &H808080
    Set UserControl.Font = Ambient.Font
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ShowScrollBars = PropBag.ReadProperty("ShowScrollBars", False)
    BackColor = PropBag.ReadProperty("BackColor", &H808080)
 
 
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", vbWhite)
    picCheck.ForeColor = PropBag.ReadProperty("ForeColor", vbWhite)
    mHeadForeColor = PropBag.ReadProperty("HeaderForeColor", vbWhite)
    mCaption = PropBag.ReadProperty("Caption", "Menu Caption")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.MaskColor = PropBag.ReadProperty("MaskColor", -2147483633)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    ThinBorder UserControl.hwnd, True
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ShowScrollBars", MScroll, &H8000000F)
    Call PropBag.WriteProperty("BackColor", mBackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", fColor, vbWhite)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, -2147483633)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("Caption", mCaption, "Menu Caption")
    Call PropBag.WriteProperty("HeaderForeColor", mHeadForeColor, vbWhite)
    
End Sub



Public Sub AddButtonFromFile(Caption As String, ImagePath As String, Optional strToolTipText As String)
   Dim i As Integer

        For i = 0 To picButtons.UBound
            If picButtons(i).Tag = "1" Or picButtons(i).Tag = "2" Then
                picButtons(i).AutoRedraw = False
                picButtons(i).Cls
                picButtons(i).Refresh
                picButtons(i).AutoRedraw = True
                picButtons(i).BackColor = BackColor
                picButtons(i).Tag = "0"
                DrawImage i
            End If
        Next
    picScroll.BackColor = mBackColor
   If Right(LCase(ImagePath), 3) = "ico" Then
        iList.AddFromFile ImagePath, IMAGE_ICON, Caption, False, BackColor
   Else
        iList.AddFromFile ImagePath, IMAGE_BITMAP, Caption, False, BackColor
   End If
   
  
   
   For i = 0 To picButtons.UBound
        picButtons(i).Visible = False
   Next
   
    If Not iList.ImageCount - 1 > picButtons.UBound Then
        picButtons(iList.ImageCount - 1).PaintPicture iList.ItemPicture(iList.ImageCount), (picButtons(iList.ImageCount - 1).Width - 32) / 2, 150
        
    Else
        If Not iList.ImageCount = 0 Then Load picButtons(picButtons.UBound + 1)
        
        picButtons(picButtons.UBound).PaintPicture iList.ItemPicture((picButtons.UBound + 1)), (picButtons(picButtons.UBound).Width - 32) / 2, 150
               
        
        'setparent lblcaption(lblcaption.UBound ).hwnd picbuttons.hwnd
        
    End If
    Dim Xoffset As Long
    Dim yOffset As Long
    
    Xoffset = 60
    yOffset = 45
   picButtons(picButtons.UBound).ToolTipText = strToolTipText
   picButtons(0).Move Xoffset, yOffset
   picButtons(0).Visible = True
   DrawImage 0
   If picHeader.Visible Then
        picUp.Move UserControl.Width - picUp.Width - 120, 120 + picHeader.Height
   Else
        picUp.Move UserControl.Width - picUp.Width - 120, 120
   End If
   For i = 1 To picButtons.UBound
        picButtons(i).Move Xoffset, picButtons(i - 1).tOp + picButtons(i).Height
        picButtons(i).Cls
        picButtons(i).AutoRedraw = True
        picButtons(i).Visible = True
        
        
        DrawImage i
   Next
   picScroll.Height = picButtons(picButtons.UBound).tOp + picButtons(picButtons.UBound).Height + 15
   If Not ShowScrollBars Then Exit Sub
   Dim tmpHeight As Long
   
   If picHeader.Visible Then tmpHeight = tmpHeight & picHeader.Height
   tmpHeight = tmpHeight + picBottom.Height
   If picScroll.Height > UserControl.Height - tmpHeight Then
        picDown.Move UserControl.Width - picDown.Width - 120, UserControl.Height - picDown.Height - 190
        picDown.Visible = True
   Else
        picDown.Visible = False
   End If
    'CheckForScroll
End Sub

Public Sub DrawImage(iIndex As Integer)
    If iList.ImageCount = 0 Then Exit Sub
    If Not picCheck.BackColor = BackColor Then picCheck.BackColor = BackColor

'       ThinBorder picCheck.hwnd, True
'    Else
        picCheck.AutoRedraw = False
        picCheck.Cls
        picCheck.Refresh
        picCheck.AutoRedraw = True
        Set picCheck.Picture = LoadPicture("")
        picCheck.Refresh
    Dim stdpic As New StdPicture
    Set stdpic = iList.ItemPicture((iIndex + 1))
    
    picCheck.PaintPicture stdpic, ((picCheck.Width / Screen.TwipsPerPixelX) - 32) / 2, 10
    picCheck.CurrentY = picCheck.ScaleHeight - (picCheck.TextHeight(iList.ItemKey(iIndex + 1))) - 15
    picCheck.CurrentX = (picCheck.ScaleWidth - picCheck.TextWidth(iList.ItemKey(iIndex + 1))) / 2
    picCheck.ForeColor = ForeColor
    picCheck.Print iList.ItemKey(iIndex + 1)
    DoEvents
    picButtons(iIndex).Picture = picCheck.Image
    If picButtons(iIndex).Tag = "1" Then DrawButton picButtons(iIndex), 64, 102
    'picButtons(Index).Refresh
End Sub

Public Sub AddButtonFromSTDPicture(ImageObj As StdPicture, Caption As String, Optional strToolTipText As String)
   Dim i As Integer

        For i = 0 To picButtons.UBound
            If picButtons(i).Tag = "1" Or picButtons(i).Tag = "2" Then
                picButtons(i).AutoRedraw = False
                picButtons(i).Cls
                picButtons(i).Refresh
                picButtons(i).AutoRedraw = True
                picButtons(i).BackColor = BackColor
                picButtons(i).Tag = "0"
                DrawImage i
            End If
        Next
    
    
   Dim tmpSTDPic As StdPicture
   Set tmpSTDPic = ImageObj
  
   If ImageObj.Type = 3 Then 'icon
        iList.AddFromHandle tmpSTDPic.Handle, IMAGE_ICON, Caption
    ElseIf ImageObj.Type = 1 Then
       iList.AddFromHandle tmpSTDPic.Handle, IMAGE_BITMAP, Caption
    End If
   'Clipboard.SetData iList.ItemPicture(iList.ImageCount), vbCFBitmap
   DoEvents
   
   
   
  
   
   For i = 0 To picButtons.UBound
        picButtons(i).Visible = False
   Next
   
   
   If iList.ImageCount = 0 Then Err.Raise 9002, Err.Source, Err.Description: Exit Sub
        If Not iList.ImageCount - 1 > picButtons.UBound Then
            picButtons(iList.ImageCount - 1).PaintPicture iList.ItemPicture((picButtons.UBound + 1)), (picButtons(iList.ImageCount - 1).Width - 32) / 2, 150
        Else
            If Not iList.ImageCount = 0 Then Load picButtons(picButtons.UBound + 1)
            picButtons(picButtons.UBound).PaintPicture iList.ItemPicture((picButtons.UBound + 1)), (picButtons(picButtons.UBound).Width - 32) / 2, 150
        End If
        
    Dim Xoffset As Long
    Dim yOffset As Long
    
    Xoffset = 60
    yOffset = 45
    
   picButtons(0).Move Xoffset, yOffset
   picButtons(0).Visible = True
   DrawImage 0
   
   For i = 1 To picButtons.UBound
        picButtons(i).Move Xoffset, picButtons(i - 1).tOp + picButtons(i).Height
        picButtons(i).Cls
        picButtons(i).AutoRedraw = True
        picButtons(i).Visible = True
        DrawImage i
   Next
   
   picScroll.Height = picButtons(picButtons.UBound).tOp + picButtons(picButtons.UBound).Height + 15
   picUp.left = UserControl.Width - picUp.Width - 120
   If Not ShowScrollBars Then Exit Sub
   Dim tmpHeight As Long
   If picHeader.Visible Then tmpHeight = tmpHeight & picHeader.Height
   tmpHeight = tmpHeight + picBottom.Height
   If picScroll.Height > UserControl.Height - tmpHeight Then
        picDown.Visible = False
   End If
   
End Sub

Public Sub CheckForScrollDown()
    If Not ShowScrollBars Then Exit Sub
    If Abs(picScroll.Height) - Abs(picScroll.tOp - picButtons(0).Height + picButtons(0).tOp) + 15 >= UserControl.Height Then
        picScroll.Move 0, picScroll.tOp - picButtons(0).Height - 140
        'picDown.Move UserControl.Width - picDown.Width - 120, UserControl.Height - picDown.Height - 120
    Else
        picScroll.tOp = picScroll.tOp - picButtons(0).Height - 35
        picDown.Visible = False
    End If
    
    If picScroll.tOp < 0 Then
        'picScroll.Move 0, picScroll.Top + picButtons(0).Height - 140
        
        picUp.Visible = True
    Else
        'picScroll.Top = picScroll.Top + picButtons(0).Height - 35
        picUp.Visible = False
    End If
    
   If picHeader.Visible Then
        picUp.Move UserControl.Width - picUp.Width - 120, 120 + picHeader.Height
   Else
        picUp.Move UserControl.Width - picUp.Width - 120, 120
   End If
    
End Sub
Public Sub CheckForScrollUp()
    If Not ShowScrollBars Then Exit Sub
    If Not picHeader.Visible = True Then
        If picScroll.tOp + picButtons(0).Height + 120 < 60 Then
            picScroll.Move 0, picScroll.tOp + (picButtons(0).Height + 120)
            
            picUp.Visible = True
        Else
        
            picScroll.tOp = 0
            picUp.Visible = False
        End If
    Else
        If picScroll.tOp + picButtons(0).Height + 120 + picHeader.Height < 60 Then
            picScroll.Move 0, picScroll.tOp + (picButtons(0).Height + 120) + picHeader.Height + picHeader.Height
            
            picUp.Visible = True
        Else
        
            picScroll.tOp = picHeader.Height
            picUp.Visible = False
        End If
    
    
    End If
    
    If Abs(picScroll.tOp) + picScroll.Height >= UserControl.Height Then
    
        picDown.Visible = True
    Else
    
        picDown.Visible = False
    End If
    
End Sub
Public Sub DrawImages()
    If iList.ImageCount = 0 Then Exit Sub
    Dim i As Integer
    For i = 0 To picButtons.UBound
        DrawImage i
    Next
End Sub

Public Sub ForceClick(iIndex As Integer)
    If Not iIndex > picButtons.UBound Then
        Dim i As Integer
        
        For i = 0 To picButtons.UBound
            
            If i <> iIndex Then
                picButtons(i).Tag = "0"
                DrawImage i
            End If
        Next
        picButtons(iIndex).Tag = "2"
        DrawDownButton picButtons(iIndex), 64, 102
        
    End If
    
    RaiseEvent ButtonClick(iIndex, iList.ItemKey(iIndex + 1))
End Sub

Public Function GetActiveButtonsPicture() As StdPicture
    Dim i As Integer
    
    For i = 0 To picButtons.UBound
        If picButtons(i).Tag = "2" Then
            Set GetActiveButtonsPicture = iList.ItemPicture(i + 1)
            Exit Function
        End If
    Next
    
    
End Function

Public Property Get ShowScrollBars() As Boolean
    ShowScrollBars = MScroll
End Property

Public Property Let ShowScrollBars(ByVal vNewValue As Boolean)
    
    MScroll = vNewValue
        
    UserControl_Resize
    DoEvents
    If MScroll = False Then
        picUp.Visible = False
        picDown.Visible = False
'        picHeader.Visible = False
'        picHeader.Height = 0
'        picHeader.AutoRedraw = False
'        picHeader.Cls
'        picHeader.Height = 340
'        picHeader.AutoRedraw = True
'        picHeader.Refresh
'
        
    Else
'        picHeader.Visible = False
'        picHeader.Height = 0
'        picHeader.AutoRedraw = False
'        picHeader.Cls
'        picHeader.Height = 340
'        picHeader.AutoRedraw = True
'        picHeader.Refresh
'        picHeader.Visible = True
    End If
    
    
   'DrawHeader
    
    

    PropertyChanged "ShowScrollBars"

    picScroll.Height = picButtons(picButtons.UBound).tOp + picButtons(picButtons.UBound).Height
    If Not MScroll Then Exit Property
    If picScroll.Height > UserControl.Height Then
         picDown.Move UserControl.Width - picDown.Width - 120, UserControl.Height - picDown.Height - 190
         picDown.Visible = True
    Else
         picDown.Visible = False
    End If

    
    
    
End Property

Public Property Get HeaderVisible() As Boolean
    HeaderVisible = mShowHeader
End Property

Public Property Let HeaderVisible(ByVal vNewValue As Boolean)
    mShowHeader = vNewValue
    picHeader.Visible = mShowHeader
    
    If Not mShowHeader Then
        
        If picScroll.tOp - (picHeader.tOp + picHeader.Height) < 60 Then picScroll.tOp = 0
        picHeader.Height = 0
    Else
        picHeader.Height = 340
        picHeader.Move 0, 0, UserControl.Width - 30, picHeader.Height
        picScroll.tOp = picHeader.Height + 5
        AlreadyDrawing = True
        
        
    End If
    
    If picUp.Visible And mShowHeader = False Then
        picUp.tOp = 120
    Else
        picUp.tOp = 120 + picHeader.Height
    End If
    
   picScroll.Height = picButtons(picButtons.UBound).tOp + picButtons(picButtons.UBound).Height + 15
   picUp.left = UserControl.Width - picUp.Width - 120
   DrawHeader
   If Not ShowScrollBars Then Exit Property
   If picScroll.Height > UserControl.Height Then
        picDown.Move UserControl.Width - picDown.Width - 120, UserControl.Height - picDown.Height - 190
        picDown.Visible = True
   Else
        picDown.Visible = False
   End If
   
    UserControl_Resize

End Property

Public Sub DrawHeader()
        picHeader.AutoRedraw = False
        
        lblCaption.Move 30, 70, picHeader.Width - 60, picHeader.TextHeight(Caption)
        AlreadyDrawing = True
        DoEvents
        picHeader.Cls
        DoEvents
        'picHeader.Refresh
        'DoEvents
        
       
        
        
        
        
        
        
        
       
       DrawButton picHeader, picHeader.Height / Screen.TwipsPerPixelY, picHeader.Width / Screen.TwipsPerPixelX
       picHeader.AutoRedraw = True
        
        'Set picHeader.Picture = picHeader.Image
        
        
        
        
        
        'picHeader.Refresh
        lblCaption.ForeColor = HeaderForeColor
        Set lblCaption.Font = Font
        lblCaption.Caption = Caption
        
        AlreadyDrawing = False
End Sub



Public Property Get Caption() As String
    Caption = mCaption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    mCaption = vNewValue
   
End Property

Public Property Get HeaderForeColor() As OLE_COLOR
    HeaderForeColor = mHeadForeColor
End Property

Public Property Let HeaderForeColor(ByVal vNewValue As OLE_COLOR)
    mHeadForeColor = vNewValue
    DrawHeader
End Property
