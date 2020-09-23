Attribute VB_Name = "modThemes"

Private m_bInhibitOptionClick As Boolean

' -----------------------------------------------------------------------
' For setting up a thin border on a picture box control:
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_STATICEDGE = &H20000
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
'Public Type RECT
'   left     As Long
'   tOp      As Long
'   Right    As Long
'   Bottom   As Long
'End Type
Public Declare Function GetSystemMenu Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal bRevert As Long) As Long


Public Declare Function RemoveMenu Lib "user32" _
    (ByVal hMenu As Long, _
    ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long
    Public Const MF_BYPOSITION = &H400&


Public OldWindowProc2 As Long

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Const GWL_WNDPROC = (-4)
Public Const WM_USER = &H400

Public MsgNames As New Collection

Public Const WM_MENUSELECT = &H11F
Public Const WM_COMMAND = &H111


Public Const WM_SYSCOMMAND = &H112
Public Const MF_SEPARATOR = &H800&
Public Const IDM_ABOUT = 1999
' *********************************************
' Pass along all messages except the one that
' makes the context menu appear.
' *********************************************
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Const SPI_GETWORKAREA& = 48
'get available real estate & 'Control+Alt+ delete
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Public resetting As Boolean
Public AppPath As String
Public Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long


Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long


Public Function AddBackSlash(ByVal sPath As String) As String
    'Returns sPath with a trailing backslash
    '     if sPath does not
    'already have a trailing backslash. Othe
    '     rwise, returns sPath.
    sPath = Trim$(sPath)


    If Len(sPath) > 0 Then
        sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")
    End If
    AddBackSlash = sPath
    
End Function


Public Function GetLongFilename(ByVal sShortFilename As String) As String
    'Returns the Long Filename associated wi
    '     th sShortFilename
    Dim lRet As Long
    Dim sLongFilename As String
    'First attempt using 1024 character buff
    '     er.
    sLongFilename = String$(1024, " ")
    lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    
    'If buffer is too small lRet contains bu
    '     ffer size needed.


    If lRet > Len(sLongFilename) Then
        'Increase buffer size...
        sLongFilename = String$(lRet + 1, " ")
        'and try again.
        lRet = GetLongPathName(sShortFilename, sLongFilename, Len(sLongFilename))
    
    End If
    
    If lRet > 0 Then
        GetLongFilename = left$(sLongFilename, lRet)
    End If


End Function

'Public Function TaskBarDimensions() As RECT
''Find out what the taskbars', left, top, right and bottom values are
''in TWIPS
'
'Const X = 0
'Const Y = 1
'
'Dim WorkArea As RECT
'Dim TaskBarDet As RECT
'Dim TwipsPP(2) As Byte 'Twips Per Pixel
'
'WorkArea = GetWorkArea
'TwipsPP(X) = Screen.TwipsPerPixelX
'TwipsPP(Y) = Screen.TwipsPerPixelY
'
''set the taskbars' default values to the screen size
'TaskBarDet.Top = 0
'TaskBarDet.Bottom = Screen.Height
'TaskBarDet.Left = 0
'TaskBarDet.Right = Screen.Width
'
''change the appropiate value according to alignment
'Select Case GetAlignment
'Case vbLeft
'    TaskBarDet.Right = (WorkArea.Left * TwipsPP(X))
'Case vbRight
'    TaskBarDet.Left = (WorkArea.Right * TwipsPP(X))
'Case vbTop
'    TaskBarDet.Bottom = (WorkArea.Top * TwipsPP(Y))
'Case vbBottom
'    TaskBarDet.Top = (WorkArea.Bottom * TwipsPP(Y))
'End Select
'
''return result
'TaskBarDimensions = TaskBarDet
'End Function
'
'
'Public Function GetUncompromisedRealEstate() As RECT
'    Dim myRect As RECT
'    Dim R As Long
'    Dim Msg As String
'
'    R = SystemParametersInfo(SPI_GETWORKAREA, 0&, myRect, 0&)
'    ' set available real estate variables
'    'UncompromisedLeft = myRect.Left
'    'UncompromisedTop = myRect.Top
'    'UncompromisedBottom = myRect.Bottom
'    'UncompromisedRight = myRect.Right
'    GetUncompromisedRealEstate = myRect
'End Function
'Public Function NewWindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Const MIN_MENU = 1001
'Const MAX_MENU = 1010
'
'    If Msg = WM_COMMAND Then
'        MsgBox "Command ID" & Str$(wParam)
'        If wParam >= MIN_MENU And _
'            wParam <= MAX_MENU _
'                Then Exit Function
'    End If
'
'    If Msg = WM_SYSCOMMAND And (wParam = IDM_ABOUT Or wParam = 1024 Or wParam = 2048) Then
'        Select Case wParam
'            Case IDM_ABOUT
'                MsgBox "About"
'            Case 1024
'                MsgBox "Exit"
'            Case 2048
'                MsgBox "Login"
'        End Select
'
'        Exit Function
'    End If
'
'    NewWindowProc = CallWindowProc( _
'        OldWindowProc2, hwnd, Msg, wParam, _
'        lParam)
'End Function

Public Sub DisableCloseWindowButton(frm As Form)
    Dim hSysMenu As Long
    'Get the handle to this windows
    'system menu
    hSysMenu = GetSystemMenu(frm.hwnd, 0)
    'This will disable the close button
    RemoveMenu hSysMenu, 6, MF_BYPOSITION
    'remove the seperator bar
    RemoveMenu hSysMenu, 5, MF_BYPOSITION
    RemoveMenu hSysMenu, 1, MF_BYPOSITION
    
End Sub
'Public Sub drawtext(hdc As Long, Text As String, xpos As Long, yPos As Long, color As Long, opacity As Double, FontName As String, FontSize As Long, Optional FontWeight As Single = 2)
'    Dim Size                                  As DWord
'    Dim Ret                                   As Long
'    Dim ndc                                   As Long
'    Dim nbmp                                  As Long
'    Dim hjunk
'    Dim Font                                  As LOGFONT
'    Dim HFont                                 As Long
'    Dim pixels()                              As RGBQUAD
'    Dim npixels()                             As RGBQUAD
'    Dim bgpixels()                            As RGBQUAD
'    Dim rgbcol(3)                             As Byte
'    Dim X, Y, yy
'    Dim bminfo                                As BITMAPINFO
'    Dim tmp                                   As Double
'    Dim alpha                                 As Double
'    With Font
'        .lfHeight = -(FontSize * 20) / Screen.TwipsPerPixelY ' set font size
'        .lfFaceName = FontName & Chr(0) 'apply font name
'        .lfWeight = FontWeight    'this is how bold the font is .. apply a in param if you want
'    End With
'
'    '-----------------------------------------
'    'create a dc for our backbuffer
'    ndc = CreateCompatibleDC(hdc)
'    'create a bitmap for our backbuffer
'    nbmp = CreateCompatibleBitmap(hdc, 1, 1) 'make a temp bitmap so we can get the size of the text
'    'attach our bitmap to our backbuffer
'    hjunk = SelectObject(ndc, nbmp)
'    'apply the font to our backbuffer
'    HFont = CreateFontIndirect(Font)
'    SelectObject ndc, HFont
'
'    'get size of the text we want to draw
'    Ret = GetTabbedTextExtent(ndc, Text, Len(Text), 0, 0)
'
'    'delete our temp bmp
'    DeleteObject HFont
'    DeleteObject ndc
'    DeleteObject nbmp
'    'this part was only to measure the size of the text
'    '----------------------------------------
'    'now lets draw the text...
'
'
'    'split our color value to a byte array
'    'this is my own invention ... pretty nice (?)
'    CopyMemoryLong VarPtr(rgbcol(0)), VarPtr(color), 4
'    'split the return value from gettextextent into two integers
'    CopyMemoryLong VarPtr(Size), VarPtr(Ret), 4
'
'    yPos = yPos - Size.high / 2
'    'create a dc for our backbuffer
'    ndc = CreateCompatibleDC(hdc)
'    'create a bitmap for our backbuffer
'    nbmp = CreateCompatibleBitmap(hdc, Size.low, Size.high)
'    'attach our bitmap to our backbuffer
'    hjunk = SelectObject(ndc, nbmp)
'    'apply the font to our backbuffer
'    HFont = CreateFontIndirect(Font)
'    SelectObject ndc, HFont
'    'set black background coloy
'    SetBkColor ndc, 0
'    'set white forecolor
'    SetTextColor ndc, vbWhite
'    'write the text to our backbuffer
'    TabbedTextOut ndc, 0, 0, Text, Len(Text), 0, 0, 0
'    'resize the arrays to the same size as the bbuffer
'    ReDim pixels(Size.low - 1, Size.high - 1)
'    ReDim npixels(Size.low - 1, Size.high - 1)
'    ReDim bgpixels(Size.low - 1, Size.high - 1)
'
'    'set the bitmap info (so we can get the gfx data in and out of our arrays
'    With bminfo.bmiHeader
'        .biSize = Len(bminfo.bmiHeader)
'        .biWidth = Size.low
'        .biHeight = Size.high
'        .biPlanes = 1
'        .biBitCount = 32
'    End With
'    'store the drawn text in our "pixels" array
'    GetDIBits ndc, nbmp, 0, Size.high, pixels(0, 0), bminfo, 1
'    'get the bg graphics into our "bgpixels" array
'    BitBlt ndc, 0, 0, Size.low, Size.high, hdc, xpos, yPos, vbSrcCopy
'    GetDIBits ndc, nbmp, 0, Size.high, bgpixels(0, 0), bminfo, 1
'    yy = Int(Size.high / 2)
'    npixels = bgpixels
'    For X = 0 To Size.low - 2 Step 2
'        For Y = 0 To Size.high - 2 Step 2
'            'alpha is the average of the color of 2*2 pixels /255
'            'now we have a value between 0 and 1
'            '0 is transparent
'            '1 is soild white
'            'now multiply alpha with the opacity factor
'            'ie if opacity is 0.5 ...  aplha will be max 0.5
'            'since we draw our text with white . we only need to check the strength of one color (in this case blue)
'            'coz red and green will always be the same as the blue
'            alpha = (((0 + (pixels(X + 0, Y + 0).rgbBlue) + (pixels(X + 1, Y + 0).rgbBlue) + (pixels(X + 0, Y + 1).rgbBlue) + (pixels(X + 1, Y + 1).rgbBlue)) / 4) / 255) * opacity
'            'alpha is now the opacity factor 0-1
'            'calculate amount of blue to apply
'            'and how much of the background that is going to be seen
'            tmp = (alpha * rgbcol(2)) + bgpixels(X / 2, Y / 2).rgbBlue * (1 - alpha)
'            'never go higher than 255
'            If tmp > 255 Then tmp = 255
'            'store the result at x/2 and y/2 (the new picture is only 0.5 times as high and wide
'            npixels(X / 2, Y / 2).rgbBlue = tmp
'            'calculate amount of red to apply
'            'and how much of the background that is going to be seen
'            tmp = (alpha * rgbcol(0)) + bgpixels(X / 2, Y / 2).rgbRed * (1 - alpha)
'            'never go higher than 255
'            If tmp > 255 Then tmp = 255
'            npixels(X / 2, Y / 2).rgbRed = tmp
'            'calculate amount of green to apply
'            'and how much of the background that is going to be seen
'            tmp = (alpha * rgbcol(1)) + bgpixels(X / 2, Y / 2).rgbGreen * (1 - alpha)
'            'never go higher than 255
'            If tmp > 255 Then tmp = 255
'            npixels(X / 2, Y / 2).rgbGreen = tmp
'        Next
'    Next
'    'apply the new picture to our bbuffer-dc
'    SetDIBits ndc, nbmp, 0, Size.high, npixels(0, 0), bminfo, 1
'    'blit our bbuffer-dc to the screen
'    BitBlt hdc, xpos, yPos, Size.low, Size.high, ndc, 0, 0, vbSrcCopy
'    'clean up
'    DeleteObject HFont
'    DeleteObject ndc
'    DeleteObject nbmp
'End Sub

Public Function ThinBorder(ByVal lhWnd As Long, ByVal bState As Boolean)
Dim lS As Long

   lS = GetWindowLong(lhWnd, GWL_EXSTYLE)
   If Not (bState) Then
      lS = lS Or WS_EX_CLIENTEDGE And Not WS_EX_STATICEDGE
   Else
      lS = lS Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
   End If
   SetWindowLong lhWnd, GWL_EXSTYLE, lS
   SetWindowPos lhWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Function
'Public Sub Refreshtheme()
'    Dim tmpform As Form
'    Dim tmpControl As Control
'    Dim tmpString As String
'    Dim inFont As Integer
'
'    On Error Resume Next
'    'Set myProgProps = New clsProgramProperties
'    'tmpString = myProgProps.ApplicationFont
'
'    Dim tmpFont As New StdFont
'    tmpFont.Name = tmpString
'    tmpFont.Bold = False
'    tmpFont.Italic = False
'    tmpFont.Underline = False
'
'
'
'        For Each tmpform In Forms
'            On Error Resume Next
'
'            tmpform.tBar.StartColor = myProgProps.TitleBarBgColor
'            tmpform.tBar.TextColor = myProgProps.TitleBarTextColor
'
'            tmpform.picIcon.BackColor = myProgProps.TitleBarBgColor
'            If tmpform.Name = "frmProps" Then
'                tmpform.Tabs.Visible = False
'                tmpform.Tabs.Visible = True
'                Set frmProps.Tabs.Font.Name = tmpFont
'                frmProps.Tabs.Captions = frmProps.Tabs.Captions & ""
'            End If
'
'            If Not tmpform.Name = "frmMain" And Not tmpform.Name = "frmNotes" And Not tmpform.Name = "frmCListOptions" And Not tmpform.Name = "frmPreview" Then ThinBorder tmpform.hwnd, True
'            If Not tmpform.Name = "frmProps" And Not tmpform.Name = "frmEditItemInfo" And Not tmpform.Name = "frmAddExistingPart" And Not tmpform.Name = "frmModDelete" And Not tmpform.Name = "frmAddUser" And Not tmpform.Name = "frmUsers" And Not tmpform.Name = "frmModUser" And Not tmpform.Name = "frmOrganise" Then
'                    If tmpform.Name = "frmMain" Then
'                        Set tmpform.MenuBar.Font = tmpFont
'                        'drawtext tmpform.tBar.hdc, "OFFICE", 30, 5, vbWhite, 1, "Advertiser", 20
'                        'drawtext tmpform.tBar.hdc, "MANAGER", 80, 5, &HFADAA5, 1, "Advertiser", 20
'                        'drawtext tmpform.tBar.hdc, "2001", 152, 3, &H80FF&, 1, "Advertiser", 25, 7
'                    End If
'                tmpform.BackColor = myProgProps.WindowObjectsColor
'                If Not tmpform.Name = "frmMain" Then tmpControl.Font.Name = tmpString
'                For Each tmpControl In tmpform.Controls
'
'                    Select Case UCase(TypeName(tmpControl))
'                        Case "CALENDAR", "SHAPE", "LABEL"
'                            If tmpControl.Name = "shpHeader" Then
'                               'tmpControl.FillColor = myProgProps.TitleBarBgColor
'                            ElseIf tmpControl.Name = "lblHeader" Then
'                               ' tmpControl.ForeColor = myProgProps.TitleBarTextColor
'                            End If
'
'
'                        Case "ASXPOWERBUTTON"
'                            tmpControl.TextColor = vbBlack
'                            tmpControl.HotTrackingColor = myProgProps.TitleBarBgColor
'                            tmpControl.BackColor = myProgProps.WindowObjectsColor
'                        Case "PICTUREBOX"
'                            If tmpform.Name = "frmPrintPreview" Or tmpform.Name = "frmPreview" Then
'                                If Not tmpControl.Name = "picPreview" And Not tmpControl.Name = "picBG" And Not tmpControl.Name = "picIcon" And Not tmpControl.Name = "picBuffer" Then  ' And Not tmpControl.Name = "picFrame"
'                                    tmpControl.BackColor = myProgProps.WindowObjectsColor
'                                End If
'                            Else
'                                If Not tmpControl.Name = "picClose" And Not tmpControl.Name = "picIcon" And Not tmpControl.Name = "Picture3" And Not tmpControl.Name = "picUserProps" And Not tmpControl.Name = "picForm" Then  'And Not tmpControl.Name = "picFrame"
'                                    tmpControl.BackColor = myProgProps.WindowObjectsColor
'                                Else
'                                    If tmpControl.Name = "picForm" Then tmpControl.BackColor = vbWhite
'
'                                End If
'                            End If
'                        Case "LISTBOX"
'
'                        Case "MSFLEXGRID"
'                            tmpControl.BackColorSel = myProgProps.TitleBarBgColor
'                            tmpControl.BackColor = vbWhite
'
'                        Case "ABSTOPICLIST"
'                            tmpControl.HighlightedColor = myProgProps.TitleBarBgColor
'                        Case "CALENDARVB"
'                        Case "LISTVIEW"
'                            'SkinLV tmpControl
'                        Case "SCROLLREPORT"
'                            If Not tmpform.Name = "frmPrintPreview" Then
'
'                                    tmpControl.BackColor = myProgProps.WindowObjectsColor
'                            End If
'                        Case "ASXPANEL"
'                                If Not tmpControl.Name = "tPanel" Then tmpControl.BackColor = myProgProps.WindowObjectsColor
'
'                        Case Else
'                            If Not TypeName(tmpControl) = "TextBox" And Not TypeName(tmpControl) = "ComboBox" And Not TypeName(tmpControl) = "Label" And Not tmpControl.Name = "WebPane" Then
'                                tmpControl.BackColor = myProgProps.WindowObjectsColor
'                            End If
'
'                    End Select
'                    ''''select case typename here
'
'                    tmpControl.Refresh
'                Next
'
'            End If
'        Next
'
'
'
'
'
'
'    Err.Clear
'End Sub


