VERSION 5.00
Begin VB.UserControl CustomButton 
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   DefaultCancel   =   -1  'True
   ScaleHeight     =   61
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   124
End
Attribute VB_Name = "CustomButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' This instance of the template does not have any drawing routines.
' Whereas the same project has some sample drawing routines.

' Update History:
' 18 Feb 09:
'   :: Added support for buttons that have the Default property set to True
'   :: Reworked UpdateState & DrawButton routines to minimize unnecessary paints even more
'   :: Added the Value property to replicate VB button's Value property
'   :: Events should be triggered after drawing, not before; missed in previous updates
'   :: Retweaked 1Jan09 patch: the patch did not account for preventing double clicks in the same scenario
' 1 Jan 09: Fixed anamoly that would fire click event if spacebar held down on button and mouse clicks outside of button
' 27 Sep 07: Not all events related to a click (btn down/up/click/etc) were in same order as VB's command button. Now are.
' 1 Jan 09: Bug noted by Juned Chippa. Focus button, press spacebar, click mouse elsewhere off button, release spacebar: Click event
'   This is not consistent with VB's command button. UpdateState modified to look for the mouse down in this case.

'-------------------------------------------------------------------------------------------------
' Add additional declarations, types, constants & enumerations here:


'-------------------------------------------------------------------------------------------------
' The following are existing properties on a command button. They are
' for reference only. You would need to add property code for these &
' also read/cache them in the ReadProperties & WriteProperties events.
' Those you don't need, simply delete/rem them out. Otherwise, each
' of these should map to the the same property in your UserControl

' Tip. If adding these to your custom button, most do not need to be cached
' as separate variables if they will be applied to your usercontrol. Instead,
' set and cache the property directly to the usercontrol. For an example, see the
' coded Public Enabled Property, Usercontrol_ReadProperty & WriteProperty routines

'Private m_Appearance As Integer   ' either 0=Flat, 1=3D
'Private m_BackColor As Long
'Private m_Font As StdFont
'Private m_Picture As StdPicture
'Private m_DisabledPicture As StdPicture
'Private m_DownPicture As StdPicture
'Private m_MouseIcon As StdPicture
'Private m_MousePointer As MousePointerConstants
'Private m_OLEdropMode As Integer ' either 0=None, 1=Manual
'' The following are also command button properties but are
'' properties not exposed in the IDE property page.
'Private m_FontBold As Boolean       ' m_Font.Bold
'Private m_FontIalic As Boolean      ' m_Font.Italic
'Private m_FontName As String        ' m_Font.Name
'Private m_FontSize As Single        ' m_Font.Size
'Private m_FontStrikethru As Boolean ' m_Font.Strikethrough
'Private m_FontUnderline As Boolean  ' m_Font.Underline
'-------------------------------------------------------------------------------------------------


' BUTTON TEMPLATE CODE BETWEEN SLASHES -- DO NOT DELETE
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Note: You may not need nor want to expose every Public event to the user.
' The below events are the same ones that a VB Command Button exposes.
' Simply remove the ones you don't want & also remove any coded RaiseEvent calls to those events

Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event Click() ' note: not fired in Usercontrol_Click event. Fired in UpdateState routine because...
' About click events> CmdButton: mouseDown,Click,mouseUp. UC: mouseDown,mouseUp,Click
' To fix, we control when, and in what order, events are sent to the user

Public Event MouseEnter()
Public Event MouseLeave()
' about MouseEnter & MouseLeave
' There seems to be no hard & fast rule regarding when to fire this & when not to
' So, here are the simple rules I have applied
'   1) Send MouseEnter whenever mouse enters the control except when...
'       - The SpaceBar button is currently being held down on the control
'       - There already has been a MouseEnter sent with no MouseLeave sent
'       - Control is disabled
'   2) Send MouseLeave whenever the mouse exits the control except when...
'       - Any mouse button is currently being held down on the control
'       - No MouseEnter was previously sent
'       - Control is disabled
Public Event DblClick() ' not a standard command button event
' ^^ like VB buttons, a dblclick will send two click events, but then this DblClick event is sent afterwards

' Much appreciation goes towards Paul Caton for his self-subclassing thunks; makes some things so much easier
'-Thunking/Callback declarations---------------------------------------------------------------------------
Private z_CbMem   As Long    'Callback allocated memory address
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC As Long = -4
'-------------------------------------------------------------------------------------------------

' Caption rendering APIs/constants
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Const DT_CALCRECT As Long = &H400
Private Const DT_NOCLIP As Long = &H100
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_MULTILINE As Long = (&H1)
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_CENTER As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_RIGHT As Long = &H2

Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32.dll" () As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
' these constants are used to simply distinguish the types of actions that effect button state
Private Const WM_ACTIVATEAPP As Long = &H1C ' application is gaining/losing focus to another window
Private Const WM_CHAR As Long = &H102       ' button's accelerator key was pressed
Private Const WM_ENABLE As Long = &HA       ' button is being enabled/disabled
Private Const WM_KILLFOCUS As Long = &H8    ' button is losing focus
Private Const WM_KEYDOWN As Long = &H100    ' key down event is occuring on the button
Private Const WM_KEYUP As Long = &H101      ' key up event is occuring on the button
Private Const WM_LBUTTONDOWN As Long = &H201 ' left mouse button is being pressed on the button
Private Const WM_LBUTTONUP As Long = &H202  ' left mouse button is being released on the button
Private Const WM_MOUSEHOVER As Long = &H2A1 ' mouse is entering the button's boundaries
Private Const WM_MOUSELEAVE As Long = &H2A3 ' mouse is leaving the button's boundaries
Private Const WM_MOUSEMOVE As Long = &H200  ' mouse is moving over the button
Private Const WM_PAINT As Long = &HF&       ' the button is to be completely repainted
Private Const WM_SETFOCUS As Long = &H7     ' the button is gaining focus
Private Const WM_SHOWWINDOW As Long = &H18  ' the button is being made visible/invisible
Private Const SWP_FRAMECHANGED As Long = &H20 ' the button's border is changing due to focus events
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Enum eBtnStates ' contains graphical, tracking & input flags
    bsNormal = 0        ' graphical state: draw normal
    bsPushed = 1        ' graphical state: draw as down
    bsHover = 2         ' graphical state: draw as mouse over
    bsFocus = 4         ' tracking/graphical state: focus rectangle
    bsDefaultBtn = 8    ' tracks whether control drawn as default (Ambient.DisplayAsDefault)
    bsMouseEntered = 32 ' tracking: MouseEnter message was sent
    bsHide = 64         ' tracking: Usercontrol.Hide event triggered
    bsAppNoFocus = 128  ' tracking: application lost focus
    bsOnClick = 2048    ' tracks the Value property as True/False
    bsKeyDown = 1024    ' input state: spacebar is treated down
    bsMouseOver = 512   ' input state: mouse is over button
    bsMouseDown = 256   ' input state: left mouse button is treated as down
    bsMaskMouseBtns = 7 ' mask for key mouse states (vbLeftButton,vbMiddleButton,vbRightButton)
    bsMaskGraphicalState = 15 'contains current graphical state (bsNormal,bsPushed,bsHover,bsFocus,bsDefaultBtn)
    bsMaskBtnState = 3  ' general state mask (bsNormal,bsPushed,bsHover)
    bsDblClick = 8      ' double clicked. Added Event. Not standard CmdButton event
End Enum

Private Enum eDrawState
    bdNormal = 0
    bdPushed = 1
    bdHover = 2
    bdDisabled = -1
End Enum
Private Enum eDrawAction
    baDrawEntire = 0
    baDrawFocusOnly = 1
    baDrawDefaultBdrOnly = 2
End Enum
Private Enum eAttributes
    attrHasFocus = 1
    attrIsDefaultBtn = 2
    attrMouseIsOver = 4
End Enum

Private m_Caption As String ' << coded, do not remove; needed should button use accelerators
Private m_Exclusions As Long    ' See UserControl_Initialize
Private m_pHwnd As Long         ' parent window
Private m_TimerActive As Long   ' active timer(s)
Private m_timerProc As Long     ' callback procedure for the timer (See TimerProc for purpose/alternatives)
Private m_State As eBtnStates   ' calculated in UpdateState routine
Private m_MouseState As Long ' contains one or more of the following:
                             ' vbLeftButton,vbRightButton,vbMiddleButton that are currently down
                             ' eBtnStates.bsDblClick if double click event occurred
                             ' most recent mouse button down * &H100
                             ' most recent mouse shift constants * &H10000
    
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

Public Property Let ShowFocusRect(bShow As Boolean)
    If Not bShow = ShowFocusRect Then   ' only modify if changinge
        ' remove existing state if any, then add new state
        ' Remember exclusions are complimentary. Therefore, if we want to show
        ' rect, we remove bsFocus else if we don't want to show it, we add bsFocus
        m_Exclusions = ((m_Exclusions And Not bsFocus) Or (Abs(Not bShow) * bsFocus))
        If ((m_State And bsFocus) = bsFocus) Then
            ' if button already has the focus, then flag is set. Remove it & update
            m_State = (m_State And Not bsFocus)
            Me.Refresh
        End If
        PropertyChanged "ShowFocusRect"
    End If
End Property
Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = CBool((m_Exclusions And bsFocus) = 0&)
End Property

Public Sub Refresh()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    UpdateState WM_PAINT, bsNormal
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Public Property Let Enabled(Enable As Boolean)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If Not UserControl.Enabled = Enable Then 'changing property value
        UserControl.Enabled = Enable
        UpdateState WM_ENABLE, Enable   ' clean up timers if needed & redraw
        PropertyChanged "Enabled"
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Property
Public Property Get Enabled() As Boolean
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Enabled = UserControl.Enabled
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Property

Public Property Let Caption(NewCaption As String)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ' Your button may be 100% graphical with no caption. However, exposing this
    ' property allows a user to still create a caption with an accelerator/shortcut
    ' where that shortcut activates the button; you simply don't need to display
    ' the caption in that case. If you absolutely don't want/need this property,
    ' also remove the related line of code in UserControl_Read/WriteProperty
    On Error GoTo ExitRoutine
    If StrComp(NewCaption, m_Caption, vbBinaryCompare) = 0 Then Exit Property
    Dim iChar As Integer, iAmp As Integer
    If Not NewCaption = vbNullString Then
        For iChar = 1 To Len(NewCaption)
            If Mid$(NewCaption, iChar, 1) = "&" Then
                If iAmp = 0 Then ' no previous ampersands
                    iAmp = iChar
                Else
                    ' if previous char was ampersand then not an accelerator
                    If iAmp = iChar - 1 Then iAmp = 0 Else iAmp = iChar
                End If
            End If
        Next
        If iAmp = iChar - 1 Then iAmp = 0 ' cannot be the last character
    End If
ExitRoutine:
    If iAmp = 0 Then
        UserControl.AccessKeys = vbNullString
    Else
        UserControl.AccessKeys = Mid$(NewCaption, iAmp + 1, 1)
    End If
    m_Caption = NewCaption
    If m_Exclusions = -1 Then Exit Property ' m_Exclusions set to -1 in ReadProperties routine
    PropertyChanged "Caption"
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    ' add your code here to update/draw the caption on your button. Can call Me.Refresh
    Me.Refresh

End Property
Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Caption = m_Caption
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Property

Public Property Let Value(ByVal ClickIt As Boolean)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If ClickIt Then
        ' prevent stack overflow. Calling Command1.Value=True inside Command1_Click
        ' event will cause a stack overflow. We want to replicate VB but not to the
        ' extremes that we are willing to replicate a design flaw.
        If Me.Value = False Then
            m_State = m_State Or bsOnClick
            RaiseEvent Click
            m_State = m_State And Not bsOnClick
        End If
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Property

Public Property Get Value() As Boolean
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Value = CBool(m_State And bsOnClick)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Property




Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    UpdateState WM_CHAR, KeyAscii ' user pressed ALT+accessKey
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Select Case PropertyName
        ' buttons should draw a different border when they have focus and when they
        ' do not. This is called by VB whenever the button should be changed to
        ' identify it is the default button or has/lost focus.
        Case "DisplayAsDefault": UpdateState SWP_FRAMECHANGED, bsNormal
        Case Else
            ' add any other ambient property changes you want to track
    End Select
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_DblClick()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Dim Button As Integer, mPT As POINTAPI, iShift As Integer
    ' When command button is double clicked, VB sends
    '   -- mouseDown, Click, mouseUp,   mouseDown, Click, mouseUp
    ' but when DblClick occurs in usercontrols, VB sends this:
    '   -- mouseDown, mouseUp, Click,   DblClick, mouseUp
    
    ' So, to send the missing mouse down & click events, we need to do it manually
    ' Also we will place the Click events between the Down & Up events like a cmdButton
    
    ' cmdButton's only fire dblClick when the left button did the double clicking
    Button = ((m_MouseState \ &H100) And &HFF)  ' determine mouse button firing this event
    iShift = (m_MouseState \ &H10000)           ' shift values when double clicked
    
    ' get mouse coords relative to the client area
    GetCursorPos mPT
    ScreenToClient UserControl.hWnd, mPT
    ' send the missing event, but don't allow Button to be m odified
    RaiseEvent MouseDown(Button + 0, iShift, Int(ScaleX(mPT.x, vbPixels, vbContainerPosition)), Int(ScaleY(mPT.y, vbPixels, vbContainerPosition)))
    If Button = vbLeftButton Then
        UpdateState WM_LBUTTONDOWN, bsNormal ' send mousedown so control draws down state
        m_MouseState = m_MouseState Or bsDblClick  ' include double click event
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_GotFocus()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    UpdateState WM_SETFOCUS, bsFocus
    ' See UserControl_Initialize to prevent receiving paint notification for change
    ' in Focus state, but do not rem out the statement.
    ' Focus notification is needed for other purposes too.
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_LostFocus()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    m_MouseState = 0&
    UpdateState WM_KILLFOCUS, bsFocus
    ' See UserControl_Initialize to prevent receiving paint notification for change
    ' in Focus state, but do not rem out the statement.
    ' Focus notification is needed for other purposes too.
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If Not KeyCode = vbKeyReturn Then RaiseEvent KeyDown(KeyCode, Shift + 0)
    Select Case KeyCode
        Case vbKeySpace ' If Alt is down without Ctrl, then do not forward
            If ((Shift And vbAltMask) = 0) Or ((Shift And vbCtrlMask) = vbCtrlMask) Then UpdateState WM_KEYDOWN, KeyCode
        Case vbKeyReturn
            ' Return=Click unless Alt, Ctrl and/or Shift is held down
            If Shift = 0 Then UpdateState WM_KEYDOWN, KeyCode
        Case Else
            UpdateState WM_KEYDOWN, KeyCode
    End Select
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If Not KeyCode = vbKeyReturn Then
        RaiseEvent KeyUp(KeyCode, Shift)
        UpdateState WM_KEYUP, KeyCode ' may have caused Click event if not previously canceled
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If Not KeyAscii = vbKeyReturn Then RaiseEvent KeyPress(KeyAscii)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ' send event converting our scale units to host's
    RaiseEvent MouseMove(Button + 0, Shift, Int(ScaleX(x, UserControl.ScaleMode, vbContainerPosition)), Int(ScaleY(y, UserControl.ScaleMode, vbContainerPosition)))
    If (m_TimerActive And 1) = 0& Then UpdateState WM_MOUSEHOVER, bsNormal
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    m_MouseState = ((m_MouseState And bsMaskMouseBtns) Or Button) Or (Button * &H100) ' track button & last button
    m_MouseState = m_MouseState Or (Shift * &H10000)                      ' track last shift values
    If Button = vbLeftButton Then UpdateState WM_LBUTTONDOWN, bsNormal ' changes graphical state
    ' send event converting our scale units to host's
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(x, UserControl.ScaleMode, vbContainerPosition)), Int(ScaleY(y, UserControl.ScaleMode, vbContainerPosition)))
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ' send event converting our scale units to host's
    m_MouseState = (m_MouseState And Not Button)    ' remove button from state
    If Button = vbLeftButton Then UpdateState WM_LBUTTONUP, bsNormal ' can fire a click event
    If (m_State And bsHide) = 0& Then
        RaiseEvent MouseUp(Button + 0, Shift, Int(ScaleX(x, UserControl.ScaleMode, vbContainerPosition)), Int(ScaleY(y, UserControl.ScaleMode, vbContainerPosition)))
        If (m_MouseState And bsDblClick) = bsDblClick Then   ' trigger dblClick if appropriate
            m_MouseState = (m_MouseState And bsMaskMouseBtns)      ' remove dblClick flag and any other flags
            If (m_State And bsMouseOver) = bsMouseOver Then RaiseEvent DblClick ' fire event
        End If
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
' add your code here as needed
    
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    RaiseEvent OLECompleteDrag(Effect)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, Int(ScaleX(x, UserControl.ScaleMode, vbContainerPosition)), Int(ScaleY(y, UserControl.ScaleMode, vbContainerPosition)))
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, Int(ScaleX(x, UserControl.ScaleMode, vbContainerPosition)), Int(ScaleY(y, UserControl.ScaleMode, vbContainerPosition)), State)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    RaiseEvent OLESetData(Data, DataFormat)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Exclusions = -1 ' prevent triggering dirty property when setting Me.Caption
    Me.Caption = PropBag.ReadProperty("Caption", vbNullString)
    m_Exclusions = PropBag.ReadProperty("Exclusions", 0&)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True  ' save Enabled state
    PropBag.WriteProperty "Caption", Me.Caption, vbNullString   ' save Caption
    PropBag.WriteProperty "Exclusions", m_Exclusions, 0&        ' save ShowFocusRect
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_Paint()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    Me.Refresh    ' Call local Refresh routine. Note: UserControl_Paint is not called if AutoRedraw=True
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_Show()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (m_State And bsHide) = bsHide Then   ' Usercontrol_Hide event was triggered
        UpdateState WM_SHOWWINDOW, bsNormal ' now it isn't; simply reset property
    Else
        If GetFocus() = UserControl.hWnd Then ' ensure we have focus state set
            If (m_Exclusions And bsFocus) = 0& Then m_State = (m_State And bsFocus)
            Me.Refresh    ' Call local Refresh routine
        Else
            If UserControl.AutoRedraw = True Then Me.Refresh    ' Call local Refresh routine
        End If
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_Hide()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    UpdateState WM_SHOWWINDOW, bsHide
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub UserControl_Initialize()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    m_timerProc = zb_AddressOf(1, 4) ' AddressOf of our TimerProc at end of module
    m_Exclusions = 0&
' Adding bsFocus to m_Exclusions prevents the DrawButton routine
'       from being called simply because the control got/lost focus.
'       i.e., m_Exclusions = m_Exclusions Or bsFocus
'       This does not prevent LostFocus & GotFocus events from firing
' Adding bsMouseOver to m_Exclusions prevents the Drawbutton routine
'       from being called simply because the mouse entered/exited the control.
'       This does not prevent MouseEnter & MouseLeave events from firing.
'       i.e., m_Exclusions = m_Exclusions Or bsMouseOver
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub

Private Sub UserControl_Terminate()
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    ManageTimers False, 1
    ManageTimers False, 2
    zb_Terminate
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

' add your code here as needed

End Sub


Private Sub UpdateState(ByVal stateMessage As Long, ByVal lParam As Long)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

    ' Function determines the graphical and tracking state of the button and also
    ' may send a Click, MouseEnter and/or a MouseLeave event.
    ' This is a bit lengthy only because a button's graphical state is dependent
    ' on so many variables: which keys are down, which mouse button is down,
    ' whether or not the mouse is over the button, whether or not the
    ' application has focus, button has focus, etc, etc, etc.
    
    ' The very end of the function is the purpose of deciphering all these
    ' conditions: It only sends a Redraw when the overall graphical state
    ' changes. Since drawing custom buttons takes the most time, this routine
    ' can prevent many unnecessary paints. There are only 4 basic states:
    ' up, down, mouse-over, disabled where up & mouse-over can have focus,
    ' down will always have focus & disabled never will.
    
    ' The following scenarios fire a repaint action (a call to DrawButton):
    ' 1. Control receives/loses focus, unless excluded
    '    - This allows adding/removing focus rectangle/graphics as needed
    '    - See UserControl_Initialize to prevent this
    ' 2. Mouse moves in or out of the control, unless excluded
    '    - This allows adding "MouseOver" graphics as needed
    '    - See UserControl_Initialize to prevent this
    ' 3. Control receives a left button click or the space bar is pressed
    '    - This allows drawing the control in a down position
    ' 4. Control changes from any non-Normal state to Normal
    '    - This allows drawing the control in an up/normal position
    ' 5. Sending a WM_Enable, WM_Paint, or WM_ShowWindow stateMessage

    If UserControl.Enabled = False Then
        ' no messages should be processed for disabled buttons; however, should
        ' this routine get any messages, other then the few below, while the
        ' control is disabled, then those messages will be ignored
        m_MouseState = 0&
        If Not stateMessage = WM_PAINT Then
            If Not stateMessage = WM_SHOWWINDOW Then
                If Not stateMessage = WM_ENABLE Then Exit Sub
            End If
        End If
    End If
    
    Dim oldState As Long
    Dim btnAction As eDrawAction
    Dim btnAttr As eAttributes
    Dim btnEvent As Long
    ' ^ 0=no event, 1=click event, 2=mouseEnter, 3=mouseLeave, 4=add/remove focus rect, 8=add/remove default button border
    
    oldState = (m_State And bsMaskGraphicalState) ' get current graphical state only (1st 4 bits)
    
    Select Case stateMessage
    Case WM_MOUSELEAVE  ' only called by TimerProc
        If (m_State And bsMouseEntered) = bsMouseEntered Then
            If (m_MouseState And bsMaskMouseBtns) = 0& Then ' no mouse button held down
                ManageTimers False, 1 ' kill timer
                btnEvent = 3          ' send MouseLeave
            End If
        End If
        m_State = (m_State And Not bsMouseOver) ' remove flag
    
    Case WM_MOUSEHOVER  ' called by TimerProc and UserControl_MouseMove
        If (m_TimerActive And 1) = 0& Then ManageTimers True, 1 ' activate a timer if needed
        If (m_State And bsMouseEntered) = 0& Then    ' was over the control
            If (m_MouseState And bsMaskMouseBtns) = 0& Then     ' no mouse button held down
                ' also see if the spacebar is held down
                If (m_State And bsKeyDown) = 0& Then btnEvent = 2 ' send MouseEnter
            End If
        End If
        m_State = m_State Or bsMouseOver    ' add flag
        
    Case WM_LBUTTONDOWN ' called by UserControl_MouseDown & UserControl_DblClick
        m_State = m_State Or bsMouseDown  ' down state via mouse
        
    Case WM_LBUTTONUP   ' only called by UserControl_MouseUp
        ' send click event only if the mouse is over the control, was previously clicked
        ' and the spacebar is not being held down
        If (m_State And bsMouseDown) = bsMouseDown Then
            If (m_State And bsMouseOver) = bsMouseOver Then btnEvent = 1
            If (m_State And bsKeyDown) = 0& Then
                If GetCapture() = UserControl.hWnd Then ReleaseCapture
            End If
            m_State = (m_State And Not bsMouseDown)   ' remove flag
        End If
        
    Case WM_KEYDOWN     ' called by UserControl_AccessKeyPress & UserControl_KeyDown
        Select Case lParam     ' which key?
        Case vbKeyReturn
            btnEvent = 1 ' pressing Return key does not change graphical state
        Case vbKeySpace
            ' only add state if the left mouse button is not being held down
            ' If it is, then buttons do not record clicks via spacebar
            If (m_State And bsMouseDown) = 0& Then
                m_State = m_State Or bsKeyDown  ' down state via keyboard
                SetCapture UserControl.hWnd     ' button keeps capture while spaceBar is down
            End If
        Case Else
            ' all other keys cancel spacebar action if spacebar is down except ALT which just releases capture
            If (m_State And bsMouseDown) = 0& Then
                If GetCapture() = UserControl.hWnd Then ReleaseCapture
            End If
            If Not lParam = vbKeyMenu Then m_State = (m_State And Not bsKeyDown)
        End Select
        
    Case WM_KEYUP   ' only called by UserControl_KeyUp
        If (m_State And bsMouseDown) = 0& Then ' release capture if mouse doesn't have it
            If GetCapture() = UserControl.hWnd Then ReleaseCapture
        End If
        If lParam = vbKeySpace Then ' click event only if the spacebar wasn't canceled previously
            If (m_State And bsKeyDown) = bsKeyDown Then
                ' however if left mouse is still down, then don't fire event
                If (m_State And bsMouseDown) = 0& Then btnEvent = 1
            End If
        End If
        m_State = (m_State And Not bsKeyDown) ' remove flag
        
    Case WM_SETFOCUS    ' only called by UserControl_GotFocus
        ' set Focus flag unless exlcuded. See UserControl_Initialize
        If (m_Exclusions And bsFocus) = 0& Then
            m_State = m_State Or bsFocus
            btnEvent = 4
        End If
        ' find our form's hWnd
        lParam = UserControl.ContainerHwnd
        m_pHwnd = 0&         ' Why do this everytime? If permanently cached the 1st time, it
        Do Until lParam = 0& ' will not be correct if: Set CustomButton.Container=SomethingElse
            m_pHwnd = lParam
            lParam = GetParent(m_pHwnd)
        Loop
        ManageTimers True, 2  ' set the timer to track app lost/got focus
    
    Case WM_KILLFOCUS   ' only called by UserControl_LostFocus
        ManageTimers False, 2   ' kill timer
        If GetCapture() = UserControl.hWnd Then ReleaseCapture
        If (m_State And bsKeyDown) = bsKeyDown Then
            ' when spacebar is held down on control and control loses focus
            ' the control should send a click event; unless the left mouse
            ' button is also held down on it
            If (m_State And bsMouseDown) = 0& Then btnEvent = 1
        End If
        ' losing focus releases all flags except the DefaultBtn, MouseEntered & MouseOver flags if they are set
        ' i.e., mouse button can be down & mouse over control & user hits the Tab key to lose focus
        m_State = (m_State And bsMouseEntered) Or (m_State And bsMouseOver) Or (m_State And bsDefaultBtn)
        btnEvent = btnEvent Or 4
    
    Case WM_ACTIVATEAPP  ' only called by TimerProc
        m_MouseState = 0&
        If lParam = 0& Then ' lost focus
            If GetCapture() = UserControl.hWnd Then ReleaseCapture
            m_State = ((m_State And Not bsMaskBtnState) Or bsAppNoFocus)
            If (m_State And bsFocus) = bsFocus Then oldState = True
            ' ^^ force a mismatch so redraw occurs if control has focus
        Else    ' else got focus
            m_State = (m_State And Not bsAppNoFocus)
            If GetFocus() = UserControl.hWnd Then
                If (m_State And bsFocus) = bsFocus Then oldState = True ' force a mismatch so redraw occurs
            Else
                ManageTimers False, 2 ' kill timer
            End If
        End If
        
    Case WM_ENABLE  ' only called by the Enabled Property
        If lParam = 0& Then ' disabling
            If GetCapture() = UserControl.hWnd Then ReleaseCapture
            ManageTimers False, 1
            ManageTimers False, 2
            m_State = 0&
        End If
        oldState = True  ' force a mismatch so redraw occurs
        
    Case WM_SHOWWINDOW  ' called by UserControl_Show & UserControl_Hide
        If lParam = 0& Then
            m_State = (m_State And Not bsHide) ' remove flag
        Else
            ' UserControl_Hide event was called; probably closing
            m_State = m_State Or bsHide ' add flag
        End If
        
    Case WM_PAINT   ' called by the UserControl.Refresh method
        ' Generic "Refresh", simply forces a call to DrawButton
        oldState = True
        
    Case WM_CHAR    ' Alt+AccessKey pressed
        If Not lParam = vbKeyReturn Then m_State = (m_State And Not bsKeyDown) ' release the spacebar flag as needed
        btnEvent = 1
    
    Case SWP_FRAMECHANGED
        btnEvent = 8 ' toggle the border as focused / not focused
        If Ambient.DisplayAsDefault = True Then
            m_State = m_State Or bsDefaultBtn
        Else
            m_State = (m_State And Not bsDefaultBtn)
        End If
    
    Case Else
        Exit Sub    ' something you added, that I have not coded for?
    End Select
    
    ' now, let's determine the graphical state of the button
    ' These are in order of priority, rearranging them produces invalid graphical states
    If Not m_State = bsNormal Then
    
        m_State = (m_State And Not bsMaskBtnState) ' cache tracking & input flags only; graphical state is next
        
        If (m_State And bsMouseDown) = bsMouseDown Then
            ' the left mouse button is down, two states possible
            If (m_State And bsMouseOver) = bsMouseOver Then m_State = m_State Or bsPushed
            ' ^ if cursor is over button, we show it as pushed else as Normal
            
        ElseIf (m_State And bsKeyDown) = bsKeyDown Then
            m_State = m_State Or bsPushed ' spacebar is down; only one state possible
            
        ElseIf (m_State And bsMouseOver) = bsMouseOver Then
            ' as long as hover state is not excluded ...
            If (m_Exclusions And bsMouseOver) = 0& Then m_State = m_State Or bsHover
            ' ^ if not down but mouse is over, hover state
        Else
            ' if all above don't trigger, then it is normal state
        End If
    End If
    
    ' if the state changed, notify drawing routine
    If (m_State And bsHide) = 0& Then   ' else Usercontrol_Hide event is in effect
        If Not (m_State And bsMaskGraphicalState) = oldState Then    ' compare focus,up,down,hover state to previous state
            If (m_State And bsAppNoFocus) = bsAppNoFocus Then
                oldState = (m_State And Not bsFocus)  ' if app doesn't have focus, neither should the control
            Else
                oldState = m_State
            End If
            If (oldState And bsFocus) = bsFocus And UserControl.Enabled = True Then btnAttr = attrHasFocus
            If (m_State And bsDefaultBtn) Then btnAttr = btnAttr Or attrIsDefaultBtn
            If (m_State And bsMouseOver) Then btnAttr = btnAttr Or attrMouseIsOver
            If btnEvent = 4 Then
                btnAction = baDrawFocusOnly
            ElseIf btnEvent = 8 Then
                btnAction = baDrawDefaultBdrOnly
            End If
            ' by Or'ing (Not UserControl.Enabled) we get -1 when the control is Disabled.
            DrawButton ((m_State And bsMaskBtnState) Or (Not UserControl.Enabled)), btnAction, btnAttr
        End If
    End If

    ' if an event is to be fired, fire it now
    Select Case (btnEvent And &H3)
        Case 1:
            m_State = m_State Or bsOnClick
            RaiseEvent Click ' click event was fired
            m_State = (m_State And Not bsOnClick)
        Case 2: ' cache mouseEnter being fired, so we can clear it later if needed
                m_State = m_State Or bsMouseEntered
                RaiseEvent MouseEnter
        Case 3: ' remove the MousEnter flag
                m_State = (m_State And Not bsMouseEntered)
                RaiseEvent MouseLeave
        Case Else
    End Select
    

' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub ManageTimers(bSet As Boolean, ByVal TimerID As Long)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If bSet = True Then
        ' See TimerProc also.
        m_TimerActive = m_TimerActive Or TimerID
        If TimerID = 1 Then     ' #1 used for mouse enter/mouse leave
            ' active whenever a MouseEnter event is detected &
            ' removed when the MouseLeave event is detected
            SetTimer UserControl.hWnd, 1, 80, m_timerProc ' 12.5x a second
        Else                    ' #2 used for app lost/got focus
            ' active whenever a control has the focus
            ' removed when control loses the focus within the parent
            SetTimer UserControl.hWnd, 2, 750, m_timerProc ' 1.25x a second
        End If
    Else
        ' remove timer(s) as needed, update active status
        If (m_TimerActive And TimerID) = TimerID Then
            KillTimer UserControl.hWnd, TimerID
            m_TimerActive = m_TimerActive And Not TimerID
        End If
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

Private Sub DrawButton(ByVal GraphicalState As eDrawState, ByVal Action As eDrawAction, ByVal Attributes As eAttributes)
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    ' This routine is called with specific flags to help prevnt unnecessary and redundant drawing
    ' The routine is not called if no changes to the graphical state occurs.
    ' If you do not want mouse enter/leave events to call this event, then set the m_Exclusion flags appropriately (see UserControl_Initialize)
    
    ' The parameters will help you decide what needs to be painted and what does not
    ' GraphicalState
    '   :: bdNormal. button is to be drawn in the up state
    '   :: bdPushed. button is to be drawn in the down state
    '   :: bdHover. button is to be drawn in the mouse-over state
    '   :: bdDisabled. button is to be drawn as disabled, normal state
    ' Action
    '   :: baDrawEntire then entire button, including borders & focus rectangle are to be drawn
    '       -- this event occurs due to refresh and changes in the GraphicalState
    '   :: baDrawDefaultBdrOnly. the button's border should be drawn to show it has focus
    '   :: baDrawFocusOnly. the button's focus rectangle is to be added/removed
    '       -- see the Attributes parameter to determine if button has focus or not
    ' Attributes
    '   :: attrHasFocus. If the button has focus Attributes includes this style
    '   :: attrIsDefaultBtn. If the button should be drawn with a border to identify it as the default button, this style is included
    '   :: attrMouseIsOver. If the mouse is currently over the button, this style is included
    
    ' Rendering notes
    ' If you are going to provide multiple styles to your buttons, you probably want to separate key portions of the rendering
    ' into separate routines to make troubleshooting easier and coding more modular. For example, you can use...
    '   DrawButtonBkg
    '   DrawButtonImage
    '   DrawButtonText
    '   DrawButtonBorders
    '   DrawFocusRectangle
    
    
    If Action = baDrawEntire Then
        
        Select Case GraphicalState
            Case bdNormal:  ' call your routine to draw button up state
    
    
            Case bdPushed:  ' call your routine to draw button down state
                
                
            Case bdHover:   ' call your routine to draw button hover state (mouse over)
                ' Note: If you are not displaying a mouse-over image, see UserControl_Initialize
                
            Case bdDisabled: ' call your routine to draw disabled button. HasFocus is always False
        
        End Select
        
        ' ******************************************************************************************
        ' Draw our Bkg, Caption, Icon, etc
        ' ******************************************************************************************
    
        ' ******************************************************************************************
        ' Draw your button border(s)
        ' ******************************************************************************************
        If GraphicalState = bdPushed Then
            ' draw borders for down state
            
        ElseIf (Attributes And attrIsDefaultBtn) Then
            ' Note about the Default property. If set you should draw a border as if the control had focus
            ' draw with "focus" type border
            
        Else ' up/hover/disabled without focus/default button state
            ' draw borders without "focus" type border
        
        End If
        
        ' ******************************************************************************************
        ' Draw your focus rectangle as needed
        ' ******************************************************************************************
        If (Attributes And attrHasFocus) Then
        ' If needed, check the Attributes to determine if the control has focus or not
        
        End If
            
    
    ElseIf Action = baDrawDefaultBdrOnly Then
    
        ' ******************************************************************************************
        ' Draw button border
        ' ******************************************************************************************
        ' Note about the Default property. If set you should draw a border as if the control had focus
        If (Attributes And attrIsDefaultBtn) Then
            ' draw with "focus" type border
        Else
            ' draw without "focus" type border
        End If

    ElseIf Action = baDrawFocusOnly Then
    
        ' ******************************************************************************************
        ' Draw focus rectangle if needed
        ' ******************************************************************************************
        ' If needed, check the Attributes to determine if the control has focus or not
        
    End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub


'-Callback code-----------------------------------------------------------------------------------
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
Const MAX_FUNKS   As Long = 1                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  Dim z_Cb()      As Long
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  
    z_Cb(3, nThunkNo) = _
              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, nThunkNo) = &HBB60E089
    z_Cb(5, nThunkNo) = z_CbMem                                             'Set the data address
    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
  
  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
  If bIdeSafety Then                                                        'If the user wants IDE protection
    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
  End If
    
  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
  zb_AddressOf = nAddr + 16                                                 'Thunk code start address
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function

'Terminate the callback thunks
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Private Sub zb_Terminate()
Const MEM_RELEASE As Long = &H8000&                                         'Release allocated memory flag

  If z_CbMem <> 0 Then                                                      'If memory allocated
    If VirtualFree(z_CbMem, 0, MEM_RELEASE) <> 0 Then                       'Release
      z_CbMem = 0                                                           'Indicate memory released
    End If
  End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  J = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < J
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function

'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************
'Callback ordinal 3 -- increment zb_AddressOf's MAX_FUNKS before adding another procedure

'Callback ordinal 2 -- increment zb_AddressOf's MAX_FUNKS before adding another procedure

'Callback ordinal 1
Private Function TimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
On Error GoTo EH

    ' why use a timer?  Why not use TrackMouseEvent for mouse enter/leave?  Why not subclass parent for loss/gain of application focus?
    ' Timers are safe and can be used in all scenarios
    
    ' For Mouse Enter/Leave events, TrackMouseEvent can be used only if this button is not windowless.
    '   Timer supports windowless controls if code is modified a tad as noted below.
    
    ' For the application losing/gaining focus, you can subclass the Parent.hWnd, however, subclassing
    ' in VB is hazardous, especially when multiple usercontrols subclass the same parent.  If subclassing, you can
    ' test for the WM_APPActivate message.  Timer is not as responsive but a lot safer.

    If TimerID = 1 Then ' mouse enter/leave timer
        ' Note: For windowless controls, you will want to pass the mPT coords
        ' to a HitTest function to determine if mouse is over the windowless control.
        ' Then call UpdateState depending on whether or not the mouse is over windowless button
        Dim mPT As POINTAPI
        GetCursorPos mPT
        If WindowFromPoint(mPT.x, mPT.y) = hWnd Then
            UpdateState WM_MOUSEHOVER, bsNormal
        Else
            UpdateState WM_MOUSELEAVE, bsNormal
        End If
    Else                ' app losing focus timer
        If GetForegroundWindow() = m_pHwnd Then
            If (m_State And bsAppNoFocus) = bsAppNoFocus Then UpdateState WM_ACTIVATEAPP, 1&
        Else
            If (m_State And bsAppNoFocus) = 0 Then UpdateState WM_ACTIVATEAPP, 0&
        End If
    End If

EH:
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
'          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
'          of a new function/sub and the callbacks will not be fired & maybe GPF.
End Function
