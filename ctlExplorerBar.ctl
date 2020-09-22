VERSION 5.00
Begin VB.UserControl XPexplorerbar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   5505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   MouseIcon       =   "ctlExplorerBar.ctx":0000
   ScaleHeight     =   5505
   ScaleWidth      =   3495
   ToolboxBitmap   =   "ctlExplorerBar.ctx":0CCA
   Begin VB.VScrollBar Scoller 
      Height          =   735
      LargeChange     =   100
      Left            =   360
      SmallChange     =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "XPexplorerbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'   WindowsXP ExplorerBar
'   Copyright Flex 2005 - flex4d@gmail.com
'
'   I hope you like this control... :o)
'
'   This control is also working in Windows 9x,
'   but with a classic look.
'   I looked at Win 9x to make that classic look.
'
'   Sorry for the non-commented code...
'
'   License:
'   Flex is not responsible for any caused damage that
'   may occur with this program.
'   Editing is at your own risk.
'   If you use this, I would like you put my name in
'   the credits.
'
'   NOTE:
'   For transparant images, I recommend .ICO files
'   Because most .GIF don't work!

'Version
Const Version As String = "2"

'Event
Public Event HeaderClick(ByVal Index As Integer)
Public Event SubItemClick(ByVal Index As Integer, ByVal SubItemIndex As Integer)
Public Event Collapse(ByVal Index As Integer)
Public Event Expand(ByVal Index As Integer)
Public Event MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Public Event MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)
Public Event MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Long, ByVal Y As Long)

'Declarations
Private Items() As MenuItem
Private ItemCount As Integer

'Constants
Private Const EBP_HEADERBACKGROUND As Integer = 1
Private Const EBP_NORMALGROUPBACKGROUND As Integer = 5
Private Const EBP_NORMALGROUPCOLLAPSE As Integer = 6
Private Const EBP_NORMALGROUPEXPAND As Integer = 7
Private Const EBP_NORMALGROUPHEAD As Integer = 8
Private Const EBP_SPECIALGROUPBACKGROUND As Integer = 9
Private Const EBP_SPECIALGROUPCOLLAPSE As Integer = 10
Private Const EBP_SPECIALGROUPEXPAND As Integer = 11
Private Const EBP_SPECIALGROUPHEAD As Integer = 12

Private Const STATE_NORMAL As Long = 1
Private Const STATE_HOT As Long = 2
Private Const STATE_PRESSED As Long = 3

Private Const MENU_NORMAL As Long = 1
Private Const MENU_SPECIAL As Long = 2
Private Const MENU_DETAILS As Long = 3

'Enums
Private Enum DrawTextFlags
    DT_TOP = &H0
    DT_LEFT = &H0
    DT_CENTER = &H1
    DT_RIGHT = &H2
    DT_VCENTER = &H4
    DT_BOTTOM = &H8
    DT_WORDBREAK = &H10
    DT_SINGLELINE = &H20
    DT_EXPANDTABS = &H40
    DT_tabstop = &H80
    DT_NOCLIP = &H100
    DT_EXTERNALLEADING = &H200
    DT_CALCRECT = &H400
    DT_NOPREFIX = &H800
    DT_INTERNAL = &H1000
    DT_EDITCONTROL = &H2000
    DT_PATH_ELLIPSIS = &H4000
    DT_END_ELLIPSIS = &H8000
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
    DT_NOFULLWIDTHCHARBREAK = &H80000
    DT_HIDEPREFIX = &H100000
    DT_PREFIXONLY = &H200000
End Enum

'Types
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MenuSubItem
    Left As Long
    Top As Long
    Height As Long
    Width As Long
    Msg As String
    SmallIc As IPictureDisp
    State As Long
End Type

Private Type MenuItem
    Left As Long
    Top As Long
    Height As Long
    Width As Long
    MenuType As Long
    Title As String
    BigIc As IPictureDisp
    BackPic As IPictureDisp
    DetailPic As IPictureDisp
    Detail_Title As String
    Detail_Msg As String
    CollapsedOrNot As Boolean
    SubItems() As MenuSubItem
    State As Long
    SubItemsCount As Long
End Type

'Api
Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'##########Subclassing code
'   Made by Paul Caton
'   See http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'   for details
'
'   This was copy-pasted, and somethings are edited.
'   It works.

Private Const WM_MOUSEMOVE           As Long = &H200
Private Const WM_MOUSELEAVE          As Long = &H2A3
Private Const WM_MOVING              As Long = &H216
Private Const WM_SIZING              As Long = &H214
Private Const WM_EXITSIZEMOVE        As Long = &H232

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                             As Long
    dwFlags                            As TRACKMOUSEEVENT_FLAGS
    hwndTrack                          As Long
    dwHoverTime                        As Long
End Type

Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'==================================================================================================
'Subclasser declarations

Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset
Private Const WM_SYSCOLORCHANGE      As Long = &H15
Private Const WM_THEMECHANGED        As Long = &H31A
Private Const WM_CTLCOLORSCROLLBAR   As Long = &H137

Private Type tSubData                                                                   'Subclass data type
    hWnd                               As Long                                            'Handle of the window being subclassed
    nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
    nMsgCntA                           As Long                                            'Msg after table entry count
    nMsgCntB                           As Long                                            'Msg before table entry count
    aMsgTblA()                         As Long                                            'Msg after table array
    aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'==================================================================================================

Private Sub Scoller_Change()
    Call ReDraw
End Sub

Private Sub Scoller_Scroll()
    Call Scoller_Change
End Sub

Private Sub UserControl_Initialize()
    ReDim Items(0)
End Sub

'UserControl events
'Read the properties from the property bag - also, a good place to start the subclassing (if we're running)

'* ========================================================================================================
'*  Subclass handler - MUST be the first Public routine in this file. That includes public properties also
'* ========================================================================================================
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    '* Parameters:
    '*  bBefore  - Indicates whether the the message is _
        being processed before or after the _
        default handler - only really needed _
        if a message is set to callback both _
        before & after.
    '*  bHandled - Set this variable to True in a before _
        callback to prevent the message being _
        subsequently processed by the default _
        handler... and if set, an after _
        callback.
    '*  lReturn  - Set this variable as per your intentions _
        and requirements, see the MSDN _
        documentation for each individual _
        message value.
    '*  hWnd     - The window handle.
    '*  uMsg     - The message number.
    '*  wParam   - Message related data.
    '*  lParam   - Message related data.
    '* Notes: _
        If you really know what youre doing, it's possible _
        to change the values of the hWnd, uMsg, wParam and _
        lParam parameters in a before callback so that _
        different values get passed to the default _
        handler... and optionaly, the after callback.
    Select Case uMsg
        Case WM_MOUSEMOVE
            If Not (bInCtrl = True) Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
            End If
        Case WM_MOUSELEAVE
            bInCtrl = False
            Call UserControl_MouseMove(0, 0, -1, -1)
        Case WM_THEMECHANGED, WM_SYSCOLORCHANGE
            ' When the theme or colors change.
            Call ReDraw
        Case WM_CTLCOLORSCROLLBAR
            bHandled = True
    End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    
    If (Ambient.UserMode = True) Then
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        If Not (bTrackUser32 = True) Then
            If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
                bTrack = False
            End If
        End If
        If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
            '* Start subclassing the UserControl.
            Call Subclass_Start(hWnd)
            Call Subclass_AddMsg(hWnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(hWnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(hWnd, WM_THEMECHANGED, MSG_AFTER)
            Call Subclass_AddMsg(hWnd, WM_SYSCOLORCHANGE, MSG_AFTER)
            Call Subclass_AddMsg(hWnd, WM_CTLCOLORSCROLLBAR, MSG_AFTER)
        End If
    End If
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
    On Error GoTo Catch
    'Stop all subclassing
    Call Subclass_StopAll
Catch:
End Sub

'======================================================================================================
'UserControl private routines
'Determine if the passed function is supported
'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    Dim hMod As Long, bLibLoaded As Boolean
    
    hMod = GetModuleHandleA(sModule)
    If (hMod = 0) Then
        hMod = LoadLibraryA(sModule)
        If (hMod) Then bLibLoaded = True
    End If
    If (hMod) Then
        If (GetProcAddress(hMod, sFunction)) Then IsFunctionExported = True
    End If
    If (bLibLoaded = True) Then Call FreeLibrary(hMod)
End Function

'* Track the mouse leaving the indicated window.
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
    Dim tme As TRACKMOUSEEVENT_STRUCT
    
    If (bTrack = True) Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
        If (bTrackUser32 = True) Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

'* =============================================================================================================================
'*  Subclass code - The programmer may call any of the following Subclass_??? routines
'*  Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages.
'* =============================================================================================================================
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    '* Parameters:
    '*  lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
    '*  uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
    '*  When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If (When) And (eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When) And (eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'* Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    '* Parameters:
    '*  lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table.
    '*  uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback.
    '*  When      - Whether the msg is to be removed from the before, after or both callback tables.
    With sc_aSubData(zIdx(lng_hWnd))
        If (When) And (eMsgWhen.MSG_BEFORE) Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When) And (eMsgWhen.MSG_AFTER) Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'* Return whether were running in the IDE.
Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'* Start subclassing the passed window handle.
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    '* Parameters:
    '*  lng_hWnd - The handle of the window to be subclassed.
    '*  Returns;
    '*  The sc_aSubData() index.
    Const CODE_LEN              As Long = 200
    Const FUNC_CWP              As String = "CallWindowProcA"
    Const FUNC_EBM              As String = "EbMode"
    Const FUNC_SWL              As String = "SetWindowLongA"
    Const MOD_USER              As String = "user32"
    Const MOD_VBA5              As String = "vba5"
    Const MOD_VBA6              As String = "vba6"
    Const PATCH_01              As Long = 18
    Const PATCH_02              As Long = 68
    Const PATCH_03              As Long = 78
    Const PATCH_06              As Long = 116
    Const PATCH_07              As Long = 121
    Const PATCH_0A              As Long = 186
    Static aBuf(1 To CODE_LEN)  As Byte
    Static pCWP                 As Long
    Static pEbMode              As Long
    Static pSWL                 As Long
    Dim i                       As Long
    Dim j                       As Long
    Dim nSubIdx                 As Long
    Dim sHex                    As String
    
    '* If it's the first time through here...
    If (aBuf(1) = 0) Then
        '* The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        '* Convert the string from hex pairs to bytes and store in the static machine code buffer.
        i = 1
        Do While (j < CODE_LEN)
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
        '* Get API function addresses.
        If (Subclass_InIDE = True) Then
            aBuf(16) = &H90
            aBuf(17) = &H90
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (pEbMode = 0) Then pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
        End If
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
        Subclass_Start = nSubIdx
    End If
    With sc_aSubData(nSubIdx)
        .hWnd = lng_hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSub)
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
    End With
End Function

'* Stop all subclassing.
Private Sub Subclass_StopAll()
    Dim i As Long
    
    On Error GoTo myErr
    i = UBound(sc_aSubData())
    Do While (i >= 0)
        With sc_aSubData(i)
            If (.hWnd <> 0) Then Call Subclass_Stop(.hWnd)
        End With
        i = i - 1
    Loop
    Exit Sub
myErr:
End Sub

'* Stop subclassing the passed window handle.
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    '* Parameters:
    '*  lng_hWnd - The handle of the window to stop being subclassed.
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hWnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB
        Erase .aMsgTblA
    End With
End Sub

'* ======================================================================================================
'*  These z??? routines are exclusively called by the Subclass_??? routines.
'*  Worker sub for Subclass_AddMsg.
'* ======================================================================================================
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long, nOff1 As Long, nOff2 As Long
    
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGES
    Else
        Do While (nEntry < nMsgCnt)
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop
        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If
    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
    Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If
    If (uMsg <> ALL_MESSAGES) Then Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

'* Return the memory address of the passed function in the passed dll.
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

'* Worker sub for Subclass_DelMsg.
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = 0
        If (When = eMsgWhen.MSG_BEFORE) Then
            nEntry = PATCH_05
        Else
            nEntry = PATCH_09
        End If
        Call zPatchVal(nAddr, nEntry, 0)
    Else
        Do While (nEntry < nMsgCnt)
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = uMsg) Then
                aMsgTbl(nEntry) = 0
                Exit Do
            End If
        Loop
    End If
End Sub

'* Get the sc_aSubData() array index of the passed hWnd.
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    '* Get the upper bound of sc_aSubData() - If you get an error here, youre probably Subclass_AddMsg-ing before Subclass_Start.
    zIdx = UBound(sc_aSubData)
    Do While (zIdx >= 0)
        With sc_aSubData(zIdx)
            If (.hWnd = lng_hWnd) And Not (bAdd = True) Then
                Exit Function
            ElseIf (.hWnd = 0) And (bAdd = True) Then
                Exit Function
            End If
        End With
        zIdx = zIdx - 1
    Loop
    If Not (bAdd = True) Then Debug.Assert False
    '* If we exit here, were returning -1, no freed elements were found.
End Function

'* Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'* Patch the machine code buffer at the indicated offset with the passed value.
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'* Worker function for Subclass_InIDE.
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function
'##########End

Private Sub UserControl_AmbientChanged(PropertyName As String)
    Call ReDraw
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Integer, GotNewmouse As Boolean, ChangedSomething As Boolean, b As Integer
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    'On error resume next
    For a = 1 To UBound(Items())
        If (X > Items(a).Left) And (X < Items(a).Left + Items(a).Width) Then
            If (Y > Items(a).Top) And (Y < Items(a).Top + Items(a).Height) Then
                If (Button = 0) Then
                    If Not (Items(a).State = STATE_HOT) Then
                        Items(a).State = STATE_HOT
                        ChangedSomething = True
                    End If
                ElseIf (Button = 1) Then
                    If Not (Items(a).State = STATE_PRESSED) Then
                        Items(a).State = STATE_PRESSED
                        ChangedSomething = True
                    End If
                End If
                If Not (UserControl.MousePointer = vbCustom) Then UserControl.MousePointer = vbCustom
                GotNewmouse = True
            ElseIf Not (Items(a).State = STATE_NORMAL) Then
                Items(a).State = STATE_NORMAL
                ChangedSomething = True
            End If
        ElseIf Not (Items(a).State = STATE_NORMAL) Then
            Items(a).State = STATE_NORMAL
            ChangedSomething = True
        End If
        If Not (Items(a).MenuType = MENU_DETAILS) Then
            If (Items(a).CollapsedOrNot = True) Then
                For b = 1 To UBound(Items(a).SubItems())
                    If (X > Items(a).SubItems(b).Left And X < Items(a).SubItems(b).Left + Items(a).SubItems(b).Width) Then
                        If (Y > Items(a).SubItems(b).Top And Y < Items(a).SubItems(b).Top + Items(a).SubItems(b).Height) Then
                            If Not (Items(a).SubItems(b).State = STATE_HOT) Then
                                Items(a).SubItems(b).State = STATE_HOT
                                ChangedSomething = True
                            End If
                            If Not (UserControl.MousePointer = vbCustom) Then UserControl.MousePointer = vbCustom
                            GotNewmouse = True
                        ElseIf Not (Items(a).SubItems(b).State = STATE_NORMAL) Then
                            Items(a).SubItems(b).State = STATE_NORMAL
                            ChangedSomething = True
                        End If
                    ElseIf Not (Items(a).SubItems(b).State = STATE_NORMAL) Then
                        Items(a).SubItems(b).State = STATE_NORMAL
                        ChangedSomething = True
                    End If
                Next
            End If
        End If
    Next
    If (GotNewmouse = False) Then UserControl.MousePointer = vbDefault
    If (ChangedSomething = True) Then Call ReDraw
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim a As Integer, GotNewmouse As Boolean, b As Integer
    
    'On error resume next
    RaiseEvent MouseUp(Button, Shift, X, Y)
    For a = 1 To UBound(Items())
        Items(a).State = STATE_NORMAL
        If (X > Items(a).Left And X < Items(a).Left + Items(a).Width) Then
            If (Y > Items(a).Top And Y < Items(a).Top + Items(a).Height) Then
                If (Button = 1) Then
                    Items(a).CollapsedOrNot = Not Items(a).CollapsedOrNot
                    If (Items(a).CollapsedOrNot = False) Then
                        RaiseEvent Collapse(a)
                    Else
                        RaiseEvent Expand(a)
                    End If
                    RaiseEvent HeaderClick(a)
                    Call UserControl_MouseMove(0, 0, X, Y)
                End If
                Exit For
            End If
        End If
        If Not (Items(a).MenuType = MENU_DETAILS) Then
            If (Items(a).CollapsedOrNot = True) Then
                For b = 1 To UBound(Items(a).SubItems())
                    If (X > Items(a).SubItems(b).Left And X < Items(a).SubItems(b).Left + Items(a).SubItems(b).Width) Then
                        If (Y > Items(a).SubItems(b).Top And Y < Items(a).SubItems(b).Top + Items(a).SubItems(b).Height) Then
                            If (Button = 1) Then
                                RaiseEvent SubItemClick(a, b)
                                Call UserControl_MouseMove(0, 0, X, Y)
                            End If
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
    Next
    Call ReDraw
End Sub

Private Sub Usercontrol_Paint()
    Call ReDraw
End Sub

Public Sub ReDraw()
    Dim MaxY As Long
    
    If Not (UserControl.Ambient.UserMode) Then
        Call DrawBackGround
        UserControl.CurrentX = 120
        UserControl.CurrentY = 120
        UserControl.ForeColor = 0
        UserControl.FontBold = True
        UserControl.FontUnderline = False
        UserControl.Print (("ExplorerBar 2 by Flex 2005"))
        UserControl.CurrentY = 300
        UserControl.CurrentX = 120
        UserControl.Print (("flex4d@gmail.com"))
        UserControl.CurrentY = 600
        UserControl.ForeColor = 0
        UserControl.FontBold = True
        UserControl.FontUnderline = False
        UserControl.Print (("<-------------->Normal Size<---------------->"))
        Exit Sub
    End If
    Call UserControl.Cls
    'This draws the bar
    Call DrawBackGround
    Call DrawItems(MaxY)
    If (MaxY > UserControl.ScaleHeight) Then
        If (Scoller.Visible = False) Then
            Scoller.Left = UserControl.ScaleWidth - Scoller.Width
            Scoller.Top = 0
            Scoller.Height = UserControl.ScaleHeight
            Scoller.Max = (MaxY - UserControl.ScaleHeight) / 10
            Scoller.Visible = True
            Call DrawBackGround
            Call DrawItems(MaxY)
        ElseIf (UserControl.ScaleHeight <> Scoller.Height) Then
            Scoller.Height = UserControl.ScaleHeight
            Scoller.Max = (MaxY - UserControl.ScaleHeight) / 10
        ElseIf (MaxY - UserControl.ScaleHeight) / 10 <> Scoller.Max Then
            Scoller.Height = UserControl.ScaleHeight
            Scoller.Max = (MaxY - UserControl.ScaleHeight) / 10
        End If
    ElseIf (Scoller.Visible = True) Then
        Scoller.Visible = False
        Scoller.Left = UserControl.ScaleWidth
        Call DrawBackGround
        Call DrawItems(MaxY)
    End If
    DoEvents
End Sub

Public Sub Cls()
    ItemCount = 0
    ReDim Items(0)
    Call ReDraw
End Sub

Private Function DrawPart(strClass As String, ByVal iPart As Long, ByVal iState As Long, pRect As RECT, pClipRect As RECT) As Boolean
    Dim hTheme As Long, lResult As Long
    
    On Error GoTo ErrorHandler
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr(strClass))
    If (hTheme) Then
        lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, pRect, pClipRect)
        DrawPart = IIf(lResult, False, True)
        Call CloseThemeData(hTheme)
    Else
        DrawPart = False
    End If
    Exit Function
ErrorHandler:
    DrawPart = False
End Function

Private Sub DrawBackGround()
    Dim bResult As Boolean, pRect As RECT
    
    'This draws the Background
    pRect.Left = 0
    pRect.Top = 0
    pRect.Right = UserControl.ScaleWidth / Screen.TwipsPerPixelX
    pRect.Bottom = UserControl.ScaleHeight / Screen.TwipsPerPixelY
    bResult = DrawPart("ExplorerBar", EBP_HEADERBACKGROUND, STATE_NORMAL, pRect, pRect)
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        UserControl.Line (0, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight), vbWhite, BF
    End If
End Sub

Private Sub DrawItems(ret_MaxY As Long)
    Dim CurrentY As Long, a As Integer
    
    If (Scoller.Visible = False) Then
        CurrentY = 240
    Else
        CurrentY = 240 - (Scoller.Value * 10)
    End If
    'On Error Resume Next
    For a = 1 To UBound(Items())
        If (Items(a).MenuType = MENU_NORMAL) Then
            Call DrawNormalHeader(CurrentY, a)
            If (Items(a).CollapsedOrNot = True) Then Call DrawNormalSubItems(CurrentY, a)
        End If
        If (Items(a).MenuType = MENU_SPECIAL) Then
            Call DrawSpecialHeader(CurrentY, a)
            If (Items(a).CollapsedOrNot = True) Then Call DrawSpecialSubItems(CurrentY, a)
        ElseIf (Items(a).MenuType = MENU_DETAILS) Then
            Call DrawNormalHeader(CurrentY, a)
            If (Items(a).CollapsedOrNot = True) Then Call DrawDetails(CurrentY, a)
        End If
        CurrentY = CurrentY + 240
    Next
    If (Scoller.Visible = False) Then
        ret_MaxY = CurrentY
    Else
        ret_MaxY = CurrentY + (Scoller.Value * 10)
    End If
End Sub

Public Sub About()
    Call MsgBox("ExplorerBar - (c) Copyright Flex 2005 - flex4d@gmail.com", vbExclamation, "About")
End Sub

Private Sub DrawNormalHeader(CurY As Long, ItemBound As Integer)
    Dim bResult As Boolean, pRect As RECT, pClipRect As RECT
    
    pRect.Left = 120 / Screen.TwipsPerPixelX
    pRect.Top = CurY / Screen.TwipsPerPixelY
    pRect.Right = (Scoller.Left - 120) / Screen.TwipsPerPixelX
    pRect.Bottom = (375 + CurY) / Screen.TwipsPerPixelY
    Items(ItemBound).Top = CurY
    Items(ItemBound).Left = 120
    Items(ItemBound).Width = Scoller.Left - 240
    Items(ItemBound).Height = 375
    bResult = DrawPart("ExplorerBar", EBP_NORMALGROUPHEAD, STATE_NORMAL, pRect, pRect)
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        UserControl.Line (120, CurY)-(Scoller.Left - 120, CurY + 375), vbButtonFace, BF
        Call DrawNormalHeaderTitle(CurY, ItemBound, False)
        CurY = CurY + 375
    Else
        Call DrawNormalHeaderTitle(CurY, ItemBound, True)
        CurY = CurY + 375
    End If
End Sub

Private Sub DrawSpecialHeader(CurY As Long, ItemBound As Integer)
    Dim bResult As Boolean, pRect As RECT, pClipRect As RECT
    
    pRect.Left = 120 / Screen.TwipsPerPixelX
    pRect.Top = CurY / Screen.TwipsPerPixelY
    pRect.Right = (Scoller.Left - 120) / Screen.TwipsPerPixelX
    pRect.Bottom = (375 + CurY) / Screen.TwipsPerPixelY
    Items(ItemBound).Top = CurY
    Items(ItemBound).Left = 120
    Items(ItemBound).Width = Scoller.Left - 240
    Items(ItemBound).Height = 375
    bResult = DrawPart("ExplorerBar", EBP_SPECIALGROUPHEAD, STATE_NORMAL, pRect, pRect)
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        UserControl.Line (120, CurY)-(Scoller.Left - 120, CurY + 375), 6956042, BF
        Call DrawSpecialHeaderTitle(CurY, ItemBound, False)
    Else
        Call DrawSpecialHeaderTitle(CurY, ItemBound, True)
    End If
    CurY = CurY + 375
End Sub

Private Sub DrawNormalSubItems(CurY As Long, ItemBound As Integer)
    Dim CalcY   As Long, BaseY    As Long, a        As Integer
    Dim bResult As Boolean, pRect As RECT, pClipRect As RECT
    Dim txtRect As RECT
    'On error resume next
    BaseY = CurY
    CalcY = 100
    For a = 1 To UBound(Items(ItemBound).SubItems())
        CalcY = CalcY + 300
    Next
    pRect.Left = 120 / Screen.TwipsPerPixelX
    pRect.Top = (Items(ItemBound).Top + Items(ItemBound).Height) / Screen.TwipsPerPixelY
    pRect.Right = (Scoller.Left - 120) / Screen.TwipsPerPixelX
    pRect.Bottom = (CalcY + CurY) / Screen.TwipsPerPixelY
    bResult = DrawPart("ExplorerBar", EBP_NORMALGROUPBACKGROUND, STATE_NORMAL, pRect, pRect)
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        UserControl.Line (120, (Items(ItemBound).Top + Items(ItemBound).Height))-(Scoller.Left - 120, (CalcY + CurY)), vbButtonFace, B
    ElseIf Not (Items(ItemBound).BackPic Is Nothing) Then
        Call UserControl.PaintPicture(Items(ItemBound).BackPic, 120, (Items(ItemBound).Top + Items(ItemBound).Height), Scoller.Left - 240, (CalcY + CurY) - (Items(ItemBound).Top + Items(ItemBound).Height))
    End If
    CurY = CurY + CalcY
    BaseY = BaseY + 100
    For a = 1 To UBound(Items(ItemBound).SubItems())
        If (bResult = True) Then
            If Not (Items(ItemBound).SubItems(a).SmallIc Is Nothing) Then Call UserControl.PaintPicture(Items(ItemBound).SubItems(a).SmallIc, 240, BaseY, 240, 240)
            Select Case Items(ItemBound).SubItems(a).State
                Case STATE_NORMAL
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 12999969
                    UserControl.FontBold = False
                    UserControl.FontUnderline = False
                Case STATE_HOT
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 16748098
                    UserControl.FontBold = False
                    UserControl.FontUnderline = True
            End Select
            
            txtRect.Top = BaseY / Screen.TwipsPerPixelY
            txtRect.Bottom = (BaseY + UserControl.TextHeight((Items(ItemBound).SubItems(a).Msg))) / Screen.TwipsPerPixelY
            txtRect.Left = 600 / Screen.TwipsPerPixelX
            txtRect.Right = (Scoller.Left - 240) / Screen.TwipsPerPixelX
            
            Items(ItemBound).SubItems(a).Top = BaseY
            Items(ItemBound).SubItems(a).Left = 240
            Items(ItemBound).SubItems(a).Height = 240
            Items(ItemBound).SubItems(a).Width = Scoller.Left - 480
            Call DrawText(UserControl.hDC, ((Items(ItemBound).SubItems(a).Msg)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
            'UserControl.Print ((Items(ItemBound).SubItems(a).Msg))
        Else
            Call UserControl.PaintPicture(Items(ItemBound).SubItems(a).SmallIc, 240, BaseY, 240, 240)
            Select Case Items(ItemBound).SubItems(a).State
                Case STATE_NORMAL
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 0
                    UserControl.FontBold = False
                    UserControl.FontUnderline = False
                Case STATE_HOT
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 0
                    UserControl.FontBold = False
                    UserControl.FontUnderline = True
            End Select
            
            
            txtRect.Top = BaseY / Screen.TwipsPerPixelY
            txtRect.Bottom = (BaseY + UserControl.TextHeight((Items(ItemBound).SubItems(a).Msg))) / Screen.TwipsPerPixelY
            txtRect.Left = 600 / Screen.TwipsPerPixelX
            txtRect.Right = (Scoller.Left - 240) / Screen.TwipsPerPixelX
            Items(ItemBound).SubItems(a).Top = BaseY
            Items(ItemBound).SubItems(a).Left = 240
            Items(ItemBound).SubItems(a).Height = 240
            Items(ItemBound).SubItems(a).Width = UserControl.TextWidth(Items(ItemBound).SubItems(a).Msg) + 360
            Call DrawText(UserControl.hDC, ((Items(ItemBound).SubItems(a).Msg)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
            
            
            'UserControl.Print ((Items(ItemBound).SubItems(a).Msg))
        End If
        BaseY = BaseY + 300
    Next
End Sub

Private Sub DrawDetails(CurY As Long, ItemBound As Integer)
    Dim CalcY   As Long, BaseY    As Long, Msg       As String
    Dim bResult As Boolean, pRect As RECT, pClipRect As RECT
    Dim txtRect As RECT
    'On error resume next
    BaseY = CurY
    CalcY = 100
    UserControl.FontName = "Tahoma"
    UserControl.ForeColor = 0
    UserControl.FontBold = False
    UserControl.FontUnderline = False
    Msg = MakeMSG(Items(ItemBound).Detail_Msg, Scoller.Left - 480)
    CalcY = CalcY + UserControl.TextHeight(Msg) + 100 + 240
    'If (UserControl.ScaleWidth - 960 - 960) > 0 And Not Items(ItemBound).DetailPic Is Nothing Then
    '    CalcY = CalcY + (UserControl.ScaleWidth - 960 - 960) + 120
    'End If
    '^^^I deleted this because of the scrollbar
    If Not (Items(ItemBound).DetailPic Is Nothing) Then
        Dim picSize As Long
        
        picSize = 1200
        If (picSize > 0) Then CalcY = CalcY + picSize + 120
    End If
    pRect.Left = 120 / Screen.TwipsPerPixelX
    pRect.Top = (Items(ItemBound).Top + Items(ItemBound).Height) / Screen.TwipsPerPixelY
    pRect.Right = (Scoller.Left - 120) / Screen.TwipsPerPixelX
    pRect.Bottom = (CalcY + CurY) / Screen.TwipsPerPixelY
    bResult = DrawPart("ExplorerBar", EBP_NORMALGROUPBACKGROUND, STATE_NORMAL, pRect, pRect)
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        UserControl.Line (120, (Items(ItemBound).Top + Items(ItemBound).Height))-(Scoller.Left - 120, (CalcY + CurY)), vbButtonFace, B
    ElseIf Not (Items(ItemBound).BackPic Is Nothing) Then
        Call UserControl.PaintPicture(Items(ItemBound).BackPic, 120, (Items(ItemBound).Top + Items(ItemBound).Height), Scoller.Left - 240, (CalcY + CurY) - (Items(ItemBound).Top + Items(ItemBound).Height))
    End If
    CurY = CurY + CalcY
    BaseY = BaseY + 100
    If Not (Items(ItemBound).DetailPic Is Nothing) Then
        If (picSize > 0) Then
            On Error Resume Next
            Call UserControl.PaintPicture(Items(ItemBound).DetailPic, Scoller.Left / 2 - picSize / 2, BaseY, picSize, picSize, , , , , vbSrcAnd)
            BaseY = BaseY + picSize + 120
            On Error GoTo 0
        End If
    End If
    pRect.Left = 240 / Screen.TwipsPerPixelX
    pRect.Top = (BaseY + 240) / Screen.TwipsPerPixelY
    pRect.Right = (Scoller.Left) / Screen.TwipsPerPixelX
    pRect.Bottom = (CalcY + BaseY) / Screen.TwipsPerPixelY
    UserControl.FontName = "Tahoma"
    UserControl.ForeColor = 0
    UserControl.FontBold = True
    UserControl.FontUnderline = False
    
    
    
    txtRect.Top = BaseY / Screen.TwipsPerPixelY
    txtRect.Bottom = (BaseY + UserControl.TextHeight((Items(ItemBound).Detail_Title))) / Screen.TwipsPerPixelY
    txtRect.Left = 240 / Screen.TwipsPerPixelX
    txtRect.Right = (Scoller.Left - 240) / Screen.TwipsPerPixelX
    
    'UserControl.Print ((Items(ItemBound).Detail_Title))
    
    Call DrawText(UserControl.hDC, ((Items(ItemBound).Detail_Title)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
    
    
    
    UserControl.FontName = "Tahoma"
    UserControl.ForeColor = 0
    UserControl.FontBold = False
    UserControl.FontUnderline = False
    Call DrawText(UserControl.hDC, Msg, -1, pRect, DT_LEFT)
End Sub

Private Sub DrawSpecialSubItems(CurY As Long, ItemBound As Integer)
    Dim CalcY   As Long, BaseY    As Long, a         As Integer
    Dim bResult As Boolean, pRect As RECT, pClipRect As RECT
    Dim txtRect As RECT
    'On error resume next
    BaseY = CurY
    CalcY = 100
    For a = 1 To UBound(Items(ItemBound).SubItems())
        CalcY = CalcY + 300
    Next
    pRect.Left = 120 / Screen.TwipsPerPixelX
    pRect.Top = (Items(ItemBound).Top + Items(ItemBound).Height) / Screen.TwipsPerPixelY
    pRect.Right = (Scoller.Left - 120) / Screen.TwipsPerPixelX
    pRect.Bottom = (CalcY + CurY) / Screen.TwipsPerPixelY
    bResult = DrawPart("ExplorerBar", EBP_SPECIALGROUPBACKGROUND, STATE_NORMAL, pRect, pRect)
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        UserControl.Line (120, (Items(ItemBound).Top + Items(ItemBound).Height))-(Scoller.Left - 120, (CalcY + CurY)), 6956042, B
    ElseIf Not (Items(ItemBound).BackPic Is Nothing) Then
        Call UserControl.PaintPicture(Items(ItemBound).BackPic, 120, (Items(ItemBound).Top + Items(ItemBound).Height), Scoller.Left - 240, (CalcY + CurY) - (Items(ItemBound).Top + Items(ItemBound).Height))
    End If
    CurY = CurY + CalcY
    BaseY = BaseY + 100
    For a = 1 To UBound(Items(ItemBound).SubItems())
        If (bResult = True) Then
            If Not (Items(ItemBound).SubItems(a).SmallIc Is Nothing) Then Call UserControl.PaintPicture(Items(ItemBound).SubItems(a).SmallIc, 240, BaseY, 240, 240)
            Select Case Items(ItemBound).SubItems(a).State
                Case STATE_NORMAL
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 12999969
                    UserControl.FontBold = False
                    UserControl.FontUnderline = False
                Case STATE_HOT
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 16748098
                    UserControl.FontBold = False
                    UserControl.FontUnderline = True
            End Select
            
            txtRect.Top = BaseY / Screen.TwipsPerPixelY
            txtRect.Bottom = (BaseY + UserControl.TextHeight((Items(ItemBound).SubItems(a).Msg))) / Screen.TwipsPerPixelY
            txtRect.Left = 600 / Screen.TwipsPerPixelX
            txtRect.Right = (Scoller.Left - 240) / Screen.TwipsPerPixelX
            Items(ItemBound).SubItems(a).Top = BaseY
            Items(ItemBound).SubItems(a).Left = 240
            Items(ItemBound).SubItems(a).Height = 240
            Items(ItemBound).SubItems(a).Width = Scoller.Left - 480
            'UserControl.Print ((Items(ItemBound).SubItems(a).Msg))
            Call DrawText(UserControl.hDC, ((Items(ItemBound).SubItems(a).Msg)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
        Else
            Call UserControl.PaintPicture(Items(ItemBound).SubItems(a).SmallIc, 240, BaseY, 240, 240)
            Select Case Items(ItemBound).SubItems(a).State
                Case STATE_NORMAL
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 0
                    UserControl.FontBold = False
                    UserControl.FontUnderline = False
                Case STATE_HOT
                    UserControl.FontName = "Tahoma"
                    UserControl.ForeColor = 0
                    UserControl.FontBold = False
                    UserControl.FontUnderline = True
            End Select
            
            txtRect.Top = BaseY / Screen.TwipsPerPixelY
            txtRect.Bottom = (BaseY + UserControl.TextHeight((Items(ItemBound).SubItems(a).Msg))) / Screen.TwipsPerPixelY
            txtRect.Left = 600 / Screen.TwipsPerPixelX
            txtRect.Right = (Scoller.Left - 240) / Screen.TwipsPerPixelX
            Items(ItemBound).SubItems(a).Top = BaseY
            Items(ItemBound).SubItems(a).Left = 240
            Items(ItemBound).SubItems(a).Height = 240
            Items(ItemBound).SubItems(a).Width = UserControl.TextWidth(Items(ItemBound).SubItems(a).Msg) + 360
            
            'UserControl.Print ((Items(ItemBound).SubItems(a).Msg))
            Call DrawText(UserControl.hDC, ((Items(ItemBound).SubItems(a).Msg)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
            
        End If
        BaseY = BaseY + 300
    Next
End Sub

Public Function AddNormalItem(Title As String, Optional CollapsedOrNot As Boolean = True, Optional BigIcon As IPictureDisp, Optional BackGround As IPictureDisp) As Integer
    ItemCount = ItemCount + 1
    ReDim Preserve Items(ItemCount)
    ReDim Items(ItemCount).SubItems(0)
    Items(ItemCount).Title = Title
    If Not (BigIcon Is Nothing) Then Set Items(ItemCount).BigIc = BigIcon
    If Not (BackGround Is Nothing) Then Set Items(ItemCount).BackPic = BackGround
    Items(ItemCount).CollapsedOrNot = CollapsedOrNot
    Items(ItemCount).MenuType = MENU_NORMAL
    Items(ItemCount).State = STATE_NORMAL
    AddNormalItem = ItemCount
    Call ReDraw
End Function

Public Function AddSpecialItem(Title As String, Optional CollapsedOrNot As Boolean = True, Optional BigIcon As IPictureDisp, Optional BackGround As IPictureDisp) As Integer
    ItemCount = ItemCount + 1
    ReDim Preserve Items(ItemCount)
    Items(ItemCount).Title = Title
    If Not (BigIcon Is Nothing) Then Set Items(ItemCount).BigIc = BigIcon
    If Not (BackGround Is Nothing) Then Set Items(ItemCount).BackPic = BackGround
    Items(ItemCount).CollapsedOrNot = CollapsedOrNot
    Items(ItemCount).MenuType = MENU_SPECIAL
    Items(ItemCount).State = STATE_NORMAL
    ReDim Items(ItemCount).SubItems(0)
    AddSpecialItem = ItemCount
    Call ReDraw
End Function

Public Function AddDetailItem(Title As String, DetailTitle As String, Details As String, Optional DetailPicture As IPictureDisp, Optional CollapsedOrNot As Boolean = True, Optional BigIcon As IPictureDisp, Optional BackGround As IPictureDisp) As Integer
    ItemCount = ItemCount + 1
    ReDim Preserve Items(ItemCount)
    Items(ItemCount).Title = Title
    Items(ItemCount).Detail_Title = DetailTitle
    Items(ItemCount).Detail_Msg = Details
    If Not (BigIcon Is Nothing) Then
        Set Items(ItemCount).BigIc = BigIcon
    End If
    If Not (BackGround Is Nothing) Then
        Set Items(ItemCount).BackPic = BackGround
    End If
    If Not (DetailPicture Is Nothing) Then
        Set Items(ItemCount).DetailPic = DetailPicture
    End If
    Items(ItemCount).CollapsedOrNot = CollapsedOrNot
    Items(ItemCount).MenuType = MENU_DETAILS
    Items(ItemCount).State = STATE_NORMAL
    ReDim Items(ItemCount).SubItems(0)
    AddDetailItem = ItemCount
    Call ReDraw
End Function

Public Function AddSubItem(Index As Integer, Caption As String, Optional SmallIcon As IPictureDisp) As Integer
    'On error resume next
    If (Items(Index).MenuType = MENU_NORMAL Or Items(Index).MenuType = MENU_SPECIAL) Then
        Items(Index).SubItemsCount = Items(Index).SubItemsCount + 1
        ReDim Preserve Items(Index).SubItems(Items(Index).SubItemsCount)
        Items(Index).SubItems(Items(Index).SubItemsCount).Msg = Caption
        If Not (SmallIcon Is Nothing) Then Set Items(Index).SubItems(Items(Index).SubItemsCount).SmallIc = SmallIcon
        Items(Index).SubItems(Items(Index).SubItemsCount).State = STATE_NORMAL
        AddSubItem = Items(Index).SubItemsCount
        Call ReDraw
    End If
End Function

Private Sub UserControl_Resize()
    If (Scoller.Visible = False) Then
        Scoller.Left = UserControl.ScaleWidth
    Else
        Scoller.Left = UserControl.ScaleWidth - Scoller.Width
    End If
    Call ReDraw
End Sub

Private Sub UserControl_Show()
    Call ReDraw
End Sub

Private Sub DrawNormalHeaderTitle(CurY As Long, ItemBound As Integer, HasTheme As Boolean)
    Dim txtRect As RECT
    If (HasTheme = True) Then
        Select Case Items(ItemBound).State
            Case STATE_NORMAL
                UserControl.FontName = "Tahoma"
                UserControl.ForeColor = 12999969
                UserControl.FontBold = True
                UserControl.FontUnderline = False
            Case STATE_HOT
                UserControl.FontName = "Tahoma"
                UserControl.ForeColor = 16748098
                UserControl.FontBold = True
                UserControl.FontUnderline = False
            Case STATE_PRESSED
                UserControl.FontName = "Tahoma"
                UserControl.ForeColor = 16748098
                UserControl.FontBold = True
                UserControl.FontUnderline = False
        End Select
        UserControl.CurrentY = CurY + 90
        If (Items(ItemBound).BigIc Is Nothing) Then
            UserControl.CurrentX = 120 + 180
        Else
            UserControl.CurrentX = 120 + 480
        End If
        
        txtRect.Top = UserControl.CurrentY / Screen.TwipsPerPixelY
        txtRect.Bottom = (UserControl.CurrentY + UserControl.TextHeight((Items(ItemBound).Title))) / Screen.TwipsPerPixelY
        txtRect.Left = UserControl.CurrentX / Screen.TwipsPerPixelX
        txtRect.Right = (Scoller.Left - 480) / Screen.TwipsPerPixelX
        
        'UserControl.Print ((Items(ItemBound).Title))
        
        
        Call DrawText(UserControl.hDC, ((Items(ItemBound).Title)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
        If Not (Items(ItemBound).BigIc Is Nothing) Then
            Call UserControl.PaintPicture(Items(ItemBound).BigIc, 120, CurY - 100, 480, 480)
        End If
        Call DrawNormalButton(CurY, ItemBound)
    Else
        UserControl.FontName = "Tahoma"
        UserControl.ForeColor = 0
        UserControl.FontBold = True
        UserControl.FontUnderline = False
        UserControl.CurrentY = CurY + 90
        If (Items(ItemBound).BigIc Is Nothing) Then
            UserControl.CurrentX = 120 + 180
        Else
            UserControl.CurrentX = 120 + 480
        End If
        txtRect.Top = UserControl.CurrentY / Screen.TwipsPerPixelY
        txtRect.Bottom = (UserControl.CurrentY + UserControl.TextHeight((Items(ItemBound).Title))) / Screen.TwipsPerPixelY
        txtRect.Left = UserControl.CurrentX / Screen.TwipsPerPixelX
        txtRect.Right = (Scoller.Left - 480) / Screen.TwipsPerPixelX
        
        'UserControl.Print ((Items(ItemBound).Title))
        
        
        Call DrawText(UserControl.hDC, ((Items(ItemBound).Title)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
        If Not (Items(ItemBound).BigIc Is Nothing) Then
            Call UserControl.PaintPicture(Items(ItemBound).BigIc, 120, CurY - 100, 480, 480)
        End If
        Call DrawNormalButton(CurY, ItemBound)
    End If
End Sub

Private Sub DrawSpecialHeaderTitle(CurY As Long, ItemBound As Integer, HasTheme As Boolean)
    Dim txtRect As RECT
    
    If (HasTheme = True) Then
        Select Case Items(ItemBound).State
            Case STATE_NORMAL
                UserControl.FontName = "Tahoma"
                UserControl.ForeColor = 16777215
                UserControl.FontBold = True
                UserControl.FontUnderline = False
            Case STATE_HOT
                UserControl.FontName = "Tahoma"
                UserControl.ForeColor = 16748098
                UserControl.FontBold = True
                UserControl.FontUnderline = False
            Case STATE_PRESSED
                UserControl.FontName = "Tahoma"
                UserControl.ForeColor = 16748098
                UserControl.FontBold = True
                UserControl.FontUnderline = False
        End Select
        UserControl.CurrentY = CurY + 90
        If (Items(ItemBound).BigIc Is Nothing) Then
            UserControl.CurrentX = 120 + 180
        Else
            UserControl.CurrentX = 120 + 480
        End If
        txtRect.Top = UserControl.CurrentY / Screen.TwipsPerPixelY
        txtRect.Bottom = (UserControl.CurrentY + UserControl.TextHeight((Items(ItemBound).Title))) / Screen.TwipsPerPixelY
        txtRect.Left = UserControl.CurrentX / Screen.TwipsPerPixelX
        txtRect.Right = (Scoller.Left - 480) / Screen.TwipsPerPixelX
        
        'UserControl.Print ((Items(ItemBound).Title))
        
        
        Call DrawText(UserControl.hDC, ((Items(ItemBound).Title)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
        If Not (Items(ItemBound).BigIc Is Nothing) Then
            Call UserControl.PaintPicture(Items(ItemBound).BigIc, 120, CurY - 100, 480, 480)
        End If
        Call DrawSpecialButton(CurY, ItemBound)
    Else
        UserControl.FontName = "Tahoma"
        UserControl.ForeColor = vbWhite
        UserControl.FontBold = True
        UserControl.FontUnderline = False
        UserControl.CurrentY = CurY + 90
        If (Items(ItemBound).BigIc Is Nothing) Then
            UserControl.CurrentX = 120 + 180
        Else
            UserControl.CurrentX = 120 + 480
        End If
        txtRect.Top = UserControl.CurrentY / Screen.TwipsPerPixelY
        txtRect.Bottom = (UserControl.CurrentY + UserControl.TextHeight((Items(ItemBound).Title))) / Screen.TwipsPerPixelY
        txtRect.Left = UserControl.CurrentX / Screen.TwipsPerPixelX
        txtRect.Right = (Scoller.Left - 480) / Screen.TwipsPerPixelX
        
        'UserControl.Print ((Items(ItemBound).Title))
        
        
        Call DrawText(UserControl.hDC, ((Items(ItemBound).Title)), -1, txtRect, DT_LEFT Or DT_WORD_ELLIPSIS)
        If Not (Items(ItemBound).BigIc Is Nothing) Then
            Call UserControl.PaintPicture(Items(ItemBound).BigIc, 120, CurY - 100, 480, 480)
        End If
        Call DrawSpecialButton(CurY, ItemBound)
    End If
End Sub

Private Sub DrawArrow(X As Long, Y As Long, DownorUp As Boolean, Color As Long)
    If (DownorUp = False) Then
        UserControl.Line (X + 135, Y + 165)-(X + 195, Y + 105), Color
        UserControl.Line (X + 150, Y + 165)-(X + 195, Y + 120), Color
        UserControl.Line (X + 225, Y + 165)-(X + 165, Y + 105), Color
        UserControl.Line (X + 210, Y + 165)-(X + 165, Y + 120), Color
        UserControl.Line (X + 135, Y + 225)-(X + 195, Y + 165), Color
        UserControl.Line (X + 150, Y + 225)-(X + 195, Y + 180), Color
        UserControl.Line (X + 225, Y + 225)-(X + 165, Y + 165), Color
        UserControl.Line (X + 210, Y + 225)-(X + 165, Y + 180), Color
    Else
        UserControl.Line (X + 135, Y + 120)-(X + 195, Y + 180), Color
        UserControl.Line (X + 150, Y + 120)-(X + 195, Y + 165), Color
        UserControl.Line (X + 225, Y + 120)-(X + 165, Y + 180), Color
        UserControl.Line (X + 210, Y + 120)-(X + 165, Y + 165), Color
        UserControl.Line (X + 135, Y + 180)-(X + 195, Y + 240), Color
        UserControl.Line (X + 150, Y + 180)-(X + 195, Y + 225), Color
        UserControl.Line (X + 225, Y + 180)-(X + 165, Y + 240), Color
        UserControl.Line (X + 210, Y + 180)-(X + 165, Y + 225), Color
    End If
End Sub

Private Sub DrawNormalButton(CurY As Long, ItemBound As Integer)
    Dim bResult As Boolean, pRect As RECT, pClipRect As RECT
    
    pRect.Left = (Scoller.Left - 120 - 360) / Screen.TwipsPerPixelX
    pRect.Top = (CurY + 20) / Screen.TwipsPerPixelY
    pRect.Right = pRect.Left + 24
    pRect.Bottom = pRect.Top + 24
    If (Items(ItemBound).CollapsedOrNot = True) Then
        bResult = DrawPart("ExplorerBar", EBP_NORMALGROUPCOLLAPSE, Items(ItemBound).State, pRect, pRect)
    Else
        bResult = DrawPart("ExplorerBar", EBP_NORMALGROUPEXPAND, Items(ItemBound).State, pRect, pRect)
    End If
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        If (Items(ItemBound).State > 1) Then
            Dim BoxLeft As Integer, BoxTop As Integer
            
            BoxTop = CurY + 80
            BoxLeft = Scoller.Left - 120 - 360
            UserControl.Line (BoxLeft, BoxTop)-(BoxLeft + 240, BoxTop), vb3DHighlight
            UserControl.Line (BoxLeft, BoxTop)-(BoxLeft, BoxTop + 240), vb3DHighlight
            UserControl.Line (BoxLeft + 240, BoxTop + 240)-(BoxLeft, BoxTop + 240), vb3DShadow
            UserControl.Line (BoxLeft + 240, BoxTop + 240)-(BoxLeft + 240, BoxTop), vb3DShadow
        End If
        Call DrawArrow(Scoller.Left - 120 - 420, CurY + 20, IIf(Items(ItemBound).CollapsedOrNot, False, True), 0)
    End If
End Sub

Private Sub DrawSpecialButton(CurY As Long, ItemBound As Integer)
    Dim bResult As Boolean, pRect As RECT, pClipRect As RECT
    
    pRect.Left = (Scoller.Left - 120 - 360) / Screen.TwipsPerPixelX
    pRect.Top = (CurY + 20) / Screen.TwipsPerPixelY
    pRect.Right = pRect.Left + 24
    pRect.Bottom = pRect.Top + 24
    If (Items(ItemBound).CollapsedOrNot = True) Then
        bResult = DrawPart("ExplorerBar", EBP_SPECIALGROUPCOLLAPSE, Items(ItemBound).State, pRect, pRect)
    Else
        bResult = DrawPart("ExplorerBar", EBP_SPECIALGROUPEXPAND, Items(ItemBound).State, pRect, pRect)
    End If
    If (bResult = False) Then
        'Draw failed, no Winxp found, draw clasic one
        If (Items(ItemBound).State > 1) Then
            Dim BoxLeft As Integer
            Dim BoxTop As Integer
            BoxTop = CurY + 80
            BoxLeft = Scoller.Left - 120 - 360
            UserControl.Line (BoxLeft, BoxTop)-(BoxLeft + 240, BoxTop), vb3DHighlight
            UserControl.Line (BoxLeft, BoxTop)-(BoxLeft, BoxTop + 240), vb3DHighlight
            UserControl.Line (BoxLeft + 240, BoxTop + 240)-(BoxLeft, BoxTop + 240), vb3DShadow
            UserControl.Line (BoxLeft + 240, BoxTop + 240)-(BoxLeft + 240, BoxTop), vb3DShadow
        End If
        Call DrawArrow(Scoller.Left - 120 - 420, CurY + 20, IIf(Items(ItemBound).CollapsedOrNot, False, True), vbWhite)
    End If
End Sub

Public Function Header(Index As Integer) As String
    'On error resume next
    Header = Items(Index).Title
End Function

Public Function HeaderIcon(Index As Integer) As IPictureDisp
    'On error resume next
    Set HeaderIcon = Items(Index).BigIc
End Function

Public Function BackGround(Index As Integer) As IPictureDisp
    'On error resume next
    Set BackGround = Items(Index).BackPic
End Function

Public Function SubItemIcon(Index As Integer, SubItemIndex As Integer) As IPictureDisp
    'On error resume next
    If (Items(Index).MenuType = MENU_NORMAL Or Items(Index).MenuType = MENU_SPECIAL) Then
        Set SubItemIcon = Items(Index).SubItems(SubItemIndex).SmallIc
    End If
End Function

Public Function SubItem(Index As Integer, SubItemIndex As Integer) As String
    'On error resume next
    If (Items(Index).MenuType = MENU_NORMAL Or Items(Index).MenuType = MENU_SPECIAL) Then
        SubItem = Items(Index).SubItems(SubItemIndex).Msg
    End If
End Function

Public Sub ClsSubItems(Index As Integer)
    'On error resume next
    If (Items(Index).MenuType = MENU_NORMAL Or Items(Index).MenuType = MENU_SPECIAL) Then
        ReDim Items(Index).SubItems(0)
        Items(Index).SubItemsCount = 0
        Call ReDraw
    End If
End Sub

Private Function MakeMSG(ByVal Msg As String, ByVal Width As Long) As String
    Dim Splitup() As String, HeelWoord As String, Letter As String
    Dim a         As Long, Woord       As String
    
    Splitup() = Split(Msg)
    For a = 0 To UBound(Splitup())
        Letter = Splitup(a)
        If (UserControl.TextWidth(Woord & " " & Letter) >= Width) Then
            HeelWoord = HeelWoord & Woord & vbCrLf
            Woord = Letter
        Else
            If (Woord = "") Then
                Woord = Letter
            Else
                Woord = Woord & " " & Letter
            End If
        End If
    Next
    HeelWoord = HeelWoord & Woord
    MakeMSG = HeelWoord
End Function

Public Sub SetBackGround(Index As Integer, Optional NewBack As IPictureDisp)
    'On error resume next
    If (NewBack Is Nothing) Then
        Set Items(Index).BackPic = Nothing
    Else
        Set Items(Index).BackPic = NewBack
    End If
    Call ReDraw
End Sub

Public Sub SetBigIcon(Index As Integer, Optional NewIcon As IPictureDisp)
    'On error resume next
    If (NewIcon Is Nothing) Then
        Set Items(Index).BigIc = Nothing
    Else
        Set Items(Index).BigIc = NewIcon
    End If
    Call ReDraw
End Sub

Public Sub SetDetailsPicture(Index As Integer, Optional NewPic As IPictureDisp)
    'On error resume next
    If (Items(Index).MenuType = MENU_DETAILS) Then
        If (NewPic Is Nothing) Then
            Set Items(Index).DetailPic = Nothing
        Else
            Set Items(Index).DetailPic = NewPic
        End If
        Call ReDraw
    End If
End Sub

Public Sub SetDetailsTitle(Index As Integer, NewTitle As String)
    'On error resume next
    If (Items(Index).MenuType = MENU_DETAILS) Then
        Items(Index).Detail_Title = NewTitle
        Call ReDraw
    End If
End Sub

Public Sub SetDetails(Index As Integer, NewDetails As String)
    'On error resume next
    If (Items(Index).MenuType = MENU_DETAILS) Then
        Items(Index).Detail_Msg = NewDetails
        Call ReDraw
    End If
End Sub




'* Formatted using this program <FormatCode 1.0> you can download _
    in this page www.geocities.com/hackprotm/FormatCode1.0.zip.

Public Function Collapse(Item As Integer)
Items(Item).CollapsedOrNot = False
ReDraw
End Function

Public Function Expand(Item As Integer)
Items(Item).CollapsedOrNot = True
ReDraw
End Function
