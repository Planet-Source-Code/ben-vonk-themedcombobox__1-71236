VERSION 5.00
Begin VB.UserControl ThemedComboBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   384
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2304
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ThemedComboBox.ctx":0000
   ScaleHeight     =   384
   ScaleWidth      =   2304
   ToolboxBitmap   =   "ThemedComboBox.ctx":0542
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   1920
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1440
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   28
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   960
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   312
   End
   Begin VB.PictureBox picButtonState 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   480
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   26
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   312
   End
   Begin VB.PictureBox picImage 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "ThemedComboBox.ctx":0854
      ScaleHeight     =   240
      ScaleWidth      =   252
      TabIndex        =   0
      Top             =   0
      Width           =   252
   End
End
Attribute VB_Name = "ThemedComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ThemedComboBox Control
'
'Author Ben Vonk
'10-10-2008 First version, included: Paul Caton's self Subclass v1.1.0008
'30-11-2011 Second version, Add properties so the user can customize the button

Option Explicit

' Private Constants
Private Const ALL_MESSAGES        As Long = -1
Private Const CB_GETDROPPEDSTATE  As Long = &H157
Private Const CBP_ARROWBTN        As Long = 1
Private Const GWL_WNDPROC         As Long = -4
Private Const PATCH_05            As Long = 93
Private Const PATCH_09            As Long = 137
Private Const WM_ACTIVATE         As Long = &H6
Private Const WM_COMMAND          As Long = &H111
Private Const WM_DESTROY          As Long = &H2
Private Const WM_LBUTTONDOWN      As Long = &H201
Private Const WM_LBUTTONUP        As Long = &H202
Private Const WM_MOUSEMOVE        As Long = &H200
Private Const WM_PAINT            As Long = &HF
Private Const WM_THEMECHANGED     As Long = &H31A
Private Const WM_TIMER            As Long = &H113

' Public Enumerations
Public Enum BorderColorStyles
   ThemeColors
   CustomColors
End Enum

Public Enum ButtonThemeTypes
   Windows
   User
End Enum

' Private Enumerations
Private Enum ButtonStates
   Normal
   Over
   Pressed
   Disabled
End Enum

Private Enum ControlState
   StateNormal
   StateOver
   StateFocus
   StateDown
   StateDisabled
   StateUp
End Enum

Private Enum MsgWhen
   MSG_AFTER = 1
   MSG_BEFORE = 2
   MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE
End Enum

' Private Types
Private Type OSVersionInfo
   dwOSVersionInfoSize            As Long
   dwMajorVersion                 As Long
   dwMinorVersion                 As Long
   dwBuildNumber                  As Long
   dwPlatformId                   As Long
   szCSDVersion                   As String * 128
End Type

Private Type PointAPI
   X                              As Long
   Y                              As Long
End Type

Private Type Rect
   Left                           As Long
   Top                            As Long
   Right                          As Long
   Bottom                         As Long
End Type

Private Type ComboBoxInfo
   cbSize                         As Long
   rcItem                         As Rect
   rcButton                       As Rect
   lStateButton                   As Long
   hWndCombo                      As Long
   hWndEdit                       As Long
   hWndList                       As Long
End Type

Private Type SubclassDataType
   hWnd                           As Long
   nAddrSclass                    As Long
   nAddrOrig                      As Long
   nMsgCountA                     As Long
   nMsgCountB                     As Long
   aMsgTabelA()                   As Long
   aMsgTabelB()                   As Long
End Type

' Private Variables
Private ButtonDown                As Boolean
Private IsThemed                  As Boolean
Private IsThemedWindows           As Boolean
Private m_Activated               As Boolean
Private MouseOver                 As Boolean
Private m_BorderColorStyle        As BorderColorStyles
Private m_ButtonThemeType         As ButtonThemeTypes
Private ButtonState               As ControlState
Private DefaultBorderColor        As Long
Private m_ComboBoxBorderColor     As Long
Private m_DriveListBoxBorderColor As Long
Private m_ImageComboBorderColor   As Long
Private SubclassMemory            As Long
Private TimerID                   As Long
Private SubclassData()            As SubclassDataType

' Private API's
Private Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "GDI32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LineTo Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As PointAPI) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function FreeLibrary Lib "Kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function StrLen Lib "Kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function FrameRect Lib "User32" (ByVal hDC As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As Rect) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As PointAPI) As Long
Private Declare Function GetComboBoxInfo Lib "User32" (ByVal hWndCombo As Long, ByRef pcbi As ComboBoxInfo) As Long
Private Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function GetSysColor Lib "User32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsWindowEnabled Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function KillTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetTimer Lib "User32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function SetWindowLongA Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function CloseThemeData Lib "UxTheme" (ByVal lngTheme As Long) As Long
Private Declare Function DrawThemeBackground Lib "UxTheme" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As Rect, pClipRect As Rect) As Long
Private Declare Function GetCurrentThemeName Lib "UxTheme" (ByVal pszThemeFileName As Long, ByVal cchMaxNameChars As Long, ByVal pszColorBuff As Long, ByVal cchMaxColorChars As Long, ByVal pszSizeBuff As Long, ByVal cchMaxSizeChars As Long) As Long
Private Declare Function GetThemeDocumentationProperty Lib "UxTheme" (ByVal pszThemeName As Long, ByVal pszPropertyName As Long, ByVal pszValueBuff As Long, ByVal cchMaxValChars As Long) As Long
Private Declare Function OpenThemeData Lib "UxTheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Subclass_WndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lhWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

Const CBN_CLOSEUP  As Long = 8

Static lngHandle   As Long

Dim lngListWindow  As Long

   Select Case uMsg
      Case WM_ACTIVATE
         If Not m_Activated Then If IsThemed Then Call Initialize
         
      Case WM_COMMAND
         If lngHandle = lParam Then
            If wParam \ &H10000 = CBN_CLOSEUP Then
               If ButtonDown Then ButtonDown = False
               
               ButtonState = StateNormal
               MouseOver = False
               TimerID = KillTimer(lhWnd, TimerID)
            End If
            
            Call DrawComboBox(lngHandle)
         End If
         
      Case WM_DESTROY
         Call Subclass_Stop(lhWnd)
         
      Case WM_LBUTTONDOWN
         If lhWnd = lngHandle Then
            ButtonState = StateDown
            
            Call DrawComboBox(lhWnd)
         End If
         
      Case WM_LBUTTONUP
         If lhWnd = lngHandle Then
            ButtonState = StateUp
            
            Call DrawComboBox(lhWnd)
         End If
         
      Case WM_MOUSEMOVE
         If InControl(lhWnd) Then
            lngHandle = lhWnd
            
            If Not MouseOver Then
               MouseOver = True
               ButtonState = StateOver
               TimerID = SetTimer(lhWnd, TimerID, 1, SubclassData(Subclass_Index(lhWnd)).nAddrSclass)
               
               Call DrawComboBox(lhWnd)
            End If
            
         Else
            ButtonState = StateDown
         End If
         
      Case WM_PAINT
         GetComboBoxButton lhWnd, lngListWindow
         
         If lhWnd = lngListWindow Then
            Call DrawComboBoxListWindow(lhWnd)
            
         Else
            Call DrawComboBox(lhWnd)
         End If
         
      Case WM_THEMECHANGED
         IsThemed = CheckIsThemed
         
      Case WM_TIMER
         If InControl(lhWnd) Then
            MouseOver = True
            
            If (ButtonState <> StateDown) And SendMessage(lhWnd, CB_GETDROPPEDSTATE, 0, ByVal 0&) Then ButtonState = StateOver
            
         Else
            MouseOver = False
            ButtonState = StateNormal
            TimerID = KillTimer(lhWnd, TimerID)
            
            Call DrawComboBox(lhWnd)
         End If
   End Select

End Sub

Private Function Subclass_AddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long

   Subclass_AddrFunc = GetProcAddress(GetModuleHandle(sDLL), sProc)
   Debug.Assert Subclass_AddrFunc

End Function

Private Function Subclass_Index(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean) As Long

   For Subclass_Index = UBound(SubclassData) To 0 Step -1
      If SubclassData(Subclass_Index).hWnd = lhWnd Then
         If Not bAdd Then Exit Function
         
      ElseIf SubclassData(Subclass_Index).hWnd = 0 Then
         If bAdd Then Exit Function
      End If
   Next 'Subclass_Index
   
   If Not bAdd Then Debug.Assert False

End Function

Private Function Subclass_InIDE() As Boolean

   Debug.Assert Subclass_SetTrue(Subclass_InIDE)

End Function

Private Function Subclass_Initialize(ByVal lhWnd As Long) As Long

Const CODE_LEN                  As Long = 200
Const GMEM_FIXED                As Long = 0
Const PATCH_01                  As Long = 18
Const PATCH_02                  As Long = 68
Const PATCH_03                  As Long = 78
Const PATCH_06                  As Long = 116
Const PATCH_07                  As Long = 121
Const PATCH_0A                  As Long = 186
Const FUNC_CWP                  As String = "CallWindowProcA"
Const FUNC_EBM                  As String = "EbMode"
Const FUNC_SWL                  As String = "SetWindowLongA"
Const MOD_USER                  As String = "User32"
Const MOD_VBA5                  As String = "vba5"
Const MOD_VBA6                  As String = "vba6"

Static bytBuffer(1 To CODE_LEN) As Byte
Static lngCWP                   As Long
Static lngEbMode                As Long
Static lngSWL                   As Long

Dim lngCount                    As Long
Dim lngIndex                    As Long
Dim strHex                      As String

   If bytBuffer(1) Then
      lngIndex = Subclass_Index(lhWnd, True)
      
      If lngIndex = -1 Then
         lngIndex = UBound(SubclassData) + 1
         
         ReDim Preserve SubclassData(lngIndex) As SubclassDataType
      End If
      
      Subclass_Initialize = lngIndex
      
   Else
      strHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D0000005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D000000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
      
      For lngCount = 1 To CODE_LEN
         bytBuffer(lngCount) = Val("&H" & Left(strHex, 2))
         strHex = Mid(strHex, 3)
      Next 'lngCount
      
      If Subclass_InIDE Then
         bytBuffer(16) = &H90
         bytBuffer(17) = &H90
         lngEbMode = Subclass_AddrFunc(MOD_VBA6, FUNC_EBM)
         
         If lngEbMode = 0 Then lngEbMode = Subclass_AddrFunc(MOD_VBA5, FUNC_EBM)
      End If
      
      lngCWP = Subclass_AddrFunc(MOD_USER, FUNC_CWP)
      lngSWL = Subclass_AddrFunc(MOD_USER, FUNC_SWL)
      
      ReDim SubclassData(0) As SubclassDataType
   End If
   
   With SubclassData(lngIndex)
      .hWnd = lhWnd
      .nAddrSclass = GlobalAlloc(GMEM_FIXED, CODE_LEN)
      .nAddrOrig = SetWindowLongA(.hWnd, GWL_WNDPROC, .nAddrSclass)
      
      Call CopyMemory(ByVal .nAddrSclass, bytBuffer(1), CODE_LEN)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_01, lngEbMode)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_02, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_03, lngSWL)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_06, .nAddrOrig)
      Call Subclass_PatchRel(.nAddrSclass, PATCH_07, lngCWP)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_0A, ObjPtr(Me))
   End With

End Function

Private Function Subclass_SetTrue(ByRef bValue As Boolean) As Boolean

   Subclass_SetTrue = True
   bValue = True

End Function

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As MsgWhen = MSG_AFTER)

   With SubclassData(Subclass_Index(lhWnd))
      If When And MSG_BEFORE Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelB, .nMsgCountB, MSG_BEFORE, .nAddrSclass)
      If When And MSG_AFTER Then Call Subclass_DoAddMsg(uMsg, .aMsgTabelA, .nMsgCountA, MSG_AFTER, .nAddrSclass)
   End With

End Sub

Private Sub Subclass_DoAddMsg(ByVal uMsg As Long, ByRef aMsgTabel() As Long, ByRef nMsgCount As Long, ByVal When As MsgWhen, ByVal nAddr As Long)

Const PATCH_04 As Long = 88
Const PATCH_08 As Long = 132

Dim lngEntry   As Long

   ReDim lngOffset(1) As Long
   
   If uMsg = ALL_MESSAGES Then
      nMsgCount = ALL_MESSAGES
      
   Else
      For lngEntry = 1 To nMsgCount - 1
         If aMsgTabel(lngEntry) = 0 Then
            aMsgTabel(lngEntry) = uMsg
            
            GoTo ExitSub
            
         ElseIf aMsgTabel(lngEntry) = uMsg Then
            GoTo ExitSub
         End If
      Next 'lngEntry
      
      nMsgCount = nMsgCount + 1
      
      ReDim Preserve aMsgTabel(1 To nMsgCount) As Long
      
      aMsgTabel(nMsgCount) = uMsg
   End If
   
   If When = MSG_BEFORE Then
      lngOffset(0) = PATCH_04
      lngOffset(1) = PATCH_05
      
   Else
      lngOffset(0) = PATCH_08
      lngOffset(1) = PATCH_09
   End If
   
   If uMsg <> ALL_MESSAGES Then Call Subclass_PatchVal(nAddr, lngOffset(0), VarPtr(aMsgTabel(1)))
   
   Call Subclass_PatchVal(nAddr, lngOffset(1), nMsgCount)
   
ExitSub:
   Erase lngOffset

End Sub

Private Sub Subclass_PatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)

End Sub

Private Sub Subclass_PatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)

   Call CopyMemory(ByVal nAddr + nOffset, nValue, 4)

End Sub

Private Sub Subclass_Stop(ByVal lhWnd As Long)

   With SubclassData(Subclass_Index(lhWnd))
      SetWindowLongA .hWnd, GWL_WNDPROC, .nAddrOrig
      
      Call Subclass_PatchVal(.nAddrSclass, PATCH_05, 0)
      Call Subclass_PatchVal(.nAddrSclass, PATCH_09, 0)
      
      GlobalFree .nAddrSclass
      .hWnd = 0
      .nMsgCountA = 0
      .nMsgCountB = 0
      Erase .aMsgTabelA, .aMsgTabelB
   End With

End Sub

Private Sub Subclass_Terminate()

Dim lngCount As Long

   For lngCount = UBound(SubclassData) To 0 Step -1
      If SubclassData(lngCount).hWnd Then Call Subclass_Stop(SubclassData(lngCount).hWnd)
   Next 'lngCount

End Sub

Public Property Get BorderColorStyle() As BorderColorStyles
Attribute BorderColorStyle.VB_Description = "Returns/sets the border style for an object."

   BorderColorStyle = m_BorderColorStyle

End Property

Public Property Let BorderColorStyle(ByVal NewBorderColorStyle As BorderColorStyles)

   If NewBorderColorStyle < ThemeColors Then NewBorderColorStyle = ThemeColors
   If NewBorderColorStyle > CustomColors Then NewBorderColorStyle = CustomColors
   
   m_BorderColorStyle = NewBorderColorStyle
   PropertyChanged "BorderColorStyle"

End Property

Public Property Get ButtonDisabled() As StdPicture
Attribute ButtonDisabled.VB_Description = "Returns/sets a graphic to be displayed when the control is disabled. (Only if ButtonThemeType is set to User!)"

   Set ButtonDisabled = picButtonState.Item(Disabled).Picture

End Property

Public Property Let ButtonDisabled(ByVal NewButtonDisabled As StdPicture)

   Set ButtonDisabled = NewButtonDisabled

End Property

Public Property Set ButtonDisabled(ByVal NewButtonDisabled As StdPicture)

   picButtonState.Item(Disabled).Picture = NewButtonDisabled
   PropertyChanged "ButtonDisabled"

End Property

Public Property Get ButtonNormal() As StdPicture
Attribute ButtonNormal.VB_Description = "Returns/sets a graphic to be displayed in an button normal state of the control. (Only if ButtonThemeType is set to User!)"

   Set ButtonNormal = picButtonState.Item(Normal).Picture

End Property

Public Property Let ButtonNormal(ByVal NewButtonNormal As StdPicture)

   Set ButtonNormal = NewButtonNormal

End Property

Public Property Set ButtonNormal(ByVal NewButtonNormal As StdPicture)

   picButtonState.Item(Normal).Picture = NewButtonNormal
   PropertyChanged "ButtonNormal"
   
   Call CheckButtonThemeType

End Property

Public Property Get ButtonOver() As StdPicture
Attribute ButtonOver.VB_Description = "Returns/sets a graphic to be displayed in an button over state of the control. (Only if ButtonThemeType is set to User!)"

   Set ButtonOver = picButtonState.Item(Over).Picture

End Property

Public Property Let ButtonOver(ByVal NewButtonOver As StdPicture)

   Set ButtonOver = NewButtonOver

End Property

Public Property Set ButtonOver(ByVal NewButtonOver As StdPicture)

   picButtonState.Item(Over).Picture = NewButtonOver
   PropertyChanged "ButtonOver"

End Property

Public Property Get ButtonPressed() As StdPicture
Attribute ButtonPressed.VB_Description = "Returns/sets a graphic to be displayed in an button pressed state of the control. (Only if ButtonThemeType is set to User!)"

   Set ButtonPressed = picButtonState.Item(Pressed).Picture

End Property

Public Property Let ButtonPressed(ByVal NewButtonPressed As StdPicture)

   Set ButtonPressed = NewButtonPressed

End Property

Public Property Set ButtonPressed(ByVal NewButtonPressed As StdPicture)

   picButtonState.Item(Pressed).Picture = NewButtonPressed
   PropertyChanged "ButtonPressed"

End Property

Public Property Get ButtonThemeType() As ButtonThemeTypes
Attribute ButtonThemeType.VB_Description = "Returns/sets a theme type of the ThemedComboBox control."

   ButtonThemeType = m_ButtonThemeType

End Property

Public Property Let ButtonThemeType(ByVal NewButtonThemeType As ButtonThemeTypes)

   m_ButtonThemeType = NewButtonThemeType
   PropertyChanged "ButtonThemeType"
   
   Call CheckButtonThemeType

End Property

Public Property Get ComboBoxBorderColor() As OLE_COLOR
Attribute ComboBoxBorderColor.VB_Description = "Returns/sets the color of an ComboBox border."

   ComboBoxBorderColor = m_ComboBoxBorderColor

End Property

Public Property Let ComboBoxBorderColor(ByVal NewComboBoxBorderColor As OLE_COLOR)

   m_ComboBoxBorderColor = NewComboBoxBorderColor
   PropertyChanged "ComboBoxBorderColor"

End Property

Public Property Get DriveListBoxBorderColor() As OLE_COLOR
Attribute DriveListBoxBorderColor.VB_Description = "Returns/sets the color of an DriveListBox border."

   DriveListBoxBorderColor = m_DriveListBoxBorderColor

End Property

Public Property Let DriveListBoxBorderColor(ByVal NewDriveListBoxBorderColor As OLE_COLOR)

   m_DriveListBoxBorderColor = NewDriveListBoxBorderColor
   PropertyChanged "DriveListBoxBorderColor"

End Property

Public Property Get ImageComboBorderColor() As OLE_COLOR
Attribute ImageComboBorderColor.VB_Description = "Returns/sets the color of an ImageCombo border."

   ImageComboBorderColor = m_ImageComboBorderColor

End Property

Public Property Let ImageComboBorderColor(ByVal NewImageComboBoxBorderColor As OLE_COLOR)

   m_ImageComboBorderColor = NewImageComboBoxBorderColor
   PropertyChanged "ImageComboBorderColor"

End Property

Public Function Activated() As Boolean

   Activated = m_Activated

End Function

Private Function CheckIsComboBox(ByRef hWnd As Long, Optional ByRef ComboBoxBorderColor As Long) As Boolean

Dim strClassName As String * 255

   Select Case Left(strClassName, GetClassName(hWnd, strClassName, Len(strClassName)))
      Case "ImageCombo20WndClass"
         CheckIsComboBox = True
         ComboBoxBorderColor = m_ImageComboBorderColor
         hWnd = FindWindowEx(hWnd, 0, "ComboBox", ByVal 0&)
         
      Case "ThunderComboBox", "ThunderRT6ComboBox"
         CheckIsComboBox = True
         ComboBoxBorderColor = m_ComboBoxBorderColor
         
      Case "ThunderDriveListBox", "ThunderRT6DriveListBox"
         CheckIsComboBox = True
         ComboBoxBorderColor = m_DriveListBoxBorderColor
   End Select

End Function

Private Function CheckIsThemed() As Boolean

Const VER_PLATFORM_WIN32_NT As Long = 2

Dim lngLibrary              As Long
Dim osvInfo                 As OSVersionInfo
Dim strName                 As String
Dim strTheme                As String

   IsThemedWindows = False
   
   With osvInfo
      .dwOSVersionInfoSize = Len(osvInfo)
      GetVersionEx osvInfo
      
      If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
         If ((.dwMajorVersion > 4) And .dwMinorVersion) Or (.dwMajorVersion > 5) Then
            IsThemedWindows = True
            lngLibrary = LoadLibrary("UxTheme")
            
            If lngLibrary Then
               strTheme = String(255, vbNullChar)
               GetCurrentThemeName StrPtr(strTheme), Len(strTheme), 0, 0, 0, 0
               strTheme = StripNull(strTheme)
               
               If Len(strTheme) Then
                  strName = String(255, vbNullChar)
                  GetThemeDocumentationProperty StrPtr(strTheme), StrPtr("ThemeName"), StrPtr(strName), Len(strName)
                  CheckIsThemed = (StripNull(strName) <> "")
               End If
               
               FreeLibrary lngLibrary
            End If
         End If
      End If
   End With

End Function

Private Function GetComboBoxButton(ByVal hWnd As Long, Optional ByRef ListWindow As Long, Optional ByRef ButtonWidth As Long) As Boolean

Dim cbiCombo As ComboBoxInfo

   With cbiCombo
      .cbSize = Len(cbiCombo)
      GetComboBoxInfo hWnd, cbiCombo
      ListWindow = .hWndList
      ButtonWidth = .rcButton.Right - .rcButton.Left + 1
      GetComboBoxButton = (.lStateButton <> &H8000&)
   End With

End Function

Private Function GetDefaultBorderColor() As Long

Const EDP_EDITTEXT As Long = 1
Const EDS_ASSIST   As Long = 1

Dim lngTheme       As Long
Dim rctWindow      As Rect

   If IsThemedWindows Then
      rctWindow.Right = 4
      rctWindow.Bottom = 4
      lngTheme = OpenThemeData(hWnd, StrPtr("Edit"))
      DrawThemeBackground lngTheme, hDC, EDP_EDITTEXT, EDS_ASSIST, rctWindow, rctWindow
      CloseThemeData lngTheme
   End If
   
   GetDefaultBorderColor = GetPixel(hDC, 0, 0)

End Function

Private Function GetLongColor(ByVal Color As Long) As Long

   If Color And &H80000000 Then
      GetLongColor = GetSysColor(Color And &H7FFFFFFF)
      
   Else
      GetLongColor = Color
   End If

End Function

Private Function InControl(ByVal hWnd As Long) As Boolean

Dim ptaMouse As PointAPI

   GetCursorPos ptaMouse
   InControl = (WindowFromPoint(ptaMouse.X, ptaMouse.Y) = hWnd)

End Function

Private Function StripNull(ByVal Text As String) As String

   StripNull = Left(Text, StrLen(StrPtr(Text)))

End Function

Private Sub CheckButtonThemeType()

   If picButtonState(Normal).Picture = 0 And (m_ButtonThemeType = User) Then ButtonThemeType = Windows

End Sub

Private Sub DrawBorder(ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal Color As Long)

Dim lngBrush As Long
Dim rctFrame As Rect

   With rctFrame
      .Top = Top
      .Left = Left
      .Right = Left + Right
      .Bottom = Top + Bottom
   End With
   
   ' Draw the border around the control with the given color
   lngBrush = CreateSolidBrush(Color)
   FrameRect hDC, rctFrame, lngBrush
   DeleteObject lngBrush

End Sub

Private Sub DrawComboBox(ByVal hWnd As Long)

Const CBXS_DISABLED As Long = 4
Const CBXS_HOT      As Long = 2
Const CBXS_NORMAL   As Long = 1
Const CBXS_PRESSED  As Long = 3

Dim blnHasButton    As Boolean
Dim intBorderLine   As Integer
Dim intLine         As Integer
Dim lngButtonWidth  As Long
Dim lngColor(1)     As Long
Dim lngDC           As Long
Dim lngStateID      As Long
Dim lngTheme        As Long
Dim lngWindow       As Long
Dim rctClient       As Rect

   ' StateDisabled
   If IsWindowEnabled(hWnd) = 0 Then
      lngStateID = CBXS_DISABLED
      
   ElseIf ButtonState = StateOver Then
      If ButtonDown Then
         lngStateID = CBXS_PRESSED
         
      Else
         lngStateID = CBXS_HOT
      End If
      
   ElseIf ButtonState = StateDown Then
      lngStateID = CBXS_PRESSED
      ButtonDown = True
      
   ElseIf ButtonState = StateUp Then
      If InControl(hWnd) Then
         lngStateID = CBXS_HOT
         
      Else
         lngStateID = CBXS_NORMAL
      End If
      
      ButtonDown = False
      
   ' StateNormal or StateFocus
   ElseIf ButtonDown Then
      lngStateID = CBXS_PRESSED
      
   Else
      lngStateID = CBXS_NORMAL
   End If
   
   If Not ButtonDown And SendMessage(hWnd, CB_GETDROPPEDSTATE, 0, ByVal 0&) Then lngStateID = CBXS_NORMAL
   
   lngDC = GetDC(hWnd)
   blnHasButton = GetComboBoxButton(hWnd, , lngButtonWidth)
   GetClientRect hWnd, rctClient
   lngColor(1) = GetPixel(lngDC, 2, 2)
   lngWindow = FindWindowEx(hWnd, 0, "Edit", ByVal 0&)
   
   If m_BorderColorStyle = ThemeColors Then
      lngColor(0) = DefaultBorderColor
      
   Else
      CheckIsComboBox hWnd, lngColor(0)
   End If
   
   With rctClient
      For intLine = 0 To 1
         Call DrawLine(lngDC, .Right - lngButtonWidth - intLine, 2, .Right - lngButtonWidth - intLine, .Bottom - 2, lngColor(1))
      Next 'intLine
      
      If Not blnHasButton Then
         intBorderLine = 21 + (3 And (Screen.TwipsPerPixelY = 12))
         
         For intLine = 19 To 25
            Call DrawLine(lngDC, 0, .Top + intLine, .Right, .Top + intLine, lngColor(1 - (1 And (intLine = intBorderLine))))
         Next 'intLine
         
      ElseIf lngWindow Then
         MoveWindow lngWindow, .Left + 3, .Top + 3, .Right - lngButtonWidth - 3, .Bottom - 5, 0
      End If
      
      Call DrawBorder(lngDC, 1, 1, .Right - 2, .Bottom - 2, lngColor(1))
      Call DrawBorder(lngDC, 0, 0, .Right, .Bottom, lngColor(0))
      
      If blnHasButton Then
         .Top = 1
         .Left = .Right - lngButtonWidth
         .Right = .Right - 1
         .Bottom = .Bottom - 1
         
         If m_ButtonThemeType = Windows Then
            lngTheme = OpenThemeData(hWnd, StrPtr("ComboBox"))
            DrawThemeBackground lngTheme, lngDC, CBP_ARROWBTN, lngStateID, rctClient, rctClient
            CloseThemeData lngTheme
            
         Else
            'lngStateID = 0 - StateNormal
            'lngStateID = 1 - StateOver
            'lngStateID = 2 - StateDown
            'lngStateID = 3 - StateDisabled
            lngStateID = lngStateID - 1
            StretchBlt lngDC, .Left, .Top, .Right - .Left, .Bottom - .Top, picButtonState.Item(lngStateID).hDC, 0, 0, picButtonState.Item(lngStateID).ScaleWidth, picButtonState.Item(lngStateID).ScaleHeight, vbSrcCopy
         End If
      End If
   End With
   
   DeleteDC hWnd
   Erase lngColor

End Sub

Private Sub DrawComboBoxListWindow(ByVal hWnd As Long)

Const GWL_EXSTYLE      As Long = -20
Const GWL_STYLE        As Long = -16
Const SWP_FRAMECHANGED As Long = &H20
Const SWP_NOACTIVATE   As Long = &H10
Const SWP_NOMOVE       As Long = &H2
Const SWP_NOSIZE       As Long = &H1
Const SWP_NOZORDER     As Long = &H4
Const WS_BORDER        As Long = &H800000
Const WS_EX_CLIENTEDGE As Long = &H200

Dim lngParent          As Long
Dim lngTop             As Long
Dim rctClient(1)       As Rect

   lngParent = GetParent(hWnd)
   GetClientRect lngParent, rctClient(0)
   GetClientRect hWnd, rctClient(1)
   
   With rctClient(1)
      ' Move the ComboBox ListWindow
      lngTop = rctClient(0).Bottom - .Bottom - 2
      MoveWindow hWnd, .Left + 1, lngTop, rctClient(0).Right - 2, .Bottom + lngTop - 7, 0
   End With
   
   ' Make the conrol flat
   SetWindowLongA hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_BORDER
   SetWindowLongA hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_CLIENTEDGE
   SetWindowPos hWnd, 0, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOMOVE
   Erase rctClient
   
   Call DrawComboBox(hWnd)
   ' No more subclassing needed for this item
   Call Subclass_Stop(hWnd)

End Sub

Public Sub DrawLine(ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)

Dim lngPen(1) As Long
Dim ptaTemp   As PointAPI

   ' Draw a line in the control with the given color
   lngPen(0) = CreatePen(0, 1, GetLongColor(Color))
   lngPen(1) = SelectObject(hDC, lngPen(0))
   MoveToEx hDC, X1, Y1, ptaTemp
   LineTo hDC, X2, Y2
   SelectObject hDC, lngPen(1)
   DeleteObject lngPen(1)
   DeleteObject lngPen(0)
   Erase lngPen

End Sub

Private Sub Initialize()

Dim ctlControl As Control
Dim lngWindow  As Long

   If Ambient.UserMode Then
      On Local Error Resume Next
      
      ' Search for all ComboBoxes on the Parent
      For Each ctlControl In Parent.Controls
         Err.Clear
         m_Activated = True
         lngWindow = ctlControl.hWnd
         
         If CheckIsComboBox(lngWindow) Then
            Call Subclass_Initialize(lngWindow)
            Call Subclass_AddMsg(lngWindow, WM_COMMAND)
            Call Subclass_AddMsg(lngWindow, WM_DESTROY, MSG_BEFORE)
            Call Subclass_AddMsg(lngWindow, WM_LBUTTONDOWN, MSG_BEFORE)
            Call Subclass_AddMsg(lngWindow, WM_LBUTTONUP)
            Call Subclass_AddMsg(lngWindow, WM_MOUSEMOVE)
            Call Subclass_AddMsg(lngWindow, WM_TIMER)
            Call Subclass_AddMsg(lngWindow, WM_PAINT)
            Call Subclass_Initialize(GetParent(lngWindow))
            Call Subclass_AddMsg(GetParent(lngWindow), WM_COMMAND)
            
            ' ComboBox Style is: 1 - Simple Combo (there is no button)
            If Not GetComboBoxButton(lngWindow, lngWindow) Then
               Call Subclass_Initialize(lngWindow)
               Call Subclass_AddMsg(lngWindow, WM_PAINT)
            End If
         End If
      Next 'ctlControl
      
      On Local Error GoTo 0
      Set ctlControl = Nothing
   End If

End Sub

Private Sub UserControl_Initialize()

   IsThemed = CheckIsThemed

End Sub

Private Sub UserControl_InitProperties()

   DefaultBorderColor = GetDefaultBorderColor
   m_ComboBoxBorderColor = DefaultBorderColor
   m_DriveListBoxBorderColor = DefaultBorderColor
   m_ImageComboBorderColor = DefaultBorderColor

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   With PropBag
      DefaultBorderColor = GetDefaultBorderColor
      m_BorderColorStyle = .ReadProperty("BorderColorStyle", ThemeColors)
      picButtonState.Item(Disabled).Picture = .ReadProperty("ButtonDisabled", Nothing)
      picButtonState.Item(Normal).Picture = .ReadProperty("ButtonNormal", Nothing)
      picButtonState.Item(Over).Picture = .ReadProperty("ButtonOver", Nothing)
      picButtonState.Item(Pressed).Picture = .ReadProperty("ButtonPressed", Nothing)
      m_ButtonThemeType = .ReadProperty("ButtonThemeType", Windows)
      m_ComboBoxBorderColor = .ReadProperty("ComboBoxBorderColor", DefaultBorderColor)
      m_DriveListBoxBorderColor = .ReadProperty("DriveListBoxBorderColor", DefaultBorderColor)
      m_ImageComboBorderColor = .ReadProperty("ImageComboBorderColor", DefaultBorderColor)
   End With
   
   Call CheckButtonThemeType
   
   If IsThemedWindows Then
      ' First subclass the Parent of the UserControl
      ' So we can catch the controls when the Parent activate
      Call Subclass_Initialize(Parent.hWnd)
      Call Subclass_AddMsg(Parent.hWnd, WM_ACTIVATE)
      Call Subclass_AddMsg(Parent.hWnd, WM_THEMECHANGED)
   End If

End Sub

Private Sub UserControl_Resize()

Static blnBusy As Boolean

   If blnBusy Then Exit Sub
   
   blnBusy = True
   Width = picImage.Width
   Height = picImage.Height
   blnBusy = False

End Sub

Private Sub UserControl_Terminate()

   On Local Error GoTo ExitSub
   
   Call Subclass_Terminate
   
ExitSub:
   On Local Error GoTo 0
   Erase SubclassData

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   With PropBag
      .WriteProperty "BorderColorStyle", m_BorderColorStyle, ThemeColors
      .WriteProperty "ButtonDisabled", picButtonState.Item(Disabled).Picture, Nothing
      .WriteProperty "ButtonNormal", picButtonState.Item(Normal).Picture, Nothing
      .WriteProperty "ButtonOver", picButtonState.Item(Over).Picture, Nothing
      .WriteProperty "ButtonPressed", picButtonState.Item(Pressed).Picture, Nothing
      .WriteProperty "ButtonThemeType", m_ButtonThemeType, Windows
      .WriteProperty "ComboBoxBorderColor", m_ComboBoxBorderColor, GetDefaultBorderColor
      .WriteProperty "DriveListBoxBorderColor", m_DriveListBoxBorderColor, GetDefaultBorderColor
      .WriteProperty "ImageComboBorderColor", m_ImageComboBorderColor, GetDefaultBorderColor
   End With

End Sub
