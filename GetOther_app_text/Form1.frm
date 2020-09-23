VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5445
   ClientLeft      =   1125
   ClientTop       =   1575
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5445
   ScaleWidth      =   6990
   Begin VB.TextBox ResultsText 
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   720
      Width           =   6975
   End
   Begin VB.TextBox AppText 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Exploring - StrAddr"
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton CmdFindText 
      Caption         =   "Find Text"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Application"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDFIRST = 0

Private Const LVM_FIRST = &H1000&
Private Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)

Private Const LVM_GETITEMA = (LVM_FIRST + 5)
Private Const LVM_GETITEMW = (LVM_FIRST + 75)

Private Const LVM_GETITEMTEXTA = (LVM_FIRST + 45)
Private Const LVM_GETITEMTEXTW = (LVM_FIRST + 115)


Private Const LVIF_TEXT = &H1
Private Const LVIF_IMAGE = &H2
Private Const LVIF_PARAM = &H4
Private Const LVIF_STATE = &H8
Private Const LVIF_INDENT = &H10
Private Const LVIF_NORECOMPUTE = &H800
Private Const LVIS_FOCUSED = &H1
Private Const LVIS_SELECTED = &H2
Private Const LVIS_CUT = &H4
Private Const LVIS_DROPHILITED = &H8
Private Const LVIS_ACTIVATING = &H20
Private Const LVIS_OVERLAYMASK = &HF00
Private Const LVIS_STATEIMAGEMASK = &HF000

Private Type LVITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As Long
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
' Return information about this window and its
' children.
Public Function WindowInfo(window_hwnd As Long)
Dim txt As String
Dim buf As String
Dim buflen As Long
Dim child_hwnd As Long
Dim children() As Long
Dim num_children As Integer
Dim i As Integer
Dim lvi As LVITEM

    ' Get the class name.
    buflen = 256
    buf = Space$(buflen - 1)
    buflen = GetClassName(window_hwnd, buf, buflen)
    buf = Left$(buf, buflen)
    txt = "Class: " & buf & vbCrLf

    ' hWnd.
    txt = txt & "    hWnd: " & _
        Format$(window_hwnd) & vbCrLf
        
    ' Associated text.
    txt = txt & "    Text: [" & _
        WindowText(window_hwnd) & "]" & vbCrLf

    ' Make a list of the child windows.
    num_children = 0
    child_hwnd = GetWindow(window_hwnd, GW_CHILD)
    Do While child_hwnd <> 0
        num_children = num_children + 1
        ReDim Preserve children(1 To num_children)
        children(num_children) = child_hwnd

        child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
    Loop

    ' Get information on the child windows.
    For i = 1 To num_children
        txt = txt & WindowInfo(children(i))
    Next i

    WindowInfo = txt
End Function
' Return the text associated with the window.
Public Function WindowText(window_hwnd As Long) As String
Dim txtlen As Long
Dim txt As String

    WindowText = ""
    If window_hwnd = 0 Then Exit Function
    
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
    If txtlen = 0 Then Exit Function
    
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
    WindowText = Left$(txt, txtlen)
End Function

Private Sub CmdFindText_Click()
Dim app_name As String
Dim parent_hwnd As Long

    app_name = AppText.Text
    parent_hwnd = FindWindow(vbNullString, app_name)
    If parent_hwnd = 0 Then
        MsgBox "Application not found."
        Exit Sub
    End If

    ResultsText.Text = app_name & vbCrLf & _
        vbCrLf & WindowInfo(parent_hwnd)
End Sub

Private Sub Form_Resize()
Dim wid As Single
Dim hgt As Single
Dim t As Single

    wid = ScaleWidth
    t = CmdFindText.Top + CmdFindText.Height
    hgt = ScaleHeight - t
    ResultsText.Move 0, t, wid, hgt
End Sub


