VERSION 5.00
Begin VB.UserControl Check 
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   ScaleHeight     =   1485
   ScaleWidth      =   1695
   ToolboxBitmap   =   "Check.ctx":0000
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1110
      Top             =   720
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   -30
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Image CheckedDisabled 
      Height          =   195
      Left            =   720
      Picture         =   "Check.ctx":0312
      Top             =   930
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image DisabledNoValue 
      Height          =   195
      Left            =   720
      Picture         =   "Check.ctx":0D3F
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image DisabledGrey 
      Height          =   195
      Left            =   720
      Picture         =   "Check.ctx":1747
      Top             =   1140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image PushedNoValue 
      Height          =   195
      Left            =   510
      Picture         =   "Check.ctx":217B
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image PushedGreyed 
      Height          =   195
      Left            =   510
      Picture         =   "Check.ctx":2CD3
      Top             =   1140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image PushedChecked 
      Height          =   195
      Left            =   510
      Picture         =   "Check.ctx":3867
      Top             =   930
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image MouseoverGreyed 
      Height          =   195
      Left            =   300
      Picture         =   "Check.ctx":440D
      Top             =   1140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image MouseoverChecked 
      Height          =   195
      Left            =   300
      Picture         =   "Check.ctx":4FE5
      Top             =   930
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IdleGreyed 
      Height          =   195
      Left            =   90
      Picture         =   "Check.ctx":5BDF
      Top             =   1140
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image MouseoverNoValue 
      Height          =   195
      Left            =   300
      Picture         =   "Check.ctx":6753
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IdleChecked 
      Height          =   195
      Left            =   90
      Picture         =   "Check.ctx":7307
      Top             =   930
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IdleNoValue 
      Height          =   195
      Left            =   90
      Picture         =   "Check.ctx":7EBE
      Top             =   720
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Website    : wwww.olivault.com
' Contact HP : ?
' E-mail     : intouch@olivault.com
'
'                                  Roes Love Maryam
'
'Note       : This Code Source is destined to You which wish to learn
'             programming.by using is Visual Basic 6.0. If You use this code source,
'             expect that remain to mention the name of me in part of Your About
'             application( Credit Title) as well as in part of Your place code source
'             using it ( IDEA of VB6). Usage of code source for the purpose of is
'             commercial / profit, HAVE TO PERMIT OF its OWNER.
'             Trespasser- an of this thing can be ensnared by penalization
'             related to misdemeanour of Copyrights and [Code/Law] Rights Of Intellectual.
'---------------------------------------------------------------------------------------
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function IsChild Lib "user32.dll" (ByVal hWndParent As Long, ByVal hWnd As Long) As Long
Private Type POINT_TYPE
x As Long
y As Long
End Type
Enum CheckTypes
Unchecked = 0
Checked = 1
Greyed = 2
End Enum
Dim Respond As Boolean
Dim OValue As CheckTypes
Dim isEnabled As Boolean
Event Click()
Event DblClick()
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Default Property Values:
Const m_def_DisabledColor = &HC0C0C0
'Property Variables:
Dim m_ForeColor As OLE_COLOR
Dim m_DisabledColor As OLE_COLOR
Private Sub picture1_DblClick()
On Error Resume Next
If Respond = False Then Exit Sub
RaiseEvent DblClick
End Sub
Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Respond = False Then Exit Sub
RaiseEvent MouseDown(Button, Shift, x, y)
If Button = 2 Then Exit Sub
Timer1.Enabled = False
Select Case OValue
Case Checked
Picture1.Picture = PushedChecked.Picture: OValue = 0
Case Unchecked
Picture1.Picture = PushedNoValue.Picture: OValue = 1
Case Greyed
Picture1.Picture = PushedGreyed.Picture: OValue = 0
End Select
End Sub
Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Respond = False Or Button = 2 Then Exit Sub
RaiseEvent Click
Timer1.Enabled = True
End Sub
Private Sub Label1_DblClick()
On Error Resume Next
RaiseEvent DblClick
End Sub
Private Sub check1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
picture1_MouseDown Button, Shift, x, y
End Sub
Private Sub check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
picture1_MouseUp Button, Shift, x, y
End Sub
Private Sub Timer1_Timer()
On Error Resume Next
Dim CursorPos As POINT_TYPE
GetCursorPos CursorPos
If Enabled = False Then Exit Sub
po = WindowFromPoint(CursorPos.x, CursorPos.y)
If WindowFromPoint(CursorPos.x, CursorPos.y) = Check1.hWnd Or WindowFromPoint(CursorPos.x, CursorPos.y) = Picture1.hWnd Then
Select Case OValue
Case Checked
Picture1.Picture = MouseoverChecked.Picture
Case Unchecked
Picture1.Picture = MouseoverNoValue.Picture
Case Greyed
Picture1.Picture = MouseoverGreyed.Picture
End Select
Else
Select Case OValue
Case Checked
Picture1.Picture = IdleChecked.Picture
Case Unchecked
Picture1.Picture = IdleNoValue.Picture
Case Greyed
Picture1.Picture = IdleGreyed.Picture
End Select
End If
End Sub
Property Let Value(Yes As CheckTypes)
On Error Resume Next
OValue = Yes
Select Case Yes
Case Checked
If isEnabled = True Then Picture1.Picture = IdleChecked.Picture Else Picture1.Picture = CheckedDisabled.Picture
Case Unchecked
If isEnabled = True Then Picture1.Picture = IdleNoValue.Picture Else Picture1.Picture = DisabledNoValue.Picture
Case Greyed
If isEnabled = True Then Picture1.Picture = IdleGreyed.Picture Else Picture1.Picture = DisabledGrey.Picture
End Select
End Property
Property Get Value() As CheckTypes
On Error Resume Next
Select Case OValue
Case Checked
If isEnabled = True Then Picture1.Picture = IdleChecked.Picture Else Picture1.Picture = CheckedDisabled.Picture
Case Unchecked
If isEnabled = True Then Picture1.Picture = IdleNoValue.Picture Else Picture1.Picture = DisabledNoValue.Picture
Case Greyed
If isEnabled = True Then Picture1.Picture = IdleGreyed.Picture Else Picture1.Picture = DisabledGrey.Picture
End Select
'************************
Value = OValue
End Property
Private Sub UserControl_InitProperties()
On Error Resume Next
Value = False
Caption = Name
Enabled = True
Set Font = Parent.Font
m_ForeColor = m_def_ForeColor
m_DisabledColor = m_def_DisabledColor
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
Value = PropBag.ReadProperty("Value", Unchecked)
Caption = PropBag.ReadProperty("Caption", Name)
Enabled = PropBag.ReadProperty("Enabled", True)
Set Font = PropBag.ReadProperty("Font", Parent.Font)
With Check1
.ForeColor = PropBag.ReadProperty("ForeColor", &H800000)
.Caption = PropBag.ReadProperty("Caption", UserControl.Name)
.Backcolor = PropBag.ReadProperty("BackColor", &HFFFFFF)
End With
m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
m_DisabledColor = PropBag.ReadProperty("DisabledColor", m_def_DisabledColor)
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
Picture1.Top = (UserControl.Height / 2) - Picture1.Height / 2
With Check1
UserControl.Backcolor = .Backcolor
If Enabled = True Then .ForeColor = ForeColor Else .ForeColor = DisabledColor
.Width = UserControl.Width - .Left
.Height = UserControl.Height
.Top = ((UserControl.Height / 2) - .Height / 2)
End With
End Sub
Private Sub UserControl_Show()
On Error Resume Next
Timer1.Enabled = UserControl.Ambient.UserMode
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty "Value", OValue, Unchecked
PropBag.WriteProperty "Caption", Check1.Caption, Name
PropBag.WriteProperty "Enabled", isEnabled, True
PropBag.WriteProperty "Font", Check1.Font, Parent.Font
Call PropBag.WriteProperty("ForeColor", Check1.ForeColor, &H800000)
Call PropBag.WriteProperty("Caption", Check1.Caption, "Check1")
Call PropBag.WriteProperty("BackColor", Check1.Backcolor, &HFFFFFF)
Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
Call PropBag.WriteProperty("DisabledColor", m_DisabledColor, m_def_DisabledColor)
End Sub
Public Property Set Font(newFont As IFontDisp)
Set Check1.Font = newFont
End Property
Public Property Get Font() As IFontDisp
Set Font = Check1.Font
End Property
Property Let Enabled(Yes As Boolean)
On Error Resume Next
isEnabled = Yes
If Yes = True Then
Select Case OValue
Case Checked
Picture1.Picture = IdleChecked.Picture
Case Unchecked
Picture1.Picture = IdleNoValue.Picture
Case Greyed
Picture1.Picture = IdleGreyed.Picture
End Select
Respond = True
Check1.ForeColor = ForeColor
Timer1.Enabled = UserControl.Ambient.UserMode
Else
Select Case OValue
Case Checked
Picture1.Picture = DisabledChecked.Picture
Case Unchecked
Picture1.Picture = DisabledNoValue.Picture
Case Greyed
Picture1.Picture = DisabledGrey.Picture
End Select
Timer1.Enabled = False
Respond = False
Check1.ForeColor = DisabledColor
End If
End Property
Property Get Enabled() As Boolean
Enabled = isEnabled
End Property

Public Property Get Caption() As String
Caption = Check1.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
Check1.Caption() = New_Caption
PropertyChanged "Caption"
End Property

Public Property Get Backcolor() As OLE_COLOR
Backcolor = Check1.Backcolor
End Property
Public Property Let Backcolor(ByVal New_BackColor As OLE_COLOR)
Check1.Backcolor() = New_BackColor
UserControl.Backcolor = New_BackColor
PropertyChanged "BackColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
m_ForeColor = New_ForeColor
PropertyChanged "ForeColor"
Check1.ForeColor = m_ForeColor
End Property
Public Property Get DisabledColor() As OLE_COLOR
DisabledColor = m_DisabledColor
End Property
Public Property Let DisabledColor(ByVal New_DisabledColor As OLE_COLOR)
m_DisabledColor = New_DisabledColor
PropertyChanged "DisabledColor"
End Property
