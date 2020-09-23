VERSION 5.00
Begin VB.Form FrmOption 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - Options"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton LblOKShortCut 
      Height          =   375
      Left            =   6240
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Apply"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Uninstaller2005.XPButton LblOK 
      Height          =   375
      Left            =   6240
      TabIndex        =   19
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      Picture         =   "FrmOptions.frx":0000
      ScaleHeight     =   4215
      ScaleWidth      =   1935
      TabIndex        =   18
      Top             =   0
      Width           =   1935
   End
   Begin Uninstaller2005.XPFrame XPFrame1 
      Height          =   4095
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7223
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin Uninstaller2005.XPFrame XPFrame4 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1085
         Caption         =   "Sort List View as :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontBold        =   0   'False
         FontItalic      =   0   'False
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   120
            ScaleHeight     =   285
            ScaleWidth      =   4935
            TabIndex        =   15
            Top             =   240
            Width           =   4935
            Begin Uninstaller2005.Check chkUrut 
               Height          =   255
               Left            =   120
               TabIndex        =   16
               Top             =   0
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   450
               Value           =   1
               Caption         =   "Ascending \ Descending"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Ascending \ Descending"
               ForeColor       =   8388608
            End
         End
      End
      Begin Uninstaller2005.XPFrame XPFrame3 
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1931
         Caption         =   "Create Shortcut on :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontBold        =   0   'False
         FontItalic      =   0   'False
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            ScaleHeight     =   735
            ScaleWidth      =   5415
            TabIndex        =   9
            Top             =   240
            Width           =   5415
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "StartUp Windows"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   4
               Left            =   1920
               TabIndex        =   13
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Start Menu"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Top             =   60
               Value           =   -1  'True
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Desktop"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   11
               Top             =   360
               Width           =   1095
            End
            Begin VB.OptionButton Option1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Quick Launch"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   3
               Left            =   1920
               TabIndex        =   10
               Top             =   60
               Width           =   1575
            End
         End
      End
      Begin Uninstaller2005.XPFrame XPFrame2 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1085
         Caption         =   "Show List View as :"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
         FontBold        =   0   'False
         FontItalic      =   0   'False
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   280
            Left            =   120
            ScaleHeight     =   285
            ScaleWidth      =   5295
            TabIndex        =   3
            Top             =   240
            Width           =   5295
            Begin VB.OptionButton OptIconBesar 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Large Icon"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   3840
               TabIndex        =   7
               Top             =   30
               Width           =   1575
            End
            Begin VB.OptionButton OptList 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "List"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Index           =   0
               Left            =   1320
               TabIndex        =   6
               Top             =   30
               Width           =   735
            End
            Begin VB.OptionButton OptDetail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Details"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   120
               TabIndex        =   5
               Top             =   30
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton OptIconKecil 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Caption         =   "Small Icon"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   2400
               TabIndex        =   4
               Top             =   30
               Width           =   1095
            End
         End
      End
      Begin Uninstaller2005.Check ChkUninstall 
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   3600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         Value           =   1
         Caption         =   "Show Uninstall Dialog !"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Show Uninstall Dialog !"
         ForeColor       =   8388608
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "FrmOptions.frx":5684
         Top             =   3240
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmOptions.frm
' Tanggal    : 8/29/2005 22:31
' Programmer : Rusman Indradi (rusman@olivault.com)
' Lokasi     : Bogor, INDONESIA
' Catatan    : Rusman Indradi ekeur stres Gw Teh euY... untuk sapa yach program ini..
'              ok deCh untuk Temen gw saudara gw yayang GW CroTZ selalu.... :)
'              tHanKz tO Rizki Priatna, Abby, Ronny, pon-pon, Maryam thaNk's for
'              yOur support Euy..... Hapy CodinG and dont forGEt me Ok....
'              unTuk mAryam And pon-pon kapan Ceng-Ceng lg euY......
'
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
Option Explicit
Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long

Private Sub chkUrut_Click()
   If chkUrut.Value = vbChecked Then
      frmmain.lstview.SortOrder = lvwAscending
    Else
      frmmain.lstview.SortOrder = lvwDescending
    End If
    frmmain.lstview.Refresh
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub LblOK_Click()
Me.Hide
End Sub


Private Sub LblOKShortCut_Click()
Dim lReturn As Long
Select Case True
  Case Option1(1).Value
    lReturn = OSfCreateShellLink("Programs" & vbNullChar, "Merry Uninstaller 2004", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Start Menu)")
  Case Option1(2).Value
    lReturn = OSfCreateShellLink("..\..\Desktop" & vbNullChar, "Merry Uninstaller 2004", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Programs)")
  Case Option1(3).Value
    lReturn = OSfCreateShellLink("..\..\Application Data\Microsoft\Internet Explorer\Quick Launch" & vbNullChar, "Merry Uninstaller 2004", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Programs)")
    Case Option1(4).Value
    lReturn = OSfCreateShellLink("StartUp" & vbNullChar, "Merry Uninstaller 2004", (App.Path & "\" & App.EXEName), "" & vbNullChar, True, "$(Programs)")
End Select
If lReturn = 0 Then
  MsgBox "Error for Create Shortcut!"
Else
  MsgBox "Success for Create Shortcut!"
End If
End Sub


Private Sub OptDetail_Click()
frmmain.lstview.view = lvwReport
End Sub

Private Sub OptIconBesar_Click()
frmmain.lstview.view = lvwIcon
End Sub

Private Sub OptIconKecil_Click()
frmmain.lstview.view = lvwSmallIcon
End Sub

Private Sub OptList_Click(Index As Integer)
frmmain.lstview.view = lvwList
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub


