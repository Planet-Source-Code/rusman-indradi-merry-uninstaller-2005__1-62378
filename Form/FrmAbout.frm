VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - About"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton cmdclose 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      caption         =   "&Close"
      font            =   "FrmAbout.frx":0000
      focusrect       =   0   'False
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "FrmAbout.frx":0034
      Top             =   2280
      Width           =   5175
   End
   Begin Uninstaller2005.XPFrame XPFrame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _extentx        =   9551
      _extenty        =   3836
      font            =   "FrmAbout.frx":0199
      backcolor       =   16777215
      fontname        =   "MS Sans Serif"
      fontsize        =   8.25
      fontbold        =   0   'False
      fontitalic      =   0   'False
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please Vote My Program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   3240
         TabIndex        =   10
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Merry Uninstaller 2005"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2001 - 2005"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Olivault™ Software"
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
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "All Rights Reserved"
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
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "If you find any bugs on my program, please send e-mail report to our technical support.thank's"
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
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label lblEmailSaya 
         BackStyle       =   0  'Transparent
         Caption         =   "intouch@olivault.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   240
         MouseIcon       =   "FrmAbout.frx":01C5
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblWebSiteSaya 
         BackStyle       =   0  'Transparent
         Caption         =   "http://olivault.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         MouseIcon       =   "FrmAbout.frx":0317
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1440
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmAbout.frm
' Tanggal    : 8/29/2005 22:33
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

Private Declare Function ShellExecute Lib _
        "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, _
        ByVal lpOperation As String, _
        ByVal lpFile As String, _
        ByVal lpParameters As String, _
        ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1

Private Sub cmdclose_Click()

    Unload Me

End Sub

Private Sub Form_Click()

    Unload Me

End Sub

Private Sub Label1_Click()

    Unload Me

End Sub

Private Sub lblEmailSaya_Click()

    ShellExecute Me.hWnd, _
                 vbNullString, _
                 "mailto:intouch@olivault.com?subject=Merry Uninstaller 2005", _
                 vbNullString, _
                 "c:\", _
                 SW_SHOWNORMAL

End Sub

Private Sub lblWebSiteSaya_Click()

    ShellExecute Me.hWnd, _
                 vbNullString, _
                 "http://www.olivault.com", _
                 vbNullString, _
                 "c:\", _
                 SW_SHOWNORMAL

End Sub

