VERSION 5.00
Begin VB.Form FrmEntryBaru 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - New Entry "
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4635
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPFrame XPFrame1 
      Height          =   2895
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
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
         Height          =   495
         Left            =   120
         Picture         =   "FrmEntryBaru.frx":0000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox TxtRegName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "Key Registry Name"
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox TxtDiplayName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "Program Name"
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox TxtUninstallString 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Text            =   "Command Uninstall"
         Top             =   1680
         Width           =   4215
      End
      Begin Uninstaller2005.XPButton LblBatal 
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin Uninstaller2005.XPButton LblOK 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   2400
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "&Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FocusRect       =   0   'False
      End
      Begin VB.Label LblDiplayName 
         BackStyle       =   0  'Transparent
         Caption         =   "Display Name :"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label LblUninstallString 
         BackStyle       =   0  'Transparent
         Caption         =   "Uninstall String :"
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label LblRegName 
         BackStyle       =   0  'Transparent
         Caption         =   "Registry Name (key) :"
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
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4320
         Y1              =   2160
         Y2              =   2160
      End
   End
End
Attribute VB_Name = "FrmEntryBaru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmEntryBaru.frm
' Tanggal    : 8/29/2005 22:32
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


Private Sub LblBatal_Click()
Unload Me
End Sub



Private Sub LblOK_Click()
On Error Resume Next
If TxtRegname.text = "" Then
MsgBox "Please Don't blank to Key registry Name !", vbCritical, "Error !!"
Exit Sub
End If
If TxtDiplayName.text = "" Then
MsgBox "Please Don't blank to Program Name !", vbCritical, "Error !!"
Exit Sub
End If
If TxtUninstallString.text = "" Then
MsgBox "Please Don't blank to Uninstall Command ", vbCritical, "Error !!"
Exit Sub
End If

Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.text, "DisplayName", TxtDiplayName.text)
Call SaveString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.text, "UninstallString", TxtUninstallString.text)
frmmain.New_Refresh
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub




Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub
