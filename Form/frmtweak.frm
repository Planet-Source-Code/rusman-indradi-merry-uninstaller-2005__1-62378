VERSION 5.00
Begin VB.Form frmtweak 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - Tweaks"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin Uninstaller2005.XPButton cmdcancel 
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Cancel"
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
   Begin Uninstaller2005.XPButton cmdapply 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
      _ExtentX        =   2143
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
   Begin Uninstaller2005.XPFrame XPFrame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7011
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
      Begin Uninstaller2005.Check Check8 
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   3600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Caption         =   "Disable Support Information"
         ForeColor       =   0
         Caption         =   "Disable Support Information"
      End
      Begin Uninstaller2005.Check Check7 
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   3120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "Hide ""Add programs from your network"" option"
         ForeColor       =   0
         Caption         =   "Hide ""Add programs from your network"" option"
      End
      Begin Uninstaller2005.Check Check6 
         Height          =   375
         Left            =   4080
         TabIndex        =   7
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "Hide ""Add programs from Microsoft"" option"
         ForeColor       =   0
         Caption         =   "Hide ""Add programs from Microsoft"" option"
      End
      Begin Uninstaller2005.Check Check5 
         Height          =   375
         Left            =   4080
         TabIndex        =   6
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "Hide ""Add a program from CD-ROM or disk"" option "
         ForeColor       =   0
         Caption         =   "Hide ""Add a program from CD-ROM or disk"" option "
      End
      Begin Uninstaller2005.Check Check4 
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "Disable Windows Components Wizard"
         ForeColor       =   0
         Caption         =   "Disable Windows Components Wizard"
      End
      Begin Uninstaller2005.Check Check3 
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         Caption         =   "Disable Add Programs "
         ForeColor       =   0
         Caption         =   "Disable Add Programs "
      End
      Begin Uninstaller2005.Check Check1 
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         Caption         =   "Disable Add/Remove Programs"
         ForeColor       =   0
         Caption         =   "Disable Add/Remove Programs"
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   120
         Picture         =   "frmtweak.frx":0000
         ScaleHeight     =   2775
         ScaleWidth      =   3855
         TabIndex        =   1
         Top             =   600
         Width           =   3855
      End
      Begin Uninstaller2005.Check Check2 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         Caption         =   "Disable Change and Remove Programs"
         ForeColor       =   0
         Caption         =   "Disable Change and Remove Programs"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tweaks Add or Remove Programs...."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmtweak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : frmtweak.frm
' Tanggal    : 8/29/2005 22:29
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
Dim Changed As Boolean


Private Sub cmdapply_Click()
Changed = True
On Error Resume Next
'Tweaks NoAddRemovePrograms - Disable Add/Remove Programs
If Check1.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddRemovePrograms", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoAddRemovePrograms", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
'Tweaks NoRemovePage - Disable Change and Remove Programs
If Check2.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoRemovePage", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoRemovePage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
'Tweaks NoAddPage - Disable Add Programs
If Check3.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddPage", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoAddPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
'Tweaks NoWindowsSetupPage - Disable Windows Components Wizard
If Check4.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoWindowsSetupPage", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoWindowsSetupPage", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
'Tweaks NoAddFromCDorFloppy - Hide "Add a program from CD-ROM or disk" option
If Check5.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromCDorFloppy", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoAddFromCDorFloppy", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
'Tweaks NoAddFromInternet - Hide "Add programs from Microsoft" option
If Check6.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromInternet", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoAddFromInternet", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
'Tweaks NoAddFromNetwork - Hide "Add programs from your network" option
If Check7.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromNetwork", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoAddFromNetwork", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
'Tweaks NoSupportInfo - Disable Support Information
If Check8.Value = Checked Then
modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoSupportInfo", 1
Else
modregistry2.DeleteSetting "", "Uninstall", "NoSupportInfo", HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies"
End If
End Sub

Private Sub cmdcancel_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Load()
'Buat Key Registry untuk tweak Addremove programs
modRegistry1.CreateKey "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall"
frmmain.gettweak
End Sub
Private Sub Form_Unload(Cancel As Integer)
'sebelum aplikasi ini di tutup maka periksa
'apakah ada perubahan
'bila ada perubahan maka tampilkan mSBOX Restart Ala WinDOwS API
On Error Resume Next
If Changed = True Then
  SHRestartSystemMB Me.hWnd, vbNullString, 2 Or 4 'tampilkan restart box bila ada perubahan
End If
End Sub
