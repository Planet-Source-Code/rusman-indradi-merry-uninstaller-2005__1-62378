VERSION 5.00
Begin VB.Form frmUninstall 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - Uninstal Program"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmUninstall.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton LblBatal 
      Height          =   375
      Left            =   2880
      TabIndex        =   8
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
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
   Begin Uninstaller2005.XPButton LblUninstall 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Uninstall"
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
   Begin Uninstaller2005.XPButton LblBersihkan 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Clear"
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
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
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
      Begin VB.PictureBox image1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   90
         Picture         =   "FrmUninstall.frx":000C
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LblInfomasi2 
         BackStyle       =   0  'Transparent
         Caption         =   "Are you sure to uninstall this program ?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label LblInfomasi 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmUninstall.frx":0C4E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin VB.TextBox TxtRegname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox TxtDname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "frmUninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmUninstall.frm
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
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub

Private Sub LblBatal_Click()
Unload Me
End Sub


Private Sub LblBersihkan_Click()
Call DeleteKey(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall\" + TxtRegname.text)
Unload Me
frmmain.New_Refresh
End Sub


Private Sub LblUninstall_Click()
LblInfomasi2.Visible = False
TxtDname.Visible = False
image1.Visible = True
LblUninstall.Visible = False
LblInfomasi.Visible = True
LblBersihkan.Visible = True
frmmain.Get_Uninstall
End Sub


