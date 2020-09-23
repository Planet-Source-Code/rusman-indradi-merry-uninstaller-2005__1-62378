VERSION 5.00
Begin VB.Form frmBantuan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FAQ Merry Uninstaller 2005"
   ClientHeight    =   7155
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   7425
   Icon            =   "frmBantuan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin Uninstaller2005.XPButton CmdExit 
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   6720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4155
      ItemData        =   "frmBantuan.frx":5E62
      Left            =   0
      List            =   "frmBantuan.frx":5FC2
      TabIndex        =   1
      Top             =   2400
      Width           =   7425
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2370
      ItemData        =   "frmBantuan.frx":79E5
      Left            =   0
      List            =   "frmBantuan.frx":7A22
      TabIndex        =   0
      Top             =   0
      Width           =   7425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pertanyaan Yang Sering Tanya"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   6720
      Width           =   3495
   End
   Begin VB.Image ImBantu 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmBantuan.frx":7D11
      Picture         =   "frmBantuan.frx":801B
      ToolTipText     =   "Bantuan"
      Top             =   6600
      Width           =   480
   End
End
Attribute VB_Name = "frmBantuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : frmBantuan.frm
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
DefLng A-W
DefSng X-Z

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub List1_Click()
Dim i, context$, res&
i = List1.ListIndex
context$ = List1.list(i) & Chr$(0)
If context$ <> "" Then
   res& = SendMessageLong(List2.hWnd, LB_FINDSTRINGEXACT, -1&, ByVal context$)
   List2.ListIndex = res
   If List2.ListIndex > 0 Then
      List2.TopIndex = List2.ListIndex - 1
   End If
End If

End Sub

