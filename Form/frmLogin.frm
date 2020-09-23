VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4080
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4080
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton cmdCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
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
   Begin Uninstaller2005.XPButton cmdok 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
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
   Begin Uninstaller2005.XPFrame XPFrame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1931
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
         Picture         =   "frmLogin.frx":5E62
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   720
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Administrator Password:"
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
         Height          =   255
         Left            =   720
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : frmLogin.frm
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

Private setX        As Integer
Private setY        As Integer
Dim encrypt As New clsEncryption
Dim Hitung As Integer
Private Sub cmdcancel_Click()
   On Error Resume Next
    ' close the form and terminate the program...
   End 'keluar bo
End Sub



Private Sub cmdOK_Click()
 On Error Resume Next
    If txtPassword.text = encrypt.Cryption(modregistry2.GetSetting("", "Desktop", "kawin ama Mery", "", HKEY_CURRENT_USER, "Control Panel"), "Mery Uninstaller 2003", False) Then
        frmmain.Show
        Unload Me
            Else
        MsgBox "Invalid administrator password... please try again...", vbCritical, App.Title
        txtPassword.text = ""
        txtPassword.SetFocus
        GoTo HitungKesalahan
        Exit Sub
    End If
HitungKesalahan:
Hitung = Hitung + 1   'Counter bertambah satu

If Hitung = 3 Then    'Jika Hitung = 3, maka...
'Tampilkan pesan
'Print "Sudah 3 kali kesempatan. Login ditolak!"
MsgBox "You already try login 3 times!please contact your administrator.", vbCritical, App.Title
txtPassword.Enabled = False
Unload Me 'bila user salah/tidak memasukan password dg benar sebanyak 3 maka tutup aplikasi ini :( LoLZ
End If
Exit Sub
End Sub
Private Sub lblTitle_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

    On Error Resume Next
        setX = x
        setY = y
    
End Sub
Private Sub lblTitle_MouseMove(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

    On Error Resume Next
        ' move the form...
        If Button = 1 Then
            Me.Left = Me.Left + (x - setX)
            Me.Top = Me.Top + (y - setY)
        End If


End Sub




Private Sub txtPassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
        'Ini adalah string yg diperbolehkan utk diinput
        'Jika ditekan Esc pd keyboard
        If KeyAscii = 27 Then
           cmdcancel_Click     'Langsung keluar
        End If
        'Jika ditekan Enter pd keyboard
        If KeyAscii = 13 Then
            txtPassword.SetFocus   'pindahkan kursor ke txtPassword
            cmdOK_Click 'eksekusi tombol cmdOK euy
        End If
End Sub


