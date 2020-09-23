VERSION 5.00
Begin VB.Form frmPassword 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - Set Password"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   ControlBox      =   0   'False
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3795
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton cmdclose 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
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
   Begin Uninstaller2005.XPButton cmdok 
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Frame fraSetPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      Begin VB.TextBox txtPassword2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtPassword1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Re-Type Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin Uninstaller2005.Check chkPassword 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      Caption         =   " Enable Password Protection."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12582912
      Caption         =   " Enable Password Protection."
      ForeColor       =   12582912
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : frmPassword.frm
' Tanggal    : 8/29/2005 22:30
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

Dim encrypt As New clsEncryption
Private Sub chkPassword_Click()
    On Error Resume Next
    If chkPassword.Value = Checked Then
      fraSetPassword.Enabled = True
        txtPassword1.text = encrypt.Cryption(modregistry2.GetSetting("", "Desktop", "Kawin ama Merry", "", HKEY_CURRENT_USER, "Control Panel"), "Merry Uninstaller 2004", False)
        txtPassword2.text = encrypt.Cryption(modregistry2.GetSetting("", "Desktop", "Kawin ama Merry", "", HKEY_CURRENT_USER, "Control Panel"), "Merry Uninstaller 2004", False)
        txtPassword1.SetFocus
    Else
      fraSetPassword.Enabled = False
        txtPassword1.text = ""
        txtPassword2.text = ""
        cmdok.SetFocus
    End If
End Sub

Private Sub cmdclose_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    If chkPassword.Value = Checked Then
        If txtPassword1.text <> txtPassword2.text Then
            MsgBox "Invalid password combination... please try again...", vbCritical, App.Title
            txtPassword1.text = "": txtPassword2.text = ""
            txtPassword1.SetFocus
            Exit Sub
        ElseIf Len(txtPassword1.text) < 6 Then
            MsgBox "A valid password must be 6 characters long...", vbCritical, App.Title
            txtPassword1.text = "": txtPassword2.text = ""
            txtPassword1.SetFocus
            Exit Sub
        Else
            modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Control Panel\Desktop", "Merry Uninstaller 2004", chkPassword.Value
            modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "kawin ama Merry", encrypt.Cryption(txtPassword1.text, "Merry Uninstaller 2004", True)
            MsgBox "Administrator password successfully applied...", vbInformation, App.Title
            Unload Me
        End If
    Else
        modRegistry1.SetDWORDValue "HKEY_CURRENT_USER\Control Panel\Desktop", "Merry Uninstaller 2004", chkPassword.Value
        modRegistry1.SetStringValue "HKEY_CURRENT_USER\Control Panel\Desktop", "Kawin ama Merry", ""
        MsgBox "Administrator password successfully removed...", vbInformation, App.Title
        Unload Me
    End If

End Sub

Private Sub Form_Load()
    On Error Resume Next
     
        
    chkPassword.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Merry Uninstaller 2004")
End Sub

