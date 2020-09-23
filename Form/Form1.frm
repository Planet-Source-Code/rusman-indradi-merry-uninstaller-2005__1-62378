VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller Add/Remove Plus! 2005 v.1.0"
   ClientHeight    =   5340
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9090
   StartUpPosition =   2  'CenterScreen
   Begin Uninstaller2005.XPButton cmdtweak 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Tweak Uninstaller"
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
   Begin Uninstaller2005.XPButton cmdoptions 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Options"
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
   Begin Uninstaller2005.XPButton cmdreport 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Report"
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
   Begin Uninstaller2005.XPButton cmdinfo 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Information Details"
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
   Begin Uninstaller2005.XPButton cmddelete 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Delete from List"
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
   Begin Uninstaller2005.XPButton cmduninstall 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Caption         =   "Uninstall..."
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Uninstaller2005.XPFrame XPFrame2 
      Height          =   4815
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin MSComctlLib.ListView lstview 
         Height          =   4455
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Select program on the list and then double click list view for run Uninstall program"
         Top             =   240
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   7858
         View            =   3
         Arrange         =   2
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Display Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Uninstall String"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Registry Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Publisher"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Display Version"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Help Link"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "URL Info About "
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Contact"
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin Uninstaller2005.XPFrame XPFrame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   8493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      FontBold        =   0   'False
      FontItalic      =   0   'False
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":1077
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label LblStatus 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label LblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Found :"
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
         Left            =   720
         TabIndex        =   3
         Top             =   4080
         Width           =   615
      End
   End
   Begin VB.Label roesmaryam1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Merry Uninstaller 2005. Copyright © 2001 - 2005 Olivault ™ Software."
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
      Height          =   195
      Left            =   3000
      TabIndex        =   12
      Top             =   4920
      Width           =   5085
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu backup 
         Caption         =   "Backup uninstal registry"
      End
      Begin VB.Menu Report 
         Caption         =   "Report uninstaller"
      End
      Begin VB.Menu we 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Entry 
         Caption         =   "Edit Entry"
      End
      Begin VB.Menu delete 
         Caption         =   "Delete Entry"
      End
      Begin VB.Menu we1 
         Caption         =   "-"
      End
      Begin VB.Menu new 
         Caption         =   "New Entry"
      End
   End
   Begin VB.Menu view 
      Caption         =   "View"
      Begin VB.Menu Detailsview 
         Caption         =   "Details"
      End
      Begin VB.Menu list 
         Caption         =   "List"
      End
      Begin VB.Menu small 
         Caption         =   "Small Icon"
      End
      Begin VB.Menu large 
         Caption         =   "Large Icon"
      End
      Begin VB.Menu we3 
         Caption         =   "-"
      End
      Begin VB.Menu Minimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu Restore 
         Caption         =   "Restore"
      End
   End
   Begin VB.Menu Uninstall 
      Caption         =   "Uninstall"
      Begin VB.Menu Uninstall1 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu we6 
         Caption         =   "-"
      End
      Begin VB.Menu Information 
         Caption         =   "Information Details"
      End
   End
   Begin VB.Menu Tool 
      Caption         =   "Tool"
      Begin VB.Menu language 
         Caption         =   "language"
         Begin VB.Menu bhsIndonesia 
            Caption         =   "Indonesia"
         End
      End
      Begin VB.Menu mig 
         Caption         =   "-"
      End
      Begin VB.Menu Options 
         Caption         =   "Options"
      End
      Begin VB.Menu we9 
         Caption         =   "-"
      End
      Begin VB.Menu tweak 
         Caption         =   "Tweak Uninstaller"
      End
      Begin VB.Menu we8 
         Caption         =   "-"
      End
      Begin VB.Menu request 
         Caption         =   "Request Uninstaller"
      End
      Begin VB.Menu we7 
         Caption         =   "-"
      End
      Begin VB.Menu Password 
         Caption         =   "Password"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu faq 
         Caption         =   "Faq"
         Begin VB.Menu faqina 
            Caption         =   "FAQ Indonesia"
         End
         Begin VB.Menu faqeng 
            Caption         =   "FAQ English"
         End
      End
      Begin VB.Menu gelo 
         Caption         =   "-"
      End
      Begin VB.Menu Readme 
         Caption         =   "Readme"
         Begin VB.Menu Readmeindo 
            Caption         =   "Readme Indonesia"
         End
         Begin VB.Menu Readmeeng 
            Caption         =   "Readme English"
         End
      End
      Begin VB.Menu faq2 
         Caption         =   "-"
      End
      Begin VB.Menu HelpTopics 
         Caption         =   "Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu we5 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : Form1.frm
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
Public FormFlag As Boolean, fx As Long, FY As Long
Public FormFirst As Boolean, AX As Long, AY As Long
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Dim GetlocString, RegName, UString, Dname, Publisher, DVersion, HelpLink, UIAbout, Contact As String
Dim iKetetapan As Integer
Dim fTimer

Sub GetInformasi()
RegName = lstview.SelectedItem.KEY
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "DisplayName")
UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "UninstallString")
Publisher = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "Publisher"))
DVersion = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "DisplayVersion"))
HelpLink = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "HelpLink"))
UIAbout = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "URLInfoAbout"))
Contact = Trim(GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "Contact"))

If Len(RegName) = 0 Then
frmInformasi.TxtRName = "(Not available)"
Else
frmInformasi.TxtRName = RegName
End If

If Len(Dname) = 0 Then
frmInformasi.TxtDname.text = "(Not available)"
Else
frmInformasi.TxtDname.text = Dname
End If

If Len(UString) = 0 Then
frmInformasi.TxtUString.text = "(Not available)"
Else
frmInformasi.TxtUString.text = UString
End If

If Len(Publisher) = 0 Then
frmInformasi.TxtPublisher.text = "(Not available)"
Else
frmInformasi.TxtPublisher.text = Publisher
End If

If Len(DVersion) = 0 Then
frmInformasi.TxtDVersion.text = "(Not available)"
Else
frmInformasi.TxtDVersion.text = DVersion
End If

If Len(HelpLink) = 0 Then
frmInformasi.TxtHLink.text = "(Not available)"
Else
frmInformasi.TxtHLink.text = HelpLink
End If

If Len(UIAbout) = 0 Then
frmInformasi.TxtUIAbout.text = "(Not available)"
Else
frmInformasi.TxtUIAbout.text = UIAbout
End If

If Len(Contact) = 0 Then
frmInformasi.TxtContact.text = "(Not available)"
Else
frmInformasi.TxtContact.text = Contact
End If
    
End Sub

Private Sub GetKetReg()
GetlocString = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
ModRegistry.GetKeyNames HKEY_LOCAL_MACHINE, GetlocString
End Sub

Private Sub ShowUninstallList()
On Error Resume Next
Dim LokasiItem As ListItem
Call GetKetReg
For iKetetapan = 1 To sKeys.Count - 0
    Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "DisplayName")
    UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "UninstallString")
    Publisher = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "Publisher")
    DVersion = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "DisplayVersion")
    HelpLink = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "HelpLink")
    UIAbout = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "URLInfoAbout")
    Contact = GetString(HKEY_LOCAL_MACHINE, GetlocString & sKeys(iKetetapan), "Contact")

    If Len(Dname) > 0 Then
        Set LokasiItem = lstview.ListItems.Add(, sKeys(iKetetapan), Dname, 1, 1)
        
        If Len(UString) = 0 Then
           LokasiItem.SubItems(1) = "(Not available)"
        Else
           LokasiItem.SubItems(1) = UString
        End If
            
        If Len(sKeys(iKetetapan)) = 0 Then
           LokasiItem.SubItems(2) = "(Not available)"
        Else
           LokasiItem.SubItems(2) = sKeys(iKetetapan)
        End If
           
        If Len(Publisher) = 0 Then
           LokasiItem.SubItems(3) = "(Not available)"
        Else
           LokasiItem.SubItems(3) = Publisher
        End If
           
        If Len(DVersion) = 0 Then
           LokasiItem.SubItems(4) = "(Not available)"
        Else
           LokasiItem.SubItems(4) = DVersion
        End If
        
        If Len(HelpLink) = 0 Then
           LokasiItem.SubItems(5) = "(Not available)"
        Else
           LokasiItem.SubItems(5) = HelpLink
        End If
        
        If Len(UIAbout) = 0 Then
           LokasiItem.SubItems(6) = "(Not available)"
        Else
           LokasiItem.SubItems(6) = UIAbout
        End If
        
        If Len(Contact) = 0 Then
           LokasiItem.SubItems(7) = "(Not available)"
        Else
           LokasiItem.SubItems(7) = Contact
        End If
End If
    LblStatus.Caption = lstview.ListItems.Count & " Installed Programs"
Next iKetetapan
    Set sKeys = Nothing
   
End Sub

Sub Show_FormUninstall()
If FrmOption.ChkUninstall.Value = vbChecked Then
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "DisplayName")
RegName = lstview.SelectedItem.KEY
frmUninstall.TxtDname.text = Dname
frmUninstall.TxtRegname.text = RegName
frmUninstall.Show vbModal, frmmain
Else
Get_Uninstall
End If
End Sub

Sub Get_Uninstall()
Dim strRemove As String
strRemove = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "UninstallString")
WinExec strRemove, 1
End Sub



Private Sub About_Click()
FrmAbout.Show vbModal, frmmain
End Sub

Private Sub backup_Click()
frmmain.Backup_Registry
FrmBackupRegistry.Show vbModal, frmmain
End Sub

Private Sub bhsIndonesia_Click()
On Error Resume Next
'Call Applybahasaindonesia
MsgBox "Data Tidak Dapat Di Temukan, File Indonesia.ini Hilang tolong install ulang Merry Uninstaller!.", vbInformation, App.Title

End Sub

Private Sub cmddelete_Click()
On Error Resume Next
MnuHapusEntry_Click
End Sub
Private Sub cmddelete_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Delete the selected program from the list of programs installed on your computer."
    On Error GoTo 0

End Sub

Private Sub cmdinfo_Click()
frmInformasi.Show vbModal, frmmain
End Sub
Private Sub cmdinfo_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Display information about the selected program."
    On Error GoTo 0

End Sub

Private Sub cmdoptions_Click()
FrmOption.Show vbModal, frmmain
End Sub
Private Sub cmdoptions_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Display Options that's may you need to Customize Mery Uninstaller 2003."
    On Error GoTo 0

End Sub

Private Sub cmdreport_Click()
FrmPilihLaporan.Show vbModal, frmmain
End Sub
Private Sub cmdreport_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Display a report programs you have uninstalled from your computer."
    On Error GoTo 0

End Sub

Private Sub cmdtweak_Click()
On Error Resume Next
frmtweak.Show vbModal, Me
End Sub
Private Sub cmdtweak_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Tweaks Add or Remove Programs for customize security."
    On Error GoTo 0

End Sub

Private Sub cmduninstall_Click()
frmmain.Show_FormUninstall
End Sub
Private Sub cmduninstall_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Uninstalls or modifies the installed component of the selected program."
    On Error GoTo 0

End Sub

Private Sub delete_Click()
frmmain.MnuHapusEntry_Click
End Sub

Private Sub Detailsview_Click()
frmmain.lstview.view = lvwReport
End Sub

Private Sub Entry_Click()
frmmain.MnuEditEntry_Click
End Sub

Private Sub exit_Click()
Unload Me
End Sub



Private Sub faqeng_Click()
On Error Resume Next
frmBantuaneng.Show vbModal, Me
End Sub

Private Sub faqina_Click()
On Error Resume Next
frmBantuan.Show vbModal, Me
End Sub

Private Sub Form_Load()
On Error Resume Next
    Set sKeys = New Collection
    lstview.Refresh
    ShowUninstallList
    lstview.view = lvwReport
   
    'Buat Startup Key pertama
      ' create key at startup...
    modRegistry1.CreateKey "HKEY_LOCAL_MACHINE\Software\Olivault Software\Merry Uninstaller 2004"
    modRegistry1.SetDWORDValue "HKEY_LOCAL_MACHINE\Software\Olivault Software\Merry Uninstaller 2004", "FirstRun", 0
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Olivault Software\Merry Uninstaller 2004", "Author", "Rusman Indradi"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Olivault Software\Merry Uninstaller 2004", "Email", "intouch@olivault.com"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Olivault Software\Merry Uninstaller 2004", "Company", "Olivault Software"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Olivault Softwar\Merry Uninstaller 2004", "Url", "http://software.olivault.com"
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Olivault Software\Merry Uninstaller 2004", "Product", App.Title
    modRegistry1.SetStringValue "HKEY_LOCAL_MACHINE\Software\Olivault Software\Merry Uninstaller 2004", "Version", App.Major & "." & App.Minor & "." & App.Revision

'Periksa tweak untuk Add or Remove Programs
'frmtweak.Check1.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddRemovePrograms")
'frmtweak.Check2.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoRemovePage")
'frmtweak.Check3.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddPage")
'frmtweak.Check4.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoWindowsSetupPage")
'frmtweak.Check5.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromCDorFloppy")
'frmtweak.Check6.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromInternet")
'frmtweak.Check7.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromNetwork")
'frmtweak.Check8.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoSupportInfo")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    GetlocString = ""
    Dim Form As Form
    For Each Form In Forms
        If Form.Name <> Me.Name Then
            Unload Form
            Set Form = Nothing
        End If
    Next Form
End Sub


Private Sub Image1_Click()
Call Backup_Registry
FrmBackupRegistry.Show vbModal, frmmain
End Sub



Private Sub ImBantu_Click()
frmBantuan.Show
End Sub



Private Sub ImgAbout_Click()
FrmAbout.Show vbModal, frmmain
End Sub

Private Sub ImgDetailInfo_Click()
frmInformasi.Show vbModal, frmmain
End Sub
Private Sub ImgKeluar_Click()
Unload Me
End Sub

Private Sub ImgLaporan_Click()
FrmPilihLaporan.Show vbModal, frmmain
End Sub
Private Sub ImgSetings_Click()
FrmOption.Show vbModal, frmmain
End Sub
Private Sub ImgUninstall_Click()
Call Show_FormUninstall
End Sub

Sub Backup_Registry()
On Error Resume Next
Dim fName As String
fName = App.Path & "\" & "temp" & ".tmp"
SaveKey "HKEY_LOCAL_MACHINE" & "\" & GetlocString & lstview.SelectedItem.KEY, fName
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "DisplayName")
FrmBackupRegistry.LblInformasi = ">> Backup For : " & Dname
End Sub
Private Sub Lblbantu_Click()
PopupMenu FrmPopup.Menu8
End Sub
Private Sub LblEdit_Click()
PopupMenu FrmPopup.Menu6
End Sub
Private Sub LblFile_Click()
PopupMenu FrmPopup.Menu3
End Sub
Private Sub LblMaximized_Click()
Me.WindowState = vbMaximized
End Sub

Private Sub LblMinimized_Click()
Me.WindowState = vbNormal
End Sub

Private Sub LblRestore_Click()

Me.WindowState = vbMinimized
End Sub

Sub MnuHapusEntry_Click()
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "DisplayName")
RegName = lstview.SelectedItem.KEY
FrmHapusEntry.LblDname.Caption = Dname
FrmHapusEntry.TxtRegname.text = RegName
FrmHapusEntry.Show vbModal, frmmain
End Sub

Sub MnuEditEntry_Click()
Dname = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "DisplayName")
UString = GetString(HKEY_LOCAL_MACHINE, GetlocString & lstview.SelectedItem.KEY, "UninstallString")
RegName = lstview.SelectedItem.KEY
FrmEditEntry.LblRegName.Caption = RegName
FrmEditEntry.TxtDname.text = Dname
FrmEditEntry.TxtUString.text = UString
FrmEditEntry.Show vbModal, frmmain
End Sub

Private Sub LblTampilan_Click()
PopupMenu FrmPopup.Menu7
End Sub



Private Sub LblTools_Click()
PopupMenu FrmPopup.Menu4
End Sub


Private Sub LblUninstall_Click()
PopupMenu FrmPopup.Menu5
End Sub



Private Sub HelpTopics_Click()
On Error Resume Next
HyperJump App.Path & "\merry.chm"

End Sub

Private Sub Information_Click()
frmInformasi.Show vbModal, frmmain
End Sub

Private Sub large_Click()
frmmain.lstview.view = lvwSmallIcon
End Sub

Private Sub list_Click()
frmmain.lstview.view = lvwList
End Sub

Private Sub lstview_DblClick()
Call Show_FormUninstall
End Sub
Private Sub lstview_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Merry Uninstaller 2005. Copyright © 2001 - 2005 Olivault ™ Software."
    On Error GoTo 0

End Sub

Private Sub lstview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu FrmPopup.Menu
End Sub


Private Sub MnuInformation_Click()
frmInformasi.Show vbModal, frmmain
End Sub

Sub New_Refresh()
lstview.ListItems.Clear
Set sKeys = New Collection
ShowUninstallList
lstview.Refresh
lstview.view = lvwReport

End Sub
Private Sub Minimize_Click()
WindowState = vbMinimized
End Sub

Private Sub new_Click()
FrmEntryBaru.Show vbModal, frmmain
End Sub

Private Sub Options_Click()
FrmOption.Show vbModal, frmmain
End Sub


Private Sub Password_Click()
On Error Resume Next
frmPassword.Show vbModal, Me
End Sub

Private Sub Readmeeng_Click()
On Error Resume Next
HyperJump App.Path & "\Readme_English.txt"
End Sub

Private Sub Readmeindo_Click()
On Error Resume Next
HyperJump App.Path & "\Readme_Indonesia.txt"
End Sub

Private Sub Report_Click()
FrmPilihLaporan.Show vbModal, frmmain
End Sub

Private Sub request_Click()
frmrequest.Show vbModal, Me
End Sub

Private Sub Restore_Click()
WindowState = vbNormal
End Sub

Private Sub small_Click()
lstview.view = lvwIcon
End Sub

Private Sub tweak_Click()
On Error Resume Next
frmtweak.Show vbModal, Me
End Sub

Private Sub Uninstall1_Click()
frmmain.Show_FormUninstall
End Sub

Private Sub XPFrame1_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     x As Single, _
                                     y As Single)

    On Error Resume Next
        ' show description...
        roesmaryam1.Caption = "Merry Uninstaller 2004. Copyright © 2001 - 2004 Olivault ™ Software."
    On Error GoTo 0

End Sub

Sub gettweak()
On Error Resume Next

'Periksa tweak untuk Add or Remove Programs
frmtweak.Check1.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddRemovePrograms")
frmtweak.Check2.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoRemovePage")
frmtweak.Check3.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddPage")
frmtweak.Check4.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoWindowsSetupPage")
frmtweak.Check5.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromCDorFloppy")
frmtweak.Check6.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromInternet")
frmtweak.Check7.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoAddFromNetwork")
frmtweak.Check8.Value = modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall", "NoSupportInfo")
End Sub

