VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBackupRegistry 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - Backup Registry"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton LblBatal 
      Height          =   375
      Left            =   7320
      TabIndex        =   13
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
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
      FocusRect       =   0   'False
   End
   Begin Uninstaller2005.XPButton LblOK 
      Height          =   375
      Left            =   6240
      TabIndex        =   12
      Top             =   3960
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
   Begin Uninstaller2005.XPFrame XPFrame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
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
      Begin Uninstaller2005.Check CheckBSUS 
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4020
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   450
         Caption         =   ""
         ForeColor       =   0
         Caption         =   ""
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   120
         Top             =   5400
      End
      Begin VB.TextBox TxtBackup 
         Appearance      =   0  'Flat
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
         Height          =   3255
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox TxtSimpan 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "Write file name for save"
         Top             =   3990
         Width           =   2625
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   240
         Pattern         =   "*.Reg"
         TabIndex        =   1
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   1575
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegbackup.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegbackup.frx":0292
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegbackup.frx":0524
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3375
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegbackup.frx":1176
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRegbackup.frx":1A52
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeViewBackup 
         Height          =   3255
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Please double click for Export backup to Registry"
         Top             =   600
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   5741
         _Version        =   393217
         Indentation     =   0
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList2"
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
      End
      Begin Uninstaller2005.Check chkBakup 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         Caption         =   "<< See list Backup"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "<< See list Backup"
         ForeColor       =   8388608
      End
      Begin VB.Label LblBSUS 
         BackStyle       =   0  'Transparent
         Caption         =   "Backup All Uninstall (Key)"
         Enabled         =   0   'False
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
         Left            =   600
         TabIndex        =   11
         Top             =   4035
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "File :"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2880
         TabIndex        =   8
         Top             =   4020
         Width           =   360
      End
      Begin VB.Label LblInformasi 
         BackStyle       =   0  'Transparent
         Caption         =   ">> New Backup :"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   5655
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   7320
         Picture         =   "FrmRegbackup.frx":232E
         Stretch         =   -1  'True
         Top             =   4680
         Width           =   480
      End
      Begin VB.Shape Shape3 
         Height          =   1035
         Left            =   165
         Top             =   4440
         Width           =   7980
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRegbackup.frx":2BF8
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
         Left            =   360
         TabIndex        =   6
         Top             =   4560
         Width           =   6735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Example >> ; Write new information "
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
         Left            =   360
         TabIndex        =   5
         Top             =   5160
         Width           =   6015
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3960
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmBackupRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmRegbackup.frm
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
Const SW_SHOWNORMAL = 1
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub CheckBSUS_Click()



    If CheckBSUS.Value = vbChecked Then

  Dim fName As String
        fName = App.Path & "\" & "temp" & ".tmp"
        SaveKey "HKEY_LOCAL_MACHINE" & "\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", fName
        LoadText
        LoadText
        Kill App.Path & "\temp" & ".tmp"
        LblInformasi.Caption = ">> Backup for : All Uninstall (Key) on registry"
        TreeViewBackup.Nodes.Clear
        chkBakup.Caption = "<< See list Backup"
        chkBakup.Value = vbUnchecked
        LblOK.Enabled = False
        TxtSimpan.text = ""
        LblBSUS.Enabled = True
      Else
        chkBakup.Caption = "<< See list Backup"
        chkBakup.Value = vbUnchecked
        TreeViewBackup.Nodes.Clear
        LblOK.Enabled = True
        LblBSUS.Enabled = False
        frmmain.Backup_Registry
        LoadText
        LoadText
        Kill App.Path & "\temp" & ".tmp"
    End If

  


End Sub

Private Sub chkBakup_Click()
        If chkBakup.Value = vbChecked Then
            Nodes
            LblInformasi.Caption = ">> Please double click for Export backup to Registry"
            TxtBackup.text = ""
            LblOK.Enabled = False
            chkBakup.Caption = "Back >>"
            TxtSimpan.text = ""
            CheckBSUS.Enabled = False
            CheckBSUS.Value = vbUnchecked
          Else
            CheckBSUS.Value = vbUnchecked
            CheckBSUS.Enabled = True
            chkBakup.Caption = "<< See list Backup"
            TreeViewBackup.Nodes.Clear
            TxtBackup.text = ""
            LblOK.Enabled = True
            frmmain.Backup_Registry
            LoadText
            LoadText
            Kill App.Path & "\temp" & ".tmp"
        End If

End Sub

Private Sub Form_Load()
On Error Resume Next
    MkDir App.Path & "\Backup Registry\"
    File1.Path = App.Path & "\Backup Registry\"
    LoadText
    LoadText
    Kill App.Path & "\temp" & ".tmp"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    FormDrag Me

End Sub

Sub Hapus()

  Dim result As Integer
    result = MsgBox("Are you sure remove Backup Registry : " & TreeViewBackup.SelectedItem.text & "?", vbInformation + vbYesNo, "Delete")
    If result = vbYes Then

        Kill App.Path & "\Backup Registry\" & TreeViewBackup.SelectedItem.text
        Nodes
    End If


End Sub

Sub kembalikan_registry()

    

    If Not TreeViewBackup.SelectedItem Is Nothing Then
        ShellExecute 0, vbNullString, App.Path & "\Backup Registry\" & TreeViewBackup.SelectedItem.text & ".", vbNullString, "", SW_SHOWNORMAL
      Else

        MsgBox "Nothing to open!", vbExclamation, App.Title
    End If


End Sub

Private Sub LblBatal_Click()

    Unload Me

End Sub

Private Sub LblOK_Click()

  Dim fName As String
  Dim num, ad As String
  Dim retval, result As String

    num = LCase$(TxtSimpan.text)

    fName = App.Path & "\Backup Registry\" & num & ".reg"
    ad = LCase$(TxtSimpan.text) & ".reg"
    If TxtBackup.text = "" Then
        MsgBox "List on text not found !!", vbCritical, "error !!"

    End If

    retval = Dir$(App.Path & "\Backup Registry\" & num & ".reg")

    If retval = ad Then
        result = MsgBox("File with name [ " & retval & " ] already exist." + vbCrLf & _
                 "Would you like to replace the existing file ?", vbInformation + vbYesNo, "Informatin !!")
        If result = vbYes Then
            fName = App.Path & "\Backup Registry\" & num & ".reg"
            Close #1
            Open fName For Output As #1
            Print #1, TxtBackup.text
            Close #1
            LblOK.Enabled = False
            TxtBackup.text = ""
            TxtSimpan.text = ""
            TreeViewBackup.Refresh

        End If

    End If

    Close #1
    Open fName For Output As #1
    Print #1, TxtBackup.text
    Close #1
    LblOK.Enabled = False
    TxtBackup.text = ""
    TxtSimpan.text = ""
    TreeViewBackup.Refresh

End Sub

Sub LoadText()
      Dim fName As String
      Dim readln As String
      Dim textload As String
        fName = App.Path & "\temp" & ".tmp"
        TxtBackup.text = ""
        Open App.Path & "\temp" & ".tmp" For Input As #1
        Do While Not EOF(1)
            Line Input #1, readln
            textload = textload + readln + Chr$(13) + Chr$(10)
            If Len(textload) >= 500000 Then
                MsgBox "This File {" & App.Path & "\temp" & ".tmp" & "} too large to Open", vbCritical, "Error !!"
                
            End If
        Loop
        TxtBackup = textload
        Close #1

End Sub

Private Sub Nodes()

  Dim s As Node
  Dim i As Integer



        File1.Refresh
        TreeViewBackup.Nodes.Clear
        Set s = TreeViewBackup.Nodes.Add(, , "r", "File Reg Backup", 3)
        For i = 0 To File1.ListCount - 1
            Set s = TreeViewBackup.Nodes.Add("r", tvwChild, , File1.list(i), 2, 1)
        Next i
        TreeViewBackup.Nodes(1).Expanded = True

End Sub

Sub panggil_Nodes() ':( Missing Scope

    Nodes

End Sub

Private Sub Timer1_Timer()

    If TxtSimpan.text = "" Then
        LblOK.Enabled = False
      Else
        LblOK.Enabled = True
    End If

End Sub

Private Sub TreeViewBackup_DblClick()

    kembalikan_registry

End Sub

Private Sub TreeViewBackup_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then PopupMenu FrmPopup.Menu2

End Sub

Private Sub TreeViewBackup_NodeClick(ByVal Node As MSComctlLib.Node)

  Dim fName As String
  Dim readln As String
  Dim wow As Integer



    fName = App.Path & "\Backup Registry\" & Node.text & "."

    TxtBackup.text = ""
    Open fName For Input As #1
    Do While Not EOF(1)
        Line Input #1, readln
        wow = wow + readln + Chr$(13) + Chr$(10)
        If Len(wow) >= 30000 Then
            MsgBox "This file {" & App.Path & "\temp" & ".tmp" & "} too large to Open", vbCritical, "Error !!"
            
        End If
    Loop
        Close #1

End Sub
