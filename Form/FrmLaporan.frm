VERSION 5.00
Begin VB.Form FrmLaporan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Merry Uninstaller 2005 - Report"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin Uninstaller2005.XPButton LblTutup 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
   End
   Begin Uninstaller2005.XPButton LblBuka 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Open"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
   End
   Begin Uninstaller2005.XPButton LblSimpan 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
   End
   Begin Uninstaller2005.XPButton LblPrint 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      Width           =   8175
      _ExtentX        =   14420
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
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4610
         Left            =   120
         ScaleHeight     =   4605
         ScaleWidth      =   7965
         TabIndex        =   1
         Top             =   800
         Width           =   7970
         Begin VB.TextBox TxtLaporan 
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
            Height          =   4360
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   3
            Top             =   240
            Width           =   7935
         End
         Begin VB.PictureBox picRuler 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   0
            Picture         =   "FrmLaporan.frx":0000
            ScaleHeight     =   270
            ScaleWidth      =   7950
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   0
            Width           =   7950
         End
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   7560
         Picture         =   "FrmLaporan.frx":A1FA
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "FrmLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmLaporan.frm
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
Sub LoadText()
Dim fName As String
Dim readln
Dim textload
On Error GoTo ld_err
fName = App.Path & "\temp" & ".tmp"
TxtLaporan.text = ""
Open App.Path & "\temp" & ".tmp" For Input As #1
        Do While Not EOF(1)
            Line Input #1, readln
           textload = textload + readln + Chr$(13) + Chr$(10)
If Len(textload) >= 500000 Then
      MsgBox "This file {" & App.Path & "\temp" & ".tmp" & "} too large for open", vbCritical
      GoTo ld_end
    End If
  Loop
 TxtLaporan = textload
ld_err:
    Resume ld_end
ld_end:
    On Error Resume Next
    Close #1
End Sub

Private Sub Form_Load()
On Error Resume Next
Call LoadText
Call LoadText
Kill App.Path & "\temp" & ".tmp"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub



Private Sub LblBuka_Click()
  Dim filename As String
   filename = OpenDialog(Me, "Text Files (*.txt)|*.txt|All files (*.*)|*.*", _
                   "Open", "")
If Len(filename) Then
 On Error GoTo ld_err
Dim readln
Dim textload
TxtLaporan.text = ""
 Open filename For Input As #1
        Do While Not EOF(1)
            Line Input #1, readln
            textload = textload + readln + Chr$(13) + Chr$(10)

   If Len(textload) >= 30000 Then
      MsgBox "This file {" & App.Path & "\temp" & ".tmp" & "} too large for open", vbCritical, "Error !!"
      GoTo ld_end
    End If
  Loop
 TxtLaporan = textload
ld_err:
 Resume ld_end
ld_end:
On Error Resume Next
  Close #1
  End If
End Sub


Private Sub LblPrint_Click()
On Error GoTo perbaiki:
Printer.Print ""
Printer.FontName = "Arial"
Printer.FontSize = 8
Printer.FontBold = False
Printer.Print Now
Printer.Print ""
Printer.Print Me.TxtLaporan.text
Printer.EndDoc
Exit Sub
perbaiki:
MsgBox "Error : " & Err.Description & ", Please Check your printer ?", vbCritical, "Print Error !!"
End Sub



Private Sub LblSimpan_Click()
  Dim filename As String
  On Local Error Resume Next
  filename = SaveDialog(Me, "Text Files (*.txt)|*.txt", _
                       "Save", "", "")
If Len(filename) Then
     Close #1
Open filename For Output As #1
Print #1, TxtLaporan.text
Close #1
  End If
End Sub



Private Sub LblTutup_Click()
Unload Me
End Sub



Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
FormDrag Me
End Sub


