VERSION 5.00
Begin VB.Form FrmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5430
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture41 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   1800
      Picture         =   "FrmSplash.frx":5E62
      ScaleHeight     =   975
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   3250
      Width           =   3495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1680
      Picture         =   "FrmSplash.frx":75EB
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   0
      Picture         =   "FrmSplash.frx":912D
      ScaleHeight     =   4335
      ScaleWidth      =   1575
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   0
      Top             =   3240
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
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
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
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Author: Rusman Indradi"
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
      Index           =   6
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   2295
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
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label lblAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmSplash.frx":E7B1
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
      Height          =   1575
      Index           =   9
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Merry Uninstaller 2005"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmSplash.frm
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
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long ' API FOR UPGRADING CONTROLS
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Const FILE_ATTRIBUTE_HIDDEN = &H2 '                     API FOR SETTING THE MANIFEST AS HIDDEN



Private Sub tmrTimer_Timer()
 tmrTimer.Enabled = False

 'Periksa apakah user memakai password login apa tidak
If modRegistry1.GetDWORDValue("HKEY_CURRENT_USER\Control Panel\Desktop", "Mery Uninstaller 2003") = 1 Then
Unload Me: frmLogin.Show
Else
Unload Me: frmmain.Show
End If

End Sub
Private Sub Form_Initialize() 'BEFORE THE USER SEES FORM
Dim xptheme As Long
Dim manifestpth As String 'DIM THE VARIBLES ETC

On Error GoTo manifestdoesnotexisT 'IF NO MANIFEST THEME FILE HAS BEEN MADE YET

If Right(App.Path, 1) = "\" Then                                 '|
    manifestpth = App.Path & App.EXEName & ".exe.manifest"       '|
Else                                                             '|
    manifestpth = App.Path & "\" & App.EXEName & ".exe.manifest" '|  FIND OUT IF MANIFEST ALREADY EXISTS
End If                                                           '|

FileCopy manifestpth, "c:\checkexist.txt"
Kill "c:\checkexist.txt"
xptheme = InitCommonControls                        ' IF MANIFEST EXISTS, EXUCUTE CONTROL UPGRADE TO XP THEME STYLE
Exit Sub

manifestdoesnotexisT:
Call makeNEWmanifest   ' IF MANIFEST DOES NOT EXIST, AND ERROR OCURRS, GO AND MAKE A NEW ONE
End Sub

Sub makeNEWmanifest()

Dim NEWmanifestpth As String
Dim xptheme As Long             ' SET VARIBLES ETC...
Dim setAShidden As Long

On Error GoTo problemARGH ' ERROR HANDLING, GOTO PROBLEMARGH ON ERROR EVENT

If Right(App.Path, 1) = "\" Then                                        '|
    NEWmanifestpth = App.Path & App.EXEName & ".exe.manifest"           '|
Else                                                                    '| SET PATH OF MANIFEST THEME FILE
    NEWmanifestpth = App.Path & "\" & App.EXEName & ".exe.manifest"     '|
End If                                                                  '|

Open NEWmanifestpth For Output As #1  '     WRITE THE MANIFEST FILE BECAUSE IT DOES NOT YET EXIST.
Print #1, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?><assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & "><assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " name=" & Chr(34) & "HybridDesign.WindowsXP.Example" & Chr(34) & " type=" & Chr(34) & "win32" & Chr(34) & " /> <description>An example of windows XP theming.</description> <dependency> <dependentAssembly> <assemblyIdentity type=" & Chr(34) & "win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " /> </dependentAssembly> </dependency> </assembly>" ' CONTENTS OF THE MANIFEST FILE...
Close #1 '                                  YOU NEED TO HAVE THIS FILE, OR THE THEME WILL NOT WORK!

xptheme = InitCommonControls                        ' IF MANIFEST EXISTS, EXUCUTE CONTROL UPGRADE TO XP THEME STYLE

setAShidden = SetFileAttributes(NEWmanifestpth, FILE_ATTRIBUTE_HIDDEN) ' HIDE THE MANIFEST THEME FILE


Exit Sub ' SKIP ANYTHING AFTER THIS MARK IN CURRENT SUB

problemARGH: ' IF AN ERROR OCCURED DURING THE CREATION OF THE MANIFEST
MsgBox "Error creating Windows XP theme file. You may be running EXE file from a network drive with which you dont have write permissions. Themes will not be enabled.", vbExclamation, "Themeing Error!" ' TELLING USER THAT THEMES WILL NOT BE ENABLED
End Sub

