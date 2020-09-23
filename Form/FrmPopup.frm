VERSION 5.00
Begin VB.Form FrmPopup 
   BorderStyle     =   0  'None
   ClientHeight    =   1260
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu MnuUninstall 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu MnuInformasi 
         Caption         =   "Information Details"
      End
      Begin VB.Menu MnuLaporan 
         Caption         =   "Report Information"
      End
      Begin VB.Menu spr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegEntry 
         Caption         =   "Registry Entry"
         Begin VB.Menu MnuEditEntry 
            Caption         =   "Edit Entry"
         End
         Begin VB.Menu MnuHEntry 
            Caption         =   "Delete Entry"
         End
         Begin VB.Menu spr2 
            Caption         =   "-"
         End
         Begin VB.Menu MnuEBaru 
            Caption         =   "New Entry"
         End
      End
      Begin VB.Menu spr23 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBakup 
         Caption         =   "Backup Uninstall Reg ..."
      End
      Begin VB.Menu spr4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Begin VB.Menu MnuEBR 
         Caption         =   "Export Backup Registry"
      End
      Begin VB.Menu MnuHapusFile 
         Caption         =   "Delete File Backup"
      End
      Begin VB.Menu spr5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuRefresh2 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "Menu3"
      Begin VB.Menu mnuBUR 
         Caption         =   "Backup Uninstall Reg ..."
      End
      Begin VB.Menu MnuLaporan2 
         Caption         =   "Report Information"
      End
      Begin VB.Menu spr6 
         Caption         =   "-"
      End
      Begin VB.Menu Keluar 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   "Menu4"
      Begin VB.Menu mnuPilihan 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu Menu5 
      Caption         =   "Menu5"
      Begin VB.Menu MnuUninstall2 
         Caption         =   "Uninstall"
      End
      Begin VB.Menu spr7 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDinformasi 
         Caption         =   "Information Detail"
      End
   End
   Begin VB.Menu Menu6 
      Caption         =   "Menu6"
      Begin VB.Menu MnuEEntry2 
         Caption         =   "Edit Entry"
      End
      Begin VB.Menu MnuHEntry2 
         Caption         =   "Delete Entry"
      End
      Begin VB.Menu spr9 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEBaru2 
         Caption         =   "New Entry"
      End
   End
   Begin VB.Menu Menu7 
      Caption         =   "Menu7"
      Begin VB.Menu MnuDetail 
         Caption         =   "Details"
      End
      Begin VB.Menu MnuList 
         Caption         =   "List"
      End
      Begin VB.Menu mnuIconBesar 
         Caption         =   "Small Icon"
      End
      Begin VB.Menu MnuIconKecil 
         Caption         =   "Large Icon "
      End
      Begin VB.Menu spr10 
         Caption         =   "-"
      End
      Begin VB.Menu MnuMinimalWindow 
         Caption         =   "Minimize"
      End
      Begin VB.Menu MnuKWindow 
         Caption         =   "Restore"
      End
      Begin VB.Menu MnuMaksWindow 
         Caption         =   "Maximize"
      End
   End
   Begin VB.Menu Menu8 
      Caption         =   "Menu8"
      Begin VB.Menu MnuBantuan 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu spr11 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Mery"
      End
   End
End
Attribute VB_Name = "FrmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Nama File  : FrmPopup.frm
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
Private Sub Keluar_Click()
Unload frmmain
End Sub

Private Sub mnuAbout_Click()
FrmAbout.Show vbModal, frmmain
End Sub

Private Sub MnuBakup_Click()
frmmain.Backup_Registry
FrmBackupRegistry.Show vbModal, frmmain
End Sub

Private Sub MnuBantuan_Click()
frmBantuan.Show
End Sub

Private Sub mnuBUR_Click()
frmmain.Backup_Registry
FrmBackupRegistry.Show vbModal, frmmain
End Sub

Private Sub MnuDetail_Click()
frmmain.lstview.view = lvwReport
End Sub


Private Sub MnuDinformasi_Click()
frmInformasi.Show vbModal, frmmain
End Sub

Private Sub MnuEBaru_Click()
FrmEntryBaru.Show vbModal, frmmain
End Sub

Private Sub MnuEBaru2_Click()
FrmEntryBaru.Show vbModal, frmmain
End Sub

Private Sub MnuEBR_Click()
FrmBackupRegistry.kembalikan_registry
End Sub

Private Sub MnuEditEntry_Click()
frmmain.MnuEditEntry_Click
End Sub

Private Sub MnuEEntry2_Click()
frmmain.MnuEditEntry_Click
End Sub


Private Sub MnuHapusFile_Click()
FrmBackupRegistry.Hapus
End Sub

Private Sub MnuHEntry_Click()
frmmain.MnuHapusEntry_Click
End Sub

Private Sub MnuHEntry2_Click()
frmmain.MnuHapusEntry_Click
End Sub

Private Sub mnuIconBesar_Click()
frmmain.lstview.view = lvwIcon
End Sub

Private Sub MnuIconKecil_Click()
frmmain.lstview.view = lvwSmallIcon
End Sub

Private Sub MnuInformasi_Click()
frmInformasi.Show vbModal, frmmain
End Sub

Private Sub MnuKWindow_Click()
frmmain.WindowState = vbNormal
End Sub

Private Sub MnuLaporan_Click()
FrmPilihLaporan.Show vbModal, frmmain
End Sub

Private Sub MnuLaporan2_Click()
FrmPilihLaporan.Show vbModal, frmmain
End Sub

Private Sub MnuList_Click()
frmmain.lstview.view = lvwList
End Sub

Private Sub MnuMaksWindow_Click()
frmmain.WindowState = vbMaximized
End Sub

Private Sub MnuMinimalWindow_Click()
frmmain.WindowState = vbMinimized
End Sub

Private Sub MnuPG_Click()
On Error Resume Next

Shell ("explorer C:\program files\"), vbNormalFocus
End Sub

Private Sub MnuPG2_Click()
Shell ("explorer C:\program files\"), vbNormalFocus
End Sub

Private Sub mnuPilihan_Click()
FrmOption.Show vbModal, frmmain
End Sub

Private Sub MnuRefresh_Click()
frmmain.New_Refresh

End Sub

Private Sub MnuRefresh2_Click()
FrmBackupRegistry.TreeViewBackup.Refresh
FrmBackupRegistry.panggil_Nodes
End Sub

Private Sub MnuUninstall_Click()
frmmain.Show_FormUninstall
End Sub

Private Sub MnuUninstall2_Click()
frmmain.Show_FormUninstall
End Sub
