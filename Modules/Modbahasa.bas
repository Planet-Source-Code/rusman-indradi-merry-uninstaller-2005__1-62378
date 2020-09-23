Attribute VB_Name = "Modbahasa"
'---------------------------------------------------------------------------------------
' Nama File  : Modbahasa.bas
' Tanggal    : 8/29/2005 22:28
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

Public Sub Applybahasaindonesia()
On Error Resume Next
'Perubaham Menu Ke Bahasa Indonesia
'Menu File
frmmain.file.Caption = "File"
frmmain.backup.Caption = "Backup uninstal registry"
frmmain.Report.Caption = "Laporan uninstaller"
frmmain.exit.Caption = "Keluar"
'Menu Edit
frmmain.Edit.Caption = "Mengedit"
frmmain.Entry.Caption = "Mengedit Masukan"
frmmain.delete.Caption = "Menghapus Masukan"
frmmain.new.Caption = "Masukan Baru"
'Menu View
frmmain.view.Caption = "Pandangan"
frmmain.Detailsview.Caption = "Details Daftar"
frmmain.list.Caption = "Daftar"
frmmain.small.Caption = "Icon Kecil"
frmmain.large.Caption = "Icon Lebar"
frmmain.Minimize.Caption = "Sembunyikan Aplikasi"
frmmain.Restore.Caption = "Kembalikan Aplikasi"
'Menu Uninstall
frmmain.Uninstall.Caption = "Uninstall"
frmmain.Uninstall1.Caption = "Uninstall"
frmmain.Information.Caption = "Informasi Program"
'Menu Tool
frmmain.Tool.Caption = "Alat"
frmmain.language.Caption = "Bahasa"

frmmain.bhsIndonesia.Caption = "Bahasa Indonesia"
frmmain.Options.Caption = "Pilihan"
frmmain.tweak.Caption = "Tweak Uninstaller"
frmmain.request.Caption = "Permintaan"
frmmain.Password.Caption = "Kata Sandi"
'Menu Help
frmmain.Help.Caption = "Pertolongan"
frmmain.faq.Caption = "FAQ (Pertanyaan Yang Sering di Tanyakan)"
frmmain.faqina.Caption = "FAQ Indonesia"
frmmain.faqeng.Caption = "FAQ Inggris"
frmmain.Readme.Caption = "Baca"
frmmain.Readmeindo.Caption = "Baca Bahasa Indonesia"
frmmain.Readmeeng.Caption = "Baca Bahasa Inggris"
frmmain.HelpTopics.Caption = "Pertolongan"
frmmain.About.Caption = "Tentang"

'Ganti Bahasa Pada Tombol FRMMAIN

frmmain.cmddelete.Caption = "hapus Dari Daftar"
frmmain.cmdinfo.Caption = "Informasi Detail"
frmmain.cmdreport.Caption = "Laporan Program"
frmmain.cmdoptions.Caption = "Plihan"
frmmain.cmdtweak.Caption = "Tweak Uninstaller"


'Ganti Bahasa Pada lstview FRMMAIN

'Ganti Bahasa pada FRMABOUT
FrmAbout.Caption = "Mery Uninstaller - Tentang"
FrmAbout.Label1.Caption = "Jika kamu temukan dimanapun kutu busuk pada program ini,silahkan mengirimkan e-mail atau melaporkan kepada teknisi kami.Terima Kasih!"
FrmAbout.lblAbout(3).Caption = "Hak Cipta © 2001 - 2003"
FrmAbout.lblAbout(6).Caption = "Pembuat: Rusman Indradi"
FrmAbout.Text1.text = "Informasi yang disajikan oleh program ini dimaksudkan semata-mata untuk menyediakan bimbingan umum pada berbagai hal minat untuk penggunaan yang pribadi pemakai dari program ini,siapa yang menerima tanggung jawab penuh untuk penggunaannya.Itu disajikan 'seperti halnya', dengan tidak ada jaminan kelengkapan atau ketelitian dan tanpa jaminan keabsahan tentang segala hal, menyatakan atau menyiratkan. MENGGUNAKAN INI MENJADI RESIKO KAMU."

'Ganti Bahasa Pada FrmBackupRegistry
FrmBackupRegistry.LblOK.Caption = "Simpan"
FrmBackupRegistry.LblBatal.Caption = "Batal"
FrmBackupRegistry.Label2.Caption = "Peringatan: tolong tidak memodifikasi teks pada diatas pencatatan daftar. Jika kamu ingin menambahkan informasi atau uraian baru, kamu dapat menambahkan string ';' di depan sebelum kamu tulis uraian."
FrmBackupRegistry.Label3.Caption = "Contoh >> ;Tulis informasi baru"

'Ganti Bahasa Pada FrmEditEntry
FrmEditEntry.Caption = "Mery Uninstaller - Mengedit Masukan Pencatatan"
FrmEditEntry.Label2.Caption = "Mengedit Untuk Nama Pencatatan (Kunci) :"
FrmEditEntry.Label3.Caption = "Display Nama :"
FrmEditEntry.Label4.Caption = "Uninstall String :"
FrmEditEntry.LblOK.Caption = "Simpan"
FrmEditEntry.LblBatal.Caption = "Batal"

'Ganti Bahasa Pada FrmEntryBaru
FrmEntryBaru.Caption = "Mery uninstaller - Masukan Baru"
FrmEntryBaru.LblDiplayName.Caption = "Pajangan Nama :"
FrmEntryBaru.LblUninstallString.Caption = "Uninstall String :"
FrmEntryBaru.LblOK.Caption = "Simpan"
FrmEntryBaru.LblBatal.Caption = "Batal"

'Ganti Bahasa Pada FrmHapusEntry
FrmHapusEntry.Caption = "Mery Uninstaller - Hilangkan Masukan"
FrmHapusEntry.Label1.Caption = "Apakah kamu pasti untuk menghilangkan masukan daftar pada pencatatan ?"
FrmHapusEntry.Label2.Caption = "Perhatian : Sebelum Anda menghapus List entry pada registry, sebaiknya jalankan dulu perintah Uninstall. Jika tetap yakin untuk menghapus maka perintah untuk uninstall pada registry akan terhapus. Apakah Anda tetap yakin ?"
FrmHapusEntry.LblUninstall.Caption = "Uninstall"
FrmHapusEntry.LblHapus.Caption = "Hapus"
FrmHapusEntry.LblBatal.Caption = "Batal"

'Ganti Bahasa Pada frmInformasi
frmInformasi.LblDes.Caption = "Informasi ini ditemukan dalam registry window"

'Ganti Bahasa Pada FrmLaporan
FrmLaporan.Caption = "Mery Uninstaller - Laporan"
FrmLaporan.LblPrint.Caption = "Print"
FrmLaporan.LblSimpan.Caption = "Simpan"
FrmLaporan.LblBuka.Caption = "Buka"
FrmLaporan.LblTutup.Caption = "Tutup"

'Ganti Bahasa Pada FrmOption
FrmOption.Caption = "Mery Uninstaller - Alat"
FrmOption.XPFrame2.Caption = "Tampilan List View berdasarkan :"
FrmOption.XPFrame4.Caption = "Urut List View Secara :"
FrmOption.XPFrame3.Caption = "Buat Shortcut di :"
FrmOption.ChkUninstall.Caption = "Tampilkan Dialog Uninstall !"
FrmOption.OptDetail.Caption = "Detail"

FrmOption.OptIconKecil.Caption = "Icon Kecil"
FrmOption.OptIconBesar.Caption = "Icon Besar"

'Ganti Bahasa Pada FrmPilihLaporan.
FrmPilihLaporan.Caption = "Mery Uninstaller - Pilih Laporan"
FrmPilihLaporan.OptDName.Caption = "Hanya menampilkan 'Display Name'"
FrmPilihLaporan.OptDnameUString.Caption = "Display Name Dan Uninstall String"
FrmPilihLaporan.OptDetailInformasi.Caption = "Detail Informasi Laporan"
FrmPilihLaporan.LblBatal.Caption = "Batal"

'Ganti Bahasa Pada frmUninstall
frmUninstall.LblBersihkan.Caption = "Bersihkan"
frmUninstall.LblBatal.Caption = "Batal"
frmUninstall.LblInfomasi.Caption = "Jika Uninstall program berhasil dijalankan maka  tekan 'Bersihkan' untuk menghapus list program dari Registry. Dan tekan 'Batal' jika Uninstall program dibatalkan !"
frmUninstall.LblInfomasi2.Caption = "Apakah anda yakin untuk menjalankan Uninstall program ini ?"
End Sub
