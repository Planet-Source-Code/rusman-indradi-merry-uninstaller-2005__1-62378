Attribute VB_Name = "ModRegistry"
'---------------------------------------------------------------------------------------
' Nama File  : ModRegistry.bas
' Tanggal    : 8/29/2005 22:27
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

Public sKeys As Collection
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1
Const REG_DWORD = 4
Public Const REG_EXPAND_SZ = 2
Private strbuffer
Private lResult As String
Dim r As String
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Enum HKeyTypes
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
End Enum


Public Function GetString(hKey As HKeyTypes, strPath As String, strValue As String)
     
    Dim keyhand As Long
    Dim datatype As Long
    Dim lRegResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long
    
  lRegResult = RegOpenKey(hKey, strPath, keyhand)
  lRegResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
  intZeroPos = InStr(strbuffer, Chr$(0))

    If lValueType = REG_SZ Or REG_EXPAND_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)


        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))


            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function


Public Sub SaveString(hKey As HKeyTypes, strPath As String, strValue As String, strdata As String)
      Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub


Public Function DeleteValue(ByVal hKey As HKeyTypes, ByVal strPath As String, ByVal strValue As String)

    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
End Function


Public Function DeleteKey(ByVal hKey As HKeyTypes, ByVal strPath As String)
  
    Dim keyhand As Long
    r = RegDeleteKey(hKey, strPath)
End Function
        

Public Sub GetKeyNames(ByVal hKey As Long, ByVal strPath As String)
Dim Cnt As Long, StrBuff As String, StrKey As String, TKey As Long
    RegOpenKey hKey, strPath, TKey
    Do
        StrBuff = String(255, vbNullChar)
        If RegEnumKeyEx(TKey, Cnt, StrBuff, 255, 0, vbNullString, 0, ByVal 0&) <> 0 Then Exit Do
        Cnt = Cnt + 1
        StrKey = Left(StrBuff, InStr(StrBuff, vbNullChar) - 1)
        sKeys.Add StrKey
    Loop
End Sub

Public Sub SaveKey(mPath As String, sfile As String)
   
    Dim temp As String
    FileAppend "", sfile
    temp = GetDosPath(sfile)
    Shell "regedit /E " & temp & " " & Chr(34) & mPath & Chr(34)
End Sub

Public Sub FileAppend(text As String, FilePath As String)
On Error Resume Next
Dim f As Integer
f = FreeFile
Dim Directory As String
              Directory$ = FilePath
    Open Directory$ For Append As #f
        Print #f, text
    Close #f
Exit Sub
End Sub

Public Function GetDosPath(LongPath As String) As String
    Dim s As String
    Dim i As Long
    Dim PathLength As Long
    i = Len(LongPath) + 1
    s = String(i, 0)
    PathLength = GetShortPathName(LongPath, s, i)
    GetDosPath = Left$(s, PathLength)

End Function



