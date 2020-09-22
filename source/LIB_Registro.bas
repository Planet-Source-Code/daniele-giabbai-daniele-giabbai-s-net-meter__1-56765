Attribute VB_Name = "LIB_Registro"
'------------------------------------------------------------------
' Modulo    : modRegistroSistema
' DataOra   : 07/06/2003 14.50
' Autore    : Samuele Battarra
' Scopo     : Permettere l'accesso alle chiavi del registro
'------------------------------------------------------------------
Option Explicit

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Private Const REG_OPTION_NON_VOLATILE = 0&

Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_ALL_ACCESS = &H3F
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_MORE_DATA = 234&
Private Const ERROR_NO_MORE_ITEMS = 259&

Private Const REG_SZ = 1&
Private Const REG_EXPAND_SZ = 2&
Private Const REG_BINARY = 3&
Private Const REG_DWORD = 4&

Private Const SE_BACKUP_NAME = "SeBackupPrivilege"

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
      'Note that if you declare the lpData parameter as String, you must pass it ByVal.

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
      'Note that if you declare the lpData parameter as String, you must pass it ByVal.

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
      
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

'------------------------------------------------------------------
' Procedura : TrovaBase
' DataOra   : 07/06/2003 14.52
' Autore    : Samuele Battarra
' Scopo     : Restituisce il codice relativo al ramo principale della chiave
'------------------------------------------------------------------
Private Function TrovaBase(ByRef sNome As String) As Long

   Dim sBase As String
   Dim lPos As Long

   lPos = InStr(sNome, "\")
   If lPos > 0 Then
      sBase = UCase$(Left$(sNome, lPos - 1))
      sNome = Mid$(sNome, lPos + 1)
   Else
      sBase = sNome
      sNome = ""
   End If
   Select Case sBase
      Case "HKEY_CLASSES_ROOT": TrovaBase = HKEY_CLASSES_ROOT
      Case "HKEY_CURRENT_USER": TrovaBase = HKEY_CURRENT_USER
      Case "HKEY_LOCAL_MACHINE": TrovaBase = HKEY_LOCAL_MACHINE
      Case "HKEY_USERS": TrovaBase = HKEY_USERS
      Case "HKEY_PERFORMANCE_DATA": TrovaBase = HKEY_PERFORMANCE_DATA
      Case "HKEY_CURRENT_CONFIG": TrovaBase = HKEY_CURRENT_CONFIG
      Case "HKEY_DYN_DATA": TrovaBase = HKEY_DYN_DATA
      Case Else: TrovaBase = &H88888888   'Valore non valido
   End Select

End Function

'------------------------------------------------------------------
' Procedura : CreaChiave
' DataOra   : 07/06/2003 14.53
' Autore    : Samuele Battarra
' Scopo     : Creare una nuova chiave
'------------------------------------------------------------------
Public Function CreaChiave(ByVal sNome As String) As Boolean

   Dim lChiave As Long
   Dim lBase As Long
   Dim lRis As Long

   lBase = TrovaBase(sNome)
   CreaChiave = (RegCreateKeyEx(lBase, sNome, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_QUERY_VALUE, ByVal 0&, lChiave, lRis) = ERROR_SUCCESS)
   If CreaChiave Then CreaChiave = ChiudiChiave(lChiave)

End Function

'------------------------------------------------------------------
' Procedura : CancellaChiave
' DataOra   : 07/06/2003 15.01
' Autore    : Samuele Battarra
' Scopo     : Elimina una chiave
'------------------------------------------------------------------
Public Function CancellaChiave(ByVal sNome As String) As Boolean

   Dim lBase As Long

   lBase = TrovaBase(sNome)
   CancellaChiave = (RegDeleteKey(lBase, sNome) = ERROR_SUCCESS)

End Function

'------------------------------------------------------------------
' Procedura : ApriChiave
' DataOra   : 07/06/2003 15.05
' Autore    : Samuele Battarra
' Scopo     : Apre una chiave e restituisce il suo handle
'------------------------------------------------------------------
Private Function ApriChiave(ByVal sNome As String, _
   ByRef hChiave As Long, ByVal lAccesso As Long) As Boolean

   Dim lBase As Long

   lBase = TrovaBase(sNome)
   ApriChiave = (RegOpenKeyEx(lBase, sNome, 0&, lAccesso, hChiave) = ERROR_SUCCESS)

End Function

'------------------------------------------------------------------
' Procedura : ChiudiChiave
' DataOra   : 07/06/2003 15.14
' Autore    : Samuele Battarra
' Scopo     : Chiude una chiave dato l'handle
'------------------------------------------------------------------
Private Function ChiudiChiave(ByVal hChiave As Long) As Boolean

   ChiudiChiave = (RegCloseKey(hChiave) = ERROR_SUCCESS)

End Function

'------------------------------------------------------------------
' Procedura : EnumeraSottoChiavi
' DataOra   : 07/06/2003 20.03
' Autore    : Samuele Battarra
' Scopo     : Elenca le sottochiavi di una chiave
'------------------------------------------------------------------
Public Function EnumeraSottoChiavi(ByVal sChiave As String, ByVal lIndice As Long, _
   ByRef sSubChiave As String) As Boolean

   Dim hChiave As Long
   Dim lNumCar As Long
   Dim uData As FILETIME
   Dim lRet As Long

   EnumeraSottoChiavi = False
   If ApriChiave(sChiave, hChiave, KEY_ENUMERATE_SUB_KEYS Or KEY_QUERY_VALUE) Then
      sSubChiave = Space$(10000)
      lNumCar = 10000
      lRet = RegEnumKeyEx(hChiave, lIndice, sSubChiave, lNumCar, 0&, 0&, 0&, uData)
      If lRet = ERROR_MORE_DATA Then
         sSubChiave = Left$(sSubChiave, lNumCar)
         EnumeraSottoChiavi = True
      End If
      EnumeraSottoChiavi = (lRet <> ERROR_NO_MORE_ITEMS)
      ChiudiChiave hChiave
   End If

End Function

'------------------------------------------------------------------
' Procedura : LeggiChiaveStringa
' DataOra   : 07/06/2003 15.15
' Autore    : Samuele Battarra
' Scopo     : Legge una stringa dalla chiave specificata,
'             se non ci riesce restituisce il valore di dafault
'------------------------------------------------------------------
Public Function LeggiChiaveStringa(ByVal sChiave As String, ByVal sNome As String, _
   ByRef sValore As String, Optional ByVal sDefault As String = "") As Boolean

   Dim hChiave As Long
   Dim lDimensione As Long
   Dim lTipo As Long

   LeggiChiaveStringa = False
   If ApriChiave(sChiave, hChiave, KEY_QUERY_VALUE) Then
      If (RegQueryValueEx(hChiave, sNome, 0&, lTipo, ByVal 0&, _
         lDimensione) = ERROR_SUCCESS) And (lTipo = REG_SZ) Then
         sValore = Space$(lDimensione)
         If RegQueryValueEx(hChiave, sNome, 0&, ByVal 0&, ByVal sValore, _
            lDimensione) = ERROR_SUCCESS Then
            If lDimensione > 0 Then sValore = Left$(sValore, lDimensione - 1)
            LeggiChiaveStringa = True
         End If
      End If
      ChiudiChiave hChiave
   End If
   If Not LeggiChiaveStringa Then sValore = sDefault

End Function

'------------------------------------------------------------------
' Procedura : LeggiChiaveBinario
' DataOra   : 07/06/2003 19.24
' Autore    : Samuele Battarra
' Scopo     : Legge un valore binario dalla chiave specificata,
'             se non ci riesce restituisce il valore di dafault.
'             Qualsiasi chiave, anche di tipo non binario, può
'             essere letta come binario.
'------------------------------------------------------------------
Public Function LeggiChiaveBinario(ByVal sChiave As String, ByVal sNome As String, _
   ByRef byValore() As Byte) As Boolean

   Dim hChiave As Long
   Dim lDimensione As Long
   Dim lTipo As Long

   LeggiChiaveBinario = False
   If ApriChiave(sChiave, hChiave, KEY_QUERY_VALUE) Then
      If (RegQueryValueEx(hChiave, sNome, 0&, lTipo, ByVal 0&, _
         lDimensione) = ERROR_SUCCESS) Then
         ReDim byValore(0 To lDimensione - 1)
         LeggiChiaveBinario = (RegQueryValueEx(hChiave, sNome, 0&, ByVal 0&, _
            byValore(0), lDimensione) = ERROR_SUCCESS)
      End If
      ChiudiChiave hChiave
   End If

End Function

'------------------------------------------------------------------
' Procedura : LeggiChiaveNumero
' DataOra   : 07/06/2003 19.37
' Autore    : Samuele Battarra
' Scopo     : Legge un numero dalla chiave specificata,
'             se non ci riesce restituisce il valore di dafault
'------------------------------------------------------------------
Public Function LeggiChiaveNumero(ByVal sChiave As String, ByVal sNome As String, _
   ByRef lValore As Long, Optional ByVal lDefault As Long = 0) As Boolean

   Dim hChiave As Long
   Dim lTipo As Long

   LeggiChiaveNumero = False
   If ApriChiave(sChiave, hChiave, KEY_QUERY_VALUE) Then
      LeggiChiaveNumero = ((RegQueryValueEx(hChiave, sNome, 0&, _
         lTipo, lValore, 4) = ERROR_SUCCESS) And (lTipo = REG_DWORD))
      ChiudiChiave hChiave
   End If
   If Not LeggiChiaveNumero Then lValore = lDefault

End Function

'------------------------------------------------------------------
' Procedura : LeggiChiaveBooleano
' DataOra   : 07/06/2003 19.38
' Autore    : Samuele Battarra
' Scopo     : Legge un valore vero/falso dalla chiave specificata,
'             se non ci riesce restituisce il valore di dafault
'------------------------------------------------------------------
Public Function LeggiChiaveBooleano(ByVal sChiave As String, ByVal sNome As String, _
   ByRef bValore As Boolean, Optional ByVal bDefault As Boolean = False) As Boolean

   Dim lRet As Long

   LeggiChiaveBooleano = LeggiChiaveNumero(sChiave, sNome, lRet, IIf(bDefault, 1, 0))
   bValore = (lRet <> 0)

End Function

'------------------------------------------------------------------
' Procedura : ScriviChiaveStringa
' DataOra   : 07/06/2003 19.40
' Autore    : Samuele Battarra
' Scopo     : Scrive una stringa nella chiave se il suo valore è diverso
'             da quello di default, altrimenti cancella la stringa dalla chiave
'------------------------------------------------------------------
Public Function ScriviChiaveStringa(ByVal sChiave As String, ByVal sNome As String, _
   ByVal sValore As String, Optional ByVal sDefault As String = "") As Boolean

   Dim hChiave As Long

   ScriviChiaveStringa = False
   If ApriChiave(sChiave, hChiave, KEY_SET_VALUE) Then
      If sValore <> sDefault Then
         ScriviChiaveStringa = (RegSetValueEx(hChiave, sNome, 0&, REG_SZ, _
            ByVal sValore, LenB(StrConv(sValore, vbFromUnicode)) + 1) = ERROR_SUCCESS)
      Else
         ScriviChiaveStringa = CancellaValore(hChiave, sNome)
      End If
      ChiudiChiave hChiave
   End If

End Function

'------------------------------------------------------------------
' Procedura : ScriviChiaveBinario
' DataOra   : 07/06/2003 19.56
' Autore    : Samuele Battarra
' Scopo     : Scrive un valore binario nella chiave
'------------------------------------------------------------------
Public Function ScriviChiaveBinario(ByVal sChiave As String, ByVal sNome As String, _
   ByRef byValore() As Byte) As Boolean

   Dim hChiave As Long
   Dim lDimensione As Long

   ScriviChiaveBinario = False
   If ApriChiave(sChiave, hChiave, KEY_SET_VALUE) Then
      lDimensione = UBound(byValore) - LBound(byValore) + 1
      ScriviChiaveBinario = (RegSetValueEx(hChiave, sNome, 0&, _
         REG_BINARY, byValore(0), lDimensione) = ERROR_SUCCESS)
      ChiudiChiave hChiave
   End If

End Function

'------------------------------------------------------------------
' Procedura : ScriviChiaveNumero
' DataOra   : 07/06/2003 19.57
' Autore    : Samuele Battarra
' Scopo     : Scrive un numero nella chiave se il suo valore è diverso
'             da quello di default, altrimenti cancella il numero dalla chiave
'------------------------------------------------------------------
Public Function ScriviChiaveNumero(ByVal sChiave As String, ByVal sNome As String, _
   ByVal lValore As Long, Optional ByVal lDefault As Long = 0) As Boolean

   Dim hChiave As Long

   ScriviChiaveNumero = False
   If ApriChiave(sChiave, hChiave, KEY_SET_VALUE) Then
      If lValore <> lDefault Then
         ScriviChiaveNumero = (RegSetValueEx(hChiave, sNome, 0&, _
            REG_DWORD, lValore, 4) = ERROR_SUCCESS)
      Else
         ScriviChiaveNumero = CancellaValore(hChiave, sNome)
      End If
      ChiudiChiave hChiave
   End If

End Function

'------------------------------------------------------------------
' Procedura : ScriviChiaveBooleano
' DataOra   : 07/06/2003 19.58
' Autore    : Samuele Battarra
' Scopo     : 'Scrive un valore vero/falso nella chiave se il suo valore è
'             diverso da quello di default, altrimenti cancella il valore dalla chiave
'------------------------------------------------------------------
Public Function ScriviChiaveBooleano(ByVal sChiave As String, ByVal sNome As String, _
   ByVal bValore As Boolean, Optional ByVal bDefault As Boolean = False) As Boolean

   ScriviChiaveBooleano = ScriviChiaveNumero(sChiave, sNome, _
      IIf(bValore, 1, 0), IIf(bDefault, 1, 0))

End Function

'------------------------------------------------------------------
' Procedura : CancellaValore
' DataOra   : 07/06/2003 19.46
' Autore    : Samuele Battarra
' Scopo     : Cancella un valore (stringa o numero) da una chiave
' Modifica  : 28/04/2004 11:00 Daniele Giabbai
'------------------------------------------------------------------
Public Function CancellaValore(ByVal sChiave As String, _
    ByVal sNome As String) As Boolean
  
  Dim hChiave As Long
  
  If Not ApriChiave(sChiave, hChiave, KEY_ALL_ACCESS) Then Exit Function
  CancellaValore = (RegDeleteValue(hChiave, sNome) = ERROR_SUCCESS)
  ChiudiChiave hChiave
End Function

Public Function EnumeraValori(ByVal sChiave As String, ByVal lIndice As Long, _
   ByRef sNome As String, ByRef lTipo As Long) As Boolean

   Dim hChiave As Long
   Dim lNumCar As Long
   Dim lRet As Long

   EnumeraValori = False
   If ApriChiave(sChiave, hChiave, KEY_QUERY_VALUE) Then
      sNome = Space$(10000)
      lNumCar = 10000
      lRet = RegEnumValue(hChiave, lIndice, sNome, _
         lNumCar, 0&, lTipo, ByVal 0&, ByVal 0&)
      If lRet = ERROR_SUCCESS Then
         sNome = Left$(sNome, lNumCar)
         EnumeraValori = (lRet <> ERROR_NO_MORE_ITEMS)
      End If
      ChiudiChiave hChiave
   End If

End Function

'------------------------------------------------------------------
' Procedura : SalvaChiave
' DataOra   : 08/06/2003 18.11
' Autore    : Samuele Battarra
' Scopo     : Salvare il contenuto di una chiave in un file .reg
' Note      : Se bAppend è True le informazioni verranno aggiunte
'             al file, altrimenti questo verrà sovrascritto
'------------------------------------------------------------------
Public Function SalvaChiave(ByVal sNomeChiave As String, _
   ByVal sNomeFile As String, Optional ByVal bAppend As Boolean = False) As Boolean

   Dim sRegEdit As String
   Dim lCar As Long

   sRegEdit = Space$(256)
   lCar = GetWindowsDirectory(sRegEdit, Len(sRegEdit))
   sRegEdit = Left$(sRegEdit, lCar) & "\regedit.exe /save "
   If Not bAppend Then
      Shell sRegEdit & """" & sNomeFile & """ """ & sNomeChiave & """"
   Else
      Shell sRegEdit & """" & App.Path & "\tmppmt.reg"" """ & sNomeChiave & """"
      UnisciReg sNomeFile, App.Path & "\tmppmt.reg"
      Kill App.Path & "\tmppmt.reg"
   End If

End Function

'------------------------------------------------------------------
' Procedura : ImportaReg
' DataOra   : 08/06/2003 18.11
' Autore    : Samuele Battarra
' Scopo     : Importare il contenuto di una file .reg nel registro
'------------------------------------------------------------------
Public Function ImportaReg(ByVal sNomeFile As String) As Boolean

   Dim sPath As String

   sPath = Space$(256)
   GetWindowsDirectory sPath, Len(sPath)
   Shell sPath & "\regedit.exe """ & sNomeFile & """"
      
End Function

'------------------------------------------------------------------
' Procedura : UnisciReg
' DataOra   : 12/07/2003 21.38
' Autore    : Samuele Battarra
' Scopo     : Aggiunge un file reg ad un'altro
'------------------------------------------------------------------
Private Sub UnisciReg(ByVal sOutputFile As String, ByVal sInputFile As String)

   Dim lFileInput As Long
   Dim lFileOutput As Long
   Dim sTmp As String
   Dim bUnicode As Boolean

   On Error GoTo UnisciReg_Errore

   bUnicode = UnicodeFile(sOutputFile)
   lFileInput = FreeFile
   Open sInputFile For Input As lFileInput
   lFileOutput = FreeFile
   Open sOutputFile For Append As lFileOutput
   If Not EOF(lFileInput) Then Line Input #lFileInput, sTmp
   Do Until EOF(lFileInput)
      Line Input #lFileInput, sTmp
      If bUnicode Then
         Print #lFileOutput, StrConv(sTmp & vbNewLine, vbUnicode);
      Else
         Print #lFileOutput, sTmp
      End If
   Loop
   Close lFileInput, lFileOutput
   Exit Sub

UnisciReg_Errore:

   MsgBox "Errore " & Err.Number & " nella procedura modRegistroSistema." & _
      "UnisciReg." & vbCrLf & vbCrLf & Err.Description, vbCritical

End Sub

'------------------------------------------------------------------
' Procedura : UnicodeFile
' DataOra   : 14/07/2003 18.25
' Autore    : Samuele Battarra
' Scopo     : Dato un file di testo, dice se è un file unicode o no
'------------------------------------------------------------------
Private Function UnicodeFile(ByVal sFile As String) As Boolean

   Const txtfmtUnicode = &HFEFF
   Const txtfmtBigEndianUnicode = &HFFFE

   Dim lFile As Long
   Dim iFlag As Integer

   On Error GoTo UnicodeFile_Errore

   UnicodeFile = False
   If FileLen(sFile) >= 2 Then
      lFile = FreeFile
      Open sFile For Binary Access Read As lFile
      Get #lFile, , iFlag
      Close lFile
      UnicodeFile = (iFlag = txtfmtUnicode)
   End If
   Exit Function

UnicodeFile_Errore:

   MsgBox "Errore " & Err.Number & " nella procedura modRegistroSistema." & _
      "UnicodeFile." & vbCrLf & vbCrLf & Err.Description, vbCritical
   
End Function
