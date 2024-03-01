Attribute VB_Name = "modCrypto"
Private Declare Function EncryptFile Lib "ADVAPI32" Alias _
   "EncryptFileA" (ByVal lpFileName As String) As Boolean
Private Declare Function DecryptFile Lib "ADVAPI32" Alias _
    "DecryptFileA" (ByVal lpFileName As String, ByVal dwReserved As Long) As Boolean

Public Function Encrypt(ByVal strFileName As String) As Boolean
    Encrypt = (EncryptFile(strFileName) = True)
End Function

Public Function Decrypt(ByVal strFileName As String)
   Decrypt = (DecryptFile(strFileName, 0) = True)
End Function

