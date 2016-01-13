
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

Namespace Code
	Public Class EncryptionService
		' This constant is used to determine the keysize of the encryption algorithm
		Private Const Keysize As Integer = 256
		'Encrypt
		Public Shared Function EncryptString(plainText As String, passPhrase As String) As String
			' This size of the IV (in bytes) must = (keysize / 8).  Default keysize is 256, so the IV must be
			' 32 bytes long.  Using a 16 character string here gives us 32 bytes when converted to a byte array.
			Dim initVector as string = ConfigurationManager.AppSettings("as:AudienceId")
			initVector = initVector.ToString().Substring(0, 16)
			passPhrase = initVector

			Dim initVectorBytes As Byte() = Encoding.UTF8.GetBytes(initVector)
			Dim plainTextBytes As Byte() = Encoding.UTF8.GetBytes(plainText)
			Dim password As New Rfc2898DeriveBytes(passPhrase, initVectorBytes)
			Dim keyBytes As Byte() = password.GetBytes(Keysize / 8)
			Dim symmetricKey As New RijndaelManaged()
			symmetricKey.Mode = CipherMode.CBC
			Dim encryptor As ICryptoTransform = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes)
			Dim memoryStream As New MemoryStream()
			Dim cryptoStream As New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)
			cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length)
			cryptoStream.FlushFinalBlock()
			Dim cipherTextBytes As Byte() = memoryStream.ToArray()
			memoryStream.Close()
			cryptoStream.Close()
			Return Convert.ToBase64String(cipherTextBytes)
		End Function
		'Decrypt
		Public Shared Function DecryptString(cipherText As String, passPhrase As String) As String
			Dim initVector as string = ConfigurationManager.AppSettings("as:AudienceId")
			initVector = initVector.ToString().Substring(0, 16)
			passPhrase = initVector

			Dim initVectorBytes As Byte() = Encoding.ASCII.GetBytes(InitVector)
			Dim cipherTextBytes As Byte() = Convert.FromBase64String(cipherText)
			Dim password As New Rfc2898DeriveBytes(passPhrase, initVectorBytes)
			Dim keyBytes As Byte() = password.GetBytes(Keysize / 8)
			Dim symmetricKey As New RijndaelManaged()
			symmetricKey.Mode = CipherMode.CBC
			Dim decryptor As ICryptoTransform = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes)
			Dim memoryStream As New MemoryStream(cipherTextBytes)
			Dim cryptoStream As New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)
			Dim plainTextBytes As Byte() = New Byte(cipherTextBytes.Length - 1) {}
			Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)
			memoryStream.Close()
			cryptoStream.Close()
			Return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount)
		End Function

	End Class
End NameSpace