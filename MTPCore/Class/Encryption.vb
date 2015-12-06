Imports System.Text
Imports System

Namespace Encryption

    Public Class XTEA

        Public Function Encrypt(num_rounds As Byte, Data As String, key As String) As Byte()
            Dim outputByte() As Byte = Nothing
            Dim dataBytes() As Byte
            Dim formattedKey() As UInteger = FormatKey(key)
            Dim fourBytes(3) As Byte
            Dim tempData(1) As UInteger
            Dim i_increment As Integer = 0
            Dim cipherDataLength As Integer
            Try
                'make sure data is in length of multiples of 8
                If Data.Length Mod 8 <> 0 Then
                    For i = 0 To (8 - Data.Length Mod 8) - 1
                        Data += " "
                    Next i
                End If
                dataBytes = Encoding.UTF8.GetBytes(Data)
                cipherDataLength = CInt(dataBytes.Length / 4) - 1
                Dim cipherBytes(dataBytes.Length - 1) As Byte
                Dim cipherData(cipherDataLength) As UInteger

                For i = 0 To Data.Length - 1 Step 8

                    For j = 0 To 3
                        fourBytes(j) = dataBytes(i + j)
                    Next
                    tempData(0) = BitConverter.ToUInt32(fourBytes, 0)
                    For j = 0 To 3
                        fourBytes(j) = dataBytes(i + 4 + j)
                    Next j
                    tempData(1) = BitConverter.ToUInt32(fourBytes, 0)

                    'XTEA Algorithm
                    Dim y As UInteger = tempData(0)
                    Dim z As UInteger = tempData(1)
                    Dim sum As UInteger = 0
                    Dim delta As UInteger = &H9E3779B9UI

                    For j = 0 To num_rounds - 1
                        Try
                            y += (CUInt(z * CUInt(2 ^ 4)) Xor CUInt(Math.Floor(z / CUInt(2 ^ 5)))) + z Xor sum + formattedKey(CInt(sum And 3))
                            sum += delta
                            z += (CUInt(y * CUInt(2 ^ 4)) Xor CUInt(Math.Floor(y / CUInt(2 ^ 5)))) + y Xor sum + formattedKey(CInt(CUInt(Math.Floor(sum / CUInt(2 ^ 11))) And 3))
                        Catch ex As Exception

                        End Try
                    Next

                    cipherData(i_increment) = y
                    cipherData(i_increment + 1) = z
                    i_increment += 2

                Next i

                i_increment = 0
                For i = 0 To cipherDataLength
                    Dim tempBytes() As Byte
                    tempBytes = BitConverter.GetBytes(cipherData(i))
                    cipherBytes(i_increment) = tempBytes(0)
                    cipherBytes(i_increment + 1) = tempBytes(1)
                    cipherBytes(i_increment + 2) = tempBytes(2)
                    cipherBytes(i_increment + 3) = tempBytes(3)
                    i_increment += 4
                Next

                outputByte = cipherBytes

            Catch ex As Exception
                Debug.Write(ex.Message)
            End Try

            Return outputByte

        End Function

        Public Function Decrypt(num_rounds As Byte, Data As Byte(), key As String) As String
            Dim DecryptedData As String = ""
            Dim dataBytes(Data.Length \ 2) As Byte
            Dim formattedKey() As UInteger = FormatKey(key)
            Dim x As Integer = 0
            Dim tempData(1) As UInteger
            Dim decipherDataLength As Integer = CInt(Data.Length / 4) - 1
            Dim decipherData(decipherDataLength) As UInteger
            Dim i_increment As Integer = 0

            Try
                For i As Integer = 0 To Data.Length - 1 Step 8

                    tempData(0) = BitConverter.ToUInt32(Data, i)
                    tempData(1) = BitConverter.ToUInt32(Data, i + 4)

                    'XTEA Algorithm
                    Dim y As UInteger = tempData(0)
                    Dim z As UInteger = tempData(1)
                    Dim delta As UInteger = &H9E3779B9UI
                    Dim sum As UInteger = &HC6EF3720UI
                    Dim n As UInteger = 32

                    For j = 0 To num_rounds - 1
                        Try
                            z -= (CUInt(y * CUInt(2 ^ 4)) Xor CUInt(Math.Floor(y / CUInt(2 ^ 5)))) + y Xor sum + formattedKey(CInt(CUInt(Math.Floor(sum / CUInt(2 ^ 11))) And 3))
                            sum -= delta
                            y -= (CUInt(z * CUInt(2 ^ 4)) Xor CUInt(Math.Floor(z / CUInt(2 ^ 5)))) + z Xor sum + formattedKey(CInt(sum And 3))
                        Catch ex As Exception

                        End Try
                    Next

                    decipherData(i_increment) = y
                    decipherData(i_increment + 1) = z
                    i_increment += 2

                Next

                For i = 0 To decipherDataLength
                    Dim tempBytes() As Byte
                    tempBytes = BitConverter.GetBytes(decipherData(i))
                    DecryptedData += Encoding.UTF8.GetChars(tempBytes)
                Next
            Catch ex As Exception

            End Try

            Return DecryptedData

        End Function

        'format the key to make it 16Bytes long
        Private Function FormatKey(key As String) As UInteger()
            Dim formattedKey(3) As UInteger

            Try
                If key.Length <= 0 Or key.Length > 16 Then
                    Throw New ArgumentException("Key must be between 1 And 16 characters in length.")
                End If

                'padding the string to ensure the key is 16 chars in length
                For i = 0 To 15 - key.Length
                    key += " "
                Next

                'Get every byte of key
                Dim keybyte() As Byte = Encoding.UTF8.GetBytes(key)
                Try
                    Dim j As Integer = 0
                    For i As Integer = 0 To key.Length - 1 Step 4
                        formattedKey(j) = CUInt(keybyte(i) + keybyte(i + 1) * 2 ^ 8 + keybyte(i + 2) * 2 ^ 16 + keybyte(i + 3) * 2 ^ 24)
                        j = j + 1
                    Next
                Catch ex As Exception
                    Debug.Write(ex.Message)
                End Try

            Catch ex As Exception
                Debug.Write(ex.Message)
            End Try

            Return formattedKey

        End Function

    End Class

End Namespace

Namespace Unicode
    ''' <summary>
    ''' http://anubis.dkuug.dk/JTC1/SC2/WG2/docs/n1335
    ''' 
    ''' http://www.cl.cam.ac.uk/~mgk25/ucs/ISO-10646-UTF-8.html
    ''' 
    ''' http://www.unicode.org/versions/corrigendum1.html
    ''' 
    ''' http://www.ietf.org/rfc/rfc2279.txt
    ''' 
    ''' </summary>
    Public Class Utf8Checker

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="buffer"></param>
        ''' <param name="length"></param>
        ''' <returns></returns>
        Public Shared Function IsUtf8(buffer As Byte(), length As Integer) As Boolean
            Dim position As Integer = 0
            Dim bytes As Integer = 0
            While position < length
                If Not IsValid(buffer, position, length, bytes) Then
                    Return False
                End If
                position += bytes
            End While
            Return True
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="buffer"></param>
        ''' <param name="position"></param>
        ''' <param name="length"></param>
        ''' <param name="bytes"></param>
        ''' <returns></returns>
        Public Shared Function IsValid(buffer As Byte(), position As Integer, length As Integer, ByRef bytes As Integer) As Boolean
            If length > buffer.Length Then
                Throw New ArgumentException("Invalid length")
            End If

            If position > length - 1 Then
                bytes = 0
                Return True
            End If

            Dim ch As Byte = buffer(position)

            If ch <= &H7F Then
                bytes = 1
                Return True
            End If

            If ch >= &HC2 AndAlso ch <= &HDF Then
                If position >= length - 2 Then
                    bytes = 0
                    Return False
                End If
                If buffer(position + 1) < &H80 OrElse buffer(position + 1) > &HBF Then
                    bytes = 0
                    Return False
                End If
                bytes = 2
                Return True
            End If

            If ch = &HE0 Then
                If position >= length - 3 Then
                    bytes = 0
                    Return False
                End If

                If buffer(position + 1) < &HA0 OrElse buffer(position + 1) > &HBF OrElse buffer(position + 2) < &H80 OrElse buffer(position + 2) > &HBF Then
                    bytes = 0
                    Return False
                End If
                bytes = 3
                Return True
            End If


            If ch >= &HE1 AndAlso ch <= &HEF Then
                If position >= length - 3 Then
                    bytes = 0
                    Return False
                End If

                If buffer(position + 1) < &H80 OrElse buffer(position + 1) > &HBF OrElse buffer(position + 2) < &H80 OrElse buffer(position + 2) > &HBF Then
                    bytes = 0
                    Return False
                End If

                bytes = 3
                Return True
            End If

            If ch = &HF0 Then
                If position >= length - 4 Then
                    bytes = 0
                    Return False
                End If

                If buffer(position + 1) < &H90 OrElse buffer(position + 1) > &HBF OrElse buffer(position + 2) < &H80 OrElse buffer(position + 2) > &HBF OrElse buffer(position + 3) < &H80 OrElse buffer(position + 3) > &HBF Then
                    bytes = 0
                    Return False
                End If

                bytes = 4
                Return True
            End If

            If ch = &HF4 Then
                If position >= length - 4 Then
                    bytes = 0
                    Return False
                End If

                If buffer(position + 1) < &H80 OrElse buffer(position + 1) > &H8F OrElse buffer(position + 2) < &H80 OrElse buffer(position + 2) > &HBF OrElse buffer(position + 3) < &H80 OrElse buffer(position + 3) > &HBF Then
                    bytes = 0
                    Return False
                End If

                bytes = 4
                Return True
            End If

            If ch >= &HF1 AndAlso ch <= &HF3 Then
                If position >= length - 4 Then
                    bytes = 0
                    Return False
                End If

                If buffer(position + 1) < &H80 OrElse buffer(position + 1) > &HBF OrElse buffer(position + 2) < &H80 OrElse buffer(position + 2) > &HBF OrElse buffer(position + 3) < &H80 OrElse buffer(position + 3) > &HBF Then
                    bytes = 0
                    Return False
                End If

                bytes = 4
                Return True
            End If

            Return False
        End Function
    End Class

End Namespace
