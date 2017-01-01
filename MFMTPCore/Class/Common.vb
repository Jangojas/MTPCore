Imports System
Imports System.Collections

''' <summary>
''' Provide access to global definition
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class Common

#Region "Structures"

    Public Structure WritingPair
        Dim Index As String
        Dim Value As String
    End Structure

    Public Structure ReturnedData
        Dim Code As Integer
        Dim Type As String
        Dim Value As String
    End Structure

#End Region

#Region "Constant"

    ''' <summary>
    ''' Default MarathonTP UDP port
    ''' </summary>
    ''' <remarks></remarks>
    Public Const Default_UDP_port As Integer = 8384
    ''' <summary>
    ''' String representation of MarathonTP Boolean datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonBoolean As String = "Bo"
    ''' <summary>
    ''' String representation of MarathonTP Integer datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonInteger As String = "In"
    ''' <summary>
    ''' String representation of MarathonTP Short datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonShort As String = "Sh"
    ''' <summary>
    ''' String representation of MarathonTP UShort datatype
    ''' </summary>
    Public Const MarathonUShort As String = "USh"
    ''' <summary>
    ''' String representation of MarathonTP Long datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonLong As String = "Lo"
    ''' <summary>
    ''' String representation of MarathonTP Single datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonSingle As String = "Si"
    ''' <summary>
    ''' String representation of MarathonTP Double datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonDouble As String = "Do"
    ''' <summary>
    ''' String representation of MarathonTP Byte datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonByte As String = "By"
    ''' <summary>
    ''' String representation of MarathonTP String datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonString As String = "St"
    ''' <summary>
    ''' String representation of MarathonTP NULL datatype
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MarathonNil As String = "Nil"

#End Region

#Region "Transmission Symbols"

    ''' <summary>
    ''' Request code
    ''' </summary>
    ''' <remarks></remarks>
    Public Const RequestCode As Char = CChar("R")
    ''' <summary>
    ''' Answer code
    ''' </summary>
    ''' <remarks></remarks>
    Public Const AnswerCode As Char = CChar("A")
    ''' <summary>
    ''' Beginning of message symbol
    ''' </summary>
    ''' <remarks></remarks>
    Public Const STX As Char = CChar("{")
    ''' <summary>
    ''' End of message symbol
    ''' </summary>
    ''' <remarks></remarks>
    Public Const ETX As Char = CChar("}")
    ''' <summary>
    ''' Item separator
    ''' </summary>
    ''' <remarks></remarks>
    Public Const DIV As Char = CChar(":")

#End Region

#Region "Enumeration"

    ''' <summary>
    ''' Available versions
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum MarathonVersion As Integer
        V_1 = 1
        V_1_1 = 2
    End Enum

    ''' <summary>
    ''' Available message codes
    ''' </summary>
    Public Enum MessageCodes
        ''' <summary>
        ''' Used to read tag(s) value(s) on remote device
        ''' </summary>
        ReadVariables = 1
        ''' <summary>
        ''' Used to write value(s) in tag(s) on remote device
        ''' </summary>
        WriteVariables = 2
        ''' <summary>
        ''' Used to Ping remote device
        ''' </summary>
        Discovery = 3
    End Enum

    ''' <summary>
    ''' Available error codes
    ''' </summary>
    Public Enum ErrorCodes
        ''' <summary>
        ''' The message is successfully processed
        ''' </summary>
        OperationSuccessful = 0
        ''' <summary>
        ''' The specified item does not exist
        ''' </summary>
        ElementNotFound = 1
        ''' <summary>
        ''' The datatype is incompatible
        ''' </summary>
        IncompatibleData = 2
        ''' <summary>
        ''' The specified index is out of range
        ''' </summary>
        IndexOutOfRange = 3
    End Enum

    ''' <summary>
    ''' Available Network Type
    ''' </summary>
    Public Enum NetworkTypes
        ''' <summary>
        ''' Ethernet/IP network layer
        ''' </summary>
        EthernetIP = 0
        ''' <summary>
        ''' Serial network layer
        ''' </summary>
        Serial = 1
        ''' <summary>
        ''' Radio Frequency nRF24L01p network layer
        ''' </summary>
        RF = 2
    End Enum

    ''' <summary>
    ''' Security Mode
    ''' </summary>
    Public Enum SecurityModes
        ''' <summary>
        ''' None
        ''' </summary>
        None = 0
        ''' <summary>
        ''' XTEA
        ''' </summary>
        XTEA = 1
        ''' <summary>
        ''' Advanced
        ''' </summary>
        Advanced = 2
    End Enum

#End Region

#Region "Fonctions"

    ''' <summary>
    ''' Return a string that contain protocol version
    ''' </summary>
    ''' <param name="Version"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Versioning(Version As Integer) As String
        Dim myVersion As String

        Select Case Version
            Case Common.MarathonVersion.V_1
                myVersion = "1.0"
            Case Common.MarathonVersion.V_1_1
                myVersion = "1.1"
            Case Else
                myVersion = "1.1"
        End Select

        Return myVersion
    End Function

#End Region

#Region "CRC Checksum"

    ''' <summary>
    ''' Table to calculate a CRC
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared aCRC16Table() As UInt16 = {&H0, &HC0C1, &HC181, &H140, &HC301, &H3C0, &H280, &HC241, _
        &HC601, &H6C0, &H780, &HC741, &H500, &HC5C1, &HC481, &H440, _
        &HCC01, &HCC0, &HD80, &HCD41, &HF00, &HCFC1, &HCE81, &HE40, _
        &HA00, &HCAC1, &HCB81, &HB40, &HC901, &H9C0, &H880, &HC841, _
        &HD801, &H18C0, &H1980, &HD941, &H1B00, &HDBC1, &HDA81, &H1A40, _
        &H1E00, &HDEC1, &HDF81, &H1F40, &HDD01, &H1DC0, &H1C80, &HDC41, _
        &H1400, &HD4C1, &HD581, &H1540, &HD701, &H17C0, &H1680, &HD641, _
        &HD201, &H12C0, &H1380, &HD341, &H1100, &HD1C1, &HD081, &H1040, _
        &HF001, &H30C0, &H3180, &HF141, &H3300, &HF3C1, &HF281, &H3240, _
        &H3600, &HF6C1, &HF781, &H3740, &HF501, &H35C0, &H3480, &HF441, _
        &H3C00, &HFCC1, &HFD81, &H3D40, &HFF01, &H3FC0, &H3E80, &HFE41, _
        &HFA01, &H3AC0, &H3B80, &HFB41, &H3900, &HF9C1, &HF881, &H3840, _
        &H2800, &HE8C1, &HE981, &H2940, &HEB01, &H2BC0, &H2A80, &HEA41, _
        &HEE01, &H2EC0, &H2F80, &HEF41, &H2D00, &HEDC1, &HEC81, &H2C40, _
        &HE401, &H24C0, &H2580, &HE541, &H2700, &HE7C1, &HE681, &H2640, _
        &H2200, &HE2C1, &HE381, &H2340, &HE101, &H21C0, &H2080, &HE041, _
        &HA001, &H60C0, &H6180, &HA141, &H6300, &HA3C1, &HA281, &H6240, _
        &H6600, &HA6C1, &HA781, &H6740, &HA501, &H65C0, &H6480, &HA441, _
        &H6C00, &HACC1, &HAD81, &H6D40, &HAF01, &H6FC0, &H6E80, &HAE41, _
        &HAA01, &H6AC0, &H6B80, &HAB41, &H6900, &HA9C1, &HA881, &H6840, _
        &H7800, &HB8C1, &HB981, &H7940, &HBB01, &H7BC0, &H7A80, &HBA41, _
        &HBE01, &H7EC0, &H7F80, &HBF41, &H7D00, &HBDC1, &HBC81, &H7C40, _
        &HB401, &H74C0, &H7580, &HB541, &H7700, &HB7C1, &HB681, &H7640, _
        &H7200, &HB2C1, &HB381, &H7340, &HB101, &H71C0, &H7080, &HB041, _
        &H5000, &H90C1, &H9181, &H5140, &H9301, &H53C0, &H5280, &H9241, _
        &H9601, &H56C0, &H5780, &H9741, &H5500, &H95C1, &H9481, &H5440, _
        &H9C01, &H5CC0, &H5D80, &H9D41, &H5F00, &H9FC1, &H9E81, &H5E40, _
        &H5A00, &H9AC1, &H9B81, &H5B40, &H9901, &H59C0, &H5880, &H9841, _
        &H8801, &H48C0, &H4980, &H8941, &H4B00, &H8BC1, &H8A81, &H4A40, _
        &H4E00, &H8EC1, &H8F81, &H4F40, &H8D01, &H4DC0, &H4C80, &H8C41, _
        &H4400, &H84C1, &H8581, &H4540, &H8701, &H47C0, &H4680, &H8641, _
        &H8201, &H42C0, &H4380, &H8341, &H4100, &H81C1, &H8081, &H4040}

    '*********************************
    '* This CRC use a table verification algorith for faster calculation
    '*********************************

    ''' <summary>
    ''' Calculate a CRC
    ''' </summary>
    ''' <param name="DataInput">Bytes array containing a complete MarathonTP message frame</param>
    ''' <returns>Calculated CRC</returns>
    Public Shared Function CalculateCRC16(ByVal DataInput() As Byte) As UShort
        Dim iCRC As UShort = 0
        Dim bytT As Byte

        For i As Integer = 0 To DataInput.Length - 1
            bytT = CByte((iCRC And &HFF)) Xor DataInput(i)
            iCRC = (iCRC \ CUShort(2 ^ 8)) Xor aCRC16Table(bytT)
        Next

        Return iCRC
    End Function

#End Region

End Class
