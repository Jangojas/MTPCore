Imports System.Net
Imports System.Collections

Namespace Packet.Request

    ''' <summary>
    ''' Value reading request definition.
    ''' </summary>
    Public Class ReadVariablesPacketV1_1
        ''' <summary>
        ''' MarathonTP protocol version
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly VER As String

        ''' <summary>
        ''' MarathonTP command type
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly CMD As Byte

        ''' <summary>
        ''' MarathonTP message type code
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly RA As Char

        ''' <summary>
        ''' MarathonTP message transaction number
        ''' </summary>
        Public Property TNS As UShort
            Get
                Return m_TNS
            End Get
            Set(value As UShort)
                m_TNS = value
            End Set
        End Property
        Private m_TNS As UShort

        ''' <summary>
        ''' MarathonTP element(s) to read
        ''' </summary>
        ''' <remarks>Can be TAG or index, but must be all the same method.</remarks>
        Public Property VALUESLIST As ArrayList
            Get
                Return m_VALUESLIST
            End Get
            Set(value As ArrayList)
                m_VALUESLIST = value
            End Set
        End Property
        Private m_VALUESLIST As ArrayList

        ''' <summary>
        ''' MarathonTP message Timestamp for tracking and retries
        ''' </summary>
        Public Property STAMP As Date
            Get
                Return m_STAMP
            End Get
            Set(value As Date)
                m_STAMP = value
            End Set
        End Property
        Private m_STAMP As Date

        Public Sub New()
            VER = Common.Versioning(Common.MarathonVersion.V_1_1)
            RA = Common.RequestCode
            VALUESLIST = New ArrayList
            CMD = CByte(Common.MessageCodes.ReadVariables)
        End Sub
    End Class

    ''' <summary>
    ''' Value writing request definition.
    ''' </summary>
    Public Class WriteVariablesPacketV1_1
        ''' <summary>
        ''' MarathonTP protocol version
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly VER As String

        ''' <summary>
        ''' MarathonTP command type
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly CMD As Byte

        ''' <summary>
        ''' MarathonTP message type code
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly RA As Char

        ''' <summary>
        ''' MarathonTP message transaction number
        ''' </summary>
        Public Property TNS As UShort
            Get
                Return m_TNS
            End Get
            Set(value As UShort)
                m_TNS = value
            End Set
        End Property
        Private m_TNS As UShort

        ''' <summary>
        ''' MarathonTP elements to write on
        ''' </summary>
        ''' <remarks>Must be a combination of an element (TAG or index) and a value.</remarks>
        Public Property VALUESLIST As ArrayList
            Get
                Return m_VALUESLIST
            End Get
            Set(value As ArrayList)
                m_VALUESLIST = value
            End Set
        End Property
        Private m_VALUESLIST As ArrayList


        Public Sub New()
            VER = Common.Versioning(Common.MarathonVersion.V_1_1)
            RA = Common.RequestCode
            CMD = CByte(Common.MessageCodes.WriteVariables)
            VALUESLIST = New ArrayList
        End Sub
    End Class

    ''' <summary>
    ''' Discover request definition.
    ''' </summary>
    Public Class DiscoverPacketV1_1

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly VER As String

        ''' <summary>
        ''' MarathonTP command type
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly CMD As Byte

        ''' <summary>
        ''' MarathonTP message type code
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly RA As Char

        ''' <summary>
        ''' MarathonTP message transaction number
        ''' </summary>
        Public Property TNS As UShort
            Get
                Return m_TNS
            End Get
            Set(value As UShort)
                m_TNS = value
            End Set
        End Property
        Private m_TNS As UShort

        ''' <summary>
        ''' MarathonTP element(s) to read
        ''' </summary>
        ''' <remarks>Can be TAG or index, but must be all the same method.</remarks>
        Public Property VALUESLIST As ArrayList
            Get
                Return m_VALUESLIST
            End Get
            Set(value As ArrayList)
                m_VALUESLIST = value
            End Set
        End Property
        Private m_VALUESLIST As ArrayList

        Public Sub New()
            VER = Common.Versioning(Common.MarathonVersion.V_1_1)
            RA = Common.RequestCode
            CMD = CByte(Common.MessageCodes.Discovery)
        End Sub
    End Class

End Namespace

Namespace Packet.Answer

    ''' <summary>
    ''' Value reading answer definition.
    ''' </summary>
    Public Class ReturnedVariablesPacketV1_1
        ''' <summary>
        ''' MarathonTP protocol version
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly VER As String

        ''' <summary>
        ''' MarathonTP command type
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly CMD As Byte

        ''' <summary>
        ''' MarathonTP message type code
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly RA As Char

        ''' <summary>
        ''' MarathonTP message transaction number
        ''' </summary>
        Public Property TNS As UShort
            Get
                Return m_TNS
            End Get
            Set(value As UShort)
                m_TNS = value
            End Set
        End Property
        Private m_TNS As UShort

        ''' <summary>
        ''' MarathonTP data value list
        ''' </summary>
        ''' <remarks>Represented by a combination of error code, datatype and value.</remarks>
        Public Property VALUESLIST As ArrayList
            Get
                Return m_VALUESLIST
            End Get
            Set(value As ArrayList)
                m_VALUESLIST = value
            End Set
        End Property
        Private m_VALUESLIST As ArrayList

        Public Sub New()
            VER = Common.Versioning(Common.MarathonVersion.V_1_1)
            RA = Common.AnswerCode
            CMD = CByte(Common.MessageCodes.ReadVariables)
            VALUESLIST = New ArrayList
        End Sub
    End Class

    ''' <summary>
    ''' Value writing answer definition.
    ''' </summary>
    Public Class ResponseGenericPacketV1_1

        ''' <summary>
        ''' MarathonTP protocol version
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly VER As String

        ''' <summary>
        ''' MarathonTP command type
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly CMD As Byte

        ''' <summary>
        ''' MarathonTP message type code
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly RA As Char

        ''' <summary>
        ''' MarathonTP message transaction number
        ''' </summary>
        Public Property TNS As UShort
            Get
                Return m_TNS
            End Get
            Set(value As UShort)
                m_TNS = value
            End Set
        End Property
        Private m_TNS As UShort

        ''' <summary>
        ''' MarathonTP error code list
        ''' </summary>
        Public Property CODES As ArrayList
            Get
                Return m_CODES
            End Get
            Set(value As ArrayList)
                m_CODES = value
            End Set
        End Property
        Private m_CODES As ArrayList

        Public Sub New()
            VER = Common.Versioning(Common.MarathonVersion.V_1_1)
            RA = Common.AnswerCode
            CMD = CByte(Common.MessageCodes.WriteVariables)
            CODES = New ArrayList
        End Sub
    End Class

    ''' <summary>
    ''' Discover request answer definition.
    ''' </summary>
    Public Class DiscoverResponsePacketV1_1

        ''' <summary>
        ''' MarathonTP protocol version
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly VER As String

        ''' <summary>
        ''' MarathonTP command type
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly CMD As Byte

        ''' <summary>
        ''' MarathonTP message type code
        ''' </summary>
        ''' <remarks></remarks>
        Public ReadOnly RA As Char

        ''' <summary>
        ''' MarathonTP message transaction number
        ''' </summary>
        Public Property TNS As UShort
            Get
                Return m_TNS
            End Get
            Set(value As UShort)
                m_TNS = value
            End Set
        End Property
        Private m_TNS As UShort

       ''' <summary>
        ''' MarathonTP data value list
        ''' </summary>
        ''' <remarks>Represented by a combination of error code, datatype and value.</remarks>
        Public Property VALUESLIST As ArrayList
            Get
                Return m_VALUESLIST
            End Get
            Set(value As ArrayList)
                m_VALUESLIST = value
            End Set
        End Property
        Private m_VALUESLIST As ArrayList

        Public Sub New()
            VER = Common.Versioning(Common.MarathonVersion.V_1_1)
            RA = Common.AnswerCode
            CMD = CByte(Common.MessageCodes.Discovery)
            VALUESLIST = New ArrayList
        End Sub
    End Class

End Namespace