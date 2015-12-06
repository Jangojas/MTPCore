Imports System

Namespace DataTypes

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Class DeviceObject

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Password As String
            Get
                Return m_Password
            End Get
            Set(value As String)
                m_Password = value
            End Set
        End Property
        Private m_Password As String

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property NetworkType As Common.NetworkTypes
            Get
                Return m_NetworkType
            End Get
            Set(value As Common.NetworkTypes)
                m_NetworkType = value
            End Set
        End Property
        Private m_NetworkType As Common.NetworkTypes

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property NetworkAddress As String
            Get
                Return m_NetworkAddress
            End Get
            Set(value As String)
                m_NetworkAddress = value
            End Set
        End Property
        Private m_NetworkAddress As String

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <remarks></remarks>
        Public PendingMessages(0) As MTPMessage

    End Class

    Public Class MTPMessage
        Public Message As Byte()
        Public InitialTimestamp As Date
        Public LastTimestamp As Date
        Public TNS As UShort
        Public Retries As UShort
        Public TimeOut As UShort
    End Class

    Public Class NetworkDiscoveryObject
        Public NetworkDiscoveryAddress As String
        Public LastDiscoveryTime As Date

        Public Sub New()
            LastDiscoveryTime = Date.Now
        End Sub

    End Class


End Namespace
