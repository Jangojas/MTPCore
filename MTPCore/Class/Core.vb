Imports System.Text
Imports System.Threading

Public Class Core

#Region "Variables"

    ''' <summary>
    ''' Random number generator.
    ''' </summary>
    Private rnd As New Random

    ''' <summary>
    ''' Occurs when a value(s) writing is returned.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event WritingReturn(sender As Packet.Answer.ResponseGenericPacketV1_1, e As MessageEventArgs)

    ''' <summary>
    ''' Occurs when value(s) is returned.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event ReadingReturn(sender As Packet.Answer.ReturnedVariablesPacketV1_1, e As MessageEventArgs)

    ''' <summary>
    ''' Occurs when Discover is returned.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event DiscoverReturn(sender As Packet.Answer.DiscoverResponsePacketV1_1, e As MessageEventArgs)

    ''' <summary>
    ''' Occurs when a value(s) writing is requested.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event NewWritingReceive(sender As Packet.Request.WriteVariablesPacketV1_1, e As MessageEventArgs)

    ''' <summary>
    ''' Occurs when value(s) is requested.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event NewReadingReceive(sender As Packet.Request.ReadVariablesPacketV1_1, e As MessageEventArgs)

    ''' <summary>
    ''' Occurs when Discover is requested.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event NewDiscoverReceive(sender As Packet.Request.DiscoverPacketV1_1, e As MessageEventArgs)

    ''' <summary>
    ''' Occurs when Message Sending is requested.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event MessageSendingRequest(sender As Byte(), e As MessageEventArgs)

    ''' <summary>
    ''' Occurs when a Message Failure occurs.
    ''' </summary>
    ''' <param name="sender">Message transaction number</param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Event MessageFailed(sender As UShort, e As MessageEventArgs)

    ''' <summary>
    ''' Message recycling timer.
    ''' </summary>
    ''' <remarks></remarks>
    Private RecylingTimer As New Timer(New TimerCallback(AddressOf RecylingTimer_Tick), Nothing, 0, 10)

    ''' <summary>
    ''' Timer to evaluate protocol stats.
    ''' </summary>
    ''' <remarks></remarks>
    Private MessageStatTimer As New Timer(New TimerCallback(AddressOf MessageStatTimer_Tick), Nothing, 0, 1000)

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public DeviceList As List(Of Object)

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private DiscoveryNetworkList As List(Of Object)

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private SuccessfulMessageCountMemory As Integer

#End Region

#Region "Properties"

    ''' <summary>
    ''' Actual message transaction number.
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
    ''' Maximum number of device in device list. 1 is the minimum value. Negative number not allowed.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    Public Property MaxDeviceObjectCount As UShort
        Get
            Return m_MaxDeviceObjectCount
        End Get
        Set(value As UShort)
            If value > UShort.MaxValue Then
                m_MaxDeviceObjectCount = UShort.MaxValue
            ElseIf value < 1 Then
                m_MaxDeviceObjectCount = 1
            Else : m_MaxDeviceObjectCount = value
            End If
        End Set
    End Property
    Private m_MaxDeviceObjectCount As UShort

    ''' <summary>
    ''' Time before a message is considered lost. Expressed in milliseconds. This property is application wide. Valid range is 100 to 65535.
    ''' </summary>
    Public Property MessageTimeOut As UShort
        Get
            Return m_MessageTimeOut
        End Get
        Set(value As UShort)
            If value > UShort.MaxValue Then
                m_MessageTimeOut = UShort.MaxValue
            ElseIf value < 100 Then
                m_MessageTimeOut = 100
            Else : m_MessageTimeOut = value
            End If
        End Set
    End Property
    Private m_MessageTimeOut As UShort

    ''' <summary>
    ''' Number in milliseconds representing the absolute maximum time limit for issuing a message. Valid range is 0 to 65535.
    ''' </summary>
    Public Property MaxRetransmitInterval As UShort
        Get
            Return m_MaxRetransmitInterval
        End Get
        Set(value As UShort)
            If value > UShort.MaxValue Then
                m_MaxRetransmitInterval = UShort.MaxValue
            ElseIf value < 1000 Then
                m_MaxRetransmitInterval = 1000
            Else : m_MaxRetransmitInterval = value
            End If
        End Set
    End Property
    Private m_MaxRetransmitInterval As UShort

    ''' <summary>
    ''' Number of allowed retry. Valid range is 0 to 255.
    ''' </summary>
    Public Property MaxRetryAttempt As Byte
        Get
            Return m_MaxRetryAttempt
        End Get
        Set(value As Byte)
            If value > Byte.MaxValue Then
                m_MaxRetryAttempt = Byte.MaxValue
            ElseIf value < 0 Then
                m_MaxRetryAttempt = 0
            Else : m_MaxRetryAttempt = value
            End If
        End Set
    End Property
    Private m_MaxRetryAttempt As Byte

    ''' <summary>
    ''' Number in milliseconds representing the minimum time delay between two subsequent message send on the network. Valid range is 0 to 65535.
    ''' </summary>
    Public Property MinimumSendDelay As UShort
        Get
            Return m_MinimumSendDelay
        End Get
        Set(value As UShort)
            If value > UShort.MaxValue Then
                m_MinimumSendDelay = UShort.MaxValue
            ElseIf value < 0 Then
                m_MinimumSendDelay = 0
            Else : m_MinimumSendDelay = value
            End If
        End Set
    End Property
    Private m_MinimumSendDelay As UShort


    ''' <summary>
    ''' Keep count of received message. Keep count of every received message. Good and bad one. Automatically reset if value exceed 2,147,483,647.
    ''' </summary>
    Public Property ReceivedMessageCount As Integer
        Get
            Return m_ReceivedMessageCount
        End Get
        Set(value As Integer)
            If value > Integer.MaxValue Then
                m_ReceivedMessageCount = 0
            Else : m_ReceivedMessageCount = value
            End If
        End Set
    End Property
    Private m_ReceivedMessageCount As Integer

    ''' <summary>
    ''' Keep count of sended message. Count every message. With or without response. Automatically reset if value exceed 2,147,483,647.
    ''' </summary>
    Public Property SendedMessageCount As Integer
        Get
            Return m_SendedMessageCount
        End Get
        Set(value As Integer)
            If value > Integer.MaxValue Then
                m_SendedMessageCount = 0
            Else
                m_SendedMessageCount = value
            End If
        End Set
    End Property
    Private m_SendedMessageCount As Integer

    ''' <summary>
    ''' Keep count of failed message. A fail occur when no response is received for a specific message. Automatically reset if value exceed 2,147,483,647.
    ''' </summary>
    Public Property FailedMessageCount As Integer
        Get
            Return m_FailedMessageCount
        End Get
        Set(value As Integer)
            If value > Integer.MaxValue Then
                m_FailedMessageCount = 0
            Else : m_FailedMessageCount = value
            End If
        End Set
    End Property
    Private m_FailedMessageCount As Integer

    ''' <summary>
    ''' Keep count of retried message. A retry occur when ne response is received after message timeout. Automatically reset if value exceed 2,147,483,647.
    ''' </summary>
    Public Property RetriedMessageCount As Integer
        Get
            Return m_RetriedMessageCount
        End Get
        Set(value As Integer)
            If value > Integer.MaxValue Then
                m_RetriedMessageCount = 0
            Else : m_RetriedMessageCount = value
            End If
        End Set
    End Property
    Private m_RetriedMessageCount As Integer

    ''' <summary>
    ''' Used to keep tracking of successfully sended message.
    ''' </summary>
    Public Property SuccessfulMessageCount As Integer
        Get
            Return m_SuccessfulMessageCount
        End Get
        Set(value As Integer)
            If value > Integer.MaxValue Then
                m_SuccessfulMessageCount = 0
            Else : m_SuccessfulMessageCount = value
            End If
        End Set
    End Property
    Private m_SuccessfulMessageCount As Integer

    ''' <summary>
    ''' Number of successfull message transaction per second. This value is evaluated every second. Automatically reset if value exceed 2,147,483,647.
    ''' </summary>
    Public Property SuccessfulMessagePerSecond As Short
        Get
            Return m_SuccessfulMessagePerSecond
        End Get
        Set(value As Short)
            If value > Short.MaxValue Then
                m_SuccessfulMessagePerSecond = 0
            Else : m_SuccessfulMessagePerSecond = value
            End If
        End Set
    End Property
    Private m_SuccessfulMessagePerSecond As Short

    ''' <summary>
    ''' Used to store Device MarathonTP Security Mode.
    ''' </summary>
    Public Property DeviceSecurityModes As Common.SecurityModes
        Get
            Return m_DeviceSecurityModes
        End Get
        Set(value As Common.SecurityModes)
            If value > Integer.MaxValue Then
                m_DeviceSecurityModes = 0
            Else : m_DeviceSecurityModes = value
            End If
        End Set
    End Property
    Private m_DeviceSecurityModes As Common.SecurityModes

    ''' <summary>
    ''' Used to store Device MarathonTP Password.
    ''' </summary>
    Public Property DevicePassword As String
        Get
            Return m_DevicePassword
        End Get
        Set(value As String)
            m_DevicePassword = value
        End Set
    End Property
    Private m_DevicePassword As String

#End Region

#Region "Routines"

    Public Sub New()

        MaxRetransmitInterval = 30000
        MessageTimeOut = 3000
        TNS = CUShort((rnd.Next And &H7F) + 2) 'lowest 7 bits 
        DeviceList = New List(Of Object)
        DiscoveryNetworkList = New List(Of Object)
        DeviceSecurityModes = Common.SecurityModes.None
        DevicePassword = "DefaultPassword"
        MaxDeviceObjectCount = 1
        MaxRetryAttempt = 0

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="NetworkType"></param>
    ''' <param name="NetworkAddress"></param>
    ''' <param name="Password"></param>
    Public Sub AssignPasswordToDevice(NetworkType As Common.NetworkTypes, NetworkAddress As String, Password As String)
        Try
            'Check if device already exist in device list
            Dim actualDevice As DataTypes.DeviceObject = Nothing
            For Each device As DataTypes.DeviceObject In DeviceList
                If device.NetworkType = NetworkType And device.NetworkAddress = NetworkAddress Then
                    actualDevice = device
                    With actualDevice
                        .Password = Password
                    End With
                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="message"></param>
    ''' <param name="NetworkType"></param>
    ''' <param name="NetworkAddress"></param>
    ''' <param name="TNS"></param>
    ''' <param name="CMD"></param>
    Public Sub MessageAllocation(message As Byte(), NetworkType As Common.NetworkTypes, NetworkAddress As String, TNS As UShort, CMD As Byte)

        'If the message is a Discovery check the time elapsed since last Dicovery. 5 seconds is the minimum time.
        If CMD = 3 Then
                'Check if the discovery Object already exist
                Dim found As Boolean = False
                If DiscoveryNetworkList.Count > 0 Then
                    For Each discoveryObject As DataTypes.NetworkDiscoveryObject In DiscoveryNetworkList
                        If discoveryObject.NetworkDiscoveryAddress = NetworkAddress Then
                            Dim elapsed As TimeSpan = Date.Now - discoveryObject.LastDiscoveryTime
                            If ((elapsed.Seconds * 1000) + elapsed.Milliseconds) > 5000 Then
                                RaiseEvent MessageSendingRequest((message), New MessageEventArgs(NetworkType, NetworkAddress))
                                discoveryObject.LastDiscoveryTime = Date.Now
                            Else
                                Throw New Exception("Minimum time between Discovery not elapsed")
                                Exit Sub
                            End If
                            found = True
                            Exit For
                        End If
                    Next
                End If
                If found = False Or DiscoveryNetworkList.Count = 0 Then
                    Dim discoveryObject As New DataTypes.NetworkDiscoveryObject
                    discoveryObject.NetworkDiscoveryAddress = NetworkAddress
                    DiscoveryNetworkList.Add(discoveryObject)
                    RaiseEvent MessageSendingRequest((message), New MessageEventArgs(NetworkType, NetworkAddress))
                End If
                Exit Sub
            End If

            'Check if device already exist in device list
            Dim actualDevice As DataTypes.DeviceObject = Nothing
            For Each device As DataTypes.DeviceObject In DeviceList
                If device.NetworkType = NetworkType And device.NetworkAddress = NetworkAddress Then
                    actualDevice = device
                    Exit For
                End If
            Next

            If actualDevice Is Nothing Then
                'The device doesn't exist. Add device to device list.
                If DeviceList.Count < MaxDeviceObjectCount Then
                    actualDevice = New DataTypes.DeviceObject
                    With actualDevice
                        Dim newMTPMessage As New DataTypes.MTPMessage
                        With newMTPMessage
                            .Message = message
                            .InitialTimestamp = Date.Now
                            .LastTimestamp = .InitialTimestamp
                            .TNS = TNS
                            .Retries = 0
                            .TimeOut = MessageTimeOut
                        End With
                        .NetworkType = NetworkType
                        .NetworkAddress = NetworkAddress
                        .Password = DevicePassword
                        .PendingMessages(0) = newMTPMessage
                    End With
                    DeviceList.Add(actualDevice)
                Else
                    Throw New Exception("Device list already full")
                    Exit Sub
                End If
            Else
                'The Device already exist. Simply add message to list of pending message.
                Dim added As Boolean = False
                With actualDevice
                    For i = 0 To .PendingMessages.Length - 1
                        If .PendingMessages(i) Is Nothing Then
                            Dim newMTPMessage As New DataTypes.MTPMessage
                            With newMTPMessage
                                .Message = message
                                .InitialTimestamp = Date.Now
                                .LastTimestamp = .InitialTimestamp
                                .TNS = TNS
                                .Retries = 0
                                .TimeOut = MessageTimeOut
                            End With
                            .NetworkType = NetworkType
                            .NetworkAddress = NetworkAddress
                            .PendingMessages(i) = newMTPMessage
                            added = True
                            Exit For
                        End If
                    Next
                    If added = False Then
                        Throw New Exception("No free space for new message in Device Object")
                        Exit Sub
                    End If
                End With
            End If

    End Sub

    ''' <summary>
    ''' Evaluate a CRC for the byte array received on the socket. If the calculated CRC equals the CRC provided in the byte array then the message parser is called.
    ''' </summary>
    ''' <param name="data"></param>
    ''' <param name="length"></param>
    ''' <param name="NetworkType"></param>
    ''' <param name="NetworkAddress"></param>
    Public Sub OnDataReceived(data As Byte(), length As Integer, NetworkType As Common.NetworkTypes, NetworkAddress As String)
        'Check if device already exist in device list
        Dim actualDevice As DataTypes.DeviceObject = Nothing

        For Each device As DataTypes.DeviceObject In DeviceList
            If device.NetworkType = NetworkType And device.NetworkAddress = NetworkAddress Then
                actualDevice = device
                Exit For
            End If
        Next

        If actualDevice Is Nothing Then
            'The device doesn't exist. Add device to device list.
            If DeviceList.Count < MaxDeviceObjectCount Then
                actualDevice = New DataTypes.DeviceObject
                With actualDevice
                    .NetworkType = NetworkType
                    .NetworkAddress = NetworkAddress
                    .Password = "DevicePassword"
                End With
                DeviceList.Add(actualDevice)
            Else
                Throw New Exception("Device list already full")
                Exit Sub
            End If
        End If

        Dim dataNoCRC(length - 3) As Byte
        Dim dataCRC(1) As Byte
        Dim PayloadIndex As Integer
        dataCRC(0) = data(length - 2)
        dataCRC(1) = data(length - 1)
        Array.Copy(data, dataNoCRC, length - 2)
        Dim CheckSumResult As UShort = CUShort((CByte(dataCRC(0))) + (CByte(dataCRC(1))) * 256)

        'CRC validation
        If Common.CalculateCRC16(dataNoCRC) = CheckSumResult Then

            'Create a Char Array based on data buffer byte
            Dim c(dataNoCRC.Length - 1) As Char

            Dim count As Integer = 0
            For i = 0 To dataNoCRC.Length - 1
                c(i) = Strings.ChrW(dataNoCRC(i))
                If c(i) = Common.DIV Then
                    count += 1
                    If count = 4 Then
                        PayloadIndex = i + 1
                    End If
                End If
            Next

            'Treatment depending on MessageFrame
            If c(0) = Common.STX And c(c.Length - 1) = Common.ETX Then
                Dim s As String = New String(c)
                Dim trimmedMessage As String = Nothing

                trimmedMessage = s.Substring(1, s.Length - 2)

                'Find the header parameters
                Dim myParsedString() As String = trimmedMessage.Split({Common.DIV}, 5)

                'Evaluate the kind of message
                Select Case myParsedString(0)
                    Case Common.Versioning(Common.MarathonVersion.V_1_1)
                        Select Case myParsedString(1)
                            Case Common.RequestCode
                                If CInt(myParsedString(3)) = 1 Then 'Variable(s) reading
                                    Dim myParsedPayload() As String = ParsePayload(myParsedString(4), dataNoCRC, DeviceSecurityModes, PayloadIndex, DevicePassword)
                                    'For every pending items
                                    Dim ELE As New List(Of Object)
                                    For i = 0 To myParsedPayload.Length - 1
                                        Dim temp As String
                                        temp = myParsedPayload(i)
                                        ELE.Add(temp)
                                    Next

                                    Dim rMessage As New Packet.Request.ReadVariablesPacketV1_1
                                    With rMessage
                                        .TNS = CUShort(myParsedString(2))
                                        .VALUESLIST = ELE
                                    End With

                                    RaiseEvent NewReadingReceive(rMessage, New MessageEventArgs(NetworkType, NetworkAddress))

                                ElseIf CInt(myParsedString(3)) = 2 Then  'Variable(s) writing
                                    Dim myParsedPayload() As String = ParsePayload(myParsedString(4), dataNoCRC, DeviceSecurityModes, PayloadIndex, DevicePassword)
                                    'For every pending items
                                    Dim PAIR As New List(Of Object)
                                    For i = 0 To myParsedPayload.Length - 2
                                        Dim temp As Common.WritingPair
                                        'The first item is index. Second item is value
                                        temp.Index = myParsedPayload(i * 2)
                                        temp.Value = myParsedPayload((i * 2) + 1)
                                        PAIR.Add(temp)
                                    Next

                                    Dim sMessage As New Packet.Request.WriteVariablesPacketV1_1
                                    With sMessage
                                        .TNS = CUShort(myParsedString(2))
                                        .VALUESLIST = PAIR
                                    End With

                                    RaiseEvent NewWritingReceive(sMessage, New MessageEventArgs(NetworkType, NetworkAddress))

                                ElseIf CInt(myParsedString(3)) = 3 Then  'Discover Request
                                    Dim myParsedPayload() As String = ParsePayload(myParsedString(4), dataNoCRC, Common.SecurityModes.None, PayloadIndex, DevicePassword)
                                    'For every pending items
                                    Dim ELE As New List(Of Object)
                                    For i = 0 To myParsedPayload.Length - 1
                                        Dim temp As String
                                        temp = myParsedPayload(i)
                                        ELE.Add(temp)
                                    Next

                                    Dim sMessage As New Packet.Request.DiscoverPacketV1_1
                                    With sMessage
                                        .TNS = CUShort(myParsedString(2))
                                        .VALUESLIST = ELE
                                    End With

                                    RaiseEvent NewDiscoverReceive(sMessage, New MessageEventArgs(NetworkType, NetworkAddress))

                                End If

                            Case Common.AnswerCode
                                If CInt(myParsedString(3)) = 1 Then      'Return of variable(s) reading
                                    Dim myParsedPayload() As String = ParsePayload(myParsedString(4), dataNoCRC, DeviceSecurityModes, PayloadIndex, actualDevice.Password)
                                    Dim VALUES As New List(Of Object)
                                    For i = 0 To (myParsedPayload.Length / 3) - 1
                                        Dim temp As Common.ReturnedData
                                        'The firt item is error code. Second item is index. Third item is value
                                        temp.Code = CInt(myParsedPayload(CInt(3 * i)))
                                        temp.Type = myParsedPayload(CInt(1 + (3 * i)))
                                        temp.Value = myParsedPayload(CInt(2 + (3 * i)))
                                        VALUES.Add(temp)
                                    Next

                                    Dim rMessage As New Packet.Answer.ReturnedVariablesPacketV1_1
                                    With rMessage
                                        .TNS = CUShort(myParsedString(2))
                                        .VALUESLIST = VALUES
                                    End With

                                    'Response treatment
                                    For Each device As DataTypes.DeviceObject In DeviceList
                                        If device.NetworkAddress = NetworkAddress Then
                                            SyncLock device.PendingMessages
                                                For j = 0 To device.PendingMessages.Length - 1
                                                    If Not device.PendingMessages(j) Is Nothing Then
                                                        If device.PendingMessages(j).TNS = rMessage.TNS Then
                                                            RaiseEvent ReadingReturn(rMessage, New MessageEventArgs(NetworkType, NetworkAddress))
                                                            device.PendingMessages(j) = Nothing
                                                            SuccessfulMessageCount += CShort(1)
                                                        End If
                                                    End If
                                                Next
                                            End SyncLock
                                            Exit For
                                        End If
                                    Next

                                ElseIf CInt(myParsedString(3)) = 2 Then  'Return of  variable(s) writing
                                    Dim myParsedPayload() As String = ParsePayload(myParsedString(4), dataNoCRC, DeviceSecurityModes, PayloadIndex, actualDevice.Password)
                                    Dim CODES As New List(Of Object)
                                    'For every pending items
                                    For i = 0 To myParsedPayload.Length - 1
                                        Dim temp As String
                                        'The element is Error code.
                                        temp = myParsedPayload(i)
                                        CODES.Add(temp)
                                    Next

                                    Dim rMessage As New Packet.Answer.ResponseGenericPacketV1_1
                                    With rMessage
                                        .TNS = CUShort(myParsedString(2))
                                        .CODES = CODES
                                    End With

                                    'Response treatment
                                    For Each device As DataTypes.DeviceObject In DeviceList
                                        If device.NetworkAddress = NetworkAddress Then
                                            SyncLock device.PendingMessages
                                                For j = 0 To device.PendingMessages.Length - 1
                                                    If Not device.PendingMessages(j) Is Nothing Then
                                                        If device.PendingMessages(j).TNS = rMessage.TNS Then
                                                            RaiseEvent WritingReturn(rMessage, New MessageEventArgs(NetworkType, NetworkAddress))
                                                            device.PendingMessages(j) = Nothing
                                                            SuccessfulMessageCount += CShort(1)
                                                        End If
                                                    End If
                                                Next
                                            End SyncLock
                                            Exit For
                                        End If
                                    Next

                                ElseIf CInt(myParsedString(3)) = 3 Then 'Discover Response
                                    Dim myParsedPayload() As String = ParsePayload(myParsedString(4), dataNoCRC, Common.SecurityModes.None, PayloadIndex)
                                    Dim VALUES As New List(Of Object)
                                    For i = 0 To (myParsedPayload.Length / 3) - 1
                                        Dim temp As Common.ReturnedData
                                        'The firt item is error code. Second item is index. Third item is value
                                        temp.Code = CInt(myParsedPayload(CInt((3 * i))))
                                        temp.Type = myParsedPayload(CInt(1 + (3 * i)))
                                        temp.Value = myParsedPayload(CInt(2 + (3 * i)))
                                        VALUES.Add(temp)
                                    Next

                                    Dim rMessage As New Packet.Answer.DiscoverResponsePacketV1_1
                                    With rMessage
                                        .TNS = CUShort(myParsedString(2))
                                        .VALUESLIST = VALUES
                                    End With

                                    RaiseEvent DiscoverReturn(rMessage, New MessageEventArgs(NetworkType, NetworkAddress))

                                End If
                        End Select

                End Select

                ReceivedMessageCount += 1

            Else : Throw New Exception("Bad marathonTP message format")
                Exit Sub
            End If

        End If

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="StringToParse"></param>
    ''' <param name="NoCRCRawData"></param>
    ''' <param name="SecurityModes"></param>
    ''' <param name="PayloadIndex"></param>
    ''' <param name="Password"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ParsePayload(StringToParse As String, NoCRCRawData As Byte(), SecurityModes As Common.SecurityModes, PayloadIndex As Integer, Optional Password As String = "") As String()
        Dim s As String() = Nothing
        'Check if device is in encrypted mode
        Select Case SecurityModes
            Case Common.SecurityModes.None
                'Create Split with Payload
                s = StringToParse.Split(Common.DIV)

            Case Common.SecurityModes.XTEA
                Dim myXTEA As New Encryption.XTEA
                Dim b(StringToParse.Length - 1) As Byte
                For i = 0 To StringToParse.Length - 1
                    b(i) = NoCRCRawData(i + PayloadIndex)
                Next
                Dim resultdecrypt As String = myXTEA.Decrypt(32, b, Password)
                'Create Split with Payload
                s = resultdecrypt.Split(Common.DIV)
            Case Common.SecurityModes.Advanced
                Throw New Exception("Unsuported Encryption")
                s = Nothing
        End Select
        Return s
    End Function

    Private Sub MessageStatTimer_Tick()
        Try
            Dim result As Short = CShort(SuccessfulMessageCount - SuccessfulMessageCountMemory)
            If result >= 0 Then
                SuccessfulMessagePerSecond = result
            Else
                SuccessfulMessagePerSecond = CShort((Integer.MaxValue - SuccessfulMessageCountMemory) + SuccessfulMessageCount)
            End If
            SuccessfulMessageCountMemory = SuccessfulMessageCount
        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "Functions"

    ''' <summary>
    ''' Increments the value of the transaction number
    ''' </summary>
    ''' <remarks></remarks>
    Public Function IncrementTNS() As UShort
        If TNS < 65535 Then
            TNS += CUShort(1)
        Else
            TNS = 1
        End If
        Return TNS
    End Function

    Public Function MessageValidation(ByVal message As Object, encyptionMode As Inertia.MTPCore.Common.SecurityModes, Optional Password As String = "") As Byte()

        Dim MessageHeader As String = ""
        Dim Payload As String = ""
        Dim MessageBuffer As Byte() = Nothing

        Try
            Select Case True
                Case TypeOf message Is Packet.Request.ReadVariablesPacketV1_1
                    Dim obj As Packet.Request.ReadVariablesPacketV1_1 = DirectCast(message, Packet.Request.ReadVariablesPacketV1_1)
                    Dim lStringArrayBuilder As StringBuilder = New StringBuilder()
                    For Each item In obj.VALUESLIST
                        Dim temp As String = item.ToString
                        lStringArrayBuilder.Append(Common.DIV + temp)
                    Next
                    Payload = lStringArrayBuilder.ToString.TrimStart(Common.DIV)
                    With obj
                        MessageHeader = Common.STX + .VER + Common.DIV + .RA.ToString + Common.DIV + .TNS.ToString + Common.DIV + .CMD.ToString + Common.DIV
                    End With

                Case TypeOf message Is Packet.Request.WriteVariablesPacketV1_1
                    Dim obj As Packet.Request.WriteVariablesPacketV1_1 = DirectCast(message, Packet.Request.WriteVariablesPacketV1_1)
                    Dim lStringArrayBuilder As StringBuilder = New StringBuilder()
                    For Each item As Common.WritingPair In obj.VALUESLIST
                        lStringArrayBuilder.Append(Common.DIV + item.Index + Common.DIV + item.Value)
                    Next
                    Payload = lStringArrayBuilder.ToString.TrimStart(Common.DIV)
                    With obj
                        MessageHeader = Common.STX + .VER + Common.DIV + .RA.ToString + Common.DIV + .TNS.ToString + Common.DIV + .CMD.ToString + Common.DIV
                    End With

                Case TypeOf message Is Packet.Request.DiscoverPacketV1_1
                    Dim obj As Packet.Request.DiscoverPacketV1_1 = DirectCast(message, Packet.Request.DiscoverPacketV1_1)
                    Payload = "2" + Common.DIV + "3"
                    With obj
                        MessageHeader = Common.STX + .VER + Common.DIV + .RA.ToString + Common.DIV + .TNS.ToString + Common.DIV + .CMD.ToString + Common.DIV
                    End With

                Case TypeOf message Is Packet.Answer.ReturnedVariablesPacketV1_1
                    Dim obj As Packet.Answer.ReturnedVariablesPacketV1_1 = DirectCast(message, Packet.Answer.ReturnedVariablesPacketV1_1)
                    Dim lStringArrayBuilder As StringBuilder = New StringBuilder()
                    For Each item As Common.ReturnedData In obj.VALUESLIST
                        lStringArrayBuilder.Append(Common.DIV + CStr(item.Code) + Common.DIV + item.Type + Common.DIV + item.Value)
                    Next
                    Payload = lStringArrayBuilder.ToString.TrimStart(Common.DIV)
                    With obj
                        MessageHeader = Common.STX + .VER + Common.DIV + .RA.ToString + Common.DIV + .TNS.ToString + Common.DIV + .CMD.ToString + Common.DIV
                    End With

                Case TypeOf message Is Packet.Answer.ResponseGenericPacketV1_1
                    Dim obj As Packet.Answer.ResponseGenericPacketV1_1 = DirectCast(message, Packet.Answer.ResponseGenericPacketV1_1)
                    Dim lStringArrayBuilder As StringBuilder = New StringBuilder()
                    For Each item In obj.CODES
                        Dim temp As String = item.ToString
                        lStringArrayBuilder.Append(Common.DIV + temp)
                    Next
                    Payload = lStringArrayBuilder.ToString.TrimStart(Common.DIV)
                    With obj
                        MessageHeader = Common.STX + .VER + Common.DIV + .RA.ToString + Common.DIV + .TNS.ToString + Common.DIV + .CMD.ToString + Common.DIV
                    End With

                Case TypeOf message Is Packet.Answer.DiscoverResponsePacketV1_1
                    Dim obj As Packet.Answer.DiscoverResponsePacketV1_1 = DirectCast(message, Packet.Answer.DiscoverResponsePacketV1_1)
                    Dim lStringArrayBuilder As StringBuilder = New StringBuilder()
                    For Each item As Common.ReturnedData In obj.VALUESLIST
                        lStringArrayBuilder.Append(Common.DIV + CStr(item.Code) + Common.DIV + item.Type + Common.DIV + item.Value)
                    Next
                    Payload = lStringArrayBuilder.ToString.TrimStart(Common.DIV)
                    With obj
                        MessageHeader = Common.STX + .VER + Common.DIV + .RA.ToString + Common.DIV + .TNS.ToString + Common.DIV + .CMD.ToString + Common.DIV
                    End With

            End Select

            MessageBuffer = CreateMessageBuffer(MessageHeader, Payload, encyptionMode, Password)

        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try

        Return MessageBuffer

    End Function

    Private Function CreateMessageBuffer(MessageHeader As String, Payload As String, encyptionMode As Inertia.MTPCore.Common.SecurityModes, ByRef Password As String) As Byte()

        Dim CRCBuffer As Byte() = Nothing
        Dim OutBuffer As Byte() = Nothing
        Dim CheckSumCalc As UShort

        Select Case encyptionMode
            Case Common.SecurityModes.None

                CRCBuffer = Encoding.UTF8.GetBytes(MessageHeader + Payload + Common.ETX)

            Case Common.SecurityModes.XTEA
                Dim myXTEA As New Encryption.XTEA
                Dim HeaderBuffer As Byte() = Encoding.UTF8.GetBytes(MessageHeader)
                Dim XTEABuffer As Byte() = myXTEA.Encrypt(32, Payload, Password)
                Dim retBuffer(HeaderBuffer.Length + XTEABuffer.Length) As Byte
                HeaderBuffer.CopyTo(retBuffer, 0)
                XTEABuffer.CopyTo(retBuffer, HeaderBuffer.Length)
                retBuffer(retBuffer.Length - 1) = 125   'character "}"
                CRCBuffer = retBuffer

            Case Common.SecurityModes.Advanced
                Throw New Exception("Unsuported Encryption")

        End Select

        CheckSumCalc = Common.CalculateCRC16(CRCBuffer)
        Dim CRC1 As Byte = CByte((CheckSumCalc And 255))
        Dim CRC2 As Byte = CByte((CheckSumCalc \ CUShort(2 ^ 8)))

        Dim ByteToCopy As Integer = CRCBuffer.Length
        Dim MessageBuffer(CRCBuffer.Length + 1) As Byte
        Array.Copy(CRCBuffer, MessageBuffer, CRCBuffer.Length)

        MessageBuffer(ByteToCopy) = CRC1
        MessageBuffer(ByteToCopy + 1) = CRC2

        OutBuffer = MessageBuffer

        Return OutBuffer

    End Function

    Public Function GetIPBroadcastAddress(currentIP As String, ipNetMask As String) As String
        Dim strCurrentIP As String() = currentIP.Split("."c)
        Dim strIPNetMask As String() = ipNetMask.Split("."c)

        Dim arBroadCast As New List(Of Object)()

        For i As Integer = 0 To 3
            Dim nrBCOct As Integer = Integer.Parse(strCurrentIP(i)) Or (Integer.Parse(strIPNetMask(i)) Xor 255)
            arBroadCast.Add(nrBCOct.ToString())
        Next

        Dim broadcastaddress As String = CStr(arBroadCast(0)) + "." + CStr(arBroadCast(1)) + "." + CStr(arBroadCast(2)) + "." + CStr(arBroadCast(3))

        Return broadcastaddress

    End Function

    Public Function GetNetworkAddress(currentIP As String, ipNetMask As String) As String

        Dim networkaddress As String = "0.0.0.0"

        Try
            Dim strCurrentIP As String() = currentIP.Split("."c)
            Dim strIPNetMask As String() = ipNetMask.Split("."c)

            Dim arNetwork As New List(Of Object)()

            For i As Integer = 0 To 3
                Dim temp As String = strCurrentIP(i) And strIPNetMask(i)
                arNetwork.Add(temp)
            Next
            networkaddress = CStr(arNetwork(0)) + "." + CStr(arNetwork(1)) + "." + CStr(arNetwork(2)) + "." + CStr(arNetwork(3))

        Catch ex As Exception

        End Try

        Return networkaddress

    End Function


#End Region

#Region "Events"

    ''' <summary>
    ''' 
    ''' </summary>
    Private Sub RecylingTimer_Tick()

        Try
            If DeviceList.Count > 0 Then
                'For each device in Device List
                For i = 0 To DeviceList.Count - 1
                    Dim device As DataTypes.DeviceObject = DirectCast(DeviceList(i), DataTypes.DeviceObject)
                    'For each message in device message list
                    SyncLock device.PendingMessages
                        For j = 0 To device.PendingMessages.Length - 1
                            If Not device.PendingMessages(j) Is Nothing Then
                                With device.PendingMessages(j)
                                    If .Retries = 0 Then    'The first attempt must be sent
                                        RaiseEvent MessageSendingRequest((.Message), New MessageEventArgs(device.NetworkType, device.NetworkAddress))
                                        .Retries += CUShort(1)
                                    Else
                                        'Calculate elapsed time
                                        Dim elapsedLast As Integer = ((Date.Now - .LastTimestamp).Seconds * 1000) + (Date.Now - .LastTimestamp).Milliseconds
                                        Dim elapsedTotal As Integer = ((Date.Now - .InitialTimestamp).Seconds * 1000) + (Date.Now - .InitialTimestamp).Milliseconds

                                        If elapsedLast > .TimeOut Then
                                            If .Retries <= MaxRetryAttempt And elapsedTotal < MaxRetransmitInterval Then
                                                RaiseEvent MessageSendingRequest((.Message), New MessageEventArgs(device.NetworkType, device.NetworkAddress))
                                                .Retries += CUShort(1)
                                                .LastTimestamp = Date.Now
                                                .TimeOut = CUShort(device.PendingMessages(j).TimeOut * 2)   'Exponential Back-off
                                            Else
                                                FailedMessageCount += 1
                                                RaiseEvent MessageFailed(.TNS, New MessageEventArgs(device.NetworkType, device.NetworkAddress))
                                                device.PendingMessages(j) = Nothing
                                            End If
                                        End If
                                    End If
                                End With
                            End If

                        Next
                    End SyncLock
                Next
            End If

        Catch ex As Exception
            Debug.Write(ex.Message)
        End Try

    End Sub

#End Region

End Class

Public Class MessageEventArgs
    Inherits EventArgs

    Public NetworkType As Common.NetworkTypes
    Public NetworkAddress As String

    Public Sub New(NetworkType As Common.NetworkTypes, NetworkAddress As String)
        Me.NetworkType = NetworkType
        Me.NetworkAddress = NetworkAddress
    End Sub


End Class
