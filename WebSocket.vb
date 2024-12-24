Imports System.Net
Imports System.Net.Sockets
Imports System.Net.WebSockets
Imports System.Text
Imports System.Threading
Imports System.IO
Imports System.Net.Security
Imports System.Security.Authentication
Imports System.Security.Cryptography.X509Certificates


Public Class WebSocketClient

    ' ManualResetEvent instances signal completion.
    Public connectDone As New ManualResetEvent(False)
    Public sendDone As New ManualResetEvent(False)
    Public receiveDone As New ManualResetEvent(False)

    ' The response from the remote device.
    Public response As Boolean = False
    Private MySocket As Socket = Nothing
    Private Mystate As StateObject
    Private MyIAsyncResult As IAsyncResult
    Public MySocketIsClosed As Boolean = True
    Private MyReceiveCallBack As Object = Nothing
    Private MyRemoteIPAddress As String = ""
    Private MyRemoteIPPort As String = ""
    Private MyLocalIPAddress As String = ""
    Private MyLocalIPPort As String = ""
    Private MyRandomNumberGenerator As New Random()
    Private MyURL As String = ""
    Private MyWebSocketIsActive As Boolean = False
    Friend WithEvents SocketAliveTimer As Timers.Timer
    Private PongReceived As Boolean = False
    Private MySocketIsSSL As Boolean = False
    Private MySSLStream As SslStream = Nothing
    Private MySSLTCPClient As TcpClient = Nothing
    Private MySSLState As SSLStateObject
    Private MySSLCertificate As X509Certificate2 = Nothing

    Public Const OpcodeText = 1
    Public Const OpcodeBinary = 2
    Public Const OpcodeClose = 8
    Public Const OpcodePing = 9
    Public Const OpcodePong = 10

    Friend WithEvents MyTreatInDataTimer As Timers.Timer
    Private InDataStream As MemoryStream
    Private InDataStreamReadIndex As Integer = 0
    Private ReEntrancyFlag As Boolean = False


    Public Class StateObject
        ' State object for receiving data from remote device.
        ' Client socket.
        Public workSocket As Socket = Nothing
        ' Size of receive buffer.
        Public Const BufferSize As Integer = 32000
        ' Receive buffer.
        Public buffer(BufferSize) As Byte
        ' Received data string.
        'Public sb As New StringBuilder
    End Class 'StateObject

    Public Class SSLStateObject
        ' State object for receiving data from remote device.
        ' Client socket.
        Public workSocket As SslStream = Nothing
        ' Size of receive buffer.
        Public Const BufferSize As Integer = 32000
        ' Receive buffer.
        Public buffer(BufferSize) As Byte
        ' Received data string.
        'Public sb As New StringBuilder
    End Class 'SSLStateObject


    Public Sub New(SSL As Boolean)
        MyBase.New()
        MySocketIsSSL = SSL
    End Sub

    Private Sub StartTreatInDataTimer()
        Try
            If MyTreatInDataTimer Is Nothing Then
                MyTreatInDataTimer = New Timers.Timer
            End If
            MyTreatInDataTimer.Stop()
            MyTreatInDataTimer.Interval = 100 ' 100 milliseconds
            MyTreatInDataTimer.AutoReset = False
            MyTreatInDataTimer.Enabled = True
            MyTreatInDataTimer.Start()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MyUPnPDevice.StartTreatInDataTimer for IPAddress = " & MyRemoteIPAddress & ". Unable to create the Event timer with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub StopTreatInDataTimer()
        If MyTreatInDataTimer Is Nothing Then Exit Sub
        Try
            MyTreatInDataTimer.Stop()
            MyTreatInDataTimer.Dispose()
            MyTreatInDataTimer = Nothing
        Catch ex As Exception
        End Try
    End Sub

    Private Sub MyTreatInDataTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyTreatInDataTimer.Elapsed
        If PIDebuglevel > DebugLevel.dlEvents Then Log("MyTreatInDataTimer_Elapsed called for IPAddress = " & MyRemoteIPAddress, LogType.LOG_TYPE_INFO)
        ReadFromInDataStream()
        e = Nothing
        sender = Nothing
    End Sub


    Public ReadOnly Property LocalIPAddress As String
        Get
            LocalIPAddress = MyLocalIPAddress
        End Get
    End Property

    Public ReadOnly Property LocalIPPort As String
        Get
            LocalIPPort = MyLocalIPPort
        End Get
    End Property

    Public ReadOnly Property RemoteIPAddress As String
        Get
            RemoteIPAddress = MyRemoteIPAddress
        End Get
    End Property

    Public ReadOnly Property RemoteIPPort As String
        Get
            RemoteIPPort = MyRemoteIPPort
        End Get
    End Property

    Public ReadOnly Property WebSocketActive As Boolean
        Get
            Return MyWebSocketIsActive
        End Get
    End Property

    Public Function UpgradeWebSocket(URL As String, SecWebSocketkey As String, TimerValue As Integer, AddOrigin As Boolean) As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("UpgradeWebSocket called with ipAddress = " & MyRemoteIPAddress & " and URL = " & URL, LogType.LOG_TYPE_INFO)
        MyWebSocketIsActive = False ' added 12/2/2018 v .0.39 If I turn my LG TV off, this is not being reset. Putting it here will enforce it to be set properly all the time
        Dim SocketDataString As String = ""
        ' shoot removing host port causes issues for LG and so is Origin
        If AddOrigin Then
            SocketDataString = "GET /" & URL & " HTTP/1.1" & vbCrLf & "Connection: Upgrade" & vbCrLf & "Upgrade: websocket" & vbCrLf & "Sec-WebSocket-Key: " & SecWebSocketkey & vbCrLf & "Sec-WebSocket-Version: 13" & vbCrLf & "Sec-WebSocket-Protocol: dumb-increment-protocol" & vbCrLf & "Sec-WebSocket-Extensions: deflate-frame" & vbCrLf & "Host: " & MyRemoteIPAddress & ":" & MyRemoteIPPort & vbCrLf & "Cache-Control: no-cache" & vbCrLf & vbCrLf '"Origin: " & MyLocalIPAddress & vbCrLf & vbCrLf
        Else
            SocketDataString = "GET /" & URL & " HTTP/1.1" & vbCrLf & "Connection: Upgrade" & vbCrLf & "Upgrade: websocket" & vbCrLf & "Sec-WebSocket-Key: " & SecWebSocketkey & vbCrLf & "Sec-WebSocket-Version: 13" & vbCrLf & "Sec-WebSocket-Protocol: dumb-increment-protocol" & vbCrLf & "Sec-WebSocket-Extensions: deflate-frame" & vbCrLf & "Host: " & MyRemoteIPAddress & ":" & MyRemoteIPPort & vbCrLf & vbCrLf '& "Origin: " & MyRemoteIPAddress & vbCrLf & vbCrLf
        End If
        ' SocketDataString = "GET /" & URL & " HTTP/1.1" & vbCrLf & "Connection: Upgrade" & vbCrLf & "Upgrade: websocket" & vbCrLf & "Sec-WebSocket-Key: " & SecWebSocketkey & vbCrLf & "Sec-WebSocket-Version: 13" & vbCrLf & "Sec-WebSocket-Protocol: dumb-increment-protocol" & vbCrLf & "Sec-WebSocket-Extensions: deflate-frame" & vbCrLf & "Host: " & MyRemoteIPAddress & ":" & MyRemoteIPPort & vbCrLf & vbCrLf '& "Origin: " & MyRemoteIPAddress & vbCrLf & vbCrLf

        Dim WaitForConnection As Integer = 0
        Do While MySocketIsClosed
            wait(1)
            WaitForConnection = WaitForConnection + 1
            If WaitForConnection > 10 Then
                ' unsuccesfull connection
                Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & ". Unable to open TCP connection within 10 seconds", LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return False
            End If
        Loop

        Try
            'Receive()  ' removed in v.039 and moved to where the socket is connected
        Catch ex As Exception
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
        response = False
        Try
            If Not Send(System.Text.ASCIIEncoding.ASCII.GetBytes(SocketDataString)) Then
                Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket", LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return False
            End If
        Catch ex As Exception
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Return False
        End Try
        Try
            sendDone.WaitOne()
        Catch ex As Exception
            Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Return False
        End Try
        WaitForConnection = 0
        Do While MySocketIsClosed
            wait(1)
            WaitForConnection = WaitForConnection + 1
            If WaitForConnection > 10 Then
                ' unsuccesfull connection
                Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & ". Unable to open TCP connection within 10 seconds", LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return False
            End If
        Loop

        If InDataStream Is Nothing Then
            Try
                InDataStream = New MemoryStream()
            Catch ex As Exception
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in UpgradeWebSocket for ipAddress = " & MyRemoteIPAddress & ". Unable to allocate MemoryStream with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If TimerValue <> 0 Then
            PongReceived = True
            SocketAliveTimer = New Timers.Timer
            SocketAliveTimer.Interval = TimerValue * 3000 ' 30 seconds
            SocketAliveTimer.AutoReset = True
            SocketAliveTimer.Enabled = True
            SendPing(True)
        End If

        Return True

    End Function

    Public Function ConnectSocket(Server As String, ipPort As String) As Boolean
        ' Establish the remote endpoint for the socket.
        If MySocketIsSSL Then
            Return ConnectSSLSocket(Server, ipPort)
        End If
        MyRemoteIPAddress = Server
        MyRemoteIPPort = ipPort
        MySocket = Nothing
        MySocketIsClosed = True
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ConnectSocket called with ipAddress = " & Server & " and ipPort = " & ipPort, LogType.LOG_TYPE_INFO)
        Try
            Dim remoteEP As New IPEndPoint(IPAddress.Parse(Server), ipPort)
            ' Create a TCP/IP socket.
            MySocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            ' Connect to the remote endpoint.
            MySocket.BeginConnect(remoteEP, New AsyncCallback(AddressOf ConnectCallback), MySocket)
            ' Wait for connect.
            connectDone.WaitOne() ' I do this in the plugin itself, based on MySocketIsClosed because this runs in its own tread
            Return True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ConnectSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            ConnectSocket = False
        End Try
    End Function

    Public Sub CloseSocket()
        ' Release the socket.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CloseSocket called for IPAddress = " & MyRemoteIPAddress, LogType.LOG_TYPE_INFO)
        If (MySocket Is Nothing) And (MySSLTCPClient Is Nothing) And (MySSLStream Is Nothing) Then Exit Sub
        MySocketIsClosed = True
        MyWebSocketIsActive = False
        Try
            If SocketAliveTimer IsNot Nothing Then
                SocketAliveTimer.Stop()
                SocketAliveTimer.Close()
                SocketAliveTimer.Dispose()
            End If
        Catch ex As Exception
        Finally
            SocketAliveTimer = Nothing
        End Try
        Try
            If InDataStream IsNot Nothing Then
                InDataStream.Close()
                InDataStream.Dispose()
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CloseSocket with ipAddress = " & MyRemoteIPAddress & " closing the memorystream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        Finally
            InDataStream = Nothing
        End Try
        Try
            If MySSLStream IsNot Nothing Then
                MySSLStream.Flush()
                MySSLStream.Close()
                MySSLStream.Dispose()
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CloseSocket with ipAddress = " & MyRemoteIPAddress & " closing SSLStream with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MySSLStream = Nothing
        If MySSLTCPClient IsNot Nothing Then
            Try
                MySSLTCPClient.Close()
            Catch ex As Exception
            Finally
                OnWebSocketClose()
            End Try
        End If
        MySSLTCPClient = Nothing

        If MySocket IsNot Nothing Then
            Try
                receiveDone.Set()
                MySocket.Shutdown(SocketShutdown.Both)
                MySocket.Close()
            Catch ex As Exception
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CloseSocket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Finally
                OnWebSocketClose()
            End Try

        End If
        MySocket = Nothing

    End Sub

    Private Sub ConnectCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        Dim client As Socket = CType(ar.AsyncState, Socket)
        ' Complete the connection.
        Try
            client.EndConnect(ar)
            MySocketIsClosed = False
            Dim MylocalEndPoint As System.Net.IPEndPoint
            MylocalEndPoint = client.LocalEndPoint
            MyLocalIPAddress = MylocalEndPoint.Address.ToString
            MyLocalIPPort = MylocalEndPoint.Port.ToString
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ConnectCallback connected a socket to " & client.RemoteEndPoint.ToString(), LogType.LOG_TYPE_INFO)
            connectDone.Set()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ConnectCallback calling EndConnect with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function ConnectSSLSocket(Server As String, ipPort As String) As Boolean
        MyRemoteIPAddress = Server
        MyRemoteIPPort = ipPort
        MySSLTCPClient = Nothing
        MySocketIsClosed = True
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ConnectSSLSocket called with ipAddress = " & Server & " and ipPort = " & ipPort, LogType.LOG_TYPE_INFO)
        Try
            Dim remoteEP As New IPEndPoint(IPAddress.Parse(Server), ipPort)
            ' Create a TCP/IP socket.
            MySSLTCPClient = New TcpClient()
            connectDone.Reset()
            MySSLTCPClient.BeginConnect(MyRemoteIPAddress, MyRemoteIPPort, New AsyncCallback(AddressOf TCPClientConnectCallback), MySSLTCPClient)
            ' Wait for connect.
            connectDone.WaitOne() ' I do this in the plugin itself, based on MySocketIsClosed because this runs in its own tread
            Return True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ConnectSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try
    End Function

    Private Sub TCPClientConnectCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        Dim SSLClient As TcpClient = CType(ar.AsyncState, TcpClient)
        ' Complete the connection.
        Try
            SSLClient.EndConnect(ar)
            Dim MylocalEndPoint As System.Net.IPEndPoint
            MylocalEndPoint = SSLClient.Client.LocalEndPoint
            MyLocalIPAddress = MylocalEndPoint.Address.ToString
            MyLocalIPPort = MylocalEndPoint.Port.ToString
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TCPClientConnectCallback connected a socket to " & SSLClient.Client.RemoteEndPoint.ToString(), LogType.LOG_TYPE_INFO)
            connectDone.Set()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TCPClientConnectCallback calling EndConnect with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            MySocketIsClosed = True
            If MySSLStream IsNot Nothing Then MySSLStream.Close()
            If MySSLStream IsNot Nothing Then MySSLStream.Dispose()
            MySSLStream = Nothing
            If MySSLTCPClient IsNot Nothing Then MySSLTCPClient.Close()
            MySSLTCPClient = Nothing
            connectDone.Set()
            Exit Sub
        End Try
        Try
            If Not AuthenticateSocket() Then
                MySocketIsClosed = True
                If MySSLStream IsNot Nothing Then MySSLStream.Close()
                If MySSLStream IsNot Nothing Then MySSLStream.Dispose()
                MySSLStream = Nothing
                If MySSLTCPClient IsNot Nothing Then MySSLTCPClient.Close()
                MySSLTCPClient = Nothing
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TCPClientConnectCallback calling Authenticate Socket with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            MySocketIsClosed = True
            Exit Sub
        End Try
        MySocketIsClosed = False
    End Sub


    Private Function AuthenticateSocket() As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("AuthenticateSocket called with ipAddress = " & MyRemoteIPAddress, LogType.LOG_TYPE_INFO)

        If MySSLTCPClient Is Nothing Then Return False
        If MySSLCertificate Is Nothing Then MySSLCertificate = New X509Certificate2()
        Try
            If MySSLStream IsNot Nothing Then
                MySSLStream.Close()
                MySSLStream.Dispose()
            End If
        Catch ex As Exception
        End Try
        MySSLStream = Nothing
        Try
            MySSLStream = New SslStream(MySSLTCPClient.GetStream(), False, (AddressOf ServerCertificateValidation), (AddressOf UserCertificateSelection))
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AuthenticateSocket setting up an SSLStream with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            Return False
        End Try

        ' The server name must match the name on the server certificate.
        Try
            MySSLStream.AuthenticateAsClient(MyRemoteIPAddress & ":" & MyRemoteIPPort)
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AuthenticateSocket authenticating the server with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Return False
        End Try

        Try
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then
                Try
                    Log("AuthenticateSocket received Cipher: " & MySSLStream.CipherAlgorithm.ToString & " strength = " & MySSLStream.CipherStrength.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    Log("AuthenticateSocket received HashAlgorithm Type: " & MySSLStream.HashAlgorithm.ToString & " strength = " & MySSLStream.HashStrength.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    Log("AuthenticateSocket received KeyExchange Type: {" & MySSLStream.KeyExchangeAlgorithm.ToString & "} strength = " & MySSLStream.KeyExchangeStrength.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    Log("AuthenticateSocket received SSL Protocol: " & MySSLStream.SslProtocol.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    Log("AuthenticateSocket received isAuthenticated: " & MySSLStream.IsAuthenticated.ToString & " as server = " & MySSLStream.IsServer.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    Log("AuthenticateSocket received isSigned: " & MySSLStream.IsSigned.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    Log("AuthenticateSocket received isEncrypted: " & MySSLStream.IsEncrypted.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                DisplayCertificateInformation(MySSLStream)
                Try
                    Log("AuthenticateSocket received can read: " & MySSLStream.CanRead.ToString & ", write = " & MySSLStream.CanWrite.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    Log("AuthenticateSocket received can timeout: " & MySSLStream.CanTimeout.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in AuthenticateSocket calling EndConnect with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        'Set timeouts for the read And write to 5 seconds.
        MySSLStream.ReadTimeout = 5000
        MySSLStream.WriteTimeout = 5000
        Return True
    End Function

    Private Function ServerCertificateValidation(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ServerCertificateValidation calling EndConnect with ipAddress = " & MyRemoteIPAddress & " and SSLPolicyErrors = " & sslPolicyErrors.ToString, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then
            If (certificate IsNot Nothing) Then
                Log("ServerCertificateValidation received cert was issued to " & certificate.Subject & " and is valid from " & certificate.GetEffectiveDateString() & " until " & certificate.GetExpirationDateString(), LogType.LOG_TYPE_INFO)
            Else
                Log("ServerCertificateValidation received certificate is null.", LogType.LOG_TYPE_INFO)
            End If
        End If
        Return True
    End Function

    Private Function UserCertificateSelection(sender As Object, targetHost As String, localCertificates As X509CertificateCollection, remoteCertificate As X509Certificate, acceptableIssuers As String()) As X509Certificate2
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("UserCertificateSelection calling EndConnect with ipAddress = " & MyRemoteIPAddress & " and Target Host = " & targetHost, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then
            If acceptableIssuers IsNot Nothing Then
                For Each accetableIssuer As String In acceptableIssuers
                    Log("UserCertificateSelection calling EndConnect and acceptable Issuer = " & accetableIssuer, LogType.LOG_TYPE_INFO)
                Next
            End If
            If localCertificates IsNot Nothing Then
                For Each localcertificate As X509Certificate In localCertificates
                    If (localcertificate IsNot Nothing) Then
                        Log("UserCertificateSelection received Local cert was issued to " & localcertificate.Subject & " and is valid from " & localcertificate.GetEffectiveDateString() & " until " & localcertificate.GetExpirationDateString(), LogType.LOG_TYPE_INFO)
                    Else
                        Log("UserCertificateSelection received Local certificate is null.", LogType.LOG_TYPE_INFO)
                    End If
                Next
            End If
            'Display the properties of the client's certificate.
            If (remoteCertificate IsNot Nothing) Then
                Log("UserCertificateSelection found Remote cert was issued to " & remoteCertificate.Subject & " and is valid from " & remoteCertificate.GetEffectiveDateString() & " until " & remoteCertificate.GetExpirationDateString(), LogType.LOG_TYPE_INFO)
            Else
                Log("UserCertificateSelection found Remote certificate is null.", LogType.LOG_TYPE_INFO)
            End If
        End If
        Return MySSLCertificate
    End Function

    Private Sub DisplayCertificateInformation(stream As SslStream)
        Try
            Log("DisplayCertificateInformation received Certificate revocation list checked: " & stream.CheckCertRevocationStatus, LogType.LOG_TYPE_INFO)
            Dim localCertificate As X509Certificate = stream.LocalCertificate
            If (stream.LocalCertificate IsNot Nothing) Then
                Log("DisplayCertificateInformation received Local cert was issued to " & localCertificate.Subject & " and is valid from " & localCertificate.GetEffectiveDateString() & " until " & localCertificate.GetExpirationDateString(), LogType.LOG_TYPE_INFO)
            Else
                Log("DisplayCertificateInformation received Local certificate is null.", LogType.LOG_TYPE_INFO)
            End If
            'Display the properties of the client's certificate.
            Dim remoteCertificate As X509Certificate = stream.RemoteCertificate
            If (stream.RemoteCertificate IsNot Nothing) Then
                Log("DisplayCertificateInformation received Remote cert was issued to " & remoteCertificate.Subject & " and is valid from " & remoteCertificate.GetEffectiveDateString() & " until " & remoteCertificate.GetExpirationDateString(), LogType.LOG_TYPE_INFO)
            Else
                Log("DisplayCertificateInformation received Remote certificate is null.", LogType.LOG_TYPE_INFO)
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in DisplayCertificateInformation with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Delegate Sub DataReceivedEventHandler(sender As Object, e As Byte())
    Public Event DataReceived As DataReceivedEventHandler

    Public Delegate Sub WebSocketClosedEventHandler(sender As Object)
    Public Event WebSocketClosed As WebSocketClosedEventHandler

    Protected Overridable Sub OnReceive(e As Byte())
        Try
            RaiseEvent DataReceived(Me, e)
        Catch ex As Exception
            Log("Error in OnReceive with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Protected Overridable Sub OnWebSocketClose()
        Try
            RaiseEvent WebSocketClosed(Me)
        Catch ex As Exception
            Log("Error in OnWebSocketClose with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function Receive() As Boolean
        If MySocketIsSSL Then Return ReceiveSSLSocket()
        Receive = False
        If MySocket Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Receive with ipAddress = " & MyRemoteIPAddress & ". No Socket", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            ' Create the state object.
            Mystate = New StateObject
            Mystate.workSocket = MySocket
            ' Begin receiving the data from the remote device.
            MyIAsyncResult = MySocket.BeginReceive(Mystate.buffer, 0, StateObject.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), Mystate)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Receive called and state = " & MyIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_INFO)
            Receive = True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Receive with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ReceiveSSLSocket() As Boolean
        ReceiveSSLSocket = False
        If MySSLStream Is Nothing Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ReceiveSSLSocket with ipAddress = " & MyRemoteIPAddress & ". No Socket", LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        Try
            ' Create the state object.
            MySSLState = New SSLStateObject
            MySSLState.workSocket = MySSLStream
            ' Begin receiving the data from the remote device.
            MyIAsyncResult = MySSLStream.BeginRead(MySSLState.buffer, 0, SSLStateObject.BufferSize, New AsyncCallback(AddressOf ReceiveCallbackSSLSocket), MySSLState)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ReceiveSSLSocket called and state = " & MyIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_WARNING)
            ReceiveSSLSocket = True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ReceiveSSLSocket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub ReceiveCallback(ByVal ar As IAsyncResult)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then log( "ReceiveCallback called")
        ' Retrieve the state object and the client socket 
        ' from the asynchronous state object.
        Try
            If MySocketIsClosed Then
                receiveDone.Set()
                Exit Sub
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ReceiveCallback closing socket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Dim state As StateObject = CType(ar.AsyncState, StateObject)
            Dim client As Socket = state.workSocket

            ' Read data from the remote device.
            Dim bytesRead As Integer = client.EndReceive(ar)

            If bytesRead > 0 Then
                ' There might be more data, so store the data received so far.
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReceiveCallback received data = " & Encoding.UTF8.GetString(state.buffer, 0, bytesRead), LogType.LOG_TYPE_INFO)
                'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReceiveCallback received data = " & Encoding.UTF8.GetString(state.buffer, 0, bytesRead), LogType.LOG_TYPE_INFO)
                response = True
                Dim ByteA As Byte()
                ReDim ByteA(bytesRead - 1)
                Array.Copy(state.buffer, ByteA, bytesRead)
                ByteA = TreatWebSocketData(ByteA)
                If ByteA IsNot Nothing Then OnReceive(ByteA)
                ByteA = Nothing
                ' Get the rest of the data.
                If MySocket IsNot Nothing Then MyIAsyncResult = client.BeginReceive(state.buffer, 0, StateObject.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), state)
            Else
                ' All the data has arrived; put it in response.
                response = True
                ' Signal that all bytes have been received. ' added connected state on 12/2/2018 
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReceiveCallback with ipAddress = " & MyRemoteIPAddress & " received all data and connected state = " & client.Connected.ToString, LogType.LOG_TYPE_WARNING)
                receiveDone.Set()
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ReceiveCallback with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub ReceiveCallbackSSLSocket(ByVal ar As IAsyncResult)
        ' Retrieve the state object and the client socket 
        ' from the asynchronous state object.
        Try
            If MySocketIsClosed Then
                receiveDone.Set()
                Exit Sub
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ReceiveCallbackSSLSocket closing socket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Dim SSLState As SSLStateObject = CType(ar.AsyncState, SSLStateObject)
            'MySSLState = CType(ar.AsyncState, SSLStateObject)
            Dim client As SslStream = SSLState.workSocket

            ' Read data from the remote device.
            Dim bytesRead As Integer = client.EndRead(ar)

            If bytesRead > 0 Then
                ' There might be more data, so store the data received so far.
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReceiveCallbackSSLSocket received data = " & Encoding.UTF8.GetString(SSLState.buffer, 0, bytesRead), LogType.LOG_TYPE_INFO)
                'If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReceiveCallbackSSLSocket received datalength = " & bytesRead.ToString, LogType.LOG_TYPE_INFO)
                response = True
                Dim ByteA As Byte()
                ReDim ByteA(bytesRead - 1)
                Array.Copy(SSLState.buffer, ByteA, bytesRead)
                ByteA = TreatWebSocketData(ByteA)
                If ByteA IsNot Nothing Then OnReceive(ByteA)
                ByteA = Nothing
                MyIAsyncResult = client.BeginRead(SSLState.buffer, 0, SSLStateObject.BufferSize, New AsyncCallback(AddressOf ReceiveCallbackSSLSocket), SSLState)
            Else
                ' All the data has arrived; put it in response.
                response = True
                ' Signal that all bytes have been received. ' added connected state on 12/2/2018 
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReceiveCallbackSSLSocket with ipAddress = " & MyRemoteIPAddress & " received all data", LogType.LOG_TYPE_WARNING)
                receiveDone.Set()
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ReceiveCallbackSSLSocket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function Send(ByVal inData As Byte()) As Boolean
        If MySocketIsSSL Then Return SendSSLSocket(inData)
        ' Convert the string data to byte data using ASCII encoding.
        If PIDebuglevel > DebugLevel.dlEvents Then Log("Send called with Data = " & Encoding.ASCII.GetString(inData, 0, inData.Length), LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Send called with Data = " & Encoding.ASCII.GetString(inData, 0, inData.Length), LogType.LOG_TYPE_INFO)
        Send = False
        Try
            If MySocket Is Nothing Then
                sendDone.Set()
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & ". No Socket", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MySocketIsClosed Then
                sendDone.Set()
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & ". Socket is closed", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' Begin sending the data to the remote device.
        Try
            MyIAsyncResult = MySocket.BeginSend(inData, 0, inData.Length, SocketFlags.None, New AsyncCallback(AddressOf SendCallback), MySocket)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Send called and state = " & MyIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_INFO)
            Send = True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Send with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SendSSLSocket(ByVal inData As Byte()) As Boolean
        ' Convert the string data to byte data using ASCII encoding.
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SendSSLSocket called with Data = " & Encoding.ASCII.GetString(inData, 0, UBound(inData)), LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendSSLSocket called with Data = " & Encoding.ASCII.GetString(inData, 0, UBound(inData)), LogType.LOG_TYPE_INFO)
        SendSSLSocket = False
        Try
            If MySSLStream Is Nothing Then
                sendDone.Set()
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendSSLSocket with ipAddress = " & MyRemoteIPAddress & ". No Socket", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendSSLSocket with ipAddress = " & MyRemoteIPAddress & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If MySocketIsClosed Then
                sendDone.Set()
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendSSLSocket with ipAddress = " & MyRemoteIPAddress & ". Socket is closed", LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendSSLSocket with ipAddress = " & MyRemoteIPAddress & " calling SendDone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        ' Begin sending the data to the remote device.
        Try
            MyIAsyncResult = MySSLStream.BeginWrite(inData, 0, inData.Length, New AsyncCallback(AddressOf SendCallbackSSLSocket), MySSLStream)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SendSSLSocket called and state = " & MyIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_INFO)
            SendSSLSocket = True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendSSLSocket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub SendCallback(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        If MySocketIsClosed Then
            sendDone.Set()
            Exit Sub
        End If
        Try
            Dim client As Socket = CType(ar.AsyncState, Socket)
            ' Complete sending the data to the remote device.
            Dim bytesSent As Integer = client.EndSend(ar)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SendCallback has sent " & bytesSent & " bytes to server.", LogType.LOG_TYPE_INFO)
            ' Signal that all bytes have been sent.
            sendDone.Set()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendCallback with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub SendCallbackSSLSocket(ByVal ar As IAsyncResult)
        ' Retrieve the socket from the state object.
        If MySocketIsClosed Then
            sendDone.Set()
            Exit Sub
        End If
        Try
            Dim SSLClient As SslStream = CType(ar.AsyncState, SslStream)
            ' Complete sending the data to the remote device.
            SSLClient.EndWrite(ar)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("SendCallbackSSLSocket has finished sending", LogType.LOG_TYPE_INFO)
            ' Signal that all bytes have been sent.
            sendDone.Set()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SendCallbackSSLSocket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub   '

    Public Function SendDataOverWebSocket(Opcode As Integer, SocketData As Byte(), UseMask As Boolean) As Boolean
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SendDataOverWebSocket called for ipAddress = " & MyRemoteIPAddress & " with DataLength = " & SocketData.Length & ", Opcode = " & Opcode.ToString & " and UseMask = " & UseMask.ToString, LogType.LOG_TYPE_INFO)
        SendDataOverWebSocket = False
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SendDataOverWebSocket for ipAddress - " & MyRemoteIPAddress & " will send data = " & Encoding.UTF8.GetString(SocketData, 0, SocketData.Length), LogType.LOG_TYPE_INFO)
        Dim GenerateANewIndex As Integer = MyRandomNumberGenerator.Next(1, 429496729)

        Dim Header(3) As Byte ' think of this as the header
        Dim PayLoad As Byte() = SocketData   ' and this the payload
        Header(0) = 128 + Opcode  ' FIN + Opcode = Text

        If SocketData.Length >= 126 Then
            Header(1) = 126 + 128 ' Mask set , 126 means next 2 bytes represent lenght (which must be >126) = FE Hex
            Header(2) = Int(SocketData.Length / 256)
            Header(3) = SocketData.Length Mod 256
        Else
            Header(1) = 128 + CByte(SocketData.Length) ' Mask set
        End If

        Dim startIndex As Integer = 4
        If SocketData.Length < 126 Then startIndex = 2

        If UseMask Then
            Dim Mask() As Byte = BitConverter.GetBytes(GenerateANewIndex)
            Array.Resize(Header, startIndex + PayLoad.Length + 4)
            Mask.CopyTo(Header, startIndex)
            For Index = 0 To UBound(PayLoad)
                PayLoad(Index) = PayLoad(Index) Xor Mask(Index Mod 4)
            Next
            startIndex += 4
        Else
            Array.Resize(Header, startIndex + PayLoad.Length)
        End If

        PayLoad.CopyTo(Header, startIndex)

        Try
            'If Not MySocketIsSSL Then Receive() ' need need to set to receive again when using SSLStream
            'Receive() ' removed in version .039 . We cannot call receive multiple times , it cause memory leak
        Catch ex As Exception
            Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to receive data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        response = False
        Try
            If Not Me.Send(Header) Then
                Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket", LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Exit Function
            End If
        Catch ex As Exception
            Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Exit Function
        End Try
        Try
            sendDone.WaitOne()
        Catch ex As Exception
            Log("Error in SendDataOverWebSocket for ipAddress = " & MyRemoteIPAddress & " unable to send data to Socket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CloseSocket()
            Exit Function
        End Try

        Return True
    End Function

    Private Function TreatWebSocketData(inData As Byte()) As Byte()

        TreatWebSocketData = Nothing
        If inData Is Nothing Then Return Nothing
        If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData called for ipAddress = " & MyRemoteIPAddress & " Datasize = " & inData.Length.ToString & " and Data = " & ASCIIEncoding.ASCII.GetString(inData), LogType.LOG_TYPE_INFO) ' dcorssl
        If inData.Length = 0 Then Return Nothing
        If Not MyWebSocketIsActive Then
            ' this is most likely the response to the websocket upgrade
            If ASCIIEncoding.ASCII.GetString(inData).IndexOf("101 Switching Protocols") <> -1 Then
                MyWebSocketIsActive = True
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatWebSocketData received Websocket upgrade for ipAddress = " & MyRemoteIPAddress & " and Header = " & ASCIIEncoding.ASCII.GetString(inData, 0, inData.Length), LogType.LOG_TYPE_INFO)
                Return Nothing
            End If
        End If

        'For i = 0 To inData.Length - 1
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("     data(" & i.ToString & ")" & inData(i).ToString, LogType.LOG_TYPE_INFO)
        'ReadFromInDataStream for ipAddressNext

        Try
            If InDataStream IsNot Nothing Then
                SyncLock (InDataStream)
                    InDataStream.Position = InDataStream.Length
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is going to write = " & inData.Length.ToString & " bytes and is now at position = " & InDataStream.Position.ToString, LogType.LOG_TYPE_INFO)
                    InDataStream.Write(inData, 0, inData.Length)
                    If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is ended writing and is now at position = " & InDataStream.Position.ToString, LogType.LOG_TYPE_INFO)
                End SyncLock
                StartTreatInDataTimer()
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in TreatWebSocketData writing data to stream for ipAddress = " & MyRemoteIPAddress & " and Header = " & ASCIIEncoding.ASCII.GetString(inData, 0, inData.Length) & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Return Nothing ' dcor to fix. remove code below if fifo mechanism is working

        Dim FIN As Boolean = False
        Dim Opcode As Integer = 0
        Dim InfoLength As Integer = 0
        Dim InfoStartOffset As Integer = 2
        Dim MaskBit As Boolean = False
        Dim Mask(3) As Byte
        Dim TextInformation As String = ""

        If inData(0) And 128 <> 0 Then
            ' FIN flag is set, no further 
            FIN = True
        End If
        Opcode = inData(0) And 15
        InfoLength = inData(1) And 127
        MaskBit = (inData(1) And 128) <> 0
        If InfoLength = 126 Then
            ' actually the next 2 bytes represent length
            InfoStartOffset = InfoStartOffset + 2
            InfoLength = inData(2) * 256 + inData(3)
        ElseIf InfoLength = 127 Then
            ' actually the next 8 bytes represent length
            InfoStartOffset = InfoStartOffset + 8
            InfoLength = inData(2) * 256 * 256 * 256 * 256 * 256 * 256 * 256 + inData(3) * 256 * 256 * 256 * 256 * 256 * 256 + inData(4) * 256 * 256 * 256 * 256 * 256 + inData(5) * 256 * 256 * 256 * 256 + inData(6) * 256 * 256 * 256 + inData(7) * 256 * 256 + inData(8) * 256 + inData(9)
        End If
        Dim DecodeBytes As Byte() = Nothing
        If MaskBit Then
            If UBound(inData) >= InfoStartOffset + 3 Then
                For i = 0 To 3
                    Mask(i) = inData(InfoStartOffset + i)
                Next
                InfoStartOffset = InfoStartOffset + 4
                If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Mask = " & Mask(0).ToString & " " & Mask(1).ToString & " " & Mask(2).ToString & " " & Mask(3).ToString, LogType.LOG_TYPE_INFO)
                ReDim DecodeBytes(UBound(inData) - InfoStartOffset)
                For Index = InfoStartOffset To UBound(inData)
                    DecodeBytes(Index - InfoStartOffset) = inData(Index) Xor Mask((Index - InfoStartOffset) Mod 4)
                Next
                If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " decoded info = " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            End If
        End If

        ' opcode
        '       *  %x0 denotes a continuation frame
        '       *  %x1 denotes a text frame
        '       *  %x2 denotes a binary frame
        '       *  %x3-7 are reserved for further non-control frames
        '       *  %x8 denotes a connection close
        '       *  %x9 denotes a ping
        '       *  %xA denotes a pong
        '       *  %xB-F are reserved for further control frames
        Dim OpCodeAsText As String
        If Opcode = OpcodeClose Then
            OpCodeAsText = "Close"
        ElseIf Opcode = OpcodePing Then
            OpCodeAsText = "Ping"
        ElseIf Opcode = OpcodePong Then
            OpCodeAsText = "Pong"
        ElseIf Opcode = OpcodeText Then
            OpCodeAsText = "Text"
        ElseIf Opcode = OpcodeBinary Then
            OpCodeAsText = "Binary"
        Else
            OpCodeAsText = Opcode.ToString
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received OpCode = " & OpCodeAsText & ", FIN = " & FIN.ToString & " and Length = " & InfoLength.ToString, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received maskbit = " & MaskBit.ToString, LogType.LOG_TYPE_INFO)
        If MaskBit Then
            If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for device - " & MyRemoteIPAddress & " received mask = " & ASCIIEncoding.ASCII.GetChars(Mask), LogType.LOG_TYPE_INFO)
        End If

        If UBound(inData) > 0 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(1) = " & inData(1).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 1 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(2) = " & inData(2).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 2 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(3) = " & inData(3).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 3 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Byte(4) = " & inData(4).ToString, LogType.LOG_TYPE_INFO)

        If Opcode = OpcodeClose Then
            ' close the connection
            If Not MaskBit Then
                TextInformation = ASCIIEncoding.ASCII.GetChars(inData, InfoStartOffset, inData.Length - InfoStartOffset - 1)
            Else
                If DecodeBytes IsNot Nothing Then TextInformation = ASCIIEncoding.ASCII.GetChars(DecodeBytes)
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " received close connection with Info = " & TextInformation, LogType.LOG_TYPE_WARNING)
            ' return a close 
            Try
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " is sending a close after receiving a close", LogType.LOG_TYPE_INFO)
                response = False
                If Not Send(inData) Then
                    Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a close", LogType.LOG_TYPE_ERROR)
                    CloseSocket()
                    Return Nothing
                End If
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a close with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return Nothing
            End Try
            Try
                sendDone.WaitOne()
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a close with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return Nothing
            End Try
            CloseSocket()
        ElseIf Opcode = OpcodePing Then ' Ping
            ' send a Pong
            Dim GenerateANewIndex As Integer = MyRandomNumberGenerator.Next(1, 429496729)
            ' use infolength and startinfoindex
            Dim Array1(1) As Byte
            Array1(0) = 138 '128 + 10 ' set FIN and pong opcode ;
            Array1(1) = 128 + InfoLength ' set mask bit and length

            ' it appears the MASK must be set
            Dim SendMask() As Byte = BitConverter.GetBytes(GenerateANewIndex)

            Array.Resize(Array1, 2 + 4) ' 4 is for the mask
            SendMask.CopyTo(Array1, 2) ' fixed 10/1/2018 v35

            If InfoLength > 0 Then
                ' copy the rest
                Array.Resize(Array1, 6 + InfoLength) ' 6 is header + mask
                Try
                    For Index = 0 To InfoLength - 1
                        Array1(Index + 6) = inData(Index + InfoStartOffset) Xor Mask(Index Mod 4)
                    Next
                Catch ex As Exception
                    Log("Error in TreatWebSocketData while preparing a pong for ipAddress = " & MyRemoteIPAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If

            Try
                If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " is sending a pong after receiving a ping", LogType.LOG_TYPE_INFO)
                response = False
                If Not Send(Array1) Then
                    Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a pong", LogType.LOG_TYPE_ERROR)
                    CloseSocket()
                    Return Nothing
                End If
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a pong with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                Return Nothing
            End Try
            Try
                sendDone.WaitOne()
            Catch ex As Exception
                Log("Error in TreatWebSocketData for ipAddress = " & MyRemoteIPAddress & " while sending a pong with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
            End Try
            Return Nothing
        ElseIf Opcode = OpcodePong Then ' Pong
            PongReceived = True
        ElseIf Opcode = OpcodeText Then ' TextFrame
            If Not MaskBit Then
                Array.Resize(DecodeBytes, inData.Length - InfoStartOffset - 1)
                Buffer.BlockCopy(inData, InfoStartOffset, DecodeBytes, 0, inData.Length - InfoStartOffset - 1)
                'TextInformation = ASCIIEncoding.ASCII.GetChars(inData, InfoStartOffset, inData.Length - InfoStartOffset - 1)
            End If
            If DecodeBytes IsNot Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Text Info =  " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received empty Text Info", LogType.LOG_TYPE_INFO)
            End If
            Return DecodeBytes
        ElseIf Opcode = OpcodeBinary Then ' BinaryFrame
            ' not sure the treatment here is different from a TextFrame
            If Not MaskBit Then
                Array.Resize(DecodeBytes, inData.Length - InfoStartOffset - 1)
                Buffer.BlockCopy(inData, InfoStartOffset, DecodeBytes, 0, inData.Length - InfoStartOffset - 1)
            End If
            If DecodeBytes IsNot Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received Binary Info =  " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("TreatWebSocketData for ipAddress - " & MyRemoteIPAddress & " received empty Binary Info", LogType.LOG_TYPE_INFO)
            End If
            Return DecodeBytes
        End If
        Return Nothing
    End Function

    Private Sub ReadFromInDataStream()

        If InDataStream Is Nothing Then
            StopTreatInDataTimer()
            Exit Sub
        End If

        If InDataStream.Length = 0 Then
            StopTreatInDataTimer()
            Exit Sub
        End If

        If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has " & InDataStream.Length.ToString & " bytes to process and is at Index = " & InDataStreamReadIndex.ToString, LogType.LOG_TYPE_INFO)
        If ReEntrancyFlag Then
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has reentracy", LogType.LOG_TYPE_WARNING)
            StartTreatInDataTimer() ' reset the timer
            Exit Sub
        End If



        ' read up on Websocket protocols here https://tools.ietf.org/html/rfc6455
        ' https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API/Writing_WebSocket_servers

        '     0                   1                   2                   3
        '     0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1 2 3 4 5 6 7 8 9 0 1
        '     +-+-+-+-+-------+-+-------------+-------------------------------+
        '     |F|R|R|R| opcode|M| Payload len |    Extended payload length    |
        '     |I|S|S|S|  (4)  |A|     (7)     |             (16/64)           |
        '     |N|V|V|V|       |S|             |   (if payload len==126/127)   |
        '     | |1|2|3|       |K|             |                               |
        '     +-+-+-+-+-------+-+-------------+ - - - - - - - - - - - - - - - +
        '     |     Extended payload length continued, if payload len == 127  |
        '     + - - - - - - - - - - - - - - - +-------------------------------+
        '     |                               |Masking-key, if MASK set to 1  |
        '     +-------------------------------+-------------------------------+
        '     | Masking-key (continued)       |          Payload Data         |
        '     +-------------------------------- - - - - - - - - - - - - - - - +
        '        Payload data continued ...                : 
        '     + - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - +
        '     |                     Payload Data continued ...                |
        '     +---------------------------------------------------------------+
        ' 
        ' Here are response results
        ' 81 7e  00 ff 7b 22 65 76
        '	           {  "
        '1000 0001 0111 1110 0000 0000 1111 1111
        'FIN  opcode text
        '          mask = 0 should be set to 1 when sent from client to server 	 	
        '           payload 7e = 126 which means payload Is more Then 125 bytes And Next two bytes are used As an unsigned Integer representing payload lenth
        '               extended payload lenth = 127

        ' when a connection is forced closed
        ' 88 02 03 EA 00
        ' 88 = FIN mask and Opcode = connection closed
        ' 02  = Mask 0; lenght = 2
        ' 03EA = 1002 decimal = protocol error 
        ' 00 I guess end 

        ' opcode
        '       *  %x0 denotes a continuation frame
        '       *  %x1 denotes a text frame
        '       *  %x2 denotes a binary frame
        '       *  %x3-7 are reserved for further non-control frames
        '       *  %x8 denotes a connection close
        '       *  %x9 denotes a ping
        '       *  %xA denotes a pong
        '       *  %xB-F are reserved for further control frames


        ' If InDataStream.Length < 2 Then Exit Sub ' not sure this is good 

        ReEntrancyFlag = True

        Dim inData As Byte()
        ReDim inData(1)
        Dim TempPosition = 0
        InDataStream.Seek(InDataStreamReadIndex, SeekOrigin.Begin)
        Dim DataRead As Integer = InDataStream.Read(inData, 0, 2)

        If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has read = " & DataRead.ToString & " bytes from Index = " & InDataStreamReadIndex.ToString & " and is now at position = " & InDataStream.Position.ToString, LogType.LOG_TYPE_INFO)

        Dim FIN As Boolean = False
        Dim Opcode As Integer = 0
        Dim InfoLength As Integer = 0
        Dim InfoStartOffset As Integer = 2
        Dim MaskBit As Boolean = False
        Dim Mask(3) As Byte
        Dim TextInformation As String = ""

        If inData(0) And 128 <> 0 Then
            ' FIN flag is set, no further 
            FIN = True
        End If
        Opcode = inData(0) And 15
        InfoLength = inData(1) And 127
        MaskBit = (inData(1) And 128) <> 0
        If InfoLength = 126 Then
            ' actually the next 2 bytes represent length
            InfoStartOffset = InfoStartOffset + 2
            ReDim inData(3)
            TempPosition = InDataStream.Position
            DataRead = InDataStream.Read(inData, 2, 2)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has read = " & DataRead.ToString & " bytes from Index = " & TempPosition.ToString & " and is now at position = " & InDataStream.Position.ToString, LogType.LOG_TYPE_INFO)
            InfoLength = inData(2) * 256 + inData(3)
        ElseIf InfoLength = 127 Then
            ' actually the next 8 bytes represent length
            InfoStartOffset = InfoStartOffset + 8
            ReDim inData(9)
            TempPosition = InDataStream.Position
            InDataStream.Read(inData, 2, 8)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has read = " & DataRead.ToString & " bytes from Index = " & TempPosition.ToString & " and is now at position = " & InDataStream.Position.ToString, LogType.LOG_TYPE_INFO)
            InfoLength = inData(2) * 256 * 256 * 256 * 256 * 256 * 256 * 256 + inData(3) * 256 * 256 * 256 * 256 * 256 * 256 + inData(4) * 256 * 256 * 256 * 256 * 256 + inData(5) * 256 * 256 * 256 * 256 + inData(6) * 256 * 256 * 256 + inData(7) * 256 * 256 + inData(8) * 256 + inData(9)
        End If

        If InDataStream.Length - InDataStreamReadIndex < (InfoStartOffset + InfoLength) Then
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has not enough data", LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has datastream length = " & InDataStream.Length.ToString & " and readIndex = " & InDataStreamReadIndex.ToString, LogType.LOG_TYPE_INFO)
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has Info length = " & InfoLength.ToString & " and infoStartOffset = " & InfoStartOffset.ToString, LogType.LOG_TYPE_INFO)
            ReEntrancyFlag = False
            StartTreatInDataTimer() ' make sure the timer is rearmed
            Exit Sub  ' not enough data in the buffer
        End If

        If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has infolength = " & InfoLength.ToString, LogType.LOG_TYPE_INFO)

        TempPosition = InDataStream.Position
        If MaskBit Then
            ReDim inData(InfoStartOffset + InfoLength + 4 - 1)
            DataRead = InDataStream.Read(inData, InfoStartOffset, InfoLength + 4)
        Else
            ReDim inData(InfoStartOffset + InfoLength - 1)
            DataRead = InDataStream.Read(inData, InfoStartOffset, InfoLength)
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " has read = " & DataRead.ToString & " bytes from Index = " & TempPosition.ToString & " and is now at position = " & InDataStream.Position.ToString, LogType.LOG_TYPE_INFO)

        Dim DecodeBytes As Byte() = Nothing
        If MaskBit Then
            If UBound(inData) >= InfoStartOffset + 3 Then
                For i = 0 To 3
                    Mask(i) = inData(InfoStartOffset + i)
                Next
                InfoStartOffset = InfoStartOffset + 4
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received Mask = " & Mask(0).ToString & " " & Mask(1).ToString & " " & Mask(2).ToString & " " & Mask(3).ToString, LogType.LOG_TYPE_INFO)
                ReDim DecodeBytes(UBound(inData) - InfoStartOffset)
                For Index = InfoStartOffset To UBound(inData)
                    DecodeBytes(Index - InfoStartOffset) = inData(Index) Xor Mask((Index - InfoStartOffset) Mod 4)
                Next
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " decoded info = " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            End If
        End If

        ' opcode
        '       *  %x0 denotes a continuation frame
        '       *  %x1 denotes a text frame
        '       *  %x2 denotes a binary frame
        '       *  %x3-7 are reserved for further non-control frames
        '       *  %x8 denotes a connection close
        '       *  %x9 denotes a ping
        '       *  %xA denotes a pong
        '       *  %xB-F are reserved for further control frames
        Dim OpCodeAsText As String
        If Opcode = OpcodeClose Then
            OpCodeAsText = "Close"
        ElseIf Opcode = OpcodePing Then
            OpCodeAsText = "Ping"
        ElseIf Opcode = OpcodePong Then
            OpCodeAsText = "Pong"
        ElseIf Opcode = OpcodeText Then
            OpCodeAsText = "Text"
        ElseIf Opcode = OpcodeBinary Then
            OpCodeAsText = "Binary"
        Else
            OpCodeAsText = Opcode.ToString
        End If
        If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received OpCode = " & OpCodeAsText & ", FIN = " & FIN.ToString & " and Length = " & InfoLength.ToString, LogType.LOG_TYPE_INFO)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received maskbit = " & MaskBit.ToString, LogType.LOG_TYPE_INFO)
        If MaskBit Then
            If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for device - " & MyRemoteIPAddress & " received mask = " & ASCIIEncoding.ASCII.GetChars(Mask), LogType.LOG_TYPE_INFO)
        End If

        If UBound(inData) > 0 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received Byte(1) = " & inData(1).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 1 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received Byte(2) = " & inData(2).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 2 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received Byte(3) = " & inData(3).ToString, LogType.LOG_TYPE_INFO)
        If UBound(inData) > 3 Then If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received Byte(4) = " & inData(4).ToString, LogType.LOG_TYPE_INFO)

        If Opcode = OpcodeClose Then
            ' close the connection
            If Not MaskBit Then
                TextInformation = ASCIIEncoding.ASCII.GetChars(inData, InfoStartOffset, inData.Length - InfoStartOffset - 1)
            Else
                If DecodeBytes IsNot Nothing Then TextInformation = ASCIIEncoding.ASCII.GetChars(DecodeBytes)
            End If
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " received close connection with Info = " & TextInformation, LogType.LOG_TYPE_WARNING)
            ' return a close 
            Try
                If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " is sending a close after receiving a close", LogType.LOG_TYPE_INFO)
                response = False
                If Not Send(inData) Then
                    Log("Error in ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " while sending a close", LogType.LOG_TYPE_ERROR)
                    CloseSocket()
                    ReEntrancyFlag = False
                    Exit Sub
                End If
            Catch ex As Exception
                Log("Error in ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " while sending a close with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                ReEntrancyFlag = False
                Exit Sub
            End Try
            Try
                sendDone.WaitOne()
            Catch ex As Exception
                Log("Error in ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " while sending a close with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                ReEntrancyFlag = False
                Exit Sub
            End Try
            CloseSocket()
        ElseIf Opcode = OpcodePing Then ' Ping
            ' send a Pong
            Dim GenerateANewIndex As Integer = MyRandomNumberGenerator.Next(1, 429496729)
            ' use infolength and startinfoindex
            Dim Array1(1) As Byte
            Array1(0) = 138 '128 + 10 ' set FIN and pong opcode ;
            Array1(1) = 128 + InfoLength ' set mask bit and length

            ' it appears the MASK must be set
            Dim SendMask() As Byte = BitConverter.GetBytes(GenerateANewIndex)

            Array.Resize(Array1, 2 + 4) ' 4 is for the mask
            SendMask.CopyTo(Array1, 2) ' fixed 10/1/2018 v35

            If InfoLength > 0 Then
                ' copy the rest
                Array.Resize(Array1, 6 + InfoLength) ' 6 is header + mask
                Try
                    For Index = 0 To InfoLength - 1
                        Array1(Index + 6) = inData(Index + InfoStartOffset) Xor Mask(Index Mod 4)
                    Next
                Catch ex As Exception
                    Log("Error in ReadFromInDataStream while preparing a pong for ipAddress = " & MyRemoteIPAddress & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If

            Try
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " is sending a pong after receiving a ping", LogType.LOG_TYPE_INFO)
                response = False
                If Not Send(Array1) Then
                    Log("Error in ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " while sending a pong", LogType.LOG_TYPE_ERROR)
                    CloseSocket()
                    ReEntrancyFlag = False
                    Exit Sub
                End If
            Catch ex As Exception
                Log("Error in ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " while sending a pong with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                ReEntrancyFlag = False
                Exit Sub
            End Try
            Try
                sendDone.WaitOne()
            Catch ex As Exception
                Log("Error in ReadFromInDataStream for ipAddress = " & MyRemoteIPAddress & " while sending a pong with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                CloseSocket()
                ReEntrancyFlag = False
                Exit Sub
            End Try
            MoreBytesInDataStream(inData.Length)
            ReEntrancyFlag = False
            Exit Sub
        ElseIf Opcode = OpcodePong Then ' Pong
            PongReceived = True
        ElseIf Opcode = OpcodeText Then ' TextFrame
            If Not MaskBit Then
                Array.Resize(DecodeBytes, inData.Length - InfoStartOffset)
                Buffer.BlockCopy(inData, InfoStartOffset, DecodeBytes, 0, inData.Length - InfoStartOffset)
                'TextInformation = ASCIIEncoding.ASCII.GetChars(inData, InfoStartOffset, inData.Length - InfoStartOffset - 1)
            End If
            If DecodeBytes IsNot Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received Text Info =  " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received empty Text Info", LogType.LOG_TYPE_INFO)
            End If
            MoreBytesInDataStream(inData.Length)
            OnReceive(DecodeBytes)
            ReEntrancyFlag = False
            Exit Sub
        ElseIf Opcode = OpcodeBinary Then ' BinaryFrame
            ' not sure the treatment here is different from a TextFrame
            If Not MaskBit Then
                Array.Resize(DecodeBytes, inData.Length - InfoStartOffset)
                Buffer.BlockCopy(inData, InfoStartOffset, DecodeBytes, 0, inData.Length - InfoStartOffset)
            End If
            If DecodeBytes IsNot Nothing Then
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received Binary Info =  " & ASCIIEncoding.ASCII.GetChars(DecodeBytes), LogType.LOG_TYPE_INFO)
            Else
                If PIDebuglevel > DebugLevel.dlEvents Then Log("ReadFromInDataStream for ipAddress - " & MyRemoteIPAddress & " received empty Binary Info", LogType.LOG_TYPE_INFO)
            End If
            MoreBytesInDataStream(inData.Length)
            OnReceive(DecodeBytes)
            ReEntrancyFlag = False
            Exit Sub
        End If
        MoreBytesInDataStream(inData.Length)
        ReEntrancyFlag = False
    End Sub

    Private Sub MoreBytesInDataStream(BytesTreated As Integer)
        InDataStreamReadIndex += BytesTreated
        If InDataStream Is Nothing Then
            InDataStreamReadIndex = 0
            Exit Sub
        End If
        Try
            If PIDebuglevel > DebugLevel.dlEvents Then Log("MoreBytesInDataStream for ipAddress - " & MyRemoteIPAddress & " has read = " & InDataStreamReadIndex.ToString & " bytes and has bytes in buffer = " & InDataStream.Length.ToString, LogType.LOG_TYPE_INFO)
            If InDataStream.Length = InDataStreamReadIndex Then
                ' all is read' flush the buffer
                InDataStream.Close()
                InDataStream.Dispose()
                InDataStream = New MemoryStream()
                InDataStreamReadIndex = 0
            Else
                StartTreatInDataTimer() ' make sure the timer is rearmed
            End If
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in MoreBytesInDataStream for ipAddress - " & MyRemoteIPAddress & " and Error= " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub SendPing(UseMask As Boolean)
        If PIDebuglevel > DebugLevel.dlEvents Then Log("SendPing called for ipAddress = " & MyRemoteIPAddress, LogType.LOG_TYPE_INFO)
        SendDataOverWebSocket(OpcodePing, ASCIIEncoding.ASCII.GetBytes("Hello"), UseMask)
    End Sub

    Private Sub SocketAliveTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles SocketAliveTimer.Elapsed
        ' we need to check if we received a response AND send out a new ping
        If Not PongReceived Then
            ' not good we need to release
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("SocketAliveTimer_Elapsed called for ipAddress = " & MyRemoteIPAddress & ". no Pong received in time. Closing socket", LogType.LOG_TYPE_WARNING)
            CloseSocket()
        Else
            PongReceived = False
            SendPing(True)
        End If
        e = Nothing
        sender = Nothing
    End Sub

    Public Sub ParseHTTPHeader(inHTTP As String)
        ' Return the first line, which should be M-Search, Notify, HTTP
        Dim Lines As String() = Split(inHTTP, {vbCr(0), vbLf(0)})
        For Each line As String In Lines
            Log("ParseHTTPHeader found entry = " & line, LogType.LOG_TYPE_INFO)
        Next
    End Sub

End Class
