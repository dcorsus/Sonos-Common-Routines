Imports System
Imports System.Text
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports Microsoft.VisualBasic
Imports System.Threading

Public Class AsynchronousClient

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

    ReadOnly Property LocalIPAddress As String
        Get
            LocalIPAddress = MyLocalIPAddress
        End Get
    End Property

    ReadOnly Property LocalIPPort As String
        Get
            LocalIPPort = MyLocalIPPort
        End Get
    End Property

    ReadOnly Property RemoteIPAddress As String
        Get
            RemoteIPAddress = MyRemoteIPAddress
        End Get
    End Property

    ReadOnly Property RemoteIPPort As String
        Get
            RemoteIPPort = MyRemoteIPPort
        End Get
    End Property

    Public Class StateObject
        ' State object for receiving data from remote device.
        ' Client socket.
        Public workSocket As Socket = Nothing
        ' Size of receive buffer.
        Public Const BufferSize As Integer = 9000
        ' Receive buffer.
        Public buffer(BufferSize) As Byte
        ' Received data string.
        'Public sb As New StringBuilder
    End Class 'StateObject

    Public Function ConnectSocket(Server As String, ipPort As String) As Socket
        ' Establish the remote endpoint for the socket.
        MyRemoteIPAddress = Server
        MyRemoteIPPort = ipPort
        ConnectSocket = Nothing
        MySocket = Nothing
        MySocketIsClosed = True
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("ConnectSocket called with ipAddress = " & Server & " and ipPort = " & ipPort, LogType.LOG_TYPE_INFO)

        Try
            Dim remoteEP As New IPEndPoint(IPAddress.Parse(Server), ipPort) 'ip_Address, ipPort)
            ' Create a TCP/IP socket.
            MySocket = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            ' Connect to the remote endpoint.
            MySocket.BeginConnect(remoteEP, New AsyncCallback(AddressOf ConnectCallback), MySocket)
            ' Wait for connect.
            connectDone.WaitOne() ' I do this in the plugin itself, based on MySocketIsClosed because this runs in its own tread
            Return MySocket
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in ConnectSocket with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            ConnectSocket = Nothing
        End Try
    End Function

    Public Sub CloseSocket()
        ' Release the socket.
        If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("CloseSocket called for IPAddress = " & MyRemoteIPAddress, LogType.LOG_TYPE_INFO)
        If MySocket Is Nothing Then Exit Sub
        MySocketIsClosed = True
        Try
            receiveDone.Set()
            MySocket.Shutdown(SocketShutdown.Both)
            MySocket.Close()
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in CloseSocket with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
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

    '
    Public Delegate Sub DataReceivedEventHandler(sender As Object, e As Byte())

    Public Event DataReceived As DataReceivedEventHandler

    Protected Overridable Sub OnReceive(e As Byte())
        Try
            RaiseEvent DataReceived(Me, e)
        Catch ex As Exception
            Log("Error in OnReceive with ipAddress = " & MyRemoteIPAddress & " with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Function Receive() As Boolean
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
            If PIDebuglevel > DebugLevel.dlEvents Then Log("Receive called and state = " & MyIAsyncResult.IsCompleted.ToString, LogType.LOG_TYPE_WARNING)
            Receive = True
        Catch ex As Exception
            If PIDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in Receive with ipAddress = " & MyRemoteIPAddress & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
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
                ReDim ByteA(bytesRead)
                Array.Copy(state.buffer, ByteA, bytesRead)
                OnReceive(ByteA)
                ByteA = Nothing
                ' Get the rest of the data.
                If MySocketIsClosed Or (client Is Nothing) Then
                    ' this could be if the OnReceive data forced the connection to close
                    MySocketIsClosed = True
                    receiveDone.Set()
                    Exit Sub
                End If
                MyIAsyncResult = client.BeginReceive(state.buffer, 0, StateObject.BufferSize, SocketFlags.None, New AsyncCallback(AddressOf ReceiveCallback), state)
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

    Public Function Send(ByVal inData As Byte()) As Boolean
        ' Convert the string data to byte data using ASCII encoding.
        If PIDebuglevel > DebugLevel.dlEvents Then Log("Send called with Data = " & Encoding.ASCII.GetString(inData, 0, UBound(inData)), LogType.LOG_TYPE_INFO)
        'If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Send called with Data = " & Encoding.ASCII.GetString(inData, 0, UBound(inData)), LogType.LOG_TYPE_INFO)
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


End Class

