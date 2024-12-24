Imports System.Text
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.NetworkInformation
Imports System.Xml
Imports System.Web

Partial Public Class HSPI

    Private MyUPnPListener As MyTcpListener = Nothing
    Private UPnPListenerPort As Integer = 0

    Public Sub MSearchEvent(MSearchInfo As String)
        'If UPnPDebuglevel > DebugLevel.dlOff Then Log("UPNPListener.MSearchEvent called with MSearchInfo = " & MSearchInfo, LogType.LOG_TYPE_WARNING)
        If UPnPDebuglevel > DebugLevel.dlEvents Then Log("UPNPListener.MSearchEvent called with MSearchInfo = " & MSearchInfo, LogType.LOG_TYPE_INFO, LogColorNavy)
        Dim Header As String = ParseUPNPMessage(MSearchInfo, "") ' calling this procedure with empty parameter returns the first line in the response
        Dim URI As String = ParseUPNPMessage(MSearchInfo, "location")
        Dim USN As String = ParseUPNPMessage(MSearchInfo, "usn")
        Dim CacheControl As String = ParseUPNPMessage(MSearchInfo, "cache-control")
        Dim NT As String = ParseUPNPMessage(MSearchInfo, "nt")                       ' Notification Target
        Dim NTType As NTtype = ParseNT(NT)
        Dim ST As String = ParseUPNPMessage(MSearchInfo, "st")                       ' Notification Target
        Dim NTS As String = ParseUPNPMessage(MSearchInfo, "nts")                       ' Search Target
        Dim Server As String = ParseUPNPMessage(MSearchInfo, "server")
        Dim BootID As String = ParseUPNPMessage(MSearchInfo, "bootid.upnp.org")
        Dim ConfigID As String = ParseUPNPMessage(MSearchInfo, "configid.upnp.org")
        Dim SearchPort As String = ParseUPNPMessage(MSearchInfo, "searchport.upnp.org")  ' port the device will respond to for Unicast M-Search
        Dim ReceiveEP As String = ParseUPNPMessage(MSearchInfo, "receiveep")
        Dim IpParts As String()
        IpParts = Split(ReceiveEP, ":")
        Dim ReceiveIP As String = ""
        Dim ReceivePort As String = ""
        If ReceiveEP <> "" Then
            Try
                IpParts = Split(ReceiveEP, ":")
                ReceiveIP = IpParts(0)
                ReceivePort = IpParts(1)
            Catch ex As Exception
            End Try
        End If

        Dim MAN As String = ParseUPNPMessage(MSearchInfo, "man")
        Dim postData As String = ""

        If Header = "M-SEARCH" And MAN = """ssdp:discover""" And ReceiveIP <> "" And ReceivePort <> "" Then ' added in version .29 because someone had an error Error in UPNPListener.MSearchEvent with Error = Conversion from string "" to type 'Integer' is not valid."
            If ST = "urn:schemas-upnp-org:device:basic:1" Or ST = "urn:Belkin:device:**" Or ST = "upnp:rootdevice" Or ST = "ssdp:all" Then
                ' this is our test. It is the echo looking for the Philips HUE, so let responds asif
                postData = "HTTP/1.1 200 OK" & vbCrLf
                postData &= "CACHE-CONTROL: max-age=86400" & vbCrLf
                'postData &= "DATE: " & DateAndTime.Now.ToString & vbCrLf
                postData &= "EXT:" & vbCrLf
                postData &= "LOCATION: http://" & PlugInIPAddress & ":" & UPnPListenerPort.ToString & "/MediaController/Echo/Amazon-Echo-HA-Bridge.xml" & vbCrLf
                postData &= "SERVER: Win/7 UPnP/1.1 Echo/1" & vbCrLf
                postData &= "ST: urn:schemas-upnp-org:device:basic:1" & vbCrLf
                postData &= "USN: uuid:88f6698f-2c83-4393-bd03-cd54a9f8596::urn:Belkin:device:*:*" & vbCrLf
                postData &= "OPT: ""http://schemas.upnp.org/upnp/1/0/""; ns=01" & vbCrLf
                postData &= "01-NLS: 88f6698f-2c83-4393-bd03-cd54a9f8596" & vbCrLf
                'postData &= "USN: uuid:Socket-1_0-221438K0100073::urn:Belkin:device:*:*" & vbCrLf
                'postData &= "BOOTID.UPNP.ORG:"
                'postData &= "CONFIGID.UPNP.ORG:"
                'postData &= "SEARCHPORT.UPNP.ORG:"
                postData &= vbCrLf

                Try
                    Dim udpClient As New UdpClient
                    Dim bytCommand As Byte() = New Byte() {}
                    udpClient.Connect(ReceiveIP, ReceivePort)
                    bytCommand = Encoding.UTF8.GetBytes(postData)
                    Dim pRet As Integer = udpClient.Send(bytCommand, bytCommand.Length)
                    If (UPnPDebuglevel > DebugLevel.dlErrorsOnly) And (pRet <> bytCommand.Length) Then Log("Error in UPNPListener.MSearchEvent transmitted a different length from transmit string  to " & ReceiveIP & ":" & ReceivePort.ToString & " with data = " & postData.ToString, LogType.LOG_TYPE_ERROR)
                    'If (UPnPDebuglevel > DebugLevel.dlOff) And (pRet = bytCommand.Length) Then Log("UPNPListener.MSearchEvent responded to " & ReceiveIP & ":" & ReceivePort.ToString & " with data = " & postData.ToString, LogType.LOG_TYPE_INFO, LogColorNavy)
                Catch ex As Exception
                    If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in UPNPListener.MSearchEvent with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try


            End If





        End If
    End Sub

    Private Sub StartUPnPListener()
        ' this procedure listens to the event Notifications on my own proprietary port 12291/12292
        If UPnPDebuglevel > DebugLevel.dlErrorsOnly Then Log("UPNPListener.StartUPnPListener called", LogType.LOG_TYPE_INFO, LogColorGreen)
        If MyUPnPListener Is Nothing Then
            Try
                MyUPnPListener = New MyTcpListener
                UPnPListenerPort = MyUPnPListener.Start(PlugInIPAddress)
                AddHandler MyUPnPListener.Connection, AddressOf HandleUPnPListenerEventConnection
                AddHandler MyUPnPListener.recOK, AddressOf HandleUPnPListenerEventReceive
                AddHandler MyUPnPListener.DataReceived, AddressOf HandleUPnPListenerDataReceive
                PostUPNPDeviceXMLFile(UPnPListenerPort)
            Catch ex As Exception
                If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in UPNPListener.StartUPnPListener for Instance = " & MainInstance & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                StopUPnPListener()
                'RestartListenerFlag = True
            End Try
        End If
    End Sub

    Private Sub StopUPnPListener()
        If UPnPDebuglevel > DebugLevel.dlErrorsOnly Then Log("UPNPListener.StopUPnPListener called for Instance = " & MainInstance, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            If MyUPnPListener IsNot Nothing Then
                Try
                    RemoveHandler MyUPnPListener.Connection, AddressOf HandleUPnPListenerEventConnection
                    RemoveHandler MyUPnPListener.recOK, AddressOf HandleUPnPListenerEventReceive
                    RemoveHandler MyUPnPListener.DataReceived, AddressOf HandleUPnPListenerDataReceive
                Catch ex As Exception
                End Try
                MyUPnPListener.Close()
            End If
        Catch ex As Exception
            If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in UPNPListener.StopUPnPListener for Instance = " & MainInstance & " stopping the EventListener with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MyUPnPListener = Nothing
    End Sub

    Public Sub HandleUPnPListenerEventConnection(Connection As Boolean)
        If UPnPDebuglevel > DebugLevel.dlOff Then Log("UPNPListener.HandleUPnPListenerEventConnection called with Connection = " & Connection, LogType.LOG_TYPE_INFO, LogColorGreen)
        If Connection Then
            'RestartListenerFlag = False
            'ListenerIsActive = True
        Else
            'RestartListenerFlag = True
            'ListenerIsActive = False
        End If
    End Sub

    Public Sub HandleUPnPListenerEventReceive(Receive As Boolean)
        If UPnPDebuglevel > DebugLevel.dlOff Then Log("UPNPListener.HandleUPnPListenerEventReceive called with Receive = " & Receive, LogType.LOG_TYPE_INFO, LogColorGreen)
    End Sub


    Public Sub HandleUPnPListenerDataReceive(Data As String)
        MyUPnPListener.TCPResponse = "HTTP/1.1 200 OK" & vbCrLf & "Connection: close" & vbCrLf & "Content-Length: 0" & vbCrLf & vbCrLf
        If UPnPDebuglevel > DebugLevel.dlOff Then Log("UPNPListener.HandleUPnPListenerDataReceive called with Length = " & Data.Length & " and Data = " & Data, LogType.LOG_TYPE_INFO, LogColorGreen)
        Try
            Dim HTTPCommand As String() = GetHTTPCommand(Data)
            If HTTPCommand IsNot Nothing Then
                If UBound(HTTPCommand) >= 0 Then
                    Select Case HTTPCommand(0)
                        Case "GET"
                            Dim GetURL As String = HTTPCommand(1)
                            MyUPnPListener.TCPResponse = TCPSendFile(GetURL, False)
                        Case "POST"
                        Case "HEAD"
                            Dim GetURL As String = HTTPCommand(1)
                            MyUPnPListener.TCPResponse = TCPSendFile(GetURL, True)
                        Case "PUT"
                            ' example of the request PUT /api/bB8QWJYkeFvnXiESrAaqO9RIaREfTLitW9XSbsjB/lights/5337df69-bbbb-4534-bb86-f0ef8145b45d/state HTTP/1.1 Host: 192.168.1.100:62222 Accept: */* Content-type: application/x-www-form-urlencoded Content-Length: 12 {"on": true}
                            ' example of response  [{"success":{"/lights/5337df69-bbbb-4534-bb86-f0ef8145b45d/state/on":true}}]
                            ' PUT /api/bB8QWJYkeFvnXiESrAaqO9RIaREfTLitW9XSbsjB/lights/3/state HTTP/1.1 Host: 192.168.1.100:53467 Accept: */* Content-type: application/x-www-form-urlencoded Content-Length: 22 {"on": true,"bri":127} 
                            MyUPnPListener.TCPResponse = SendTCPResponse("[{""success"":{""/lights/3/state/on"":true}}]")
                        Case "DELETE"
                        Case "TRACE"
                        Case "OPTIONS"
                        Case "CONNECT"
                        Case "PATCH"

                    End Select
                End If
            End If
        Catch ex As Exception
            If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in UPNPListener.HandleUPnPListenerDataReceive called with Data = " & Data & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        MyUPnPListener.sendDone.Set()
    End Sub


    Private Function TCPSendFile(ByVal FileName As String, ByVal headerOnly As Boolean) As String
        If UPnPDebuglevel > DebugLevel.dlOff Then Log("UPNPListener.TCPSendFile called with FileName = " & FileName, LogType.LOG_TYPE_INFO, LogColorGreen)
        TCPSendFile = ""
        Dim F As String = ""
        Dim iFileLen As Integer
        Dim st As New StringBuilder
        Dim tp As String = ""
        Dim ErrorCode As Integer

        If FileName = "/api/bB8QWJYkeFvnXiESrAaqO9RIaREfTLitW9XSbsjB" Then
            FileName = "/MediaController/Echo/devices.txt"
        ElseIf FileName = "/api/bB8QWJYkeFvnXiESrAaqO9RIaREfTLitW9XSbsjB/lights" Then
            FileName = "/MediaController/Echo/lights.txt"
        ElseIf FileName = "/api/bB8QWJYkeFvnXiESrAaqO9RIaREfTLitW9XSbsjB/lights/3" Then
            FileName = "/MediaController/Echo/light3.txt"
        End If

        ' /api/bB8QWJYkeFvnXiESrAaqO9RIaREfTLitW9XSbsjB/lights/3
        Try
            tp = Mid(F, F.LastIndexOf(".") + 1, F.Length - F.LastIndexOf(".") - 1)
        Catch ex As Exception

        End Try

        ' try to pick out /api/bB8QWJYkeFvnXiESrAaqO9RIaREfTLitW9XSbsjB
        Try

            F = CurrentAppPath & "\html" & FileName

            F = System.Web.HttpUtility.UrlDecode(F)
            'If InStr(1, F, "../") <> 0 And gWebAllowFileAccess = 0 Then
            'GoTo Error1
            'End If
            If Not My.Computer.FileSystem.FileExists(F) Then GoTo Error1
        Catch ex As Exception
            'WriteMon("ERROR", "InTCPSendFile: " & ex.Message)
            GoTo Error1
        End Try

        Dim bteFileContents() As Byte
        Try
            bteFileContents = My.Computer.FileSystem.ReadAllBytes(F)
        Catch ex As Exception
            ErrorCode = Err.Number
            GoTo Error1
        End Try

        If bteFileContents.Length < 1 Then GoTo Error1

        GoTo Continue_Here

Error1:
        Try
            'If gWebSvrLogErrorsEnabled Then
            'WriteMon("Error", "Web Server Error 404, cannot serve file: " & F)
            'End If
            st.Append("HTTP/1.0 404 OK" & vbCrLf)
            st.Append("Server: UPnPListener" & vbCrLf)
            st.Append("<HTML><h2>Bad Request: " & FileName & "</h2></HTML>" & vbCrLf)
            'st.Append(TextFileToString("tail.htm", False, ErrorCode))
            TCPSendFile = st.ToString & vbCrLf
        Catch
        End Try
        Exit Function

Continue_Here:
        st.Append("HTTP/1.0 200 OK" & vbCrLf)
        st.Append("Server: HomeSeer" & vbCrLf)

        Dim sTemp As String = FileName.ToLower
        If sTemp.Contains("/images/homeseer") Or sTemp.Contains("/images/cached") Then
            ' allow browser to cache the homeseer images folder and cached folder for third parties
            st.Append("Expires: Sun, 10 Jan 2040 19:20:00 GMT" + vbCrLf)
        End If

        sTemp = FileName.ToUpper
        If sTemp.Contains("GIF") Then
            st.Append("Content-Type: image/gif" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_PN=GIF_LRG;DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("JPG") Or sTemp.Contains("JPEG") Then
            st.Append("Content-Type: image/jpeg" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_PN=JPEG_LRG;DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("PNG") Then
            st.Append("Content-Type: image/png" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_PN=PNG_LRG;DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("HTM") Or sTemp.Contains("HTML") Then
            st.Append("Content-Type: text/html" & vbCrLf)
        ElseIf sTemp.Contains("WML") Then
            st.Append("Content-Type: text/vnd.wap.wml" & vbCrLf)
        ElseIf sTemp.Contains("WBMP") Then
            st.Append("Content-Type: image/vnd.wap.wbmp" & vbCrLf)
        ElseIf sTemp.Contains("WMLC") Or sTemp.Contains("WSC") Then
            st.Append("Content-Type: application/vnd.wap.wmlc" & vbCrLf)
        ElseIf sTemp.Contains("WMLS") Or sTemp.Contains("WMLSCRIPT") Or sTemp.Contains("WS") Then
            st.Append("Content-Type: text/vnd.wap.wmlscript" & vbCrLf)
        ElseIf sTemp.Contains("WAV") Then
            st.Append("Content-Type: audio/x-wav" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_OP=00;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("SWF") Then
            st.Append("Content-Type: application/x-shockwave-flash" & vbCrLf)
        ElseIf sTemp.Contains("SPF") Then
            st.Append("Content-Type: application/futuresplash" & vbCrLf)
        ElseIf sTemp.Contains("CSS") Then
            st.Append("Content-Type: text/css" & vbCrLf)
        ElseIf sTemp.Contains("MP3") Then
            st.Append("Content-Type: audio/mpeg" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_PN=MP3;DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("MPG") Then
            st.Append("Content-Type: video/mpeg" & vbCrLf)
            st.Append("contentFeatures.dlna.org: :DLNA.ORG_PN=MPEG1;DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("AVI") Then
            st.Append("Content-Type: video/avi" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("MP4") Then
            st.Append("Content-Type: video/mp4" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("WMV") Then
            st.Append("Content-Type: video/x-ms-wmv" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_PN=WMVMED_BASE;DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("WMA") Then
            st.Append("Content-Type: audio/x-ms-wma" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_PN=WMABASE;DLNA.ORG_OP=01;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        ElseIf sTemp.Contains("OGG") Then
            st.Append("Content-Type: audio/ogg" & vbCrLf)
            st.Append("contentFeatures.dlna.org: DLNA.ORG_OP=00;DLNA.ORG_CI=1;DLNA.ORG_FLAGS=01500000000000000000000000000000" & vbCrLf)
            st.Append("transferMode.dlna.org: Streaming" & vbCrLf)

        Else
            st.Append("Content-Type: application/" & tp & vbCrLf)
        End If

        iFileLen = bteFileContents.Length
        st.Append("Accept-Ranges: bytes" & vbCrLf)
        st.Append("Content-Length: " & iFileLen.ToString & vbCrLf & vbCrLf)

        'WriteMon("SERVER DEBUG", "Sending file " & F & " size: " & Str(i))

        ' send header
        'State.disconnect += 1
        'Send(State, st.ToString)

        If Not headerOnly Then
            If iFileLen <= 32768 Then
                st.Append(System.Text.Encoding.UTF8.GetString(bteFileContents))
                'Send(State, bteFileContents, iFileLen)
            Else
                'Dim iPtr As Integer
                'Dim Amt As Integer
                'Dim bteSend() As Byte
                'For iPtr = 0 To (iFileLen - 1) Step 32768
                'If (iPtr + 32768) < (iFileLen - 1) Then
                'State.disconnect += 1
                'Amt = 32768
                'Else
                'Amt = iFileLen - iPtr
                'End If
                'ReDim bteSend(Amt - 1)
                'Array.Copy(bteFileContents, iPtr, bteSend, 0, Amt)
                'Send(State, bteSend, Amt)
                'Next
            End If
        Else

        End If
        TCPSendFile = st.ToString
        'WriteMon("SERVER DEBUG", "Done sending file " & F)


    End Function

    Private Function SendTCPResponse(ByVal Response As String) As String
        If UPnPDebuglevel > DebugLevel.dlOff Then Log("UPNPListener.SendTCPResponse called with Response = " & Response, LogType.LOG_TYPE_INFO, LogColorGreen)
        SendTCPResponse = ""
        Dim st As New StringBuilder
        st.Append("HTTP/1.0 200 OK" & vbCrLf)
        st.Append("Server: UPnPListener" & vbCrLf)
        st.Append("Content-Type: application/json;charset=UTF-8" & vbCrLf)
        st.Append("Connection: close" & vbCrLf)
        st.Append("Content-Length: " & Response.Length.ToString & vbCrLf & vbCrLf)
        If Response <> "" Then
            st.Append(Response)
        End If
        SendTCPResponse = st.ToString
    End Function



    Private Sub PostUPNPDeviceXMLFile(ReceiverPort As Integer)
        Dim DeviceXML As New XmlDocument
        Try
            DeviceXML.Load("file://" & CurrentAppPath & "/html/MediaController/Echo/Amazon-Echo-HA-Bridge.xml")
        Catch ex As Exception
            If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in UPNPListener.PostUPNPDeviceXMLFile called with ReceiverPort = " & ReceiverPort.ToString & " failed to load XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        ' fix this node <URLBase>http://192.168.1.175:8080/</URLBase>
        Dim MyXMLNodeList As XmlNodeList = DeviceXML.GetElementsByTagName("URLBase")
        If MyXMLNodeList IsNot Nothing Then
            MyXMLNodeList.Item(0).InnerText = "http://" & PlugInIPAddress & ":" & ReceiverPort.ToString
        End If
        Try
            DeviceXML.Save(CurrentAppPath & "/html/MediaController/Echo/Amazon-Echo-HA-Bridge.xml")
        Catch ex As Exception
            If UPnPDebuglevel > DebugLevel.dlOff Then Log("Error in UPNPListener.PostUPNPDeviceXMLFile called with ReceiverPort = " & ReceiverPort.ToString & " failed to save XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function ParseUPNPMessage(response As String, SearchItem As String) As String
        ParseUPNPMessage = ""
        If SearchItem <> "" Then
            Dim FuncReturn As String = (From line In response.Split({vbCr(0), vbLf(0)})
                                        Where line.ToLowerInvariant().StartsWith(SearchItem & ":")
                                        Select (line.Substring(SearchItem.Length + 1).Trim())).FirstOrDefault
            If FuncReturn IsNot Nothing Then
                Return FuncReturn
            Else
                Return ""
            End If
        Else
            ' Return the first line, which should be M-Search, Notify, HTTP
            Dim Lines As String() = Split(response, {vbCr(0), vbLf(0)})
            If Lines IsNot Nothing Then
                If Lines(0).ToLowerInvariant().StartsWith("m-search * ") Then
                    Return "M-SEARCH"
                ElseIf Lines(0).ToLowerInvariant().StartsWith("http/") Then
                    Return "HTTP"
                ElseIf Lines(0).ToLowerInvariant().StartsWith("notify * ") Then
                    Return "NOTIFY"
                End If
            End If
        End If
    End Function

    Private Function ParseNT(inNT As String) As NTtype
        If UPnPDebuglevel > DebugLevel.dlEvents Then Log("MySSDP.ParseNT received inNT = " & inNT.ToString, LogType.LOG_TYPE_INFO, LogColorGreen)
        If inNT = "" Then Return NTtype.NTTypeUnknown
        Dim NTParts As String() = Split(inNT, ":")
        Try
            Select Case NTParts(0).ToUpper
                Case "UPNP"
                    If NTParts(1).ToUpper = "ROOTDEVICE" Then Return NTtype.NTTypeRootDevice
                Case "UUID"
                    Return NTtype.NTTypeDevice
                Case "URN"
                    Select Case NTParts(2).ToUpper
                        Case "DEVICE"
                            Return NTtype.NTTypeURNDevice
                        Case "SERVICE"
                            Return NTtype.NTTypeURNService
                    End Select
            End Select
        Catch ex As Exception
        End Try
        Return NTtype.NTTypeUnknown
    End Function

    Private Enum NTtype
        NTTypeRootDevice = 0
        NTTypeDevice = 1
        NTTypeURNDevice = 2
        NTTypeURNService = 3
        NTTypeUnknown = 4
    End Enum

End Class
