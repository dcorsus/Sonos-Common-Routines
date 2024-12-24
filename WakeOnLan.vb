
Imports System.Net

Partial Public Class HSPI


    ' If IPv4 is used, the destination IP address MUST be the IP subnet broadcast address.
    ' If IPv6 is used, the destination IP address MUST be FF:2::1.

    Private Function GetIP(ByVal DNSName As String) As String
        Try
            Return Net.Dns.GetHostEntry(DNSName).AddressList.GetLowerBound(0).ToString
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Public Function GetSubnetMask(inIPAddress As String) As String
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSubnetMask called with IP Address = " & inIPAddress, LogType.LOG_TYPE_INFO)
        GetSubnetMask = ""
        If inIPAddress = "" Then Exit Function
        Dim searchIPAddress As IPAddress = Nothing
        Try
            System.Net.IPAddress.TryParse(inIPAddress, searchIPAddress)
        Catch ex As Exception
            Exit Function
        End Try
        Try
            For Each nic As System.Net.NetworkInformation.NetworkInterface In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
                If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("The MAC address of {0} is {1}{2}", nic.Description, Environment.NewLine, nic.GetPhysicalAddress()), LogType.LOG_TYPE_INFO)
                For Each Ipa In nic.GetIPProperties.UnicastAddresses
                    If piDebuglevel > DebugLevel.dlEvents Then Log(String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString), LogType.LOG_TYPE_INFO)
                    'If piDebuglevel > DebugLevel.dlErrorsOnly Then hs.writelog(IFACE_NAME, String.Format("The IPaddress address of {0} is {1}{2}", nic.Description, Environment.NewLine, Ipa.Address.ToString))
                    If Ipa.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                        ' IPV4. Mask both address with the network mask to make sure this ip Address is part of this NIC
                        Dim mask As Byte() = Ipa.IPv4Mask.GetAddressBytes
                        Dim nicIpAddress As Byte() = Ipa.Address.GetAddressBytes
                        Dim searchIpWMask As Byte() = searchIPAddress.GetAddressBytes
                        Dim equal As Boolean = True
                        For i = 0 To 3
                            If (nicIpAddress(i) And mask(i)) <> (searchIpWMask(i) And mask(i)) Then
                                equal = False
                                Exit For
                            End If
                        Next
                        If equal Then
                            ' OK we found our IPaddress
                            GetSubnetMask = Ipa.IPv4Mask.ToString
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSubnetMask found IP Mask = " & Ipa.IPv4Mask.ToString, LogType.LOG_TYPE_INFO)
                            Exit Function
                        End If
                    Else
                        If Ipa.Address.ToString = inIPAddress Then ' this would be IPV6, we don't support it so not sure what the code would be
                            ' OK we found our IPaddress
                            GetSubnetMask = Ipa.IPv4Mask.ToString
                            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("GetSubnetMask found IP Mask = " & Ipa.IPv4Mask.ToString, LogType.LOG_TYPE_INFO)
                            Exit Function
                        End If
                    End If

                Next
            Next
        Catch ex As Exception
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in GetSubnetMask, none found", LogType.LOG_TYPE_ERROR)
    End Function

    Private Sub SendMagicPacket(MacAddress As String, inIPAddress As String)
        If MacAddress = "" Then Exit Sub
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMagicPacket called for MacAddress = " & MacAddress, LogType.LOG_TYPE_INFO)
        Dim udpClient As System.Net.Sockets.UdpClient = Nothing
        Dim sendBytes As [Byte]() = Nothing
        Dim myAddress As String
        Dim Port As Integer = 9

        Try

            udpClient = New System.Net.Sockets.UdpClient(9, Net.Sockets.AddressFamily.InterNetwork)
            udpClient.EnableBroadcast = True
            Dim buf(101) As Char
            sendBytes = System.Text.Encoding.UTF8.GetBytes(buf)

            For x As Integer = 0 To 5
                sendBytes(x) = CByte(&HFF)
            Next

            MacAddress = MacAddress.Replace("-", "").Replace(":", "")

            Dim i As Integer = 6

            For x As Integer = 1 To 16
                sendBytes(i) = Convert.ToByte(MacAddress.Substring(0, 2), 16)
                sendBytes(i + 1) = Convert.ToByte(MacAddress.Substring(2, 2), 16)
                sendBytes(i + 2) = Convert.ToByte(MacAddress.Substring(4, 2), 16)
                sendBytes(i + 3) = Convert.ToByte(MacAddress.Substring(6, 2), 16)
                sendBytes(i + 4) = Convert.ToByte(MacAddress.Substring(8, 2), 16)
                sendBytes(i + 5) = Convert.ToByte(MacAddress.Substring(10, 2), 16)
                i += 6
            Next

            myAddress = Net.IPAddress.Parse(PlugInIPAddress).ToString

            If myAddress = String.Empty Then
                Log("Error in SendMagicPacket. Invalid IP address/Host Name given", LogType.LOG_TYPE_ERROR)
                Return
            End If

            Dim mySubnetArray() As String
            Dim myIPAddressArray() As String
            Dim sm1, sm2, sm3, sm4 As Byte
            Dim ip1, ip2, ip3, ip4 As Byte

            mySubnetArray = GetSubnetMask(inIPAddress).Split("."c)
            myIPAddressArray = myAddress.Split("."c)

            For i = 0 To mySubnetArray.GetUpperBound(0)
                Select Case i
                    Case Is = 0
                        sm1 = Convert.ToByte(mySubnetArray(i))
                    Case Is = 1
                        sm2 = Convert.ToByte(mySubnetArray(i))
                    Case Is = 2
                        sm3 = Convert.ToByte(mySubnetArray(i))
                    Case Is = 3
                        sm4 = Convert.ToByte(mySubnetArray(i))
                End Select
            Next
            For i = 0 To myIPAddressArray.GetUpperBound(0)
                Select Case i
                    Case Is = 0
                        ip1 = Convert.ToByte(myIPAddressArray(i))
                        ip1 = ip1 And sm1 Or (Not sm1) ' Xor 0)
                    Case Is = 1
                        ip2 = Convert.ToByte(myIPAddressArray(i))
                        ip2 = ip2 And sm2 Or (Not sm2)
                    Case Is = 2
                        ip3 = Convert.ToByte(myIPAddressArray(i))
                        ip3 = ip3 And sm3 Or (Not sm3)
                    Case Is = 3
                        ip4 = Convert.ToByte(myIPAddressArray(i))
                        ip4 = ip4 And sm4 Or (Not sm4)
                End Select
            Next

            myAddress = ip1.ToString & "." & ip2.ToString & "." & ip3.ToString & "." & ip4.ToString
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMagicPacket has Broadcast IPAddress = " & myAddress, LogType.LOG_TYPE_INFO)

        Catch ex As Exception
            Log("Error in SendMagicPacket with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim bytesSent As Integer = 0
        Try
            bytesSent = udpClient.Send(sendBytes, sendBytes.Length, myAddress, Port)
        Catch ex As Exception
            Log("Error in SendMagicPacket sending bytes with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SendMagicPacket for MacAddress = " & MacAddress & " has sent = " & bytesSent.ToString & " bytes", LogType.LOG_TYPE_INFO)
        Try
            udpClient.Close()
        Catch ex As Exception
        End Try

        udpClient = Nothing

    End Sub


End Class

