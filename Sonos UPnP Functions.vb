
Imports System.Math
Imports System.Text.RegularExpressions.Regex
Imports System.Runtime.InteropServices
Imports System.Xml
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.IO.Path
Imports System.Drawing
Imports System.IO
Imports System.Data.SQLite



'<Serializable()> _
'Public Class ZonePlayerController
Public Class SonosUPNPFunctions

    'Inherits MarshalByRefObject

    Public WithEvents myAVTransportCallback As New myUPnPControlCallback
    Public WithEvents myRenderingControlCallback As New myUPnPControlCallback
    Public WithEvents myContentDirectoryCallback As New myUPnPControlCallback
    Public WithEvents myAudioInCallback As New myUPnPControlCallback
    Public WithEvents myDevicePropertiesCallback As New myUPnPControlCallback

    Public WithEvents myAlarmClockCallback As New myUPnPControlCallback
    Public WithEvents myMusicServicesCallback As New myUPnPControlCallback
    Public WithEvents mySystemPropertiesCallback As New myUPnPControlCallback
    Public WithEvents myZonegroupTopologyCallback As New myUPnPControlCallback
    Public WithEvents myGroupManagementCallback As New myUPnPControlCallback
    Public WithEvents myConnectionManagerCallback As New myUPnPControlCallback
    Public WithEvents myQueueServiceCallback As New myUPnPControlCallback
    Public WithEvents myVirtualLineInCallback As New myUPnPControlCallback

    Private MyUPnPDevice As MyUPnPDevice
    Private MediaServer As MyUPnPDevice = Nothing
    Private MediaRenderer As MyUPnPDevice = Nothing
    Private AudioIn As MyUPnPService = Nothing
    Private DeviceProperties As MyUPnPService = Nothing
    Private AVTransport As MyUPnPService = Nothing
    Private RenderingControl As MyUPnPService = Nothing
    Private ContentDirectory As MyUPnPService = Nothing
    Private AlarmClock As MyUPnPService = Nothing
    Private ZoneGroupTopology As MyUPnPService = Nothing
    Private MusicServices As MyUPnPService = Nothing
    Private QueueService As MyUPnPService = Nothing     ' added 2/24/2019 version 3.1.0.29
    Private VirtualLineIn As MyUPnPService = Nothing    ' added 2/24/2019 version 3.1.0.29

    Private Properties(4) As String
    Public AudioInState(6)
    Private UDN As String = ""
    Private IPAddress As String = ""
    Private IPPort As String = ""
    Private MyDeviceStatus As String = "Offline"
    Private DockedDeviceStatus As String = "Offline"
    Private ZoneName As String = ""
    Private HSRefPlayer As Integer = -1
    Private HSRefTrack As Integer = -1
    Private HSRefNextTrack As Integer = -1
    Private HSRefArtist As Integer = -1
    Private HSRefNextArtist As Integer = -1
    Private HSRefAlbum As Integer = -1
    Private HSRefNextAlbum As Integer = -1
    Private HSRefArt As Integer = -1
    Private HSRefNextArt As Integer = -1
    Private HSRefPlayState As Integer = -1
    Private HSRefVolume As Integer = -1
    Private HSRefMute As Integer = -1
    Private HSRefLoudness As Integer = -1
    Private HSRefBalance As Integer = -1
    Private HSRefTrackLength As Integer = -1
    Private HSRefTrackPos As Integer = -1
    Private HSRefRadiostationName As Integer = -1
    Private HSRefDockDeviceName As Integer = -1
    Private HSRefTrackDescr As Integer = -1
    Private HSRefRepeat As Integer = -1
    Private HSRefShuffle As Integer = -1
    Private HSRefGenre As Integer = -1
    'Private HSDeviceCodeRendering As Integer = -1
    Private MyHSTMusicIndex As Integer = 0
    Private MyZoneIsLinked As Boolean = False
    Private MySourceLinkedZone As String = ""
    Private MyHSPIControllerRef As Object ' HSPI = Me
    Private MyZoneIsSourceForLinkedZone As Boolean = False
    Private MyTargetZoneLinkedList As String = "" ' this is a list of Zone Names when this zone is source for other linked zone players
    Private WaitingToReConnect As Boolean = False
    Private MyPlayerTimeoutActionArray(MaxPlayerTOActionArray) As Integer
    Private MyFailedPingCount As Integer = 0
    Private MyConnectRetryCount As Integer = 0
    Private LinkGroupInfoArray As Object
    Private LinkGroupInfoArrayIndex As Integer = 0
    Private MyCurrentTrackInfo As Object = {"", "", "", "", NoArtPath, "", "", "0", "", "", "False", "", "", "", "", ""}
    Private MyCurrentTransportState As String = ""
    Private MyCurrentMuteState As Boolean = -1
    Private MyCurrentLoudnessState As Boolean = -1
    Private MyCurrentMasterVolumeLevel As Integer = 0
    Private MyCurrentFixedVolumeState As Boolean = False
    Private MyRightVolume As Integer = 100
    Private MyLeftVolume As Integer = 100
    Private MyBalance As Integer = 0
    Private ZoneHasFixedVolume As Boolean = False
    Private MyZoneModel As String = ""
    Private MyDestinationZone As String = ""
    Private MyWirelessDockSourcePlayer As Object = Nothing 'HSPI = Nothing
    Private MyWirelessDockDestinationPlayer As Object = Nothing ' HSPI = Nothing
    Private MyWirelessDockZoneName As String = ""
    Private MyAutoPlayRoomUUID As String = ""
    Private MyAutoPlayLinkedZones As Boolean = False
    Private MyAutoPlayVolume As Integer = 0
    Private TimerReEntry As Boolean = False
    Private MyLineInputConnected As Boolean = False
    Private MyZoneIsStored As Boolean = False
    Private MyPlayerStateBeforeAnnouncement As player_state_values
    Private MyPreviousArtist As String = ""
    Private MyPreviousTrack As String = ""
    Private MyPreviousAlbum As String = ""
    Private MyPreviousNextArtist As String = ""
    Private MyPreviousNextTrack As String = ""
    Private MyPreviousNextAlbum As String = ""
    Public MyChannelMapSet As String = ""
    Public MyHTSatChanMapSet As String = ""
    Private MyZoneIsPlaybarMaster As Boolean = False
    Private MyZoneIsPlaybarSlave As Boolean = False
    Private MyZonePlayBarLeftRearUDN As String = ""
    Private MyZonePlayBarRightRearUDN As String = ""
    Private MyZonePlayBarUDN As String = ""
    Private MyZonePlayBarLeftFrontUDN As String = ""
    Private MyZonePlayBarRightFrontUDN As String = ""

    Private MyZoneIsPairMaster As Boolean = False
    Private MyZoneIsPairSlave As Boolean = False
    Private MyZonePairLeftFrontUDN As String = ""
    Private MyZonePairRightFrontUDN As String = ""

    Private MyZonePairMasterZoneName As String = ""
    Private MyZonePairMasterZoneUDN As String = ""
    Private MyZonePairSlaveZoneUDN As String = ""
    Private MyZonePairSubWooferZoneUDN As String = ""
    Private MyPreviousAlbumArtPath As String = ""
    Private MyPreviousNextAlbumArtPath As String = ""
    Private MyPreviousAlbumURI As String = ""
    Private MyPreviousNextAlbumURI As String = ""
    Private MyIconURL As String = ""
    Private MyRetryLinkAZone As Boolean = False
    Private MyDeleteQueueName As String = ""
    Private MyAdminStateActive As Boolean = False
    Private MyMissedPings As Integer = 0
    Private MyZoneGroupState As String = ""
    Public MyZoneGroupStateHasChanged As Boolean = False
    Private MySWVersion As Integer = 0
    Private MyGroupCoordinatorIsLocal As Boolean = True
    Private MyLocalGroupUUID As String = ""
    Private MyIPAddress As String = ""
    Private MyMACAddress As String = ""
    Private MyHouseHoldID As String = ""
    Private MyThirdPartyMediaServicesX As String = ""
    Private MyEnqueuedTransportURI As String = ""
    Private MyEnqueuedTransportURIMetaData As String = ""
    Private MyQueueServiceLastInfo As String = ""

    ' here is more info on how to browse the Sonos Player
    'A: Music Library (organized by id3tags) 'used device spy and type in A: 
    ' A:ARTIST
    ' A:ALBUMARTIST
    ' A:ALBUM
    ' A:GENRE
    ' A:COMPOSER
    ' A:TRACKS
    ' A:PLAYLISTS
    'Q: Queues
    'SQ: Sonos Queue
    'S: Storage/Servers (cifs shares, browsable)
    'R: Radio Stations
    'G: Now playing (always just a single item) ??? not sure this is still supported
    'AI: for AudioInput
    'EN: Entire Network
    'RP: Recently Played
    'FV:2 Favorites
    ' you get the following with Browse ContentDirectory ObjectID = 0; BrowseDirectChildren
    '<DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
    '<container id="A:" parentID="0" restricted="true"><dc:title>Attributes</dc:title><upnp:class>object.container</upnp:class></container>
    '<container id="S:" parentID="0" restricted="false"><dc:title>Music Shares</dc:title><upnp:class>object.container</upnp:class></container>
    '<container id="Q:" parentID="0" restricted="true"><dc:title>Queues</dc:title><upnp:class>object.container</upnp:class></container>
    '<container id="SQ:" parentID="0" restricted="true"><dc:title>Saved Queues</dc:title><upnp:class>object.container</upnp:class></container>
    '<container id="R:" parentID="0" restricted="true"><dc:title>Internet Radio</dc:title><upnp:class>object.container</upnp:class></container>
    '<container id="AI:" parentID="0" restricted="true"><dc:title>Audio Inputs</dc:title><upnp:class>object.container</upnp:class></container>
    '<container id="EN:" parentID="0" restricted="true"><dc:title>Entire Network</dc:title><upnp:class>object.container</upnp:class></container>



    'The main queue (Q:0) gives you the active playlist.

    ' For WD100 you need to call SetAvTransportURI for the zone that is selected as ouput. The URI is the ObjectId prepended by x-sonos-dock:UDN/ObjectId
    '  exmaple : x-sonos-dock:RINCON_000E5860905A01400/1:2:5
    Public WriteOnly Property pDevice As MyUPnPDevice
        Set(value As MyUPnPDevice)
            MyUPnPDevice = value
        End Set
    End Property

    Public Property MissedPings As Integer
        Get
            MissedPings = MyMissedPings
        End Get
        Set(value As Integer)
            MyMissedPings = value
        End Set
    End Property

    Private Sub MyZonePlayerTimer_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles MyZonePlayerTimer.Elapsed
        If TimerReEntry Then Exit Sub
        If MyRetryLinkAZone Then
            TimerReEntry = True
            MyRetryLinkAZone = False
            If MySourceLinkedZone <> "" Then
                HandleLinkedZones(MySourceLinkedZone)
            End If
        End If
        If MyDeleteQueueName <> "" Then
            Dim TempStr As String = MyDeleteQueueName
            MyDeleteQueueName = "" ' this will prevent re-entrancy
            DeletePreviousSavedQueues(TempStr)
        End If
        TimerReEntry = True
        Try
            Reachable()
            If ConnectToIPod Then
                ConnectWireLessDock("uuid:DOCK" & UDN & "_MS")
                ConnectToIPod = False
            End If
            If ConnectPlayer Then
                Connect("uuid:" & UDN)
                ConnectPlayer = False
            End If
            sender = Nothing
            e = Nothing
        Catch ex As Exception
        End Try
        TimerReEntry = False
    End Sub

    Private Function CheckIPAddressChange() As Boolean
        CheckIPAddressChange = False
        If MyUPnPDevice Is Nothing Then Exit Function ' no device is alive on the network
        If g_bDebug Then Log("CheckIPAddressChange for UPnPDevice = " & ZoneName & " found IPAddress = " & MyUPnPDevice.IPAddress, LogType.LOG_TYPE_INFO)
        If g_bDebug Then Log("CheckIPAddressChange for UPnPDevice = " & ZoneName & " found IPPort    = " & MyUPnPDevice.IPPort, LogType.LOG_TYPE_INFO)

        If MyUPnPDevice.IPAddress <> IPAddress Then
            CheckIPAddressChange = True
            Log("IPAddress for UPnPDevice = " & ZoneName & " has changed. Old = " & IPAddress & ". New = " & MyUPnPDevice.IPAddress, LogType.LOG_TYPE_INFO)
            IPAddress = MyUPnPDevice.IPAddress
            WriteStringIniFile(UDN, DeviceInfoIndex.diIPAddress.ToString, IPAddress)
        End If
    End Function

    Private Function Reachable() As Boolean
        Reachable = False
        If MyUPnPDevice Is Nothing Then Exit Function ' no device is alive on the network
        Try
            If MyUPnPDevice.Alive Then
                If DeviceStatus = "Offline" Then
                    If SuperDebug Then Log("Reachable called for UPnPDevice " & ZoneName & " which is reachable on network. Attempting to reconnect", LogType.LOG_TYPE_INFO)
                    Connect("uuid:" & UDN)
                End If
                Reachable = True
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in Reachable for UPnPDevice " & ZoneName & " calling the ping status with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Function

    Private Sub DestroyObjects(Full As Boolean)
        DeviceStatus = "Offline"
        'Try
        'If MyUPnPDevice IsNot Nothing Then '  If MyUPnPDevice IsNot Nothing And Full Then
        'RemoveHandler MyUPnPDevice.DeviceDied, AddressOf myDeviceFinderCallback_DeviceLost
        'RemoveHandler MyUPnPDevice.DeviceAlive, AddressOf myDeviceFinderCallback_DeviceAlive
        'End If
        'Catch ex As Exception
        'End Try
        Try
            If AudioIn IsNot Nothing Then AudioIn.RemoveCallback()
        Catch ex As Exception
        End Try
        AudioIn = Nothing
        Try
            If DeviceProperties IsNot Nothing Then DeviceProperties.RemoveCallback()
        Catch ex As Exception
        End Try
        DeviceProperties = Nothing
        Try
            If AVTransport IsNot Nothing Then AVTransport.RemoveCallback()
        Catch ex As Exception
        End Try
        AVTransport = Nothing
        Try
            If RenderingControl IsNot Nothing Then RenderingControl.RemoveCallback()
        Catch ex As Exception
        End Try
        RenderingControl = Nothing
        Try
            If ContentDirectory IsNot Nothing Then ContentDirectory.RemoveCallback()
        Catch ex As Exception
        End Try
        ContentDirectory = Nothing
        Try ' added 2/24/2019 in v 3.1.0.29
            If QueueService IsNot Nothing Then QueueService.RemoveCallback()
        Catch ex As Exception
        End Try
        QueueService = Nothing
        Try ' added 2/24/2019 in v 3.1.0.29
            If VirtualLineIn IsNot Nothing Then VirtualLineIn.RemoveCallback()
        Catch ex As Exception
        End Try
        VirtualLineIn = Nothing
        MediaServer = Nothing
        MediaRenderer = Nothing
        Try
            If AlarmClock IsNot Nothing Then AlarmClock.RemoveCallback()
        Catch ex As Exception
        End Try
        AlarmClock = Nothing
        Try
            If ZoneGroupTopology IsNot Nothing Then ZoneGroupTopology.RemoveCallback()
        Catch ex As Exception
        End Try
        ZoneGroupTopology = Nothing
        MusicServices = Nothing

        If MyUPnPDevice IsNot Nothing Then
            Try
                MyUPnPDevice.Dispose(True)
            Catch ex As Exception
            End Try
        End If
        MyUPnPDevice = Nothing
    End Sub

    Public WriteOnly Property HSPIControllerRef As Object 'HSPI
        ' pass a reference to the Sonos HS plugin. This is needed to find other zone players in case zones are linked and updates need to be send to the other instances of ZonePlayerController
        ' to have their events and status updated in HS
        Set(ByVal value As Object) 'HSPI)
            MyHSPIControllerRef = value
        End Set
    End Property

    Public Property SourcingZone As Boolean
        ' this is called to set or get info on whether this Sonos is part of a linked zone and more importantly, whether it is the source for it or not
        Set(ByVal value As Boolean)
            MyZoneIsSourceForLinkedZone = value
            If g_bDebug Then Log("SourceZone Called for ZoneName " & ZoneName & " with Value: " & SourcingZone.ToString, LogType.LOG_TYPE_INFO)
        End Set
        Get
            SourcingZone = MyZoneIsSourceForLinkedZone
        End Get
    End Property

    Public ReadOnly Property DestinationZone As String
        ' I should use this only for WD100 devices but in principle any app can use it but it will return a list
        Get
            DestinationZone = ""
            'If Not SourcingZone Then Exit Property
            If MyZoneModel <> "WD100" Then
                DestinationZone = MyTargetZoneLinkedList
                Exit Property
            End If
            ' Let's retrieve the dest zone for the WD100
            Try
                DestinationZone = GetAutoplayRoomUUID()
                'If DestinationZone = "" Then Exit Property ' nothing specified
                ' translate the UUID into a zonename
                'DestinationZone = GetZoneByUDN(DestinationZone)
            Catch ex As Exception
                Log("Error in get DestinationZone with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End Get
    End Property

    Public ReadOnly Property ZoneIsLinked As Boolean
        Get
            ZoneIsLinked = MyZoneIsLinked
        End Get
    End Property

    Public Property DeviceStatus As String
        Get
            DeviceStatus = MyDeviceStatus
        End Get
        Set(value As String)
            If value <> MyDeviceStatus Then
                If value.ToUpper = "ONLINE" Then
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.Online)
                Else
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.Offline)
                End If
            End If
            MyDeviceStatus = value
        End Set
    End Property

    Public Sub PassZoneName(ByVal strZoneName As String)
        ZoneName = strZoneName
    End Sub

    Public Sub PassUDN(ByVal strUDN As String)
        UDN = strUDN
    End Sub

    Public Function GetHSDeviceRefPlayer() As Integer
        GetHSDeviceRefPlayer = HSRefPlayer
    End Function

    Public Property ZoneModel As String
        Get
            ZoneModel = MyZoneModel
        End Get
        Set(ByVal value As String)
            MyZoneModel = value
        End Set
    End Property

    Public Sub SetHSDeviceRefPlayer(ByVal strDeviceRef As Integer)
        ' this holds the housecode/devicecode that was created by HS so it can be used to send events back to HS to update status etc
        HSRefPlayer = strDeviceRef
        If g_bDebug Then Log("SetHSDeviceRefPlayer called for zone - " & ZoneName & " with HS Ref = " & strDeviceRef, LogType.LOG_TYPE_INFO)
    End Sub

    Public Property HSTMusicIndex As Integer
        ' The HS Touch Designer and Client, especually the specially created "MUSIC" devices, do not use House/Device-codes but use the zone or player name. A special index is stored
        ' here to make the link between zone names called by HST (HSTouch) and the instance of the zone player. It is called from HSPI_Sonosc
        Get
            HSTMusicIndex = MyHSTMusicIndex
        End Get
        Set(ByVal value As Integer)
            MyHSTMusicIndex = value
        End Set
    End Property

    Public ReadOnly Property GroupCoordinatorIsLocal As Boolean
        Get
            GroupCoordinatorIsLocal = MyGroupCoordinatorIsLocal
        End Get
    End Property

    Public ReadOnly Property LocalGroupUUID As String
        Get
            LocalGroupUUID = MyLocalGroupUUID
        End Get
    End Property

    Public ReadOnly Property ZoneIPAddress As String
        Get
            ZoneIPAddress = MyIPAddress
        End Get
    End Property

    Public ReadOnly Property ZoneMACAddress As String
        Get
            ZoneMACAddress = MyMACAddress
        End Get
    End Property

    Public ReadOnly Property ZoneThirdPartyMediaServicesX As String
        Get
            ZoneThirdPartyMediaServicesX = MyThirdPartyMediaServicesX
        End Get
    End Property

    Private Function WaitToReconnect() As Boolean
        ' I noted that immediately reconnecting after a player died isn't working, so I added some time and prevent multiple attempts while waiting
        WaitToReconnect = False
        If WaitingToReConnect Then
            Log("WaitToReconnect - already waiting for zoneplayer - " & ZoneName, LogType.LOG_TYPE_INFO)
            Exit Function
        End If
        WaitingToReConnect = True
        wait(10)
        WaitingToReConnect = False
        WaitToReconnect = True
    End Function

    Public Function GetZoneByUDN(ByVal UDN As String) As String
        GetZoneByUDN = ""
        If UDN = "" Then Exit Function
        If SuperDebug Then Log("GetZoneByUDN called with UDN = " & UDN, LogType.LOG_TYPE_INFO)
        Dim ZoneIndex As Integer = 0
        If InStr(UDN, "-sonos-dock:") <> 0 Then
            Try
                UDN = Replace(UDN, "x-sonos-dock:", "") ' "x-sonos-dock:" Remove the x-sonos-dock:
            Catch ex As Exception
                Log("Issue with UDN in GetZoneByUDN. UDN = " & UDN & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            Try
                UDN = Replace(UDN, "x-rincon:", "") ' "x-rincon:" Remove the x-rincon
            Catch ex As Exception
                Log("Issue with UDN in GetZoneByUDN. UDN = " & UDN & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        End If
        GetZoneByUDN = GetStringIniFile(UDN, DeviceInfoIndex.diFriendlyName.ToString, "")
    End Function

    Public Sub DirectConnect(ByVal pDevice As MyUPnPDevice, ByVal inUDN As String)

        If pDevice Is Nothing Then
            Log("Error in DirectConnect called for Zone " & ZoneName & " with UDN = " & inUDN & ". The pDevice pointer is NILL", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If


        Dim Children As MyUPnPDevices
        Dim objChild As MyUPnPDevice

        ' Initialize all the State/Info Variables
        Properties = {"", "", "", "", ""}
        AudioInState = {"", "", 0, False, False, False, False}

        Log("DirectConnect called for Zone " & ZoneName & " with device name = " & pDevice.UniqueDeviceName & " and Model = " & pDevice.ModelNumber, LogType.LOG_TYPE_INFO)

        If Mid(pDevice.UniqueDeviceName, 1, 12) <> "uuid:RINCON_" Then
            If g_bDebug Then Log("DirectConnect for zoneplayer = " & ZoneName & " found non Sonos device with UDN =  " & pDevice.UniqueDeviceName & " Friendly Name = " & pDevice.FriendlyName, LogType.LOG_TYPE_WARNING)
            Exit Sub ' this is the UPNP service of HS itself on an XP machine responding
        End If
        UDN = Replace(inUDN, "uuid:", "")
        MyZoneModel = pDevice.ModelNumber  ' store the Model here ZP80/90/100/120 or WD100. The latter is needed for a multitude of checks
        MyUPnPDevice = pDevice
        IPAddress = pDevice.IPAddress
        IPPort = pDevice.IPPort

        pDevice.AddHandlers(Me)

        Try
            MyIconURL = pDevice.IconURL("image/png", 200, 200, 16) 'image/png image/x-png image/tiff image/bmp image/pjpeg image/jpeg
            If g_bDebug Then Log("DirectConnect for zoneplayer = " & ZoneName & " found IconURL = " & MyIconURL, LogType.LOG_TYPE_INFO)
            If SuperDebug Then Log("DirectConnect for zoneplayer = " & ZoneName & " checking for file = " & CurrentAppPath & "/html" & URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", LogType.LOG_TYPE_INFO)
            Dim FileExists As Boolean = False
            If ImRunningLocal Then
                FileExists = File.Exists(CurrentAppPath & "\html" & URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png")
            Else
                ' this is where I need the new HTMLCommands to check for a file on the remote system
                FileExists = False ' dcor tobefixed
            End If
            If Not FileExists Then
                Dim IConImage As Image
                IConImage = GetPicture(MyIconURL)
                If Not IConImage Is Nothing Then
                    Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                    MyIconURL = URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png"
                    Dim SuccesfullSave As Boolean = False
                    SuccesfullSave = hs.WriteHTMLImage(IConImage, FileArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", True)
                    If Not SuccesfullSave Then
                        If g_bDebug Then Log("Error in DirectConnect for zoneplayer = " & ZoneName & " had error storing Icon at " & FileArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", LogType.LOG_TYPE_ERROR)
                    Else
                        If g_bDebug Then Log("DirectConnect for zoneplayer = " & ZoneName & " stored Icon at " & FileArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", LogType.LOG_TYPE_INFO)
                    End If
                    IConImage.Dispose()
                End If
            Else
                MyIconURL = URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png"
                If g_bDebug Then Log("DirectConnect for zoneplayer = " & ZoneName & " found Icon already stored at " & MyIconURL, LogType.LOG_TYPE_INFO)
            End If

            If HSRefPlayer <> -1 Then
                Dim dv As Scheduler.Classes.DeviceClass
                dv = hs.GetDeviceByRef(HSRefPlayer)
                dv.Image(hs) = MyIconURL
                dv.ImageLarge(hs) = MyIconURL
                If g_bDebug Then Log("DirectConnect for zoneplayer = " & ZoneName & " added Image = " & MyIconURL & " for HSRef = " & HSRefPlayer.ToString, LogType.LOG_TYPE_INFO)
                WriteStringIniFile(UDN, DeviceInfoIndex.diDeviceIConURL.ToString, MyIconURL)
            End If

        Catch ex As Exception
            If g_bDebug Then Log("Error in DirectConnect. Could not get ICON info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        hs.SetDeviceValue(HSRefPlayer, CurrentPlayerState)

        Try
            Children = pDevice.Children
            If Children IsNot Nothing Then
                If Children.Count > 0 Then
                    For Each objChild In Children
                        If objChild IsNot Nothing Then
                            If objChild.UniqueDeviceName = ("uuid:" & UDN & "_MS") Then
                                MediaServer = objChild
                                If SuperDebug Then Log(ZoneName & " : MediaServer Friendly Name = " & MediaServer.UniqueDeviceName, LogType.LOG_TYPE_INFO)
                            ElseIf objChild.UniqueDeviceName = ("uuid:" & UDN & "_MR") Then
                                MediaRenderer = objChild
                                If SuperDebug Then Log(ZoneName & " : MediaRenderer Friendly Name = " & MediaRenderer.FriendlyName, LogType.LOG_TYPE_INFO)
                            Else
                                If SuperDebug Then Log(ZoneName & " : Additional service Fiendly Name = " & objChild.UniqueDeviceName, LogType.LOG_TYPE_INFO) ' for testing proposes
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Log(ZoneName & " : There was a problem in getting MediaServer & MediaRenderer" & ex.Message, LogType.LOG_TYPE_ERROR)
            Log(ZoneName & " : Seach item = uuid:" & UDN & "_MS / _MR", LogType.LOG_TYPE_ERROR)
        End Try

        Try
            For Each Serviceid As MyUPnPService In pDevice.Services
                If g_bDebug Then Log("service = " & Serviceid.Id & " found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Next
        Catch ex As Exception

        End Try
        If pDevice IsNot Nothing Then
            Try
                AudioIn = pDevice.Services.Item("urn:upnp-org:serviceId:AudioIn")
                If SuperDebug And (AudioIn IsNot Nothing) Then Log("AudioIn service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                AlarmClock = pDevice.Services.Item("urn:upnp-org:serviceId:AlarmClock")
                If SuperDebug And (AlarmClock IsNot Nothing) Then Log("AlarmClock service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                DeviceProperties = pDevice.Services.Item("urn:upnp-org:serviceId:DeviceProperties")
                If SuperDebug And (DeviceProperties IsNot Nothing) Then Log("DeviceProperties service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                AVTransport = pDevice.Services.Item("urn:upnp-org:serviceId:AVTransport")
                If SuperDebug And (AVTransport IsNot Nothing) Then Log("AVTransport service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                ZoneGroupTopology = pDevice.Services.Item("urn:upnp-org:serviceId:ZoneGroupTopology")
                If SuperDebug And (ZoneGroupTopology IsNot Nothing) Then Log("ZoneGroupTopology service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                MusicServices = pDevice.Services.Item("urn:upnp-org:serviceId:MusicServices")
                If SuperDebug And (MusicServices IsNot Nothing) Then Log("MusicServices service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
        End If

        If Not MediaRenderer Is Nothing Then
            Try
                AVTransport = MediaRenderer.Services.Item("urn:upnp-org:serviceId:AVTransport")
                If SuperDebug And (AVTransport IsNot Nothing) Then Log("AVTransport service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                RenderingControl = MediaRenderer.Services.Item("urn:upnp-org:serviceId:RenderingControl")
                If SuperDebug And (RenderingControl IsNot Nothing) Then Log("RenderingControl service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try ' added 2/24/2019 in v3.1.0.29
                QueueService = MediaRenderer.Services.Item("urn:sonos-com:serviceId:Queue")
                If SuperDebug And (QueueService IsNot Nothing) Then Log("Queue service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try ' added 2/24/2019 in v3.1.0.29
                VirtualLineIn = MediaRenderer.Services.Item("urn:upnp-org:serviceId:VirtualLineIn")
                If SuperDebug And (VirtualLineIn IsNot Nothing) Then Log("VirtualLineIn service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
        End If

        If Not MediaServer Is Nothing Then
            Try
                ContentDirectory = MediaServer.Services.Item("urn:upnp-org:serviceId:ContentDirectory")
                If SuperDebug And (ContentDirectory IsNot Nothing) Then Log("ContentDirectory service added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in MediaServer.Services.ContentDirectory for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If MyZoneModel.ToUpper = "SUB" Then
            If g_bDebug Then Log("DirectConnect for zoneplayer = " & ZoneName & " is ignoring found Sonos SUB device with UDN =  " & pDevice.UniqueDeviceName & " Friendly Name = " & pDevice.FriendlyName, LogType.LOG_TYPE_INFO)
            DeviceStatus = "Online"
            Exit Sub
        End If
        If Not ZoneGroupTopology Is Nothing Then
            Try
                ZoneGroupTopology.AddCallback(myZonegroupTopologyCallback)
                If g_bDebug Then Log("ZoneGroupTopology ControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding ZoneGroupTopology Control Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If Not ZoneGroupTopology Is Nothing Then
            GetZoneGroupState()
        End If

        If MyZoneIsPlaybarSlave Then
            RenderingControl = Nothing    ' added 7/13/2019 in v3.1.0.34 to prevent errors for players that are paired to a play bar/beam/base
            QueueService = Nothing
            VirtualLineIn = Nothing
        End If

        If Not AudioIn Is Nothing Then
            Try
                AudioIn.AddCallback(myAudioInCallback)
                If g_bDebug Then Log("AudioInCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding AudioIn Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If Not DeviceProperties Is Nothing Then
            Try
                DeviceProperties.AddCallback(myDevicePropertiesCallback)
                If g_bDebug Then Log("DevicePropertiesCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding DeviceProperties Call Back for zoneplayer = " & ZoneName & ". Error" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If Not AVTransport Is Nothing Then
            Try
                AVTransport.AddCallback(myAVTransportCallback)
                If g_bDebug Then Log("AvTransportCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding Transport Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If RenderingControl IsNot Nothing Then
            Try
                RenderingControl.AddCallback(myRenderingControlCallback)
                If g_bDebug Then Log("RenderingControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                If AVTransport IsNot Nothing Then Log("Error in Adding RenderingControl Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR) ' if no AVTransport, this is a 5:1 surround sound speaker
            End Try
        End If

        If Not ContentDirectory Is Nothing Then
            Try
                ContentDirectory.AddCallback(myContentDirectoryCallback)
                If g_bDebug Then Log("ContentDirectory ControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding ContentDirectoryControl Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If QueueService IsNot Nothing Then  ' added 2/24/2019 in v3.1.0.29
            Try
                QueueService.AddCallback(myQueueServiceCallback)
                If g_bDebug Then Log("QueueService added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding QueueService Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If VirtualLineIn IsNot Nothing Then  ' added 2/24/2019 in v3.1.0.29  
            Try
                VirtualLineIn.AddCallback(myVirtualLineInCallback)
                If g_bDebug Then Log("VirtualLineIn added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding VirtualLineIn Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If Not AlarmClock Is Nothing Then
            Try
                AlarmClock.AddCallback(myAlarmClockCallback)
                If g_bDebug Then Log("AlarmClock ControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding AlarmClock Control Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If



        If DeviceProperties Is Nothing Then '    If AVTransport Is Nothing Or DeviceProperties Is Nothing Then
            Log("Error in DirectConnect for zoneplayer = " & ZoneName & ". Some crucial services are not on-line. The plug-in cannot work for this zone", LogType.LOG_TYPE_ERROR)
        Else
            DeviceStatus = "Online"
        End If
        If Not RenderingControl Is Nothing Then
            Try
                MyCurrentMasterVolumeLevel = GetVolumeLevel("Master")
                SetVolume = MyCurrentMasterVolumeLevel
                MyLeftVolume = GetVolumeLevel("LF")
                MyRightVolume = GetVolumeLevel("RF")
                UpdateBalance()
                SetMuteState = GetMuteState("Master")
                SetLoudness = GetLoudnessState("Master")
            Catch ex As Exception
            End Try
        End If
        If Not AudioIn Is Nothing Then
            'If ZoneModel = "WD100" Then
            ' we need to find out at start up whether something is docked. The Zone player doesn't fire off automatically
            'End If
            GetAudioInputAttributes()
        End If
        If Not DeviceProperties Is Nothing Then
            Try
                If ZoneModel <> "WD100" Then
                    Dim InArg(0) As String
                    Dim OutArg(2) As String
                    Dim CurrentZoneName As String = ""
                    Try
                        MyUPnPDevice.Services.Item("urn:upnp-org:serviceId:DeviceProperties").InvokeAction("GetZoneAttributes", InArg, OutArg)
                        CurrentZoneName = OutArg(0) ' this is CurrentZoneName
                        If g_bDebug Then Log("DirectConnect found Zone Name = " & CurrentZoneName, LogType.LOG_TYPE_INFO)
                    Catch ex As Exception
                        If g_bDebug Then Log("Error in DirectConnect while getting ZoneAttributes with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    If (CurrentZoneName <> "") And ZoneName <> CurrentZoneName Then ZoneNameChanged(CurrentZoneName)
                End If
            Catch ex As Exception
                If g_bDebug Then Log("Warning in DirectConnect for zoneplayer = " & ZoneName & " when getting the ChannelMapSet/ZoneName with Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            End Try
        End If
        Try
            'IPAddress = GetIPAddress()
            If g_bDebug Then Log("IPAddress for zoneplayer = " & ZoneName & " = " & IPAddress, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Issue Calling GetIPAddress for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            MySWVersion = GetSoftwareVersion()
            'if g_bDebug then Log("SW Version for zoneplayer = " & ZoneName & " = " & MySWVersion.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Issue Calling GetSoftwareVersion for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If Not AVTransport Is Nothing Then
            GetPlayMode()
            UpdateHS(True)
        End If
    End Sub


    Private Sub myDeviceFinderCallback_DeviceFound(ByVal pDevice As MyUPnPDevice)

        If pDevice Is Nothing Then
            Log("Error in myDeviceFinderCallback_DeviceFound called for Zone " & ZoneName & ". The pDevice pointer is NILL", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Dim Children As MyUPnPDevices
        Dim objChild As MyUPnPDevice

        ' Indicates new Device

        If g_bDebug Then Log("myDeviceFinderCallback_DeviceFound received for Zone " & ZoneName & " with device name = " & pDevice.UniqueDeviceName & " and Model = " & pDevice.ModelNumber, LogType.LOG_TYPE_INFO)

        If Mid(pDevice.UniqueDeviceName, 1, 12) <> "uuid:RINCON_" Then
            If g_bDebug Then Log("Device Finder Call Back for zoneplayer = " & ZoneName & " found non Sonos device with UDN =  " & pDevice.UniqueDeviceName & " Friendly Name = " & pDevice.FriendlyName, LogType.LOG_TYPE_WARNING)
            Exit Sub ' this is the UPNP service of HS itself on an XP machine responding
        End If

        ConnectPlayer = False ' make sure we don't reconnect on timer-ex v.92. received a disconnect and findercallback all automously 

        MyZoneModel = pDevice.ModelNumber  ' store the Model here ZP80/90/100/120 or WD100. The latter is needed for a multitude of checks

        MyUPnPDevice = pDevice
        IPAddress = pDevice.IPAddress
        IPPort = pDevice.IPPort

        pDevice.AddHandlers(Me)

        Try
            MyIconURL = pDevice.IconURL("image/png", 200, 200, 16) 'image/png image/x-png image/tiff image/bmp image/pjpeg image/jpeg
            If g_bDebug Then Log("Device Finder CallBack for zoneplayer = " & ZoneName & " found IconURL = " & MyIconURL, LogType.LOG_TYPE_INFO)
            If SuperDebug Then Log("Device Finder CallBack for zoneplayer = " & ZoneName & " checking for file = " & CurrentAppPath & "\html" & URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", LogType.LOG_TYPE_INFO)
            Dim FileExists As Boolean = False
            If ImRunningLocal Then
                FileExists = File.Exists(CurrentAppPath & "\html" & URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png")
            Else
                FileExists = False ' dcor tobefixed
            End If
            If Not FileExists Then
                Dim IConImage As Image
                IConImage = GetPicture(MyIconURL)
                If Not IConImage Is Nothing Then
                    Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                    MyIconURL = URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png"
                    Dim SuccesfullSave As Boolean = False
                    SuccesfullSave = hs.WriteHTMLImage(IConImage, FileArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", True)
                    If Not SuccesfullSave Then
                        If g_bDebug Then Log("Error in myDeviceFinderCallback_DeviceFound for zoneplayer = " & ZoneName & " had error storing Icon at " & FileArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", LogType.LOG_TYPE_ERROR)
                    Else
                        If g_bDebug Then Log("myDeviceFinderCallback_DeviceFound for zoneplayer = " & ZoneName & " stored Icon at " & FileArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", LogType.LOG_TYPE_INFO)
                    End If
                    'IConImage.Save(hs.GetAppPath & "\html" & URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png", ImageFormat)
                    IConImage.Dispose()
                End If
            Else
                MyIconURL = URLArtWorkPath & "PlayerIcon_" & pDevice.ModelNumber.ToString & ".png"
                If g_bDebug Then Log("myDeviceFinderCallback_DeviceFound for zoneplayer = " & ZoneName & " found Icon already stored at " & MyIconURL, LogType.LOG_TYPE_INFO)
            End If
            If HSRefPlayer <> -1 Then
                Dim dv As Scheduler.Classes.DeviceClass
                dv = hs.GetDeviceByRef(HSRefPlayer)
                dv.Image(hs) = MyIconURL
                dv.ImageLarge(hs) = MyIconURL
                WriteStringIniFile(UDN, DeviceInfoIndex.diDeviceIConURL.ToString, MyIconURL)
                If g_bDebug Then Log("Device Finder Call Back  for zoneplayer = " & ZoneName & " added Image = " & MyIconURL, LogType.LOG_TYPE_INFO)
            End If
        Catch ex As Exception
            Log("Error in Device Finder Call Back for zoneplayer = " & ZoneName & ". Could not get ICON info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            Children = pDevice.Children
            If Children IsNot Nothing Then
                If Children.Count > 0 Then
                    For Each objChild In Children
                        If objChild.UniqueDeviceName = ("uuid:" & UDN & "_MS") Then
                            MediaServer = objChild
                            If SuperDebug Then Log(ZoneName & " : MediaServer Friendly Name = " & MediaServer.UniqueDeviceName, LogType.LOG_TYPE_INFO)
                        ElseIf objChild.UniqueDeviceName = ("uuid:" & UDN & "_MR") Then
                            MediaRenderer = objChild
                            If SuperDebug Then Log(ZoneName & " : MediaRenderer Friendly Name = " & MediaRenderer.FriendlyName, LogType.LOG_TYPE_INFO)
                        Else
                            If SuperDebug Then Log(ZoneName & " : Additional service Fiendly Name = " & objChild.UniqueDeviceName, LogType.LOG_TYPE_WARNING) ' for testing proposes
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Log(ZoneName & " : There was a problem in getting MediaServer & MediaRenderer" & ex.Message, LogType.LOG_TYPE_ERROR)
            Log(ZoneName & " : Seach item = uuid:" & UDN & "_MS / _MR", LogType.LOG_TYPE_ERROR)
        End Try

        If pDevice IsNot Nothing Then
            Try
                AudioIn = pDevice.Services.Item("urn:upnp-org:serviceId:AudioIn")
                If SuperDebug And (AudioIn IsNot Nothing) Then Log("AudioIn service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                AlarmClock = pDevice.Services.Item("urn:upnp-org:serviceId:AlarmClock")
                If SuperDebug And (AlarmClock IsNot Nothing) Then Log("AlarmClock service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                DeviceProperties = pDevice.Services.Item("urn:upnp-org:serviceId:DeviceProperties")
                If SuperDebug And (DeviceProperties IsNot Nothing) Then Log("DeviceProperties service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                AVTransport = pDevice.Services.Item("urn:upnp-org:serviceId:AVTransport")
                If SuperDebug And (AVTransport IsNot Nothing) Then Log("AVTransport service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                ZoneGroupTopology = pDevice.Services.Item("urn:upnp-org:serviceId:ZoneGroupTopology")
                If SuperDebug And (ZoneGroupTopology IsNot Nothing) Then Log("ZoneGroupTopology service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                MusicServices = pDevice.Services.Item("urn:upnp-org:serviceId:MusicServices")
                If SuperDebug And (MusicServices IsNot Nothing) Then Log("MusicServices service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
        End If
        If Not MediaRenderer Is Nothing Then
            Try
                AVTransport = MediaRenderer.Services.Item("urn:upnp-org:serviceId:AVTransport")
                If SuperDebug And (AVTransport IsNot Nothing) Then Log("AVTransport service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                RenderingControl = MediaRenderer.Services.Item("urn:upnp-org:serviceId:RenderingControl")
                If SuperDebug And (RenderingControl IsNot Nothing) Then Log("RenderingControl service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try ' added 2/24/2019 in v3.1.0.29
                QueueService = MediaRenderer.Services.Item("urn:sonos-com:serviceId:Queue")
                If SuperDebug And (QueueService IsNot Nothing) Then Log("Queue service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try ' added 2/24/2019 in v3.1.0.29
                VirtualLineIn = MediaRenderer.Services.Item("urn:upnp-org:serviceId:VirtualLineIn")
                If SuperDebug And (VirtualLineIn IsNot Nothing) Then Log("VirtualLineIn service found for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
        End If

        If Not MediaServer Is Nothing Then
            Try
                ContentDirectory = MediaServer.Services.Item("urn:upnp-org:serviceId:ContentDirectory")
                If SuperDebug And (ContentDirectory IsNot Nothing) Then Log("ContentDirectory service added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in myDeviceFinderCallback_DeviceFound for MediaServer.Services.ContentDirectory for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If MyZoneModel.ToUpper = "SUB" Then
            If g_bDebug Then Log("Device Finder Call Back for zoneplayer = " & ZoneName & " is ignoring found Sonos SUB device with UDN =  " & pDevice.UniqueDeviceName & " Friendly Name = " & pDevice.FriendlyName, LogType.LOG_TYPE_WARNING)
            DeviceStatus = "Online"
            Exit Sub
        End If
        If Not ZoneGroupTopology Is Nothing Then
            Try
                ZoneGroupTopology.AddCallback(myZonegroupTopologyCallback)
                If g_bDebug Then Log("ZoneGroupTopology ControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding ZoneGroupTopology Control Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If Not ZoneGroupTopology Is Nothing Then
            GetZoneGroupState()
        End If

        If MyZoneIsPlaybarSlave Then
            RenderingControl = Nothing    ' added 7/13/2019 in v3.1.0.34 to prevent errors for players that are paired to a play bar/beam/base
            QueueService = Nothing
            VirtualLineIn = Nothing
        End If

        If Not AudioIn Is Nothing Then
            Try
                AudioIn.AddCallback(myAudioInCallback)
                If g_bDebug Then Log("AudioInCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding AudioIn Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If Not DeviceProperties Is Nothing Then
            Try
                DeviceProperties.AddCallback(myDevicePropertiesCallback)
                If g_bDebug Then Log("DevicePropertiesCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding DeviceProperties Call Back for zoneplayer = " & ZoneName & ". Error" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If Not AVTransport Is Nothing Then
            Try
                AVTransport.AddCallback(myAVTransportCallback)
                If g_bDebug Then Log("AvTransportCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding Transport Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If RenderingControl IsNot Nothing Then
            Try
                RenderingControl.AddCallback(myRenderingControlCallback)
                If g_bDebug Then Log("RenderingControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                If AVTransport IsNot Nothing Then Log("Error in Adding RenderingControl Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR) ' if no AVTransport, this is a 5:1 surround sound speaker
            End Try
        End If
        If Not ContentDirectory Is Nothing Then
            Try
                ContentDirectory.AddCallback(myContentDirectoryCallback)
                If g_bDebug Then Log("ContentDirectoryControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding ContentDirectoryControl Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If QueueService IsNot Nothing Then  ' added 2/24/2019 in v3.1.0.29
            Try
                QueueService.AddCallback(myQueueServiceCallback)
                If g_bDebug Then Log("QueueService added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding QueueService Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If VirtualLineIn IsNot Nothing Then  ' added 2/24/2019 in v3.1.0.29 
            Try
                VirtualLineIn.AddCallback(myVirtualLineInCallback)
                If g_bDebug Then Log("VirtualLineIn added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding VirtualLineIn Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If Not AlarmClock Is Nothing Then
            Try
                AlarmClock.AddCallback(myAlarmClockCallback)
                If g_bDebug Then Log("AlarmClock ControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in Adding AlarmClock Control Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If

        If DeviceProperties Is Nothing Then '   If AVTransport Is Nothing Or DeviceProperties Is Nothing Then
            Log("Error in myDeviceFinderCallback_DeviceFound for zoneplayer = " & ZoneName & ". Some crucial services are not on-line. The plug-in cannot work for this zone", LogType.LOG_TYPE_ERROR)
        Else
            DeviceStatus = "Online"
        End If
        If Not RenderingControl Is Nothing Then
            Try
                MyCurrentMasterVolumeLevel = GetVolumeLevel("Master")
                SetVolume = MyCurrentMasterVolumeLevel
                MyLeftVolume = GetVolumeLevel("LF")
                MyRightVolume = GetVolumeLevel("RF")
                UpdateBalance()
                SetMuteState = GetMuteState("Master")
                SetLoudness = GetLoudnessState("Master")
            Catch ex As Exception
            End Try
        End If
        If Not AudioIn Is Nothing Then
            'If ZoneModel = "WD100" Then
            ' we need to find out at start up whether something is docked. The Zone player doesn't fire off automatically
            'End If
            GetAudioInputAttributes()
        End If
        If Not DeviceProperties Is Nothing Then
            Try
                If ZoneModel <> "WD100" Then
                    Dim InArg(0) As String
                    Dim OutArg(2) As String
                    Dim CurrentZoneName As String = ""
                    Try
                        MyUPnPDevice.Services.Item("urn:upnp-org:serviceId:DeviceProperties").InvokeAction("GetZoneAttributes", InArg, OutArg)
                        CurrentZoneName = OutArg(0) ' this is CurrentZoneName
                        If g_bDebug Then Log("DirectConnect found Zone Name = " & CurrentZoneName, LogType.LOG_TYPE_INFO)
                    Catch ex As Exception
                        If g_bDebug Then Log("Error in DirectConnect while getting ZoneAttributes with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    If (CurrentZoneName <> "") And ZoneName <> CurrentZoneName Then ZoneNameChanged(CurrentZoneName)
                End If
            Catch ex As Exception
                If g_bDebug Then Log("Warning in DirectConnect for zoneplayer = " & ZoneName & " when getting the ChannelMapSet/ZoneName with Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
            End Try
        End If
        Try
            'IPAddress = GetIPAddress()
            If g_bDebug Then Log("IPAddress for zoneplayer = " & ZoneName & "= " & IPAddress, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Issue in myDeviceFinderCallback_DeviceFound Calling GetIPAddress for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            MySWVersion = GetSoftwareVersion()
            If g_bDebug Then Log("SW Version for zoneplayer = " & ZoneName & "= " & MySWVersion.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Issue in myDeviceFinderCallback_DeviceFound Calling GetSoftwareVersion for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If Not AVTransport Is Nothing Then
            GetPlayMode()
            UpdateHS(True)
        End If
    End Sub

    Public Sub DeviceLostCallBack()
        ' tobefixed dcor
        'If DockedDeviceToFind = FindData Then
        'DockedDeviceCallbackDeviceLost(FindData, bstrUDN)
        'Exit Sub
        'End If
        Try
            ' Device is Removed or Lost
            Log("ZonePlayer " & ZoneName & " has been disconnected from the network in DeviceLostCallBack.", LogType.LOG_TYPE_WARNING)
            DeviceStatus = "Offline"
            DestroyObjects(False)
        Catch ex As Exception
            Log("Well we were here in the Device Lost Callback proc but messed up along the way for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub DeviceAliveCallback()
        Try
            ' Device is Removed or Lost
            Log("ZonePlayer " & ZoneName & " just became alive on the network in DeviceAliveCallback.", LogType.LOG_TYPE_WARNING)
            'Reachable()
        Catch ex As Exception
            Log("Error in myDeviceFinderCallback_DeviceAlive for ZonePlayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub GetZoneGroupState()
        If ZoneGroupTopology Is Nothing Then Exit Sub
        If g_bDebug Then Log("GetZoneGroupState called for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            ZoneGroupTopology.InvokeAction("GetZoneGroupState", InArg, OutArg)
            ProcessZoneGroupState(OutArg(0))
        Catch ex As Exception
            If g_bDebug Then Log("Error in GetZoneGroupState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub ProcessZoneGroupState(inVar As String)
        If g_bDebug Then Log("ProcessZoneGroupState called for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)

        ' example
        '<ZoneGroups>
        '   <ZoneGroup Coordinator="RINCON_000E5860905A01400" ID="RINCON_000E5860905A01400:0">
        '   <ZoneGroupMember UUID="RINCON_000E5860905A01400" 
        '       Location="http://192.168.1.109:1400/xml/zone_dock.xml" 
        '       ZoneName="Wireless Dock" Icon="x-rincon-roomicon:dock" 
        '       Invisible="1" 
        '       IsZoneBridge="1" 
        '       SoftwareVersion="14.4-33290d" 
        '       MinCompatibleVersion="14.0-00000" 
        '       BootSeq="25"/>
        '</ZoneGroup>
        '<ZoneGroup Coordinator="RINCON_000E5825227A01400" ID="RINCON_000E5825227A01400:105">
        '   <ZoneGroupMember UUID="RINCON_000E5825227A01400" 
        '       Location="http://192.168.1.102:1400/xml/zone_player.xml" 
        '      ZoneName="Master Bedroom" Icon="x-rincon-roomicon:masterbedroom" 
        '       SoftwareVersion="14.4-33290" 
        '       MinCompatibleVersion="13.0-00000" 
        '       BootSeq="72"/>
        '</ZoneGroup>
        '<ZoneGroup Coordinator="RINCON_000E5832D2D401400" ID="RINCON_000E5832D2D401400:34">
        '   <ZoneGroupMember UUID="RINCON_000E5832D2D401400" 
        '       Location="http://192.168.1.108:1400/xml/zone_player.xml" 
        '      ZoneName="Patio" Icon="x-rincon-roomicon:patio" 
        '       SoftwareVersion="14.4-33290" 
        '       MinCompatibleVersion="13.0-00000" 
        '       BootSeq="44"/>
        '</ZoneGroup>
        '<ZoneGroup Coordinator="RINCON_000E5833F3CC01400" ID="RINCON_000E5833F3CC01400:61">
        '<ZoneGroupMember UUID="RINCON_000E5833F3CC01400" Location="http://192.168.1.101:1400/xml/zone_player.xml" 
        '       ZoneName="Kitchen" Icon="x-rincon-roomicon:kitchen" 
        '       SoftwareVersion="14.4-33290" 
        '       MinCompatibleVersion="13.0-00000" 
        '       BootSeq="38"/>
        '</ZoneGroup>
        '<ZoneGroup Coordinator="RINCON_000E5824C3B001400" ID="RINCON_000E5824C3B001400:55">
        '   <ZoneGroupMember UUID="RINCON_000E5824C3B001400" Location="http://192.168.1.103:1400/xml/zone_player.xml" 
        '       ZoneName="Family Room" Icon="x-rincon-roomicon:family" 
        '       SoftwareVersion="14.4-33290" 
        '       MinCompatibleVersion="13.0-00000" '
        '       BootSeq="66"/>
        '   </ZoneGroup>
        '</ZoneGroups>
        '

        Dim xmlData As New XmlDocument
        Try
            xmlData.LoadXml(inVar.ToString)
        Catch ex As Exception
            Log("Error in ProcessZoneGroupState loading xml for ZonePlayer - " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim ZoneGroupsList As XmlNodeList = Nothing
        Dim ZoneGroupList As XmlNode = Nothing
        Try ' added 7/12/2019 in V3.1.0.31
            ZoneGroupsList = xmlData.DocumentElement.GetElementsByTagName("ZoneGroups") ' changed on 7/12/2019 in V3.1.0.31 from "ChildNodes"
        Catch ex As Exception
            Log("Error in ProcessZoneGroupState loading ZoneGroups xml for ZonePlayer - " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        If ZoneGroupsList Is Nothing Then
            Log("Error in ProcessZoneGroupState there are no ZoneGroups in xml for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If
        Try
            ZoneGroupList = ZoneGroupsList.Item(0)
            ZoneGroupsList = ZoneGroupList.ChildNodes()
        Catch ex As Exception
            Log("Error in ProcessZoneGroupState finding ZoneGroups in xml for ZonePlayer - " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        If ZoneGroupList Is Nothing Then
            Log("Error in ProcessZoneGroupState there is no ZoneGroup in xml for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End If

        Try
            If ZoneGroupsList.Count = 0 Then Exit Try
            'Parse through all nodes
            For Each outerNode As XmlNode In ZoneGroupsList
                If outerNode.Name.ToUpper = "ZONEGROUP" Then
                    ' rewritten 7/14/2019 in version v3.1.0.34 because S18 players are following some new schema as to who is master (right player)
                    Dim ZoneGroupCoordinatorUDN As String = ""
                    Try
                        ZoneGroupCoordinatorUDN = outerNode.Attributes("Coordinator").Value
                    Catch ex As Exception
                    End Try
                    Dim ZoneGroupID As String = ""
                    Try
                        ZoneGroupID = outerNode.Attributes("ID").Value
                    Catch ex As Exception
                    End Try
                    Try
                        If SuperDebug Then Log("ZonePlayer = " & ZoneName & " ProcessZoneGroupState ----->ZoneGroup Coordinator = " & ZoneGroupCoordinatorUDN & " ID = " & ZoneGroupID, LogType.LOG_TYPE_INFO)
                    Catch ex As Exception
                    End Try
                    ' now we need to figure out whether this zonegroup is
                    ' a/ just a player by itself so it only has itself as member (ZoneGroupMember)
                    ' b/ a bunch of grouped players that all play the same content (multiple ZonegroupMember, no ChannelMapSet & no HTSatChanMapSet)
                    ' c/ a stereo pair such as two s5, s1, one ... (multiple ZonegroupMember with ChannelMapSet)
                    ' d/ a HomeTheater (HT) pairing such as playbar/beam/base with S1, SubWoofer as satellite members (multiple Satellite with HTSatChanMapSet)
                    ' 
                    ' There are two scenarios we need to extract from this data
                    ' a/ this player instance is ZoneGroupCoordinator and we need to mark up out local data reflecting that this player is master of something
                    ' b/ this player is a member of another group but it is not the coordinator so we also need to mark up our local data to reflect this player is a slave of something
                    If outerNode.HasChildNodes Then
                        For Each Level1Node As XmlNode In outerNode.ChildNodes
                            Dim ZoneGroupMember As String = ""
                            If Level1Node.Name.ToUpper = "ZONEGROUPMEMBER" Then
                                Try
                                    ZoneGroupMember = Level1Node.Attributes("UUID").Value
                                    If SuperDebug Then Log("ProcessZoneGroupState ---------->ZoneGroupMember UUID = " & ZoneGroupMember, LogType.LOG_TYPE_INFO)
                                Catch ex As Exception
                                End Try
                                Dim HTSatChanMapSet As String = ""
                                Dim ChannelMapSet As String = ""
                                If ZoneGroupMember = UDN Then
                                    ' so this entry is about THIS player
                                    ' if we now check whether the ZoneGroupCoordinatorUDN is this player as well, we know it is master of something. Now we need to check what
                                    Try
                                        HTSatChanMapSet = Level1Node.Attributes("HTSatChanMapSet").Value    ' HomeTheater settings. The coordinator has a full view. Members only theirs and coordinator
                                    Catch ex As Exception
                                    End Try
                                    Try
                                        ChannelMapSet = Level1Node.Attributes("ChannelMapSet").Value    ' StereoPairing setting.
                                    Catch ex As Exception
                                    End Try
                                    If SuperDebug Then Log("ProcessZoneGroupState ------------>matching UDNs at Memberlevel with HTSatChanMapSet = " & HTSatChanMapSet & " and ChannelMapSet = " & ChannelMapSet, LogType.LOG_TYPE_INFO)
                                    If MyHTSatChanMapSet <> HTSatChanMapSet Then
                                        PlaybarPairingChanged(HTSatChanMapSet, ZoneGroupCoordinatorUDN)
                                        If g_bDebug Then Log("ProcessZoneGroupState for ZonePlayer = " & ZoneName & " and UDN = " & UDN & " updated HTSatChanMapSet = " & HTSatChanMapSet, LogType.LOG_TYPE_INFO)
                                    End If
                                    If MyChannelMapSet <> ChannelMapSet Then
                                        ZonePairingChanged(ChannelMapSet, ZoneGroupCoordinatorUDN)
                                        If g_bDebug Then Log("ProcessZoneGroupState for ZonePlayer = " & ZoneName & " and UDN = " & UDN & " updated ChannelMapSet = " & ChannelMapSet, LogType.LOG_TYPE_INFO)
                                    End If
                                End If
                                If Level1Node.HasChildNodes Then
                                    For Each Level2Node As XmlNode In Level1Node.ChildNodes
                                        If Level2Node.Name.ToUpper = "SATELLITE" Then
                                            Dim SatelliteZoneMember As String = ""
                                            Try
                                                SatelliteZoneMember = Level2Node.Attributes("UUID").Value
                                                If SuperDebug Then Log("ProcessZoneGroupState ---------->ZoneGroupMember UUID = " & SatelliteZoneMember, LogType.LOG_TYPE_INFO)
                                            Catch ex As Exception
                                            End Try
                                            If SatelliteZoneMember = UDN Then
                                                ' So this entry is our player. It can now only be a slave of a HT pairing
                                                Try
                                                    HTSatChanMapSet = Level2Node.Attributes("HTSatChanMapSet").Value
                                                Catch ex As Exception
                                                End Try
                                                Try ' I don't think this is a valid case. When players are paired in Stereo they will show up as a ZoneGroupMember not Sattelite but no harm to leave it here
                                                    ChannelMapSet = Level2Node.Attributes("ChannelMapSet").Value
                                                Catch ex As Exception
                                                End Try
                                                If SuperDebug Then Log("ProcessZoneGroupState ------------>matching UDNs at Sattelite level with HTSatChanMapSet = " & HTSatChanMapSet & " and ChannelMapSet = " & ChannelMapSet, LogType.LOG_TYPE_INFO)
                                                If MyHTSatChanMapSet <> HTSatChanMapSet Then
                                                    PlaybarPairingChanged(HTSatChanMapSet, ZoneGroupCoordinatorUDN)
                                                    If g_bDebug Then Log("ProcessZoneGroupState for ZonePlayer = " & ZoneName & " and UDN = " & UDN & " updated HTSatChanMapSet = " & HTSatChanMapSet, LogType.LOG_TYPE_INFO)
                                                End If
                                                If MyChannelMapSet <> ChannelMapSet Then
                                                    ZonePairingChanged(ChannelMapSet, ZoneGroupCoordinatorUDN)
                                                    If g_bDebug Then Log("ProcessZoneGroupState for ZonePlayer = " & ZoneName & " and UDN = " & UDN & " updated ChannelMapSet = " & ChannelMapSet, LogType.LOG_TYPE_INFO)
                                                End If
                                            End If
                                        Else
                                            If g_bDebug Then Log("Warning in ProcessZoneGroupState for ZonePlayer = " & ZoneName & " and UDN = " & UDN & " found unknown Level2Node = " & Level2Node.Name, LogType.LOG_TYPE_WARNING)
                                        End If
                                    Next
                                End If
                            Else
                                If g_bDebug Then Log("Warning in ProcessZoneGroupState for ZonePlayer = " & ZoneName & " and UDN = " & UDN & " found unknown Level1Node = " & Level1Node.Name, LogType.LOG_TYPE_WARNING)
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in ProcessZoneGroupState for ZonePlayer = " & ZoneName & " processing XML with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MyZoneGroupState = inVar
    End Sub

    Public Sub UpdateHS(ByVal GetUpdate As Boolean)
        Try
            Dim TrackInfo
            'Dim TransportState As String()
            'TransportState = {"", "0", "0:00:00", "", "", "", "", "", "", "", "", NoArtPath, NoArtPath, "", ""}
            If GetUpdate Then
                TrackInfo = GetCurrentTrackInfo() ' this will force an update of the info
                'If ZoneSource <> "Linked" Then TransportState(0) = GetTransportState()
            Else
                TrackInfo = MyCurrentTrackInfo
                'TransportState(0) = MyCurrentTransportState
            End If

            If GetUpdate And ZoneSource <> "Linked" Then ' if Linked, the master zone will update the state
                PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState)
                PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                Track = TrackInfo(2)
                Artist = TrackInfo(0)
                Album = TrackInfo(1)
                MyPreviousArtist = Artist
                MyPreviousTrack = Track
                MyPreviousAlbum = Album
                ArtworkURL = TrackInfo(4)
                SetHSPlayerInfo()
                hs.SetDeviceValueByRef(HSRefPlayer, CurrentPlayerState, True)
                If g_bDebug Then Log("HS updated in UpdateHS. HSRef = " & HSRefPlayer, LogType.LOG_TYPE_INFO)
            End If

            If MyZoneIsSourceForLinkedZone Then
                ' Update all linked Zones
                Dim TargetZones() = {""}
                Dim TargetZone As String
                Dim TempMusicApi As HSPI
                Dim AlreadyExist As Boolean = False
                Dim TempHSDeviceCodeTransport As String
                Dim TempTransportInfo As String = ""
                Try
                    TargetZones = Split(MyTargetZoneLinkedList, ";")
                Catch ex As Exception
                End Try
                For Each TargetZone In TargetZones
                    Try
                        TempMusicApi = MyHSPIControllerRef.GetAPIByUDN(TargetZone.ToString)
                        TempMusicApi.Track = TrackInfo(2)
                        TempMusicApi.Artist = TrackInfo(0)
                        TempMusicApi.Album = TrackInfo(1)
                        TempMusicApi.NextTrack = NextTrack
                        TempMusicApi.NextAlbum = NextAlbum
                        TempMusicApi.NextArtist = NextArtist
                        TempMusicApi.RadiostationName = RadiostationName
                        TempMusicApi.MyZoneSourceExt = MyZoneSourceExt
                        TempMusicApi.ArtworkURL = TrackInfo(4)
                        TempMusicApi.NextArtworkURL = NoArtPath
                        TempMusicApi.SetTrackLength(MyTrackLength)
                        TempMusicApi.SetPlayerPosition(MyPlayerPosition)
                        TempMusicApi.SetNbrOfTracks(MyNbrOfTracksInQueue)
                        TempMusicApi.SetTrackNbr(MyTrackInQueueNbr)
                        TempMusicApi.MusicService = "Linked to - " & ZoneName
                        TempMusicApi.CurrentPlayerState = CurrentPlayerState
                        TempMusicApi.SetShuffleState(MyShuffleState)
                        If g_bDebug Then Log("UpdateHS is updating other linked zones. SourceZone=" & ZoneName & ". TargetZone=" & TargetZone.ToString, LogType.LOG_TYPE_INFO)
                        ' notify HS if they have the callback linked
                        TempMusicApi.PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState)
                        TempMusicApi.PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                        TempHSDeviceCodeTransport = MyHSPIControllerRef.GetHSDeviceReference(TargetZone.ToString)
                        TempMusicApi.SetHSPlayerInfo()
                        hs.SetDeviceValueByRef(TempHSDeviceCodeTransport, CurrentPlayerState, True)
                        If g_bDebug Then Log("HS updated in UpdateHS. HS Code = " & TempHSDeviceCodeTransport & " Info = " & TempTransportInfo, LogType.LOG_TYPE_INFO)
                    Catch ex As Exception
                        ' we have this pesky error at start up trying to update an HS zone that hasn't been discovered yet. Use the interface status to suppress the error
                        If gInterfaceStatus = ERR_NONE Then Log("Error in UpdateHS updating other linked zones. SourceZone=" & ZoneName & ". TargetZone=" & TargetZone.ToString & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Next
                TargetZones = Nothing
                TargetZone = Nothing
                TempMusicApi = Nothing
            End If
        Catch ex As Exception
            Log("Error in UpdateHS for Zone = " & ZoneName & ". Couldn't get transport Info with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function ConnectWireLessDock(ByVal inUDN As String) As String
        ConnectWireLessDock = ""
        Properties = {"", "", "", "", ""}
        AudioInState = {"", "", 0, False, False, False, False}
        Log("ConnectWireLessDock is connecting to ZonePlayer - " & ZoneName & ": UDN = " & inUDN, LogType.LOG_TYPE_INFO)
        Dim MyDevice As MyUPnPDevice = Nothing
        Try
            MyDevice = MySSDPDevice.Item(inUDN)
        Catch ex As Exception
            If g_bDebug Then Log("Error in ConnectWireLessDock for ZonePlayer - " & ZoneName & " and UDN = " & inUDN & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            ConnectWireLessDock = "Failed"
            Exit Function
        End Try
        If MyDevice IsNot Nothing Then
            If g_bDebug Then Log("ConnectWireLessDock: Connecting to ZonePlayer - " & ZoneName & ": " & inUDN, LogType.LOG_TYPE_INFO)
            ConnectWireLessDock = "OK"
            ConnectToIPod = False
            DockedDeviceCallbackDeviceFound(MyDevice)
        Else
            ConnectWireLessDock = "Failed"
        End If
    End Function

    Public Function Connect(ByVal inUDN As String) As String
        Connect = ""
        ' Initialize all the State/Info Variables
        Properties = {"", "", "", "", ""}
        AudioInState = {"", "", 0, False, False, False, False}
        If g_bDebug Then Log("Connect: Connecting to ZonePlayer - " & ZoneName & ": UDN = " & inUDN, LogType.LOG_TYPE_INFO)
        Dim MyDevice As MyUPnPDevice = Nothing
        Try
            MyDevice = MySSDPDevice.Item(inUDN)
        Catch ex As Exception
            If g_bDebug Then Log("Error in Connect for ZonePlayer - " & ZoneName & " and UDN = " & inUDN & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Connect = "Failed"
            Exit Function
        End Try
        If MyDevice IsNot Nothing Then
            UDN = Replace(inUDN, "uuid:", "")
            'If SuperDebug Then Log("Connect called for ZonePlayer - " & ZoneName & " with UDN = " & inUDN & " and Device to find = " & Val(DeviceToFind), LogType.LOG_TYPE_INFO)
            Connect = "OK"
            ConnectPlayer = False
            myDeviceFinderCallback_DeviceFound(MyDevice)
        Else
            Connect = "Failed"
            ' Log("Connect: ERROR: " & ZoneName & " Connecting has failed. Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End If
    End Function

    Public Sub Disconnect(Full As Boolean)
        'If DeviceStatus = "Offline" Then Exit Sub ' already disconnected removed in HS3 for testing purposes
        DeviceStatus = "Offline"
        Log("Disconnect: Disconnected from ZonePlayer - " & ZoneName & " with Full = " & Full.ToString, LogType.LOG_TYPE_INFO)
        DestroyObjects(Full)
    End Sub

    Public Sub SetAdministrativeState(Active As Boolean)
        If g_bDebug Then Log("SetAdministrativeState called for device - " & ZoneName & " and Active = " & Active.ToString, LogType.LOG_TYPE_INFO)
        If Active Then
            MyAdminStateActive = True
            WriteBooleanIniFile(UDN, DeviceInfoIndex.diAdminState.ToString, True)
            Dim PlayerType As String = GetStringIniFile(UDN, DeviceInfoIndex.diSonosPlayerType.ToString, "")
            If GetStringIniFile(UDN, DeviceInfoIndex.diSonosPlayerType.ToString, "").ToUpper = "SUB" Or GetStringIniFile(UDN, DeviceInfoIndex.diSonosPlayerType.ToString, "").ToUpper = "BRIDGE" Then
                Exit Sub
            End If
            HSRefTrack = GetIntegerIniFile(UDN, DeviceInfoIndex.diTrackHSRef.ToString, -1)
            If HSRefTrack = -1 Then
                HSRefTrack = CreateHSDevice(DeviceInfoIndex.diTrackHSRef, "Track")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diTrackHSRef.ToString, HSRefTrack)
            End If
            HSRefNextTrack = GetIntegerIniFile(UDN, DeviceInfoIndex.diNextTrackHSRef.ToString, -1)
            If HSRefNextTrack = -1 And ZoneModel <> "WD100" Then
                HSRefNextTrack = CreateHSDevice(DeviceInfoIndex.diNextTrackHSRef, "Next Track")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diNextTrackHSRef.ToString, HSRefNextTrack)
            End If
            HSRefArtist = GetIntegerIniFile(UDN, DeviceInfoIndex.diArtistHSRef.ToString, -1)
            If HSRefArtist = -1 Then
                HSRefArtist = CreateHSDevice(DeviceInfoIndex.diArtistHSRef, "Artist")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diArtistHSRef.ToString, HSRefArtist)
            End If
            HSRefNextArtist = GetIntegerIniFile(UDN, DeviceInfoIndex.diNextArtistHSRef.ToString, -1)
            If HSRefNextArtist = -1 And ZoneModel <> "WD100" Then
                HSRefNextArtist = CreateHSDevice(DeviceInfoIndex.diNextArtistHSRef, "Next Artist")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diNextArtistHSRef.ToString, HSRefNextArtist)
            End If
            HSRefAlbum = GetIntegerIniFile(UDN, DeviceInfoIndex.diAlbumHSRef.ToString, -1)
            If HSRefAlbum = -1 Then
                HSRefAlbum = CreateHSDevice(DeviceInfoIndex.diAlbumHSRef, "Album")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diAlbumHSRef.ToString, HSRefAlbum)
            End If
            HSRefNextAlbum = GetIntegerIniFile(UDN, DeviceInfoIndex.diNextAlbumHSRef.ToString, -1)
            If HSRefNextAlbum = -1 And ZoneModel <> "WD100" Then
                HSRefNextAlbum = CreateHSDevice(DeviceInfoIndex.diNextAlbumHSRef, "Next Album")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diNextAlbumHSRef.ToString, HSRefNextAlbum)
            End If
            HSRefArt = GetIntegerIniFile(UDN, DeviceInfoIndex.diArtHSRef.ToString, -1)
            If HSRefArt = -1 Then
                HSRefArt = CreateHSDevice(DeviceInfoIndex.diArtHSRef, "Art")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diArtHSRef.ToString, HSRefArt)
                MyCurrentArtworkURL = "" ' this will make sure that the value is set when calling ArtworkURL
                ArtworkURL = NoArtPath
            End If
            HSRefNextArt = GetIntegerIniFile(UDN, DeviceInfoIndex.diNextArtHSRef.ToString, -1)
            If HSRefNextArt = -1 And ZoneModel <> "WD100" Then
                HSRefNextArt = CreateHSDevice(DeviceInfoIndex.diNextArtHSRef, "Next Art")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diNextArtHSRef.ToString, HSRefNextArt)
                MyNextAlbumURI = "" ' this will make sure that the value is set  when calling NextArtworkURL
                NextArtworkURL = NoArtPath
            End If
            HSRefPlayState = GetIntegerIniFile(UDN, DeviceInfoIndex.diPlayStateHSRef.ToString, -1)
            If HSRefPlayState = -1 Then
                HSRefPlayState = CreateHSDevice(DeviceInfoIndex.diPlayStateHSRef, "State")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diPlayStateHSRef.ToString, HSRefPlayState)
            Else
                Try
                    Dim dv As Scheduler.Classes.DeviceClass
                    dv = hs.GetDeviceByRef(HSRefPlayState)
                    If dv IsNot Nothing Then
                        MyCurrentPlayerState = dv.devValue(hs)
                    End If
                Catch ex As Exception

                End Try
                hs.GetDeviceByRef(HSRefPlayState)
            End If
            HSRefVolume = GetIntegerIniFile(UDN, DeviceInfoIndex.diVolumeHSRef.ToString, -1)
            If HSRefVolume = -1 Then
                HSRefVolume = CreateHSDevice(DeviceInfoIndex.diVolumeHSRef, "Volume")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diVolumeHSRef.ToString, HSRefVolume)
            End If
            HSRefMute = GetIntegerIniFile(UDN, DeviceInfoIndex.diMuteHSRef.ToString, -1)
            If HSRefMute = -1 Then
                HSRefMute = CreateHSDevice(DeviceInfoIndex.diMuteHSRef, "Mute")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diMuteHSRef.ToString, HSRefMute)
                hs.SetDeviceValueByRef(HSRefMute, msUnmuted, True)
            End If
            HSRefLoudness = GetIntegerIniFile(UDN, DeviceInfoIndex.diLoudnessHSRef.ToString, -1)
            If HSRefLoudness = -1 Then
                HSRefLoudness = CreateHSDevice(DeviceInfoIndex.diLoudnessHSRef, "Loudness")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diLoudnessHSRef.ToString, HSRefLoudness)
                hs.SetDeviceValueByRef(HSRefLoudness, lsLoudnessOff, True)
            End If
            HSRefBalance = GetIntegerIniFile(UDN, DeviceInfoIndex.diBalanceHSRef.ToString, -1)
            If HSRefBalance = -1 Then
                HSRefBalance = CreateHSDevice(DeviceInfoIndex.diBalanceHSRef, "Balance")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diBalanceHSRef.ToString, HSRefBalance)
                hs.SetDeviceValueByRef(HSRefBalance, 0, True)
            End If
            HSRefTrackLength = GetIntegerIniFile(UDN, DeviceInfoIndex.diTrackLengthHSRef.ToString, -1)
            If HSRefTrackLength = -1 Then
                HSRefTrackLength = CreateHSDevice(DeviceInfoIndex.diTrackLengthHSRef, "Track Length")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diTrackLengthHSRef.ToString, HSRefTrackLength)
                'hs.SetDeviceString(HSRefTrackLength, "0", True)
            End If
            HSRefTrackPos = GetIntegerIniFile(UDN, DeviceInfoIndex.diTrackPosHSRef.ToString, -1)
            If HSRefTrackPos = -1 Then
                HSRefTrackPos = CreateHSDevice(DeviceInfoIndex.diTrackPosHSRef, "Track Position")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diTrackPosHSRef.ToString, HSRefTrackPos)
                'hs.SetDeviceString(HSRefTrackPos, "0", True)
            End If
            HSRefRadiostationName = GetIntegerIniFile(UDN, DeviceInfoIndex.diRadiostationNameHSRef.ToString, -1)
            If HSRefRadiostationName = -1 And ZoneModel <> "WD100" Then
                HSRefRadiostationName = CreateHSDevice(DeviceInfoIndex.diRadiostationNameHSRef, "Radiostation Name")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diRadiostationNameHSRef.ToString, HSRefRadiostationName)
            End If
            HSRefDockDeviceName = GetIntegerIniFile(UDN, DeviceInfoIndex.diDockDeviceNameHSRef.ToString, -1)
            If HSRefDockDeviceName = -1 And ZoneModel = "WD100" Then
                HSRefDockDeviceName = CreateHSDevice(DeviceInfoIndex.diDockDeviceNameHSRef, "Docked Device Name")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diDockDeviceNameHSRef.ToString, HSRefDockDeviceName)
            End If
            'HSRefTrackDescr = GetIntegerIniFile(UDN, DeviceInfoIndex.diTrackDescrHSRef.ToString, -1)
            'If HSRefTrackDescr = -1 Then
            'HSRefTrackDescr = CreateHSDevice(DeviceInfoIndex.diTrackDescrHSRef, "Track Desc")
            'WriteIntegerIniFile(UDN, DeviceInfoIndex.diTrackDescrHSRef.ToString, HSRefTrackDescr)
            'End If
            HSRefRepeat = GetIntegerIniFile(UDN, DeviceInfoIndex.diRepeatHSRef.ToString, -1)
            If HSRefRepeat = -1 Then
                HSRefRepeat = CreateHSDevice(DeviceInfoIndex.diRepeatHSRef, "Repeat")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diRepeatHSRef.ToString, HSRefRepeat)
                hs.SetDeviceValueByRef(HSRefRepeat, rsnoRepeat, True)
            End If
            HSRefShuffle = GetIntegerIniFile(UDN, DeviceInfoIndex.diShuffleHSRef.ToString, -1)
            If HSRefShuffle = -1 Then
                HSRefShuffle = CreateHSDevice(DeviceInfoIndex.diShuffleHSRef, "Shuffle")
                WriteIntegerIniFile(UDN, DeviceInfoIndex.diShuffleHSRef.ToString, HSRefShuffle)
                hs.SetDeviceValueByRef(HSRefShuffle, ssNoShuffle, True)
            End If
            'HSRefGenre = GetIntegerIniFile(UDN, DeviceInfoIndex.diGenreHSRef.ToString, -1)
            'If HSRefGenre = -1 Then
            'HSRefGenre = CreateHSDevice(DeviceInfoIndex.diGenreHSRef, "Genre")
            'WriteIntegerIniFile(UDN, DeviceInfoIndex.diGenreHSRef.ToString, HSRefGenre)
            'End If
        Else
            MyAdminStateActive = False
            WriteBooleanIniFile(UDN, DeviceInfoIndex.diAdminState.ToString, False)
        End If
    End Sub

    Private Function CreateHSDevice(ByVal DeviceType As DeviceInfoIndex, DevString As String) As Integer
        CreateHSDevice = -1
        Dim dv As Scheduler.Classes.DeviceClass
        Dim DevName As String = DevString 'ZoneName & " - " & DevString
        Dim dvParent As Scheduler.Classes.DeviceClass = Nothing
        Dim HSRef As Integer = -1
        Dim RoomIcon As String = GetStringIniFile(UDN, DeviceInfoIndex.diRoomIcon.ToString, "")
        Try
            HSRef = hs.NewDeviceRef(DevName)
            If g_bDebug Then Log("CreateHSDevice for Zone = " & ZoneName & " created deviceType " & DeviceType.ToString & " with Ref " & HSRef.ToString, LogType.LOG_TYPE_INFO)
            ' Force HomeSeer to save changes to devices and events so we can find our new device
            hs.SaveEventsDevices()
            dv = hs.GetDeviceByRef(HSRef)
            dv.Interface(hs) = sIFACE_NAME
            dv.InterfaceInstance(hs) = MainInstance
            dv.Location2(hs) = tIFACE_NAME
            dv.Location(hs) = ZoneName
            'dv.MISC_Set(hs, Enums.dvMISC.HIDDEN)
            dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)
            Dim DT As New DeviceTypeInfo
            DT.Device_API = DeviceTypeInfo.eDeviceAPI.Media
            If RoomIcon <> "" Then
                dv.Image(hs) = ImagesPath & RoomIcon & ".png"
                dv.ImageLarge(hs) = ImagesPath & RoomIcon & ".png"
                If g_bDebug Then Log("CreateHSDevice added image  " & ImagesPath & RoomIcon & ".png", LogType.LOG_TYPE_INFO)
            End If
            Select Case DeviceType
                Case DeviceInfoIndex.diTrackHSRef
                    dv.Device_Type_String(hs) = "Track"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Track
                    DT.Device_SubType_Description = "Sonos Player Track Name"
                    dv.Address(hs) = "S09"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Track"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Track", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)    ' set high value so stupid lamp/dim symbols don't show when stringvalue is set to empty
                Case DeviceInfoIndex.diNextTrackHSRef
                    dv.Device_Type_String(hs) = "Next Track"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Next Track Name"
                    dv.Address(hs) = "S13"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Track"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Track", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)    ' set high value so stupid lamp/dim symbols don't show when stringvalue is set to empty
                Case DeviceInfoIndex.diArtistHSRef
                    dv.Device_Type_String(hs) = "Artist"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Artist
                    DT.Device_SubType_Description = "Sonos Player Artist Name"
                    dv.Address(hs) = "S11"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Artist"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Artist", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)    ' set high value so stupid lamp/dim symbols don't show when stringvalue is set to empty
                Case DeviceInfoIndex.diNextArtistHSRef
                    dv.Device_Type_String(hs) = "Next Artist"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Next Artist Name"
                    dv.Address(hs) = "S15"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Artist"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Artist", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)
                Case DeviceInfoIndex.diAlbumHSRef
                    dv.Device_Type_String(hs) = "Album"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Album
                    DT.Device_SubType_Description = "Sonos Player Album Name"
                    dv.Address(hs) = "S10"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Album"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Album", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)
                Case DeviceInfoIndex.diNextAlbumHSRef
                    dv.Device_Type_String(hs) = "Next Album"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Next Album Name"
                    dv.Address(hs) = "S14"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Album"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Album", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)
                Case DeviceInfoIndex.diArtHSRef
                    dv.Device_Type_String(hs) = "Art"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player AlbumArt URL"
                    dv.Address(hs) = "S16"
                    CreateArtImagePairs(HSRef)
                    hs.SetDeviceString(HSRef, NoArtPath, True)
                Case DeviceInfoIndex.diNextArtHSRef
                    dv.Device_Type_String(hs) = "Next Art"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player AlbumArt URL"
                    dv.Address(hs) = "S17"
                    CreateArtImagePairs(HSRef)
                    hs.SetDeviceString(HSRef, NoArtPath, True)
                Case DeviceInfoIndex.diPlayStateHSRef
                    dv.Device_Type_String(hs) = "State"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status
                    DT.Device_SubType_Description = "Sonos Player State"
                    dv.Address(hs) = "S01"
                    CreateStatePairs(HSRef)
                    hs.SetDeviceValueByRef(HSRef, player_state_values.Stopped, True)
                Case DeviceInfoIndex.diVolumeHSRef
                    dv.Device_Type_String(hs) = "Volume"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Volume
                    DT.Device_SubType_Description = "Sonos Player Volume"
                    dv.Address(hs) = "S02"
                    CreateVolumePairs(HSRef)
                    hs.SetDeviceValueByRef(HSRef, 0, True)
                Case DeviceInfoIndex.diMuteHSRef
                    dv.Device_Type_String(hs) = "Mute"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Mute State"
                    dv.Address(hs) = "S06"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diMuteHSRef)
                Case DeviceInfoIndex.diLoudnessHSRef
                    dv.Device_Type_String(hs) = "Loudness"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Loudness State"
                    dv.Address(hs) = "S07"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diLoudnessHSRef)
                Case DeviceInfoIndex.diBalanceHSRef
                    dv.Device_Type_String(hs) = "Balance"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Balance State"
                    dv.Address(hs) = "S03"
                    CreateSliderPairs(HSRef, DeviceInfoIndex.diBalanceHSRef)
                    hs.SetDeviceValueByRef(HSRef, 100, False)
                Case DeviceInfoIndex.diTrackLengthHSRef
                    dv.Device_Type_String(hs) = "Track Length"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Track Length"
                    dv.Address(hs) = "S21"
                    hs.SetDeviceValueByRef(HSRef, 0, True)
                    Select Case MyHSTrackPositionFormat
                        Case HSSTrackLengthSettings.TLSSeconds
                            hs.SetDeviceString(HSRef, "0", True)
                        Case HSSTrackLengthSettings.TLSHoursMinutesSeconds
                            hs.SetDeviceString(HSRef, "00:00:00", True)
                    End Select
                Case DeviceInfoIndex.diTrackPosHSRef
                    dv.Device_Type_String(hs) = "Track Position"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Track Position"
                    dv.Address(hs) = "S08"
                    CreateSliderPairs(HSRef, DeviceInfoIndex.diTrackPosHSRef)
                    hs.SetDeviceValueByRef(HSRef, 0, True)
                    Select Case MyHSTrackPositionFormat
                        Case HSSTrackPositionSettings.TPSSeconds
                            hs.SetDeviceString(HSRef, "0", True)
                        Case HSSTrackPositionSettings.TPSHoursMinutesSeconds
                            hs.SetDeviceString(HSRef, "00:00:00", True)
                        Case HSSTrackPositionSettings.TPSPercentage
                            hs.SetDeviceString(HSRef, "0", True)
                    End Select
                Case DeviceInfoIndex.diRadiostationNameHSRef
                    dv.Device_Type_String(hs) = "Radiostation Name"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Radiostation Name"
                    dv.Address(hs) = "S12"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Radio"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Radio", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)
                Case DeviceInfoIndex.diDockDeviceNameHSRef
                    dv.Device_Type_String(hs) = "Docked Device Name"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Docked Device Name"
                    dv.Address(hs) = "S18"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Device"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Device", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)
                Case DeviceInfoIndex.diTrackDescrHSRef
                    dv.Device_Type_String(hs) = "Track Descr"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Status_Additional
                    DT.Device_SubType_Description = "Sonos Player Track Descriptor"
                    dv.Address(hs) = "S20"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Descr"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Descr", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)
                Case DeviceInfoIndex.diRepeatHSRef
                    dv.Device_Type_String(hs) = "Repeat"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Repeat
                    DT.Device_SubType_Description = "Sonos Player Repeat State"
                    dv.Address(hs) = "S04"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diRepeatHSRef)
                    hs.SetDeviceValueByRef(HSRef, rsnoRepeat, True)
                Case DeviceInfoIndex.diShuffleHSRef
                    dv.Device_Type_String(hs) = "Shuffle"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Player_Shuffle
                    DT.Device_SubType_Description = "Sonos Player Shuffle State"
                    dv.Address(hs) = "S05"
                    CreateOnOffTogglePairs(HSRef, DeviceInfoIndex.diShuffleHSRef)
                    hs.SetDeviceValueByRef(HSRef, ssNoShuffle, True)
                Case DeviceInfoIndex.diGenreHSRef
                    dv.Device_Type_String(hs) = "Genre"
                    DT.Device_Type = DeviceTypeInfo.eDeviceType_Media.Media_Genre
                    DT.Device_SubType_Description = "Sonos Player Track Genre"
                    dv.Address(hs) = "S19"
                    Dim Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                    Pair.PairType = VSVGPairType.SingleValue
                    Pair.Value = 1000
                    Pair.Status = "No Genre"
                    hs.DeviceVSP_AddPair(HSRef, Pair)
                    hs.SetDeviceString(HSRef, "No Genre", True)
                    hs.SetDeviceValueByRef(HSRef, 1000, True)
            End Select

            dv.DeviceType_Set(hs) = DT
            dv.Status_Support(hs) = True
            ' This device is a child device, the parent being the root device for the entire security system. 
            ' As such, this device needs to be associated with the root (Parent) device.
            dvParent = hs.GetDeviceByRef(HSRefPlayer)
            If dvParent.AssociatedDevices_Count(hs) < 1 Then
                ' There are none added, so it is OK to add this one.
                dvParent.AssociatedDevice_Add(hs, HSRef)
            Else
                Dim Found As Boolean = False
                For Each ref As Integer In dvParent.AssociatedDevices(hs)
                    If ref = HSRef Then
                        Found = True
                        Exit For
                    End If
                Next
                If Not Found Then
                    dvParent.AssociatedDevice_Add(hs, HSRef)
                Else
                    ' This is an error condition likely as this device's reference ID should not already be associated.
                End If
            End If

            ' Now, we want to make sure our child device also reflects the relationship by adding the parent to
            '   the child's associations.
            dv.AssociatedDevice_ClearAll(hs)  ' There can be only one parent, so make sure by wiping these out.
            dv.AssociatedDevice_Add(hs, dvParent.Ref(hs))
            dv.Relationship(hs) = Enums.eRelationship.Child
            hs.SaveEventsDevices()
            WriteStringIniFile("UPnP HSRef to UDN", HSRef, UDN)
            Return HSRef
        Catch ex As Exception
            Log("Error in CreateHSDevice with Error =  " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub CreateVolumePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        ' add a Down button
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpDown
        Pair.Status = "Down"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 1
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)


        ' add Volume Slider
        Pair = New VSPair(ePairStatusControl.Both)
        Pair.PairType = VSVGPairType.Range
        Pair.Value = vpSlider
        Pair.RangeStart = 0
        Pair.RangeEnd = 100
        Pair.RangeStatusPrefix = "Volume "
        Pair.RangeStatusSuffix = "%"
        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
        Pair.Render_Location.Column = 2
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        ' add an Up button

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpUp
        Pair.Status = "Up"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 3
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

    End Sub

    Private Sub CreateOnOffTogglePairs(HSRef As Integer, DeviceType As DeviceInfoIndex)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        Dim GraphicsPair As VGPair

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpOff
        Pair.Status = "Off"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpOn
        Pair.Status = "On"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = tpToggle
        Pair.Status = "Toggle"
        Pair.Render = Enums.CAPIControlType.Button
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Select Case DeviceType
            Case DeviceInfoIndex.diMuteHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = msMuted
                Pair.Status = "Muted"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "muted.gif".ToLower
                GraphicsPair.Set_Value = msMuted
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = msUnmuted
                Pair.Status = "Unmuted"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "unmuted.gif"
                GraphicsPair.Set_Value = msUnmuted
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diLoudnessHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = lsLoudnessOn
                Pair.Status = "Loudness On"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "loudness.gif"
                GraphicsPair.Set_Value = lsLoudnessOn
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = lsLoudnessOff
                Pair.Status = "Loudness Off"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "loudnessoff.gif"
                GraphicsPair.Set_Value = lsLoudnessOff
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diRepeatHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = rsnoRepeat
                Pair.Status = "No Repeat"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "repeatoff.gif"
                GraphicsPair.Set_Value = rsnoRepeat
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = rsRepeat
                Pair.Status = "Repeat all"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "repeatall.gif"
                GraphicsPair.Set_Value = rsRepeat
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
            Case DeviceInfoIndex.diShuffleHSRef
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = ssNoShuffle
                Pair.Status = "Ordered"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "ordered.gif"
                GraphicsPair.Set_Value = ssNoShuffle
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = ssShuffled
                Pair.Status = "Shuffled"
                hs.DeviceVSP_AddPair(HSRef, Pair)
                GraphicsPair = New VGPair()
                GraphicsPair.PairType = VSVGPairType.SingleValue
                GraphicsPair.Graphic = ImagesPath & "shuffled.gif"
                GraphicsPair.Set_Value = ssShuffled
                hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
        End Select
    End Sub

    Private Sub CreateStatePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        Dim GraphicsPair As VGPair

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psStop
        Pair.Status = "Stop"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 1
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPlay
        Pair.Status = "Play"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 2
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPause
        Pair.Status = "Pause"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 3
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPlayPause
        Pair.Status = "Play-Pause"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Row = 1
        Pair.Render_Location.Column = 4
        hs.DeviceVSP_AddPair(HSRef, Pair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPlaying
        Pair.Status = "Playing"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "playing.gif"
        GraphicsPair.Set_Value = psPlaying
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psStopped
        Pair.Status = "Stopped"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "stopped.gif"
        GraphicsPair.Set_Value = psStopped
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)

        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = psPaused
        Pair.Status = "Paused"
        hs.DeviceVSP_AddPair(HSRef, Pair)
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ImagesPath & "paused.jpg"
        GraphicsPair.Set_Value = psPaused
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
    End Sub

    Private Sub CreateSliderPairs(HSRef As Integer, deviceType As DeviceInfoIndex)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair

        ' add a Dow button
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpDown
        Pair.Status = "Down"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 1
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        ' add Volume Slider
        Pair = New VSPair(ePairStatusControl.Both)
        Pair.PairType = VSVGPairType.Range
        Pair.Render = Enums.CAPIControlType.ValuesRangeSlider
        Select Case deviceType
            Case DeviceInfoIndex.diBalanceHSRef
                Pair.Value = vpSlider
                Pair.RangeStart = -100
                Pair.RangeEnd = 100
                Pair.RangeStatusPrefix = "Balance L (-100) <-> R (+100) : "
                'Pair.RangeStatusSuffix = "%"
            Case DeviceInfoIndex.diTrackPosHSRef
                Pair.Value = vpSlider
                Pair.RangeStart = 0
                Pair.RangeEnd = 200
                'Pair.RangeStatusPrefix = "Volume "
                'Pair.RangeStatusSuffix = "%"

        End Select
        Pair.Render_Location.Column = 2
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)

        ' add an Up button
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = vpUp
        Pair.Status = "Up"
        Pair.Render = Enums.CAPIControlType.Button
        Pair.Render_Location.Column = 3
        Pair.Render_Location.Row = 1
        hs.DeviceVSP_AddPair(HSRef, Pair)
    End Sub

    Private Sub CreateArtImagePairs(HSRef As Integer)
        hs.DeviceVSP_ClearAll(HSRef, True)
        Dim Pair As VSPair
        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
        Pair.PairType = VSVGPairType.SingleValue
        Pair.Value = 1000
        Pair.Status = NoArtPath
        hs.DeviceVSP_AddPair(HSRef, Pair)
        Dim GraphicsPair As VGPair
        GraphicsPair = New VGPair()
        GraphicsPair.PairType = VSVGPairType.SingleValue
        GraphicsPair.Graphic = ""
        GraphicsPair.Set_Value = 1000
        hs.DeviceVGP_AddPair(HSRef, GraphicsPair)
    End Sub


    Public Sub TreatSetIOEx(CC As CAPIControl)
        If g_bDebug Then Log("TreatSetIOEx called for Zone = " & ZoneName & " with  Ref = " & CC.Ref.ToString & ", Index " & CC.CCIndex.ToString & ", controlFlag = " & CC.ControlFlag.ToString &
                 ", ControlString" & CC.ControlString.ToString & ", ControlType = " & CC.ControlType.ToString & ", ControlValue = " & CC.ControlValue.ToString &
                  ", Label = " & CC.Label.ToString, LogType.LOG_TYPE_INFO)
        Dim SonosPlayer As HSPI = Me
        If ZoneModel = "WD100" Then
            ' The WD100 pretty much does everything through its streaming zone
            If CC.Ref = HSRefPlayer And CC.ControlValue = psBuildiPodDB Then
                If DockediPodPlayerName <> "" Then BuildDockedPlayerDatabase(CurrentAppPath & DockedPlayersDBPath & SonosPlayer.DockediPodPlayerName & ".sdb", DockediPodPlayerName)
                Exit Sub
            End If

            Dim DestZonePlayerName As String = ""
            Dim DestZonePlayerUDN As String = ""
            Dim DestZonePlayerHSRef As Integer = -1
            DestZonePlayerUDN = DestinationZone
            Dim DestPlayer As HSPI
            DestPlayer = MyHSPIControllerRef.GetAPIByUDN(DestZonePlayerUDN)
            If DestPlayer Is Nothing Then
                If CC.Ref = HSRefPlayer And (CC.ControlValue = psPlay Or CC.ControlValue = psPlayPause) Then
                    If Not SourcingZone Then
                        ' this indicates that the zoneplayer is not yet linked
                        DestPlayer.PlayURI("x-sonos-dock:" & GetUDN(), "") ' group
                    End If
                    TogglePlay()    ' SetTransportState("Play")
                Else
                    If g_bDebug Then Log("Warning in TreatSetIOEx: Wireless Dock is not connected to streaming zone", LogType.LOG_TYPE_WARNING)
                End If
                Exit Sub
            End If
            SonosPlayer = DestPlayer ' forward to right zone
        End If
        If ZoneIsLinked Then
            ' this zone is sourcing from another zone. If the source zone is a WD100 zone then we need to send Stop/Play/Pause/Shuffle/Repeat commands to that zone
            Dim SourcePlayer As HSPI = MyHSPIControllerRef.GetAPIByUDN(MySourceLinkedZone)
            If SourcePlayer IsNot Nothing Then
                If SourcePlayer.ZoneModel <> "WD100" Then
                    Select Case CC.Ref
                        Case HSRefBalance, HSRefLoudness, HSRefMute, HSRefVolume
                            ' don't do anything right player
                        Case HSRefPlayer
                            Select Case CC.ControlValue
                                Case psPlay, psStop, psPause, psNext, psPrevious, psShuffle, psRepeat, psPlayPause
                                    SonosPlayer = SourcePlayer
                                Case Else
                                    ' right player
                            End Select
                        Case Else ' only update playstate 
                            SonosPlayer = SourcePlayer
                    End Select
                End If
            End If
        End If
        Dim tempVolumeStep As Integer
        Try
            'tempVolumeStep = CInt(data2) tobefixed
            If tempVolumeStep = 0 Then tempVolumeStep = VolumeStep
        Catch ex As Exception
            tempVolumeStep = VolumeStep
        End Try
        ' treat with "local" requests
        Select Case CC.Ref
            Case HSRefBalance
                Select Case CC.ControlValue
                    Case vpDown
                        SonosPlayer.ChangeBalanceLevel("LF", 10)
                    Case vpUp
                        SonosPlayer.ChangeBalanceLevel("RF", 10)
                    Case Else
                        SonosPlayer.SetBalance(CC.ControlValue)
                        ' this should be the slider
                End Select
            Case HSRefLoudness
                Select Case CC.ControlValue
                    Case tpOn
                        SonosPlayer.SetLoudnessState("Master", True)    ' fixed v.12
                    Case tpOff
                        SonosPlayer.SetLoudnessState("Master", False)   ' fixed v.12
                    Case tpToggle
                        SonosPlayer.ToggleLoudnessState("Master")
                End Select
            Case HSRefMute
                Select Case CC.ControlValue
                    Case tpOn
                        SonosPlayer.SetMute("Master", True)
                    Case tpOff
                        SonosPlayer.SetMute("Master", False)
                    Case tpToggle
                        SonosPlayer.ToggleMuteState("Master")
                End Select
            Case HSRefPlayState
                Select Case CC.ControlValue
                    Case psPlay
                        SonosPlayer.SetTransportState("Play")
                    Case psStop
                        SonosPlayer.SetTransportState("Stop")
                    Case psPause
                        SonosPlayer.TogglePause()
                    Case psPlayPause
                        SonosPlayer.TogglePlay()
                End Select
            Case HSRefPlayer
                Select Case CC.ControlValue
                    Case psPlay
                        SonosPlayer.SetTransportState("Play")
                    Case psStop
                        SonosPlayer.SetTransportState("Stop")
                    Case psPause
                        SonosPlayer.TogglePause()
                    Case psPlayPause
                        SonosPlayer.TogglePlay()
                    Case psPrevious
                        SonosPlayer.SetTransportState("Previous")
                    Case psNext
                        SonosPlayer.SetTransportState("Next")
                    Case psShuffle  'Shuffle
                        SonosPlayer.ToggleShuffle()
                    Case psRepeat  ' Repeat
                        SonosPlayer.ToggleRepeat()
                    Case psVolUp  ' Up
                        SonosPlayer.ChangeVolumeLevel("Master", tempVolumeStep)
                    Case psVolDown   ' Down
                        SonosPlayer.ChangeVolumeLevel("Master", -tempVolumeStep)
                    Case psMute   ' mute
                        SonosPlayer.ToggleMuteState("Master")
                    Case psBalanceLeft    ' Left
                        SonosPlayer.ChangeBalanceLevel("LF", 10)
                    Case psBalanceRight   ' Right
                        SonosPlayer.ChangeBalanceLevel("RF", 10)
                    Case psLoudness   ' Loudness
                        SonosPlayer.ToggleLoudnessState("Master")
                    Case Else
                        If CC.ControlValue <= 100 Then
                            SonosPlayer.SetVolumeLevel("Master", CC.ControlValue)
                        ElseIf CC.ControlValue >= 200 Then
                            SonosPlayer.SetBalance(CC.ControlValue - vpMidPoint)
                        End If
                End Select
            Case HSRefRepeat
                Select Case CC.ControlValue
                    Case tpOn
                        SonosPlayer.SonosRepeat = repeat_modes.repeat_all
                    Case tpOff
                        SonosPlayer.SonosRepeat = repeat_modes.repeat_off
                    Case tpToggle
                        SonosPlayer.ToggleRepeat()
                End Select
            Case HSRefShuffle
                Select Case CC.ControlValue
                    Case tpOn
                        SonosPlayer.SonosShuffle = Shuffle_modes.Shuffled
                    Case tpOff
                        SonosPlayer.SonosShuffle = Shuffle_modes.Ordered
                    Case tpToggle
                        SonosPlayer.ToggleShuffle()
                End Select
            Case HSRefTrackPos
                Select Case CC.ControlValue
                    Case vpDown
                        SonosPlayer.SeekTime(ConvertSecondsToTimeFormat(MyPlayerPosition - 10))
                    Case vpUp
                        SonosPlayer.SeekTime(ConvertSecondsToTimeFormat(MyPlayerPosition + 10))
                    Case Else
                        Try
                            If MyHSTrackPositionFormat = HSSTrackPositionSettings.TPSPercentage Then
                                SonosPlayer.SeekTime(ConvertSecondsToTimeFormat(Val(CC.ControlValue / 100 * MyTrackLength)))
                            Else
                                SonosPlayer.SeekTime(ConvertSecondsToTimeFormat(Val(CC.ControlValue)))
                            End If
                        Catch ex As Exception
                        End Try
                End Select
            Case HSRefVolume
                Select Case CC.ControlValue
                    Case vpDown
                        SonosPlayer.ChangeVolumeLevel("Master", -tempVolumeStep)
                    Case vpUp
                        SonosPlayer.ChangeVolumeLevel("Master", tempVolumeStep)
                    Case Else
                        ' this should be the slider
                        SonosPlayer.SetVolumeLevel("Master", CC.ControlValue)
                End Select
        End Select
    End Sub



    Private Sub DockedDeviceCallbackDeviceFound(ByVal pDevice As MyUPnPDevice)
        Log("Callback received for Docked ZonePlayer " & ZoneName & " with device name = " & pDevice.UniqueDeviceName & " and Model = " & pDevice.ModelNumber, LogType.LOG_TYPE_INFO)
        If Mid(pDevice.UniqueDeviceName, 1, 16) <> "uuid:DOCKRINCON_" Then
            If g_bDebug Then Log("Device Finder Call Back for Wireless Dock zoneplayer = " & ZoneName & " found non Sonos device with UDN =  " & pDevice.UniqueDeviceName & " Friendly Name = " & pDevice.FriendlyName, LogType.LOG_TYPE_WARNING)
            Exit Sub ' this is the UPNP service of HS itself on an XP machine responding
        End If
        Try
            ContentDirectory = pDevice.Services.Item("urn:upnp-org:serviceId:ContentDirectory")
            If g_bDebug Then Log("ContentDirectory for zoneplayer = " & ZoneName & ". Last Transport Status = " & ContentDirectory.LastTransportStatus, LogType.LOG_TYPE_INFO)
            DockedDeviceStatus = "Online"
            If Not ContentDirectory Is Nothing Then
                Try
                    ContentDirectory.AddCallback(myContentDirectoryCallback)
                    Log("ContentDirectoryControlCallback added for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
                    If MyDockediPodPlayerName <> "" And AutoBuildDockedDB Then BuildDockedPlayerDatabase(CurrentAppPath & DockedPlayersDBPath & MyDockediPodPlayerName & ".sdb", MyDockediPodPlayerName)
                Catch ex As Exception
                    Log("Error in Adding ContentDirectoryControl Call Back for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
            ConnectToIPod = False
        Catch ex As Exception
            Log("Error in Services.ContentDirectory for Wireless Dock zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DockedDeviceCallbackDeviceLost(ByVal FindData As Integer, ByVal bstrUDN As String)
        Try
            ' Device is Removed or Lost
            Log("Ipod for ZonePlayer " & ZoneName & " has been disconnected from the network in myDockedDeviceFinderCallback_DeviceLost.", LogType.LOG_TYPE_INFO)
            DockedDeviceStatus = "Offline"
            'DestroyObjects()
            ConnectWireLessDock("uuid:DOCK" & UDN & "_MS")
        Catch ex As Exception
            Log("Well we were here in the Device Lost Callback proc but messed up along the way for Wireless Dock zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DockedDeviceCallbackSearchOperationCompleted(ByVal FindData As Integer)
        Try
            ' Search is Complete
            If DockedDeviceStatus = "Online" Then
                Log("Search for Docked iPod in " & ZoneName & " complete.", LogType.LOG_TYPE_INFO)
            Else
                Log("DockedDeviceCallbackSearchOperationCompleted: iPod for ZonePlayer " & ZoneName & " could not be located. Reattempting", LogType.LOG_TYPE_WARNING)
                'Disconnect()
                ConnectWireLessDock("uuid:DOCK" & UDN & "_MS")
            End If
        Catch ex As Exception
            Log("Well we were here in the Search Complete Callback proc but messed up along the way for zoneplayer = " & ZoneName & "Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub TransportStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myAVTransportCallback.ControlStateChange
        If SuperDebug Then
            Log("TransportChangeCallback for player - " & ZoneName & " VarName = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        ElseIf g_bDebug Then
            Log("TransportChangeCallback for player - " & ZoneName & " VarName = " & StateVarName, LogType.LOG_TYPE_INFO)
        End If

        Dim TransportState As String()
        TransportState = {"", "0", "0:00:00", "", "", "", "", "", "", "", "", NoArtPath, NoArtPath, "", ""}
        Dim UpdateSongChange As Boolean = False
        Dim UpdatePlayChange As Boolean = False
        Try
            Dim MetaData As String = ""        'As XML Document
            Dim NextMetaData As String = ""    'As XML Document
            Dim InternetRadio As Boolean = False
            Dim TrackInfo(16)
            TrackInfo = {"", "", "", "", NoArtPath, "", "", "0", "", "", "False", "", "", "", "", "", ""}
            Dim infos As String()
            Dim xmlData As XmlDocument = New XmlDocument
            Dim UpdateDeviceValue As Boolean = False

            If Trim(Value.ToString) <> "" Then
                ' when sonos links/unlinks zone, it does call this procedure with StateVarName = LastChange (which it always does) but no XML
                ' call GetCurrentTrackInfo to find out
                Try
                    xmlData.LoadXml(Value.ToString)
                    'Log( "XML=" & Value.ToString) ' used for testing
                Catch ex As Exception
                    Log("Error in TransportStateChange loading XML. Error = " & ex.Message & ". XML = " & Value.ToString, LogType.LOG_TYPE_ERROR)
                    Log("Error in TransportStateChange loading XML. Parameters:  VarName = " & StateVarName, LogType.LOG_TYPE_ERROR)
                    Exit Sub
                End Try
            End If

            ' TransportState (0)  = PlayerState 
            ' TransportState (1)  = Current Track
            ' TransportState (2)  = Current Track Nbr
            ' TransportState (3)  = Nbr of Tracks
            ' TransportState (4)  = Artist/Creator
            ' TransportState (5)  = Tittle track
            ' TransportState (6)  = Album
            ' TransportState (7)  = Next Artist
            ' TransportState (8)  = Next Track
            ' TransportState (9)  = Next Album
            ' TransportState (10) = Current URI
            ' TransportState (11) = Current Album URI
            ' TransportState (12) = Next Album URI
            ' TransportState (13) = Current Play Mode (repeat/shuffle)
            ' TransportState (14) = Current Section

            Try
                TransportState(0) = xmlData.GetElementsByTagName("TransportState").Item(0).Attributes("val").Value
            Catch ex As Exception
                ' I had an example like this
                ' <Event xmlns="urn:schemas-upnp-org:metadata-1-0/AVT/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/">
                ' <InstanceID val="0">
                '       <TransportStatus val="ERROR_UNSUPPORTED_FORMAT"/>
                '       <TransportErrorDescription val="8,7,Vamo Pa el Banco,Napster,x-sonos-mms:track%3a23784440?sid=0&flags=32,"/>
                '       <TransportErrorURI val="x-sonos-mms:track%3a23784440?sid=0&flags=32"/>
                ' </InstanceID>
                ' </Event>
                ' other example
                ' <Event xmlns="urn:schemas-upnp-org:metadata-1-0/AVT/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/">
                '<InstanceID val="0">
                '       <TransportStatus val="ERROR_LOST_CONNECTION"/>
                '       <TransportErrorDescription val="7,7,So Appalled [Album Version (Explicit)],Napster,x-sonos-mms:track%3a36054680?sid=0&flags=32,"/>
                '       <TransportErrorURI val="x-sonos-mms:track%3a36054680?sid=0&flags=32"/>
                '</InstanceID>
                '</Event>
                ' another case: when I grab a groupcoordator to be a slave player in another grouping a receive a transportchange event at one of the slave players that was originally linked to the 
                ' old groupcoordinator and there is no transportstate info in it. Just AVTRANSPORTAVI (indicating linking to a new master) and AVTRANSPORTURIMETADATA
                Try
                    TransportState(0) = xmlData.GetElementsByTagName("TransportStates").Item(0).Attributes("val").Value
                    'If g_bDebug Then Log("Warning in TransportStateChange, Transportstatus = " & TransportState(0) & ". XML = " & Value.ToString & " Parameters:  VarName = " & StateVarName, LogType.LOG_TYPE_WARNING)
                Catch ex1 As Exception
                    'If g_bDebug Then Log("Warning in TransportStateChange finding transportstate in XML. Error = " & ex.Message & ". XML = " & Value.ToString & " Parameters:  VarName = " & StateVarName, LogType.LOG_TYPE_WARNING)
                End Try
            End Try

            MyCurrentTransportState = TransportState(0)
            '
            If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(TransportState) = " & TransportState(0).ToString & " and previous state was = " & CurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)

            'If Not (TransportState(0) = "STOPPED" And CurrentPlayerState = player_state_values.stopped) Then ' this is when we load for example the queue
            Try
                TrackInfo = GetCurrentTrackInfo()   ' this won't work for Wireless Dock players
            Catch ex As Exception
                Log("Error in TransportStateChange getting track info. Error = " & ex.Message & ". XML = " & Value.ToString, LogType.LOG_TYPE_ERROR)
                xmlData = Nothing
                Exit Sub
            End Try
            'End If

            If TrackInfo(10) = "True" Then InternetRadio = True

            If ZoneSource = "Docked" Then
                Exit Sub
            End If
            Try
                TransportState(1) = xmlData.GetElementsByTagName("CurrentTrack").Item(0).Attributes("val").Value
                If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(CurrentTrack) = " & TransportState(1).ToString, LogType.LOG_TYPE_INFO)
                SetTrackNbr(TransportState(1))
            Catch ex As Exception
            End Try
            Try
                TransportState(2) = xmlData.GetElementsByTagName("CurrentTrackDuration").Item(0).Attributes("val").Value
                If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(CurrentTrackDuration) = " & TransportState(2).ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                TransportState(3) = xmlData.GetElementsByTagName("NumberOfTracks").Item(0).Attributes("val").Value
                If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(NumberOfTracks) = " & TransportState(3).ToString, LogType.LOG_TYPE_INFO)
                SetNbrOfTracks(TransportState(3))
            Catch ex As Exception
            End Try

            Try
                TransportState(13) = xmlData.GetElementsByTagName("CurrentPlayMode").Item(0).Attributes("val").Value
                If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(CurrentPlayMode) = " & TransportState(13).ToString, LogType.LOG_TYPE_INFO)
                SetShuffleState(TransportState(13))
            Catch ex As Exception
            End Try
            Try
                TransportState(14) = xmlData.GetElementsByTagName("CurrentSection").Item(0).Attributes("val").Value
                If SuperDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(CurrentSection) = " & TransportState(14).ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Dim TempInfo
            Try
                TempInfo = xmlData.GetElementsByTagName("LastChange").Item(0).Attributes("val").Value
                If SuperDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(LastChange) = " & TempInfo, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                TempInfo = xmlData.GetElementsByTagName("TransportStatus").Item(0).Attributes("val").Value
                If SuperDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(TransportStatus) = " & TempInfo, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                TempInfo = ""
                TempInfo = xmlData.GetElementsByTagName("r:EnqueuedTransportURIMetaData").Item(0).Attributes("val").Value
                If TempInfo <> "" Then MyEnqueuedTransportURIMetaData = TempInfo
                If SuperDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(r:EnqueuedTransportURIMetaData) = " & MyEnqueuedTransportURIMetaData, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                TempInfo = ""
                TempInfo = xmlData.GetElementsByTagName("r:EnqueuedTransportURI").Item(0).Attributes("val").Value
                If TempInfo <> "" Then MyEnqueuedTransportURI = TempInfo
                If SuperDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(r:EnqueuedTransportURI) = " & MyEnqueuedTransportURI, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try

            If MyZoneModel = "WD100" Then
                ' retrieve more info because we cannot call gettrackinfo
                Try
                    TrackInfo(3) = xmlData.GetElementsByTagName("CurrentTrackURI").Item(0).Attributes("val").Value
                    If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(CurrentTrackURI) = " & TrackInfo(3).ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
            End If

            Try
                MetaData = xmlData.GetElementsByTagName("CurrentTrackMetaData").Item(0).Attributes("val").Value
                'Log( "MetaData = " & MetaData.ToString)
                Try
                    xmlData.LoadXml(MetaData)
                Catch ex As Exception
                    If Trim(MetaData) <> "" Then
                        Log("Error in TransportStateChange loading XML MetaData. XML = " & MetaData.ToString, LogType.LOG_TYPE_ERROR)
                        xmlData = Nothing
                        Exit Sub
                    Else
                        ' when sonos links/unlinks zone, it does call this procedure with StateVarName = LastChange (which it always does) but no XML
                        ' call GetCurrentTrackInfo to find out
                        'TrackInfo = GetCurrentTrackInfo()
                    End If
                End Try
            Catch ex As Exception
            End Try

            ' Changed in V101 due to a bug in Sonos when Play to is issued from WMP, the GetMediaInfo returns corrupt metadata

            Try
                If TrackInfo(0) = "" Then
                    TrackInfo(0) = xmlData.GetElementsByTagName("dc:creator").Item(0).InnerText
                    MyCurrentTrackInfo(0) = TrackInfo(0)
                End If
            Catch ex As Exception
            End Try
            Try
                If TrackInfo(1) = "" Then
                    TrackInfo(1) = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
                    MyCurrentTrackInfo(1) = TrackInfo(1)

                End If
            Catch ex As Exception
            End Try
            Try
                If TrackInfo(2) = "" Then
                    TrackInfo(2) = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                    MyCurrentTrackInfo(2) = TrackInfo(2)
                End If
            Catch ex As Exception
            End Try
            Try
                If TrackInfo(4) = "" Or TrackInfo(4) = NoArtPath Then
                    TrackInfo(4) = xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText
                    If TrackInfo(4) = "" Then
                        TrackInfo(4) = NoArtPath '"http://" & hs.GetIPAddress & "/Sonos/images/noart.jpg"  '"file:///" & EncodeURI(hs.getAppPath) & "\html\Sonos\images\noart.jpg"
                    Else
                        'TrackInfo(4) = DecodeURI(TrackInfo(4)) remove in version 3.1.0.19 because failing to retrieve the Art
                        TrackInfo(4) = BuildArtURL(TrackInfo(4))
                        TrackInfo(4) = GetAlbumArtPath(TrackInfo(4), False)
                    End If
                    MyCurrentTrackInfo(4) = TrackInfo(4)
                End If
            Catch ex As Exception
            End Try
            Try
                If TrackInfo(9) = "" Then
                    TrackInfo(9) = "Tracks"
                    ZoneSource = TrackInfo(9)
                    MyZoneSourceExt = ZoneSource
                    MyCurrentTrackInfo(9) = TrackInfo(9)
                End If
            Catch ex As Exception
            End Try
            Try
                If TrackInfo(6) = "" And TransportState(2) <> "" Then
                    TrackInfo(6) = TransportState(2)
                    SetTrackLength(GetSeconds(TrackInfo(6)))
                    MyCurrentTrackInfo(6) = TrackInfo(6)
                End If
                If TrackInfo(7) = "" Then
                    TrackInfo(7) = TransportState(1)
                    MyCurrentTrackInfo(7) = TrackInfo(7)
                End If
                SetTrackNbr(TrackInfo(7))
            Catch ex As Exception
            End Try

            ' added V114, the stream content is not always available in getpositioninfo but is available in the transportchange event info
            If TrackInfo(15) = "" Then
                Try
                    TrackInfo(15) = xmlData.GetElementsByTagName("r:streamContent").Item(0).InnerText
                    If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState (Streamcontent/Title) = " & TrackInfo(15).ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                    TrackInfo(15) = ""
                End Try
            End If
            If TrackInfo(16) = "" Then
                Try
                    TrackInfo(16) = xmlData.GetElementsByTagName("r:radioShowMd").Item(0).InnerText
                    If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState (RadioShow/Title) = " & TrackInfo(16).ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                    TrackInfo(16) = ""
                End Try
            End If

            If ZoneSource = "TV" Then
                ' reset trackinfo
                Track = "TV Input"
                Album = ""
                Artist = ""
                ArtworkURL = NoArtPath
                TransportState(5) = Track
                TransportState(4) = Artist
                TransportState(6) = Album
                ' reset nexttrackinfo
                NextTrack = ""
                NextAlbum = ""
                NextArtist = ""
                NextArtworkURL = NoArtPath
                GoTo updateHSDevices ' changed in v3.1.0.10
            End If

            ' reload the original XML
            Try
                xmlData.LoadXml(Value.ToString)
                'Log( "XML=" & Value.ToString) ' used for testing
            Catch ex As Exception
                If Trim(Value.ToString) = "" Then
                    ' when sonos links/unlinks zone, it does call this procedure with StateVarName = LastChange (which it always does) but no XML
                    ' call GetCurrentTrackInfo to find out
                Else
                    Log("Error in TransportStateChange loading XML. Error = " & ex.Message & ". XML = " & Value.ToString, LogType.LOG_TYPE_ERROR)
                    Log("Error in TransportStateChange loading XML. Parameters:  VarName = " & StateVarName, LogType.LOG_TYPE_ERROR)
                End If
                xmlData = Nothing
                If g_bDebug Then Log("TransportStateChange has no XML and is done for ZP - " & ZoneName, LogType.LOG_TYPE_INFO)
                'Exit Sub ' removed for v80 because we fail to generate play events when the zone is linked
            End Try
            Try
                NextMetaData = xmlData.GetElementsByTagName("r:NextTrackMetaData").Item(0).Attributes("val").Value 'r:NextTrackMetaData
                MyNextURIMetaData = NextMetaData
            Catch ex As Exception
                NextMetaData = ""
            End Try
            'If g_bDebug Then Log( "TransportChangeCallback for player - " & ZoneName & " NextMetaData = " & NextMetaData.ToString)
            'TrackInfo(0) = Artist
            'TrackInfo(1) = Album
            'TrackInfo(2) = Title
            'TrackInfo(3) = CurrentURI
            'TrackInfo(4) = AlbumArtURI
            'TrackInfo(5) = Track Position
            'TrackInfo(6) = Track Duration
            'TrackInfo(7) = Queue Position
            'TrackInfo(8) = % played
            'TrackInfo(9) = Source
            'TrackInfo(10)= Internet Radio True/False
            'TrackInfo(11)= TrackInfo(0) and equal to radiostation name or Service
            'TrackInfo(12)= CurrentURIMetaData
            'TrackInfo(13)= RadioCurrentURI
            'TrackInfo(14)= RadioCurrentMetaData
            'TrackInfo(15)= StreamContent added v107
            'TrackInfo(16)= RadioShow added v114
            TransportState(4) = TrackInfo(0) ' Artist/creator
            If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(Artist) = " & TransportState(4).ToString, LogType.LOG_TYPE_INFO)
            TransportState(5) = TrackInfo(2) ' Title
            If InternetRadio Then ' Internet radio
                If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(Radio) = " & MyZoneSourceExt & " - Title info = " & TransportState(5).ToString, LogType.LOG_TYPE_INFO)
            Else
                If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(title) = " & TransportState(5).ToString, LogType.LOG_TYPE_INFO)
            End If
            TransportState(6) = TrackInfo(1) ' Album
            If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(Album) = " & TransportState(6).ToString, LogType.LOG_TYPE_INFO)
            TransportState(10) = TrackInfo(3) ' CurrentURI
            If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(CurrentURI) = " & TransportState(10).ToString, LogType.LOG_TYPE_INFO)
            TransportState(11) = TrackInfo(4) ' AlbumArtURI
            If TransportState(11) = "" Then
                TransportState(11) = NoArtPath
            End If

            If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(AlbumArtURL) = " & TransportState(11).ToString, LogType.LOG_TYPE_INFO)

            Try
                NextMetaData = Trim(NextMetaData)
                If NextMetaData <> "" And MyZoneSourceExt <> "Stream Radio" Then
                    Try
                        xmlData.LoadXml(NextMetaData)
                    Catch ex As Exception
                        Log("Error in TransportStateChange loading XML NextMetaData. XML = " & NextMetaData, LogType.LOG_TYPE_ERROR)
                        xmlData = Nothing
                        Exit Sub
                    End Try
                    ' now we are looking at what comes next
                    Try
                        TransportState(7) = xmlData.GetElementsByTagName("dc:creator").Item(0).InnerText
                        If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(Next Artist) = " & TransportState(7).ToString, LogType.LOG_TYPE_INFO)
                        NextArtist = TransportState(7)
                    Catch ex As Exception
                    End Try
                    Try
                        TransportState(8) = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                        If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(Next Title) = " & TransportState(8).ToString, LogType.LOG_TYPE_INFO)
                        NextTrack = TransportState(8)
                    Catch ex As Exception
                    End Try
                    Try
                        TransportState(9) = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
                        If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(Next Album) = " & TransportState(9).ToString, LogType.LOG_TYPE_INFO)
                        NextAlbum = TransportState(9)
                    Catch ex As Exception
                    End Try
                    Try
                        TransportState(12) = xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText
                        TransportState(12) = BuildArtURL(TransportState(12))
                        If TransportState(12) = "" Then
                            TransportState(12) = NoArtPath
                        Else
                            TransportState(12) = GetAlbumArtPath(TransportState(12), True)
                        End If
                        If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " TransportState(Next albumArtURI) = " & TransportState(12).ToString, LogType.LOG_TYPE_INFO)
                    Catch ex As Exception
                        TransportState(12) = NoArtPath
                    End Try
                    NextArtworkURL = TransportState(12)
                End If
            Catch ex As Exception
                ' there probably wasn't any
            End Try

            If SuperDebug Then Log("TransportStateChange processed XML succesfully for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
            ' Get more Info on the track. Is this internet radio or we are linked or source from another zone's input
            'TrackInfo = GetCurrentTrackInfo()
            '
            ' filter out ZPSTR_BUFFERING and ZPSTR_CONNECTING

            If UCase(TransportState(5)) = "ZPSTR_BUFFERING" Or UCase(TransportState(5)) = "ZPSTR_CONNECTING" Then
                If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " is filtered out because player is connecting to Radiostation", LogType.LOG_TYPE_INFO)
                xmlData = Nothing
                Exit Sub
            End If

            If InternetRadio Then
                If (MyZoneSourceExt = "Sirius") Or (MyZoneSourceExt = "SiriusXM" And TrackInfo(15).ToString.Contains("BR P|")) Or (TrackInfo(15).ToString.Contains("TYPE=SNG|")) Then
                    ' Sirius Provides Info in the following Form
                    ' BR P|TYPE=SNG|TITLE Billionaire|ARTIST Travie McCoy
                    ' New form I believe from Apple Music TYPE=SNG|TITLE Rapture|ARTIST Blondie|ALBUM Autoamerican 
                    Dim Tempstring As String
                    Tempstring = TrackInfo(15).ToString ' changed in v109 from TrackInfo (2) to trackinfo (15) = r:streamContent
                    infos = Tempstring.Split("|")
                    Dim Title As String = ""
                    Dim Artist As String = ""
                    Dim Album As String = ""
                    If infos IsNot Nothing And (UBound(infos) > 0) Then
                        For Each Field As String In infos
                            If Field.IndexOf("TITLE") = 0 Then
                                Mid(Field, 1, 5) = "     " ' Remove the Title part
                                Title = Trim(Field)
                                TransportState(5) = infos(2)
                            ElseIf Field.IndexOf("ARTIST") = 0 Then
                                Mid(Field, 1, 6) = "      " ' Remove the ARTIST part
                                Artist = Trim(Field)
                            ElseIf Field.IndexOf("ALBUM") = 0 Then
                                Mid(Field, 1, 5) = "     " ' Remove the ALBUM part
                                Album = Trim(Field)
                            End If
                        Next
                    End If
                    TransportState(5) = Title
                    TransportState(4) = Artist
                    TransportState(6) = Album
                    'If UBound(infos) >= 2 Then
                    'Mid(infos(2), 1, 5) = "     " ' Remove the Title part
                    'infos(2) = Trim(infos(2))
                    'TransportState(5) = infos(2)
                    'Else
                    'TransportState(5) = "" ' added in V110
                    'End If
                    'If UBound(infos) >= 3 Then
                    'Mid(infos(3), 1, 6) = "      " ' Remove the ARTIST part
                    'infos(3) = Trim(infos(3))
                    'TransportState(4) = Trim(infos(3))
                    'Else
                    'TransportState(4) = "" ' added in V110
                    'End If
                    If TrackInfo(11) <> "" Then
                        ' we have a radio station name, don't lose it
                        MyZoneSourceExt = MyZoneSourceExt & " - " & TrackInfo(11)
                    End If
                ElseIf (MyZoneSourceExt = "Pandora") Or (MyZoneSourceExt = "Rhapsody") Or (MyZoneSourceExt = "Lastfm") Then ' some radio stations do provide the info like LastFM, Sirius
                    ' don't do anything
                    If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " is internet radio " & MyZoneSourceExt, LogType.LOG_TYPE_INFO)
                    If TrackInfo(11) <> "" Then
                        ' we have a radio station name, don't lose it
                        MyZoneSourceExt = MyZoneSourceExt & " - " & TrackInfo(11)
                    End If
                ElseIf MyZoneSourceExt <> "Tracks" Then ' some radio stations do provide the info like LastFM, Sirius
                    If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " is not tracks and radio " & MyZoneSourceExt, LogType.LOG_TYPE_INFO)
                    If TrackInfo(11) <> "" Then
                        ' we have a radio station name, don't lose it
                        MyZoneSourceExt = MyZoneSourceExt & " - " & TrackInfo(11)
                    End If
                    If TrackInfo(0) = "" And TrackInfo(15) <> "" And Not TrackInfo(15).ToString.Contains("ZPSTR_") Then ' added v107/V114+ because slacker radio track and artist are swapped
                        Dim Tempstring As String
                        Tempstring = TrackInfo(15).ToString
                        infos = Tempstring.Split("-")
                        TransportState(4) = infos(0)
                        If UBound(infos) > 1 Then
                            TransportState(5) = "" ' there are multiple "-" so I'm not attempting to split
                            TransportState(4) = TrackInfo(2)
                        ElseIf UBound(infos) = 1 Then
                            TransportState(5) = infos(1)
                        Else
                            ' this means no Info. Stick Radio Name in Track
                            TransportState(5) = TrackInfo(0)
                        End If
                    End If
                End If
                'SetPlayerPosition(0) ' removed because slacker radio events next track which causes the track position to go to zero' do this because streaming does not provide info and HST doesn't update screens
            End If

            If MyZoneSource = "Linked" Then 'MyZoneSourceExt = "Linked" Then Changed V12
                ' the updates should have come in from calls from the Master PI
                TransportState(5) = Track
                TransportState(4) = Artist
                TransportState(6) = Album
                TransportState(11) = ArtworkURL
                TrackInfo(11) = RadiostationName
                Exit Sub ' DONE
            Else
                Track = TransportState(5)
                Artist = TransportState(4)
                Album = TransportState(6)
                ArtworkURL = TransportState(11)
                RadiostationName = TrackInfo(11) ' radiostation channel name
            End If

            '            SetHSPlayerInfo() ' this was moved further in this procedure in v.17
updateHSDevices:

            If TransportState(0) = "PLAYING" Then
                TransportState(0) = "Play"
                MyPlayerWentThroughPlayState = True
                If CurrentPlayerState = player_state_values.Playing Then
                    If (MyPreviousArtist = TransportState(4)) And (MyPreviousTrack = TransportState(5)) And (MyPreviousAlbum = TransportState(6)) Then
                        ' this is a radiostation sending the same info, no need to process what we already have and generate the wrong events
                        If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " is the same from the previous event", LogType.LOG_TYPE_INFO)
                        xmlData = Nothing
                    Else
                        MyPreviousArtist = Artist
                        MyPreviousTrack = Track
                        MyPreviousAlbum = Album
                        'If g_bDebug Then Log( "TransportChangeCallback for player - " & ZoneName & " is different from the previous event - " & TempValue)
                        PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState) ' notify HS if they have the callback linked
                        UpdateSongChange = True
                        If ZoneModel = "WD100" Then
                            ' we need to manually reset the PlayerPosition because we cannot call get track info
                            SetPlayerPosition(0)
                            If Not MyWirelessDockDestinationPlayer Is Nothing Then
                                MyWirelessDockDestinationPlayer.SetPlayerPosition(0)
                            End If
                        End If
                    End If

                Else
                    CurrentPlayerState = player_state_values.Playing
                    PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                    If (MyPreviousArtist = TransportState(4)) And (MyPreviousTrack = TransportState(5)) And (MyPreviousAlbum = TransportState(6)) Then
                        ' this is a radiostation sending the same info, no need to process what we already have and generate the wrong events
                        If g_bDebug Then Log("TransportChangeCallback for player - " & ZoneName & " is the same from the previous event", LogType.LOG_TYPE_INFO)
                        'Exit Sub ' changed this in v.79 when bedroom amp remained on because an on-event wasn't sent
                    Else
                        MyPreviousArtist = Artist
                        MyPreviousTrack = Track
                        MyPreviousAlbum = Album
                        PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState) ' notify HS if they have the callback linked
                        UpdateSongChange = True
                    End If
                    'MyPreviousArtist = ""
                    UpdatePlayChange = True
                    UpdateDeviceValue = True
                End If
            ElseIf TransportState(0) = "STOPPED" Then
                TransportState(0) = "Stop"
                If ZoneModel = "WD100" Then
                    ' we need to manually reset the PlayerPosition because we cannot call get track info
                    SetPlayerPosition(0)
                    If Not MyWirelessDockDestinationPlayer Is Nothing Then
                        MyWirelessDockDestinationPlayer.SetPlayerPosition(0)
                    End If
                End If
                If CurrentPlayerState = player_state_values.Stopped Then
                    'PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState)
                    'PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                    'UpdateSongChange = True
                Else
                    CurrentPlayerState = player_state_values.Stopped
                    If Not MyZoneIsStored Then
                        PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                        UpdatePlayChange = True
                    End If
                    NextTrack = ""
                    NextArtist = ""
                    NextAlbum = ""
                    NextArtworkURL = NoArtPath
                    RadiostationName = ""
                    UpdateDeviceValue = True
                End If
            ElseIf TransportState(0) = "PAUSED_PLAYBACK" Then
                TransportState(0) = "Pause"
                If CurrentPlayerState = player_state_values.Paused Then
                    'PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState)
                    'PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                    'UpdateSongChange = True
                Else
                    CurrentPlayerState = player_state_values.Paused
                    PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                    UpdatePlayChange = True
                    UpdateDeviceValue = True
                End If
            ElseIf TransportState(0).ToUpper = "TRANSITIONING" Then
                ' don't do anything
                UpdateSongChange = True
            Else
                TransportState(0) = "Unknown"
                PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                UpdatePlayChange = True
                'PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState)
                'UpdateSongChange = True
                UpdateDeviceValue = True
                If ZoneModel = "WD100" Then
                    ' we need to manually reset the PlayerPosition because we cannot call get track info
                    SetPlayerPosition(0)
                    If Not MyWirelessDockDestinationPlayer Is Nothing Then
                        MyWirelessDockDestinationPlayer.SetPlayerPosition(0)
                    End If
                End If
            End If
            SetHSPlayerInfo() ' this was moved here in v.17 because the player device would reflect the wrong state

            If g_bDebug Then Log("Checking linked zones for SourceZone = " & ZoneName & ". MyZoneIsSourceForLinkedZone=" & MyZoneIsSourceForLinkedZone.ToString & " and TargetZones = " & MyTargetZoneLinkedList.ToString, LogType.LOG_TYPE_INFO)
            If MyZoneIsSourceForLinkedZone Then
                ' Update all linked Zones
                Dim TargetZones() = {""}
                Dim TargetZone As String
                Dim TempMusicApi As HSPI
                Dim AlreadyExist As Boolean = False
                Dim TempHSDeviceCodeTransport As Integer = -1
                'Dim TempTransportInfo As String = ""
                Try
                    TargetZones = Split(MyTargetZoneLinkedList, ";")
                Catch ex As Exception
                End Try
                For Each TargetZone In TargetZones
                    Try
                        TempMusicApi = MyHSPIControllerRef.GetAPIByUDN(TargetZone.ToString)
                        TempMusicApi.Track = TransportState(5)
                        TempMusicApi.Artist = TransportState(4)
                        TempMusicApi.Album = TransportState(6)
                        TempMusicApi.ArtworkURL = TransportState(11)
                        TempMusicApi.NextArtworkURL = TransportState(12)
                        TempMusicApi.NextTrack = NextTrack
                        TempMusicApi.NextAlbum = NextAlbum
                        TempMusicApi.NextArtist = NextArtist
                        TempMusicApi.RadiostationName = RadiostationName
                        TempMusicApi.MyZoneSourceExt = MyZoneSourceExt
                        TempMusicApi.MusicService = "Linked to - " & ZoneName
                        TempMusicApi.CurrentPlayerState = CurrentPlayerState
                        TempMusicApi.SetTrackLength(MyTrackLength)
                        TempMusicApi.SetPlayerPosition(MyPlayerPosition)
                        TempMusicApi.SetNbrOfTracks(MyNbrOfTracksInQueue)
                        TempMusicApi.SetTrackNbr(MyTrackInQueueNbr)
                        TempMusicApi.SetShuffleState(MyShuffleState)
                        TempMusicApi.MyQueueHasChanged = MyQueueHasChanged
                        Dim TempZoneName As String = TempMusicApi.ZonePlayerName
                        If g_bDebug Then Log("Updating other linked zones. SourceZone=" & ZoneName & ". TargetZone=" & TargetZone.ToString, LogType.LOG_TYPE_INFO)
                        ' notify HS if they have the callback linked
                        If UpdateSongChange Then TempMusicApi.PlayChangeNotifyCallback(player_status_change.SongChanged, CurrentPlayerState)
                        If UpdatePlayChange Then TempMusicApi.PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, CurrentPlayerState)
                        TempHSDeviceCodeTransport = MyHSPIControllerRef.GetHSDeviceReference(TargetZone.ToString)
                        TempMusicApi.SetHSPlayerInfo()
                        If UpdateDeviceValue Then
                            hs.SetDeviceValue(TempHSDeviceCodeTransport, CurrentPlayerState)
                            If g_bDebug Then Log("HS DeviceValue updated in TransportChanged for linked zonePlayer " & TempZoneName & ". HS Code = " & TempHSDeviceCodeTransport & " and value = " & CurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
                        End If
                        If SuperDebug Then Log("HS updated in TransportChanged for linked zonePlayer " & TempZoneName & ". HS Code = " & TempHSDeviceCodeTransport & " and updated Device Flag = " & UpdateDeviceValue.ToString, LogType.LOG_TYPE_INFO)
                    Catch ex As Exception
                        Log("Error updating other linked zones. SourceZone=" & ZoneName & ". TargetZone=" & TargetZone.ToString & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Next
                TargetZones = Nothing
                TargetZone = Nothing
                TempMusicApi = Nothing
            End If
            Try
                'hs.SetDeviceString(GetHSDeviceRefPlayer, TransportInfo, False)
                If UpdateDeviceValue Then
                    hs.SetDeviceValue(HSRefPlayer, CurrentPlayerState)
                    If g_bDebug Then Log("HS DeviceValue updated in TransportChanged for zonePlayer " & ZoneName & ". HSRef = " & HSRefPlayer.ToString & " and value = " & CurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
                End If
                If g_bDebug Then Log("HS updated in TransportChanged for zonePlayer " & ZoneName & ". HSRef = " & HSRefPlayer.ToString & " and updated DeviceValue = " & UpdateDeviceValue.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            xmlData = Nothing
        Catch ex As Exception
            Log("ERROR: Not Successful in getting Transport info for zoneplayer = " & ZoneName & ". Error=" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub TransportDied() Handles myAVTransportCallback.ControlDied
        Log("TransportDied. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in TransportDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("Transport Died. ZonePlayer - " & ZoneName & "Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub DevicePropertiesStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myDevicePropertiesCallback.ControlStateChange
        If SuperDebug Then Log("DevicePropertiesStateChange for ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        'If g_bDebug And Not SuperDebug Then Log("DevicePropertiesStateChange for ZonePlayer " & ZoneName & ": Var Name = " & StateVarName, LogType.LOG_TYPE_INFO)
        Try
            If (StateVarName = "ZoneName") Then
                Properties(0) = Value.ToString
                If g_bDebug Then Log("DevicePropertiesStateChange for ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
                ' this could indicate that the zone name has changed.
                ' Two posibilities
                ' a/ Name Changed by user
                ' b/ this is an Sxx player and it was just paired or unpaired
                'If ZoneName <> Value.ToString Then ' this was moved to the ZoneNameChanged procedure
                'If g_bDebug Then Log("DevicePropertiesStateChange: Zone Name Changed from = " & ZoneName & " to " & Value.ToString, LogType.LOG_TYPE_INFO)
                'PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ZoneName)
                'End If
                If Not CheckPlayerIsPairable(MyZoneModel) Then
                    ' the User just changed the Zone Name
                    If (ZoneName <> Value.ToString) And (MyZonePairMasterZoneName <> Value.ToString) And (Value.ToString <> "") Then ZoneNameChanged(Value.ToString)
                Else
                    ' either pairing change or zone Name change
                    Try
                        'Dim ChannelMapSet
                        'ChannelMapSet = DeviceProperties.QueryStateVariable("ChannelMapSet")
                        'Dim HTSatChanMapSet
                        'HTSatChanMapSet = DeviceProperties.QueryStateVariable("HTSatChanMapSet") ' this variable may not be available yet ... use ZoneGroupTopology:GetZoneGroupState
                        If MyChannelMapSet <> "" Then ' removed 7/16/2019 in v3.1.0.35
                            'If g_bDebug Then Log("DevicePropertiesStateChange: Player: " & ZoneName & " was paired and ZoneName changed to " & Value.ToString, LogType.LOG_TYPE_INFO)
                            ' do nothing an event will fire on the ChannelMapSet Change where all updates will happen
                        ElseIf MyHTSatChanMapSet <> "" Then ' removed 7/16/2019 in v3.1.0.35
                            'If g_bDebug Then Log("DevicePropertiesStateChange: Player: " & ZoneName & " was paired to a Playbar and ZoneName changed to " & Value.ToString, LogType.LOG_TYPE_INFO)
                        Else
                            If ZoneName <> Value.ToString Then ZoneNameChanged(Value.ToString)
                        End If
                    Catch ex As Exception
                        'If ZoneName <> Value.ToString Then ZoneNameChanged(Value.ToString) ' removed 7/16/2019 in v3.1.0.35
                    End Try
                End If
            ElseIf (StateVarName = "Icon") Then
                ' "x-rincon-roomicon:garden"
                Properties(1) = Replace(Value.ToString, "x-rincon-roomicon:", "")
                Dim RoomIcon As String = GetStringIniFile(UDN, DeviceInfoIndex.diRoomIcon.ToString, "")
                If Properties(1).ToString <> RoomIcon Then
                    MyHSPIControllerRef.ZoneRoomIconChanged(GetUDN, Properties(1).ToString)
                End If
            ElseIf (StateVarName = "Invisible") Then
                Properties(2) = Value
            ElseIf (StateVarName = "SettingsReplicationState") Then
                ' this is an important value. It shows some sequence nbr for each zone and looking at the sequence #, one can derive whether something was
                ' changed. I think I'm going to save this in the DB and use it to compare against. If changed, then I trigger a new DB rebuild
                ' the format is UDN,<seq number>, UDN, <seq number>
                ' maybe for time being, I store it in the ini file and see what I can do with it
                'UserRadioUpdateID(Value = RINCON_000E5825227A01400, 13)
                'SavedQueuesUpdateID(Value = RINCON_000E5833F3CC01400, 116)
                'ShareListUpdateID(Value = RINCON_000E5824C3B001400, 330)
                'RecentlyPlayedUpdateID(Value = RINCON_000E5824C3B001400, 73)
                'RINCON_000E5825227A01400,9,	<< UserRadioUpdateID
                'RINCON_FFFFFFFFFFFF99999,0,
                'RINCON_000E5833F3CC01400,116,	<< SavedQueuesUpdateID
                'RINCON_000E5824C3B001400,330,	<< ShareListUpdateID
                'RINCON_000E5824C3B001400,44,
                'RINCON_000E5824C3B001400,18,
                'RINCON_000E5824C3B001400,384,	<< ServiceListVersion ?
                'RINCON_000E5824C3B001400,73,	<< RecentlyPlayedUpdateID
                'RINCON_000E5824C3B001400, 3

                '* SystemUpdateID
                '* ContainerUpdateID
                '* ShareListRefreshState [NOTRUN|RUNNING|DONE]
                '* ShareIndexInProgress
                '* ShareIndexLastError
                '* UserRadioUpdateID
                '* MasterRadioUpdateID
                '* SavedQueuesUpdateID
                '* ShareListUpdateID

                Properties(3) = Value.ToString
                If Not MyZoneModel = "WD100" Then
                    ' the info from the WD100 is different
                    ' objIniFile.WriteString("SettingsReplicationState", "UPnPMaster", Value.ToString)
                    Dim LastSettingInfo As String = ""
                    LastSettingInfo = GetStringIniFile("Saved SettingsReplicationState", "UPnPMaster", "")
                    If LastSettingInfo <> Properties(3) Then
                        ' Sonos Settings have Changed
                        CompareReplicationChanges(LastSettingInfo, Properties(3))
                        WriteStringIniFile("Saved SettingsReplicationState", "UPnPMaster", Value.ToString)
                    End If
                Else
                    WriteStringIniFile("DockedSettingsReplicationState", ZoneName, Value.ToString)
                    Dim LastSettingInfo As String = ""
                    LastSettingInfo = GetStringIniFile("Saved DockedSettingsReplicationState", ZoneName, "")
                    If LastSettingInfo <> Properties(3) Then
                        ' Wireless Dock iPod Settings have Changed
                        If g_bDebug Then Log("DevicePropertiesStateChange: Sonos Config has changed for Wireless Dock Player = " & ZoneName & " with iPod = " & MyDockediPodPlayerName.ToString & ". Change Flag Set", LogType.LOG_TYPE_INFO)
                        If g_bDebug Then Log("DevicePropertiesStateChange: Saved Setting = " & LastSettingInfo, LogType.LOG_TYPE_INFO)
                        If g_bDebug Then Log("DevicePropertiesStateChange: New Setting = " & Properties(3), LogType.LOG_TYPE_INFO)
                        WirelessSettingsHaveChanged = True
                    End If
                End If
            ElseIf (StateVarName = "IsZoneBridge") Then
                Properties(4) = Value
            ElseIf (StateVarName = "ChannelMapSet") Then
                ' when an S5 is paired, you get something like this
                ' Office:  Var Name = ChannelMapSet Value = RINCON_000E5858C97A01400:LF,LF;RINCON_000E5859008Axxxx:RF,RF
                ' Office2: Var Name = ChannelMapSet Value = RINCON_000E5858C97A01400:LF,LF;RINCON_000E5859008Axxxx:RF,RF
                '  with the UDN on the left being the "Master"
                ' When they Unpair, the value is empty ""
                'ZonePairingChanged(Value) removed on 7/14/2019 in v3.1.0.34
            ElseIf (StateVarName = "HTSatChanMapSet") Then
                'PlaybarPairingChanged(Value) removed on 7/14/2019 in v3.1.0.34
            ElseIf (StateVarName = "Configuration") Then
            ElseIf (StateVarName = "ZoneGroupState") Then ' added on 7/14/2019 in v3.1.0.34
                ProcessZoneGroupState(Value)
            End If
        Catch ex As Exception
            Log("ERROR: " & ZoneName & " : Something went wrong in Device Property Callback" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub DevicePropertiesDied() Handles myDevicePropertiesCallback.ControlDied
        Log("Device Property Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in DeviceProperties.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in DevicePropertiesDied Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub AudioInStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myAudioInCallback.ControlStateChange
        If g_bDebug Then Log("AudioInCallback for ZonePlayer - " & ZoneName & " Statename = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Try
            If (StateVarName = "AudioInputName") Then
                AudioInState(0) = Value.ToString
                If ZoneModel = "WD100" Then
                    If MyDockediPodPlayerName <> AudioInState(0) Then
                        MyDockediPodPlayerName = AudioInState(0)
                        If HSRefDockDeviceName <> -1 Then
                            hs.SetDeviceString(HSRefDockDeviceName, MyDockediPodPlayerName, True)
                        End If
                    End If
                End If
            ElseIf (StateVarName = "Icon") Then
                AudioInState(1) = Value.ToString
            ElseIf (StateVarName = "LeftLineInLevel") Then
                AudioInState(2) = Value.ToString
            ElseIf (StateVarName = "RightLineInLevel") Then
                AudioInState(3) = Value.ToString
            ElseIf (StateVarName = "LineInConnected") And Value <> "" Then
                ' LineInConnected is used for any zoneplayer and will show true if anything is connected to the line-in interface. For the WD100 it will show iPod docked or not
                ' this is important for WD100's because the event shows when it is inserted or removed!
                'If g_bDebug Then Log( "AudioInCallback for ZonePlayer - " & ZoneName & " LineInConnected for ZoneModel = " & ZoneModel & " and current state = " & AudioInState(4).ToString)
                If MyZoneModel = "WD100" Then
                    iPodDockChange(CBool(Value))
                End If
                AudioInState(4) = CBool(Value)
                MyLineInputConnected = CBool(Value)
            ElseIf (StateVarName = "Playing") And Value <> "" Then ' the Wireless Dock will use this with Value = true/false indicating play or pause
                AudioInState(5) = CBool(Value)
            End If
        Catch ex As Exception
            Log("ERROR: " & ZoneName & " Something went wrong in AudioInStateChange for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub AudioInDied() Handles myAudioInCallback.ControlDied
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in AudioInCallBack.", LogType.LOG_TYPE_WARNING)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in AudioInDied for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub RenderingControlStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myRenderingControlCallback.ControlStateChange
        If SuperDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        Dim NewMasterVolume As Integer
        Dim NewMasterVolumeString As String = ""
        Dim NewMasterMute As Boolean
        Dim NewMasterMuteString As String = ""
        Dim xmlData As XmlDocument = New XmlDocument
        '<Event xmlns="urn:schemas-upnp-org:metadata-1-0/RCS/">
        '   <InstanceID val="0">
        '       <Volume channel="Master" val="23"/>
        '       <Volume channel="LF" val="100"/>
        '       <Volume channel="RF" val="100"/>
        '       <Mute channel="Master" val="0"/>
        '       <Mute channel="LF" val="0"/>
        '       <Mute channel="RF" val="0"/>
        '       <Bass val="10"/><Treble val="6"/>
        '       <Loudness channel="Master" val="1"/>
        '       <OutputFixed val="0"/>
        '       <HeadphoneConnected val="0"/>
        '       <PresetNameList>FactoryDefaults</PresetNameList>
        '       <SpeakerSize val="3"/>
        '       <SubGain val="0"/>
        '       <SubCrossover val="0"/>
        '       <SubPolarity val="0"/>
        '       <SubEnabled val="1"/>
        '   </InstanceID>
        '</Event>
        '<Event xmlns="urn:schemas-upnp-org:metadata-1-0/RCS/"><InstanceID val="0"><Volume channel="Master" val="9"/><Volume channel="LF" val="100"/><Volume channel="RF" val="100"/><Bass val="0"/><Treble val="0"/><Loudness channel="Master" val="1"/><SpeakerSize val="3"/><SubGain val="0"/><SubCrossover val="0"/><SubPolarity val="0"/><SubEnabled val="1"/></InstanceID></Event>
        Try
            Try
                xmlData.LoadXml(Value.ToString)
            Catch ex As Exception
                Log("Error in RenderingControlStateChange for " & ZoneName & " loading XML. XML = " & Value.ToString, LogType.LOG_TYPE_ERROR)
                Exit Sub
            End Try

            Try
                NewMasterVolumeString = xmlData.GetElementsByTagName("Volume").Item(0).Attributes("val").Value   ' this is Master Volume
                If NewMasterVolumeString <> "" Then
                    NewMasterVolume = Val(NewMasterVolumeString)
                    If g_bDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (Master Volume) = " & NewMasterVolume.ToString, LogType.LOG_TYPE_INFO)
                    SetVolume = NewMasterVolume
                    Dim FixedVolume As Integer
                    FixedVolume = GetIntegerIniFile("Fixed Volume", ZoneName, 200)
                    If FixedVolume <> 200 Then
                        ' this means we have a fixed setting for this zone, force it back
                        If FixedVolume <> NewMasterVolume Then
                            SetVolumeLevel("Master", FixedVolume)
                            NewMasterVolume = FixedVolume
                        End If
                    End If
                    MyCurrentMasterVolumeLevel = NewMasterVolume
                End If
            Catch ex As Exception
            End Try
            Try
                Dim NewLeftVolumeString As String = xmlData.GetElementsByTagName("Volume").Item(1).Attributes("val").Value
                If NewLeftVolumeString <> "" Then
                    Dim NewLeftVolume As Integer = Val(NewLeftVolumeString)   ' This is Left Volume (Balance) between 0 and 100 - 100 is balanced in middle
                    If MyLeftVolume <> NewLeftVolume Then
                        MyLeftVolume = NewLeftVolume
                        UpdateBalance()
                    End If
                    If g_bDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (LF Volume) = " & NewLeftVolume.ToString, LogType.LOG_TYPE_INFO)
                End If
            Catch ex As Exception
            End Try
            Try
                Dim NewRightVolumeString As String = xmlData.GetElementsByTagName("Volume").Item(2).Attributes("val").Value
                If NewRightVolumeString <> "" Then
                    Dim NewRightVolume As Integer = Val(NewRightVolumeString)  ' This is Right Volume (Balance) between 0 and 100 - 100 is balanced in middle
                    If MyRightVolume <> NewRightVolume Then
                        MyRightVolume = NewRightVolume
                        UpdateBalance()
                    End If
                    If g_bDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (RF Volume) = " & NewRightVolume.ToString, LogType.LOG_TYPE_INFO)
                End If
            Catch ex As Exception
            End Try
            Try
                NewMasterMuteString = xmlData.GetElementsByTagName("Mute").Item(0).Attributes("val").Value      ' this is Master Mute Volume Level 
                If NewMasterMuteString <> "" Then
                    NewMasterMute = CBool(NewMasterMuteString)
                    If g_bDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (Mute Master Volume) = " & NewMasterMute.ToString, LogType.LOG_TYPE_INFO)
                    If NewMasterMute <> MyCurrentMuteState Then
                        SetMuteState = NewMasterMute
                        'PlayChangeNotifyCallback(player_status_change.SongChanged, player_state_values.UpdateHSServerOnly, False) ' notify HS if they have the callback linked
                    End If
                    If NewMasterMute = False Then MyCurrentMuteState = False Else MyCurrentMuteState = True
                End If
            Catch ex As Exception
            End Try
            Try
                Dim NewLeftMuteVolume As Integer = Val(xmlData.GetElementsByTagName("Mute").Item(1).Attributes("val").Value)     ' This is Left Mute Volume (Balance) between 0 and 100 - 100 is balanced in middle
                If SuperDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (Mute LF Volume) = " & NewLeftMuteVolume.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                Dim NewRightMuteVolume As Integer = Val(xmlData.GetElementsByTagName("Mute").Item(2).Attributes("val").Value)     ' This is Right Mute Volume (Balance) between 0 and 100 - 100 is balanced in middle
                If SuperDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (Mute RF Volume) = " & NewRightMuteVolume.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                Dim NewTrebleValue As Integer = Val(xmlData.GetElementsByTagName("Treble").Item(0).Attributes("val").Value)    ' This is Treble value between -10 and +10
                If SuperDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (Treble) = " & NewTrebleValue.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                Dim NewBassValue As Integer = Val(xmlData.GetElementsByTagName("Bass").Item(0).Attributes("val").Value)     ' This is Bass value between -10 and +10
                If SuperDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (Bass) = " & NewBassValue.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                Dim NewLoudnessStateString As String = xmlData.GetElementsByTagName("Loudness").Item(0).Attributes("val").Value
                If NewLoudnessStateString <> "" Then
                    Dim NewLoudnessState As Boolean = CBool(NewLoudnessStateString)  ' This is Loudness off=0; on=1
                    If NewLoudnessState <> MyCurrentLoudnessState Then
                        SetLoudness = NewLoudnessState
                    End If
                    If NewLoudnessState = False Then MyCurrentLoudnessState = False Else MyCurrentLoudnessState = True
                    If g_bDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (Loudness) = " & NewLoudnessState.ToString, LogType.LOG_TYPE_INFO)
                End If
            Catch ex As Exception
            End Try
            Try
                Dim NewFixedOutputState As Boolean = CBool(xmlData.GetElementsByTagName("OutputFixed").Item(0).Attributes("val").Value)   ' Fixed = 1, Variable = 0 ; when you change this to fixed, the output level Master,lf and rf are all 100
                If SuperDebug Then Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (OutputFixed) = " & NewFixedOutputState.ToString, LogType.LOG_TYPE_INFO)
                MyCurrentFixedVolumeState = NewFixedOutputState
            Catch ex As Exception
            End Try
            Try
                'Rendering(10) = xmlData.GetElementsByTagName("VolumeDB").Item(0).Attributes("val").Value   ' not sure what this is
                'If SuperDebug Then Log(ZoneName & " : Rendering (VoluemDB) = " & Rendering(10).ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Try
                'Rendering(11) = xmlData.GetElementsByTagName("LastChange").Item(0).Attributes("val").Value   ' not sure what this is
                'If SuperDebug Then Log(ZoneName & " : Rendering (LastChange) = " & Rendering(11).ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
            End Try
            Dim MoreInfo As String = ""
            If SuperDebug Then
                Try
                    MoreInfo = xmlData.GetElementsByTagName("SpeakerSize").Item(0).Attributes("val").Value
                    Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (SpeakerSize) = " & MoreInfo.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    MoreInfo = xmlData.GetElementsByTagName("SubGain").Item(0).Attributes("val").Value
                    Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (SubGain) = " & MoreInfo.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    MoreInfo = xmlData.GetElementsByTagName("SubCrossover").Item(0).Attributes("val").Value
                    Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (SubCrossover) = " & MoreInfo.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    MoreInfo = xmlData.GetElementsByTagName("SubPolarity").Item(0).Attributes("val").Value
                    Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (SubPolarity) = " & MoreInfo.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    MoreInfo = xmlData.GetElementsByTagName("SubEnabled").Item(0).Attributes("val").Value
                    Log("Rendering Change callback - ZonePlayer " & ZoneName & " : Rendering (SubEnabled) = " & MoreInfo.ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
            End If
        Catch ex As Exception
            Log("Error in Rendering Change callback - ZonePlayer " & ZoneName & ". This rendering didn't work too well with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub RenderingControlDied() Handles myRenderingControlCallback.ControlDied
        Log("Rendering Control Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in RenderingControlDied. Attempting to reconnect...", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("Error in RenderingControlDied. This rendering didn't work too well for zoneplayer = " & ZoneName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub AlarmClockStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myAlarmClockCallback.ControlStateChange
        If g_bDebug Then Log("AlarmClock callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        If Value = "AlarmListVersion" Then
            ' something about an alarm was changed
            ' the 

        End If
        ' TimeZone String
        ' TimeServer String
        ' TimeGeneration UI4
        ' AlarmListVersion String
        ' DailyIndexRefreshTime String
        ' TimeFormat As string
        ' DateFormat As String

    End Sub

    Private Sub AlarmClockControlDied() Handles myAlarmClockCallback.ControlDied
        Log("AlarmClock Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in AlarmClockDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("Error in AlarmClockControlDied: Something went wrong for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub MusicServicesStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myMusicServicesCallback.ControlStateChange
        If g_bDebug Then Log("MusicServices callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        ' ServiceListVersion

    End Sub

    Private Sub MusicServicesControlDied() Handles myMusicServicesCallback.ControlDied
        Log("MusicServices Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in MusicServicesDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in MusicServicesDied Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub SystemPropertiesStateChange(ByVal StateVarName As String, ByVal Value As String) Handles mySystemPropertiesCallback.ControlStateChange
        If g_bDebug Then Log("SystemProperties callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
    End Sub

    Private Sub SystemPropertiesControlDied() Handles mySystemPropertiesCallback.ControlDied
        Log("SystemProperties Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in SystemPropertiesDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in SystemPropertiesDied Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub GroupManagementStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myGroupManagementCallback.ControlStateChange
        If SuperDebug Then
            Log("GroupManagement callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        ElseIf g_bDebug Then
            'Log("GroupManagement callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName, LogType.LOG_TYPE_INFO)
        End If
        If StateVarName = "BufferingResultCode" Then
            ' example
            ' A_ARG_TYPE_BufferingResultCode
        ElseIf StateVarName = "MemberID" Then
            ' A_ARG_TYPE_MemberID
        ElseIf StateVarName = "TransportSettings" Then
            ' A_ARG_TYPE_TransportSettings

        ElseIf StateVarName = "GroupCoordinatorIsLocal" Then
            ' GroupCoordinatorIsLocal
            ' If this value is false, the player belongs to a zonegroup but is not the coordinator
            MyGroupCoordinatorIsLocal = Value
        ElseIf StateVarName = "LocalGroupUUID" Then
            ' LocalGroupUUID
            MyLocalGroupUUID = Value
        ElseIf StateVarName = "VolumeAVTransportURI" Then
            ' VolumeAVTransportURI
        End If
    End Sub

    Private Sub GroupManagementControlDied() Handles myGroupManagementCallback.ControlDied
        Log("GroupManagement Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in GroupManagementDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in GroupManagementDied Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub ZonegroupTopologyStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myZonegroupTopologyCallback.ControlStateChange
        'If g_bDebug Then Log("ZonegroupTopology callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        'If g_bDebug Then Log("ZonegroupTopology callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName, LogType.LOG_TYPE_INFO)
        If StateVarName = "ZoneGroupState" Then
            'If g_bDebug Then Log("ZonegroupTopology callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
            If MyZoneGroupState <> Value.ToString Then
                MyZoneGroupStateHasChanged = True ' used by Ajax for web player
                ProcessZoneGroupState(Value)
            End If
        ElseIf StateVarName = "AvailableSoftwareUpdate" Then
            If g_bDebug Then Log("ZonegroupTopology callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
            ' <UpdateItem xmlns="urn:schemas-rinconnetworks-com:update-1-0" 
            '   Type="Software" Version="14.4-33290" 
            '   UpdateURL="http://update.sonos.com/firmware/Gold/v3.3-Clash/^14.4-33290" 
            '   DownloadSize="0"/>
        ElseIf StateVarName = "ThirdPartyMediaServers" Then
            '<MediaServers>
            '<Ex CURL="http://192.168.1.109:1400/MediaServer/ContentDirectory/Control" EURL="http://192.168.1.109:1400/MediaServer/ContentDirectory/Event" T="3" EXT=""/>
            '<MediaServer Name="SONJA&apos;S NAN" UDN="DOCKRINCON_000E5860905A01400_MS" Location="http://192.168.1.109:1400/xml/dock_cd.xml"/>
            '   <Service UDN="SA_RINCON7_dcorsus" Md="" Password="tomas3120" NumAccounts="1" Username0="dcorsus" Md0="" Password0="tomas3120"/>
            '   <Service UDN="SA_RINCON11_aplasticfeast" Md="" Password="tomrules" NumAccounts="1" Username0="aplasticfeast" Md0="" Password0="tomrules"/>
            '   <Service UDN="SA_RINCON3_dirk@famcorsus.com" Md="" Password="tomas3120" NumAccounts="1" Username0="dirk@famcorsus.com" Md0="" Password0="tomas3120"/>
            '   <Service UDN="SA_RINCON2_00_0e_58_25_22_7a_9@sonos.com" Md="40134" Password="00M0qM58M25M22M7mM9" TrialDays="0" NumAccounts="1" Username0="00_0e_58_25_22_7a_9@sonos.com" Md0="40134" Password0="00M0qM58M25M22M7mM9"/>
            '   <Service UDN="SA_RINCON1543_Sonos" Md="" Password="" NumAccounts="1" Username0="Sonos" Md0="" Password0=""/>
            '   <Service UDN="SA_RINCON6_" Md="010E5825227A" Password="" TrialDays="0" NumAccounts="1" Username0="" Md0="010E5825227A" Password0=""/>
            '   <Service UDN="SA_RINCON1799_dirk@famcorsus.com" Md="" Password="mywolfgang" NumAccounts="1" Username0="dirk@famcorsus.com" Md0="" Password0="mywolfgang"/>
            '</MediaServers>
            'MyThirdPartyMediaServices = Value
        ElseIf StateVarName = "ThirdPartyMediaServersX" Then ' if you ask me, this is ThirdPartyMediaServers in encrypted form
            'If g_bDebug Then Log( "ZonegroupTopology callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString)
            ' this apparantly called when an iPhone is being inserted or a windows DLNA server is detected
            MyThirdPartyMediaServicesX = Value
        ElseIf StateVarName = "AlarmRunSequence" Then
            If g_bDebug Then Log("ZonegroupTopology callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
            ' the alarm just went off. The Value is something like RINCONxxxx:yy:z    
            ' RINCON_000E5824C3B001400:66:14 -->>> you get this at start-up
            If GetUDN() = GoFindTheAlarmZone(GetUDN) Then
                PlayChangeNotifyCallback(player_status_change.AlarmStart, player_state_values.Playing)
            End If
        End If
    End Sub

    Private Sub ZonegroupTopologyControlDied() Handles myZonegroupTopologyCallback.ControlDied
        Log("ZonegroupTopology Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in ZonegroupTopologyDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in ZonegroupTopologyDied Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub ContentDirectoryStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myContentDirectoryCallback.ControlStateChange
        ' ContentDirectoryControlStateChange
        If g_bDebug Then Log("ContentDirectoryStateChange for ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        'Log( "ContentDirectoryStateChange for ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString)
        ' Browseable = Boolena
        ' RecentlyPlayedUpdateID = String
        ' ShareListUpdateID = String
        ' SavedQueuesUpdateID = String
        ' UserRadioUpdateID = String
        ' ShareIndexLastError = String
        ' ShareIndexInProgress = String
        ' ShareListRefreshState = NOTRUN RUNNING DONE
        ' ContainerUpdateIDs = String
        ' SystemUpdateID = UI4

        ' dcor nothing realy implemented yet
    End Sub

    Private Sub ContentDirectoryDied() Handles myContentDirectoryCallback.ControlDied
        Log("Content Directory Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in ContentDirectory.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in ContentDirectoryDied Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub


    Private Sub ConnectionManagerStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myConnectionManagerCallback.ControlStateChange
        'If g_bDebug Then Log( "ConnectionManagerStateChange for ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString)
    End Sub

    Private Sub ConnectionManagerDied() Handles myConnectionManagerCallback.ControlDied
        Log("ConnectionManagerControlDied for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in ConnectionManagerControlDied.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in ConnectionManagerControlDied Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub QueueServiceStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myQueueServiceCallback.ControlStateChange
        If g_bDebug Then Log("QueueServiceStateChange callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)
        ' 
        If StateVarName = "LastChange" Then
            ' it appears all updates are done through XML
            If Value <> "" Then
                MyQueueServiceLastInfo = Value
                Try
                    Dim xmlData As XmlDocument = New XmlDocument
                    xmlData.LoadXml(Value)
                    ' examples:
                    ' <Event xmlns="urn:schemas-sonos-com:metadata-1-0/Queue/"><QueueID val="0"><UpdateID val="34"/><Curated val="0"/><QueueOwnerID val=""/></QueueID><QueueID val="5"><UpdateID val="1"/><Curated val="0"/><QueueOwnerID val=""/></QueueID></Event>
                    ' <Event xmlns="urn:schemas-sonos-com:metadata-1-0/Queue/"><QueueID val="6"><UpdateID val="2"/><QueueOwnerID val="alexa.bridgeAlexa"/></QueueID><QueueID val="5"><UpdateID val="0"/></QueueID></Event>
                    ' <Event xmlns="urn:schemas-sonos-com:metadata-1-0/Queue/"><QueueID val="6"><UpdateID val="3"/></QueueID></Event>
                    ' fields of interest are QueueID, QueueOwnerID, UpdateID
                    ' I wonder when the UpdateID is set to 0 whether that means the queue is released??
                    ' note !!! there can be more than one node of QueueID
                Catch ex As Exception
                    Log("Error in QueueServiceStateChange callback for ZonePlayer " & ZoneName & " with XML = " & Value.ToString & " And Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            End If
        End If

    End Sub

    Private Sub QueueServiceDied() Handles myQueueServiceCallback.ControlDied
        Log("QueueServiceStateChange Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in QueueServiceStateChange.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in QueueServiceStateChange Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub



    Private Sub VirtualLineInStateChange(ByVal StateVarName As String, ByVal Value As String) Handles myVirtualLineInCallback.ControlStateChange
        If g_bDebug Then Log("VirtualLineInStateChange callback ZonePlayer " & ZoneName & ": Var Name = " & StateVarName & " Value = " & Value.ToString, LogType.LOG_TYPE_INFO)

    End Sub

    Private Sub VirtualLineInDied() Handles myVirtualLineInCallback.ControlDied
        Log("VirtualLineInCallback Callback Died. ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Log("UPnP connection to ZonePlayer " & ZoneName & " was lost in VirtualLineInCallback.", LogType.LOG_TYPE_INFO)
            Disconnect(False)
        Catch ex As Exception
            Log("ERROR: Something went wrong in VirtualLineInCallback Callback for ZonePlayer - " & ZoneName & " Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function GetSeconds(ByVal Time As String) As Integer
        GetSeconds = 0
        If Time = "" Then Exit Function
        Try
            Dim Conversion
            Conversion = Split(Time, ":")
            GetSeconds = (CInt(Conversion(0)) * 60 * 60) + (CInt(Conversion(1)) * 60) + CInt(Conversion(2))
        Catch ex As Exception
            GetSeconds = 0
        End Try
    End Function


    Public Function SeekTrack(ByVal TrackNumber As String)
        SeekTrack = ""
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SeekTrack(TrackNumber)
            Catch ex As Exception
                If g_bDebug Then Log("SeekTrack called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If g_bDebug Then Log("SeekTrack called for zoneplayer = " & ZoneName & " with TrackNumber = " & TrackNumber, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)

            InArg(0) = 0                    ' UI4
            InArg(1) = "TRACK_NR"           ' String
            InArg(2) = TrackNumber          ' String

            AVTransport.InvokeAction("Seek", InArg, OutArg)

            SeekTrack = "OK"

        Catch ex As Exception
            If g_bDebug Then Log("ERROR: In Seek Track for zoneplayer = " & ZoneName & " with TrackNbr = " & TrackNumber & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SeekTime(ByVal NewTime As String)
        SeekTime = "" ' use this function to change the trackposition. The input needs to be in the xx:xx:xx format!
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If g_bDebug Then Log("SeekTime called for zoneplayer = " & ZoneName & " with Time = " & NewTime, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = 0            ' UI4
            InArg(1) = "REL_TIME"   ' String
            InArg(2) = NewTime      ' String
            AVTransport.InvokeAction("Seek", InArg, OutArg)
            SeekTime = "OK"
        Catch ex As Exception
            If g_bDebug Then Log("ERROR: In Seek Time for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


    Private Function UPnP_Error(ByVal ErrNumber) As String
        'If (ErrNumber < &H8004042B) And (ErrNumber >= &H80040300) Then ' UPNP_E_ACTION_SPECIFIC_BASE <= UPnP_Error <= UPNP_E_ACTION_SPECIFIC_MAX 
        Dim SaveError
        SaveError = ErrNumber
        If (ErrNumber < &H8004042B) And (ErrNumber >= &H80040200) Then
            ErrNumber = ErrNumber - &H80040300 + &H258
        ElseIf ErrNumber < 300 Then
            ErrNumber = ErrNumber + 700
        End If
        'UPNP_E_ROOT_ELEMENT_EXPECTED = 0x80040200, = 344 dec
        'UPNP_E_DEVICE_ELEMENT_EXPECTED = 0x80040201,
        'UPNP_E_SERVICE_ELEMENT_EXPECTED = 0x80040202,
        'UPNP_E_SERVICE_NODE_INCOMPLETE = 0x80040203,
        'UPNP_E_DEVICE_NODE_INCOMPLETE = 0x80040204,
        'UPNP_E_ICON_ELEMENT_EXPECTED = 0x80040205,
        'UPNP_E_ICON_NODE_INCOMPLETE = 0x80040206,
        'UPNP_E_INVALID_ACTION = 0x80040207,
        'UPNP_E_INVALID_ARGUMENTS = 0x80040208,
        'UPNP_E_OUT_OF_SYNC = 0x80040209,
        'UPNP_E_ACTION_REQUEST_FAILED = 0x80040210,
        'UPNP_E_TRANSPORT_ERROR = 0x80040211,
        'UPNP_E_VARIABLE_VALUE_UNKNOWN = 0x80040212,
        'UPNP_E_INVALID_VARIABLE = 0x80040213,
        'UPNP_E_DEVICE_ERROR = 0x80040214,
        'UPNP_E_PROTOCOL_ERROR = 0x80040215,
        'UPNP_E_ERROR_PROCESSING_RESPONSE = 0x80040216,
        'UPNP_E_DEVICE_TIMEOUT = 0x80040217,

        Select Case ErrNumber
            Case 0
                UPnP_Error = "Successfull Action"
            Case 344
                UPnP_Error = "UPNP_E_ROOT_ELEMENT_EXPECTED"
            Case 345
                UPnP_Error = "UPNP_E_DEVICE_ELEMENT_EXPECTED "
            Case 346
                UPnP_Error = "UPNP_E_SERVICE_ELEMENT_EXPECTED"
            Case 347
                UPnP_Error = "UPNP_E_SERVICE_NODE_INCOMPLETE"
            Case 348
                UPnP_Error = "UPNP_E_DEVICE_NODE_INCOMPLETE"
            Case 349
                UPnP_Error = "UPNP_E_ICON_ELEMENT_EXPECTED"
            Case 350
                UPnP_Error = "UPNP_E_ICON_NODE_INCOMPLETE"
            Case 351
                UPnP_Error = "UPNP_E_INVALID_ACTION"
            Case 352
                UPnP_Error = "UPNP_E_INVALID_ARGUMENTS "
            Case 353
                UPnP_Error = "UPNP_E_OUT_OF_SYNC"
            Case 360
                UPnP_Error = "UPNP_E_ACTION_REQUEST_FAILED "
            Case 361
                UPnP_Error = "UPNP_E_TRANSPORT_ERROR "
            Case 362
                UPnP_Error = "UPNP_E_VARIABLE_VALUE_UNKNOWN "
            Case 363
                UPnP_Error = "UPNP_E_INVALID_VARIABLE"
            Case 364
                UPnP_Error = "UPNP_E_DEVICE_ERROR "
            Case 365
                UPnP_Error = "UPNP_E_PROTOCOL_ERROR "
            Case 366
                UPnP_Error = "UPNP_E_ERROR_PROCESSING_RESPONSE "
            Case 367
                UPnP_Error = "UPNP_E_DEVICE_TIMEOUT"
            Case 351
                UPnP_Error = "UPNP_E_INVALID_DOCUMENT"
            Case 352
                UPnP_Error = "UPNP_E_EVENT_SUBSCRIPTION_FAILED"
            Case 401
                UPnP_Error = "Invalid Action"
            Case 402
                UPnP_Error = "Invalid args"
            Case 403
                UPnP_Error = "Out of Sync"
            Case 404
                UPnP_Error = "Invalid Var"
            Case 501
                UPnP_Error = "Action failed"
            Case 701
                UPnP_Error = "No such object / Transition no available / Incompatible protocol info"
            Case 702
                UPnP_Error = "Invalid CurrentTagValue / Invalid InstanceID / No contents / Incompatible directions"
            Case 703
                UPnP_Error = "Invalid NewTagValue / Read error / Insufficient network resources"
            Case 704
                UPnP_Error = "Required tag / Format not supported for playback / Local restrictions"
            Case 705
                UPnP_Error = "Read only tag / Transport is locked / Access denied"
            Case 706
                UPnP_Error = "Paramater Mismatch / Write Error / Invalid connection reference"
            Case 707
                UPnP_Error = "Media is protected / Not writable / Not in Network"
            Case 708
                UPnP_Error = "Unsupported or invalid search criteria / Format no supported for recording"
            Case 709
                UPnP_Error = "Unsupported or invalid sort criteria / Media is full"
            Case 710
                UPnP_Error = "No such container / Seek mode not supported"
            Case 711
                UPnP_Error = "Restricted object / Illegal seek target"
            Case 712
                UPnP_Error = "Bad metadata / Playmode not supported"
            Case 713
                UPnP_Error = "Restricted parent object"
            Case 714
                UPnP_Error = "No such source resource / Invalid MIME-type"
            Case 715
                UPnP_Error = "Source resource access denied / Content 'BUSY'"
            Case 716
                UPnP_Error = "Transfer busy / Resource not found"
            Case 717
                UPnP_Error = "No such file transfer / Play speed not supported"
            Case 718
                UPnP_Error = "No such destination source / Invalid InstanceID"
            Case 719
                UPnP_Error = "Destination resource access denied"
            Case 720
                UPnP_Error = "Cannot process the request"
            Case 801
                UPnP_Error = "Access denied"
            Case 802
                UPnP_Error = "Not Enough Room"
            Case Else
                If ErrNumber >= 600 And ErrNumber <= 699 Then
                    UPnP_Error = "Common action error - undefined. Error=" & SaveError.ToString
                ElseIf ErrNumber >= 800 And ErrNumber <= 899 Then
                    UPnP_Error = "Sonos error - undefined. Error=" & SaveError.ToString
                Else
                    UPnP_Error = ErrNumber & ": Unknown error type. OrgError=" & SaveError.ToString
                End If
        End Select
    End Function


    Public Function DestroyObject(ByVal ObjectID As String) As String
        DestroyObject = ""
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("DestroyObject called with ObjectID = " & ObjectID & " for Zone = " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = ObjectID
            ContentDirectory.InvokeAction("DestroyObject", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in DestroyObject for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            DestroyObject = "NOK"
        End Try
    End Function


    Public Function GetMuteState(ByVal Channel As String) As Boolean
        GetMuteState = False
        If g_bDebug Then Log("GetMuteState called. DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' need to forward this
            Return MyWirelessDockDestinationPlayer.GetMuteState(Channel)
            Exit Function
        End If
        ' removed this 7/13/2019 in v3.1.0.34 It appears you can get the volume from the slave player, it will actually show the right side volume. If you however set it for a slave, it will cause balance to shift
        'If ZoneIsASlave Then
        'Dim LinkedZone As HSPI  'HSMusicAPI
        'LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
        'Try
        'Return LinkedZone.GetMuteState(Channel)
        'Catch ex As Exception
        'If g_bDebug Then Log("GetMuteState called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        'Return False
        'End Try
        'Exit Function
        'End If
        If DeviceStatus = "Offline" Or RenderingControl Is Nothing Or MyZoneModel = "WD100" Then Exit Function
        Try
            Dim InArg(1)                'InstanceID UI4
            Dim OutArg(0)               'Channel String
            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                InArg(0) = 0
                InArg(1) = Channel
            Else
                GetMuteState = "Error in GetMuteState: strChannel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If
            RenderingControl.InvokeAction("GetMute", InArg, OutArg)
            GetMuteState = OutArg(0) ' boolean
            'If g_bDebug Then Log( "GetMuteState called. Value =" & GetMuteState)
        Catch ex As Exception
            Log("ERROR in GetMuteState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetVolumeLevel(ByVal Channel As String) As String
        GetVolumeLevel = ""
        If g_bDebug Then Log("GetVolumeLevel called for ZonePlayer = " & ZoneName & " with values Channel=" & Channel & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' need to forward this
            Return MyWirelessDockDestinationPlayer.GetVolumeLevel(Channel)
            Exit Function
        End If
        ' removed this 7/13/2019 in v3.1.0.34 It appears you can get the volume from the slave player, it will actually show the right side volume. If you however set it for a slave, it will cause balance to shift
        'If ZoneIsASlave Then
        'Dim LinkedZone As HSPI  'HSMusicAPI
        'LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
        'Try
        'Return LinkedZone.GetVolumeLevel(Channel)
        'Catch ex As Exception
        'If g_bDebug Then Log("GetVolumeLevel called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        'Return ""
        'End Try
        'Exit Function
        'End If
        If DeviceStatus = "Offline" Or RenderingControl Is Nothing Or MyZoneModel = "WD100" Then Exit Function
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = 0
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetVolume", InArg, OutArg)
            GetVolumeLevel = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetVolumeLevel for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetLoudnessState(ByVal Channel As String) As Boolean
        GetLoudnessState = False
        If g_bDebug Then Log("GetLoudnessState called for ZonePlayer = " & ZoneName & " with values Channel=" & Channel & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' need to forward this
            Return MyWirelessDockDestinationPlayer.GetLoudnessState(Channel)
            Exit Function
        End If
        ' removed this 7/13/2019 in v3.1.0.34 It appears you can get the volume from the slave player, it will actually show the right side volume. If you however set it for a slave, it will cause balance to shift
        'If ZoneIsASlave Then
        'Dim LinkedZone As HSPI  'HSMusicAPI
        'LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
        'Try
        'Return LinkedZone.GetLoudnessState(Channel)
        'Catch ex As Exception
        'If g_bDebug Then Log("GetLoudnessState called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        'Return False
        'End Try
        'Exit Function
        'End If
        If DeviceStatus = "Offline" Or RenderingControl Is Nothing Or MyZoneModel = "WD100" Then Exit Function
        Try
            Dim InArg(1)                'InstanceID UI4
            Dim OutArg(0)               'Channel String
            InArg(0) = 0
            InArg(1) = Channel
            RenderingControl.InvokeAction("GetLoudness", InArg, OutArg)
            GetLoudnessState = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetLoudnessState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ToggleLoudnessState(ByVal Channel As String) As String
        ToggleLoudnessState = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim LoudnessState
        LoudnessState = GetLoudnessState(Channel)
        If LoudnessState Then
            LoudnessState = False
        Else
            LoudnessState = True
        End If
        SetLoudnessState(Channel, LoudnessState)
    End Function

    Public Function SetLoudnessState(ByVal Channel As String, ByVal NewState As Boolean) As String
        SetLoudnessState = ""
        If SuperDebug Then Log("SetLoudnessState called for zoneplayer = " & ZoneName & " with Channel = " & Channel & " and NewState = " & NewState & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' need to forward this
            MyWirelessDockDestinationPlayer.SetLoudnessState(Channel, NewState)
            Exit Function
        End If
        If ZoneIsASlave Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
            Try
                SetLoudnessState = LinkedZone.SetLoudnessState(Channel, NewState)
            Catch ex As Exception
                If g_bDebug Then Log("SetLoudnessState called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Or RenderingControl Is Nothing Or MyZoneModel = "WD100" Then Exit Function
        Try
            Dim InArg(2)            'InstanceID UI4
            Dim OutArg(0)           'Channel String
            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                InArg(0) = 0
                InArg(1) = Channel
                InArg(2) = NewState ' Disired Loudness Boolean
            Else
                SetLoudnessState = "Error: strChannel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If
            RenderingControl.InvokeAction("SetLoudness", InArg, OutArg)
            SetLoudnessState = "OK"
        Catch ex As Exception
            Log("ERROR in SetLoudnessState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function SetMute(ByVal Channel As String, ByVal NewState As Boolean) As String
        SetMute = ""
        If SuperDebug Then Log("SetMute called for zoneplayer = " & ZoneName & " with Channel = " & Channel & " and NewState = " & NewState & " and DeviceStatus =" & DeviceStatus, LogType.LOG_TYPE_INFO)
        If ZoneModel = "WD100" And MyWirelessDockDestinationPlayer IsNot Nothing Then ' need to forward this
            MyWirelessDockDestinationPlayer.SetMute(Channel, NewState)
            Exit Function
        End If
        If ZoneIsASlave Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
            Try
                SetMute = LinkedZone.SetMute(Channel, NewState)
            Catch ex As Exception
                If g_bDebug Then Log("SetMute called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Or RenderingControl Is Nothing Or MyZoneModel = "WD100" Then Exit Function
        Try
            Dim InArg(2)
            Dim OutArg(0)

            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                InArg(0) = 0
                InArg(1) = Channel
                InArg(2) = NewState
            Else
                SetMute = "Error: strChannel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If

            RenderingControl.InvokeAction("SetMute", InArg, OutArg)

            SetMute = "OK"
        Catch ex As Exception
            Log("ERROR in SetMute for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ToggleMuteState(ByVal Channel As String) As String
        ToggleMuteState = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim MuteState As Boolean
        MuteState = GetMuteState(Channel)
        If MuteState Then
            MuteState = False
        Else
            MuteState = True
        End If
        ToggleMuteState = SetMute(Channel, MuteState)
    End Function


    Public Function SetVolumeLevel(ByVal Channel As String, ByVal NewLevel As Integer) As String
        SetVolumeLevel = ""
        If g_bDebug Then Log("SetVolumeLevel called for ZonePlayer = " & ZoneName & " with values Channel=" & Channel & " Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
        If ZoneIsASlave Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            Try
                LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
                SetVolumeLevel = LinkedZone.SetVolumeLevel(Channel, NewLevel)
            Catch ex As Exception
                If g_bDebug Then Log("SetVolumeLevel called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Or RenderingControl Is Nothing Or MyZoneModel = "WD100" Then Exit Function
        If MyCurrentFixedVolumeState Then Exit Function
        Try
            Dim InArg(2)
            Dim OutArg(0)

            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                InArg(0) = 0
                InArg(1) = Channel
                If NewLevel < 0 Then
                    NewLevel = 0
                ElseIf NewLevel > 100 Then
                    NewLevel = 100
                End If
                InArg(2) = NewLevel
            Else
                Log("ERROR in SetVolumeLevel for zoneplayer = " & ZoneName & ". strChannel = " & Channel & " and Channel value must be 'Master', 'LF', or 'RF'", LogType.LOG_TYPE_INFO)
                SetVolumeLevel = "Error: Channel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If
            Try
                RenderingControl.InvokeAction("SetVolume", InArg, OutArg)
                SetVolumeLevel = "OK"
            Catch ex As Exception
                Log("ERROR in SetVolumeLevel for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("ERROR in SetVolumeLevel for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ChangeVolumeLevel(ByVal Channel As String, ByVal NewLevel As Integer) As String
        ChangeVolumeLevel = ""
        If g_bDebug Then Log("ChangeVolumeLevel called for ZonePlayer = " & ZoneName & " for Channel=" & Channel & " Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
        If ZoneIsASlave Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
            Try
                ChangeVolumeLevel = LinkedZone.ChangeVolumeLevel(Channel, NewLevel)
            Catch ex As Exception
                If g_bDebug Then Log("ChangeVolumeLevel called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Or RenderingControl Is Nothing Or MyZoneModel = "WD100" Then Exit Function
        If MyCurrentFixedVolumeState Then Exit Function
        Try
            If Channel = "Master" Or Channel = "LF" Or Channel = "RF" Then
                Dim InArg(1)
                Dim OutArg(0)
                InArg(0) = 0
                InArg(1) = Channel
                Try
                    RenderingControl.InvokeAction("GetVolume", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in ChangeVolumeLevel for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                NewLevel = OutArg(0) + NewLevel
                If NewLevel < 0 Then
                    NewLevel = 0
                ElseIf NewLevel > 100 Then
                    NewLevel = 100
                End If
            Else
                ChangeVolumeLevel = "Error: strChannel value must be 'Master', 'LF', or 'RF'"
                Exit Function
            End If
            Try
                Dim InArg(2)
                Dim OutArg(0)
                InArg(0) = 0
                InArg(1) = Channel
                InArg(2) = NewLevel
                RenderingControl.InvokeAction("SetVolume", InArg, OutArg)
                ChangeVolumeLevel = "OK"
                If g_bDebug Then Log("ChangeVolumeLevel called for ZonePlayer = " & ZoneName & " for Channel=" & Channel & " New Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
            Catch ex As Exception
                Log("Error in ChangeVolumeLevel for zoneplayer = " & ZoneName & " for Channel=" & Channel & " New Value=" & NewLevel.ToString & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Catch ex As Exception
            Log("Error in ChangeVolumeLevel for zoneplayer = " & ZoneName & " for Channel=" & Channel & " New Value=" & NewLevel.ToString & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function ChangeBalanceLevel(ByVal Channel As String, ByVal NewLevel As Integer) As String
        ChangeBalanceLevel = ""
        If g_bDebug Then Log("ChangeBalanceLevel called for ZonePlayer = " & ZoneName & " with values Channel=" & Channel & " Value=" & NewLevel.ToString, LogType.LOG_TYPE_INFO)
        If ZoneIsASlave Then
            Dim LinkedZone As HSPI  'HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(ZoneMasterUDN)
            Try
                ChangeBalanceLevel = LinkedZone.ChangeBalanceLevel(Channel, NewLevel)
            Catch ex As Exception
                If g_bDebug Then Log("ChangeBalanceLevel called for Zone - " & ZoneName & " which was linked to " & ZoneMasterUDN.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Then Exit Function
        If (Channel <> "LF" And Channel <> "RF") Or NewLevel = 0 Then
            ' wrong input
            Exit Function
        End If
        Dim LeftLevel, RightLevel As Integer
        LeftLevel = GetVolumeLevel("LF")
        RightLevel = GetVolumeLevel("RF")
        ' when the balance is in the middle both LF and RF are equal to 100
        ' if the balance is to the left from midpoint and moving left then only left should be touched else right and left may need ajustment
        Dim DirectionLeft As Boolean
        If Channel = "LF" And NewLevel > 0 Then
            ' we're moving to the left
            DirectionLeft = True
        ElseIf Channel = "LF" And NewLevel < 0 Then
            ' we're moving to the right
            DirectionLeft = False
        ElseIf Channel = "RF" And NewLevel > 0 Then
            ' we're moving to the right
            DirectionLeft = False
        Else
            ' we're moving to the left
            DirectionLeft = True
        End If
        If DirectionLeft Then
            ' decrease right unless left < 100 then increase left first
            If LeftLevel < 100 Then
                ' adjust right and left levels
                ChangeVolumeLevel("LF", NewLevel)
            Else
                ChangeVolumeLevel("RF", -NewLevel)
            End If
        Else
            ' decrease left unless Right < 100 then increase right first
            If RightLevel < 100 Then
                ' adjust right and left levels
                ChangeVolumeLevel("RF", NewLevel)
            Else
                ChangeVolumeLevel("LF", -NewLevel)
            End If
        End If
    End Function


    Private Function GetMatchingTracks(ByVal artist As String, ByVal album As String, ByVal genre As String, ByVal Track As String, Optional ByVal iPodDBName As String = "") As System.Array
        'Returns a list of tracks matching the parameters provided by Artist, Album, and Genre.  If empty strings are provided as parameters, then all tracks from the library are returned.  
        'Note, only the track names are returned, duplicates included.  
        If g_bDebug Then Log("GetMatchingTracks called for Zone - " & ZoneName & " with Artist=" & artist & " Album=" & album & " Genre=" & genre, LogType.LOG_TYPE_INFO)
        Dim MyTracks As String()
        MyTracks = {""}
        Dim MyTrackNbrs As Integer()
        MyTrackNbrs = {}
        GetMatchingTracks = MyTracks
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        Dim tempstring As String
        Dim Index As Integer = 0
        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" And iPodDBName = "" Then Exit Function
            If iPodDBName = "" Then iPodDBName = MyDockediPodPlayerName
            ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & MyDockediPodPlayerName.ToString & ".sdb"
        Else
            ConnectionString = "Data Source=" & CurrentAppPath & MusicDBPath
        End If
        Dim QueryString As String
        If genre <> "" Then
            QueryString = "SELECT * FROM Tracks WHERE Genre='" & PrepareForQuery(genre) & "'"
        Else
            QueryString = "SELECT * FROM Tracks"
            If artist <> "" And album <> "" Then
                QueryString = QueryString & " WHERE Artist = '" & PrepareForQuery(artist) & "' AND Album='" & PrepareForQuery(album) & "'"
            ElseIf artist <> "" Then
                QueryString = QueryString & " WHERE Artist = '" & PrepareForQuery(artist) & "'"
            ElseIf album <> "" Then
                QueryString = QueryString & " WHERE Album='" & PrepareForQuery(album) & "'"
            End If
        End If
        If album <> "" Then QueryString = QueryString & " ORDER BY TrackNo"
        If g_bDebug Then Log("GetMatchingTracks for " & ZoneName & " with Query=" & QueryString, LogType.LOG_TYPE_INFO)

        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("GetMatchingTracks unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        Try
            Index = 0
            While SQLreader.Read()
                If SQLreader("ParentID").ToString = "A:TRACKS" Then ' Recordset.Fields("ParentID").Value.ToString = "A:TRACKS" Then
                    tempstring = SQLreader("Name").ToString
                    'Log( "Record #" & Index.ToString & " with value " & tempstring)
                    If SQLreader("Album").ToString <> "All Tracks" Then ' this is how the WD100 does it
                        ReDim Preserve MyTracks(Index)
                        MyTracks(Index) = tempstring
                        Index = Index + 1
                    End If
                ElseIf genre <> "" And SQLreader("ParentID").ToString = "A:GENRE" Then
                    Dim Artists
                    Dim ArtistsIndex As Integer = 0
                    Dim Albums
                    Dim Tracks
                    Try
                        Artists = GetTracks(SQLreader("Id").ToString, False, True)
                    Catch ex As Exception
                        Artists = Nothing
                        Log("Error GetMatchingTracks for zoneplayer " & ZoneName & "unable to put in tracks with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        While ArtistsIndex < UBound(Artists, 1)
                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & ArtistsIndex.ToString & " for Artist (0) with value " & Artists(ArtistsIndex, 0).ToString)
                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & ArtistsIndex.ToString & " for Artist (1) with value " & Artists(ArtistsIndex, 1).ToString)
                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & ArtistsIndex.ToString & " for Artist (2) with value " & Artists(ArtistsIndex, 2).ToString) 'name
                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & ArtistsIndex.ToString & " for Artist (3) with value " & Artists(ArtistsIndex, 3).ToString)
                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & ArtistsIndex.ToString & " for Artist (6) with value " & Artists(ArtistsIndex, 6).ToString) ' ID
                            If Artists(ArtistsIndex, 0) = artist Or (artist = "" And (Artists(ArtistsIndex, 0) <> "All" And Artists(ArtistsIndex, 0) <> "All Albums")) Then ' All Albums is WD100
                                ' now pull up its list of albums
                                'If g_bDebug Then Log("GetMatchingTracks Record #" & ArtistsIndex.ToString & " for Artist (0) with value " & Artists(ArtistsIndex, 0).ToString, LogType.LOG_TYPE_INFO)

                                Albums = GetTracks(Artists(ArtistsIndex, 6), False, True) ' this is the track id so we can pull up album

                                Dim AlbumIndex As Integer = 0
                                For AlbumIndex = 0 To UBound(Albums, 1) - 1
                                    'If g_bDebug Then Log( "GetMatchingTracks Record #" & AlbumIndex.ToString & " for Album (0) with value " & Albums(AlbumIndex, 0).ToString)
                                    'If g_bDebug Then Log( "GetMatchingTracks Record #" & AlbumIndex.ToString & " for Album (1) with value " & Albums(AlbumIndex, 1).ToString)
                                    'If g_bDebug Then Log( "GetMatchingTracks Record #" & AlbumIndex.ToString & " for Album (2) with value " & Albums(AlbumIndex, 2).ToString) 'name
                                    'If g_bDebug Then Log( "GetMatchingTracks Record #" & AlbumIndex.ToString & " for Album (3) with value " & Albums(AlbumIndex, 3).ToString)
                                    'If g_bDebug Then Log( "GetMatchingTracks Record #" & AlbumIndex.ToString & " for Album (6) with value " & Albums(AlbumIndex, 6).ToString) ' ID
                                    If Albums(AlbumIndex, 0) = album Or (album = "" And (Albums(AlbumIndex, 0) <> "All") And Albums(AlbumIndex, 0) <> "All Tracks") Then
                                        ' now go get the tracks of the album
                                        ' now pull up its list of albums
                                        'If g_bDebug Then Log("GetMatchingTracks Record #" & AlbumIndex.ToString & " for Album (0) with value " & Albums(AlbumIndex, 0).ToString, LogType.LOG_TYPE_INFO)

                                        Tracks = GetTracks(Albums(AlbumIndex, 6), True, False) ' this is the track id so we can pull up album

                                        Dim TrackIndex As Integer = 0
                                        For TrackIndex = 0 To UBound(Tracks, 1) - 1
                                            'If g_bDebug Then Log("GetMatchingTracks Record #" & TrackIndex.ToString & " for Tracks (0) with value " & Tracks(TrackIndex, 0).ToString, LogType.LOG_TYPE_INFO)
                                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & TrackIndex.ToString & " for Tracks (1) with value " & Tracks(TrackIndex, 1).ToString)
                                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & TrackIndex.ToString & " for Tracks (2) with value " & Tracks(TrackIndex, 2).ToString) 'name
                                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & TrackIndex.ToString & " for Tracks (3) with value " & Tracks(TrackIndex, 3).ToString)
                                            'If g_bDebug Then Log( "GetMatchingTracks Record #" & TrackIndex.ToString & " for Tracks (6) with value " & Tracks(TrackIndex, 6).ToString) ' ID

                                            If Tracks(TrackIndex, 0) = Track Or Track = "" Then
                                                ReDim Preserve MyTracks(Index + 3)
                                                MyTracks(Index) = Tracks(TrackIndex, 0).ToString
                                                MyTracks(Index + 1) = Tracks(TrackIndex, 3).ToString
                                                MyTracks(Index + 2) = Tracks(TrackIndex, 6).ToString
                                                MyTracks(Index + 3) = Tracks(TrackIndex, 10).ToString ' added v.78 representing Track Nbr
                                                If g_bDebug Then Log("GetMatchingTracks for zoneplayer " & ZoneName & " Tracks " & Tracks(TrackIndex, 0).ToString, LogType.LOG_TYPE_INFO)
                                                Index = Index + 4 ' changed from 3 to 4 in v.78
                                            End If

                                        Next
                                    End If
                                Next
                            End If
                            ArtistsIndex = ArtistsIndex + 1
                        End While
                    Catch ex As Exception
                        Log("Error GetMatchingTracks for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                End If
            End While
            GetMatchingTracks = MyTracks
        Catch ex As Exception
            Log("Error GetMatchingTracks for zoneplayer " & ZoneName & " unable read record with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("GetMatchingTracks to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Function

    Public Sub PlayMusicOnZone(ByVal MusicDBName As String, ByVal ZoneUDN As String, ByVal SourceSonosPlayer As HSPI, ByVal Artist As String, Optional ByVal Album As String = "", Optional ByVal PlayList As String = "", Optional ByVal Genre As String = "", Optional ByVal Track As String = "", Optional ByVal StartWithArtist As String = "", Optional ByVal StartWithTrack As String = "", Optional ByVal TrackMatch As String = "", Optional ByVal Favorite As String = "", Optional ByVal ClearPlayerQueue As Boolean = False, Optional QueueAction As QueueActions = QueueActions.qaDontPlay)
        ' Search with provided info in the DB and call PlayURI
        ' ZoneUDN is typically empty except for WD100 players, it will hold the "x-sonos-dock:RINCONxxxx/"
        If g_bDebug Then Log("PlayMusicOnZone for Zone " & ZoneName & " and MusicDBName = " & MusicDBName & ", ZoneUDN = " & ZoneUDN & ", Artist = " & Artist & " Album = " & Album & " Playlist = " & PlayList & ", Genre = " & Genre & ", Track = " & Track & ", StartWithArtist = " & StartWithArtist & ", StartWithTrack = " & StartWithTrack & ", TrackMatch = " & TrackMatch & ", Favorite = " & Favorite & ", ClearPlayerQueue = " & ClearPlayerQueue.ToString & ", QueueAction = " & QueueAction.ToString, LogType.LOG_TYPE_INFO)
        If Artist = "" And Album = "" And PlayList = "" And Genre = "" And Track = "" And StartWithArtist = "" And StartWithTrack = "" And TrackMatch = "" And Favorite = "" Then Exit Sub
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Sub
        Dim xmlData As XmlDocument = New XmlDocument
        Dim ConnectionString As String
        Dim QueryString As String = "SELECT * FROM Tracks WHERE "

        ConnectionString = "Data Source=" & MusicDBName 'hs.GetAppPAth & MusicDBPath

        ' create Query String
        If Album <> "" And (Track <> "" Or StartWithTrack <> "") Then
            ' this overrules genre. The track is very specifically identified. Doing a seach with Genre is a pain in the ass
            Genre = ""
        End If
        If Genre <> "" Then
            ' this trumps almost everything
            QueryString = QueryString & " Name = 'All Genres' AND Genre = '" & PrepareForQuery(Genre) & "'"
        Else
            If (Artist <> "" Or StartWithArtist <> "") And Album <> "" And (Track <> "") Then ' If (Artist <> "" Or StartWithArtist <> "") And Album <> "" And (Track <> "" Or StartWithTrack <> "") Then '
                ' when we search for a track from an artist on a specific album, unfortunatly for albums by various artists, depending on how the track was navigated into
                ' we either need to search by creator (artist) or Albumartist. 
                If Artist <> "" Then QueryString = QueryString & " Artist = '" & PrepareForQuery(Artist) & "' AND "
            Else
                If Artist <> "" Then QueryString = QueryString & " Artist = '" & PrepareForQuery(Artist) & "' AND "
            End If
            If Album <> "" Then
                If Track = "" And StartWithTrack = "" Then '
                    QueryString = QueryString & " Name = 'All Tracks' AND Album = '" & PrepareForQuery(Album) & "' AND "
                Else
                    QueryString = QueryString & " Album = '" & PrepareForQuery(Album) & "' AND "
                End If
            Else
                If PlayList = "" Then
                    QueryString = QueryString & " Name <> 'All Tracks' AND "
                End If
            End If

            If PlayList <> "" Then
                ' another special case. This can be a Playlist, a track from a playlist or a radiostation
                If Mid(PlayList, 1, 14) = "RadioStation: " Then
                    QueryString = QueryString & " Name = 'All RadioStations' AND Artist = '" & PrepareForQuery(PlayList) & "' AND "
                ElseIf Mid(PlayList, 1, 11) = "Audiobook: " Then
                    QueryString = QueryString & " Name = 'All Audiobooks' AND Artist = '" & PrepareForQuery(PlayList) & "' AND "
                ElseIf Mid(PlayList, 1, 9) = "Podcast: " Then
                    QueryString = QueryString & " Name = 'All Podcasts' AND Artist = '" & PrepareForQuery(PlayList) & "' AND "
                Else
                    If Track = "" Then 'If Track = "" And StartWithTrack = "" Then '
                        QueryString = QueryString & " Name = 'All Playlists' AND Artist = '" & PrepareForQuery(PlayList) & "' AND "
                    Else
                        ' this is problematic as I have no way (yet) to store Track+Playlist. This would require first to get tracks from the  playlist
                        ' then search for track, then play. So let's get the play list first, we can later retrieve the exact track 
                        QueryString = QueryString & " Name = 'All Playlists' AND Artist = '" & PrepareForQuery(PlayList) & "' AND "
                    End If
                End If
            End If
            If Track <> "" And PlayList = "" Then QueryString = QueryString & " Name = '" & PrepareForQuery(Track) & "' AND "
            If StartWithArtist <> "" Then QueryString = QueryString & " Artist = '" & PrepareForQuery(StartWithArtist) & "' AND "
            'If StartWithTrack <> "" And PlayList = "" Then QueryString = QueryString & " Name = '" & PrepareForQuery(StartWithTrack) & "'"
        End If

        QueryString = Trim(QueryString)
        If Mid(QueryString, Len(QueryString) - 2, 3) = "AND" Then Mid(QueryString, Len(QueryString) - 2, 3) = "   "
        QueryString = Trim(QueryString)

        If Album <> "" Then QueryString = QueryString & " ORDER BY TrackNo"

        If Favorite <> "" Then
            QueryString = "SELECT * FROM Tracks WHERE `ParentID`='FV:2' and Artist = '" & PrepareForQuery(Favorite) & "'"
        End If

        If g_bDebug Then Log("PlayMusicOnZone for Zone " & ZoneName & " and query = " & QueryString, LogType.LOG_TYPE_INFO)
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("PlayMusicOnZone unsuccesful query  for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("PlayMusicOnZone - Query string: " & QueryString, LogType.LOG_TYPE_ERROR)
            Log("PlayMusicOnZone - Query: Artist=" & Artist & " and Album=" & Album & " and Playlist=" & PlayList & "and Genre=" & Genre & " and Track=" & Track, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        If Not SQLreader.HasRows Then
            If g_bDebug Then Log("PlayMusicOnZone for zoneplayer = " & ZoneName & " found No Tracks", LogType.LOG_TYPE_INFO)
            Exit Sub
        End If

        Dim TrackURI As String = ""
        Dim TrackIndex As Integer = 1
        Dim TrackMetaData As String = ""
        Dim StartTrackIndex As Integer = 0
        Dim CurrentTrackIndexInQueue As Integer = 0

        If ClearPlayerQueue Then
            SourceSonosPlayer.SetTransportState("Stop")
            ClearQueue()
        End If

        If ZoneIsLinked Then
            Unlink()
            wait(1)
        End If

        Dim InArg(0)

        ' go figure out where in the queue things are by doing a getpositioninfo
        InArg(0) = 0 ' InstanceID = 0
        Dim PositionInfo(7)
        Try
            AVTransport.InvokeAction("GetPositionInfo", InArg, PositionInfo)
            CurrentTrackIndexInQueue = PositionInfo(pmiTrack) + 1
            TrackIndex = CurrentTrackIndexInQueue
            If g_bDebug Then Log("PlayMusicOnZone for zoneplayer = " & ZoneName & " found TrackIndex = " & TrackIndex.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("ERROR in PlayMusicOnZone while doing a GetPositionInfo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Dim MediaInfo(8)
        InArg = {""}
        MediaInfo = {0, "", "", "", "", "", "", "", ""}
        Dim CurrentURI As String
        Dim PlayingFromQueue As Boolean = False
        InArg(0) = 0 ' InstanceID = 0
        Try
            AVTransport.InvokeAction("GetMediaInfo", InArg, MediaInfo)
            CurrentURI = MediaInfo(gmiCurrentURI)
            If CurrentURI.ToString.IndexOf("x-rincon-queue:") <> -1 Then
                PlayingFromQueue = True
                If QueueAction = QueueActions.qaPlayNow Or QueueAction = QueueActions.qaDontPlay Then ' Means it goes to end of queue
                    CurrentTrackIndexInQueue = MediaInfo(gmiNrTracks) + 1
                    TrackIndex = CurrentTrackIndexInQueue
                End If
            Else
                ' reset trackindex any additions need to go to the end of the queue
                TrackIndex = 0
            End If
        Catch ex As Exception
            Log("ERROR in PlayMusicOnZone while doing GetMediaInfo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Dim PlayRadiostation As Boolean = False

        Try
            While SQLreader.Read()
                TrackURI = ZoneUDN & SQLreader("URI").ToString
                If SQLreader("Name").ToString = "All RadioStations" Then
                    ' don't add to Queue, play now
                    SourceSonosPlayer.SetTransportState("Stop")
                    PlayURI(TrackURI, SQLreader("Id").ToString)
                    PlayRadiostation = True
                    Exit Try
                ElseIf SQLreader("ParentID").ToString = "FV:2" Then
                    ' don't add to Queue, play now
                    SourceSonosPlayer.SetTransportState("Stop")
                    PlayURI(TrackURI, SQLreader("Id").ToString, True)
                    PlayRadiostation = True
                    Exit Try
                ElseIf SQLreader("Name").ToString = "All Playlists" Then
                    ' this is a playlist. If a track was specified, go find it
                    If Track = "" And StartWithTrack = "" And TrackMatch = "" Then
                        TrackMetaData = SourceSonosPlayer.GetTrackMetaData(SQLreader("Id").ToString, False)
                        If ZoneUDN <> "" Then
                            PlayURI(TrackURI, "") ' TrackMetaData) ' does it give errors?
                            SonosPlay()
                            Exit Try
                        End If
                        AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue))
                        TrackIndex = TrackIndex + 1
                    Else
                        ' Go find the track
                        Dim Tracks
                        Dim LoopIndex As Integer = 0
                        Try
                            Tracks = SourceSonosPlayer.GetTracks(SQLreader("Id").ToString, True, True)
                        Catch ex As Exception
                            Tracks = Nothing
                            Log("PlayMusicOnZone unable to put in tracks for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                        Try
                            While LoopIndex < UBound(Tracks, 1)
                                'If g_bDebug Then Log( "Record #" & LoopIndex.ToString & " with value " & Tracks(LoopIndex, 0).ToString)
                                If (Track <> "" And Tracks(LoopIndex, 0).ToString = Track) Or (TrackMatch <> "") Then
                                    If TrackMatch <> "" Then
                                        If InStr(UCase(Tracks(LoopIndex, 0).ToString), UCase(TrackMatch)) <> 0 Then
                                            TrackMetaData = SourceSonosPlayer.GetTrackMetaData(Tracks(LoopIndex, 6).ToString, True) ' this is the "Id"
                                            TrackURI = ZoneUDN & Tracks(LoopIndex, 3).ToString ' this is the "URI"
                                            AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue))
                                            TrackIndex = TrackIndex + 1
                                        End If
                                    Else
                                        TrackMetaData = SourceSonosPlayer.GetTrackMetaData(Tracks(LoopIndex, 6).ToString, True) ' this is the "Id"
                                        TrackURI = ZoneUDN & Tracks(LoopIndex, 3).ToString ' this is the "URI"
                                        AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex + LoopIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex + LoopIndex = CurrentTrackIndexInQueue))
                                        TrackIndex = TrackIndex + 1
                                    End If
                                Else
                                    If (StartWithTrack <> "" And UCase(Tracks(LoopIndex, 0).ToString) = UCase(StartWithTrack)) Then
                                        StartTrackIndex = TrackIndex
                                    End If
                                    TrackMetaData = SourceSonosPlayer.GetTrackMetaData(Tracks(LoopIndex, 6).ToString, True) ' this is the "Id"
                                    TrackURI = ZoneUDN & Tracks(LoopIndex, 3).ToString ' this is the "URI"
                                    AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex + LoopIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex + LoopIndex = CurrentTrackIndexInQueue))
                                    TrackIndex = TrackIndex + 1
                                End If
                                LoopIndex = LoopIndex + 1
                            End While
                        Catch ex As Exception
                            Log("PlayMusicOnZone unable read record  for zoneplayer = " & ZoneName & "with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                        End Try
                    End If
                ElseIf SQLreader("Name").ToString = "All Genres" Then
                    ' this is a Genre. Either get all Genres or If a track was specified, go find it
                    If Track = "" And StartWithTrack = "" And TrackMatch = "" And Artist = "" And Album = "" Then
                        TrackMetaData = SourceSonosPlayer.GetTrackMetaData(SQLreader("Id").ToString, False) ' it's a container
                        If ZoneUDN <> "" Then
                            PlayURI(TrackURI, "") ' TrackMetaData) ' does it give errors?
                            SonosPlay()
                            Exit Try
                        End If
                        AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue))
                        TrackIndex = TrackIndex + 1
                        Exit Try
                    Else

                        ' Go find the tracks

                        Dim Tracks As System.Array
                        Dim LoopIndex As Integer = 0

                        If StartWithTrack <> "" Then Track = StartWithTrack
                        If StartWithArtist <> "" Then Artist = StartWithArtist

                        Try
                            Tracks = SourceSonosPlayer.GetMatchingTracks(Artist, Album, Genre, Track)
                        Catch ex As Exception
                            Tracks = Nothing
                            Log("PlayMusicOnZone unable to put in tracks with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Sub
                        End Try
                        Try
                            LoopIndex = 0
                            While LoopIndex < UBound(Tracks)
                                'If g_bDebug Then Log( "PlayMusicOnZone Track Info (0) = " & Tracks(LoopIndex + 0).ToString)
                                'If g_bDebug Then Log( "PlayMusicOnZone Track Info (1) = " & Tracks(LoopIndex + 1).ToString)
                                'If g_bDebug Then Log( "PlayMusicOnZone Track Info (2) = " & Tracks(LoopIndex + 2).ToString)

                                TrackMetaData = SourceSonosPlayer.GetTrackMetaData(Tracks(LoopIndex + 2).ToString, True) ' this is the "Id"
                                TrackURI = ZoneUDN & Tracks(LoopIndex + 1).ToString ' this is the "URI"
                                AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue)) ' in stead of TrackIndex
                                TrackIndex = TrackIndex + 1
                                LoopIndex = LoopIndex + 4 ' changed from 3 to 4 in v.78
                            End While
                        Catch ex As Exception
                            Log("PlayMusicOnZone unable read record  for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Exit Sub
                        End Try
                    End If
                ElseIf SQLreader("Name").ToString = "All Tracks" And Track = "" And StartWithTrack = "" And TrackMatch = "" And Artist = "" And Album = "" Then
                    ' this must be selected because an artist or an album was entered
                    TrackMetaData = SourceSonosPlayer.GetTrackMetaData(SQLreader("Id").ToString, False) ' The ID here in fact points to the album ID
                    If ZoneUDN <> "" Then
                        PlayURI(TrackURI, "") ' TrackMetaData) ' does it give errors?
                        SonosPlay()
                        Exit Try
                    End If
                    AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue))
                    TrackIndex = TrackIndex + 1
                Else
                    ' this is track info
                    If TrackMatch <> "" Then
                        If InStr(UCase(SQLreader("Name").ToString), UCase(TrackMatch)) <> 0 Then
                            TrackMetaData = SourceSonosPlayer.GetTrackMetaData(SQLreader("Id").ToString, True) ' The ID here in fact points a track
                            If ZoneUDN <> "" Then
                                PlayURI(TrackURI, "") ' TrackMetaData) ' does it give errors?
                                SonosPlay()
                                Exit Try
                            End If
                            AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue)) ' trackindex
                            TrackIndex = TrackIndex + 1
                        End If
                    ElseIf StartWithTrack <> "" Then
                        If SQLreader("Name").ToString <> "All Tracks" Then
                            TrackMetaData = SourceSonosPlayer.GetTrackMetaData(SQLreader("Id").ToString, True) ' The ID here in fact points to the album ID
                            Dim xmlDoc As New XmlDocument
                            Dim TrackTitle As String = ""
                            Try
                                xmlDoc.LoadXml(TrackMetaData)
                                TrackTitle = xmlDoc.GetElementsByTagName("dc:title").Item(0).InnerText
                                If UCase(TrackTitle) = UCase(StartWithTrack) Then
                                    StartTrackIndex = TrackIndex
                                End If
                                Log("PlayMusicOnZone found for zoneplayer = " & ZoneName & " following track = " & xmlDoc.GetElementsByTagName("dc:title").Item(0).InnerText, LogType.LOG_TYPE_INFO)
                            Catch ex As Exception
                                If g_bDebug Then Log("Error in PlayMusicOnZone unable to get track title out of TrackMetaData for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
                            End Try
                            xmlDoc = Nothing
                            If ZoneUDN <> "" Then
                                PlayURI(TrackURI, "") ' TrackMetaData) ' does it give errors?
                                SonosPlay()
                                Exit Try
                            End If
                            AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue))
                            TrackIndex = TrackIndex + 1
                        End If
                    Else
                        TrackMetaData = SourceSonosPlayer.GetTrackMetaData(SQLreader("Id").ToString, True) ' The ID here in fact points to the album ID
                        AddTrackToQueue(TrackURI, TrackMetaData, TrackIndex, QueueAction = QueueActions.qaPlayNext And (TrackIndex = CurrentTrackIndexInQueue))
                        TrackIndex = TrackIndex + 1
                    End If
                End If
            End While
            If ClearPlayerQueue And QueueAction <> QueueActions.qaDontPlay Then
                PlayFromQueue("Q:")
            End If
        Catch ex As Exception
            Log("PlayMusicOnZone unsuccesful reading of record set  for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        If ZoneUDN <> "" Then
            StartTrackIndex = 0
            QueueAction = QueueActions.qaDontPlay
        End If

        If StartTrackIndex <> 0 Then SeekTrack(StartTrackIndex)

        If Not PlayingFromQueue And Not PlayRadiostation And QueueAction <> QueueActions.qaDontPlay Then
            SourceSonosPlayer.SetTransportState("Stop")
            PlayFromQueue("Q:")
        End If

        If QueueAction = QueueActions.qaPlayNow And Not PlayRadiostation Then
            SeekTrack(CurrentTrackIndexInQueue)
        End If

        If QueueAction <> QueueActions.qaDontPlay Then SetTransportState("Play")

        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("PlayMusicOnZone to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub PlayRadioStation(ByVal RadioStation As String)
        If g_bDebug Then Log("PlayRadioStation called for Zone " & ZoneName & " and RadioStation = " & RadioStation, LogType.LOG_TYPE_INFO)
        If Mid(RadioStation, 1, 14) = "RadioStation: " Then
            ' this a a Sonos favorite Radiostation not one learned by the plugin
            PlayMusicOnZone(CurrentAppPath & MusicDBPath, "", Me, "", "", RadioStation, "", "", "", "", "", "", False, QueueActions.qaPlayNow)
            Exit Sub
        End If
        Dim ConnectionString As String
        Dim QueryString As String = "SELECT * FROM RadioStations WHERE Name='" & PrepareForQuery(RadioStation) & "'"

        ConnectionString = "Data Source=" & CurrentAppPath & RadioStationsDBPath
        '
        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("PlayRadioStation unsuccesful query  for zoneplayer = " & ZoneName & "with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("PlayRadioStation - Query string: " & QueryString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim StationURI As String = ""
        Dim StationMetaData As String = ""
        SetTransportState("Stop")
        'ClearQueue() ' why do I do this? Different from v.114
        Try
            If SQLreader.Read() Then
                StationURI = SQLreader("URI").ToString
                StationMetaData = SQLreader("MetaData").ToString
                PlayURI(StationURI, StationMetaData)
                SetTransportState("Play")
            End If
        Catch ex As Exception
            Log("PlayRadioStation unsuccesful reading of record set  for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("PlayRadioStation to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub PlayAudioBookPodCast(ByVal TrackName As String, ByVal iPodPlayerName As String, ByVal ZoneUDN As String, ByVal SourceSonosPlayer As HSPI, ByVal AudioBook As Boolean)
        Dim ConnectionString As String
        Dim QueryString As String = ""
        If AudioBook Then
            QueryString = "SELECT * FROM Tracks WHERE `ParentID`='A:AUDIOBOOK' and Name='" & PrepareForQuery(TrackName) & "'"
        Else
            QueryString = "SELECT * FROM Tracks WHERE `ParentID`='A:PODCAST' and Name='" & PrepareForQuery(TrackName) & "'"
        End If

        ConnectionString = "Data Source=" & CurrentAppPath & DockedPlayersDBPath & iPodPlayerName & ".sdb"

        Dim SQLconnect As SQLiteConnection = Nothing
        Dim SQLCommand As SQLiteCommand = Nothing
        Dim SQLreader As SQLiteDataReader = Nothing
        Try
            'Create a new database connection
            SQLconnect = New SQLiteConnection(ConnectionString)
            'Open the connection
            SQLconnect.Open()
            SQLCommand = SQLconnect.CreateCommand
            SQLCommand.CommandText = QueryString
            SQLreader = SQLCommand.ExecuteReader()
        Catch ex As Exception
            Log("Error PlayAudioBookPodCast unsuccesful query  for zoneplayer = " & ZoneName & "with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("Error PlayAudioBookPodCast - Query string: " & QueryString, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try
        Dim TrackURI As String = ""
        Dim TrackMetaData As String = ""
        SetTransportState("Stop")
        ClearQueue()
        Try
            If SQLreader.Read() Then
                TrackURI = ZoneUDN & SQLreader("URI").ToString
                TrackMetaData = SourceSonosPlayer.GetTrackMetaData(SQLreader("Id").ToString, True) ' The ID here in fact points a track
                PlayURI(TrackURI, TrackMetaData)
                SetTransportState("Play")
                Exit Try
            End If
        Catch ex As Exception
            Log("PlayAudioBookPodCast unsuccesful reading of record set  for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            'Cleanup and close the connection
            If SQLreader IsNot Nothing Then SQLreader = Nothing
            If Not IsNothing(SQLconnect) Then
                SQLconnect.Close()
            End If
            Try
                If SQLCommand IsNot Nothing Then SQLCommand.Dispose()
                If Not IsNothing(SQLconnect) Then SQLconnect.Dispose()
            Catch ex As Exception
            End Try
        Catch ex As Exception
            Log("PlayAudioBookPodCast to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub PlayTV()
        If g_bDebug Then Log("PlayTV called for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Then Exit Sub
        If Not CheckPlayerCanPlayTV(ZoneModel) Then Exit Sub ' changed on 7/12/2019 in v3.1.0.31
        PlayURI("x-sonos-htastream:" & UDN & ":spdif", "", False)
    End Sub

    Private Sub iPodDockChange(ByVal LineConnected As Boolean)
        If MyZoneModel = "WD100" Then
            If CBool(AudioInState(4)) <> LineConnected Then
                'If g_bDebug Then Log( "AudioInCallback for ZonePlayer - " & ZoneName & " LineInConnected has changed")
                If LineConnected Then
                    CurrentPlayerState = player_state_values.Docked
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.Docked)
                    If HSRefDockDeviceName <> -1 Then
                        hs.SetDeviceString(HSRefDockDeviceName, MyDockediPodPlayerName, True)
                    End If
                Else
                    CurrentPlayerState = player_state_values.Undocked
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.Undocked)
                    If HSRefDockDeviceName <> -1 Then
                        hs.SetDeviceString(HSRefDockDeviceName, "", True)
                    End If
                End If
                If Not LineConnected Then
                    ' an event should only be received with AudioInstate to false, when the iPod is undocked
                    MyCurrentTrackInfo(0) = "" ' Artist
                    MyCurrentTrackInfo(1) = "" 'Album
                    MyCurrentTrackInfo(2) = "" 'Title
                    MyCurrentTrackInfo(3) = "" 'CurrentURI
                    MyCurrentTrackInfo(4) = NoArtPath
                    MyCurrentTrackInfo(5) = "0:00:00" 'Track Position
                    MyCurrentTrackInfo(6) = "0:00:00" 'Track Duration
                    MyCurrentTrackInfo(7) = "0" ' Queue Position
                    MyCurrentTrackInfo(8) = "0" ' % played
                    MyCurrentTrackInfo(9) = "Undocked"
                    MyCurrentTrackInfo(10) = "False" 'Internet Radio True/False
                    ZoneSource = "Undocked"
                End If
            End If
            If LineConnected Then
                'ConnectToIPod = True
                FindDockedPlayerSettings()
                AddiPodPlayerNameToINIFile(MyDockediPodPlayerName)
                ConnectWireLessDock("uuid:DOCK" & UDN & "_MS")
            Else
                ContentDirectory = Nothing
                DockedDeviceStatus = "Offline"
                ResetDockedPlayerSettings()
                SetHSPlayerInfo()
                hs.SetDeviceValue(HSRefPlayer, CurrentPlayerState)
                If g_bDebug Then Log("HS updated in AudioInStateChange. HSRef = " & HSRefPlayer.ToString, LogType.LOG_TYPE_INFO)
            End If
        Else
            ' the audio in state changed 
            If CBool(AudioInState(4)) <> LineConnected Then
                'If g_bDebug Then Log( "AudioInCallback for ZonePlayer - " & ZoneName & " LineInConnected has changed")
                If LineConnected Then
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.AudioInTrue)
                Else
                    PlayChangeNotifyCallback(player_status_change.DeviceStatusChanged, player_state_values.AudioInFalse)
                End If
            End If
        End If
    End Sub

    Private Function BuildArtURL(ByVal AlbumArtURL As String) As String ' example /getaa?u=aac://4723.live.streamtheworld.com:3690/KLLCFMAACCMP3&v=269
        'If g_bDebug Then Log( "BuildArtURL. Input = " & strAlbumArtURL)
        ' removed because art work wouldn't show up if spaces or other characters are included in filename    strAlbumArtURL = DecodeURI(strAlbumArtURL)
        If Mid(AlbumArtURL, 1, 9) = "/getaa?u=" Then
            AlbumArtURL = "http://" & IPAddress.ToString & ":" & IPPort.ToString & AlbumArtURL
        ElseIf Mid(AlbumArtURL, 1, 9) = "/getaa?s=" Then ' example /getaa?s=1&amp;u=x-sonos-mms:track:33491428?sid=0&flags=32
            AlbumArtURL = Replace(AlbumArtURL, "&amp;", "&")  ' added v21 on Sept 26/2017 because the &amp; was giving errors on Amazon radio
            '            If Mid(AlbumArtURL, 1, 14) = "getaa?s=1&amp;" Then 
            '           AlbumArtURL = "/getaa?s=1&" & AlbumArtURL.Remove(1, 14)
            '          AlbumArtURL = Replace(AlbumArtURL, "&amp;", "&")
            '     End If
            AlbumArtURL = "http://" & IPAddress & ":" & IPPort.ToString & AlbumArtURL
        ElseIf Mid(AlbumArtURL, 1, 9) = "/getaa?m=" Then ' example /getaa?m=1&u=http%3a%2f%2f79f84a69-0f02-4bd8-b6b0-f72f89ed86f3.x-udn%2fWMPNSSv4%2f33788882%2f1_ezBFQTc3NUU2LTUyNDctNEJDNS05QkVBLTYzMUJBMkZCMDIyRn0uMC5EOUQ4QTBDRQ.mp3%3falbumArt%3dtrue
            AlbumArtURL = "http://" & IPAddress & ":" & IPPort.ToString & AlbumArtURL
        ElseIf Mid(AlbumArtURL, 1, 4) = "http" Then
            'strAlbumArtURL = DecodeURI(strAlbumArtURL)
        ElseIf Mid(AlbumArtURL, 1, 9) = "/getaa?r=" Then ' example from Rhapsody /getaa?r=1&amp;u=radio-radea%3aps.8647980%3aTraS.2058013.mp3
            AlbumArtURL = "http://" & IPAddress & ":" & IPPort.ToString & AlbumArtURL
        Else
            AlbumArtURL = ""
        End If
        BuildArtURL = AlbumArtURL
    End Function

    Public Function PrepareForQuery(ByVal inString As String) As String
        ' this function deals with ' in query names
        PrepareForQuery = inString.Replace("'", "''")
    End Function

    Private Function AddDescToMetaData(ByVal MetaData As String) As String
        ' this function will add an element to the meta data only if the Sonos SW version is > 3.2 and the desc is missing
        ' More info on Radiostations at http://forums.sonos.com/archive/index.php?t-7718.html
        ' Something changed between SW R3.2 and SW R3.3 on Sonos
        ' This is what Browsing the Radio station returns
        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
        '<item id="R:0/0/0" parentID="R:0/0" restricted="false">
        '<dc:title>104.5 | KFOG (AAA)</dc:title>
        '<upnp:class>object.item.audioItem.audioBroadcast</upnp:class>
        '<res protocolInfo="x-rincon-mp3radio:*:*:*">x-sonosapi-stream:s32698?sid=254&amp;flags=32</res></item></DIDL-Lite>

        ' this is what GetMedia Info returns . Using this works in R33
        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
        '<item id="-1" parentID="-1" restricted="true">
        '   <dc:title>KFOG</dc:title>
        '   <upnp:class>object.item.audioItem.audioBroadcast</upnp:class>
        '   <desc id="cdudn" nameSpace="urn:schemas-rinconnetworks-com:metadata-1-0/">SA_RINCON65031_</desc> <++ it is this part that makes it work ??
        '</item></DIDL-Lite>
        If g_bDebug Then Log("AddDescToMetaData called for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        AddDescToMetaData = MetaData
        Dim xmlData As XmlDocument = New XmlDocument
        Try
            xmlData.LoadXml(MetaData)
        Catch ex As Exception
            Log("Error in AddDescToMetaData for zoneplayer = " & ZoneName & " loading XML. XML = " & MetaData & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            If xmlData.GetElementsByTagName("desc").Item(0).InnerText() <> "" Then
                ' the desc element is already present, don't do anything
                Exit Function
            End If
        Catch ex As Exception
            ' this means the element is not present
        End Try

        Dim Position As Integer
        Position = InStr(MetaData, "</item></DIDL-Lite>") ' </item></DIDL-Lite>
        If Position = 0 Then Exit Function
        MetaData = MetaData.Insert(Position - 1, "<desc id=""cdudn"" nameSpace=""urn:schemas-rinconnetworks-com:metadata-1-0/"">SA_RINCON65031_</desc>")
        AddDescToMetaData = MetaData
        If g_bDebug Then Log("AddDescToMetaData called for zoneplayer " & ZoneName & " with modified XML = " & AddDescToMetaData.ToString, LogType.LOG_TYPE_INFO)

    End Function

    Public Function GetTrackMetaData(ByVal ObjectID As String, ByVal BrowseOnlyMetaData As Boolean) As String
        ' More info on Radiostations at http://forums.sonos.com/archive/index.php?t-7718.html
        GetTrackMetaData = ""
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("GetTrackMetaData called  for zoneplayer = " & ZoneName & " with URI= " & ObjectID & " and BrowseOnlyMetaData = " & BrowseOnlyMetaData.ToString, LogType.LOG_TYPE_INFO)
        ObjectID = Trim(ObjectID)
        If ObjectID = "" Then Exit Function
        Dim strMetaData As String
        Dim strSearchObject As String
        Dim NumberReturned As Integer = 0
        Dim RequestedCount As Integer = 10
        strSearchObject = ObjectID
        If Mid(ObjectID, 1, 6) = "R:0/0/" Then
            ' we're looking for a radio station
            strSearchObject = "R:0/0"
            RequestedCount = 0 ' ask for all
        End If
        If strSearchObject <> "" Then
            ' go get the XML
            Dim InArg(5)
            Dim OutArg(3)
            InArg(0) = strSearchObject
            If BrowseOnlyMetaData Then
                InArg(1) = "BrowseMetadata"
            Else
                InArg(1) = "BrowseDirectChildren"
            End If
            InArg(2) = "*"
            InArg(3) = 0
            InArg(4) = RequestedCount
            InArg(5) = ""
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                strMetaData = OutArg(0)
                NumberReturned = OutArg(1)
                If strSearchObject = "R:0/0" Then
                    ' we first need to find the right element
                    Dim xmlData As XmlDocument = New XmlDocument
                    Dim LoopIndex As Integer
                    Try
                        xmlData.LoadXml(strMetaData)
                        For LoopIndex = 0 To NumberReturned - 1
                            Try
                                If xmlData.GetElementsByTagName("item").Item(LoopIndex).Attributes("id").Value = ObjectID Then
                                    ' Ok this is what we are looking for
                                    Try
                                        InArg(3) = LoopIndex
                                        InArg(4) = 1 ' ask for only one entry
                                        ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                                        strMetaData = OutArg(0)
                                        strMetaData = AddDescToMetaData(strMetaData)
                                    Catch ex As Exception
                                        Log("Error in GetTrackMetaData  for zoneplayer = " & ZoneName & " finding XML for index " & LoopIndex.ToString, LogType.LOG_TYPE_ERROR)
                                    End Try
                                    Exit For
                                End If
                            Catch ex As Exception
                            End Try
                        Next
                    Catch ex As Exception
                        Log("Error in GetTrackMetaData  for zoneplayer = " & ZoneName & " loading XML. XML = " & strMetaData, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                End If
                GetTrackMetaData = strMetaData
            Catch ex As Exception
                Log("GetTrackMetaData error in getting XML. ObjectID = " & ObjectID & "  for zoneplayer = " & ZoneName & " and error - " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
    End Function

    Public Function PlayURI(ByVal URI As String, ByVal MetaData As String, Optional isObjectID As Boolean = False) As String
        ' It works only passing the URI and in my case it selects the song but doesn't start playing
        PlayURI = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If SuperDebug Then
            Log("PlayURI called for zoneplayer " & ZoneName & " with strURI = " & URI & " and MetaData = " & MetaData & " and isObjectID = " & isObjectID.ToString, LogType.LOG_TYPE_INFO)
        ElseIf g_bDebug Then
            Log("PlayURI called for zoneplayer " & ZoneName & " with strURI = " & URI & " and isObjectID = " & isObjectID.ToString, LogType.LOG_TYPE_INFO)
        End If

        If isObjectID Then
            MetaData = GetTrackMetaData(MetaData, True)
        ElseIf MetaData <> "" And Mid(MetaData, 1, 1) <> "<" Then
            MetaData = GetTrackMetaData(MetaData, False)
        End If

        If MetaData <> "" Then  ' added to play Favorites FV in version 3.1.0.10
            Try
                Dim MetaDataDoc As New XmlDocument
                MetaDataDoc.LoadXml(MetaData)
                If MetaDataDoc.GetElementsByTagName("r:resMD").Item(0).InnerXml <> "" Then
                    MetaData = WebUtility.HtmlDecode(MetaDataDoc.GetElementsByTagName("r:resMD").Item(0).InnerXml)
                    Try
                        MetaDataDoc.LoadXml(MetaData)
                        If MetaDataDoc.GetElementsByTagName("item").Item(0).Attributes("parentID").Value = "SQ:" Then
                            Dim ObjectID As String = MetaDataDoc.GetElementsByTagName("item").Item(0).Attributes("id").Value
                            If ObjectID <> "" Then
                                MetaData = GetTrackMetaData(ObjectID, False)
                            End If
                        End If
                        Dim UPnPClass As String = ProcessClassInfo(MetaDataDoc.GetElementsByTagName("upnp:class").Item(0).InnerText)
                        If UPnPClass.ToUpper = "MUSICTRACK" Or UPnPClass.ToUpper = "" Or UPnPClass.ToUpper = "MUSICALBUM" Then
                            AddTrackToQueue(URI, MetaData, 0, True)
                            PlayURI("x-rincon-queue:" & GetUDN() & "#0", "")
                            PlayURI = "OK"
                            Exit Function
                        ElseIf UPnPClass.ToUpper = "#PLAYLISTVIEW" Then ' this case was added to support Apple Music Playlists which have class = object.container.playlistContainer.#PlaylistView
                            ' added 5/22/2019 in v3.1.0.30
                            ClearQueue()
                            AddTrackToQueue(URI, MetaData, 0, True)
                            PlayURI("x-rincon-queue:" & GetUDN() & "#0", "")
                            PlayURI = "OK"
                            Exit Function
                        End If
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            End Try
        End If

        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = 0
            InArg(1) = URI
            InArg(2) = MetaData
            AVTransport.InvokeAction("SetAVTransportURI", InArg, OutArg)
            PlayURI = "OK"
        Catch ex As Exception
            Log("ERROR in PlayURI for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". URI=" & URI & " and isObjectID = " & isObjectID.ToString & ", MetaData=" & MetaData & ", Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
            PlayURI = "NOK"
        End Try
    End Function


    Public Function CheckQueueForMMS() As Boolean
        CheckQueueForMMS = False
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("CheckQueueForMMS called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim LoopIndex As Integer
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim TrackURI As String = ""

        Try

            Dim InArg(5)
            Dim OutArg(3)

            InArg(0) = "Q:0"
            InArg(1) = "BrowseDirectChildren"
            InArg(2) = "*"
            InArg(3) = "0"
            InArg(4) = MaxNbrOfUPNPObjects
            InArg(5) = ""

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in CheckQueueForMMS for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try

            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)

            If g_bDebug Then Log("CheckQueueForMMS found " & TotalMatches.ToString & " queue entries for ZonePlayer - " & ZoneName & " and NbrRet= " & NumberReturned.ToString, LogType.LOG_TYPE_INFO)
            If NumberReturned < 1 Then
                Exit Function
            End If
            '
            StartIndex = 0
            '
            Do
                InArg(3) = CStr(StartIndex)
                InArg(4) = MaxNbrOfUPNPObjects
                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in CheckQueueForMMS/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try
                Try
                    xmlData.LoadXml(OutArg(0))
                    NumberReturned = OutArg(1)
                    For LoopIndex = 0 To NumberReturned - 1
                        Try
                            TrackURI = xmlData.GetElementsByTagName("res").Item(LoopIndex).InnerText
                            If InStr(TrackURI, "x-sonos-mms") <> 0 Then
                                CheckQueueForMMS = True
                                If g_bDebug Then Log("CheckQueueForMMS found x-sonos-mms for zoneplayer = " & ZoneName & " in TrackURI " & TrackURI, LogType.LOG_TYPE_INFO)
                                Exit Function
                            End If
                        Catch ex As Exception
                        End Try
                    Next
                Catch ex As Exception
                    Log("Error in CheckQueueForMMS loading XML. XML = " & OutArg(0), LogType.LOG_TYPE_INFO)
                    Exit Function
                End Try
                StartIndex = StartIndex + NumberReturned
                If StartIndex >= TotalMatches Then
                    Exit Do
                End If
                If g_bDebug Then Log("CheckQueueForMMS for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
                'hs.WaitEvents()
            Loop
        Catch ex As Exception
            Log("Error in CheckQueueForMMS for zoneplayer = " & ZoneName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetQueue() As String() ' this will return an array of titles
        GetQueue = Nothing
        If DeviceStatus = "Offline" Then Exit Function
        If ZoneModel = "WD100" Then Exit Function
        'If g_bDebug Then Log("GetQueue called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO) ' DCORMEDIAAPI
        Dim xmlData As XmlDocument = New XmlDocument
        Dim Queue As String() = Nothing
        Dim I As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim ArrayIndex As Integer = 0
        Dim MetaData As String = ""

        Try

            Dim InArg(5) As String
            Dim OutArg(3) As String

            InArg(0) = "Q:0"
            InArg(1) = "BrowseDirectChildren"
            InArg(2) = "dc:title"
            InArg(3) = "0"
            InArg(4) = MaxNbrOfUPNPObjects.ToString
            InArg(5) = ""

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in GetQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            MetaData = OutArg(0)
            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)
            ReDim Queue(TotalMatches - 1) '  sizes the queue for maximum read

            'If g_bDebug Then Log("GetQueue found " & TotalMatches.ToString & " queue entries for ZonePlayer - " & ZoneName & " and NbrRet= " & NumberReturned.ToString, LogType.LOG_TYPE_INFO)
            If NumberReturned < 1 Then
                GetQueue = Nothing
                Queue = Nothing
                Exit Function
            End If
            '
            StartIndex = 0
            ArrayIndex = 0
            '
            Do
                Try
                    xmlData.LoadXml(MetaData)
                Catch ex As Exception
                    Log("ERROR in GetQueue for zoneplayer = " & ZoneName & " at Index = " & StartIndex.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    GetQueue = Queue
                    Exit Function
                End Try
                Try
                    Dim TitleList As XmlNodeList = xmlData.GetElementsByTagName("dc:title")
                    For Each Title As XmlNode In TitleList
                        Queue(ArrayIndex) = Title.InnerText
                        ArrayIndex += 1
                    Next
                Catch ex As Exception
                    Log("ERROR in GetQueue for zoneplayer = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
                StartIndex += NumberReturned
                If StartIndex >= TotalMatches Then Exit Do
                InArg(3) = StartIndex.ToString
                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in GetQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                MetaData = OutArg(0)
                NumberReturned = OutArg(1)
                If NumberReturned < 1 Then Exit Do
                'If g_bDebug Then Log("GetQueue found " & TotalMatches.ToString & " queue entries for ZonePlayer - " & ZoneName & " at Index = " & StartIndex.ToString & " and NbrRet= " & NumberReturned.ToString, LogType.LOG_TYPE_INFO)
            Loop
        Catch ex As Exception
            Log("Error in GetQueue with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Queue = Nothing
            Exit Function
        End Try
        'If g_bDebug Then Log("GetQueue for ZonePlayer - " & ZoneName & " found " & Queue.Count.ToString & " entries", LogType.LOG_TYPE_INFO) ' DCORMEDIAPI
        GetQueue = Queue
        Queue = Nothing

    End Function

    Public Function GetQueueElement(ElementID As Integer, ElementTitle As String) As HomeSeerAPI.Lib_Entry
        GetQueueElement = Nothing
        If DeviceStatus = "Offline" Then Exit Function
        If ZoneModel = "WD100" Then Exit Function
        If SuperDebug Or DCORMEDIAAPITrace Then Log("GetQueueElement called for ZonePlayer = " & ZoneName & ", ElementID = " & ElementID.ToString & ", ElementTitle = " & ElementTitle, LogType.LOG_TYPE_INFO)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim I As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim ArrayIndex As Integer = 0
        Dim MetaData As String = ""

        Dim Result As HomeSeerAPI.Lib_Entry
        Result.Album = ""
        Result.Artist = ""
        Result.Cover_Back_path = NoArtPath
        Result.Cover_path = NoArtPath
        Result.Genre = ""
        Result.Key.iKey = 1
        Result.Key.Library = MyLibraryTypes.LibraryQueue
        Result.Key.sKey = ""
        Result.Key.Title = ElementTitle
        Result.Key.WhichKey = eKey_Type.eEither
        Result.Kind = ""
        Result.LengthSeconds = 0
        Result.Lib_Media_Type = eLib_Media_Type.Music
        Result.Lib_Type = 0
        Result.PlayedCount = 0
        Result.Rating = 0
        Result.Title = ""
        Result.Year = 0

        Try

            Dim InArg(5)
            Dim OutArg(3)

            InArg(0) = "Q:0"
            InArg(1) = "BrowseDirectChildren"
            InArg(2) = "*"
            If ElementID <> 0 Then
                InArg(3) = (ElementID - 1).ToString
                InArg(4) = "1"
            Else
                InArg(3) = "0"
                InArg(4) = MaxNbrOfUPNPObjects.ToString
            End If
            InArg(5) = ""

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in GetQueueElement for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            MetaData = OutArg(0)
            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)

            If SuperDebug Or DCORMEDIAAPITrace Then Log("GetQueueElement found " & TotalMatches.ToString & " queue entries for ZonePlayer - " & ZoneName & " and NbrRet= " & NumberReturned.ToString, LogType.LOG_TYPE_INFO)
            If NumberReturned < 1 Then
                GetQueueElement = Nothing
                Exit Function
            End If



            If ElementID = 0 Then
                ' we are searching by tittle
                '
                StartIndex = 0
                ArrayIndex = 0
                '
                Do
                    Try
                        xmlData.LoadXml(MetaData)
                    Catch ex As Exception
                        Log("ERROR in GetQueueElement for zoneplayer = " & ZoneName & " at Index = " & StartIndex.ToString & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                    Try
                        Dim TrackNodes As XmlNodeList = xmlData.GetElementsByTagName("item")
                        Dim tracktitle As String = ""
                        Dim Index As Integer = 0
                        If TrackNodes IsNot Nothing Then
                            If TrackNodes.Count > 0 Then
                                For Each TrackDataNode As XmlNode In TrackNodes
                                    If TrackDataNode.Item("dc:title").InnerText = ElementTitle Then
                                        MetaData = TrackDataNode.OuterXml
                                        Exit Do
                                    End If
                                Next
                            End If
                        End If
                    Catch ex As Exception
                        Log("ERROR in GetQueueElement for zoneplayer = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    StartIndex += NumberReturned
                    If StartIndex >= TotalMatches Then Exit Do
                    InArg(3) = StartIndex.ToString
                    Try
                        ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                    Catch ex As Exception
                        Log("ERROR in GetQueueElement for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Exit Function
                    End Try
                    MetaData = OutArg(0)
                    NumberReturned = OutArg(1)
                    If NumberReturned < 1 Then Exit Do
                    If SuperDebug Or DCORMEDIAAPITrace Then Log("GetQueueElement found " & TotalMatches.ToString & " queue entries for ZonePlayer - " & ZoneName & " at Index = " & StartIndex.ToString & " and NbrRet= " & NumberReturned.ToString, LogType.LOG_TYPE_INFO)
                Loop
            End If
            ' we're searching by index ID
            xmlData.LoadXml(MetaData)
            Try
                Result.Title = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                Result.Key.Title = Result.Title
            Catch ex As Exception
            End Try
            Try
                Result.Album = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
            Catch ex As Exception
            End Try
            Try
                Result.Artist = xmlData.GetElementsByTagName("dc:creator").Item(0).InnerText
            Catch ex As Exception
            End Try
            Try
                'Result.Cover_path = BuildArtURL(DecodeURI(xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText))
                Result.Cover_path = BuildArtURL(xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText) ' removed in R3.1.0.19 due failure to retrieve art
            Catch ex As Exception
            End Try

        Catch ex As Exception
            Log("Error in GetQueueElement with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If SuperDebug Or DCORMEDIAAPITrace Then Log("GetQueueElement called for ZonePlayer = " & ZoneName & ", ElementID = " & ElementID.ToString & ", ElementTitle = " & ElementTitle & ", Title = " & Result.Title & ", Album = " & Result.Album & ", Artist = " & Result.Artist & ", CoverPath = " & Result.Cover_path, LogType.LOG_TYPE_INFO)
        Return Result

    End Function



    Public Function PlayFromQueue(ByVal QueueId As String)
        PlayFromQueue = ""
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.PlayFromQueue(QueueId)
            Catch ex As Exception
                If g_bDebug Then Log("PlayFromQueue called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Then Exit Function
        If ZoneModel = "WD100" Then Exit Function
        Dim xmlData As XmlDocument = New XmlDocument
        Dim Genre(6)
        Dim InArg(5)
        Dim OutArg(3)

        If g_bDebug Then Log("PlayFromQueue called for ZonePlayer = " & ZoneName & " and QueueId = " & QueueId, LogType.LOG_TYPE_INFO)

        InArg(0) = QueueId                  ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 0                        ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in PlayFromQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            PlayFromQueue = "NOK"
            Exit Function
        End Try

        Try
            xmlData.LoadXml(OutArg(0))
        Catch ex As Exception
            Log("PlayFromQueue unpack the XML data for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            PlayFromQueue = "NOK"
            Exit Function
        End Try

        Dim QueueInfo As String

        Try
            QueueInfo = xmlData.GetElementsByTagName("res").Item(0).InnerText
            PlayFromQueue = PlayURI(QueueInfo, "")
        Catch ex As Exception
            Log("PlayFromQueue unable to find protocol info in XML data for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("PlayFromQueue XML Data = " & OutArg(0).ToString, LogType.LOG_TYPE_ERROR)
            PlayFromQueue = "NOK"
            Exit Function
        End Try

    End Function

    Public Function AddTrackToQueue(ByVal URI As String, ByVal MetaData As String, ByVal QueuePosition As Integer, ByVal EnqueuedNext As Boolean)
        AddTrackToQueue = ""
        If g_bDebug Then Log("AddTrackToQueue called for zoneplayer " & ZoneName & " with URI = " & URI & " MetaData=" & MetaData & " QueuePosition=" & QueuePosition.ToString & " EnqueuedNext=" & EnqueuedNext.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        URI = Trim(URI)
        Try
            Dim InArg(4)
            Dim OutArg(2)
            InArg(0) = 0
            InArg(1) = URI              ' EnqueuedURI String
            InArg(2) = MetaData         ' enqueuedURIMetaData String
            InArg(3) = QueuePosition    ' desiredFirstTrackNumberEnqueued UI4
            InArg(4) = EnqueuedNext     ' enqueuedAsNext Boolean
            AVTransport.InvokeAction("AddURIToQueue", InArg, OutArg)
            MyQueueHasChanged = True
            AddTrackToQueue = "OK"
        Catch ex As Exception
            Log("ERROR in AddTrackToQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("AddTrackToQueue called with URI = " & URI & " MetaData=" & MetaData & " QueuePosition=" & QueuePosition.ToString & " EnqueuedNext=" & EnqueuedNext.ToString, LogType.LOG_TYPE_ERROR)
            AddTrackToQueue = "NOK"
        End Try
    End Function

    Public Function ClearQueue()
        ClearQueue = ""
        If g_bDebug Then Log("ClearQueue called for zoneplayer " & ZoneName, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked Then
            Dim LinkedZone As HSPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.ClearQueue()
            Catch ex As Exception
                If g_bDebug Then Log("ClearQueue called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If ZoneModel = "WD100" Then Exit Function
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = 0
            AVTransport.InvokeAction("RemoveAllTracksFromQueue", InArg, OutArg)
            MyQueueHasChanged = True
            ClearQueue = "OK"
        Catch ex As Exception
            If g_bDebug Then Log("Warning in ClearQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
        End Try
    End Function

    Public Function SaveQueue(ByVal QueueName As String)
        SaveQueue = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If g_bDebug Then Log("SaveQueue called with QueueName = " & QueueName & " for Zone = " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = "0"
            InArg(1) = QueueName
            InArg(2) = ""
            AVTransport.InvokeAction("SaveQueue", InArg, OutArg)
            SaveQueue = OutArg(0) ' this is the assigned queue name
        Catch ex As Exception
            Log("ERROR in SaveQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function TracksInQueue(QueueID As String) As Integer ' this will return the number of tracks in the queue
        TracksInQueue = 0
        If DeviceStatus = "Offline" Then Exit Function
        If ZoneModel = "WD100" Then Exit Function

        Dim InArg(5)
        Dim OutArg(3)

        InArg(0) = "Q:" & QueueID
        InArg(1) = "BrowseDirectChildren"
        InArg(2) = "*"
        InArg(3) = "0"
        InArg(4) = 1
        InArg(5) = ""

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in TracksInQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            TracksInQueue = CInt(OutArg(2))
            If g_bDebug Then Log("TracksInQueue called for ZonePlayer = " & ZoneName & " and found " & OutArg(2).ToString & " tracks.", LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("ERROR in TracksInQueue for zoneplayer = " & ZoneName & " with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Function


    Public Function BrowseQueue(QueueID As String, ByRef UpdateID As Integer) As String
        BrowseQueue = ""
        If g_bDebug Then Log("BrowseQueue called for zoneplayer " & ZoneName, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Or QueueService Is Nothing Then Exit Function
        Try
            Dim InArg(2)
            Dim OutArg(3)
            InArg(0) = QueueID
            InArg(1) = "0"  ' StartingIndex
            InArg(2) = "0"  ' RequestedCount 0 means all
            ' Output is Result, NumberReturned, TotalMatches, UpdateID
            QueueService.InvokeAction("Browse", InArg, OutArg)
            If g_bDebug Then Log("BrowseQueue succesfully called for zoneplayer " & ZoneName & " with Result = " & OutArg(0).ToString & " with NumberReturned = " & OutArg(1).ToString & " with TotalMatched = " & OutArg(2).ToString & " NUpdateId=" & OutArg(3).ToString, LogType.LOG_TYPE_INFO)
            BrowseQueue = OutArg(0)
            UpdateID = OutArg(3)
        Catch ex As Exception
            If g_bDebug Then Log("Warning in BrowseQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
        End Try
    End Function


    Public Function AddURI(QueueID As Integer, UpdateID As Integer, ByVal EnqueuedURI As String, ByVal EnqueuedMetaData As String, ByVal DesiredFirstTrackNumber As Integer, ByVal EnqueuedAsNext As Boolean) As Integer
        AddURI = 0
        If g_bDebug Then Log("AddURI called for zoneplayer " & ZoneName & " with QueueID = " & QueueID & " with UpdateID = " & UpdateID & " with EnqueuedURI = " & EnqueuedURI & " EnqueuedMetaData=" & EnqueuedMetaData & " DesiredFirstTrackNumber=" & DesiredFirstTrackNumber.ToString & " EnqueuedAsNext=" & EnqueuedAsNext.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Or QueueService Is Nothing Then Exit Function
        Try
            Dim InArg(5)
            Dim OutArg(3)
            InArg(0) = QueueID
            InArg(1) = UpdateID
            InArg(2) = EnqueuedURI
            InArg(3) = EnqueuedMetaData
            InArg(4) = DesiredFirstTrackNumber
            InArg(5) = EnqueuedAsNext
            QueueService.InvokeAction("AddURI", InArg, OutArg)
            If g_bDebug Then Log("AddURI succesfully called for zoneplayer " & ZoneName & " with FirstTrackNumberEnqueued = " & OutArg(0).ToString & " with NumTracksAdded = " & OutArg(1).ToString & " with NewQueueLength = " & OutArg(2).ToString & " NewUpdateId=" & OutArg(3).ToString, LogType.LOG_TYPE_INFO)
            MyQueueHasChanged = True
            AddURI = OutArg(3)
        Catch ex As Exception
            Log("ERROR in AddURI for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("AddURI called for zoneplayer " & ZoneName & " with QueueID = " & QueueID & " with UpdateID = " & UpdateID & " with EnqueuedURI = " & EnqueuedURI & " EnqueuedMetaData=" & EnqueuedMetaData & " DesiredFirstTrackNumber=" & DesiredFirstTrackNumber.ToString & " EnqueuedAsNext=" & EnqueuedAsNext.ToString, LogType.LOG_TYPE_ERROR)
            AddURI = 0
        End Try
    End Function

    Public Function CreateQueue(QueueOwnerID As String, QueueOwnerContext As String, ByVal QueuePolicy As String) As Integer
        CreateQueue = 0
        If g_bDebug Then Log("CreatQueue called for zoneplayer " & ZoneName & " with QueueOwnerID = " & QueueOwnerID & " with QueueOwnerContext = " & QueueOwnerContext & " with QueuePolicy = " & QueuePolicy, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Or QueueService Is Nothing Then Exit Function
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = QueueOwnerID
            InArg(1) = QueueOwnerContext
            InArg(2) = QueuePolicy
            QueueService.InvokeAction("CreateQueue", InArg, OutArg)
            If g_bDebug Then Log("CreateQueue succesfully called for zoneplayer " & ZoneName & " with QueueID = " & OutArg(0).ToString, LogType.LOG_TYPE_INFO)
            MyQueueHasChanged = True
            CreateQueue = OutArg(0)
        Catch ex As Exception
            Log("ERROR in CreateQueue for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("CreatQueue called for zoneplayer " & ZoneName & " with QueueOwnerID = " & QueueOwnerID & " with QueueOwnerContext = " & QueueOwnerContext & " with QueuePolicy = " & QueuePolicy, LogType.LOG_TYPE_ERROR)
            CreateQueue = 0
        End Try
    End Function


    Public Function DestroySonosObject(ByVal ObjectID As String) As String
        DestroySonosObject = ""
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("DestroySonosObject called with ObjectID = " & ObjectID & " for Zone = " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = ObjectID
            ContentDirectory.InvokeAction("DestroyObject", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in DestroySonosObject for zoneplayer = " & ZoneName & " with ObjectID = " & ObjectID & " and UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            DestroySonosObject = "NOK"
        End Try
    End Function

    Public Function GetCurrentTrackInfo()
        'TrackInfo(0) = Artist
        'TrackInfo(1) = Album
        'TrackInfo(2) = Title
        'TrackInfo(3) = CurrentURI
        'TrackInfo(4) = AlbumArtURI
        'TrackInfo(5) = Track Position
        'TrackInfo(6) = Track Duration
        'TrackInfo(7) = Queue Position
        'TrackInfo(8) = % played
        'TrackInfo(9) = Source
        'TrackInfo(10)= Internet Radio True/False
        'TrackInfo(11)= TrackInfo(0) and equal to radiostation name or Service
        'TrackInfo(12)= CurrentURIMetaData
        'TrackInfo(13)= RadioCurrentURI
        'TrackInfo(14)= RadioCurrentMetaData
        'TrackInfo(15)= StreamContent added v107
        'TrackInfo(16)= RadioShow added v114
        Dim TrackInfo(15)
        Dim I As Integer
        Dim InArg(0)
        Dim MediaInfo(8)
        Dim PositionInfo(7) ' 0-Track, 1-TrackDuration, 2-TrackMetaData, 3-TrackURI, 4-RelTime, 5-AbsTime, 6-RelCount, 7-AbsCount
        InArg = {""}
        MediaInfo = {0, "", "", "", "", "", "", "", ""}
        PositionInfo = {0, "", "", "", "", 0, 0, 0}
        TrackInfo = {"", "", "", "", NoArtPath, "", "", "0", "", "", "False", "", "", "", "", "", ""}
        GetCurrentTrackInfo = TrackInfo
        If DeviceStatus = "Offline" Or ZoneModel = "WD100" Or AVTransport Is Nothing Then Exit Function

        Dim InternetRadio As Boolean = False
        Dim InputInfo() As String
        Dim xmlData As XmlDocument = New XmlDocument
        Dim CurrentURI, CurrentURIMetaData As String

        If g_bDebug Then Log("GetcurrentTrackInfo called for zoneplayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        InArg(0) = 0 ' InstanceID = 0
        Try
            AVTransport.InvokeAction("GetMediaInfo", InArg, MediaInfo)
        Catch ex As Exception
            Log("ERROR in GetCurrentTrackInfo/GetMediaInfo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyCurrentTrackInfo = TrackInfo
            Exit Function
        End Try
        CurrentURI = MediaInfo(gmiCurrentURI)
        CurrentURIMetaData = MediaInfo(gmiCurrentURIMetaData)
        If Trim(CurrentURIMetaData) <> "" Then MyCurrentURIMetaData = CurrentURIMetaData

        TrackInfo(12) = CurrentURIMetaData
        If SuperDebug Then
            Log(ZoneName & " TrackInfo (CurrentURI) = " & CurrentURI.ToString & " and TrackInfo (CurrentURIMetaData) = " & TrackInfo(12).ToString, LogType.LOG_TYPE_INFO)
        ElseIf g_bDebug Then
            Log(ZoneName & " TrackInfo (CurrentURI) = " & CurrentURI.ToString, LogType.LOG_TYPE_INFO)
        End If

        ' I should check here on arg OutArg1(2) which represents CurrentURI. If x-sonosapi-stream: then we have internet radio
        Try
            If MediaInfo(gmiCurrentURIMetaData) <> "" Then
                Try
                    xmlData.LoadXml(MediaInfo(gmiCurrentURIMetaData)) ' CurrentURIMetaData = 3
                Catch ex As Exception
                    If g_bDebug Then Log("Error in GetCurrentTrackInfo for zoneplayer " & ZoneName & " loading XML. XML = " & MediaInfo(3), LogType.LOG_TYPE_ERROR)
                    MyCurrentTrackInfo = TrackInfo
                    Exit Function
                End Try
                Try
                    TrackInfo(11) = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText ' this is typically the name of the radio station or service
                Catch ex As Exception
                    TrackInfo(11) = ""
                End Try
                If g_bDebug Then Log(ZoneName & " TrackInfo (MediaInfo/Title) = " & TrackInfo(11).ToString, LogType.LOG_TYPE_INFO)

                ' this is for our own announcements
                Try
                    TrackInfo(2) = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                    If g_bDebug Then Log(ZoneName & " TrackInfo (title) = " & TrackInfo(2).ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    TrackInfo(0) = xmlData.GetElementsByTagName("dc:creator").Item(0).InnerText
                    If g_bDebug Then Log(ZoneName & " TrackInfo (creator) = " & TrackInfo(0).ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    TrackInfo(1) = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
                    If g_bDebug Then Log(ZoneName & " TrackInfo (album) = " & TrackInfo(1).ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try
                Try
                    'TrackInfo(4) = DecodeURI(xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText) removed in version 3.1.0.19 because failure to retrieve art
                    TrackInfo(4) = BuildArtURL(TrackInfo(4))
                    If g_bDebug Then Log(ZoneName & " TrackInfo (albumArtURI) = " & TrackInfo(4).ToString, LogType.LOG_TYPE_INFO)
                Catch ex As Exception
                End Try

            End If
            TrackInfo(3) = MediaInfo(gmiCurrentURI) ' CurrentURI = 2
        Catch ex As Exception
            Log("Error in GetCurrentTrackInfo. ZonePlayer = " & ZoneName & " Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
            For I = 1 To UBound(MediaInfo)
                Log("MediaInfo(" & I.ToString & ")= " & MediaInfo(I).ToString, LogType.LOG_TYPE_ERROR)
            Next
        End Try

        Try
            AVTransport.InvokeAction("GetPositionInfo", InArg, PositionInfo)
        Catch ex As Exception
            Log("ERROR in GetCurrentTrackInfo/GetPositionInfo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            MyCurrentTrackInfo = TrackInfo
            Exit Function
        End Try

        If Mid(CurrentURI, 1, 9) = "x-rincon:" Then
            ' CurrentURI - you get this when zone are linked
            TrackInfo(9) = "Linked"
            ZoneSource = "Linked"
            MyZoneSourceExt = "Linked"
            If g_bDebug Then Log(ZoneName & " TrackInfo (CurrentURI) = " & DecodeURI(TrackInfo(3).ToString), LogType.LOG_TYPE_INFO) ' I believe the CurrentURI points to the UDN of the source player
        ElseIf Mid(CurrentURI, 1, 16) = "x-rincon-stream:" Then
            ' seen this when I source from another zoneplayers input line or own
            ' Meta Data Present
            TrackInfo(9) = "Input"
            ZoneSource = "Input"
            MyZoneSourceExt = "Input"
            If g_bDebug Then Log(ZoneName & " TrackInfo (Streaming) = " & TrackInfo(3).ToString, LogType.LOG_TYPE_INFO)
            ' this is received and separated by : as follows "input name : zoneplayer name"
            If g_bDebug Then Log(ZoneName & " TrackInfo (Streaming Source) = " & TrackInfo(0).ToString, LogType.LOG_TYPE_INFO)
            InputInfo = Split(TrackInfo(0), ": ")
        ElseIf Mid(CurrentURI, 1, 17) = "x-rincon-mp3radio" Then
            ' Internet Radio
            TrackInfo(9) = "Radio"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Radio"
            InternetRadio = True
            TrackInfo(4) = BuildArtURL("/getaa?s=1&u=" & TrackInfo(3))
            If g_bDebug Then Log(ZoneName & " TrackInfo () = Internet Radio " & TrackInfo(3).ToString, LogType.LOG_TYPE_INFO)
        ElseIf Mid(CurrentURI, 1, 17) = "x-sonosapi-stream" Then
            ' Internet Radio - MetaData present
            TrackInfo(9) = "Stream Radio"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Stream Radio"
            InternetRadio = True
            TrackInfo(4) = BuildArtURL("/getaa?s=1&u=" & TrackInfo(3))
            If g_bDebug Then Log(ZoneName & " TrackInfo (Internet Radio) = " & TrackInfo(0).ToString, LogType.LOG_TYPE_INFO)
            If g_bDebug Then Log(ZoneName & " TrackInfo (Internet Radio Art) = " & TrackInfo(3).ToString, LogType.LOG_TYPE_INFO)
            'If TrackInfo(0) <> "" And TrackInfo(0).ToString.Contains("iHeartRadio ") Then
            ' this is iHeartRadio, store the info
            If TrackInfo(11) <> "" Then   ' Changed in v101 to learn any type of radiostation
                If LearnRadioStations Then UpdateRadioStationsInfo(CurrentURI, CurrentURIMetaData, "Learned")
            End If
        ElseIf Mid(CurrentURI, 1, 7) = "lastfm:" Then
            ' Internet Radio LastFm - MetaData present
            TrackInfo(9) = "Lastfm"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Lastfm"
            InternetRadio = True
            If LearnRadioStations Then UpdateRadioStationsInfo(CurrentURI, CurrentURIMetaData, "LastFM")
        ElseIf Mid(CurrentURI, 1, 10) = "pndrradio:" Then
            ' Internet Radio Pandora Radio - MetaData present
            TrackInfo(9) = "Pandora"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Pandora"
            InternetRadio = True
            If LearnRadioStations Then UpdateRadioStationsInfo(CurrentURI, CurrentURIMetaData, "Pandora")
        ElseIf Mid(CurrentURI, 1, 15) = "x-rincon-queue:" Then
            ' Normal tracks - Meta Data not present ' there is a queue # at the end
            TrackInfo(9) = "Tracks"
            If g_bDebug Then Log(ZoneName & " TrackInfo (CurrentURI) = " & DecodeURI(TrackInfo(3).ToString), LogType.LOG_TYPE_INFO)
            ZoneSource = "Tracks"
            MyZoneSourceExt = "Tracks"
        ElseIf Mid(CurrentURI, 1, 5) = "http:" Then
            ' This is when I play announcements directly from a PC file pointed to by an IP address
            TrackInfo(9) = "Tracks"
            If g_bDebug Then Log(ZoneName & " TrackInfo (CurrentURI) = " & DecodeURI(TrackInfo(3).ToString), LogType.LOG_TYPE_INFO)
            ZoneSource = "Tracks"
            MyZoneSourceExt = "Tracks"
        ElseIf Mid(CurrentURI, 1, 9) = "sirradio:" Then
            ' Internet Radio Pandora Sirius - MetaData present
            TrackInfo(9) = "Sirius"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Sirius"
            InternetRadio = True
            If LearnRadioStations Then UpdateRadioStationsInfo(CurrentURI, CurrentURIMetaData, "Sirius")
        ElseIf Mid(CurrentURI, 1, 8) = "rdradio:" Then
            ' Internet Radio Rhapsody Sirius - MetaData present
            TrackInfo(9) = "Rhapsody"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Rhapsody"
            InternetRadio = True
            If LearnRadioStations Then UpdateRadioStationsInfo(CurrentURI, CurrentURIMetaData, "Rhapsody")
        ElseIf Mid(CurrentURI, 1, 13) = "x-sonos-dock:" Then
            ' seen this when I source from a wireless dock
            TrackInfo(9) = "Linked"
            ZoneSource = "Docked"
            MyZoneSourceExt = "Linked"
            If g_bDebug Then Log(ZoneName & " TrackInfo (CurrentURI) = " & DecodeURI(TrackInfo(3).ToString), LogType.LOG_TYPE_INFO) ' I believe the CurrentURI points to the UDN of the source player
        ElseIf Mid(CurrentURI, 1, 16) = "x-rincon-buzzer:" Then
            ' alarm went off and chime is now playing
            TrackInfo(9) = "Chime"
            If g_bDebug Then Log(ZoneName & " TrackInfo (CurrentURI) = " & DecodeURI(TrackInfo(3).ToString), LogType.LOG_TYPE_INFO)
            ZoneSource = "Input"
            MyZoneSourceExt = "Chime"
        ElseIf Mid(CurrentURI, 1, 17) = "x-sonosapi-radio:" Then
            ' Internet Radio
            TrackInfo(9) = "Radio"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Radio"
            InternetRadio = True
            'TrackInfo(4) = BuildArtURL("/getaa?s=1&u=" & TrackInfo(3))
            If LearnRadioStations Then UpdateRadioStationsInfo(CurrentURI, CurrentURIMetaData, "Learned")
            If g_bDebug Then Log(ZoneName & " TrackInfo () = Internet Radio " & TrackInfo(3).ToString, LogType.LOG_TYPE_INFO)
        ElseIf Mid(CurrentURI, 1, 15) = "x-sonosapi-hls:" Then ' added v114
            ' Internet Radio like SiriusXM?
            TrackInfo(9) = "SiriusXM"
            ZoneSource = "Radio"
            MyZoneSourceExt = "SiriusXM"
            InternetRadio = True
            If LearnRadioStations Then UpdateRadioStationsInfo(CurrentURI, CurrentURIMetaData, "SiriusXM")
            If g_bDebug Then Log(ZoneName & " TrackInfo () = Internet SiriusXM " & TrackInfo(2).ToString, LogType.LOG_TYPE_INFO)
        ElseIf Mid(CurrentURI, 1, 18) = "x-sonos-htastream:" Then ' added v3.1.0.3
            ' Playbar sourcing from spdif interface?
            TrackInfo(9) = "TV"
            ZoneSource = "TV"
            MyZoneSourceExt = "TV"
            If g_bDebug Then Log(ZoneName & " TrackInfo () = TV " & TrackInfo(2).ToString, LogType.LOG_TYPE_INFO)
            ' ElseIF .... x-rincon-cpcontainer probably need to add this , this is Google Music podcasts
        Else
            ' Probably some Internet streaming. Save URI and MetaData
            TrackInfo(9) = "Unknown"
            ZoneSource = "Radio"
            MyZoneSourceExt = "Unknown"
            InternetRadio = True
        End If
        If g_bDebug Then Log(ZoneName & " TrackInfo(Source) = " & MyZoneSourceExt, LogType.LOG_TYPE_INFO)

        If MyZoneSourceExt = "Linked" Then
            HandleLinkedZones(TrackInfo(3))
        Else
            If MyZoneIsLinked Then HandleUnlinkZone()
        End If
        If InternetRadio Then
            TrackInfo(10) = "True"
        Else
            TrackInfo(11) = "" ' reset any RadioStationName that we have mistakenly picked up
        End If

        If PositionInfo(pmiTrackMetaData).ToString = "" Or PositionInfo(pmiTrackMetaData).ToString = "NOT_IMPLEMENTED" Then
            GetCurrentTrackInfo = TrackInfo
            MyCurrentTrackInfo = TrackInfo
            Exit Function ' nothing available when in input mode
        End If
        If SuperDebug Then Log("GetcurrentTrackInfo called GetPositionInfo for zoneplayer - " & ZoneName & " and MetaData " & PositionInfo(pmiTrackMetaData).ToString, LogType.LOG_TYPE_INFO)
        Try
            xmlData.LoadXml(PositionInfo(pmiTrackMetaData).ToString)   ' TrackMetaData
        Catch ex As Exception
            Log("Error in GetCurrentTrackInfo for zoneplayer " & ZoneName & " PlayMusicOnZone loading XML trackmetadata. XML = " & PositionInfo(pmiTrackMetaData).ToString, LogType.LOG_TYPE_ERROR)
            For I = 1 To UBound(MediaInfo)
                Log("MediaInfo(" & I.ToString & ")= " & MediaInfo(I).ToString, LogType.LOG_TYPE_ERROR)
            Next
            For I = 1 To UBound(PositionInfo)
                Log("PositionInfo(" & I.ToString & ")= " & PositionInfo(I).ToString, LogType.LOG_TYPE_ERROR)
            Next
            MyCurrentTrackInfo = TrackInfo
            Exit Function
        End Try

        MyCurrentTrackInfo = TrackInfo

        Try
            TrackInfo(15) = xmlData.GetElementsByTagName("r:streamContent").Item(0).InnerText
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (Streamcontent/Title) = " & TrackInfo(15).ToString, LogType.LOG_TYPE_INFO)
            If MyZoneSourceExt = "Sirius" Then
                ' overwrite TrackInfo (2) -> title 
                TrackInfo(2) = TrackInfo(15)
            End If
        Catch ex As Exception
            TrackInfo(15) = ""
        End Try
        Try ' added v114
            TrackInfo(16) = xmlData.GetElementsByTagName("r:radioShowMd").Item(0).InnerText
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (RadioShow/Title) = " & TrackInfo(16).ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            TrackInfo(16) = ""
        End Try

        Try
            TrackInfo(0) = xmlData.GetElementsByTagName("dc:creator").Item(0).InnerText
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (creator) = " & TrackInfo(0).ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            'If g_bDebug Then Log( "Error for " & ZoneName & " TrackInfo (creator) = not found with error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            TrackInfo(1) = xmlData.GetElementsByTagName("upnp:album").Item(0).InnerText
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (album) = " & TrackInfo(1).ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        Try
            If (MyZoneSourceExt <> "Sirius") And (MyZoneSourceExt <> "SiriusXM") And (MyZoneSourceExt <> "Stream Radio") Then ' changed v114
                TrackInfo(2) = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (title) = " & TrackInfo(2).ToString, LogType.LOG_TYPE_INFO)
            End If
        Catch ex As Exception
        End Try
        Try
            If TrackInfo(4) = NoArtPath Or TrackInfo(4).trim = "" Then 'Else the info was already retrieved from the radio station
                'TrackInfo(4) = DecodeURI(xmlData.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText) removed in v3.1.0.19 because of failure to retrieve art
                TrackInfo(4) = BuildArtURL(TrackInfo(4))
                ' /getaa?m=1&u=http%3a%2f%2f79f84a69-0f02-4bd8-b6b0-f72f89ed86f3.x-udn%2fWMPNSSv4%2f33788882%2f1_ezBFQTc3NUU2LTUyNDctNEJDNS05QkVBLTYzMUJBMkZCMDIyRn0uMC5EOUQ4QTBDRQ.mp3%3falbumArt%3dtrue
                If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (albumArtURI) = " & TrackInfo(4).ToString, LogType.LOG_TYPE_INFO)
            End If

        Catch ex As Exception
            'If g_bDebug Then Log( "Error for " & ZoneName & " TrackInfo (albumArtURI) = not found with error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        TrackInfo(4) = GetAlbumArtPath(TrackInfo(4), False)
        Try
            TrackInfo(5) = PositionInfo(pmiRelTime)    'Track Position
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (Track Position) = " & TrackInfo(5).ToString, LogType.LOG_TYPE_INFO)
            If ZoneSource <> "Linked" Then SetPlayerPosition(GetSeconds(TrackInfo(5)))
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (Track Position Seconds) = " & GetSeconds(TrackInfo(5)).ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            If g_bDebug Then Log(ZoneName & " TrackInfo Failed to convert to Seconds, with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            TrackInfo(6) = PositionInfo(pmiTrackDuration)    'Track Duration
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (Track Duration) = " & TrackInfo(6).ToString, LogType.LOG_TYPE_INFO)
            If ZoneSource <> "Linked" Then SetTrackLength(GetSeconds(TrackInfo(6)))
        Catch ex As Exception
        End Try
        Try
            TrackInfo(7) = CStr(PositionInfo(pmiTrack))     'Queue Position
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (Queue Position) = " & TrackInfo(7).ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        Try
            If GetSeconds(PositionInfo(pmiTrackDuration)) <> 0 Then
                TrackInfo(8) = CStr(Math.Round(GetSeconds(PositionInfo(pmiRelTime)) / GetSeconds(PositionInfo(pmiTrackDuration)) * 100)) '% played
            Else
                TrackInfo(8) = 100
            End If
            If g_bDebug Then Log(ZoneName & " TrackInfo PositionInfo (% played) = " & TrackInfo(8).ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
        End Try
        If MyZoneSourceExt = "Lastfm" Or MyZoneSourceExt = "Pandora" Or MyZoneSourceExt = "Sirius" Or MyZoneSourceExt = "Rhapsody" Then
            ' these stations have metadata and URIs that show up only in the GetPositionInfo call
            If PositionInfo(pmiTrackURI) <> "" Then TrackInfo(13) = PositionInfo(pmiTrackURI) ' Store the CurrentURI
            If PositionInfo(pmiTrackMetaData) <> "" Then TrackInfo(14) = PositionInfo(pmiTrackMetaData) ' Store the CurrentMetaData
        End If
        MyCurrentTrackInfo = TrackInfo
        GetCurrentTrackInfo = TrackInfo
    End Function



    Public Function SaveCurrentTrackInfo(ByVal LinkgroupName As String, ByVal SaveQueueFlag As Boolean, Optional DeleteSavedQueues As Boolean = False)
        SaveCurrentTrackInfo = ""
        MyDeleteQueueName = "" ' this will prevent a race condition when there was already an unlink with a subsequent link and timer running in the background
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("SaveCurrentTrackInfo called for zoneplayer - " & ZoneName & " for Linkgroup = " & LinkgroupName & " and SaveQueueFlag = " & SaveQueueFlag.ToString & " and DeleteSavedQueue = " & DeleteSavedQueues.ToString, LogType.LOG_TYPE_INFO)
        Dim LinkgroupInfo As SavedLinkInfo = Nothing
        Dim TempIndex As Integer
        If LinkGroupInfoArrayIndex = 0 Then
            ReDim Preserve LinkGroupInfoArray(LinkGroupInfoArrayIndex)
            LinkGroupInfoArray(LinkGroupInfoArrayIndex) = New SavedLinkInfo
            LinkgroupInfo = LinkGroupInfoArray(LinkGroupInfoArrayIndex)
            LinkGroupInfoArrayIndex = LinkGroupInfoArrayIndex + 1
            LinkgroupInfo.LinkgroupName = LinkgroupName
        Else
            For TempIndex = 0 To LinkGroupInfoArrayIndex - 1
                LinkgroupInfo = LinkGroupInfoArray(TempIndex)
                If LinkgroupInfo.LinkgroupName = LinkgroupName Then
                    ' the instance is already created
                    If LinkgroupInfo.InfoIsSaved Then
                        Log("Warning in SaveCurrentTrackInfo for zoneplayer - " & ZoneName & " Track info is being stored and overwriting previously stored info", LogType.LOG_TYPE_WARNING)
                        DeletePreviousSavedQueues("SCQueue" & "-" & ZoneName & "-" & LinkgroupName)
                    End If
                    LinkgroupInfo.FlushInfo()
                    Exit For
                End If
            Next
            If LinkgroupInfo.LinkgroupName <> LinkgroupName Then
                ' the instance doesn't exist create it
                ReDim Preserve LinkGroupInfoArray(LinkGroupInfoArrayIndex)
                LinkGroupInfoArray(LinkGroupInfoArrayIndex) = New SavedLinkInfo
                LinkgroupInfo = LinkGroupInfoArray(LinkGroupInfoArrayIndex)
                LinkGroupInfoArrayIndex = LinkGroupInfoArrayIndex + 1
                LinkgroupInfo.LinkgroupName = LinkgroupName
            End If
        End If
        MyPlayerStateBeforeAnnouncement = CurrentPlayerState
        LinkgroupInfo.MySavedZoneisLinked = MyZoneIsLinked
        If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved ZoneIsLinked = " & LinkgroupInfo.MySavedZoneisLinked.ToString, LogType.LOG_TYPE_INFO)
        LinkgroupInfo.MySavedSourceLinkedZone = MySourceLinkedZone
        If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved SourceLinkedZone = " & LinkgroupInfo.MySavedSourceLinkedZone.ToString, LogType.LOG_TYPE_INFO)
        Try
            ' now that we don't call GetCurrentTrackInfo anymore, the trackposition info is wrong. I can use the own tracking of the Musicapi
            MyCurrentTrackInfo(5) = ConvertSecondsToTimeFormat(SonosPlayerPosition)
            LinkgroupInfo.MySavedTrackInfo = MyCurrentTrackInfo 'GetCurrentTrackInfo()
            If g_bDebug Then ' changed on 7/14/2019 in v3.1.0.34 from superdebug
                Dim Index As Integer
                For Index = 0 To 12
                    Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " with Track(" & Index.ToString & ") = " & LinkgroupInfo.MySavedTrackInfo(Index).ToString, LogType.LOG_TYPE_INFO)
                Next
            End If
        Catch ex As Exception
            Log("Error in SaveCurrentTrackInfo ZonePlayer = " & ZoneName & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Queue position = " & LinkgroupInfo.MySavedQueuePosition.ToString, LogType.LOG_TYPE_INFO)
            LinkgroupInfo.MySavedTransportState = MyCurrentTransportState 'GetTransportState()
            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Player State = " & LinkgroupInfo.MySavedTransportState.ToString, LogType.LOG_TYPE_INFO)
            LinkgroupInfo.MySavedMasterVolumeLevel = MyCurrentMasterVolumeLevel ' GetVolumeLevel("Master")
            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Master Volume = " & LinkgroupInfo.MySavedMasterVolumeLevel.ToString, LogType.LOG_TYPE_INFO)
            LinkgroupInfo.MySavedMuteState = MyCurrentMuteState 'GetMuteState("Master")
            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Mute State = " & LinkgroupInfo.MySavedMuteState.ToString, LogType.LOG_TYPE_INFO)
            LinkgroupInfo.MySavedPlayMode = MyShuffleState
            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved PlayMode State = " & LinkgroupInfo.MySavedPlayMode.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("Error in SaveCurrentTrackInfo (1) for ZonePlayer = " & ZoneName & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

        Try
            If MySWVersion < 420 Then If DeleteSavedQueues Then DeletePreviousSavedQueues("SCQueue" & "-" & ZoneName & "-" & LinkgroupName)
        Catch ex As Exception
        End Try

        Try

            If (SaveQueueFlag Or (MySWVersion < 420)) Then '  And TracksInQueue("0") <> 0 Then ' removed in V21
                SaveQueueFlag = False
                If MyCurrentTrackInfo IsNot Nothing Then ' added in v21 on Sept 26,2017
                    If Mid(MyCurrentTrackInfo(3), 1, 15) = "x-rincon-queue:" Then
                        ' ok we are playing from our queue, question is which one .. x-rincon-queue:RINCON_5CAAFD9CFF5601400#16
                        Dim QueueParts As String() = Split(MyCurrentTrackInfo(3), "#")
                        If QueueParts IsNot Nothing Then
                            If UBound(QueueParts) > 0 Then
                                Dim QueueNbr As String = QueueParts(1)
                                If QueueNbr = "0" And TracksInQueue("0") <> 0 Then
                                    ' we also need to save the queue info
                                    LinkgroupInfo.MySavedQueueObjectID = SaveQueue("SCQueue" & "-" & ZoneName & "-" & LinkgroupName)
                                    If LastStoredQueueID = LinkgroupInfo.MySavedQueueObjectID Then ' this has failed where the queueid of one but last player suddenly shows up
                                        ' OK, this is a problem
                                        If g_bDebug Then Log("WARNING : SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Queue with ID = " & LinkgroupInfo.MySavedQueueObjectID & " but this QueueID is already in use, re-try ....", LogType.LOG_TYPE_WARNING)
                                        wait(1)  ' needed to let all players sync
                                        LinkgroupInfo.MySavedQueueObjectID = SaveQueue("SCQueue" & "-" & ZoneName & "-" & LinkgroupName)
                                        If LastStoredQueueID = LinkgroupInfo.MySavedQueueObjectID Then
                                            ' OK, this is a problem
                                            If g_bDebug Then Log("ERROR : SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Queue with ID = " & LinkgroupInfo.MySavedQueueObjectID & " but this QueueID is already in use", LogType.LOG_TYPE_ERROR)
                                        Else
                                            LastStoredQueueID = LinkgroupInfo.MySavedQueueObjectID
                                            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Queue with ID = " & LinkgroupInfo.MySavedQueueObjectID, LogType.LOG_TYPE_INFO)
                                            SaveQueueFlag = True
                                        End If
                                    Else
                                        LastStoredQueueID = LinkgroupInfo.MySavedQueueObjectID
                                        If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved Queue with ID = " & LinkgroupInfo.MySavedQueueObjectID, LogType.LOG_TYPE_INFO)
                                        SaveQueueFlag = True
                                    End If
                                    If MySWVersion < 420 Then wait(0.5) ' needed to let all players sync. Changed from .25 to .50 in v.92 because problem still appears
                                Else    ' added 2/24/2019 in v3.1.0.29
                                    ' this means the player is playing something from their non standard queue like for example playing through Alexa!
                                    ' I will have to save the queue the old way, including restoring it using the new queue service as opposed to functions under AVTransport
                                    ' dcorAlexa
                                    Dim UpdateId As Integer = 0
                                    LinkgroupInfo.MySavedQueue = BrowseQueue(QueueNbr, UpdateId)
                                    LinkgroupInfo.MySavedTrackInfo(3) = QueueParts(0)  ' remove the queueID , which will be added when restored AFTER a new queue was created
                                    SaveQueueFlag = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If MyCurrentTrackInfo(9) = "Linked" And InStr(MyCurrentTrackInfo(3), "-sonos-dock:") <> 0 Then ' this zone is linked to a wireless dock. First pause it
                MyWirelessDockSourcePlayer.SetTransportState("Pause")
            End If
            LinkgroupInfo.MySavedSavedQueueFlag = SaveQueueFlag
            LinkgroupInfo.InfoIsSaved = True
            LinkgroupInfo.MySavedTargetZoneLinkedList = MyTargetZoneLinkedList
            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved MyTargetZoneLinkedList = " & LinkgroupInfo.MySavedTargetZoneLinkedList.ToString, LogType.LOG_TYPE_INFO)
            LinkgroupInfo.MySavedZoneIsPairMaster = MyZoneIsPairMaster
            LinkgroupInfo.MySavedZoneIsPairSlave = MyZoneIsPairSlave
            LinkgroupInfo.MySavedZonePairMasterZoneName = MyZonePairMasterZoneName
            LinkgroupInfo.MySavedZonePairMasterUDN = MyZonePairMasterZoneUDN
            LinkgroupInfo.MySavedZonePairSlaveUDN = MyZonePairSlaveZoneUDN
            LinkgroupInfo.MySavedZonePairSubWooferZoneUDN = MyZonePairSubWooferZoneUDN
            If MyZoneIsPairMaster Then
                If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & ". Zone is pair master and PairSlaveZoneUDN = " & MyZonePairSlaveZoneUDN.ToString, LogType.LOG_TYPE_INFO)
            ElseIf MyZoneIsPairSlave Then
                If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " . Zone is pair slave and PairMasterZoneUDN =" & MyZonePairMasterZoneUDN.ToString, LogType.LOG_TYPE_INFO)
            End If
            LinkgroupInfo.MySavedChannelMapSet = MyChannelMapSet
            If g_bDebug Then Log("SaveCurrentTrackInfo for zoneplayer " & ZoneName & " saved ChannelMapSet = " & LinkgroupInfo.MySavedChannelMapSet.ToString, LogType.LOG_TYPE_INFO)
            MyZoneIsStored = True

        Catch ex As Exception
            Log("Error in SaveCurrentTrackInfo (2) for ZonePlayer = " & ZoneName & " Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function RestoreCurrentTrackInfo(ByVal LinkgroupName As String, ByVal VolumeOnly As Boolean) As Boolean
        RestoreCurrentTrackInfo = True
        MyZoneIsStored = False
        If DeviceStatus = "Offline" Then Exit Function
        Dim xmlData As XmlDocument = New XmlDocument
        Dim LinkgroupInfo As SavedLinkInfo = Nothing
        Dim TempIndex As Integer
        If LinkGroupInfoArrayIndex = 0 Then
            Log("Warning in RestoreCurrentTrackInfo. Saved Link Info not found for Zonename = " & ZoneName & " and LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_WARNING)
            Exit Function ' nothing was ever saved
        End If
        For TempIndex = 0 To LinkGroupInfoArrayIndex - 1
            LinkgroupInfo = LinkGroupInfoArray(TempIndex)
            If LinkgroupInfo.LinkgroupName = LinkgroupName Then
                ' the instance is existing
                Exit For
            End If
        Next
        If LinkgroupInfo.LinkgroupName <> LinkgroupName Then
            ' the instance doesn't exist 
            If g_bDebug Then Log("Error in RestoreCurrentTrackInfo. Saved Link Info not found for Zonename = " & ZoneName & " and LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_ERROR)
            Exit Function
        End If
        If Not LinkgroupInfo.InfoIsSaved Then
            Log("Warning in RestoreCurrentTrackInfo. No Saved Link Info for Zonename = " & ZoneName & " and LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_WARNING)
            Exit Function
        End If
        If VolumeOnly Then
            SetVolumeLevel("Master", LinkgroupInfo.MySavedMasterVolumeLevel) ' restore volume
            LinkgroupInfo.FlushInfo()
            Exit Function
        End If
        If g_bDebug Then Log("RestoreCurrentTrackInfo for Zonename = " & ZoneName & " and MySavedTrackinfo = " & LinkgroupInfo.MySavedTrackInfo(3).ToString, LogType.LOG_TYPE_INFO)
        ' the reason this is here, else I cannot restore the queue because the players are still linked! The PlayURI will unlink and I suspect the reason why I sometimes lose commands
        ' is because they are busy just doing that ... unlinking
        ' as of version .95 an unlink is issued for all zone so I can remove this
        'If Mid(LinkgroupInfo.MySavedTrackInfo(3), 1, 10) = "pndrradio:" Then
        '    try to unlink first before playing the radio, see if that takes away the error
        '    If g_bDebug Then Log( "Warning RestoreCurrentTrackInfo detected Pandora for zoneplayer " & ZoneName & " and therefore issues an unlink first")
        '    PlayURI("x-rincon-queue:" & GetUDN() & "#0", "")
        'End If
        If LinkgroupInfo.MySavedChannelMapSet <> "" And LinkgroupInfo.MySavedZonePairSubWooferZoneUDN <> "" Then ' do this before to trap the case where the zone is linked to another player, which the sub doesn't like
            Dim SUBZonePlayer As HSPI '  HSMusicAPI
            SUBZonePlayer = MyHSPIControllerRef.GetAPIByUDN(LinkgroupInfo.MySavedZonePairSubWooferZoneUDN)
            If Not SUBZonePlayer Is Nothing Then
                SUBZonePlayer.PlayURI("x-rincon:" & UDN, "") ' group it back to the zone here v.92
                wait(0.5) ' give it some breathing room
            End If
        End If

        ' Some folks have been complaining that they have a brief high volume restored track afer an announcement is over.
        ' The volume adjustment was adjusted later in the flow but I'm going to add it here. 7/9/2019 version 3.1.0.31
        ' I believe it was done later in the flow because I had issues where I was setting the volume and the player wasn't unlinked yet, causing the command to be refused.
        SetVolumeLevel("Master", LinkgroupInfo.MySavedMasterVolumeLevel) ' restore volume
        SetMute("Master", LinkgroupInfo.MySavedMuteState)

        If LinkgroupInfo.MySavedQueue = "" Then PlayURI(LinkgroupInfo.MySavedTrackInfo(3), LinkgroupInfo.MySavedTrackInfo(12))  ' not sure why I do this before I restore the queue ???

        wait(0.25) ' give it some breathing room

        If LinkgroupInfo.MySavedQueueObjectID <> "" Then
            If g_bDebug Then Log("RestoreCurrentTrackInfo is restoring a saved queue for zoneplayer " & ZoneName & " with MySavedQueueObjectID = " & LinkgroupInfo.MySavedQueueObjectID, LogType.LOG_TYPE_INFO)
            ClearQueue()
            Dim MetaData As String
            Dim QueueURI As String
            MetaData = GetTrackMetaData(LinkgroupInfo.MySavedQueueObjectID, True)
            Try
                xmlData.LoadXml(MetaData)
                QueueURI = xmlData.GetElementsByTagName("res").Item(0).InnerText
                AddTrackToQueue(QueueURI, MetaData, 0, False)
                DestroySonosObject(LinkgroupInfo.MySavedQueueObjectID)
                LinkgroupInfo.MySavedQueueObjectID = ""
            Catch ex As Exception
                Log("Error in RestoreCurrentTrackInfo while restoring queueinfo for zoneplayer " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        ElseIf LinkgroupInfo.MySavedSavedQueueFlag And LinkgroupInfo.MySavedQueue <> "" Then ' totally rewritten 2/26/2019 in v3.1.29
            ' we also need to restore the queue info
            If g_bDebug Then Log("RestoreCurrentTrackInfo is restoring a queue from memory for zoneplayer " & ZoneName, LogType.LOG_TYPE_INFO)
            Dim QueueURI As String = ""
            Try
                xmlData.LoadXml(LinkgroupInfo.MySavedQueue)
                Dim Items As XmlNodeList = Nothing
                Items = xmlData.GetElementsByTagName("item")
                If Items IsNot Nothing Then
                    ' creata a new queue 
                    Dim UpdateID As Integer = 1
                    Dim QueueID As Integer = CreateQueue("MediaController", "", "")
                    If QueueID <> 0 Then
                        For Each Item As XmlNode In Items
                            Dim ItemxmlData As New XmlDocument
                            ItemxmlData.LoadXml(Item.OuterXml)
                            QueueURI = ItemxmlData.GetElementsByTagName("res").Item(0).InnerText
                            UpdateID = AddURI(QueueID, UpdateID, QueueURI, ItemxmlData.OuterXml, 0, True)
                            If UpdateID = 0 Then Exit For
                        Next
                        PlayURI(LinkgroupInfo.MySavedTrackInfo(3) & "#" & QueueID.ToString, "")
                    Else
                        If g_bDebug Then Log("Error in RestoreCurrentTrackInfo for zoneplayer " & ZoneName & ". Can't create a Queue", LogType.LOG_TYPE_ERROR)
                    End If
                End If
            Catch ex As Exception
                Log("Error in RestoreCurrentTrackInfo for zoneplayer " & ZoneName & " restoring queue from memory with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        End If
        If Mid(LinkgroupInfo.MySavedTrackInfo(3), 1, 15) = "x-rincon-queue:" Then
            If LinkgroupInfo.MySavedTrackInfo(7) <> "" And LinkgroupInfo.MySavedTrackInfo(7) <> "0" Then
                SeekTrack(LinkgroupInfo.MySavedTrackInfo(7))
                'End If
                'If LinkgroupInfo.MySavedTrackInfo(5) <> "" And LinkgroupInfo.MySavedTrackInfo(5) <> "00:00:00" Then
                SeekTime(LinkgroupInfo.MySavedTrackInfo(5)) ' set track position
            End If
        End If
        If Mid(LinkgroupInfo.MySavedTrackInfo(3), 1, 16) = "x-rincon-stream:" Then
            ' this is linked to an Audio Input
            ' the zone starts playing automatically so if it was paused before, do it as quick as possible
            SetTransportState(LinkgroupInfo.MySavedTransportState, True) ' ' Set Playerstate
        End If
        SetVolumeLevel("Master", LinkgroupInfo.MySavedMasterVolumeLevel) ' restore volume
        SetMute("Master", LinkgroupInfo.MySavedMuteState)
        SetPlayMode(LinkgroupInfo.MySavedPlayMode)
        If Mid(LinkgroupInfo.MySavedTrackInfo(3), 1, 16) <> "x-rincon-stream:" Then
            SetTransportState(LinkgroupInfo.MySavedTransportState, True) ' ' Set Playerstate and OverRuleLinkState
        End If
        ' check wether this zone was part of a pair or linked. If it was linked before the announcement and some of the other players did not participate in the announcement then we need to restore them
        If LinkgroupInfo.MySavedChannelMapSet <> "" And LinkgroupInfo.MySavedZoneIsPairMaster Then
            ' this zone was part of a pair group before being saved
        ElseIf LinkgroupInfo.MySavedZoneIsPairSlave Then
            Dim MasterPairZonePlayer As HSPI ' HSMusicAPI
            MasterPairZonePlayer = MyHSPIControllerRef.GetAPIByUDN(LinkgroupInfo.MySavedZonePairMasterUDN)
        End If
        LinkgroupInfo.FlushInfo()
        If MyPlayerStateBeforeAnnouncement <> player_state_values.playing Then
            ' we've been supressing these but now there is no off event to turn amps off if the zone is silent
            PlayChangeNotifyCallback(player_status_change.PlayStatusChanged, MyPlayerStateBeforeAnnouncement)
            hs.SetDeviceValue(HSRefPlayer, MyPlayerStateBeforeAnnouncement)
            If g_bDebug Then Log("RestoreCurrentTrackInfo updated HS DeviceValue for zonePlayer " & ZoneName & ". HSRef = " & HSRefPlayer.ToString & " and value = " & MyPlayerStateBeforeAnnouncement.ToString, LogType.LOG_TYPE_INFO)
        End If
        MyDeleteQueueName = "SCQueue" & "-" & ZoneName & "-" & LinkgroupName ' This will clean up any "hanging" saved queues added in v.100
        If g_bDebug Then Log("RestoreCurrentTrackInfo is done for zonePlayer " & ZoneName, LogType.LOG_TYPE_INFO)
    End Function


    Public Function CheckAnnouncementHasStarted() As Boolean
        CheckAnnouncementHasStarted = False
        If g_bDebug Then Log("CheckAnnouncementHasStarted called for Zoneplayer = " & ZoneName & " and CurrentplayerState = " & CurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then
            CheckAnnouncementHasStarted = True
            HasAnnouncementStarted = True
            Exit Function
        End If
        Dim TempTransportstate As String = MyCurrentTransportState
        GetTransportState()
        If TempTransportstate <> MyCurrentTransportState Then
            UpdateHS(True)
        End If
        If g_bDebug Then Log("CheckAnnouncementHasStarted called for Zoneplayer = " & ZoneName & " and New CurrentplayerState = " & CurrentPlayerState.ToString, LogType.LOG_TYPE_INFO)
        Dim InArg(0)
        Dim PositionInfo(7) ' 0-Track, 1-TrackDuration, 2-TrackMetaData, 3-TrackURI, 4-RelTime, 5-AbsTime, 6-RelCount, 7-AbsCount
        InArg = {""}

        Try
            AVTransport.InvokeAction("GetPositionInfo", InArg, PositionInfo)
        Catch ex As Exception
            Log("Error in CheckAnnouncementHasStarted/GetPositionInfo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Dim Position As Integer = GetSeconds(PositionInfo(pmiRelTime))
        If MyCurrentPlayerState = player_state_values.Playing And Position > 0 Then ' changed in v3.0.0.18 & 3.0.0.20 (in v20 the GetTransportState would show playing whilst still processing a transport event indicating stopped or transitioning)
            ' this part is a problem. I see the trackpos <> 0 but player is still stopped so what happens next is that the PI assumes the announcement has started and next 500ms, 
            ' it checks the state and detect is is still in stopped state and assumes the announcement to be over and unlinks
            CheckAnnouncementHasStarted = True
            HasAnnouncementStarted = True
        ElseIf Position > 1 Then
            If g_bDebug Then Log("CheckAnnouncementHasStarted called for Zoneplayer = " & ZoneName & " has trackposition = " & Position.ToString & " has Previoustrackposition = " & MyPreviousAnnouncementPosition.ToString & " has NbrofTimesUnchanged = " & MyAnnouncementPositionHasNotChangedNbrofTimes.ToString, LogType.LOG_TYPE_INFO)
            If Position <> MyPreviousAnnouncementPosition Then
                MyAnnouncementPositionHasNotChangedNbrofTimes = 0
            Else
                MyAnnouncementPositionHasNotChangedNbrofTimes += 1
                If MyAnnouncementPositionHasNotChangedNbrofTimes > MaxAllowedAnnouncementPositionHasNotChanged Then
                    ' we assume here that the announcement really started, it stopped, but state was never updated
                    CheckAnnouncementHasStarted = True
                    HasAnnouncementStarted = True
                End If
            End If
        Else
            If g_bDebug Then Log("CheckAnnouncementHasStarted called for Zoneplayer = " & ZoneName & " has trackposition = " & Position.ToString, LogType.LOG_TYPE_INFO)
        End If
    End Function

    Private Sub DeletePreviousSavedQueues(QueueName As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim InArg(5)
        Dim OutArg(3)
        Dim LoopIndex As Integer
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXMLData As String
        Dim OuterXML As XmlDocument = New XmlDocument

        If g_bDebug Then Log("DeletePreviousSavedQueues called for ZonePlayer = " & ZoneName & " with QueueName = " & QueueName, LogType.LOG_TYPE_INFO)

        InArg(0) = "SQ:"                    ' Object ID     String
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in DeletePreviousSavedQueues/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        If g_bDebug Then Log("DeletePreviousSavedQueues found " & TotalMatches.ToString & " playlists for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        If TotalMatches = 0 Or NumberReturned = 0 Then Exit Sub ' added to v101

        StartIndex = 0

        Do
            InArg(3) = CStr(StartIndex)
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in DeletePreviousSavedQueues/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try

            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                If TotalMatches = 0 Or NumberReturned = 0 Then Exit Sub ' added to v101
                For LoopIndex = 0 To NumberReturned - 1
                    Try
                        OuterXMLData = xmlData.GetElementsByTagName("container").Item(LoopIndex).OuterXml
                        OuterXML.LoadXml(OuterXMLData)
                    Catch ex As Exception
                        Log("Error in DeletePreviousSavedQueues extracting an container element with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        If OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText = QueueName Then
                            Dim ObjectID As String = ""
                            Try
                                ObjectID = xmlData.GetElementsByTagName("container").Item(LoopIndex).Attributes("id").Value
                                If ObjectID <> "" Then DestroySonosObject(ObjectID)
                            Catch ex As Exception
                            End Try
                        End If
                    Catch ex As Exception
                    End Try
                Next
            Catch ex As Exception
                Log("DeletePreviousSavedQueues for zoneplayer = " & ZoneName & ": Error " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            If g_bDebug Then Log("DeletePreviousSavedQueues for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
            'hs.WaitEvents()
        Loop
        If g_bDebug Then Log("Finished DeletePreviousSavedQueues for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub HandleLinkedZones(ByVal LinkedUDN As String)
        If g_bDebug Then Log("HandleLinkedZones called for ZonePlayer = " & ZoneName & ". LinkedUDN = " & LinkedUDN, LogType.LOG_TYPE_INFO)
        If Mid(LinkedUDN, 1, 9) = "x-rincon:" Then
            Mid(LinkedUDN, 1, 9) = "         "
            LinkedUDN = Trim(LinkedUDN)
            MusicService = "Linked to " & MyHSPIControllerRef.GetZoneByUDN(LinkedUDN)
        ElseIf Mid(LinkedUDN, 1, 13) = "x-sonos-dock:" Then
            Mid(LinkedUDN, 1, 13) = "             "
            LinkedUDN = Trim(LinkedUDN)
            MusicService = "Linked to " & MyHSPIControllerRef.GetZoneByUDN(LinkedUDN)
        End If
        If MyZoneIsLinked Then
            ' this zone might be linked to something else
            If MySourceLinkedZone <> LinkedUDN Then
                If g_bDebug Then Log("HandleLinkedZones called for ZonePlayer = " & ZoneName & " is going to unlink first. MySourceLinkedUDN = " & MySourceLinkedZone & " and New LinkedUDN = " & LinkedUDN, LogType.LOG_TYPE_INFO)
                HandleUnlinkZone()
            Else
                If g_bDebug Then Log("HandleLinkedZones called for ZonePlayer = " & ZoneName & " and already linked", LogType.LOG_TYPE_INFO)
                Exit Sub
            End If
        End If

        MySourceLinkedZone = LinkedUDN
        ' Challenge is to now notify the other zone that it is the master and that any change to that zone needs to be reported back
        ' For WD100 wireless dock device, the oposite is true, this zone is the master!

        Try
            'MyHSPIControllerRef.LinkAZone(ZoneName, LinkedUDN)
            If Not MyHSPIControllerRef.LinkAZone(UDN, LinkedUDN) Then
                MyRetryLinkAZone = True
                Exit Sub
            Else
                MyZoneIsLinked = True
            End If
            Dim TempSourceZonePlayer As HSPI = MyHSPIControllerRef.GetAPIByUDN(LinkedUDN)
            If TempSourceZonePlayer IsNot Nothing Then
                If TempSourceZonePlayer.ZoneModel = "WD100" Then
                    ' OK store the reference to the controller of the source WD100 player
                    MyWirelessDockSourcePlayer = TempSourceZonePlayer
                    'MyWirelessDockSourceMusicAPI = MyHSPIControllerRef.GetMusicApI(TempSourceZoneName)
                    MyWirelessDockZoneName = TempSourceZonePlayer.GetZoneName
                    If g_bDebug Then Log("HandleLinkedZones stored info on Wireless Doc for ZonePlayer = " & ZoneName & ". Wireless Dock = " & MyWirelessDockZoneName, LogType.LOG_TYPE_INFO)
                    MyWirelessDockSourcePlayer.SetShuffleState(MyShuffleState)
                    MyWirelessDockSourcePlayer.SetVolume = MyCurrentMasterVolumeLevel
                    MyWirelessDockSourcePlayer.SetMuteState = MyCurrentMuteState
                    MyWirelessDockSourcePlayer.SetLoudness = MyCurrentLoudnessState
                    MyWirelessDockSourcePlayer.SetHSBalance = MyBalance
                End If
            Else
                Log("Error in HandleLinkedZones finding target zone. ZonePlayer = " & ZoneName & ". LinkedUDN = " & LinkedUDN, LogType.LOG_TYPE_INFO)
            End If
        Catch ex As Exception
            Log("Error in HandleLinkedZones calling Sonos. ZonePlayer = " & ZoneName & ". LinkedUDN = " & LinkedUDN & " Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Sub HandleUnlinkZone()
        MyZoneIsLinked = False
        If g_bDebug Then Log("HandleUnlinkZone called for ZonePlayer = " & ZoneName & ". LinkedUDN = " & MySourceLinkedZone, LogType.LOG_TYPE_INFO)
        Try
            MyHSPIControllerRef.UnlinkAZone(UDN, MySourceLinkedZone)
        Catch ex As Exception
            Log("Error in HandleUnlinkZone calling Sonos. ZonePlayer = " & ZoneName & ". LinkedUDN = " & MySourceLinkedZone & " Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        MySourceLinkedZone = ""
        ' OK remove the reference to the controller and the MusicAPI
        MyWirelessDockSourcePlayer = Nothing
        MyWirelessDockZoneName = ""
    End Sub

    Public Sub AddTargetLinkZone(ByVal ZoneUDN As String)
        ' this procedure is used for Zones that are linked and this instance is controlling the source to that linked zone. Each time a zone is added this procudure is called
        ' It keeps a list of zones so that each time there is an event at the source, all target instances of zoneplayercontroller are informed so they subsequently can generate
        ' events for HS
        Dim TargetZones() = {}
        Dim TargetZone
        Dim AlreadyExist As Boolean = False
        Try
            TargetZones = Split(MyTargetZoneLinkedList, ";")
        Catch ex As Exception
        End Try
        For Each TargetZone In TargetZones
            If TargetZone.ToString = ZoneUDN Then
                ' Zone already exist .... shouldn't be but it is
                AlreadyExist = True
                Exit For
            End If
        Next
        If Not AlreadyExist Then
            MyTargetZoneLinkedList = MyTargetZoneLinkedList + ";" + ZoneUDN
            MyTargetZoneLinkedList = Trim(MyTargetZoneLinkedList).Trim(Char.Parse(";"))
        End If
        MyTargetZoneLinkedList = Trim(MyTargetZoneLinkedList)
        If MyTargetZoneLinkedList <> "" Then
            MyZoneIsSourceForLinkedZone = True
            If ZoneModel = "WD100" Then
                ' keep a reference to the target zone for a WD100 player so we can easily update/retrieve info from the target player like mutestate/vol/
                MyWirelessDockDestinationPlayer = MyHSPIControllerRef.GetAPIByUDN(ZoneUDN)
                MyWirelessDockDestinationPlayer.Artist = Artist
                MyWirelessDockDestinationPlayer.Track = Track
                MyWirelessDockDestinationPlayer.Album = Album
            ElseIf Not AlreadyExist Then
                ' update the newly added
                Dim NewPlayer As HSPI = Nothing
                NewPlayer = MyHSPIControllerRef.GetAPIByUDN(ZoneUDN)
                If NewPlayer IsNot Nothing Then
                    NewPlayer.Artist = Artist
                    NewPlayer.Track = Track
                    NewPlayer.Album = Album
                    NewPlayer.ArtworkURL = ArtworkURL
                    NewPlayer.LinkedCurrentPlayerState = CurrentPlayerState
                    NewPlayer.SetTrackLength(MyTrackLength)
                    NewPlayer.SetPlayerPosition(MyPlayerPosition)
                    NewPlayer.NextTrack = NextTrack
                    NewPlayer.NextAlbum = NextAlbum
                    NewPlayer.NextArtist = NextArtist
                    NewPlayer.RadiostationName = RadiostationName
                    NewPlayer.NextArtworkURL = NextArtworkURL
                    NewPlayer.MyZoneSourceExt = MyZoneSourceExt
                    NewPlayer.SetNbrOfTracks(MyNbrOfTracksInQueue)
                    NewPlayer.SetTrackNbr(MyTrackInQueueNbr)
                    NewPlayer.SetShuffleState(MyShuffleState)
                    NewPlayer.SetHSPlayerInfo()
                Else
                    If g_bDebug Then Log("AddTargetLinkZone called for ZoneName " & ZoneName & " but target Zone is not online. TargetUDN = " & ZoneUDN & " and AlreadExist = " & AlreadyExist.ToString, LogType.LOG_TYPE_INFO)
                End If
            End If
        End If
        If g_bDebug Then Log("AddTargetLinkZone called for ZoneName " & ZoneName & " with NewUDN = " & ZoneUDN & " and new Target UDN List= " & MyTargetZoneLinkedList & " and AlreadExist = " & AlreadyExist.ToString, LogType.LOG_TYPE_INFO)
    End Sub

    Public Sub RemoveTargetLinkZone(ByVal TargetUDN As String)
        ' this procedure is used for Zones that are linked and this instance is controlling the source to that linked zone. Each time a zone is removed this procudure is called
        ' It keeps a list of zones so that each time there is an event at the source, all target instances of zoneplayercontroller are informed so they subsequently can generate
        ' events for HS
        Dim TargetZones() = {}
        Dim TargetZone
        Dim AlreadyExist As Boolean = False
        Dim OldTargetList As String = MyTargetZoneLinkedList
        Try
            TargetZones = Split(MyTargetZoneLinkedList, ";")
        Catch ex As Exception
        End Try
        MyTargetZoneLinkedList = ""
        For Each TargetZone In TargetZones
            If TargetZone.ToString = TargetUDN Then
                ' Zone already exist, skip this
                AlreadyExist = True
            Else
                MyTargetZoneLinkedList = MyTargetZoneLinkedList + TargetZone.ToString + ";"
            End If
        Next
        If MyTargetZoneLinkedList <> "" Then
            ' remove last ";"
            MyTargetZoneLinkedList = Trim(MyTargetZoneLinkedList).Trim(Char.Parse(";"))
        End If
        If Not AlreadyExist Then
            ' wasn't found, shouldn't be
            If g_bDebug Then Log("WARNING: RemoveTargetLinkZone called for ZoneName " & ZoneName & " with not found TargetUDN = " & TargetUDN & " and current Zone Target List= " & OldTargetList, LogType.LOG_TYPE_WARNING)
            Exit Sub
        End If
        MyTargetZoneLinkedList = Trim(MyTargetZoneLinkedList)
        If MyTargetZoneLinkedList = "" Then
            MyZoneIsSourceForLinkedZone = False
            If ZoneModel = "WD100" Then
                ' keep a reference to the target zone for a WD100 player so we can easily update/retrieve info from the target player like mutestate/vol/
                MyWirelessDockDestinationPlayer = Nothing
            End If
        End If
        If g_bDebug Then Log("RemoveTargetLinkZone called for ZoneName " & ZoneName & " with TargetUDN= " & TargetUDN & " and new Zone Target List= " & MyTargetZoneLinkedList, LogType.LOG_TYPE_INFO)
    End Sub

    Public Function GetTargetZoneLinkedList() As String
        GetTargetZoneLinkedList = MyTargetZoneLinkedList
    End Function

    Public Function BecomeCoordinatorOfStandaloneGroup(Optional InstanceID As Integer = 0)
        BecomeCoordinatorOfStandaloneGroup = ""
        If g_bDebug Then Log("BecomeCoordinatorOfStandaloneGroup called for zoneplayer " & ZoneName, LogType.LOG_TYPE_INFO)
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If ZoneModel = "WD100" Then Exit Function
        Try
            Dim InArg(0)
            Dim OutArg(0)
            InArg(0) = InstanceID
            AVTransport.InvokeAction("BecomeCoordinatorOfStandaloneGroup", InArg, OutArg)
            BecomeCoordinatorOfStandaloneGroup = "OK"
        Catch ex As Exception
            If g_bDebug Then Log("Warning in BecomeCoordinatorOfStandaloneGroup for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_WARNING)
        End Try
    End Function


    Public Function GetZoneSourceName() As String
        GetZoneSourceName = ""
        If g_bDebug Then Log("GetZoneSourceName called for ZoneName " & ZoneName & " ZoneIsLinked = " & MyZoneIsLinked.ToString & " and SourceLinkedZone = " & MySourceLinkedZone, LogType.LOG_TYPE_INFO)
        If MyZoneIsLinked And MySourceLinkedZone <> "" Then
            ' MySourceLinkedZone holds the UDN to the zone it is linked to. Retrieve the Zone Name from the .ini file
            Dim SourceZoneUDN As String()
            SourceZoneUDN = MySourceLinkedZone.Split(":") ' this is used to remove the uuid: if it is there
            If UBound(SourceZoneUDN) > 0 Then
                GetZoneSourceName = GetStringIniFile(SourceZoneUDN(1), DeviceInfoIndex.diFriendlyName.ToString, "")
                If g_bDebug Then Log("GetZoneSourceName called for ZoneName " & ZoneName & " Sourcename = " & GetZoneSourceName & " and SourceUDN = " & SourceZoneUDN(1), LogType.LOG_TYPE_INFO)
            Else
                GetZoneSourceName = GetStringIniFile(SourceZoneUDN(0), DeviceInfoIndex.diFriendlyName.ToString, "")
                If g_bDebug Then Log("GetZoneSourceName called for ZoneName " & ZoneName & " Sourcename = " & GetZoneSourceName & " and SourceUDN = " & SourceZoneUDN(0), LogType.LOG_TYPE_INFO)
            End If
        End If
    End Function

    Public Function GetZoneDestination() As String()
        'If g_bDebug Then Log("GetZoneDestination called for ZoneName " & ZoneName & " and TargetZones = " & MyTargetZoneLinkedList, LogType.LOG_TYPE_INFO)
        GetZoneDestination = {""}
        Dim TargetZones() = {}
        Try
            TargetZones = Split(MyTargetZoneLinkedList, ";")
            GetZoneDestination = TargetZones
        Catch ex As Exception
        End Try
    End Function

    Public Function GetZoneSourceUDN() As String
        GetZoneSourceUDN = ""
        If g_bDebug Then Log("GetZoneSourceUDN called for ZoneName " & ZoneName & " ZoneIsLinked = " & MyZoneIsLinked.ToString & " and SourceLinkedZone = " & MySourceLinkedZone, LogType.LOG_TYPE_INFO)
        If MyZoneIsLinked And MySourceLinkedZone <> "" Then
            ' MySourceLinkedZone holds the UDN to the zone it is linked to. Retrieve the Zone Name from the .ini file
            Dim SourceZoneUDN As String()
            SourceZoneUDN = MySourceLinkedZone.Split(":") ' this is used to remove the uuid: if it is there
            If UBound(SourceZoneUDN) > 0 Then
                GetZoneSourceUDN = SourceZoneUDN(1)
            Else
                GetZoneSourceUDN = SourceZoneUDN(0)
            End If
        End If
    End Function

    Public Function PlayModeNormal()
        PlayModeNormal = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If g_bDebug Then Log("PlayModeNormal called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)

            InArg(0) = 0
            InArg(1) = "NORMAL"

            AVTransport.InvokeAction("SetPlayMode", InArg, OutArg)
            PlayModeNormal = "OK"
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in PlayModeNormal for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function PlayModeShuffle()
        PlayModeShuffle = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If g_bDebug Then Log("PlayModeShuffle called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        Try

            Dim InArg(1)
            Dim OutArg(0)

            InArg(0) = 0
            InArg(1) = "SHUFFLE"

            AVTransport.InvokeAction("SetPlayMode", InArg, OutArg)
            PlayModeShuffle = "OK"
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in PlayModeShuffle for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function PlayModeShuffleNoRepeat()
        PlayModeShuffleNoRepeat = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If g_bDebug Then Log("PlayModeShuffleNoRepeat called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        Try

            Dim InArg(1)
            Dim OutArg(0)

            InArg(0) = 0
            InArg(1) = "SHUFFLE_NOREPEAT"

            AVTransport.InvokeAction("SetPlayMode", InArg, OutArg)

            PlayModeShuffleNoRepeat = "OK"
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in PlayModeShuffleNoRepeat for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function PlayModeRepeatAll()
        PlayModeRepeatAll = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If g_bDebug Then Log("PlayModeRepeatAll called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        Try

            Dim InArg(1)
            Dim OutArg(0)

            InArg(0) = 0
            InArg(1) = "REPEAT_ALL"

            AVTransport.InvokeAction("SetPlayMode", InArg, OutArg)

            PlayModeRepeatAll = "OK"
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in PlayModeRepeatAll for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Sub ToggleRepeat()
        If SonosRepeat() = 0 Then SonosRepeat = 2 Else SonosRepeat = 0
    End Sub

    Public Sub ToggleShuffle()
        If SonosShuffle() = 1 Then SonosShuffle = 2 Else SonosShuffle = 1
    End Sub


    Public Function SetTransportState(ByVal NewState As String, Optional ByVal OverRuleLinkState As Boolean = False)
        SetTransportState = ""
        If g_bDebug Then Log("SetTransportState called for zoneplayer - " & ZoneName & " with value = " & NewState, LogType.LOG_TYPE_INFO)
        If ZoneIsLinked And (Not OverRuleLinkState) Then
            Dim LinkedZone As HSPI ' HSMusicAPI
            LinkedZone = MyHSPIControllerRef.GetAPIByUDN(LinkedZoneSource.ToString)
            Try
                LinkedZone.SetTransportState(NewState, True)
            Catch ex As Exception
                If g_bDebug Then Log("SetTransportState called for Zone - " & ZoneName & " which was linked to " & LinkedZoneSource.ToString & " but ended in Error: " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
            Exit Function
        End If
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        If ZoneSource = "Input" And (UCase(NewState) = "STOP" Or UCase(NewState) = "STOPPED") Then
            ' Sonos player ignores Stop when in sourcing from input from another player"
            NewState = "Pause"
        End If
        Try

            Dim InArg(1)
            Dim OutArg(0)

            If UCase(NewState) = "PLAY" Or UCase(NewState) = "PLAYING" Then
                InArg(0) = 0
                InArg(1) = 1
                Try
                    AVTransport.InvokeAction("Play", InArg, OutArg)
                Catch ex As Exception
                    If g_bDebug Then Log("ERROR in SetTransportState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            ElseIf UCase(NewState) = "STOP" Or UCase(NewState) = "STOPPED" Then
                ReDim InArg(0)
                InArg(0) = 0
                'InArg(1) = 1
                Try
                    AVTransport.InvokeAction("Stop", InArg, OutArg)
                Catch ex As Exception
                    If g_bDebug Then Log("ERROR in SetTransportState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            ElseIf UCase(NewState) = "PAUSE" Or UCase(NewState) = "PAUSED_PLAYBACK" Then
                ReDim InArg(0)
                InArg(0) = 0
                Try
                    AVTransport.InvokeAction("Pause", InArg, OutArg)
                Catch ex As Exception
                    If g_bDebug Then Log("ERROR in SetTransportState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            ElseIf UCase(NewState) = "NEXT" Then
                ReDim InArg(0)
                InArg(0) = 0
                Try
                    AVTransport.InvokeAction("Next", InArg, OutArg)
                Catch ex As Exception
                    If g_bDebug Then Log("ERROR in SetTransportState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            ElseIf UCase(NewState) = "PREVIOUS" Then
                ReDim InArg(0)
                InArg(0) = 0
                Try
                    AVTransport.InvokeAction("Previous", InArg, OutArg)
                Catch ex As Exception
                    If g_bDebug Then Log("ERROR in SetTransportState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                End Try
            Else
                SetTransportState = "Error: strNewState value must be 'Play', 'Pause', 'Stop', 'Next', or 'Previous'"
                Exit Function
            End If
            SetTransportState = "OK"
        Catch ex As Exception
            Log("Error in SetTransportState for zoneplayer = " & ZoneName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetTransportState()
        GetTransportState = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        Try
            Dim InArg(0)
            Dim OutArg(2)
            InArg(0) = 0
            AVTransport.InvokeAction("GetTransportInfo", InArg, OutArg)
            MyCurrentTransportState = OutArg(0)
            If OutArg(0) = "PLAYING" Then
                GetTransportState = "Play"
                CurrentPlayerState = player_state_values.playing
            ElseIf OutArg(0) = "STOPPED" Then
                GetTransportState = "Stop"
                CurrentPlayerState = player_state_values.stopped
            ElseIf OutArg(0) = "PAUSED_PLAYBACK" Then
                GetTransportState = "Pause"
                CurrentPlayerState = player_state_values.paused
            Else
                GetTransportState = "Unknown"
                CurrentPlayerState = player_state_values.stopped
            End If
        Catch ex As Exception
            Log("ERROR in GetTransportState for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetSoftwareVersion() As Integer
        GetSoftwareVersion = 0
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(8)
        Try
            DeviceProperties.InvokeAction("GetZoneInfo", InArg, OutArg)
            Dim SWVersionParts As String() = OutArg(2).ToString.Split(".")
            Dim SwVersion As Integer = 0
            SwVersion = Val(SWVersionParts(0)) * 100
            If UBound(SWVersionParts) > 0 Then
                SwVersion = SwVersion + Val(SWVersionParts(1)) * 10
                If UBound(SWVersionParts) > 1 Then
                    SwVersion = SwVersion + Val(SWVersionParts(2))
                End If
            End If
            GetSoftwareVersion = SwVersion
            MyIPAddress = OutArg(4)
            MyMACAddress = OutArg(5)
        Catch ex As Exception
            Log("ERROR in GetSoftwareVersion for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetSoftwareVersion called for zoneplayer " & ZoneName & " with SW version = " & OutArg(2).ToString, LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetHouseholdID() As String ' asked by skavan
        GetHouseholdID = 0
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(0)
        Try
            DeviceProperties.InvokeAction("GetHouseholdID", InArg, OutArg)
            GetHouseholdID = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetHouseholdID for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetHouseholdID called for zoneplayer " & ZoneName & " with HouseholdID = " & OutArg(0).ToString, LogType.LOG_TYPE_INFO)
    End Function

    Public Function ListAvailableServices() As String() ' asked by skavan
        ListAvailableServices = Nothing
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(2) As String
        Try
            MusicServices.InvokeAction("ListAvailableServices", InArg, OutArg)
            ListAvailableServices = OutArg
        Catch ex As Exception
            Log("ERROR in ListAvailableServices for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("ListAvailableServices called for zoneplayer " & ZoneName, LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetMediaInfo() As String() ' asked by skavan
        GetMediaInfo = Nothing
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        InArg(0) = 0
        Dim OutArg(8) As String
        Try
            AVTransport.InvokeAction("GetMediaInfo", InArg, OutArg)
            GetMediaInfo = OutArg
        Catch ex As Exception
            Log("ERROR in GetMediaInfo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetMediaInfo called for zoneplayer " & ZoneName, LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetPositionInfo() As String() ' asked by skavan
        GetPositionInfo = Nothing
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        InArg(0) = 0
        Dim OutArg(7) As String
        Try
            AVTransport.InvokeAction("GetPositionInfo", InArg, OutArg)
            GetPositionInfo = OutArg
        Catch ex As Exception
            Log("ERROR in GetPositionInfo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If g_bDebug Then Log("GetPositionInfo called for zoneplayer " & ZoneName, LogType.LOG_TYPE_INFO)
    End Function

    Public Function SetEQ(EQType As String, DesiredValue As String) As String ' asked by w.vuyk
        ' EQTypes are DialogLevel SurroundLevel AudioDelay AudioDelayLeftRear AudioDelayRightRear NightMode SurroundEnabled SurroundMode
        SetEQ = ""
        If DeviceStatus = "Offline" Then Exit Function
        If RenderingControl Is Nothing Then Exit Function
        Dim InArg(2) As String
        Dim OutArg(0) As String
        InArg(0) = 0 ' = InstanceID
        InArg(1) = EQType
        InArg(2) = DesiredValue
        Try
            RenderingControl.InvokeAction("SetEQ", InArg, OutArg)
            SetEQ = "Ok"
        Catch ex As Exception
            Log("ERROR in SetEQ for zoneplayer = " & ZoneName & " with EQType = " & EQType & " with DesiredValue = " & DesiredValue & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If g_bDebug Then Log("SetEQ called for zoneplayer " & ZoneName & " with EQType = " & EQType & " with DesiredValue = " & DesiredValue, LogType.LOG_TYPE_INFO)
    End Function


    Public Function GetEQ(EQType As String) As String ' asked by w.vuyk
        GetEQ = ""
        If DeviceStatus = "Offline" Then Exit Function
        If RenderingControl Is Nothing Then Exit Function
        Dim InArg(1) As String
        InArg(0) = 0    ' InstanceID
        InArg(1) = EQType
        Dim OutArg(0) As String
        Try
            RenderingControl.InvokeAction("GetEQ", InArg, OutArg)
            GetEQ = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetEQ for zoneplayer = " & ZoneName & " with EQType = " & EQType & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If g_bDebug Then Log("GetEQ called for zoneplayer " & ZoneName & " with EQType = " & EQType & " and returned = " & OutArg(0), LogType.LOG_TYPE_INFO)
    End Function


    Public Function GetIPAddress() As String
        GetIPAddress = ""
        If DeviceStatus = "Offline" Then Exit Function
        If MyUPnPDevice Is Nothing Then Exit Function
        GetIPAddress = MyUPnPDevice.IPAddress
    End Function

    Public Function GetIPPort() As String
        GetIPPort = ""
        If DeviceStatus = "Offline" Then Exit Function
        If MyUPnPDevice Is Nothing Then Exit Function
        GetIPPort = MyUPnPDevice.IPPort
    End Function


    Public Function GetZoneName()
        GetZoneName = ""
        If DeviceStatus = "Offline" Then
            ' use what is stored
            GetZoneName = ZoneName
            Exit Function
        End If
        Try
            Dim InArg(0)
            Dim OutArg(2)
            DeviceProperties.InvokeAction("GetZoneAttributes", InArg, OutArg)
            GetZoneName = OutArg(0)
            'ZoneName = GetZoneName
        Catch ex As Exception
            Log("ERROR in GetZoneName for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetUDN() As String
        GetUDN = UDN
    End Function


    Public Function GetPlayMode()
        GetPlayMode = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        Try
            Dim InArg(0)
            Dim OutArg(1)
            InArg(0) = 0
            AVTransport.InvokeAction("GetTransportSettings", InArg, OutArg)
            SetShuffleState(OutArg(0))
            GetPlayMode = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetPlayMode for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function DelegateGroupCoordinationTo(NewMaster As String, JoinGroup As Boolean)
        DelegateGroupCoordinationTo = ""
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Function
        Try
            Dim InArg(2)
            Dim OutArg(0)
            InArg(0) = 0
            InArg(1) = NewMaster
            InArg(2) = JoinGroup
            AVTransport.InvokeAction("DelegateGroupCoordinationTo", InArg, OutArg)
            DelegateGroupCoordinationTo = "Ok"
        Catch ex As Exception
            Log("ERROR in DelegateGroupCoordinationTo for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function


    Public Sub SetPlayMode(ByVal PlayMode As String)
        If DeviceStatus = "Offline" Or AVTransport Is Nothing Then Exit Sub
        If g_bDebug Then Log("SetPlayMode called for ZoneName " & ZoneName & " with value = " & PlayMode, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(1)
            Dim OutArg(0)
            InArg(0) = 0
            InArg(1) = PlayMode
            AVTransport.InvokeAction("SetPlayMode", InArg, OutArg)
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetPlayMode for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Public Function CreateStereoPair(ByVal OtherZoneUDN As String) As String ' Creates a Stereo pair with this Zone Master (Left) 
        CreateStereoPair = ""
        If DeviceStatus = "Offline" Then Exit Function
        If Not CheckPlayerIsPairable(ZoneModel) Then Exit Function
        Dim InArg(0)
        Dim OutArg(0)
        ' the Channelmap looks like this RINCON_000E5858C97A01400:LF,LF;RINCON_000E5859008A0xxxx:RF,RF;RINCON_000E5898164001400:SW,SW
        InArg(0) = GetUDN()
        Try
            InArg(0) = Replace(InArg(0), "uuid:", "")
        Catch ex As Exception
        End Try
        Try
            OtherZoneUDN = Replace(OtherZoneUDN, "uuid:", "")
        Catch ex As Exception
        End Try
        InArg(0) = InArg(0) & ":LF,LF;" & OtherZoneUDN & ":RF,RF"
        If MyZonePairSubWooferZoneUDN <> "" Then InArg(0) = InArg(0) & ";" & MyZonePairSubWooferZoneUDN & ":SW,SW" ' new to version .91 to support subwoofers
        If g_bDebug Then Log("CreateStereoPair called with ChannelMapSet = " & InArg(0), LogType.LOG_TYPE_INFO)
        If MyChannelMapSet <> "" Then
            ' I think this will cause issues
            Log("Warning CreateStereoPair called with ChannelMapSet = " & InArg(0) & " but Player is already paired with ChannelMapSet = " & MyChannelMapSet, LogType.LOG_TYPE_WARNING)
        End If
        Try
            DeviceProperties.InvokeAction("CreateStereoPair", InArg, OutArg)
            CreateStereoPair = "OK"
        Catch ex As Exception
            Log("ERROR in CreateStereoPair for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            CreateStereoPair = "NOK"
        End Try
    End Function

    Public Function SeparateStereoPair() As String ' Unpairs a Stereo pair  
        SeparateStereoPair = ""
        If g_bDebug Then Log("SeparateStereoPair called for zoneplayer = " & ZoneName & " with ChannelMapSet = " & MyChannelMapSet, LogType.LOG_TYPE_INFO)
        If MyZoneIsPairSlave Then
            Dim MasterPlayer As HSPI = MyHSPIControllerRef.GetAPIByUDN(MyZonePairMasterZoneUDN)
            If MasterPlayer IsNot Nothing Then
                Return MasterPlayer.SeparateStereoPair()
                Exit Function
            End If
        End If
        If DeviceStatus = "Offline" Then Exit Function
        If Not CheckPlayerIsPairable(ZoneModel) Then Exit Function
        If MyChannelMapSet = "" Then Exit Function ' nothing to unpair
        Dim InArg(0)
        Dim OutArg(0)
        ' the Channelmap looks like this RINCON_000E5858C97A01400:LF,LF;RINCON_000E5859008A0xxxx:RF,RF;RINCON_000E5898164001400:SW,SW
        InArg(0) = MyChannelMapSet
        Try
            DeviceProperties.InvokeAction("SeparateStereoPair", InArg, OutArg)
            SeparateStereoPair = "OK"
        Catch ex As Exception
            Log("ERROR in SeparateStereoPair for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            SeparateStereoPair = "NOK"
        End Try
    End Function

    Public Function ZoneIsPairMaster() As Boolean
        ZoneIsPairMaster = MyZoneIsPairMaster
    End Function

    Public Function ZoneIsPairSlave() As Boolean
        ZoneIsPairSlave = MyZoneIsPairSlave
    End Function


    Public ReadOnly Property ZoneIsASlave As Boolean
        Get
            ZoneIsASlave = MyZoneIsPairSlave Or MyZoneIsPlaybarSlave
        End Get
    End Property

    Public ReadOnly Property ZoneMasterUDN As String
        Get
            If MyZoneIsPairSlave Then Return MyZonePairMasterZoneUDN
            If MyZoneIsPlaybarSlave Then Return MyZonePlayBarUDN
            Return ""
        End Get
    End Property

    Public ReadOnly Property PlayBarMaster As Boolean
        Get
            PlayBarMaster = MyZoneIsPlaybarMaster
        End Get
    End Property

    Public ReadOnly Property PlayBarSlave As Boolean
        Get
            PlayBarSlave = MyZoneIsPlaybarSlave
        End Get
    End Property

    Public Function GetZonePairSlaveUDN() As String
        GetZonePairSlaveUDN = MyZonePairSlaveZoneUDN
    End Function

    Public Function GetZonePairMasterUDN() As String
        GetZonePairMasterUDN = MyZonePairMasterZoneUDN
    End Function

    Public Function BuildTrackDatabase(ByVal DatabasePath As String)
        Dim ConnectionString
        BuildTrackDatabase = "NOK"
        If g_bDebug Then Log("BuildTrackDatabase for zoneplayer = " & ZoneName & " with " & DatabasePath, LogType.LOG_TYPE_INFO)
        '
        hs.SetDeviceValueByRef(MasterHSDeviceRef, msBuildingDB, True)
        MusicDBIsBeingEstablished = True
        ' save the SettingsReplicationState in the inifile so we can detect when changes have occured
        WriteBooleanIniFile("SettingsReplicationState", "SonosSettingsHaveChanged", False)
        ' create DB and erase existing
        Try
            ConnectionString = DBConnectionString & DatabasePath
            CreateTrackDatabase(DatabasePath, ConnectionString)
            If MyMusicDBItems.Tracks Then BuildTrackDB(DatabasePath, "")
            If MyMusicDBItems.Artists Then BuildArtistDB(DatabasePath, "")
            If MyMusicDBItems.Albums Then BuildAlbumDB(DatabasePath, "")
            If MyMusicDBItems.Radiostations Then BuildRadioStationDB(DatabasePath)
            If MyMusicDBItems.Playlists Then BuildPlaylistDB(DatabasePath, "")
            If MyMusicDBItems.Playlists Then BuildSonosPlaylistDB(DatabasePath)
            If MyMusicDBItems.Genres Then BuildGenreDB(DatabasePath, "")
            BuildSpecificObjectDB(DatabasePath, "FV:2", "Favorites")
            BuildTrackDatabase = "OK"
        Catch ex As Exception
            Log("An Error occurred in BuildTrackDatabase for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        ' hs.SetDeviceString(MasterHSDeviceRef, "Connected", False)
        hs.SetDeviceValueByRef(MasterHSDeviceRef, msConnected, True)
        MusicDBIsBeingEstablished = False
        If g_bDebug Then Log("BuildTrackDatabase Done  for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        GC.Collect()
    End Function

    Public Function GetTracks(ByVal SearchString As String, ByVal SearchItemObject As Boolean, NamesOnly As Boolean)
        ' search itemobject should be true when looking for tracks or RadioStations
        ' For Genre the info is found in the <container> section not <item>

        Dim xmlData As XmlDocument = New XmlDocument
        Dim InArg(5)
        Dim OutArg(3)
        Dim Tracks(,)
        Dim LoopIndex As Integer
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXMLData As String
        Dim OuterXML As XmlDocument = New XmlDocument

        GetTracks = {}
        '
        If g_bDebug Then Log("GetTracks called for ZonePlayer = " & ZoneName & " with Input = " & SearchString & " and SearchItemObject = " & SearchItemObject.ToString, LogType.LOG_TYPE_INFO)

        If ZoneModel = "WD100" Then
            If MyDockediPodPlayerName = "" Then Exit Function
            If ContentDirectory Is Nothing Then Exit Function
            If Not MakeiPodBrowseable() Then Exit Function
        End If

        InArg(0) = SearchString             ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4 -- should be 1 but that gives errors for wireless doc Genres
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in GetTracks/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("ERROR in GetTracks/Browse for zoneplayer = " & ZoneName & " with SearchString = " & SearchString, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        ReDim Tracks(CInt(OutArg(2)), 12)

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        If g_bDebug Then Log("GetTracks found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        StartIndex = 0

        Do

            InArg(3) = StartIndex
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in GetTracks/Browse1 for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Log("ERROR in GetTracks/Browse for zoneplayer = " & ZoneName & " with SearchString = " & SearchString & " and StartIndex = " & InArg(3).ToString & " and Nbr = " & InArg(4).ToString, LogType.LOG_TYPE_ERROR)

                Exit Do
            End Try

            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                For LoopIndex = 0 To NumberReturned - 1
                    Try
                        If SearchItemObject Then
                            OuterXMLData = xmlData.GetElementsByTagName("item").Item(LoopIndex).OuterXml
                        Else
                            OuterXMLData = xmlData.GetElementsByTagName("container").Item(LoopIndex).OuterXml
                        End If
                        OuterXML.LoadXml(OuterXMLData)
                    Catch ex As Exception
                        Log("Error in GetTracks extracting an item element with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                    Try
                        Tracks(StartIndex + LoopIndex, 0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                        Tracks(StartIndex + LoopIndex, 0) = ""
                    End Try
                    If SearchItemObject Then
                        Try
                            Tracks(StartIndex + LoopIndex, 6) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("id").Value
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 6) = ""
                        End Try
                        Try
                            Tracks(StartIndex + LoopIndex, 7) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("parentID").Value
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 7) = ""
                        End Try
                    Else
                        Try
                            Tracks(StartIndex + LoopIndex, 6) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 6) = ""
                        End Try
                        Try
                            Tracks(StartIndex + LoopIndex, 7) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 7) = ""
                        End Try
                    End If
                    Try
                        Tracks(StartIndex + LoopIndex, 3) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                        If InStr(Tracks(StartIndex + LoopIndex, 3), "-sonos-dock:RINCON_") <> 0 Then
                            ' these are URIs for Wireless dock players. Remove this, it will be properly added later
                            ' to be removed part = x-sonos-dock:RINCON_000E5860905A01400/ or 38 characters
                            Tracks(StartIndex + LoopIndex, 3) = Tracks(StartIndex + LoopIndex, 3).ToString.Remove(0, 38)
                        End If
                    Catch ex As Exception
                        Tracks(StartIndex + LoopIndex, 3) = ""
                    End Try
                    If Not NamesOnly Then
                        Try
                            Tracks(StartIndex + LoopIndex, 1) = OuterXML.GetElementsByTagName("upnp:album").Item(0).InnerText
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 1) = ""
                        End Try
                        Try
                            Tracks(StartIndex + LoopIndex, 2) = OuterXML.GetElementsByTagName("dc:creator").Item(0).InnerText
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 2) = ""
                        End Try
                        Try
                            'Tracks(StartIndex + LoopIndex, 4) = DecodeURI(OuterXML.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText)
                            Tracks(StartIndex + LoopIndex, 4) = OuterXML.GetElementsByTagName("upnp:albumArtURI").Item(0).InnerText ' removed in v3.1.0.19 because of failure to retrieve art
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 4) = ""
                        End Try
                        Try
                            Tracks(StartIndex + LoopIndex, 5) = OuterXML.GetElementsByTagName("upnp:class").Item(0).InnerText
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 5) = ""
                        End Try
                        Try
                            Tracks(StartIndex + LoopIndex, 10) = OuterXML.GetElementsByTagName("upnp:originalTrackNumber").Item(0).InnerText
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 10) = ""
                        End Try
                        Try
                            Tracks(StartIndex + LoopIndex, 11) = OuterXML.GetElementsByTagName("upnp:genre").Item(0).InnerText
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 11) = ""
                        End Try
                        Try
                            Tracks(StartIndex + LoopIndex, 12) = OuterXML.GetElementsByTagName("r:albumArtist").Item(0).InnerText
                        Catch ex As Exception
                            Tracks(StartIndex + LoopIndex, 12) = ""
                        End Try
                    End If
                Next
            Catch ex As Exception
                Log("GetTracks for zoneplayer = " & ZoneName & ": Error " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            If g_bDebug Then Log("GetTracks for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
            wait(0.25)
        Loop
        GetTracks = Tracks
    End Function

    Public Function BuildTrackDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim Track(12)
        Dim ConnectionString As String
        Dim WaitIndex As Integer = 0
        BuildTrackDB = ""
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildTrackDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If ObjectID = "" Then ObjectID = "A:TRACKS"
        Try

            Dim InArg(5)
            Dim OutArg(3)

            InArg(0) = ObjectID
            InArg(1) = "BrowseDirectChildren"
            InArg(2) = "*"
            InArg(3) = 0
            InArg(4) = 10
            InArg(5) = ""

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildTrackDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                'Cleanup and close the connection
                Try
                    If Not IsNothing(objConn) Then objConn.Close()
                Catch ex1 As Exception
                End Try
                Try
                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                Catch ex1 As Exception
                End Try
                Try
                    If Not IsNothing(objConn) Then objConn.Dispose()
                Catch ex1 As Exception
                End Try
                Exit Function
            End Try

            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)
            If g_bDebug Then Log("BuildTrackDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
            StartIndex = 0
            Do
                InArg(3) = CStr(StartIndex)
                InArg(4) = MaxNbrOfUPNPObjects
                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in BuildTrackDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try

                Try
                    xmlData.LoadXml(OutArg(0))
                    NumberReturned = OutArg(1)

                    Try
                        Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("item")
                        For Each Item As XmlNode In ItemList
                            OuterXML.LoadXml(Item.OuterXml)

                            Try
                                Track(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                            Catch ex As Exception
                                Track(0) = ""
                            End Try
                            Try
                                Track(1) = OuterXML.GetElementsByTagName("upnp:album").Item(0).InnerText
                            Catch ex As Exception
                                Track(1) = ""
                            End Try
                            Try
                                Track(2) = OuterXML.GetElementsByTagName("dc:creator").Item(0).InnerText
                            Catch ex As Exception
                                Track(2) = ""
                            End Try
                            Try
                                Track(3) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                            Catch ex As Exception
                                Track(3) = ""
                            End Try
                            Try
                                Track(6) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("id").Value
                            Catch ex As Exception
                                Track(6) = ""
                            End Try
                            Try
                                Track(7) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("parentID").Value
                                If ObjectID <> "A:TRACKS" Then
                                    Track(7) = "A:TRACKS" ' overwrite this, this is a DB for an iPOD docked in a WD100
                                End If
                            Catch ex As Exception
                                Track(7) = ""
                            End Try
                            Try
                                Track(10) = OuterXML.GetElementsByTagName("upnp:originalTrackNumber").Item(0).InnerText
                            Catch ex As Exception
                                Track(10) = ""
                            End Try
                            Try
                                Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                                objCommand.Parameters.Add("@Name", Data.DbType.String).Value = Track(0)
                                objCommand.Parameters.Add("@Album", Data.DbType.String).Value = Track(1)
                                objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = Track(2)
                                objCommand.Parameters.Add("@URI", Data.DbType.String).Value = Track(3)
                                objCommand.Parameters.Add("@Id", Data.DbType.String).Value = Track(6)
                                objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = Track(7)
                                objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = Val(Track(10))
                                objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = Track(11)
                                objCommand.ExecuteNonQuery()
                            Catch ex As Exception
                                Log("BuildTrackDB unable write this record with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                                Log("Name = " & Track(0), LogType.LOG_TYPE_ERROR)
                                Log("Album = " & Track(1), LogType.LOG_TYPE_ERROR)
                                Log("Artist = " & Track(2), LogType.LOG_TYPE_ERROR)
                                Log("URI = " & Track(3), LogType.LOG_TYPE_ERROR)
                                Log("Id = " & Track(6), LogType.LOG_TYPE_ERROR)
                                Log("ParentID = " & Track(7), LogType.LOG_TYPE_ERROR)
                                Log("TrackNo = " & Track(10), LogType.LOG_TYPE_ERROR)
                                Log("Genre = " & Track(11), LogType.LOG_TYPE_ERROR)
                            End Try
                        Next
                    Catch ex As Exception
                        Log("Error in BuildTrackDB extracting an item element with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                    End Try
                Catch ex As Exception
                    Log("Error in BuildTrackDB loading XML for zoneplayer = " & ZoneName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try

                StartIndex = StartIndex + NumberReturned
                If StartIndex >= TotalMatches Then
                    Exit Do
                End If
                wait(0.25)
                If g_bDebug Then Log("BuildTrackDB for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
            Loop
        Catch ex As Exception
            Log("Error in BuildTrackDB for zoneplayer = " & ZoneName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildTrackDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building TrackDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)
    End Function

    Public Function BuildArtistDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim InArg(5)
        Dim OutArg(3)
        Dim Artist(6)
        Dim ConnectionString As String
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildArtistDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildArtistDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try
        If ObjectID = "" Then ObjectID = "A:ALBUMARTIST"
        Try
            InArg(0) = ObjectID
            InArg(1) = "BrowseDirectChildren"
            InArg(2) = "*"
            InArg(3) = 0
            InArg(4) = 10
            InArg(5) = ""

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildArtistDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                'Cleanup and close the connection
                Try
                    If Not IsNothing(objConn) Then objConn.Close()
                Catch ex1 As Exception
                End Try
                Try
                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                Catch ex1 As Exception
                End Try
                Try
                    If Not IsNothing(objConn) Then objConn.Dispose()
                Catch ex1 As Exception
                End Try
                Exit Function
            End Try
            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)
            If g_bDebug Then Log("BuildArtistDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

            StartIndex = 0
            ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
            '<container id="A:ALBUMARTIST/!!!" parentID="A:ALBUMARTIST" restricted="true">
            '   <dc:title>!!!</dc:title>
            '   <upnp:class>object.container.person.musicArtist</upnp:class>
            '   <res protocolInfo="x-rincon-playlist:*:*:*">x-rincon-playlist:RINCON_000E5824C3B001400#A:ALBUMARTIST/!!!</res></container>
            '<container id="A:ALBUMARTIST/10cc" parentID="A:ALBUMARTIST" restricted="true">
            '   <dc:title>10cc</dc:title>
            '   <upnp:class>object.container.person.musicArtist</upnp:class>
            '   <res protocolInfo="x-rincon-playlist:*:*:*">x-rincon-playlist:RINCON_000E5824C3B001400#A:ALBUMARTIST/10cc</res></container>
            Do
                InArg(3) = CStr(StartIndex)
                InArg(4) = MaxNbrOfUPNPObjects
                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in BuildArtistDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try

                Try
                    xmlData.LoadXml(OutArg(0))
                    NumberReturned = OutArg(1)
                    Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("container")
                    For Each Item As XmlNode In ItemList
                        OuterXML.LoadXml(Item.OuterXml)
                        Try
                            Artist(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                        Catch ex As Exception
                            Artist(0) = ""
                        End Try
                        Try
                            Artist(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                        Catch ex As Exception
                            Artist(1) = ""
                        End Try
                        Try
                            Artist(3) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                        Catch ex As Exception
                            Artist(3) = ""
                        End Try
                        Try
                            Artist(4) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                            If ObjectID <> "A:ALBUMARTIST" Then
                                Artist(4) = "A:ALBUMARTIST" ' overwrite this, this is a DB for an iPOD docked in a WD100
                            End If
                        Catch ex As Exception
                            Artist(4) = ""
                        End Try
                        Try
                            Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                            objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All Tracks"
                            objCommand.Parameters.Add("@Album", Data.DbType.String).Value = "All Albums"
                            objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = Artist(0)
                            objCommand.Parameters.Add("@URI", Data.DbType.String).Value = Artist(1)
                            objCommand.Parameters.Add("@Id", Data.DbType.String).Value = Artist(3)
                            objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = Artist(4)
                            objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                            objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                            objCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            Log("BuildArtistDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Log("Name = All Tracks", LogType.LOG_TYPE_ERROR)
                            Log("Album = All Albums", LogType.LOG_TYPE_ERROR)
                            Log("Artist = " & Artist(0), LogType.LOG_TYPE_ERROR)
                            Log("URI = " & Artist(1), LogType.LOG_TYPE_ERROR)
                            Log("Id = " & Artist(3), LogType.LOG_TYPE_ERROR)
                            Log("ParentID = " & Artist(4), LogType.LOG_TYPE_ERROR)
                        End Try
                    Next
                Catch ex As Exception
                    Log("Error in BuildArtistDB loading XML for zoneplayer = " & ZoneName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try
                StartIndex = StartIndex + NumberReturned
                If StartIndex >= TotalMatches Then
                    Exit Do
                End If
                wait(0.25)
                If g_bDebug Then Log("BuildArtistDB writing Artists for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
            Loop
        Catch ex As Exception
            Log("Error in BuildArtistDB for zoneplayer = " & ZoneName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildArtistDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building ArtistDB for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
    End Function

    Public Function BuildAlbumDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim InArg(5)
        Dim OutArg(3)
        Dim Album(9)
        Dim ConnectionString As String
        BuildAlbumDB = ""
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildAlbumDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try
        If ObjectID = "" Then ObjectID = "A:ALBUM"
        Try

            InArg(0) = ObjectID
            InArg(1) = "BrowseDirectChildren"
            InArg(2) = "*"
            InArg(3) = "0"
            InArg(4) = "10" ' this is really weird. the WD100 gives an error when you request 1
            InArg(5) = ""

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildAlbumDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                'Cleanup and close the connection
                Try
                    If Not IsNothing(objConn) Then objConn.Close()
                Catch ex1 As Exception
                End Try
                Try
                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                Catch ex1 As Exception
                End Try
                Try
                    If Not IsNothing(objConn) Then objConn.Dispose()
                Catch ex1 As Exception
                End Try
                Exit Function
            End Try
            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)
            If g_bDebug Then Log("BuildAlbumDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
            '
            StartIndex = 0
            '
            '<DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
            '<container id="A:ALBUM/(MIA)%3a%20The%20Complete%20Anthology" parentID="A:ALBUM" restricted="true">
            '   <dc:title>(MIA): The Complete Anthology</dc:title>
            '   <upnp:class>object.container.album.musicAlbum</upnp:class>
            '   <res protocolInfo="x-rincon-playlist:*:*:*">x-rincon-playlist:RINCON_000E5824C3B001400#A:ALBUM/(MIA)%3a%20The%20Complete%20Anthology</res>
            '   <dc:creator>Germs</dc:creator>
            '   <upnp:albumArtURI>/getaa?u=x-file-cifs%3a%2f%2fMAXIE%2fOurMusic%2fThe%2520Germs%2f(MIA)-%2520The%2520Complete%2520Anthology%2f01-Forming.mp3&amp;v=331</upnp:albumArtURI></container>
            '<container id="A:ALBUM/(What&apos;s%20the%20Story)%20Morning%20Glory%3f" parentID="A:ALBUM" restricted="true">
            '   <dc:title>(What&apos;s the Story) Morning Glory?</dc:title>
            Dim RecordIndex As Integer = 0
            Do
                InArg(3) = CStr(StartIndex)
                InArg(4) = MaxNbrOfUPNPObjects
                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in BuildAlbumDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try

                Try
                    xmlData.LoadXml(OutArg(0))
                    NumberReturned = OutArg(1)
                    Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("container")
                    For Each Item As XmlNode In ItemList
                        OuterXML.LoadXml(Item.OuterXml)
                        Try
                            Album(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                        Catch ex As Exception
                            Album(0) = ""
                        End Try
                        Try
                            Album(1) = OuterXML.GetElementsByTagName("dc:creator").Item(0).InnerText
                        Catch ex As Exception
                            Album(1) = ""
                        End Try
                        Try
                            Album(2) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                        Catch ex As Exception
                            Album(2) = ""
                        End Try
                        Try
                            Album(5) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                        Catch ex As Exception
                            Album(5) = ""
                        End Try
                        Try
                            Album(6) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                            If ObjectID <> "A:ALBUM" Then
                                Album(6) = "A:ALBUM" ' overwrite this, this is a DB for an iPOD docked in a WD100
                            End If
                        Catch ex As Exception
                            Album(6) = ""
                        End Try
                        'Log( "BuildAlbumDB found Album/Artist = " & Album(0) & " / " & Album(1) & "at Index = " & RecordIndex.Message, LogType.LOG_TYPE_ERROR)
                        RecordIndex = RecordIndex + 1
                        Try
                            Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                            objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All Tracks"
                            objCommand.Parameters.Add("@Album", Data.DbType.String).Value = Album(0)
                            objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = Album(1)
                            objCommand.Parameters.Add("@URI", Data.DbType.String).Value = Album(2)
                            objCommand.Parameters.Add("@Id", Data.DbType.String).Value = Album(5)
                            objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = Album(6)
                            objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                            objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                            objCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            Log("BuildAlbumDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Log("Name = All Tracks", LogType.LOG_TYPE_ERROR)
                            Log("Album = " & Album(0), LogType.LOG_TYPE_ERROR)
                            Log("Artist = " & Album(1), LogType.LOG_TYPE_ERROR)
                            Log("URI = " & Album(2), LogType.LOG_TYPE_ERROR)
                            Log("Id = " & Album(5), LogType.LOG_TYPE_ERROR)
                            Log("ParentID = " & Album(6), LogType.LOG_TYPE_ERROR)
                        End Try
                    Next
                Catch ex As Exception
                    Log("Error in BuildAlbumDB loading XML for zoneplayer = " & ZoneName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try
                StartIndex = StartIndex + NumberReturned
                If StartIndex >= TotalMatches Then
                    Exit Do
                End If
                wait(0.25)
                If g_bDebug Then Log("BuildAlbumDB for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
            Loop
        Catch ex As Exception
            Log("Error in BuildAlbumDB for zoneplayer = " & ZoneName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildAlbumDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building AlbumDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildRadioStationDB(ByVal DatabasePath As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim Station(6)
        Dim InArg(5)
        Dim OutArg(3)
        Dim ConnectionString As String
        Dim I As Integer = 0
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildRadioStationDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildRadioStationDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try
        '

        If g_bDebug Then Log("BuildRadioStationDB called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO)

        InArg(0) = "R:0/0"                  ' Object ID     String If I use R:0 I get list of categories with R:0/0 I get the favorites
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 1                        ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildRadioStationDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildRadioStationDB found " & OutArg(2).ToString & " Radiostations for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        If g_bDebug Then Log("BuildRadioStationDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)
        '
        StartIndex = 0
        '
        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
        '<item id="R:0/0/0" parentID="R:0/0" restricted="false">
        '   <dc:title>104.5 | KFOG (AAA)</dc:title>
        '   <upnp:class>object.item.audioItem.audioBroadcast</upnp:class>
        '   <res protocolInfo="x-rincon-mp3radio:*:*:*">x-sonosapi-stream:s32698?sid=254&amp;flags=32</res></item>
        '<item id="R:0/0/1" parentID="R:0/0" restricted="false">
        '   <dc:title>99.7 | Norcal&apos;s Hit Music Ch. - KMVQ-HD2 (Top 40-Pop)</dc:title>
        Do
            InArg(3) = CStr(StartIndex)
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildRadioStationDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("item")
                For Each Item As XmlNode In ItemList
                    OuterXML.LoadXml(Item.OuterXml)
                    Try
                        Station(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                        Station(0) = ""
                    End Try
                    Try
                        Station(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                    Catch ex As Exception
                        Station(1) = ""
                    End Try
                    Try
                        Station(3) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("id").Value
                    Catch ex As Exception
                        Station(3) = ""
                    End Try
                    Try
                        Station(4) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("parentID").Value
                    Catch ex As Exception
                        Station(4) = ""
                    End Try
                    Try
                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All RadioStations"
                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = "RadioStation: " & Station(0)
                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = Station(1)
                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = Station(3)
                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = Station(4)
                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = I
                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                        objCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Log("BuildRadioStationDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("Name = All RadioStations", LogType.LOG_TYPE_ERROR)
                        Log("Album = ", LogType.LOG_TYPE_ERROR)
                        Log("Artist = " & Station(0), LogType.LOG_TYPE_ERROR)
                        Log("URI = " & Station(1), LogType.LOG_TYPE_ERROR)
                        Log("ArtURI = " & Station(7), LogType.LOG_TYPE_ERROR)
                        Log("Id = " & Station(3), LogType.LOG_TYPE_ERROR)
                        Log("ParentID = " & Station(4), LogType.LOG_TYPE_ERROR)
                    End Try
                Next
            Catch ex As Exception
                Log("Error in BuildRadioStationDB loading XML for zoneplayer = " & ZoneName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                'Cleanup and close the connection
                Try
                    If Not IsNothing(objConn) Then objConn.Close()
                Catch ex1 As Exception
                End Try
                Try
                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                Catch ex1 As Exception
                End Try
                Try
                    If Not IsNothing(objConn) Then objConn.Dispose()
                Catch ex1 As Exception
                End Try
                Exit Function
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            wait(0.25)
            If g_bDebug Then Log("BuildRadioStationDB for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
        Loop
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildRadioStationDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building RadioStationDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildPlaylistDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim PlayList(6)
        Dim InArg(5)
        Dim OutArg(3)
        Dim ConnectionString As String
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildPlaylistDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildPlaylistDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildPlaylistDB called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        If ObjectID = "" Then ObjectID = "A:PLAYLISTS"
        InArg(0) = ObjectID                 ' Object ID     String
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildPlaylistDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        ' OutArg (0) = Result
        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        If g_bDebug Then Log("BuildPlaylistDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        StartIndex = 0

        Do

            InArg(3) = CStr(StartIndex)
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildPlaylistDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try

            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("container")
                For Each Item As XmlNode In ItemList
                    OuterXML.LoadXml(Item.OuterXml)
                    Try
                        PlayList(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                        PlayList(0) = ""
                    End Try
                    Try
                        PlayList(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                    Catch ex As Exception
                        PlayList(1) = ""
                    End Try
                    Try
                        PlayList(3) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                        If ObjectID <> "A:PLAYLISTS" Then
                            PlayList(1) = PlayList(3) ' overwrite this, this is a DB for an iPOD docked in a WD100
                        End If
                    Catch ex As Exception
                        PlayList(3) = ""
                    End Try
                    Try
                        PlayList(4) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                        If ObjectID <> "A:PLAYLISTS" Then
                            PlayList(4) = "A:PLAYLISTS" ' overwrite this, this is a DB for an iPOD docked in a WD100
                        End If
                    Catch ex As Exception
                        PlayList(4) = ""
                    End Try
                    Try
                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All Playlists"
                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = PlayList(0)
                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = PlayList(1)
                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = PlayList(3)
                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = PlayList(4)
                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                        objCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Log("BuildPlaylistDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("Name = All Playlists", LogType.LOG_TYPE_ERROR)
                        Log("Album = ", LogType.LOG_TYPE_ERROR)
                        Log("Artist = " & PlayList(0), LogType.LOG_TYPE_ERROR)
                        Log("URI = " & PlayList(1), LogType.LOG_TYPE_ERROR)
                        Log("Id = " & PlayList(3), LogType.LOG_TYPE_ERROR)
                        Log("ParentID = " & PlayList(4), LogType.LOG_TYPE_ERROR)
                    End Try
                Next
            Catch ex As Exception
                Log("BuildPlaylistDB for zoneplayer = " & ZoneName & ": Error " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            wait(0.25)
            If g_bDebug Then Log("BuildPlaylistDB for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
        Loop
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildPlaylistDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building BuildPlaylistDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildSonosPlaylistDB(ByVal DatabasePath As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim PlayList(6)
        Dim InArg(5)
        Dim OutArg(3)
        Dim ConnectionString As String
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildSonosPlaylistDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildSonosPlaylistDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildSonosPlaylistDB called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO)

        InArg(0) = "SQ:"                    ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildSonosPlaylistDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        If g_bDebug Then Log("BuildSonosPlaylistDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        StartIndex = 0
        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
        '<container id="SQ:1" parentID="SQ:" restricted="true">
        '   <dc:title>Best Of Damien Mrley</dc:title>
        '   <res protocolInfo="file:*:audio/mpegurl:*">file:///jffs/settings/savedqueues.rsq#1</res>
        '   <upnp:class>object.container.playlistContainer</upnp:class></container>
        '<container id="SQ:33" parentID="SQ:" restricted="true">
        '   <dc:title>Dirk</dc:title>
        '   <res protocolInfo="file:*:audio/mpegurl:*">file:///jffs/settings/savedqueues.rsq#33</res>
        '   <upnp:class>object.container.playlistContainer</upnp:class></container>

        Do
            InArg(3) = CStr(StartIndex)
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildSonosPlaylistDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try

            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("container")
                For Each Item As XmlNode In ItemList
                    OuterXML.LoadXml(Item.OuterXml)
                    Try
                        PlayList(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                        PlayList(0) = ""
                    End Try
                    Try
                        PlayList(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                    Catch ex As Exception
                        PlayList(1) = ""
                    End Try
                    Try
                        PlayList(3) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    Catch ex As Exception
                        PlayList(3) = ""
                    End Try
                    Try
                        PlayList(4) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                    Catch ex As Exception
                        PlayList(4) = ""
                    End Try
                    Try
                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All Playlists"
                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = PlayList(0)
                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = PlayList(1)
                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = PlayList(3)
                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = PlayList(4)
                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                        objCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Log("BuildSonosPlaylistDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("Name = All Playlists", LogType.LOG_TYPE_ERROR)
                        Log("Album = ", LogType.LOG_TYPE_ERROR)
                        Log("Artist = " & PlayList(0), LogType.LOG_TYPE_ERROR)
                        Log("URI = " & PlayList(1), LogType.LOG_TYPE_ERROR)
                        Log("Id = " & PlayList(3), LogType.LOG_TYPE_ERROR)
                        Log("ParentID = " & PlayList(4), LogType.LOG_TYPE_ERROR)
                    End Try
                Next
            Catch ex As Exception
                Log("BuildSonosPlaylistDB the unpacking didn't work for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            wait(0.25)
            If g_bDebug Then Log("BuildSonosPlaylistDB for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
        Loop
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildSonosPlaylistDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building BuildSonosPlaylistDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildGenreDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim Genre(6)
        Dim InArg(5)
        Dim OutArg(3)
        Dim ConnectionString As String
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildGenreDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildGenreDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildGenreDB called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        If ObjectID = "" Then ObjectID = "A:GENRE"
        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildGenreDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        If g_bDebug Then Log("BuildGenreDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        StartIndex = 0

        Do
            InArg(3) = CStr(StartIndex)
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildGenreDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try

            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("container")
                For Each Item As XmlNode In ItemList
                    OuterXML.LoadXml(Item.OuterXml)
                    Try
                        Genre(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                    End Try
                    Try
                        Genre(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                    Catch ex As Exception
                    End Try
                    Try
                        Genre(3) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value ' use this to find tracks linked to this Genre
                        If ObjectID <> "A:GENRE" Then
                            Genre(1) = Genre(3)
                        End If
                    Catch ex As Exception
                    End Try
                    Try
                        Genre(4) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                        If ObjectID <> "A:GENRE" Then
                            Genre(4) = "A:GENRE" ' overwrite this, this is a DB for an iPOD docked in a WD100
                        End If
                    Catch ex As Exception
                    End Try
                    Try
                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All Genres"
                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = Genre(1)
                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = Genre(3)
                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = Genre(4)
                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = Genre(0)
                        objCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Log("BuildGenreDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("Name = All Genres", LogType.LOG_TYPE_ERROR)
                        Log("Album = ", LogType.LOG_TYPE_ERROR)
                        Log("Artist = ", LogType.LOG_TYPE_ERROR)
                        Log("URI = " & Genre(1), LogType.LOG_TYPE_ERROR)
                        Log("Id = " & Genre(3), LogType.LOG_TYPE_ERROR)
                        Log("ParentID = " & Genre(4), LogType.LOG_TYPE_ERROR)
                        Log("Genre = " & Genre(0), LogType.LOG_TYPE_ERROR)
                    End Try
                Next
            Catch ex As Exception
                Log("BuildGenreDB the unpacking didn't work for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            wait(0.25)
            If g_bDebug Then Log("BuildGenreDB for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
        Loop
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildGenreDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building BuildGenreDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildAudioBooksDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim AudioBook(3)
        Dim InArg(5)
        Dim OutArg(3)
        Dim ConnectionString As String
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildAudioBooksDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildAudioBooksDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildAudioBooksDB called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO)

        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildAudioBooksDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        Log("BuildAudioBooksDB found " & TotalMatches.ToString & " for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        StartIndex = 0
        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
        '<item id="7:0" parentID="7" restricted="false">
        '   <dc:title>Revelation: the Word of Promise Audio Bible: NKJV: Free Audiobook(Unabridged)</dc:title>
        '   <upnp:class>object.item.audioItem</upnp:class>
        '   <res protocolInfo="http-get:*:audio/wav:*">x-sonos-dock:RINCON_000E5860905A01400/7:0</res>
        '   <upnp:albumArtURI></upnp:albumArtURI>
        '</item>
        '</DIDL-Lite>
        Do
            InArg(3) = CStr(StartIndex)
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildAudioBooksDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try

            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("item")
                For Each Item As XmlNode In ItemList
                    OuterXML.LoadXml(Item.OuterXml)
                    Try
                        AudioBook(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                    End Try
                    Try
                        AudioBook(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                    Catch ex As Exception
                    End Try
                    Try
                        AudioBook(2) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("id").Value ' use this to find tracks linked to this Genre
                    Catch ex As Exception
                    End Try
                    Try
                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = AudioBook(0)
                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = AudioBook(2)
                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = AudioBook(2)
                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = "A:AUDIOBOOK"
                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                        objCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Log("BuildAudioBooksDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("Name = " & AudioBook(0), LogType.LOG_TYPE_ERROR)
                        Log("Album = ", LogType.LOG_TYPE_ERROR)
                        Log("Artist = ", LogType.LOG_TYPE_ERROR)
                        Log("URI = " & AudioBook(2), LogType.LOG_TYPE_ERROR)
                        Log("Id = " & AudioBook(2), LogType.LOG_TYPE_ERROR)
                        Log("ParentID = A:AUDIOBOOK", LogType.LOG_TYPE_ERROR)
                        Log("Genre = ", LogType.LOG_TYPE_ERROR)
                    End Try
                Next
            Catch ex As Exception
                Log("BuildAudioBooksDB the unpacking didn't work for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            wait(0.25)
            If g_bDebug Then Log("BuildAudioBooksDB for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
        Loop
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildAudioBooksDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        Log("Finished Building BuildAudioBooksDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildPodCastDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlContainerData As XmlDocument = New XmlDocument
        Dim xmlPodcastData As XmlDocument = New XmlDocument
        Dim PodCastFamily(2)
        Dim Podcast(4)
        Dim InArg(5)
        Dim OutArg(3)
        Dim ConnectionString As String
        Dim ContainerLoopIndex As Integer = 0
        Dim ContainerStartIndex As Integer = 0
        Dim NumberReturnedContainers As Integer = 0
        Dim TotalContainerMatches As Integer = 0
        Dim PodcastLoopIndex As Integer = 0
        Dim PodcastStartIndex As Integer = 0
        Dim NumberReturnedPodcasts As Integer = 0
        Dim TotalPodcastMatches As Integer = 0
        Dim OuterXMLData As String
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildPodCastDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildPodCastDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildPodCastDB called for ZonePlayer = " & ZoneName, LogType.LOG_TYPE_INFO)

        InArg(0) = ObjectID                 ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 10                       ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildPodCastDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        NumberReturnedContainers = OutArg(1)
        TotalContainerMatches = OutArg(2)
        Log("BuildPodCastDB found " & TotalContainerMatches.ToString & " Podcast families for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

        ContainerStartIndex = 0


        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
        '<container id="8:0:5" parentID="8" restricted="false">
        '   <dc:title>Beautiful Creatures</dc:title>
        '   <upnp:class>object.container</upnp:class>
        '</container>
        '<container id="8:1:5" parentID="8" restricted="false">
        '   <dc:title>Eden - A free audiobook by Phil Rossi</dc:title>
        '   <upnp:class>object.container</upnp:class>
        '</container>
        '<container id="8:2:5" parentID="8" restricted="false">
        '   <dc:title>Motley Fool Money</dc:title>
        '   <upnp:class>object.container</upnp:class>
        '</container>
        '<container id="8:3:5" parentID="8" restricted="false">
        '   <dc:title>The Random House Audio Podcast</dc:title>
        '   <upnp:class>object.container</upnp:class>
        '</container>
        '</DIDL-Lite>

        Do
            InArg(0) = ObjectID                 ' Object ID     String 
            InArg(3) = ContainerStartIndex
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildPodCastDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try

            Try
                xmlContainerData.LoadXml(OutArg(0))
                NumberReturnedContainers = OutArg(1)
                Dim ItemList As XmlNodeList = xmlContainerData.GetElementsByTagName("container")
                For Each Item As XmlNode In ItemList
                    OuterXML.LoadXml(Item.OuterXml)
                    Try
                        PodCastFamily(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                    End Try

                    Try
                        PodCastFamily(1) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    Catch ex As Exception
                    End Try
                    Try
                        PodCastFamily(2) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                    Catch ex As Exception
                    End Try
                    Try
                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = PodCastFamily(0)
                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = "All Podcasts"
                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = PodCastFamily(1)
                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = PodCastFamily(1)
                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = "A:PODCASTFAMILY" 'PodCastFamily(2)
                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                        objCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Log("BuildPodCastDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("Name = " & PodCastFamily(0), LogType.LOG_TYPE_ERROR)
                        Log("Album = All Podcasts", LogType.LOG_TYPE_ERROR)
                        Log("Artist = ", LogType.LOG_TYPE_ERROR)
                        Log("URI = " & PodCastFamily(1), LogType.LOG_TYPE_ERROR)
                        Log("Id = " & PodCastFamily(1), LogType.LOG_TYPE_ERROR)
                        Log("ParentID = A:PODCASTFAMILY", LogType.LOG_TYPE_ERROR) ' & PodCastFamily(2)))
                        Log("Genre = ", LogType.LOG_TYPE_ERROR)
                    End Try
                    If PodCastFamily(1) <> "" Then
                        ' now go pick up the individual podcasts
                        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
                        '<item id="8:0:5:0" parentID="8:0:5" restricted="false">
                        '   <dc:title>Beautiful Creatures - Audiobook Sample</dc:title>
                        '   <upnp:class>object.item.audioItem</upnp:class>
                        '   <res protocolInfo="http-get:*:audio/wav:*">x-sonos-dock:RINCON_000E5860905A01400/8:0:5:0</res>
                        '   <upnp:albumArtURI></upnp:albumArtURI>
                        '</item>
                        '</DIDL-Lite>
                        InArg(0) = PodCastFamily(1)         ' Object ID     String 
                        InArg(3) = 0                        ' Index         UI4
                        InArg(4) = 10                       ' Count         UI4

                        Try
                            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                        Catch ex As Exception
                            Log("ERROR in BuildPodCastDB/Browse1 for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                            'Cleanup and close the connection
                            Try
                                If Not IsNothing(objConn) Then objConn.Close()
                            Catch ex1 As Exception
                            End Try
                            Try
                                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                            Catch ex1 As Exception
                            End Try
                            Try
                                If Not IsNothing(objConn) Then objConn.Dispose()
                            Catch ex1 As Exception
                            End Try
                            Exit Function
                        End Try

                        NumberReturnedPodcasts = OutArg(1)
                        TotalPodcastMatches = OutArg(2)
                        Log("BuildPodCastDB found " & TotalPodcastMatches.ToString & " Podcast entries for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

                        PodcastStartIndex = 0

                        Do
                            InArg(3) = PodcastStartIndex
                            InArg(4) = MaxNbrOfUPNPObjects
                            Try
                                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                            Catch ex As Exception
                                Log("ERROR in BuildPodCastDB/Browse2 for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                Exit Do
                            End Try

                            Try
                                xmlPodcastData.LoadXml(OutArg(0))
                                NumberReturnedPodcasts = OutArg(1)
                                For PodcastLoopIndex = 0 To NumberReturnedPodcasts - 1
                                    Try
                                        OuterXMLData = xmlPodcastData.GetElementsByTagName("item").Item(PodcastLoopIndex).OuterXml
                                        OuterXML.LoadXml(OuterXMLData)
                                    Catch ex As Exception
                                        Log("Error in BuildPodCastDB1 extracting an item element with error " & ex.Message, LogType.LOG_TYPE_ERROR)
                                    End Try
                                    Try
                                        Podcast(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                                    Catch ex As Exception
                                    End Try

                                    Try
                                        Podcast(1) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("id").Value
                                    Catch ex As Exception
                                        Podcast(1) = ""
                                    End Try
                                    Try
                                        Podcast(2) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("parentID").Value
                                    Catch ex As Exception
                                        Podcast(2) = ""
                                    End Try
                                    Try
                                        Podcast(3) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                                    Catch ex As Exception
                                        Podcast(3) = ""
                                    End Try
                                    Try
                                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = Podcast(0)
                                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = ""
                                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = ""
                                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = Podcast(1)
                                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = Podcast(1)
                                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = "A:PODCAST" 'Podcast(2)
                                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                                        objCommand.ExecuteNonQuery()
                                    Catch ex As Exception
                                        Log("BuildPodCastDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                                        Log("Name = " & Podcast(0), LogType.LOG_TYPE_ERROR)
                                        Log("Album = ", LogType.LOG_TYPE_ERROR)
                                        Log("Artist = ", LogType.LOG_TYPE_ERROR)
                                        Log("URI = " & Podcast(1), LogType.LOG_TYPE_ERROR)
                                        Log("Id = " & Podcast(1), LogType.LOG_TYPE_ERROR)
                                        Log("ParentID = A:PODCAST" & Podcast(2), LogType.LOG_TYPE_ERROR)
                                        Log("Genre = ", LogType.LOG_TYPE_ERROR)
                                    End Try
                                Next
                            Catch ex As Exception
                                Log("BuildPodCastDB1 the unpacking didn't work for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                                Exit Do
                            End Try
                            PodcastStartIndex = PodcastStartIndex + NumberReturnedPodcasts
                            If PodcastStartIndex >= TotalPodcastMatches Then
                                Exit Do
                            End If
                            If g_bDebug Then Log("BuildPodCastDB for zoneplayer = " & ZoneName & ". Count =" & PodcastStartIndex.ToString, LogType.LOG_TYPE_INFO)
                        Loop
                    End If
                Next
            Catch ex As Exception
                Log("BuildPodCastDB the unpacking didn't work for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            ContainerStartIndex = ContainerStartIndex + NumberReturnedContainers
            If ContainerStartIndex >= TotalContainerMatches Then
                Exit Do
            End If
            wait(0.25)
            If g_bDebug Then Log("BuildPodCastDB for zoneplayer = " & ZoneName & ". Count =" & ContainerStartIndex.ToString, LogType.LOG_TYPE_INFO)
        Loop
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildPodCastDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        Log("Finished Building BuildPodCastDB for zoneplayer = " & ZoneName & "", LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildSpecificObjectDB(ByVal DatabasePath As String, ObjectID As String, ObjectDBName As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim ObjectInfo(6)
        Dim InArg(5)
        Dim OutArg(3)
        Dim ConnectionString As String
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer
        Dim TotalMatches As Integer
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildSpecificObjectDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildSpecificObjectDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildSpecificObjectDB called for ZonePlayer = " & ZoneName & " with ObjectID = " & ObjectID & " and ObjectDBName = " & ObjectDBName, LogType.LOG_TYPE_INFO)

        InArg(0) = ObjectID                 ' Object ID     String If I use R:0 I get list of categories with R:0/0 I get the favorites
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 1                        ' Count         UI4
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildSpecificObjectDB/Browse for zoneplayer = " & ZoneName & " with ObjectDBName = " & ObjectDBName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        If g_bDebug Then Log("BuildSpecificObjectDB found " & OutArg(2).ToString & " Objects for ZonePlayer - " & ZoneName & " with ObjectDBName = " & ObjectDBName, LogType.LOG_TYPE_INFO)

        NumberReturned = OutArg(1)
        TotalMatches = OutArg(2)
        '
        StartIndex = 0
        '
        ' <DIDL-Lite xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:upnp="urn:schemas-upnp-org:metadata-1-0/upnp/" xmlns:r="urn:schemas-rinconnetworks-com:metadata-1-0/" xmlns="urn:schemas-upnp-org:metadata-1-0/DIDL-Lite/">
        '<item id="R:0/0/0" parentID="R:0/0" restricted="false">
        '   <dc:title>104.5 | KFOG (AAA)</dc:title>
        '   <upnp:class>object.item.audioItem.audioBroadcast</upnp:class>
        '   <res protocolInfo="x-rincon-mp3radio:*:*:*">x-sonosapi-stream:s32698?sid=254&amp;flags=32</res></item>
        '<item id="R:0/0/1" parentID="R:0/0" restricted="false">
        '   <dc:title>99.7 | Norcal&apos;s Hit Music Ch. - KMVQ-HD2 (Top 40-Pop)</dc:title>
        Do
            InArg(3) = CStr(StartIndex)
            InArg(4) = MaxNbrOfUPNPObjects
            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildSpecificObjectDB/Browse for zoneplayer = " & ZoneName & " with ObjectDBName = " & ObjectDBName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            Try
                xmlData.LoadXml(OutArg(0))
                NumberReturned = OutArg(1)
                Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("item")
                For Each Item As XmlNode In ItemList
                    OuterXML.LoadXml(Item.OuterXml)
                    Try
                        ObjectInfo(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText
                    Catch ex As Exception
                        ObjectInfo(0) = ""
                    End Try
                    Try
                        ObjectInfo(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                    Catch ex As Exception
                        ObjectInfo(1) = ""
                    End Try
                    Try
                        ObjectInfo(3) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("id").Value
                    Catch ex As Exception
                        ObjectInfo(3) = ""
                    End Try
                    Try
                        ObjectInfo(4) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("parentID").Value
                    Catch ex As Exception
                        ObjectInfo(4) = ""
                    End Try
                    Try
                        Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                        objCommand.Parameters.Add("@Name", Data.DbType.String).Value = ObjectDBName
                        objCommand.Parameters.Add("@Album", Data.DbType.String).Value = ""
                        objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = ObjectInfo(0)
                        objCommand.Parameters.Add("@URI", Data.DbType.String).Value = ObjectInfo(1)
                        objCommand.Parameters.Add("@Id", Data.DbType.String).Value = ObjectInfo(3)
                        objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = ObjectInfo(4)
                        objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                        objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                        objCommand.ExecuteNonQuery()
                    Catch ex As Exception
                        Log("BuildSpecificObjectDB unable write this record for zoneplayer = " & ZoneName & " with ObjectDBName = " & ObjectDBName & " and error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                        Log("Name = " & ObjectDBName, LogType.LOG_TYPE_ERROR)
                        Log("Album = ", LogType.LOG_TYPE_ERROR)
                        Log("Artist = " & ObjectInfo(0), LogType.LOG_TYPE_ERROR)
                        Log("URI = " & ObjectInfo(1), LogType.LOG_TYPE_ERROR)
                        Log("Id = " & ObjectInfo(3), LogType.LOG_TYPE_ERROR)
                        Log("ParentID = " & ObjectInfo(4), LogType.LOG_TYPE_ERROR)
                    End Try
                Next
            Catch ex As Exception
                Log("Error in BuildSpecificObjectDB loading XML for zoneplayer = " & ZoneName & " with ObjectDBName = " & ObjectDBName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                'Cleanup and close the connection
                Try
                    If Not IsNothing(objConn) Then objConn.Close()
                Catch ex1 As Exception
                End Try
                Try
                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                Catch ex1 As Exception
                End Try
                Try
                    If Not IsNothing(objConn) Then objConn.Dispose()
                Catch ex1 As Exception
                End Try
                Exit Function
            End Try
            StartIndex = StartIndex + NumberReturned
            If StartIndex >= TotalMatches Then
                Exit Do
            End If
            wait(0.25)
            If g_bDebug Then Log("BuildSpecificObjectDB for zoneplayer = " & ZoneName & " with ObjectDBName = " & ObjectDBName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
        Loop
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildSpecificObjectDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        If g_bDebug Then Log("Finished Building BuildSpecificObjectDB for zoneplayer = " & ZoneName & " with ObjectDBName = " & ObjectDBName, LogType.LOG_TYPE_INFO)

    End Function

    Public Function BuildiPODTrackDB(ByVal DatabasePath As String, ByVal ObjectID As String)
        Dim xmlData As XmlDocument = New XmlDocument
        Dim AlbumxmlData As XmlDocument = New XmlDocument
        Dim TrackxmlData As XmlDocument = New XmlDocument
        Dim InArg(5)
        Dim OutArg(3)
        Dim ArtistInArg(5)
        Dim ArtistOutArg(3)
        Dim AlbumInArg(5)
        Dim AlbumOutArg(3)
        Dim TrackInArg(5)
        Dim TrackOutArg(3)
        Dim Artist(6)
        Dim TrackInfo(4) ' in the sequence Artist, album, track, albumartist?
        Dim ConnectionString As String
        Dim StartIndex As Integer = 0
        Dim NumberReturned As Integer = 0
        Dim TotalMatches As Integer = 0
        Dim AlbumLoopIndex As Integer = 0
        Dim AlbumStartIndex As Integer = 0
        Dim AlbumNumberReturned As Integer = 0
        Dim AlbumTotalMatches As Integer = 0
        Dim TrackLoopIndex As Integer = 0
        Dim TrackStartIndex As Integer = 0
        Dim TrackNumberReturned As Integer = 0
        Dim TrackTotalMatches As Integer = 0
        Dim ArtistObjectID As String = ""
        Dim AlbumObjectID As String = ""
        Dim TrackObjectID As String = ""
        Dim OuterXML As XmlDocument = New XmlDocument

        BuildiPODTrackDB = ""

        ConnectionString = DBConnectionString & DatabasePath
        '
        Dim objConn As SQLiteConnection = Nothing
        Dim SQLtransaction As SQLiteTransaction = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
            SQLtransaction = objConn.BeginTransaction()
        Catch ex As Exception
            Log("BuildiPODTrackDB unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Function
        End Try

        Try
            InArg(0) = ObjectID
            InArg(1) = "BrowseDirectChildren"
            InArg(2) = "*"
            InArg(3) = 0
            InArg(4) = 10
            InArg(5) = ""

            Try
                ContentDirectory.InvokeAction("Browse", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in BuildiPODTrackDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                'Cleanup and close the connection
                Try
                    If Not IsNothing(objConn) Then objConn.Close()
                Catch ex1 As Exception
                End Try
                Try
                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                Catch ex1 As Exception
                End Try
                Try
                    If Not IsNothing(objConn) Then objConn.Dispose()
                Catch ex1 As Exception
                End Try
                Exit Function
            End Try
            NumberReturned = OutArg(1)
            TotalMatches = OutArg(2)
            If g_bDebug Then Log("BuildiPODTrackDB found " & TotalMatches.ToString & " Artists for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

            StartIndex = 0

            Do
                InArg(3) = StartIndex
                InArg(4) = MaxNbrOfUPNPObjects
                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in BuildiPODTrackDB/Browse for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try

                Try
                    xmlData.LoadXml(OutArg(0))
                    NumberReturned = OutArg(1)
                    Dim ItemList As XmlNodeList = xmlData.GetElementsByTagName("container")
                    For Each Item As XmlNode In ItemList
                        OuterXML.LoadXml(Item.OuterXml)
                        ArtistObjectID = ""
                        Try
                            Artist(0) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText ' this is the Artist title
                            TrackInfo(0) = Artist(0)
                        Catch ex As Exception
                            Artist(0) = ""
                        End Try
                        Try
                            Artist(1) = OuterXML.GetElementsByTagName("res").Item(0).InnerText
                        Catch ex As Exception
                            Artist(1) = ""
                        End Try
                        Try
                            Artist(3) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                            AlbumObjectID = Artist(3)
                        Catch ex As Exception
                            Artist(3) = ""
                            AlbumObjectID = Artist(3)
                        End Try
                        Try
                            Artist(4) = OuterXML.GetElementsByTagName("container").Item(0).Attributes("parentID").Value
                            If ObjectID <> "A:ALBUMARTIST" Then
                                Artist(4) = "A:ALBUMARTIST" ' overwrite this, this is a DB for an iPOD docked in a WD100
                            End If
                        Catch ex As Exception
                            Artist(4) = ""
                        End Try
                        Try
                            Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                            objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All Tracks"
                            objCommand.Parameters.Add("@Album", Data.DbType.String).Value = "All Albums"
                            objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = Artist(0)
                            objCommand.Parameters.Add("@URI", Data.DbType.String).Value = Artist(3)
                            objCommand.Parameters.Add("@Id", Data.DbType.String).Value = Artist(3)
                            objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = Artist(4)
                            objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                            objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                            objCommand.ExecuteNonQuery()
                        Catch ex As Exception
                            Log("BuildArtistDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                            Log("Name = All Tracks", LogType.LOG_TYPE_ERROR)
                            Log("Album = All Albums", LogType.LOG_TYPE_ERROR)
                            Log("Artist = " & Artist(0), LogType.LOG_TYPE_ERROR)
                            Log("URI = " & Artist(3), LogType.LOG_TYPE_ERROR)
                            Log("Id = " & Artist(3), LogType.LOG_TYPE_ERROR)
                            Log("ParentID = " & Artist(4), LogType.LOG_TYPE_ERROR)
                        End Try
                        ' now that we just stored the Album, now let's loop through this for all 
                        If AlbumObjectID <> "" Then
                            AlbumInArg(0) = AlbumObjectID
                            AlbumInArg(1) = "BrowseDirectChildren"
                            AlbumInArg(2) = "*"
                            AlbumInArg(3) = 0
                            AlbumInArg(4) = 10
                            AlbumInArg(5) = ""
                            Try
                                ContentDirectory.InvokeAction("Browse", AlbumInArg, AlbumOutArg)
                            Catch ex As Exception
                                Log("ERROR in BuildiPODTrackDB/Album for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                Log("ERROR in BuildiPODTrackDB/Album for zoneplayer = " & ZoneName & " with ObjectID = " & ArtistObjectID, LogType.LOG_TYPE_ERROR)
                                If Not IsNothing(objConn) Then
                                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                                    objConn.Close()
                                End If
                                Exit Function
                            End Try
                            AlbumNumberReturned = AlbumOutArg(1)
                            AlbumTotalMatches = AlbumOutArg(2)
                            If g_bDebug Then Log("BuildiPODTrackDB found " & AlbumTotalMatches.ToString & " albums for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

                            AlbumStartIndex = 0

                            Do
                                AlbumInArg(3) = AlbumStartIndex
                                AlbumInArg(4) = MaxNbrOfUPNPObjects
                                Try
                                    ContentDirectory.InvokeAction("Browse", AlbumInArg, AlbumOutArg)
                                    'Log( "BuildiPODTrackDB/Album1 for zoneplayer = " & ZoneName & " with ObjectID = " & AlbumObjectID & " and StartIndex = " & AlbumStartIndex.Message, LogType.LOG_TYPE_ERROR)
                                Catch ex As Exception
                                    Log("ERROR in BuildiPODTrackDB/Album1 for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                    Log("ERROR in BuildiPODTrackDB/Album1 for zoneplayer = " & ZoneName & " with ObjectID = " & AlbumObjectID, LogType.LOG_TYPE_ERROR)
                                    Exit Do
                                End Try

                                Try
                                    AlbumxmlData.LoadXml(AlbumOutArg(0))
                                    AlbumNumberReturned = AlbumOutArg(1)
                                    Dim ContainerList As XmlNodeList = AlbumxmlData.GetElementsByTagName("container")
                                    For Each Container As XmlNode In ContainerList
                                        OuterXML.LoadXml(Container.OuterXml)
                                        TrackObjectID = ""
                                        Try
                                            TrackInfo(1) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText ' this is the Album title
                                        Catch ex As Exception
                                            TrackInfo(1) = ""
                                        End Try
                                        Try
                                            TrackObjectID = OuterXML.GetElementsByTagName("container").Item(0).Attributes("id").Value
                                        Catch ex As Exception
                                            TrackObjectID = ""
                                        End Try
                                        If TrackInfo(1) <> "All Tracks" Then
                                            ' store this as a record. You have here Artist + Album
                                            ' now save the whole record
                                            Try
                                                Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                                                objCommand.Parameters.Add("@Name", Data.DbType.String).Value = "All Tracks"
                                                objCommand.Parameters.Add("@Album", Data.DbType.String).Value = TrackInfo(1)
                                                objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = TrackInfo(0)
                                                objCommand.Parameters.Add("@URI", Data.DbType.String).Value = TrackObjectID
                                                objCommand.Parameters.Add("@Id", Data.DbType.String).Value = TrackObjectID
                                                objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = "A:ALBUM"
                                                objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = 0
                                                objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                                                objCommand.ExecuteNonQuery()
                                            Catch ex As Exception
                                                Log("BuildiPODTrackDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                                                Log("Name = All Tracks", LogType.LOG_TYPE_ERROR)
                                                Log("Album = " & TrackInfo(1), LogType.LOG_TYPE_ERROR)
                                                Log("Artist = " & TrackInfo(0), LogType.LOG_TYPE_ERROR)
                                                Log("URI = " & TrackObjectID, LogType.LOG_TYPE_ERROR)
                                                Log("Id = " & TrackObjectID, LogType.LOG_TYPE_ERROR)
                                                Log("ParentID = A:ALBUM", LogType.LOG_TYPE_ERROR)
                                            End Try
                                        End If
                                        If TrackObjectID <> "" Then
                                            TrackInArg(0) = TrackObjectID
                                            TrackInArg(1) = "BrowseDirectChildren"
                                            TrackInArg(2) = "*"
                                            TrackInArg(3) = 0
                                            TrackInArg(4) = 10
                                            TrackInArg(5) = ""
                                            Try
                                                ContentDirectory.InvokeAction("Browse", TrackInArg, TrackOutArg)
                                            Catch ex As Exception
                                                Log("ERROR in BuildiPODTrackDB/Track for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                                Log("ERROR in BuildiPODTrackDB/Track for zoneplayer = " & ZoneName & " with ObjectID = " & AlbumObjectID, LogType.LOG_TYPE_ERROR)
                                                'Cleanup and close the connection
                                                Try
                                                    If Not IsNothing(objConn) Then objConn.Close()
                                                Catch ex1 As Exception
                                                End Try
                                                Try
                                                    If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
                                                Catch ex1 As Exception
                                                End Try
                                                Try
                                                    If Not IsNothing(objConn) Then objConn.Dispose()
                                                Catch ex1 As Exception
                                                End Try
                                                Exit Function
                                            End Try
                                            TrackNumberReturned = TrackOutArg(1)
                                            TrackTotalMatches = TrackOutArg(2)
                                            If g_bDebug Then Log("BuildiPODTrackDB found " & TrackTotalMatches.ToString & " Tracks for ZonePlayer - " & ZoneName, LogType.LOG_TYPE_INFO)

                                            TrackStartIndex = 0

                                            Do
                                                TrackInArg(3) = TrackStartIndex
                                                TrackInArg(4) = MaxNbrOfUPNPObjects
                                                Try
                                                    ContentDirectory.InvokeAction("Browse", TrackInArg, TrackOutArg)
                                                Catch ex As Exception
                                                    Log("ERROR in BuildiPODTrackDB/Track1 for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                                                    Log("ERROR in BuildiPODTrackDB/Track1 for zoneplayer = " & ZoneName & " with ObjectID = " & AlbumObjectID, LogType.LOG_TYPE_ERROR)
                                                    Exit Do
                                                End Try

                                                Try
                                                    TrackxmlData.LoadXml(TrackOutArg(0))
                                                    TrackNumberReturned = TrackOutArg(1)
                                                    Dim TrackList As XmlNodeList = xmlData.GetElementsByTagName("item")
                                                    For Each Track As XmlNode In TrackList
                                                        OuterXML.LoadXml(Track.OuterXml)
                                                        Try
                                                            TrackInfo(2) = OuterXML.GetElementsByTagName("dc:title").Item(0).InnerText ' this is the track title
                                                        Catch ex As Exception
                                                            TrackInfo(2) = ""
                                                        End Try
                                                        Try
                                                            TrackInfo(3) = OuterXML.GetElementsByTagName("item").Item(0).Attributes("id").Value
                                                        Catch ex As Exception
                                                            TrackInfo(3) = ""
                                                        End Try
                                                        ' now save the whole record
                                                        Try
                                                            Dim objCommand As New SQLiteCommand("INSERT INTO Tracks (Name, Album, Artist, URI, Id, ParentID, TrackNo, Genre ) VALUES (@Name, @Album, @Artist, @URI, @Id, @ParentID, @TrackNo, @Genre);", objConn)
                                                            objCommand.Parameters.Add("@Name", Data.DbType.String).Value = TrackInfo(2)
                                                            objCommand.Parameters.Add("@Album", Data.DbType.String).Value = TrackInfo(1)
                                                            objCommand.Parameters.Add("@Artist", Data.DbType.String).Value = TrackInfo(0)
                                                            objCommand.Parameters.Add("@URI", Data.DbType.String).Value = TrackInfo(3)
                                                            objCommand.Parameters.Add("@Id", Data.DbType.String).Value = TrackInfo(3)
                                                            objCommand.Parameters.Add("@ParentID", Data.DbType.String).Value = "A:TRACKS"
                                                            objCommand.Parameters.Add("@TrackNo", Data.DbType.Int16).Value = (TrackLoopIndex + 1)
                                                            objCommand.Parameters.Add("@Genre", Data.DbType.String).Value = ""
                                                            objCommand.ExecuteNonQuery()
                                                        Catch ex As Exception
                                                            Log("BuildiPODTrackDB unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
                                                            Log("Name = " & TrackInfo(2), LogType.LOG_TYPE_ERROR)
                                                            Log("Album = " & TrackInfo(1), LogType.LOG_TYPE_ERROR)
                                                            Log("Artist = " & TrackInfo(0), LogType.LOG_TYPE_ERROR)
                                                            Log("URI = " & TrackInfo(3), LogType.LOG_TYPE_ERROR)
                                                            Log("Id = " & TrackInfo(3), LogType.LOG_TYPE_ERROR)
                                                            Log("ParentID = A:TRACKS", LogType.LOG_TYPE_ERROR)
                                                        End Try
                                                    Next
                                                Catch ex As Exception
                                                    Log("Error in BuildiPODTrackDB loading XML for zoneplayer = " & ZoneName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                                                    Exit Do
                                                End Try
                                                TrackStartIndex = TrackStartIndex + TrackNumberReturned
                                                If TrackStartIndex >= TrackTotalMatches Then
                                                    Exit Do
                                                End If
                                                If g_bDebug Then Log("BuildiPODTrackDB1 writing Artists for zoneplayer = " & ZoneName & ". Count =" & TrackStartIndex.ToString, LogType.LOG_TYPE_INFO)
                                            Loop
                                        End If
                                    Next
                                Catch ex As Exception
                                    Log("Error in BuildiPODTrackDB2 loading XML for zoneplayer = " & ZoneName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                                    Exit Do
                                End Try
                                AlbumStartIndex = AlbumStartIndex + AlbumNumberReturned
                                If AlbumStartIndex >= AlbumTotalMatches Then
                                    Exit Do
                                End If
                                If g_bDebug Then Log("BuildiPODTrackDB3 writing Artists for zoneplayer = " & ZoneName & ". Count =" & AlbumStartIndex.ToString, LogType.LOG_TYPE_INFO)
                            Loop
                        End If
                    Next
                Catch ex As Exception
                    Log("Error in BuildiPODTrackDB loading XML for zoneplayer = " & ZoneName & ". XML = " & OutArg(0), LogType.LOG_TYPE_ERROR)
                    Exit Do
                End Try
                StartIndex = StartIndex + NumberReturned
                If StartIndex >= TotalMatches Then
                    Exit Do
                End If
                wait(0.25)
                If g_bDebug Then Log("BuildiPODTrackDB writing Artists for zoneplayer = " & ZoneName & ". Count =" & StartIndex.ToString, LogType.LOG_TYPE_INFO)
            Loop
        Catch ex As Exception
            Log("Error in BuildiPODTrackDB for zoneplayer = " & ZoneName & ". Error =" & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If SQLtransaction IsNot Nothing Then SQLtransaction.Commit()
        Catch ex As Exception
            Log("BuildiPODTrackDB to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
        Try
            If SQLtransaction IsNot Nothing Then SQLtransaction.Dispose()
        Catch ex As Exception
        End Try
        Try
            If Not IsNothing(objConn) Then objConn.Dispose()
        Catch ex As Exception
        End Try
        Log("Finished Building BuildiPODTrackDB for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
    End Function

    Private Sub CreateTrackDatabase(ByVal DatabasePath As String, ByVal strConnectionString As String)
        'Declare the main SQLite data access objects
        Dim objConn As SQLiteConnection = Nothing
        Dim objCommand As System.Data.SQLite.SQLiteCommand
        If g_bDebug Then Log("CreateTrackDatabase called with " & DatabasePath & " and " & strConnectionString & " for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            If File.Exists(DatabasePath) Then
                File.Delete(DatabasePath)
            End If
        Catch ex As Exception
        End Try
        Try
            'Create a new database connection
            'Note - use New=True to create a new database
            objConn = New SQLiteConnection(strConnectionString & ";Version=3;New=True;")
            'Open the connection
            objConn.Open()
            'Create a new SQL command
            objCommand = objConn.CreateCommand()
            'Setup and execute the command SQL to create a new table
            objCommand.CommandText = "CREATE TABLE Tracks (Name Text,Album Text,Artist Text,URI Text,Id text PRIMARY KEY,ParentID text,TrackNo integer ,Genre text);"
            objCommand.ExecuteNonQuery()
        Catch ex As Exception
            Log("Error in CreateTrackDatabase for zoneplayer = " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
    End Sub

    Public Sub CreateRadioStationDatabase(ByVal DatabasePath As String, ByVal strConnectionString As String, ByVal DeleteDB As Boolean)
        'Declare the main SQLite data access objects
        Dim objConn As SQLiteConnection = Nothing
        Dim objCommand As System.Data.SQLite.SQLiteCommand
        If g_bDebug Then Log("CreateRadioStationDatabase called with " & DatabasePath & " and " & strConnectionString & " for zoneplayer = " & ZoneName & " and DeleteDB = " & DeleteDB.ToString, LogType.LOG_TYPE_INFO)
        Try
            If File.Exists(DatabasePath) Then
                If Not DeleteDB Then Exit Sub
                File.Delete(DatabasePath)
            End If
        Catch ex As Exception
        End Try
        Try
            'Create a new database connection
            'Note - use New=True to create a new database
            objConn = New SQLiteConnection(strConnectionString & ";Version=3; New=True;")
            'Open the connection
            objConn.Open()
            'Create a new SQL command
            objCommand = objConn.CreateCommand()
            'Setup and execute the command SQL to create a new table
            objCommand.CommandText = "CREATE TABLE RadioStations (Name Text Primary Key,URI Text,MetaData text,Type text);"
            objCommand.ExecuteNonQuery()
        Catch ex As Exception
            Log("Error in CreateRadioStationDatabase for zoneplayer = " & ZoneName & " with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            'Cleanup and close the connection
            If Not IsNothing(objConn) Then objConn.Close()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UpdateRadioStationsInfo(ByVal CurrentURI As String, ByVal CurrentURIMetaData As String, ByVal CurrentType As String)
        If SuperDebug Then
            Log("UpdateRadioStationsInfo called for Zone = " & ZoneName & " with CurrentURI " & CurrentURI & " and CurrentURIMetaData = " & CurrentURIMetaData, LogType.LOG_TYPE_INFO)
        ElseIf g_bDebug Then
            Log("UpdateRadioStationsInfo called for Zone = " & ZoneName & " with CurrentURI " & CurrentURI, LogType.LOG_TYPE_INFO)
        End If

        Dim xmlData As XmlDocument = New XmlDocument
        Dim StationName As String = ""
        Dim ConnectionString As String

        Try
            xmlData.LoadXml(CurrentURIMetaData)
        Catch ex As Exception
            If g_bDebug Then Log("Error in UpdateRadioStationsInfo loading XML metadata. XML = " & CurrentURIMetaData, LogType.LOG_TYPE_ERROR)
            If g_bDebug Then Log("Error in UpdateRadioStationsInfo unable to retrieve RadioStationName", LogType.LOG_TYPE_ERROR)
            Exit Sub
        End Try

        ' Now retrieve the name

        Try
            StationName = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
        Catch ex As Exception
            If g_bDebug Then Log("Error in UpdateRadioStationsInfo unable to retrieve RadioStationName. XML = " & CurrentURIMetaData, LogType.LOG_TYPE_INFO)
            Exit Sub
        End Try

        xmlData = Nothing

        StationName = CurrentType & " - " & StationName

        ConnectionString = DBConnectionString & CurrentAppPath & RadioStationsDBPath
        CreateRadioStationDatabase(CurrentAppPath & RadioStationsDBPath, ConnectionString, False)

        '
        Dim objConn As SQLiteConnection = Nothing

        Try
            'Create a new database connection
            objConn = New SQLiteConnection(ConnectionString)
            'Open the connection
            objConn.Open()
        Catch ex As Exception
            Log("UpdateRadioStationsInfo unable to open DB for zoneplayer = " & ZoneName & " with error- " & ex.Message, LogType.LOG_TYPE_ERROR)
            'Cleanup and close the connection
            Try
                If Not IsNothing(objConn) Then objConn.Close()
            Catch ex1 As Exception
            End Try
            Try
                If Not IsNothing(objConn) Then objConn.Dispose()
            Catch ex1 As Exception
            End Try
            Exit Sub
        End Try

        Dim QueryResult As Integer = 0
        Try
            Dim objCommand As New SQLiteCommand("INSERT OR REPLACE INTO RadioStations (Name, URI, MetaData, Type ) VALUES (@Name, @URI, @MetaData, @Type);", objConn)
            objCommand.Parameters.Add("@Name", Data.DbType.String).Value = StationName
            objCommand.Parameters.Add("@URI", Data.DbType.String).Value = CurrentURI
            objCommand.Parameters.Add("@MetaData", Data.DbType.String).Value = CurrentURIMetaData
            objCommand.Parameters.Add("@Type", Data.DbType.String).Value = CurrentType
            QueryResult = objCommand.ExecuteNonQuery()
        Catch ex As Exception
            Log("UpdateRadioStationsInfo unable write this record for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
            Log("Name = " & StationName, LogType.LOG_TYPE_ERROR)
            Log("URI = " & CurrentURI, LogType.LOG_TYPE_ERROR)
            Log("MetaData = " & CurrentURIMetaData, LogType.LOG_TYPE_ERROR)
            Log("Type = " & CurrentType, LogType.LOG_TYPE_ERROR)
        End Try

        If g_bDebug Then Log("UpdateRadioStationsInfo updated RadioStation = " & StationName & " for Zone = " & ZoneName & " with QueryResult = " & QueryResult.ToString, LogType.LOG_TYPE_INFO)

        Try
            If objConn IsNot Nothing Then objConn.Close()
        Catch ex As Exception
            If g_bDebug Then Log("Error in UpdateRadioStationsInfo to close DB for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Try
            If objConn IsNot Nothing Then objConn.Dispose()
        Catch ex As Exception
            If g_bDebug Then Log("Error in UpdateRadioStationsInfo to dispose of the DB-Object for zoneplayer = " & ZoneName & " with error - " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        objConn = Nothing
        If g_bDebug Then Log("UpdateRadioStationsInfo done for Zone = " & ZoneName, LogType.LOG_TYPE_INFO)

    End Sub

    Private Sub FindDockedPlayerSettings()
        MyAutoPlayLinkedZones = GetAutoplayLinkedZones()
        MyAutoPlayRoomUUID = GetAutoplayRoomUUID()
        MyAutoPlayVolume = GetAutoplayVolume()
        If g_bDebug Then Log("FindDockedPlayerSettings for Player = " & ZoneName & " and PlayAutoLinkedZones = " & MyAutoPlayLinkedZones.ToString & " and AutoPlayZone = " & MyAutoPlayRoomUUID & " and AutoVolume = " & MyAutoPlayVolume.ToString, LogType.LOG_TYPE_INFO)
    End Sub

    Private Sub ResetDockedPlayerSettings()
        If g_bDebug Then Log("ResetDockedPlayerSettings for Player = " & ZoneName, LogType.LOG_TYPE_INFO)
        MyAutoPlayLinkedZones = False
        MyAutoPlayRoomUUID = """"
        MyAutoPlayVolume = 0
        MyDockediPodPlayerName = ""
    End Sub

    Private Function MakeiPodBrowseable() As Boolean
        MakeiPodBrowseable = False
        Dim InArg(0)
        Dim OutArg(0)

        Try
            ContentDirectory.InvokeAction("GetBrowseable", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in MakeiPodBrowseable for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try

        ReDim InArg(0)
        ReDim OutArg(0)

        If Not OutArg(0) Then
            ' the player is not browseable, set it to true. This may interupt the music stream
            InArg(0) = True                      ' isBrowseable Boolean
            Try
                ContentDirectory.InvokeAction("SetBrowseable", InArg, OutArg)
            Catch ex As Exception
                Log("ERROR in MakeiPodBrowseable for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                ' if not succesful really no use to continue
                Exit Function
            End Try
        End If
        MakeiPodBrowseable = True
    End Function

    Public Function BuildDockedPlayerDatabase(ByVal DatabasePath As String, ByVal PlayerName As String)
        Dim ConnectionString
        BuildDockedPlayerDatabase = "NOK"
        If g_bDebug Then Log("BuildDockedPlayerDatabase for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName & " with " & DatabasePath, LogType.LOG_TYPE_INFO)
        If ContentDirectory Is Nothing Then
            If g_bDebug Then Log("BuildDockedPlayerDatabase No player connected for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName, LogType.LOG_TYPE_WARNING)
            Exit Function ' there is no iPod Docked
        End If

        '
        ' save the SettingsReplicationState in the inifile so we can detect when changes have occured
        Dim CurrentReplicationSettings As String
        CurrentReplicationSettings = GetStringIniFile("DockedSettingsReplicationState", ZoneName, "")
        WriteStringIniFile("Saved DockedSettingsReplicationState", ZoneName, CurrentReplicationSettings)
        ' first we need to pick up the ObjectId to browse from

        Dim xmlData As XmlDocument = New XmlDocument
        Dim NumberReturned As Integer
        Dim ObjectName As String
        Dim ObjectID(7)
        ObjectID = {"", "", "", "", "", "", "", ""}

        If Not MakeiPodBrowseable() Then Exit Function


        Dim InArg(5)
        Dim OutArg(3)
        InArg(0) = "0"                      ' Object ID     String 
        InArg(1) = "BrowseDirectChildren"   ' Browse Flag   String
        InArg(2) = "*"                      ' Filter        String
        InArg(3) = 0                        ' Index         UI4
        InArg(4) = 0                        ' Count         UI4  - 0 means all
        InArg(5) = ""                       ' Sort Criteria String

        Try
            ContentDirectory.InvokeAction("Browse", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in BuildDockedPlayerDatabase/Browse for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        'If g_bDebug Then Log( "BuildDockedPlayerDatabase found " & OutArg(1).ToString & " entries for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName & " with " & DatabasePath)
        Try
            NumberReturned = OutArg(1)
            InArg(4) = 1                        ' Count         UI4  - 0 means all
            For LoopIndex = 0 To NumberReturned - 1
                InArg(3) = LoopIndex            ' Index         UI4
                Try
                    ContentDirectory.InvokeAction("Browse", InArg, OutArg)
                Catch ex As Exception
                    Log("ERROR in BuildDockedPlayerDatabase/Browse for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                Try
                    xmlData.LoadXml(OutArg(0))
                Catch ex As Exception
                    Log("ERROR in BuildDockedPlayerDatabase/Browse loading XML for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
                    Exit Function
                End Try
                Try
                    ObjectName = xmlData.GetElementsByTagName("dc:title").Item(0).InnerText
                    If UCase(ObjectName) = "PLAYLISTS" Then
                        ObjectID(0) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    ElseIf UCase(ObjectName) = "ARTISTS" Then
                        ObjectID(1) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    ElseIf UCase(ObjectName) = "ALBUMS" Then
                        ObjectID(2) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    ElseIf UCase(ObjectName) = "GENRES" Then
                        ObjectID(3) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    ElseIf UCase(ObjectName) = "TRACKS" Then
                        ObjectID(4) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    ElseIf UCase(ObjectName) = "COMPOSERS" Then
                        ObjectID(5) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    ElseIf UCase(ObjectName) = "AUDIOBOOKS" Then
                        ObjectID(6) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    ElseIf UCase(ObjectName) = "PODCASTS" Then
                        ObjectID(7) = xmlData.GetElementsByTagName("container").Item(0).Attributes("id").Value
                    End If
                Catch ex As Exception
                End Try
            Next
        Catch ex As Exception
            Log("ERROR 2 in BuildDockedPlayerDatabase/Browse for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        xmlData = Nothing

        ' create DB and erase existing
        Try
            ConnectionString = DBConnectionString & DatabasePath
            'Call CreateTrackDatabase(strDatabasePath, ConnectionString)
            CreateTrackDatabase(DatabasePath, ConnectionString)
            'If MyMusicDBItems.Tracks And ObjectID(4) <> "" Then BuildTrackDB(DatabasePath, ObjectID(4))
            If MyMusicDBItems.Artists And ObjectID(1) <> "" Then BuildiPODTrackDB(DatabasePath, ObjectID(1)) 'BuildArtistDB(DatabasePath, ObjectID(1))
            'If MyMusicDBItems.Albums And ObjectID(2) <> "" Then BuildAlbumDB(DatabasePath, ObjectID(2))
            If MyMusicDBItems.Genres And ObjectID(3) <> "" Then BuildGenreDB(DatabasePath, ObjectID(3))
            If MyMusicDBItems.Playlists And ObjectID(0) <> "" Then BuildPlaylistDB(DatabasePath, ObjectID(0))
            If MyMusicDBItems.Audiobooks And ObjectID(6) <> "" Then BuildAudioBooksDB(DatabasePath, ObjectID(6))
            If MyMusicDBItems.Podcasts And ObjectID(7) <> "" Then BuildPodCastDB(DatabasePath, ObjectID(7))
            BuildSpecificObjectDB(DatabasePath, "FV:2", "Favorites")
            BuildDockedPlayerDatabase = "OK"
        Catch ex As Exception
            Log("An Error occurred in BuildDockedPlayerDatabase for zoneplayer = " & ZoneName & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        Log("BuildDockedPlayerDatabase Done  for zoneplayer = " & ZoneName & " and Docked Player Name = " & PlayerName, LogType.LOG_TYPE_INFO)

    End Function

    Private Function GetAlbumArtPath(ByVal AlbumURI As String, ByVal NextTrack As Boolean) As String
        If SuperDebug Then Log("GetAlbumArtPath called for zone " & ZoneName & " with AlbumURI = " & AlbumURI & " and NextTrack = " & NextTrack.ToString, LogType.LOG_TYPE_INFO)
        Dim AlbumArtImage As Image = Nothing
        If AlbumURI = NoArtPath Or AlbumURI = "" Then
            GetAlbumArtPath = NoArtPath
            If NextTrack Then
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            Exit Function
        End If
        ' prevent multiple saves and avoid this pesky GDI+ error
        If NextTrack Then
            If MyPreviousNextAlbumURI <> "" And MyPreviousNextAlbumURI = AlbumURI And MyPreviousAlbumArtPath <> "" Then
                GetAlbumArtPath = MyPreviousNextAlbumArtPath
                If SuperDebug Then Log("GetAlbumArtPath returned for Zone - " & ZoneName & " with AlbumURI = " & AlbumURI & " and cached returned path= " & GetAlbumArtPath, LogType.LOG_TYPE_INFO)
                Exit Function
            End If
        Else
            If MyPreviousAlbumURI <> "" And MyPreviousAlbumURI = AlbumURI And MyPreviousAlbumArtPath <> "" Then
                GetAlbumArtPath = MyPreviousAlbumArtPath
                If SuperDebug Then Log("GetAlbumArtPath returned for Zone - " & ZoneName & " with AlbumURI = " & AlbumURI & " and cached returned path= " & GetAlbumArtPath, LogType.LOG_TYPE_INFO)
                Exit Function
            End If
        End If
        AlbumArtImage = GetPicture(AlbumURI)
        If AlbumArtImage Is Nothing Then
            GetAlbumArtPath = NoArtPath
            If NextTrack Then
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            Exit Function
        End If
        If AlbumArtImage.Height = 0 Or AlbumArtImage.Width = 0 Then
            GetAlbumArtPath = NoArtPath
            If SuperDebug Then Log("GetAlbumArtPath encountered zero width/height picture for zonename = " & ZoneName, LogType.LOG_TYPE_INFO)
            AlbumArtImage.Dispose()
            GC.Collect()
            If NextTrack Then
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            Exit Function
        End If
        Dim ExtensionIndex As Integer = 0
        Dim ExtensionType As String = ".jpg"
        Dim ImageFormat As System.Drawing.Imaging.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg
        Try
            ' get the extension file type
            ExtensionIndex = AlbumURI.LastIndexOf(".")
            Dim TempExtensiontype As String = ""
            If ExtensionIndex <> -1 Then
                TempExtensiontype = AlbumURI.Substring(ExtensionIndex, AlbumURI.Length - ExtensionIndex)
                If UCase(TempExtensiontype) = ".PNG" Then
                    ExtensionType = ".png"
                    'ImageFormat = System.Drawing.Imaging.ImageFormat.Png
                    'If g_bDebug Then Log( "GetAlbumArtPath for zonename = " & ZoneName & " has set image to .PNG")
                End If
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in GetAlbumArtPath when searching for the file type with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        'Dim FilePath As String = ""
        Dim TempFilePath As String = ""
        Dim MyNextArtFileIndex As Integer = 0
        Dim MyArtFileIndex As Integer = 0

        If NextTrack Then
            MyNextArtFileIndex = GetIntegerIniFile(UDN, DeviceInfoIndex.diNextArtFileIndex.ToString, 0)
            MyNextArtFileIndex += 1
            WriteIntegerIniFile(UDN, DeviceInfoIndex.diNextArtFileIndex.ToString, MyNextArtFileIndex)
            'FilePath = CurrentAppPath & "\html\images\" & FileArtWorkPath & "NextCover" & UDN & "_" & MyNextArtFileIndex.ToString & ExtensionType 
            TempFilePath = "NextCover" & UDN & "*.*"
        Else
            MyArtFileIndex = GetIntegerIniFile(UDN, DeviceInfoIndex.diArtFileIndex.ToString, 0)
            MyArtFileIndex += 1
            WriteIntegerIniFile(UDN, DeviceInfoIndex.diArtFileIndex.ToString, MyArtFileIndex)
            'FilePath = CurrentAppPath & "\html\images\" & FileArtWorkPath & "Cover" & UDN & "_" & MyArtFileIndex.ToString & ExtensionType 
            TempFilePath = "Cover" & UDN & "*.*"
        End If
        Dim successflag As Boolean
        If ImRunningLocal Then
            Try
                ' let's try to delete the previous file
                Dim TempPath As String = FileArtWorkPath.Remove(FileArtWorkPath.Length - 1, 1) ' remove the "/' character
                For Each FileFound As String In Directory.GetFiles(CurrentAppPath & "/html/images/" & TempPath, TempFilePath)
                    File.Delete(FileFound)
                Next
            Catch ex As Exception
                If g_bDebug Then Log("Warning in GetAlbumArtPath when deleting the previous art work for device = " & ZoneName & " with Filename = " & TempFilePath & " and Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            End Try
        Else
            ' tobefixed dcor with new hs.deleteHTMLImage
            'successflag = hs.DeleteImageFile(FileFound)
            'If Not successflag Then
            'If g_bDebug Then Log("Error in GetAlbumArtPath when deleting the art work for device = " & ZoneName & " with Filename = " & FileFound, LogType.LOG_TYPE_ERROR)
            'End If
        End If

        Try
            If NextTrack Then
                successflag = hs.WriteHTMLImage(AlbumArtImage, FileArtWorkPath & "NextCover" & UDN & "_" & MyNextArtFileIndex.ToString & ExtensionType, True)
                If Not successflag Then
                    If g_bDebug Then Log("Error in GetAlbumArtPath when saving the next art work for device = " & ZoneName & " with Filename = " & FileArtWorkPath & "NextCover" & UDN & "_" & MyNextArtFileIndex.ToString & ExtensionType, LogType.LOG_TYPE_ERROR)
                End If
                GetAlbumArtPath = URLArtWorkPath & "NextCover" & UDN & "_" & MyNextArtFileIndex.ToString & ExtensionType
                MyPreviousNextAlbumArtPath = GetAlbumArtPath
                MyPreviousNextAlbumURI = AlbumURI
            Else
                successflag = hs.WriteHTMLImage(AlbumArtImage, FileArtWorkPath & "Cover" & UDN & "_" & MyArtFileIndex.ToString & ExtensionType, True)
                If Not successflag Then
                    If g_bDebug Then Log("Error in GetAlbumArtPath when saving the art work for device = " & ZoneName & " with Filename = " & FileArtWorkPath & "Cover" & UDN & "_" & MyArtFileIndex.ToString & ExtensionType, LogType.LOG_TYPE_ERROR)
                End If
                GetAlbumArtPath = URLArtWorkPath & "Cover" & UDN & "_" & MyArtFileIndex.ToString & ExtensionType
                MyPreviousAlbumArtPath = GetAlbumArtPath
                MyPreviousAlbumURI = AlbumURI
            End If
        Catch ex As Exception
            If NextTrack Then
                If g_bDebug Then Log("Error in GetAlbumArtPath storing artwork for device - " & ZoneName & " and path = " & URLArtWorkPath & "NextCover" & UDN & "_" & MyNextArtFileIndex.ToString & ExtensionType & " with error= " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyPreviousNextAlbumArtPath = ""
                MyPreviousNextAlbumURI = ""
            Else
                If g_bDebug Then Log("Error in GetAlbumArtPath storing artwork for device - " & ZoneName & " and path = " & URLArtWorkPath & "Cover" & UDN & "_" & MyArtFileIndex.ToString & ExtensionType & " with error= " & ex.Message, LogType.LOG_TYPE_ERROR)
                MyPreviousAlbumArtPath = ""
                MyPreviousAlbumURI = ""
            End If
            GetAlbumArtPath = NoArtPath
        End Try
        If SuperDebug Then Log("GetAlbumArtPath returned for Zone - " & ZoneName & " with AlbumURI = " & AlbumURI & " and returned path= " & GetAlbumArtPath, LogType.LOG_TYPE_INFO)
        Try
            AlbumArtImage.Dispose()
            GC.Collect()
        Catch ex As Exception
        End Try
    End Function

    Private Function GetPicture(ByVal url As String) As Image
        ' Get the picture at a given URL.
        Dim web_client As New WebClient()
        web_client.UseDefaultCredentials = True                         ' added 5/2/2019
        ' web_client.Credentials = CredentialCache.DefaultCredentials     ' added 5/2/2019
        GetPicture = Nothing
        Try
            url = Trim(url)
            If url = "" Then
                Return Nothing
                Exit Function
            End If
            If Not (url.ToLower().StartsWith("http://") Or url.ToLower().StartsWith("https://") Or url.ToLower().StartsWith("file:")) Then url = "http://" & url
            Dim image_stream As New MemoryStream(web_client.DownloadData(url))
            GetPicture = Image.FromStream(image_stream, True, True)
        Catch ex As Exception
            If g_bDebug Then Log("GetPicture called for Zone - " & ZoneName & " url= " & url.ToString & " caused error: " & ex.Message, LogType.LOG_TYPE_ERROR)
        Finally
            web_client.Dispose()
        End Try
    End Function

    Public Function ConvertSecondsToTimeFormat(ByVal Seconds As Integer) As String
        ConvertSecondsToTimeFormat = "00:00:00"
        If Seconds < 0 Then Exit Function
        Dim StartTime As Date = CDate("00:00:00")
        ConvertSecondsToTimeFormat = Format(DateAdd("s", CType(Seconds, Double), StartTime), "HH:mm:ss")
    End Function


    Public Sub SetBrowsable(Browse As Boolean)
        If g_bDebug Then Log("SetBrowsable called for zoneplayer = " & ZoneName & " and Browse flag = " & Browse.ToString, LogType.LOG_TYPE_INFO)
        If ContentDirectory Is Nothing Then Exit Sub
        Dim InArg(0)
        Dim OutArg(0)
        InArg(0) = Browse              ' isBrowseable Boolean
        Try
            ContentDirectory.InvokeAction("SetBrowseable", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SetBrowsable for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub CompareReplicationChanges(OldState As String, Newstate As String)
        ' The ReplicationString is something as follows
        'UserRadioUpdateID(Value = RINCON_000E5825227A01400, 13)
        'SavedQueuesUpdateID(Value = RINCON_000E5833F3CC01400, 116)
        'ShareListUpdateID(Value = RINCON_000E5824C3B001400, 330)
        'RecentlyPlayedUpdateID(Value = RINCON_000E5824C3B001400, 73)
        'RINCON_000E5825227A01400,9,	<< UserRadioUpdateID
        'RINCON_FFFFFFFFFFFF99999,0,
        'RINCON_000E5833F3CC01400,116,	<< SavedQueuesUpdateID
        'RINCON_000E5824C3B001400,330,	<< ShareListUpdateID
        'RINCON_000E5824C3B001400,44,
        'RINCON_000E5824C3B001400,18,
        'RINCON_000E5824C3B001400,384,	<< ServiceListVersion ?
        'RINCON_000E5824C3B001400,73,	<< RecentlyPlayedUpdateID
        'RINCON_000E5824C3B001400, 3
        Dim OldStates As String()
        Dim NewStates As String()
        If OldState = "" Then
            ' this is probably just after the plugin was activated
            WriteBooleanIniFile("SettingsReplicationState", "SonosSettingsHaveChanged", True)
            PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            Exit Sub
        End If
        OldStates = OldState.Split(",")
        NewStates = Newstate.Split(",")
        Try
            If OldStates(1) <> NewStates(1) Then ' UserRadioUpdateID -> called when the favorite radiostations have been changed
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for UserRadioUpdateID with Old = " & OldStates(1) & " and New = " & NewStates(1), LogType.LOG_TYPE_INFO)
                WriteBooleanIniFile("SettingsReplicationState", "SonosSettingsHaveChanged", True)
                PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            End If
            If OldStates(3) <> NewStates(3) Then ' Unknown
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for Unkown 1 with Old = " & OldStates(3) & " and New = " & NewStates(3), LogType.LOG_TYPE_INFO)
            End If
            If OldStates(5) <> NewStates(5) Then ' SavedQueuesUpdateID -> this is called when the saved queues change
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for SavedQueuesUpdateID with Old = " & OldStates(5) & " and New = " & NewStates(5), LogType.LOG_TYPE_INFO)
                If Not AnnouncementInProgress Then
                    WriteBooleanIniFile("SettingsReplicationState", "SonosSettingsHaveChanged", True)
                    PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
                End If
            End If
            If OldStates(7) <> NewStates(7) Then ' ShareListUpdateID -> this is when the music Index is reindexed
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for ShareListUpdateID with Old = " & OldStates(7) & " and New = " & NewStates(7), LogType.LOG_TYPE_INFO)
                WriteBooleanIniFile("SettingsReplicationState", "SonosSettingsHaveChanged", True)
                PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            End If
            If OldStates(9) <> NewStates(9) Then ' Unknown -> this changed when I deleted a music service
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for Unkown 4 with Old = " & OldStates(9) & " and New = " & NewStates(9), LogType.LOG_TYPE_INFO)
                PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            End If
            If OldStates(11) <> NewStates(11) Then ' Unkown -> this changed when I deleted a music service, together with UNkown 4
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for Alarms with Old = " & OldStates(11) & " and New = " & NewStates(11), LogType.LOG_TYPE_INFO)
                PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            End If
            If OldStates(13) <> NewStates(13) Then ' ServiceListVersion?
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for ServiceListVersion? with Old = " & OldStates(13) & " and New = " & NewStates(13), LogType.LOG_TYPE_INFO)
                PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            End If
            If OldStates(15) <> NewStates(15) Then ' RecentlyPlayedUpdateID
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for RecentlyPlayedUpdateID with Old = " & OldStates(15) & " and New = " & NewStates(15), LogType.LOG_TYPE_INFO)
                PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            End If
            If OldStates(17) <> NewStates(17) Then ' Unkown
                If g_bDebug Then Log("CompareReplicationChanges for zoneplayer " & ZoneName & " detected change for Unknown 8 with Old = " & OldStates(17) & " and New = " & NewStates(17), LogType.LOG_TYPE_INFO)
                PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ReplicationState)
            End If
        Catch ex As Exception
            If g_bDebug Then Log("Error in CompareReplicationChanges for zoneplayer " & ZoneName & " with Old = " & OldState & " and New = " & Newstate, LogType.LOG_TYPE_ERROR)
        End Try


    End Sub


    Public Function GetAutoplayLinkedZones() As Boolean
        GetAutoplayLinkedZones = False
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("GetAutoplayLinkedZones called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            DeviceProperties.InvokeAction("GetAutoplayLinkedZones", InArg, OutArg)
            GetAutoplayLinkedZones = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetAutoplayLinkedZones for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetAutoplayRoomUUID() As String
        GetAutoplayRoomUUID = False
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("GetAutoplayRoomUUID called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0)
            DeviceProperties.InvokeAction("GetAutoplayRoomUUID", InArg, OutArg)
            GetAutoplayRoomUUID = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetAutoplayRoomUUID for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Public Function GetAutoplayVolume() As Integer
        GetAutoplayVolume = False
        If DeviceStatus = "Offline" Then Exit Function
        If g_bDebug Then Log("GetAutoplayVolume called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        Try
            Dim InArg(0)
            Dim OutArg(0) ' ui2
            DeviceProperties.InvokeAction("GetAutoplayVolume", InArg, OutArg)
            GetAutoplayVolume = OutArg(0)
        Catch ex As Exception
            Log("ERROR in GetAutoplayVolume for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Function

    Private Sub GetAudioInputAttributes()
        If DeviceStatus = "Offline" Then Exit Sub
        If AudioIn Is Nothing Then Exit Sub
        If g_bDebug Then Log("GetAudioInputAttributes called for ZoneName " & ZoneName, LogType.LOG_TYPE_INFO)
        wait(1)
        Try
            'If ZoneModel = "WD100" Then
            ' we need to find out at start up whether something is docked. The Zone player doesn't fire off automatically
            Dim AudioInputName
            AudioInputName = AudioIn.QueryStateVariable("AudioInputName")
            If g_bDebug Then Log("GetAudioInputAttributes for zoneplayer = " & ZoneName & " found AudioInputName = " & AudioInputName.ToString, LogType.LOG_TYPE_INFO)
            If ZoneModel = "WD100" Then MyDockediPodPlayerName = AudioInputName
            Dim LineInConnected As String
            LineInConnected = AudioIn.QueryStateVariable("LineInConnected")
            If LineInConnected = "" Then LineInConnected = "False"
            iPodDockChange(LineInConnected)
            AudioInState(4) = LineInConnected
            MyLineInputConnected = LineInConnected
            'End If
        Catch ex As Exception
            Log("Error in GetAudioInputAttributes for zoneplayer = " & ZoneName & " when getting the AudioInputName with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Sub AddiPodPlayerNameToINIFile(ByVal iPodName As String)
        Dim DBiPodNamesString As String = ""
        Dim iPodDBNames() As String
        DBiPodNamesString = GetStringIniFile("iPod Player Names", "iPod Players", "")
        iPodDBNames = Split(DBiPodNamesString, ":|:")
        For Each iPodDBName In iPodDBNames
            If g_bDebug Then Log("AddiPodPlayerNameToINIFile found '" & iPodDBName & "' and is looking for '" & iPodName & "'", LogType.LOG_TYPE_INFO)
            If iPodDBName = iPodName Then
                ' name already exist
                Exit Sub
            End If
        Next
        ' this means that the name does note exist yet
        If DBiPodNamesString <> "" Then DBiPodNamesString = DBiPodNamesString & ":|:"
        DBiPodNamesString &= iPodName
        WriteStringIniFile("iPod Player Names", "iPod Players", DBiPodNamesString)
        If g_bDebug Then Log("AddiPodPlayerNameToINIFile added " & iPodName & " to the ini file", LogType.LOG_TYPE_INFO)
    End Sub

    Public Function GetTimeNow() As String
        GetTimeNow = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(3)
        Try
            AlarmClock.InvokeAction("GetTimeNow", InArg, OutArg)
            GetTimeNow = OutArg(1) ' Format = 2011-01-27 07:32:30
        Catch ex As Exception
            Log("ERROR in GetTimeNow for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
    End Function

    Private Function GoFindTheAlarmZone(ByVal UDN As String) As String
        GoFindTheAlarmZone = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(1)
        Try
            AlarmClock.InvokeAction("ListAlarms", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in GoFindTheAlarmZone for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If OutArg(0) = "" Then Exit Function ' no alarms stored
        ' OutArg(0) = CurrentAlarmList
        ' OutArg(1) = CurrentAlarmListVersion
        ' here is an example of what will be returned
        ' <Alarms>
        '   <Alarm ID="1" StartTime="18:47:00" Duration="02:00:00" Recurrence="DAILY" Enabled="1" RoomUUID="RINCON_000E5824C3B001400" ProgramURI="x-rincon-buzzer:0" 
        '           ProgramMetaData="" PlayMode="NORMAL" Volume="41" IncludeLinkedZones="0"/>
        '<  Alarm ID="15" StartTime="19:12:00" Duration="02:00:00" Recurrence="DAILY" Enabled="1" RoomUUID="RINCON_000E5825227A01400" ProgramURI="x-rincon-buzzer:0" 
        '           ProgramMetaData="" PlayMode="SHUFFLE_NOREPEAT" Volume="25" IncludeLinkedZones="0"/>
        '</Alarms>
        Dim xmlData As XmlDocument = New XmlDocument
        Try
            xmlData.LoadXml(OutArg(0))
        Catch ex As Exception
            Log("ERROR in GoFindTheAlarmZone for zoneplayer = " & ZoneName & " loading XML with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Dim LoopIndex As Integer = 0
        Dim StartTime As String = ""
        Dim RoomUUID As String = ""
        Dim SonosTime As DateTime ' = GetTimeNow()        ' Format is 2011-01-27 07:32:30
        Try
            SonosTime = GetTimeNow()    ' added 11/7/2018 because it failed on a paired player and returned an empty string
        Catch ex As Exception
            Exit Function
        End Try
        'Dim Now As DateTime = DateTime.Now
        Dim NowinMinutes As Integer = 0
        NowinMinutes = SonosTime.Hour * 60 + SonosTime.Minute
        If SonosTime.Second > 30 Then
            NowinMinutes = NowinMinutes + 1
        End If
        Dim AlarmTimeInMinutes As Integer = 0
        Dim Times
        Do
            Try
                StartTime = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("StartTime").Value
                If g_bDebug Then Log("GoFindTheAlarmZone for zoneplayer = " & ZoneName & " found StartTime = " & StartTime & " and looking for = " & SonosTime.ToString, LogType.LOG_TYPE_INFO)
                Try
                    Times = Split(StartTime, ":")
                    AlarmTimeInMinutes = Times(0) * 60 + Times(1)
                Catch ex As Exception
                End Try
                If g_bDebug Then Log("GoFindTheAlarmZone for zoneplayer = " & ZoneName & " AlarmtimeInMinutes = " & AlarmTimeInMinutes.ToString & " and NowinMinutes = " & NowinMinutes.ToString, LogType.LOG_TYPE_INFO)
                If AlarmTimeInMinutes = NowinMinutes Then
                    ' we found the alarm, now check which zone it is
                    RoomUUID = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("RoomUUID").Value
                    If RoomUUID = UDN Then
                        GoFindTheAlarmZone = UDN
                        Exit Function
                    End If
                Else
                    ' continue looking
                End If
            Catch ex As Exception
                'Log( "ERROR in GoFindTheAlarmZone for zoneplayer = " & ZoneName & " bailed out at counter =" & LoopIndex.Message, LogType.LOG_TYPE_ERROR)
                Exit Do
            End Try
            LoopIndex = LoopIndex + 1
            If LoopIndex > 100 Then Exit Do
        Loop
    End Function

    Public Function ListAlarms() As String
        ListAlarms = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(1)
        Try
            AlarmClock.InvokeAction("ListAlarms", InArg, OutArg)
            ListAlarms = OutArg(0)
            Log("ListAlarms for zoneplayer = " & ZoneName & " are " & ListAlarms.ToString, LogType.LOG_TYPE_INFO)
        Catch ex As Exception
            Log("ERROR in ListAlarms for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
    End Function

    Public Function SetAlarm(ByVal AlarmID As String, ByVal State As Boolean) As String
        SetAlarm = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(1)
        Try
            AlarmClock.InvokeAction("ListAlarms", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If OutArg(0) = "" Then Exit Function ' no alarms stored
        ' OutArg(0) = CurrentAlarmList
        ' OutArg(1) = CurrentAlarmListVersion
        ' here is an example of what will be returned
        ' <Alarms>
        '   <Alarm ID="1" StartTime="18:47:00" Duration="02:00:00" Recurrence="DAILY" Enabled="1" RoomUUID="RINCON_000E5824C3B001400" ProgramURI="x-rincon-buzzer:0" 
        '           ProgramMetaData="" PlayMode="NORMAL" Volume="41" IncludeLinkedZones="0"/>
        '<  Alarm ID="15" StartTime="19:12:00" Duration="02:00:00" Recurrence="DAILY" Enabled="1" RoomUUID="RINCON_000E5825227A01400" ProgramURI="x-rincon-buzzer:0" 
        '           ProgramMetaData="" PlayMode="SHUFFLE_NOREPEAT" Volume="25" IncludeLinkedZones="0"/>
        '</Alarms>
        Dim xmlData As XmlDocument = New XmlDocument
        Try
            xmlData.LoadXml(OutArg(0))
        Catch ex As Exception
            Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & " loading XML with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Dim LoopIndex As Integer = 0

        Do
            Try
                If xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("ID").Value = AlarmID Then
                    ' OK we found it
                    If g_bDebug Then Log("SetAlarm for zoneplayer = " & ZoneName & " found AlarmID = " & AlarmID, LogType.LOG_TYPE_INFO)
                    Exit Do
                End If
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & " bailed out at counter =" & LoopIndex.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            LoopIndex = LoopIndex + 1
            If LoopIndex > 100 Then
                If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & " bailed out at counter =" & LoopIndex.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Loop

        ' we need to set
        ' AlarmID               UI4
        ' StartTimeLocal        String
        ' Duration              String
        ' Recurrence            ONCE/WEEKDAYS/WEEKENDS/DAILY
        ' Enabled               false/true
        ' ROOMUUID              String
        ' ProgramURI            String
        ' ProgramMetaData       String
        ' PlayMode              String
        ' Volume                UI2
        ' IncludeLinkedZones    Boolean
        ReDim InArg(10)
        ReDim OutArg(0)
        InArg(0) = AlarmID
        Try
            InArg(1) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("StartTime").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving StartTime", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            InArg(2) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("Duration").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving Duration", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            InArg(3) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("Recurrence").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving Recurrence", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        InArg(4) = State
        Try
            InArg(5) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("RoomUUID").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving RoomUUID", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            InArg(6) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("ProgramURI").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving ProgramURI", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            InArg(7) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("ProgramMetaData").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving ProgramMetaData", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            InArg(8) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("PlayMode").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving PlayMode", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            InArg(9) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("Volume").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving Volume", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            InArg(10) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("IncludeLinkedZones").Value
        Catch ex As Exception
            If g_bDebug Then Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & "retrieving IncludeLinkedZones", LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Try
            AlarmClock.InvokeAction("UpdateAlarm", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & " when setting new state with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If g_bDebug Then Log("Successful SetAlarm for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        SetAlarm = "OK"
    End Function

    Public Function SetAlarmParms(ByVal AlarmID As String, Optional ByVal StartTimeLocal As String = "", Optional ByVal Duration As String = "",
                                  Optional ByVal Recurrence As String = "", Optional ByVal Enabled As String = "", Optional ByVal ROOMUUID As String = "",
                                  Optional ByVal ProgramURI As String = "", Optional ByVal ProgramMetaData As String = "", Optional ByVal PlayMode As String = "",
                                  Optional ByVal Volume As String = "", Optional ByVal IncludeLinkedZones As String = "") As String
        ' Enabled               False/True
        ' Recurrence            ONCE/WEEKDAYS/WEEKENDS/DAILY
        ' IncludeLinkedZones      False/True
        SetAlarmParms = ""
        If DeviceStatus = "Offline" Then Exit Function
        Dim InArg(0)
        Dim OutArg(1)
        Try
            AlarmClock.InvokeAction("ListAlarms", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & " with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If OutArg(0) = "" Then Exit Function ' no alarms stored
        ' OutArg(0) = CurrentAlarmList
        ' OutArg(1) = CurrentAlarmListVersion
        ' here is an example of what will be returned
        ' <Alarms>
        '   <Alarm ID="1" StartTime="18:47:00" Duration="02:00:00" Recurrence="DAILY" Enabled="1" RoomUUID="RINCON_000E5824C3B001400" ProgramURI="x-rincon-buzzer:0" 
        '           ProgramMetaData="" PlayMode="NORMAL" Volume="41" IncludeLinkedZones="0"/>
        '<  Alarm ID="15" StartTime="19:12:00" Duration="02:00:00" Recurrence="DAILY" Enabled="1" RoomUUID="RINCON_000E5825227A01400" ProgramURI="x-rincon-buzzer:0" 
        '           ProgramMetaData="" PlayMode="SHUFFLE_NOREPEAT" Volume="25" IncludeLinkedZones="0"/>
        '</Alarms>
        Dim xmlData As XmlDocument = New XmlDocument
        Try
            xmlData.LoadXml(OutArg(0))
        Catch ex As Exception
            Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & " loading XML with Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        Dim LoopIndex As Integer = 0

        Do
            Try
                If xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("ID").Value = AlarmID Then
                    ' OK we found it
                    If g_bDebug Then Log("SetAlarmParms for zoneplayer = " & ZoneName & " found AlarmID = " & AlarmID, LogType.LOG_TYPE_INFO)
                    Exit Do
                End If
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & " bailed out at counter =" & LoopIndex.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
            LoopIndex = LoopIndex + 1
            If LoopIndex > 100 Then
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & " bailed out at counter =" & LoopIndex.ToString, LogType.LOG_TYPE_ERROR)
                Exit Function
            End If
        Loop

        ' we need to set
        ' AlarmID               UI4
        ' StartTimeLocal        String
        ' Duration              String
        ' Recurrence            ONCE/WEEKDAYS/WEEKENDS/DAILY
        ' Enabled               false/true
        ' ROOMUUID              String
        ' ProgramURI            String
        ' ProgramMetaData       String
        ' PlayMode              String
        ' Volume                UI2
        ' IncludeLinkedZones    Boolean
        ReDim InArg(10)
        ReDim OutArg(0)
        InArg(0) = AlarmID
        If StartTimeLocal = "" Then
            Try
                InArg(1) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("StartTime").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving StartTime", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(1) = StartTimeLocal
        End If
        If Duration = "" Then
            Try
                InArg(2) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("Duration").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving Duration", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(2) = Duration
        End If
        If Recurrence = "" Then
            Try
                InArg(3) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("Recurrence").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving Recurrence", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(3) = Recurrence
        End If
        If Enabled = "" Then
            Try
                InArg(4) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("Enabled").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving Enabled", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(4) = CBool(Enabled)
        End If
        If ROOMUUID = "" Then
            Try
                InArg(5) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("RoomUUID").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving RoomUUID", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(5) = ROOMUUID
        End If
        If ProgramURI = "" Then
            Try
                InArg(6) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("ProgramURI").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving ProgramURI", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(6) = ProgramURI
        End If
        If ProgramMetaData = "" Then
            Try
                InArg(7) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("ProgramMetaData").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving ProgramMetaData", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(7) = ProgramMetaData
        End If
        If PlayMode = "" Then
            Try
                InArg(8) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("PlayMode").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving PlayMode", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(8) = PlayMode
        End If
        If Volume = "" Then
            Try
                InArg(9) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("Volume").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving Volume", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(9) = Volume
        End If
        If IncludeLinkedZones = "" Then
            Try
                InArg(10) = xmlData.GetElementsByTagName("Alarm").Item(LoopIndex).Attributes("IncludeLinkedZones").Value
            Catch ex As Exception
                If g_bDebug Then Log("ERROR in SetAlarmParms for zoneplayer = " & ZoneName & "retrieving IncludeLinkedZones", LogType.LOG_TYPE_ERROR)
                Exit Function
            End Try
        Else
            InArg(10) = CBool(IncludeLinkedZones)
        End If
        Try
            AlarmClock.InvokeAction("UpdateAlarm", InArg, OutArg)
        Catch ex As Exception
            Log("ERROR in SetAlarm for zoneplayer = " & ZoneName & " when setting new state with UPNP Error = " & UPnP_Error(Err.Number) & ". Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
            Exit Function
        End Try
        If g_bDebug Then Log("Successful SetAlarmParms for zoneplayer = " & ZoneName, LogType.LOG_TYPE_INFO)
        SetAlarmParms = "OK"
    End Function

    Private Sub ZoneNameChanged(ByVal NewZoneName As String)
        If MyHSPIControllerRef.ZoneNameChanged(ZoneName, GetUDN, NewZoneName, False) Then
            If g_bDebug Then Log("ZoneNameChanged called for Zone = " & ZoneName & " and changed name to " & NewZoneName.ToString, LogType.LOG_TYPE_INFO)
            PlayChangeNotifyCallback(player_status_change.ConfigChange, player_state_values.ZoneName)
            ZoneName = NewZoneName
        End If
    End Sub

    Private Sub ZonePairingChanged(ByVal ChannelMapSet As String, Coordinator As String)
        ' when an S5 is paired, you get something like this
        ' Office:  Var Name = ChannelMapSet Value = RINCON_000E5858C97A01400:LF,LF;RINCON_000E5859008Axxxx:RF,RF;RINCON_000E5898164001400:SW,SW
        ' Office2: Var Name = ChannelMapSet Value = RINCON_000E5858C97A01400:LF,LF;RINCON_000E5859008Axxxx:RF,RF
        '  with the UDN on the left being the "Master"
        ' When they Unpair, the value is empty ""
        ' The SW,SW indicated a SubWoofer
        ' rewritten on 7/14/2019 in v3.1.0.34 to support S18 players who have right as master but the statement still holds, the FIRST UDN in the string is the master
        If g_bDebug Then Log("ZonePairingChanged called for Zone = " & ZoneName & " with ChannelMapSet " & ChannelMapSet & " and Coordinator = " & Coordinator, LogType.LOG_TYPE_INFO)
        If MyChannelMapSet = ChannelMapSet Then Exit Sub ' nothing changed
        MyChannelMapSet = ChannelMapSet ' I may have to check on the subwoofer here because it is used to unpair, not sure whether I unpair with the subwoofer in the string or without
        If ChannelMapSet = "" Then
            ' this is Unpairing. We don't really have to do anything, we will receive a zone name change which will correct the zone name and we'll get an unlink
            MyZoneIsPairMaster = False
            MyZoneIsPairSlave = False
            MyZonePairMasterZoneName = ""
            MyZonePairSlaveZoneUDN = ""
            MyZonePairMasterZoneUDN = ""
            MyZonePairSubWooferZoneUDN = ""
            MyZonePairLeftFrontUDN = ""
            MyZonePairRightFrontUDN = ""
            Exit Sub
        End If
        ' this is pairing
        Dim ChannelMapSetInfos
        Dim ZoneInfos
        Try
            ChannelMapSetInfos = Split(ChannelMapSet, ";")
            If UBound(ChannelMapSetInfos) = 0 Then
                ' this should not be
                Exit Sub
            End If
            'Dim LeftFrontUDN As String = ""
            'Dim RightFrontUDN As String = ""
            'Dim LeftRearUDN As String = ""
            'Dim RightRearUDN As String = ""
            Dim SubUDN As String = ""
            Dim FirstUDN As String = ""
            ZoneInfos = Split(ChannelMapSetInfos(0), ":")
            If ZoneInfos(1).ToString.ToUpper = "LF,LF" Then
                MyZonePairLeftFrontUDN = ZoneInfos(0)
                FirstUDN = MyZonePairLeftFrontUDN
            ElseIf ZoneInfos(1).ToString.ToUpper = "RF,RF" Then
                MyZonePairRightFrontUDN = ZoneInfos(0)
                FirstUDN = MyZonePairRightFrontUDN
            ElseIf ZoneInfos(1).ToString.ToUpper = "LF,RF" Then ' added 7/25/2019 in v3.1.0.36
                MyZonePairLeftFrontUDN = ZoneInfos(0)
                MyZonePairRightFrontUDN = ZoneInfos(0)
                FirstUDN = MyZonePairRightFrontUDN
            ElseIf ZoneInfos(1).ToString.ToUpper = "SW,SW" Then
                SubUDN = ZoneInfos(0)
            End If
            ZoneInfos = Split(ChannelMapSetInfos(1), ":")

            If ZoneInfos(1).ToString.ToUpper = "LF,LF" Then
                MyZonePairLeftFrontUDN = ZoneInfos(0)
            ElseIf ZoneInfos(1).ToString.ToUpper = "RF,RF" Then
                MyZonePairRightFrontUDN = ZoneInfos(0)
            ElseIf ZoneInfos(1).ToString.ToUpper = "LF,RF" Then  ' added 7/25/2019 in v3.1.0.36
                MyZonePairLeftFrontUDN = ZoneInfos(0)
                MyZonePairRightFrontUDN = ZoneInfos(0)
            ElseIf ZoneInfos(1).ToString.ToUpper = "SW,SW" Then
                SubUDN = ZoneInfos(0)
            End If
            If UBound(ChannelMapSetInfos) >= 2 Then
                ' there is a subwoofer involved
                ZoneInfos = Split(ChannelMapSetInfos(2), ":")
                If ZoneInfos(1).ToString.ToUpper = "LF,LF" Then
                    MyZonePairLeftFrontUDN = ZoneInfos(0)
                ElseIf ZoneInfos(1).ToString.ToUpper = "RF,RF" Then
                    MyZonePairRightFrontUDN = ZoneInfos(0)
                ElseIf ZoneInfos(1).ToString.ToUpper = "LF,RF" Then  ' added 7/25/2019 in v3.1.0.36
                    MyZonePairLeftFrontUDN = ZoneInfos(0)
                    MyZonePairRightFrontUDN = ZoneInfos(0)
                ElseIf ZoneInfos(1).ToString.ToUpper = "SW,SW" Then
                    SubUDN = ZoneInfos(0)
                End If
            End If
            MyZonePairMasterZoneUDN = FirstUDN
            If MyZonePairMasterZoneUDN <> UDN Then
                ' this player is NOT the coordinator
                MyZonePairSlaveZoneUDN = UDN
            Else
                ' the slave is the other player in the pairing
                If MyZonePairMasterZoneUDN = MyZonePairLeftFrontUDN Then
                    MyZonePairSlaveZoneUDN = MyZonePairRightFrontUDN
                Else
                    MyZonePairSlaveZoneUDN = MyZonePairLeftFrontUDN
                End If
            End If
            MyZonePairSubWooferZoneUDN = SubUDN
            If (UDN = FirstUDN) And (MyZonePairSlaveZoneUDN <> "") Then ' check on slavezone is to avoid this player is just linked to a subwoofer
                ' this zone is master
                MyZoneIsPairMaster = True
                MyZoneIsPairSlave = False
                MyZonePairMasterZoneName = MyHSPIControllerRef.GetZoneNamebyUDN(Coordinator)
                If g_bDebug Then Log("ZonePairingChanged for Zone = " & ZoneName & " and became Master", LogType.LOG_TYPE_INFO)
            ElseIf (MyZonePairSlaveZoneUDN <> "") Then
                ' this zone is Slave
                MyZonePairMasterZoneName = MyHSPIControllerRef.GetZoneNamebyUDN(Coordinator)
                MyZoneIsPairMaster = False
                MyZoneIsPairSlave = True
                If g_bDebug Then Log("ZonePairingChanged for Zone = " & ZoneName & " and became Slave", LogType.LOG_TYPE_INFO)
            End If
        Catch ex As Exception
            Log("Error in ZonePairingChanged for Zone = " & ZoneName & " with Error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Private Sub PlaybarPairingChanged(ByVal HTSatChanMapSet As String, Coordinator As String)
        ' when an Playbar is paired, you get something like this
        ' S9:  Var Name = HTSatChanMapSet Value = RINCON_B8E9377C7E2601400:LF,RF;RINCON_000E5898164001400:SW;RINCON_000E58C174F201400:LR;RINCON_000E58C174DE01400:RR
        ' S1:  Var Name = HTSatChanMapSet Value = RINCON_B8E9377C7E2601400:LF,RF;RINCON_000E58C174DE01400:RR <- note this is from the S1 at the Rear LEFT!!!
        ' SUB:  Var Name = HTSatChanMapSet Value = RRINCON_B8E9377C7E2601400:LF,RF;RINCON_000E5898164001400:SW;RINCON_000E58C174F201400:LR;RINCON_000E58C174DE01400:RR
        ' if a connect:Amp is added it shows up like this HTSatChanMapSet = RINCON_7828CA57993B01400:LF,RF;RINCON_5CAAFDED681A01400:LR,RR
        ' 
        ' rewritten on 7/14/2019 in v3.1.0.34 to support S18 players who have right as master
        If g_bDebug Then Log("PlaybarPairingChanged called for Zone = " & ZoneName & " with HTSatChanMapSet = " & HTSatChanMapSet & " and Coordinator = " & Coordinator, LogType.LOG_TYPE_INFO)
        If MyHTSatChanMapSet = HTSatChanMapSet Then Exit Sub ' nothing changed
        MyHTSatChanMapSet = HTSatChanMapSet ' I may have to check on the subwoofer here because it is used to unpair, not sure whether I unpair with the subwoofer in the string or without
        If HTSatChanMapSet = "" Then
            ' this is Unpairing. We don't really have to do anything, we will receive a zone name change which will correct the zone name and we'll get an unlink
            MyZoneIsPlaybarMaster = False
            MyZoneIsPlaybarSlave = False
            MyZonePlayBarLeftRearUDN = ""
            MyZonePlayBarRightRearUDN = ""
            MyZonePlayBarLeftFrontUDN = ""
            MyZonePlayBarRightFrontUDN = ""
            MyZonePlayBarUDN = ""
            HandleUnlinkZone()
            Exit Sub
        End If
        ' this is pairing

        Dim ChannelMapSetInfos As String() = Nothing
        Dim FirstUDN As String = ""

        Try
            ChannelMapSetInfos = Split(HTSatChanMapSet, ";")
            If ChannelMapSetInfos Is Nothing Then
                ' this should not be
                Exit Sub
            End If
            For Each ChannelMapInfo As String In ChannelMapSetInfos
                Dim ZoneInfos As String() = Nothing
                ZoneInfos = Split(ChannelMapInfo, ":")
                ' example = RINCON_B8E9377C7E2601400:LF,RF
                ' so left or part (0) is UDN, part (1) defines where the UDN is used, could be LF, RF, SW, LR, RR
                If ZoneInfos IsNot Nothing Then
                    If UBound(ZoneInfos) > 0 Then
                        Dim ZoneUDN As String = ZoneInfos(0)
                        If FirstUDN = "" And ZoneUDN <> "" Then
                            FirstUDN = ZoneUDN
                            If FirstUDN = UDN Then
                                If g_bDebug Then Log("PlaybarPairingChanged for Zone = " & ZoneName & " and became Master", LogType.LOG_TYPE_INFO)
                                MyZoneIsPlaybarMaster = True
                                MyZoneIsPlaybarSlave = False
                                MyZonePlayBarUDN = ""
                            Else
                                MyZoneIsPlaybarMaster = False
                                MyZoneIsPlaybarSlave = True
                                MyZonePlayBarUDN = ZoneUDN
                            End If
                        End If
                        Dim ZoneLocationsString = ZoneInfos(1)
                        Dim ZoneLocations = Nothing
                        ZoneLocations = Split(ZoneLocationsString, ",")
                        If ZoneLocations IsNot Nothing Then
                            For Each zonelocation As String In ZoneLocations
                                Select Case Trim(UCase(zonelocation))
                                    Case "LF"
                                        ' this would be the playbar
                                        If ZoneUDN = UDN Then
                                            MyZonePlayBarLeftFrontUDN = ZoneUDN
                                            If ZoneUDN <> FirstUDN Then HandleLinkedZones(MyZonePlayBarUDN)
                                        End If
                                        'MyZonePlayBarUDN = ZoneUDN
                                        'If ZoneUDN = UDN Then
                                            'MyZoneIsPlaybarMaster = True
                                            'MyZoneIsPlaybarSlave = False
                                            'MyZonePlayBarUDN = ""
                                        'Else
                                            'MyZoneIsPlaybarMaster = False
                                            'MyZoneIsPlaybarSlave = True
                                        'End If
                                    Case "RF"
                                        ' this would be the playbar
                                        If ZoneUDN = UDN Then
                                            MyZonePlayBarRightFrontUDN = ZoneUDN
                                            If ZoneUDN <> FirstUDN Then HandleLinkedZones(MyZonePlayBarUDN)
                                        End If
                                        'MyZonePlayBarUDN = ZoneUDN
                                        'If ZoneUDN = UDN Then
                                            'MyZoneIsPlaybarMaster = True
                                            'MyZoneIsPlaybarSlave = False
                                            'MyZonePlayBarUDN = ""
                                        'Else
                                            'MyZoneIsPlaybarMaster = False
                                            'MyZoneIsPlaybarSlave = True
                                        'End If
                                    Case "SW"
                                        ' this is the subwoofer
                                        If ZoneUDN = UDN Then
                                            MyZonePairSubWooferZoneUDN = ZoneUDN
                                            'MyZoneIsPlaybarSlave = True
                                        End If
                                    Case "LR"
                                        If ZoneUDN = UDN Then
                                            MyZonePlayBarLeftRearUDN = ZoneUDN
                                            'MyZoneIsPlaybarSlave = True
                                            If ZoneUDN <> FirstUDN Then HandleLinkedZones(MyZonePlayBarUDN)
                                        End If
                                    Case "RR"
                                        If ZoneUDN = UDN Then
                                            MyZonePlayBarRightRearUDN = ZoneUDN
                                            'MyZoneIsPlaybarSlave = True
                                            If ZoneUDN <> FirstUDN Then HandleLinkedZones(MyZonePlayBarUDN)
                                        End If
                                End Select
                            Next
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            Log("Error in PlaybarPairingChanged for Zone = " & ZoneName & " with Error " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
    End Sub

    Private Function ProcessClassInfo(ClassString As String) As String
        ProcessClassInfo = ""
        If ClassString = "" Then Exit Function
        Dim ClassItems As String()
        Try
            ClassItems = ClassString.Split(".")
            If UBound(ClassItems) > 2 Then
                Return ClassItems(3)
            End If
        Catch ex As Exception
        End Try
    End Function

End Class

<Serializable()>
Public Class myUPnPControlCallback
    Public Event ControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event ControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent ControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent ControlDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPAVTransportCallback
    Public Event TransportStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event TransportDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent TransportStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent TransportDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPRenderingControlCallback
    Public Event RenderingControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event RenderingControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent RenderingControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent RenderingControlDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPAlarmClockCallback
    Public Event AlarmClockControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event AlarmClockControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent AlarmClockControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent AlarmClockControlDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPMusicServicesCallback
    Public Event MusicServicesControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event MusicServicesControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent MusicServicesControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent MusicServicesControlDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPSystemPropertiesCallback
    Public Event SystemPropertiesControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event SystemPropertiesControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent SystemPropertiesControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent SystemPropertiesControlDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPZonegroupTopologyCallback
    Public Event ZonegroupTopologyControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event ZonegroupTopologyControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent ZonegroupTopologyControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent ZonegroupTopologyControlDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPGroupManagementCallback
    Public Event GroupManagementControlStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event GroupManagementControlDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent GroupManagementControlStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent GroupManagementControlDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPConnectionManagerCallback
    Public Event ConnectionManagerStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event ConnectionManagerDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent ConnectionManagerStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent ConnectionManagerDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPContentDirectoryCallback
    Public Event ContentDirectoryStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event ContentDirectoryDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent ContentDirectoryStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent ContentDirectoryDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPAudioInCallback
    Public Event AudioInStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event AudioInDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent AudioInStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent AudioInDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPDevicePropertiesCallback
    Public Event DevicePropertiesStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event DevicePropertiesDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent DevicePropertiesStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent DevicePropertiesDied()
        Return 0
    End Function
End Class

<Serializable()>
Public Class myUPnPQueueServiceCallback
    Public Event QueueServiceStateChange(ByVal StateVarName As String, ByVal Value As String)
    Public Event QueueServiceDied()

    Public Function StateVariableChanged(ByVal StateVarName As String, ByVal Value As String) As Integer
        RaiseEvent QueueServiceStateChange(StateVarName, Value)
        Return 0
    End Function

    Public Function ServiceInstanceDied() As Integer
        RaiseEvent QueueServiceDied()
        Return 0
    End Function
End Class

Public Module UPNPDefinitions
    ' here are the index values for GetMediaInfo
    Public Const gmiNrTracks = 0            ' ui4
    Public Const gmiMediaDuration = 1       ' String
    Public Const gmiCurrentURI = 2          ' String
    Public Const gmiCurrentURIMetaData = 3  ' string
    Public Const gmiNextURI = 4             ' string
    Public Const gmiNextURIMetaData = 5     ' string
    Public Const gmiPlayMedium = 6          ' string
    Public Const gmiRecordMedium = 7        ' string
    Public Const gmiWriteStatus = 8         ' string

    ' here are the index values for GetPositionInfo
    Public Const pmiTrack = 0               ' ui4
    Public Const pmiTrackDuration = 1       ' String xx:xx:xx
    Public Const pmiTrackMetaData = 2       ' string
    Public Const pmiTrackURI = 3            ' string
    Public Const pmiRelTime = 4             ' string xx:xx:xx
    Public Const pmiAbsTime = 5             ' string
    Public Const pmiRelCount = 6            ' i4
    Public Const pmiAbsCount = 7            ' i4

End Module

<Serializable()>
Public Class SavedLinkInfo
    Public MySavedTrackInfo
    Public MySavedQueuePosition As Integer = 0
    Public MySavedTransportState As String = ""
    Public MySavedQueue As String = ""
    Public MySavedQueueObjectID As String = ""
    Public MySavedMasterVolumeLevel As Integer = 0
    Public MySavedMuteState As Boolean = False
    Public MySavedPlayMode As String = "NORMAL"
    Public MySavedLineInLeftVolume As Integer = 0
    Public MySavedLineInRightVolume As Integer = 0
    Public MySavedZoneisLinked As Boolean = False
    Public MySavedSourceLinkedZone As String
    Public MySavedSavedQueueFlag As Boolean = True
    Public LinkgroupName As String = ""
    Public InfoIsSaved As Boolean = False
    Public MySavedTargetZoneLinkedList As String = ""
    Public MySavedZoneIsPairMaster As Boolean = False
    Public MySavedZoneIsPairSlave As Boolean = False
    Public MySavedZonePairMasterZoneName As String = ""
    Public MySavedChannelMapSet As String = ""
    Public MySavedZonePairMasterUDN As String = ""
    Public MySavedZonePairSlaveUDN As String = ""
    Public MySavedZonePairSubWooferZoneUDN As String = ""

    Public Sub New()
        MyBase.New()
        FlushInfo()
        LinkgroupName = ""
    End Sub

    Public Sub FlushInfo()
        MySavedTrackInfo = {"", "", "", "", "", "", "", "0", "", "", "False", "", ""}
        MySavedQueuePosition = 0
        MySavedTransportState = ""
        MySavedQueue = ""
        MySavedQueueObjectID = ""
        MySavedMasterVolumeLevel = 0
        MySavedMuteState = False
        MySavedLineInLeftVolume = 0
        MySavedLineInRightVolume = 0
        MySavedZoneisLinked = False
        MySavedSourceLinkedZone = ""
        MySavedSavedQueueFlag = False
        InfoIsSaved = False
        MySavedTargetZoneLinkedList = ""
        MySavedZoneIsPairMaster = False
        MySavedZoneIsPairSlave = False
        MySavedZonePairMasterZoneName = ""
        MySavedChannelMapSet = ""
        MySavedZonePairMasterUDN = ""
        MySavedZonePairSlaveUDN = ""
        MySavedZonePairSubWooferZoneUDN = ""
    End Sub

End Class
