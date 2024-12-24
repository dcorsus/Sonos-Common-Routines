Imports Newtonsoft.Json
Public Module HS_GLOBAL_VARIABLES

    ' interface status
    ' for InterfaceStatus function call
    Public Const ERR_NONE = 0
    Public Const ERR_SEND = 1
    Public Const ERR_INIT = 2

    ' Master State Pairs
    Public Const msDisconnected = 100
    Public Const msConnected = 101
    Public Const msBuildingDB = 102
    Public Const msInitializing = 103
    Public Const msPauseAll = 1000
    Public Const msPlayAll = 1001
    Public Const msMuteAll = 1002
    Public Const msUnmuteAll = 1003
    Public Const msBuildDB = 1004

    ' Player State Pairs
    Public Const psStopped = 2
    Public Const psPaused = 3
    Public Const psPlaying = 1
    Public Const psPlay = 1001
    Public Const psStop = 1002
    Public Const psPause = 1003
    Public Const psNext = 1004
    Public Const psPrevious = 1005
    Public Const psBuildiPodDB = 1006
    Public Const psShuffle = 1007
    Public Const psRepeat = 1008
    Public Const psVolUp = 1009
    Public Const psVolDown = 1010
    Public Const psMute = 1 ' 1011
    Public Const psBalanceLeft = 1012
    Public Const psBalanceRight = 1013
    Public Const psLoudness = 1014
    Public Const psVolSlider = 15
    Public Const psBalanceSlider = 200
    Public Const psPlayPause = 1015
    Public Const psUnMute = 0 ' 1016
    Public Const psMuteToggle = 1002 ' changed to be compatible with HS3 1017
    Public Const psWOL = 1020

    ' MuteState Pairs
    Public Const msMuted = 1 '1001
    Public Const msUnmuted = 0 '1000

    ' ShuffleState Pairs
    Public Const ssShuffled = 1001
    Public Const ssNoShuffle = 1000

    ' RepeatState Pairs
    Public Const rsRepeat = 1001
    Public Const rsnoRepeat = 1000

    ' Down - Slider - Up pairs
    Public Const vpDown = 1000
    Public Const vpSlider = 1
    Public Const vpUp = 1001
    Public Const vpMidPoint = 300

    ' Toggle pairs
    Public Const tpOff = 1000
    Public Const tpOn = 1001
    Public Const tpToggle = 1002

    Public Const lsLoudnessOff = 1000
    Public Const lsLoudnessOn = 1001
    ' Define Trigger/Actions Variables
    '
    Public gTriggerNames(2) As String
    Public gTriggerActions(1) As String

    ' authentication values
    Public Const avNotAuthenticated = 0
    Public Const avAuthenticated = 1

    Public InterfaceVersion As Integer
    Public bShutDown As Boolean = False
    Public MyShutDownRequest As Boolean = False

    'Public instance As String = ""                             ' set when SupportMultipleInstances is TRUE
    Public gLogToDisk As Boolean = False
    Public gHSInitialized As Boolean = False
    Public ImRunningOnLinux As Boolean = False
    Public HSisRunningOnLinux As Boolean = False
    Public gInterfaceStatus As Integer = ERR_INIT           ' Interface status
    Public PlugInIPAddress As String = ""
    Public PluginIPPort As String = ""
    Public ImRunningLocal As Boolean = True

    Public MasterHSDeviceRef As Integer = -1
    Public MasterHSFeatureRef As Integer = -1
    Public gIOEnabled As Boolean = False                    ' IO interface enabled
    Public MyPIisInitialized As Boolean = False

    ' File Path Name definitions. May be overwritten in InitIO
    Public CurrentAppPath As String = ""
#If HS3 = "True" Then
    Public MusicDBPath As String = "/html/" & tIFACE_NAME & "/MusicDB/SonosDB.sdb"
        Public DockedPlayersDBPath As String = "/html/" & tIFACE_NAME & "/MusicDB/"
#Else
    Public dbFileExtension As String = ".sdb"
    Public MusicDBPath As String = "/html/" & tIFACE_NAME & "/MusicDB/SonosDB_"
#End If

    Public RadioStationsDBPath As String = "/html/" & tIFACE_NAME & "/MusicDB/SonosRadioStationsDB.sdb"
    Public FileArtWorkPath As String = tIFACE_NAME & "\Artwork\"
    Public DebugLogFileName As String = "/" & tIFACE_NAME & "/Logs/SonosDebug.txt"
    Public migrationLogFileName As String = "/" & tIFACE_NAME & "/Logs/SonosMigrationLog.txt"
    Public verificationLogFileName As String = "/" & tIFACE_NAME & "/Logs/SonosVerificationLog.txt"
    Public Const DBConnectionString As String = "Data Source=" ' 

    ' URLs
    Public AnnouncementPath As String = "\" & tIFACE_NAME & "\Announcements\"
    Public AnnouncementURL As String = "/" & tIFACE_NAME & "/Announcements/"
    Public ImagesPath As String = "/images/" & sIFACE_NAME & "/"
    Public URLImagesPath As String = "/images/" & sIFACE_NAME & "/"
    Public NoArtPath As String = "/images/" & sIFACE_NAME & "/noart.png" ' this is based on \html\Sonos\Images
    Public ArtWorkPath As String = "/images/" & tIFACE_NAME & "/Artwork/"
    Public URLArtWorkPath As String = "/images/" & tIFACE_NAME & "/Artwork/"  '"/" & tIFACE_NAME & "/Artwork/"  
    Public helpURL As String = "/" & tIFACE_NAME & "/Help/Help.htm"


    Public RenderingInfo, ContentInfo As String

    Public Const cMaxNbrOfUPNP = 1000 '400
    Public MaxNbrOfUPNPObjects As Integer = cMaxNbrOfUPNP

    Public SonosSettingsHaveChanged As Boolean = False
    Public WirelessSettingsHaveChanged As Boolean = False

    Public MusicDBIsBeingEstablished As Boolean = False

    Public MyMusicDBItems As New MusicDBItems
    Public VolumeStep As Integer = 5
    Public Const cArtworkSize = 150
    Public ArtworkHSize As Integer = 0
    Public ArtworkVSize As Integer = 0

    Public NbrOfPingRetries As Integer = 3
    Public ShowFailedPings As Boolean = False
    Public myMaxAnnouncementTime As Integer = 100
    Public MyNoPingingFlag As Boolean = False

#If HS3 = "True" Then
    public LearnRadioStations as boolean
    public AutoBuildDockedDB as boolean
#Else
    Public stateDeviceInfoChoice As stateChoices = stateChoices.scTrack
    Public hidePairedPlayers As Boolean
    Public hideStatusDevices As Boolean
    Public categoryChoice As String = ""
    Public useTransitioningState As Boolean = False
    Public useAudioClip As Boolean = True
#End If

    ' Part of the UPNP stuff
    Public TCPListenerPort As Integer = 0 '12291   ' removed 10/19/2019
    Public UPnPSubscribeTimeOut As Integer = 1800
    Public MySSDPDevice As MySSDP
    Public Const LoopBackIPv4Address As String = "127.0.0.1"
    Public Const AnyIPv4Address As String = "0.0.0.0"

    Public upnpDebuglevel As DebugLevel = DebugLevel.dlOff
    Public piDebuglevel As DebugLevel = DebugLevel.dlOff


    Public AnnouncementLink As AnnouncementItems = Nothing
    Public AnnouncementsInQueue As Boolean = False
    Public AnnouncementInProgress As Boolean = False
    Public announcementReEntryCounter As Integer = 0
    Public MyAnnouncementCountdown As Integer = 100
    Public MyAnnouncementIndex As Integer = 0
    Public LastStoredQueueID As String = ""
    Public AnnouncementTitle As String = "HomeSeer Announcement"
    Public AnnouncementAuthor As String = "dC"
    Public AnnouncementAlbum As String = "SonosController"
    Public ResendPlay As Integer = 0
    Public MyAnnouncementWaitToSendPlay As Integer = 0
    Public MyAnnouncementWaitBetweenPlayers As Integer = 0

    Public MyHSTrackLengthFormat As HSSTrackPositionSettings = HSSTrackPositionSettings.TPSSeconds
    Public MyHSTrackPositionFormat As HSSTrackPositionSettings = HSSTrackPositionSettings.TPSSeconds

    Public PreviousVersion As Integer = 0
    Public Const CurrentVersion As Integer = 2  ' added stop control in version 1;
    Public Const addStopControl As Integer = 1

    Public directNavigate As Boolean = False
    Public maxNavItems As Integer = 1000
    Public artRetrievalReEntryFlag As Boolean = False



    <Serializable()>
    Public Structure DBRecord
        Public id As String
        Public title As String
        Public parentid As String
        Public itemorcontainer As String
        Public classtype As String
        Public iconurl As String
        Public albumname As String
        Public artistname As String
        Public genre As String
        Public trackno As String
        Public URI As String
        Public type As NavEntryType
        Public protocolInfo As String
    End Structure

    <Serializable()>
    Public Structure NavItem
        Public key As String
        Public id As String
        Public value As String
        Public parentid As String
        Public itemorcontainer As String
        Public type As NavEntryType
        Public searchitem As String
        Public searchtext As String
    End Structure

    <Serializable()>
    Class VolumeAttributes
        Public volume As Integer = 0
        Public relative As Boolean = False
        Public applyToGroup As Boolean = False
    End Class
    Public Enum UPnPClassType
        ctUnknown = 0
        ctMusic = 1
        ctVideo = 2
        ctPictures = 3
    End Enum

    <Serializable()>
    Public Structure PlayerRecord
        Public UDN As String
        Public ZoneName As String
        Public ModelNbr As String
        Public pDevice As MyUPnPDevice
        Public IPAddress As String
        Public PlayerIcon As String
        Public IconURL As String
    End Structure

    'Public Enum Part
    '    player_control = 4 '2
    '    ZonePlayerName = 2 '3
    '    volume_set = 6 '4
    '    shuffle_set = 10 '5
    '    repeat_set = 12 '6
    '    genre = 3 '7
    '    artist = 5 '8
    '    album = 7 '9
    '    playlist = 9 '10
    '    radiostation = 15 '11
    '    track = 13 '12
    '    trackmatch = 11 '13
    '    Mute = 8 '14
    '    Loudness = 14 '15
    '    Left = 16 '16
    '    LineInput = 17
    '    Right = 18 '17
    '    TrackPosition = 19
    '    AudioBook = 21 ' 18
    '    PodCast = 22 '19
    '    iPodDBName = 20 ' 20
    'End Enum

    'Private Enum ActionUI_Priority
    '   playlist = 1
    '   genre = 2
    '   album = 3
    '   artist = 4
    'End Enum

    'Private Enum SonosTriggerEnum
    '    SonosTrackChange = 1
    '    SonosPlayerStop = 2
    '    SonosPlayerPaused = 3
    '    SonosPlayerStartPlaying = 4
    '    SonosVolumeUp = 5
    '    SonosVolumeDown = 6
    '    SonosPlayerDocked = 7
    '    SonosPlayerUndocked = 8
    '    SonosPlayerLineinConnected = 9
    '    SonosPlayerLineinDisconnected = 10
    '    SonosPlayerAlarmStart = 11
    '    SonosPlayerConfigChange = 12
    '    SonosPlayerTVConnected = 9
    '    SonosPlayerTVDisconnected = 10
    '    SonosAnnouncementStart = 11
    '    SonosAnnouncementStop = 12
    'End Enum

    Public Enum NavEntryType
        netRoot = 0
        netAlbums = 1
        netAlbum = 2
        netArtists = 3
        netArtist = 4
        netTracks = 5
        netTrack = 6
        netGenres = 7
        netGenre = 8
        netFavorites = 9
        netFavorite = 10
        netRadiostations = 11
        netRadioStation = 12
        netLineIputs = 13
        netPairings = 14
        netLearnedRadiostations = 15
        netPlaylists = 16
        netPlaylist = 17
        netContainerId = 18
    End Enum
    Public Enum Player_status_change
        SongChanged = 1         'raises whenever the current song changes
        PlayStatusChanged = 2   'raises when pause, stop, play, etc. pressed.
        PlayList = 3            'raises whenever the current playlist changes
        Library = 4             'raises when the library changes
        DeviceStatusChanged = 11 'raised when the player goes on/off-line or an iPod is inserted/removed from the wireless dock
        AlarmStart = 12          ' raised when the alarm goes off
        ConfigChange = 13        ' raised when the configuration of a device changes like alarm info being modified
        NextSong = 14            ' raised when the next song is about to start
        AnnouncementChange = 15
    End Enum

    Public Enum Player_state_values
        Unknown = 0
        Playing = 1
        Stopped = 2
        Paused = 3
        Forwarding = 4
        Rewinding = 5
        Transistioning = 6
        Docked = 11         ' new state to support WD100 devices being docked
        Undocked = 12       ' new state to support WD100 devices being undocked
        AudioInTrue = 13    ' new state to indicating Line-In went Up (or connected)
        AudioInFalse = 14   ' new state to indicating Line-In went down (or disconnected)
        ZoneName = 15
        ReplicationState = 16
        UpdateHSServerOnly = 17
        Online = 18
        Offline = 19
        AnnouncementStart = 20
        AnnouncementStop = 21
        Sleeping = 22
        PoweredOff = 23
        Isolated = 24
        onBlueTooth = 25
        Upgrading = 26
        LowBattery = 27
    End Enum

    Public Enum DeviceStateValues
        Unknown = 0
        Offline = 1
        Online = 2
        Sleeping = 3
        PoweredOff = 4
        Isolated = 5
        onBlueTooth = 6
        Upgrading = 7
        LowBattery = 8
    End Enum

    Public Enum PlayerTriggerEvents
        SongChanged = 1         'raises whenever the current song changes
        PlayStatusChanged = 2   'raises when pause, stop, play, etc. pressed.
        PlayList = 3            'raises whenever the current playlist changes
        Library = 4             'raises when the library changes
        DeviceStatusChanged = 11 'raised when the player goes on/off-line or an iPod is inserted/removed from the wireless dock
        AlarmStart = 12          ' raised when the alarm goes off
        ConfigChange = 13        ' raised when the configuration of a device changes like alarm info being modified
        NextSong = 14            ' raised when the next song is about to start
        AnnouncementChange = 15
        VolumeUp = 16
        VolumeDown = 17
        AlarmStop = 18
        OnLine = 19
        OffLine = 20
        Sleeping = 21
        PowerOff = 22
        onBluetooth = 23
        AnnouncementStart = 24
        AnnouncementStop = 25
        playing = 26
        paused = 27
        stopped = 28
        lineinConnected = 29
        lineinDisconnected = 30
        transitioning = 31
        LowBattery = 32
    End Enum

    Public ReadOnly Property TriggerCommandSelectOptions As Dictionary(Of PlayerTriggerEvents, String)
        Get
            Return New Dictionary(Of PlayerTriggerEvents, String) From {
            {PlayerTriggerEvents.PlayStatusChanged, "play status change"},
            {PlayerTriggerEvents.SongChanged, "is changing Tracks"},
            {PlayerTriggerEvents.playing, "starts Playing"},
            {PlayerTriggerEvents.stopped, "Stops"},
            {PlayerTriggerEvents.paused, "is being Paused"},
            {PlayerTriggerEvents.transitioning, "is Transitioning"},
            {PlayerTriggerEvents.VolumeUp, "Volume Up"},
            {PlayerTriggerEvents.VolumeDown, "Volume Down"},
            {PlayerTriggerEvents.lineinConnected, "Line-in Connected"},
            {PlayerTriggerEvents.lineinDisconnected, "Line-in Disconnected"},
            {PlayerTriggerEvents.ConfigChange, "has Config Change"},
            {PlayerTriggerEvents.OnLine, "goes Online"},
            {PlayerTriggerEvents.OffLine, "goes Offline"},
            {PlayerTriggerEvents.Sleeping, "goes to Sleeping"},
            {PlayerTriggerEvents.PowerOff, "goes to Powered Off"},
            {PlayerTriggerEvents.LowBattery, "goes to Low Battery"},
            {PlayerTriggerEvents.onBluetooth, "switches to Bluetooth"},
            {PlayerTriggerEvents.NextSong, "Next Track Change"},
            {PlayerTriggerEvents.AnnouncementStart, "Announcement Start"},
            {PlayerTriggerEvents.AnnouncementStop, "Announcement Stop"},
            {PlayerTriggerEvents.AlarmStart, "Alarm Start"}
            }
        End Get
    End Property

    Public Enum Repeat_modes
        repeat_off = 0
        repeat_one = 1
        repeat_all = 2
    End Enum

    Public Enum Shuffle_modes
        Shuffled = 1
        Ordered = 2
        Sorted = 3
    End Enum

    Public Enum QueueActions
        qaDontPlay
        qaPlayLast
        qaPlayNext
        qaPlayNow
    End Enum

    Public Enum MyLibraryTypes
        LibraryQueue = 1
        LibraryDB = 2
    End Enum


    Public Const MaxPlayerTOActionArray = 10
    '
    ' Timeout indexes
    Public Const TOReachable = 0
    Public Const TOPositionUpdate = 1
    '
    ' Timeout Values
    Public Const TOReachableValue = 10
    Public Const ToPositionUpdateValue = 1

    Public Enum SonosHSDevices
        Player = 0
        Status = 1
        Control = 2
        Repeat = 3
        Shuffle = 4
        Volume = 5
        Balance = 6
        Mute = 7
        Loudness = 8
        Tittle = 9
        NexTittle = 10
        Artist = 11
        NextArtist = 12
        Album = 13
        NextAlbum = 14
        Art = 15
        NextArt = 16
        TrackLength = 17
        TrackPosition = 18
        RadioStationName = 19
        TrackDescriptor = 20
        Genre = 21
        TrackDate = 22
        Master = 23
        RootDevice = 24
    End Enum

    Public Enum Player_selections
        Playlist_Track = 0
        Artist_Album_Track = 1
        Artist_Track = 2
        Album_Track = 3
        Playlist = 4
        Artist_Album = 5
        Album = 6
        Artist = 7
        Genre = 8
        audiobook = 9
        podcast = 10
        LineInput = 11
    End Enum

    Public Enum AnnouncementState
        asIdle = 0
        asLinking = 1
        asLinked = 2
        asSpeaking = 3
        asFilePlayed = 4
        asUnlinking = 5
        asUnDefined = 6
    End Enum

    <Serializable()>
    Public Class MusicDBItems
        Public Sub New()
            MyBase.New()
        End Sub
        Public Genres As Boolean = True
        Public Tracks As Boolean = True
        Public Artists As Boolean = True
        Public Albums As Boolean = True
        Public Playlists As Boolean = True
        Public Radiostations As Boolean = True
        Public Audiobooks As Boolean = True
        Public Podcasts As Boolean = True
    End Class

    <Serializable()>
    Public Class AnnouncementItems
        Public Sub New()
            MyBase.New()
        End Sub
        Public LinkGroupName As String = ""
        Public device As Short = 0
        Public text As String = ""
        Public host As String = ""
        Public IsFile As Boolean = False
        Public State_ As AnnouncementState = AnnouncementState.asIdle
        Public Next_ As Object = Nothing
        Public Previous_ As Object = Nothing
#If HS3 = "True" Then
        Public SourceZoneMusicAPI As HSPI = Nothing
#Else
        Public SourceZoneMusicAPI As MediaDevice
        Public destinationZonesUDNs As String = ""
        Public announcementInfoXML As String = ""
#End If
        Public DelayinSec As Integer = 0
        Public AbsoluteTime As DateTime
    End Class


    <Serializable()> Public Structure IPAddressInfo
        Public IPAddress As String
        Public IPPort As String
    End Structure

    <Serializable()> Public Class PingArrayElement
        Public IPAddress As String
        Public Status As String
        Public FailedPingCount As Integer
        Public ClientNameList As String()
    End Class


    Public Enum DebugLevel
        dlOff = 0
        dlErrorsOnly = 1
        dlEvents = 2
        dlVerbose = 3
    End Enum

    Public DebugLevelConstants As String() = {"Off", "Errors Only", "Events and Errors", "Verbose"}

    Public NavigationTabConstants As String() = {"General", "Debug", "Grouptable", "Playertable", "Network", "UPNP", "Register", "Help"}

    Public stateSelectConstants As String() = {"Legacy", "State", "Track", "Track & Artist"}

    Public Enum ManagePingActions
        mpAdd
        mpRemove
        mpTimeOut
        mpStatus
    End Enum

    Public Enum DeviceInfoIndex
        diGivenName = 0
        diUDN = 1
        diDeviceModelName = 2
        diDeviceType = 3
        diFriendlyName = 4
        diIPAddress = 5
        diIPPort = 6
        diMusicAPIIndex = 7
        diAdminState = 8
        diDeviceIsAdded = 9
        diHSDeviceCode = 10
        diDeviceAPIIndex = 11
        diTimeBetweenPictures = 12
        diArtistObjectID = 13
        diAlbumObjectID = 14
        diTrackObjectID = 15
        diGenreObjectID = 16
        diPlayListObjectID = 17
        diPollTransportChanges = 18
        diRepeat = 19
        diShuffle = 20
        diServerUDN = 21
        diSystemUpdateID = 22
        diPictureSize = 23
        diPollVolumeChanges = 24
        diRemoteControl = 25
        diRegistered = 26
        diRemoteType = 27
        diNextAV = 28
        diUseNextAV = 29
        diAnnouncementMP3 = 30
        diMusicObjectID = 31
        diPhotosObjectID = 32
        diVideosObjectID = 33
        diSystemUpdateIDAtDBCreation = 34
        diMACAddress = 35
        ' these are for HS3
        diTrackHSRef = 36
        diNextTrackHSRef = 37
        diArtistHSRef = 38
        diNextArtistHSRef = 39
        diAlbumHSRef = 40
        diNextAlbumHSRef = 41
        diArtHSRef = 42
        diNextArtHSRef = 43
        diPlayStateHSRef = 44
        diVolumeHSRef = 45
        diMuteHSRef = 46
        diLoudnessHSRef = 47
        diBalanceHSRef = 48
        diTrackLengthHSRef = 49
        diTrackPosHSRef = 50
        diRadiostationNameHSRef = 51
        diTrackDescrHSRef = 52
        diRepeatHSRef = 53
        diShuffleHSRef = 54
        diPlayerHSRef = 55
        diGenreHSRef = 56
        ' These specific for Sonos
        diSonosPlayerType = 57
        diSonosReplicationInfo = 58
        diDeviceIConURL = 59
        diRoomIcon = 60
        diDockDeviceNameHSRef = 61
        diArtFileIndex = 62
        diNextArtFileIndex = 63
        diControlHSRef = 64

        diHidePlayer = 65
        diHidePlayerPaired = 66
        diBatteryHSRef = 67
        diHouseholdId = 68
        diMuseHouseholdId = 69
        diWebsocketKey = 70
        diGroupInfo = 71
        diVanishedReason = 72
        diLastKnownIPAddress = 73
        diWOLHSRef = 74
        diSummaryHSRef = 75
        diNightModeHSRef = 76
        diDialogLevelHSRef = 77



        ' because this is shared with Sonos v3 and MediaController, I need to make sure they don't overlap
        diInstanceDebugFlag = 201
        diInstanceDebugParams = 202
        diUPnPDebugParams = 203
        diUPnPDebugFlag = 204
    End Enum

    Public Enum displayInfoIndex
        diSettingsinstanceflagsselection = 1
        diSettingsupnpinstanceflagsselection = 2
        diSettingsnavbarselection = 3
    End Enum

    Public Enum apiInfoIndex
        aiOAuthCode = 1
        aiClientID = 2
        aiClientSecret = 3
        aiRefreshtoken = 4
        aiTokenExpiry = 5
        aiAccessToken = 6
        aiOAuthCodeTimeStamp = 7
        aiHSFeatRef = 8
    End Enum

    Public Enum settingInfoIndex
        siMasterHSDeviceRef = 1
        siMasterHSFeatureRef = 2
        gWebSvrPort = 3
        gServerAddressBind = 4
    End Enum

    Public Enum PostAnnouncementAction
        paaAlwaysForward = 0
        paaForwardNoMatch = 1
        paaAlwaysDrop = 2
    End Enum

    Public Enum HSSTrackLengthSettings
        TLSSeconds = 0
        TLSHoursMinutesSeconds = 1
    End Enum

    Public Enum HSSTrackPositionSettings
        TPSSeconds = 0
        TPSHoursMinutesSeconds = 1
        TPSPercentage = 2
    End Enum

    <Flags>
    Enum PlayerAttributesSet
        none = 0
        canPlayTV = 1
        hasAudioInput = 2
        excludeOwn = 4
        ' = 4
        ' = 8
    End Enum

    Public Enum stateChoices
        scLegacy = 0
        scState = 1
        scTrack = 2
        scTrackArtist = 3
    End Enum
End Module

<Serializable>
Class playertableinfo
    Public devicename As String
    Public deviceudn As String
    Public deviceonline As String
    Public devicemodel As String
    Public deviceswgen As String
    Public deviceipaddress As String
    Public deviceinconsistency As Boolean
    Public devicepairslave As Boolean
    Public instancetraceflag As Boolean
    Public instancedebugparms As String
    Public upnptraceflag As Boolean
    Public upnpdebugparms As String
    Public household As String
    Public supportaudioclip As Boolean
    Public supportaudioin As Boolean
    Public nbrOfDisconnects As Integer
End Class

<Serializable>
Class SelectionInfo
    Public key As String
    Public value As String
    Public selected As Boolean
End Class

<Serializable>
Class TabInfo
    Public id As String
    Public name As String
    Public selected As String
End Class

<Serializable>
Class UPNPDeviceInfo
    Public deviceicon As String
    Public devicefriendlyname As String
    Public deviceudn As String
    Public deviceonline As String
    Public devicemodel As String
    Public deviceipaddress As String
    Public deviceinconsistency As Boolean
    Public deviceipport As String
    Public devicelocation As String
    Public devicetype As String
    Public Errrors412 As Integer = 0
End Class


<Serializable>
Class NetworkInfo
    Public description As String
    Public mac As String
    Public operationalstate As String
    Public id As String
    Public name As String
    Public interfacetype As String
    Public addressinfo As List(Of NetworkInfoAddrMask)
End Class

<Serializable>
Class NetworkInfoAddrMask
    Public address As String
    Public mask As String
    Public addrtype As String
End Class

<Serializable()>
Public Class MyUPnpDeviceInfo
    Public ZoneName As String
    Public ZoneModel As String
    Public ZoneUDN As String
    Public ZonePlayerRef As Integer
    Public ZoneControlFeatureRef As Integer
    Public ZoneOnLine As Boolean
#If HS3 = "True" Then
    Public ZonePlayerControllerRef As HSPI
    public ZoneDeviceAPIIndex as integer = 0
#Else
    Public ZonePlayerControllerRef As MediaDevice
#End If
    Public ZoneWeblinkCreated As Boolean
    Public ZoneHasInput As Boolean
    Public ZoneIsHSInput As Boolean
    Public ZoneAdminState As Boolean
    Public Device As MyUPnPDevice
    Public UPnPDeviceIPAddress As String = ""
    Public UPnPDeviceAdminStateActive As Boolean = False
    Public UPnPDeviceIsAddedToHS As Boolean = False
    Public ZoneDeviceIconURL As String = ""
    Public ZoneCurrentIcon As String = ""

    Public Sub New()
        MyBase.New()
        ZoneName = ""
        ZoneModel = ""
        ZoneUDN = ""
        ZonePlayerRef = -1
        ZoneControlFeatureRef = -1
        ZoneOnLine = False
        ZonePlayerControllerRef = Nothing
        ZoneWeblinkCreated = False
        ZoneHasInput = False
        ZoneIsHSInput = False
        ZoneAdminState = True
        Device = Nothing
        UPnPDeviceAdminStateActive = False
        UPnPDeviceIsAddedToHS = False
        UPnPDeviceIPAddress = ""
        ZoneDeviceIconURL = ""
        ZoneCurrentIcon = ""
    End Sub

    Public Sub Close()
        ZoneName = ""
        ZoneModel = ""
        ZoneUDN = ""
        ZonePlayerRef = -1
        ZoneControlFeatureRef = -1
        ZoneOnLine = False
        ZonePlayerControllerRef = Nothing
        ZoneWeblinkCreated = False
        ZoneHasInput = False
        ZoneIsHSInput = False
        ZoneAdminState = True
        Device = Nothing
        UPnPDeviceAdminStateActive = False
        UPnPDeviceIsAddedToHS = False
        UPnPDeviceIPAddress = ""
        ZoneDeviceIconURL = ""
        ZoneCurrentIcon = ""
    End Sub

End Class

<Serializable()>
Public Class SavedPlayerList
    Public MyLinkListIndex As Integer = 0
    Public MyPlayerUDNArray() As String
    Public MyCurrentLinkListIndex As Integer = 0


    Public Sub New()
        MyBase.New()
        FlushInfo()
    End Sub

    Public Sub FlushInfo()
        MyLinkListIndex = 0
        MyCurrentLinkListIndex = 0
        Try
            Erase MyPlayerUDNArray
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SavedPlayerList.FlushInfo with error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try

    End Sub

    Public Sub ResetIndex()
        MyCurrentLinkListIndex = 0
    End Sub

    Public Function GetNext() As String
        If MyLinkListIndex = 0 Then
            Return ""
        End If
        If MyCurrentLinkListIndex >= MyLinkListIndex Then
            Return ""
        End If
        GetNext = MyPlayerUDNArray(MyCurrentLinkListIndex)
        MyCurrentLinkListIndex += 1
    End Function

    Public Sub Add(ByVal PlayerUDN As String)
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SavedPlayerList.add called with PlayerUDN = " & PlayerUDN & " And MyLinkListIndex = " & MyLinkListIndex.ToString, LogType.LOG_TYPE_INFO)
        Dim TempIndex As Integer
        If MyLinkListIndex = 0 Then
            ReDim Preserve MyPlayerUDNArray(MyLinkListIndex)
            MyPlayerUDNArray(MyLinkListIndex) = PlayerUDN
            MyLinkListIndex += 1
        Else
            For TempIndex = 0 To MyLinkListIndex - 1
                If MyPlayerUDNArray(TempIndex) = PlayerUDN Then
                    ' the Index is already stored
                    Exit Sub
                End If
            Next
            ' the Index doesn't exist store it
            ReDim Preserve MyPlayerUDNArray(MyLinkListIndex)
            MyPlayerUDNArray(MyLinkListIndex) = PlayerUDN
            MyLinkListIndex += 1
        End If
    End Sub

    Public Function IsAlreadyStored(PlayerUDN As String) As Boolean
        IsAlreadyStored = False
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SavedPlayerList.IsAlreadyStored called with PlayerUDN = " & PlayerUDN & " And MyLinkListIndex = " & MyLinkListIndex.ToString, LogType.LOG_TYPE_INFO)
        If MyLinkListIndex = 0 Then Exit Function
        Try
            Dim TempIndex As Integer
            For TempIndex = 0 To MyLinkListIndex - 1
                If MyPlayerUDNArray(TempIndex) = PlayerUDN Then
                    ' the Index is already stored
                    Return True
                End If
            Next
        Catch ex As Exception
            If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("Error in SavedPlayerList.IsAlreadyStored with PlayerUDN = " & PlayerUDN & " And MyLinkListIndex = " & MyLinkListIndex.ToString & " And Error = " & ex.Message, LogType.LOG_TYPE_ERROR)
        End Try
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SavedPlayerList.IsAlreadyStored for PlayerUDN = " & PlayerUDN & " didn't find anything stored", LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetLastUDN() As String   ' note this procedure also removes the last player info
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SavedPlayerList.GetLastIndex called" & " and MyLinkListIndex = " & MyLinkListIndex.ToString, LogType.LOG_TYPE_INFO)
        GetLastUDN = ""
        If MyLinkListIndex = 0 Then Exit Function
        GetLastUDN = MyPlayerUDNArray(MyLinkListIndex - 1)
        MyLinkListIndex -= 1
        If MyLinkListIndex <= 0 Then
            FlushInfo()
        Else
            ReDim Preserve MyPlayerUDNArray(MyLinkListIndex)
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("SavedPlayerList.GetLastIndex called and returned UDN = " & GetLastUDN, LogType.LOG_TYPE_INFO)
    End Function
End Class

<Serializable()>
Public Class LinkGroupInfo
    Public MyLinkGroupName As String = ""
    Public MySavedPlayerList As SavedPlayerList
    Public Sub New()
        MyBase.New()
        MySavedPlayerList = New SavedPlayerList
    End Sub
    Public Sub FlushInfo()
        MySavedPlayerList = Nothing
    End Sub
End Class

<Serializable()>
Public Class LinkGroupInfoArray
    Public MyLinkListIndex As Integer = 0
    Public MyLinkGroupInfoArray()

    Public Sub New()
        MyBase.New()
        MyLinkListIndex = 0
        MyLinkGroupInfoArray = Nothing
    End Sub

    Public Sub FlushInfo()
        Dim TempIndex As Integer
        For TempIndex = 0 To MyLinkListIndex - 1
            MyLinkGroupInfoArray(TempIndex).flushinfo()
            MyLinkGroupInfoArray(TempIndex) = Nothing
        Next
        MyLinkListIndex = 0
        Erase MyLinkGroupInfoArray
    End Sub

    Public Function AddLinkGroupInfo(ByVal LinkgroupName As String) As LinkGroupInfo
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddLinkGroupInfo called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        AddLinkGroupInfo = Nothing
        Dim TempIndex As Integer
        If MyLinkListIndex = 0 Then
            ReDim Preserve MyLinkGroupInfoArray(MyLinkListIndex)
            MyLinkGroupInfoArray(MyLinkListIndex) = New LinkGroupInfo With {
                .MyLinkGroupName = LinkgroupName
            }
            AddLinkGroupInfo = MyLinkGroupInfoArray(MyLinkListIndex)
            MyLinkListIndex += 1
        Else
            For TempIndex = 0 To MyLinkListIndex - 1
                If MyLinkGroupInfoArray(TempIndex).MyLinkgroupName = LinkgroupName Then
                    AddLinkGroupInfo = MyLinkGroupInfoArray(TempIndex)
                    Exit Function
                End If
            Next
            ' the Index doesn't exist store it
            ReDim Preserve MyLinkGroupInfoArray(MyLinkListIndex)
            MyLinkGroupInfoArray(MyLinkListIndex) = New LinkGroupInfo With {
                .MyLinkGroupName = LinkgroupName
            }
            AddLinkGroupInfo = MyLinkGroupInfoArray(MyLinkListIndex)
            MyLinkListIndex += 1
        End If
        If piDebuglevel > DebugLevel.dlErrorsOnly Then Log("AddLinkGroupInfo done for LinkgroupName = " & LinkgroupName & ". LinkListIndex = " & MyLinkListIndex, LogType.LOG_TYPE_INFO)
    End Function

    Public Function GetLinkGroupInfo(ByVal LinkgroupName As String) As LinkGroupInfo
        If piDebuglevel > DebugLevel.dlEvents Then Log("GetLinkGroupInfo called with LinkgroupName = " & LinkgroupName, LogType.LOG_TYPE_INFO)
        GetLinkGroupInfo = Nothing
        If MyLinkListIndex <> 0 Then
            Dim TempIndex As Integer
            For TempIndex = 0 To MyLinkListIndex - 1
                If MyLinkGroupInfoArray(TempIndex).MyLinkgroupName = LinkgroupName Then
                    GetLinkGroupInfo = MyLinkGroupInfoArray(TempIndex)
                    Exit Function
                End If
            Next
        End If
        ' entry did not exist, add it
        GetLinkGroupInfo = AddLinkGroupInfo(LinkgroupName)
    End Function



End Class

<Serializable>
Public Class TrackItemInfo
    Public artist As String = ""
    Public album As String = ""
    Public title As String = ""
    Public currentURI As String = ""
    Public albumArtURI As String = NoArtPath
    Public trackPosition As String = ""
    Public trackDuration As String = ""
    Public queuePosition As String = ""
    Public playedPercentage As String = ""
    Public source As String = ""
    Public internetRadio As Boolean = False
    Public radiostationOrService As String = ""
    Public currentURIMetaData As String = ""
    Public radioCurrentURI As String = ""
    Public radioCurrentMetaData As String = ""
    Public streamContent As String = ""
    Public radioShow As String = ""
End Class

<Serializable>
Public Class TransportInfo
    Public transportState As String = "N/A"
    Public currentTrack As String = "N/A"
    Public currentTrackDuration As String = "N/A"
    'Public currentTrackNbr As String = "N/A"
    Public currentTrackURI As String = "N/A"
    Public currentTrackMetaData As String = "N/A"
    Public currentPlayMode As String = "N/A"
    Public currentSection As String = "N/A"

    Public nbrOfTracks As String = "N/A"
    Public artist As String = "N/A"
    Public track As String = "N/A"
    Public album As String = "N/A"
    Public art As String = NoArtPath
    Public protocolInfo As String = "N/A"
    Public duration As String = "N/A"
    Public upnpClass As String = "N/A"

    Public enqueuedTransportURI As String = "N/A"
    Public enqueuedTransportURIMetaData As String = "N/A"
    Public enqArtist As String = "N/a"
    Public enqTitle As String = "N/A"   ' I believe the playlist name when loaded in a queue or audiobroadcast name
    Public enqAlbum As String = "N/A"
    Public enqArt As String = NoArtPath
    Public enqProtocolInfo As String = "N/A"
    Public enqDuration As String = "N/A"
    Public enqUpnpClass As String = "N/A"
    Public enqSourceName As String = "N/A"
    Public enqDescription As String = "N/A" ' sometimes station name appears here, especially for amazon music

    Public avTransportURI As String = "N/A"
    Public avTransportURIMetaData As String = "N/A"

    ' from nextMetaData
    Public nextArtist As String = "N/A"
    Public nextTrack As String = "N/A"
    Public nextAlbum As String = "N/A"
    Public nextArt As String = NoArtPath
    Public nextProtocolInfo As String = "N/A"
    Public nextDuration As String = "N/A"
    Public NextUpnpClass As String = "N/A"
    ' from analyzing the protocol info etc
    Public source As String = "Unknown"
    Public internetRadio As Boolean = False
    'Public radiostationOrService As String = ""
    Public radioCurrentURI As String = "N/A"
    Public radioCurrentMetaData As String = "N/A"
    Public streamContent As String = "N/A"
    Public radioShow As String = "N/A"
End Class

<Serializable>
Public Class AVTransportInfo
    Public artist As String = ""
    Public title As String = ""
    Public album As String = ""
    Public art As String = NoArtPath
    Public desc As String = ""
    Public upnpClass As String = ""
    Public albumArtist As String = ""
    Public description As String = ""
    Public source As String = "Unknown"
    Public internetRadio As Boolean = False
    Public streamContent As String = ""
End Class

Public Class playerControlLastSettings
    Public Duration As Integer = 0
    Public LinkState As String = ""
    Public PlayerState As Player_state_values
    Public Volume As Integer = 0
    Public Track As String = ""
    Public Album As String = ""
    Public Artist As String = ""
    Public Art As String = ""
    Public RadioStation As String = ""
    Public NextTrack As String = ""
    Public NextAlbum As String = ""
    Public NextArtist As String = ""
    Public Mute As Boolean = False
    Public Shuffle As Boolean = False
    Public Repeat As Boolean = False
    Public TrackPosition As Integer = 0
    Public TrackLength As Integer = 0
    Public queueHasChanged As Boolean = False
End Class

Public Class navControlList
    Public navigationInfo As List(Of NavItem) = New List(Of NavItem)
End Class


<Serializable>
Public Class ArtItem
    Public ArtUri As String = ""
    Public nextTrack As Boolean = False
End Class

#If HS3 = "" Then   ' JSON definitions
<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class UpdateDivInfoList
    <JsonProperty(PropertyName:="result")> Public Property result As String = "updatedivs"
    <JsonProperty(PropertyName:="actionInfo")> Public Property actionInfo As String = ""
    <JsonProperty(PropertyName:="updateDivs")> Public Property updateDivs As UpdateDivInfo()
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class UpdateDivInfo
    <JsonProperty(PropertyName:="divId")> Public Property divId As String = ""
    <JsonProperty(PropertyName:="updateDivInfo")> Public Property updateDivInfo As String = ""
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class SourcePlayersInfo
    <JsonProperty(PropertyName:="result")> Public Property result As String = ""
    <JsonProperty(PropertyName:="sourcePlayers")> Public Property sourcePlayers As SourcePlayerInfo()
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class SourcePlayerInfo
    <JsonProperty(PropertyName:="index")> Public Property index As String = ""
    <JsonProperty(PropertyName:="sourcePlayerXML")> Public Property sourcePlayerXML As String = ""
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Public Class LGInfo
    <JsonProperty(PropertyName:="name")> Public Property name As String = ""
    <JsonProperty(PropertyName:="sourcePlayer")> Public Property sourcePlayer As String = ""
    <JsonProperty(PropertyName:="ttsOption")> Public Property ttsOption As String = "ttsregular"
    <JsonProperty(PropertyName:="destinationInfo")> Public destinationInfo As LGDestinationInfo()
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Public Class LGDestinationInfo
    <JsonProperty(PropertyName:="udn")> Public Property udn As String = ""
    <JsonProperty(PropertyName:="muteOverride")> Public Property muteOverride As Boolean = False
    <JsonProperty(PropertyName:="volume")> Public Property volume As String = ""
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class htmlClientJSONResponse
    <JsonProperty(PropertyName:="reason")> Public Property reason As String = ""
    <JsonProperty(PropertyName:="selection")> Public Property selection As String = ""
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class htmlServerJSONResponse
    <JsonProperty(PropertyName:="result")> Public Property result As String = ""
    <JsonProperty(PropertyName:="resultString")> Public Property resultString As String = ""
    <JsonProperty(PropertyName:="migrationFileName")> Public Property migrationFileName As String = ""
End Class

' definitions for HS state device
<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class PlayerInfo
    <JsonProperty(PropertyName:="state")> Public Property state As String
    <JsonProperty(PropertyName:="source")> Public Property source As String = ""
    <JsonProperty(PropertyName:="sourceExt")> Public Property sourceExt As String = ""
    <JsonProperty(PropertyName:="currentTrackInfo")> Public Property currentTrackInfo As TrackInfo
    <JsonProperty(PropertyName:="nextTrackInfo")> Public Property nextTrackInfo As TrackInfo
    <JsonProperty(PropertyName:="pairingInfo")> Public Property pairingInfo As PairingInfo
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class TrackInfo
    <JsonProperty(PropertyName:="track")> Public Property track As String = ""
    <JsonProperty(PropertyName:="album")> Public Property album As String = ""
    <JsonProperty(PropertyName:="artist")> Public Property artist As String = ""
    <JsonProperty(PropertyName:="art")> Public Property art As String = ""
End Class

<JsonObject(MemberSerialization.OptIn)>
<Serializable()>
Class PairingInfo
    <JsonProperty(PropertyName:="state")> Public Property state As String = ""
    <JsonProperty(PropertyName:="UDNsOtherPlayers")> Public Property UDNsOtherPlayers As String = ""
    Enum PairStateEnum
        pairedTo    ' slave
        pairedWith  ' master
        linkedTo    ' slave of group
        linkedWith  ' maser of group
    End Enum

End Class



#End If  ' JSON definitions

<Serializable()>
Public Class MyUPnPControlCallback
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
Public Class MyUPnPAVTransportCallback
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
Public Class MyUPnPRenderingControlCallback
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
Public Class MyUPnPAlarmClockCallback
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
Public Class MyUPnPMusicServicesCallback
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
Public Class MyUPnPSystemPropertiesCallback
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
Public Class MyUPnPZonegroupTopologyCallback
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
Public Class MyUPnPGroupManagementCallback
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
Public Class MyUPnPConnectionManagerCallback
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
Public Class MyUPnPContentDirectoryCallback
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
Public Class MyUPnPAudioInCallback
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
Public Class MyUPnPDevicePropertiesCallback
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
Public Class MyUPnPQueueServiceCallback
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