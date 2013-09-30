Attribute VB_Name = "modGuild"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'Max Members Per Guild
Public Const MAX_GUILD_MEMBERS As Long = 50
'Max Ranks Guilds Can Have
Public Const MAX_GUILD_RANKS As Long = 6
'Max Different Permissions
Public Const MAX_GUILD_RANKS_PERMISSION As Long = 6
'Max guild save files(aka max guilds)
Public Const MAX_GUILD_SAVES As Long = 200

'Default Ranks Info
'1: Open Admin
'2: Can Recruit
'3: Can Kick
'4: Can Edit Ranks
'5: Can Edit Users
'6: Can Edit Options

Public Guild_Ranks_Premission_Names(1 To MAX_GUILD_RANKS_PERMISSION) As String
Public Default_Ranks(1 To MAX_GUILD_RANKS_PERMISSION) As Byte


'Max is set to MAX_PLAYERS so each online player can have his own guild
Public GuildData(1 To MAX_PLAYERS) As GuildRec

Public Type GuildRanksRec
    'General variables
    Used As Boolean
    Name As String
    
    'Rank Variables
    RankPermission(1 To MAX_GUILD_RANKS_PERMISSION) As Byte
End Type

Public Type GuildMemberRec
    'User login/name
    Used As Boolean
    
    User_Login As String
    User_Name As String
    Founder As Boolean
    
    Online As Boolean
    
    'Guild Variables
    Rank As Integer
    Comment As String * 100
     
End Type

Public Type GuildRec
    In_Use As Boolean
    
    Guild_Name As String
    Guild_Tag As String * 3
    
    'Guild file number for saving
    Guild_Fileid As Long
    
    Guild_Members(1 To MAX_GUILD_MEMBERS) As GuildMemberRec
    Guild_Ranks(1 To MAX_GUILD_RANKS) As GuildRanksRec
    
    'Message of the day
    Guild_MOTD As String * 100
    
    'The rank recruits start at
    Guild_RecruitRank As Integer
    'Color of guild name
    Guild_Color As Long
    Guild_Logo As Long
End Type
Public Sub Set_Default_Guild_Ranks()
    'Max sure this starts at 1 and ends at MAX_GUILD_RANKS_PERMISSION (Default 7)
    '0 = Cannot, 1 = Able To
    Guild_Ranks_Premission_Names(1) = "Open Admin"
    Default_Ranks(1) = 0
    
    Guild_Ranks_Premission_Names(2) = "Can Recruit"
    Default_Ranks(2) = 1
    
    Guild_Ranks_Premission_Names(3) = "Can Kick"
    Default_Ranks(3) = 0
    
    Guild_Ranks_Premission_Names(4) = "Can Edit Ranks"
    Default_Ranks(4) = 0
    
    Guild_Ranks_Premission_Names(5) = "Can Edit Users"
    Default_Ranks(5) = 0
    
    Guild_Ranks_Premission_Names(6) = "Can Edit Options"
    Default_Ranks(6) = 0
End Sub
Public Function GuildCheckName(Index As Long, MemberSlot As Long, AttemptCorrect As Boolean) As Boolean
Dim i As Integer

    If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(MemberSlot).User_Login = Player(Index).Login Then
        GuildCheckName = True
        Exit Function
    End If
    
    If AttemptCorrect = True Then
        If TempPlayer(Index).tmpGuildSlot > 0 And Player(Index).GuildFileId > 0 Then
            'did they get moved?
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).User_Login = Player(Index).Login Then
                    Player(Index).GuildMemberId = i
                    Call SavePlayer(Index)
                    GuildCheckName = True
                    Exit Function
                Else
                    Player(Index).GuildMemberId = 0
                End If
            Next
                
            'Remove from guild if we can't find them
            If Player(Index).GuildMemberId = 0 Then
                Player(Index).GuildFileId = 0
                TempPlayer(Index).tmpGuildSlot = 0
                Call SavePlayer(Index)
                PlayerMsg Index, "We can't seem to find you on your guilds member list this could mean a couple things.", BrightRed
                PlayerMsg Index, "1)They kicked you out   2)Your guild was deleted and replaced by another", BrightRed
            End If
        End If
    End If
    
    
    GuildCheckName = False


End Function
Public Sub MakeGuild(Founder_Index As Long, Name As String, Tag As String)
    Dim GuildSlot As Long
    Dim GuildFileId As Long
    Dim itemamount As Long
    Dim i As Integer
    Dim b As Integer
    
    If Player(Founder_Index).GuildFileId > 0 Then
        PlayerMsg Founder_Index, "You must leave your current guild before you make this one!", BrightRed
        Exit Sub
    End If
    
    
    
    GuildFileId = Find_Guild_Save
    GuildSlot = FindOpenGuildSlot
    
    'We are unable for an unknown reason
    If GuildSlot = 0 Or GuildFileId = 0 Then
        PlayerMsg Founder_Index, "Unable to make guild, sorry!", BrightRed
        Exit Sub
    End If
    
    If Name = "" Then
        PlayerMsg Founder_Index, "Your guild needs a name!", BrightRed
        Exit Sub
    End If
    
    'Change 1 to item number
    itemamount = HasItem(Founder_Index, 1)
    
    'Change 5000 to amount
    If itemamount = 0 Or itemamount < 5000 Then
        PlayerMsg Founder_Index, "Not enough Gold.", BrightRed
        Exit Sub
    End If
    
    'Change 1 to item number 5000 to amount
    TakeInvItem Founder_Index, 1, 5000
    
    GuildData(GuildSlot).Guild_Name = Name
    GuildData(GuildSlot).Guild_Tag = Tag
    GuildData(GuildSlot).Guild_MOTD = "Welcome to the " & Name & "!"
    GuildData(GuildSlot).In_Use = True
    GuildData(GuildSlot).Guild_Fileid = GuildFileId
    GuildData(GuildSlot).Guild_Members(1).Founder = True
    GuildData(GuildSlot).Guild_Members(1).User_Login = Player(Founder_Index).Login
    GuildData(GuildSlot).Guild_Members(1).User_Name = Player(Founder_Index).Name
    GuildData(GuildSlot).Guild_Members(1).Rank = MAX_GUILD_RANKS
    GuildData(GuildSlot).Guild_Members(1).Comment = "Guild Founder"
    GuildData(GuildSlot).Guild_Members(1).Used = True
    GuildData(GuildSlot).Guild_Members(1).Online = True
    GuildData(GuildSlot).Guild_Logo = RAND(1, MAX_GUILD_LOGO)

    'Set up Admin Rank with all permission which is just the max rank
    GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).Name = "Leader"
    GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).Used = True
    
    For b = 1 To MAX_GUILD_RANKS_PERMISSION
        GuildData(GuildSlot).Guild_Ranks(MAX_GUILD_RANKS).RankPermission(b) = 1
    Next
    
    'Set up rest of the ranks with default permission
    For i = 1 To MAX_GUILD_RANKS - 1
        GuildData(GuildSlot).Guild_Ranks(i).Name = "Rank " & i
        GuildData(GuildSlot).Guild_Ranks(i).Used = True
        
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            GuildData(GuildSlot).Guild_Ranks(i).RankPermission(b) = Default_Ranks(b)
        Next
    Next
    
    Player(Founder_Index).GuildFileId = GuildFileId
    Player(Founder_Index).GuildMemberId = 1
    TempPlayer(Founder_Index).tmpGuildSlot = GuildSlot
    
    
    'Save
    Call SaveGuild(GuildSlot)
    Call SavePlayer(Founder_Index)
    
    'Send to player
    Call SendGuild(False, Founder_Index, GuildSlot)
    
    'Inform users
    PlayerMsg Founder_Index, "Guild Successfully Created!", BrightGreen
    PlayerMsg Founder_Index, "Welcome to " & GuildData(GuildSlot).Guild_Name & ".", BrightGreen
    PlayerMsg Founder_Index, "Your Guild Logo Randomly [" & GuildData(GuildSlot).Guild_Logo & "].", BrightGreen
    PlayerMsg Founder_Index, "You can talk in guild chat with:  ;Message ", BrightRed
    
    'Update user for guild name display
    Call SendPlayerData(Founder_Index)

    
End Sub
Public Function CheckGuildPermission(Index As Long, Permission As Integer) As Boolean
Dim GuildSlot As Long

    'Get slot
    GuildSlot = TempPlayer(Index).tmpGuildSlot
    
    'Make sure we are talking about the same person
    If Not GuildData(GuildSlot).Guild_Members(Player(Index).GuildMemberId).User_Login = Player(Index).Login Then
        'Something went wrong and they are not allowed to do anything
        CheckGuildPermission = False
        Exit Function
    End If
    
    'If founder, true in every case
    If GuildData(GuildSlot).Guild_Members(Player(Index).GuildMemberId).Founder = True Then
        CheckGuildPermission = True
        Exit Function
    End If
    
    'Make sure this slot is being used aka they are still a member
    If GuildData(GuildSlot).Guild_Members(Player(Index).GuildMemberId).Used = False Then
        'Something went wrong and they are not allowed to do anything
        CheckGuildPermission = False
        Exit Function
    End If
    
    'Check if they are able to
    If GuildData(GuildSlot).Guild_Ranks(GuildData(GuildSlot).Guild_Members(Player(Index).GuildMemberId).Rank).RankPermission(Permission) = 1 Then
        CheckGuildPermission = True
    Else
        CheckGuildPermission = False
    End If
    
End Function
Public Sub Request_Guild_Invite(Index As Long, GuildSlot As Long, Inviter_Index As Long)

    If Player(Index).GuildFileId > 0 Then
        PlayerMsg Index, "You must leave your current guild before you can join " & GuildData(GuildSlot).Guild_Name & "!", BrightRed
        PlayerMsg Inviter_Index, "They are unable to join because they are already in a guild!", BrightRed
        Exit Sub
    End If

    If TempPlayer(Index).tmpGuildInviteSlot > 0 Then
        PlayerMsg Inviter_Index, "This user has a pending invite try again.", BrightRed
        Exit Sub
    End If

    'Permission 2 = Can Recruit
    If CheckGuildPermission(Inviter_Index, 2) = False Then
        PlayerMsg Inviter_Index, "Sorry your rank is not high enough!", BrightRed
        Exit Sub
    End If
    
    TempPlayer(Index).tmpGuildInviteSlot = GuildSlot
    '2 minute
    TempPlayer(Index).tmpGuildInviteTimer = timeGetTime + 120000
    TempPlayer(Index).tmpGuildInviteId = Player(Inviter_Index).GuildFileId
    
    PlayerMsg Inviter_Index, "Guild invite sent!", Green
    PlayerMsg Index, Player(Inviter_Index).Name & " has invited you to join the guild " & GuildData(GuildSlot).Guild_Name & "!", Green
End Sub
Public Sub Join_Guild(Index As Long, GuildSlot As Long)
Dim OpenSlot As Long

    OpenSlot = FindOpenGuildMemberSlot(GuildSlot)
        'Guild full?
        If OpenSlot > 0 Then
            'Set guild data
            GuildData(GuildSlot).Guild_Members(OpenSlot).Used = True
            GuildData(GuildSlot).Guild_Members(OpenSlot).User_Login = Player(Index).Login
            GuildData(GuildSlot).Guild_Members(OpenSlot).User_Name = Player(Index).Name
            GuildData(GuildSlot).Guild_Members(OpenSlot).Rank = GuildData(GuildSlot).Guild_RecruitRank
            GuildData(GuildSlot).Guild_Members(OpenSlot).Comment = "Joined: " & DateValue(Now)
            GuildData(GuildSlot).Guild_Members(OpenSlot).Online = True
            
            'Set player data
            Player(Index).GuildFileId = GuildData(GuildSlot).Guild_Fileid
            Player(Index).GuildMemberId = OpenSlot
            TempPlayer(Index).tmpGuildSlot = GuildSlot
            
            'Save
            Call SaveGuild(GuildSlot)
            Call SavePlayer(Index)
            
            'Send player guild data and display welcome
            Call SendGuild(True, Index, GuildSlot)
            PlayerMsg Index, "Welcome to " & GuildData(GuildSlot).Guild_Name & ".", BrightGreen
            
            PlayerMsg Index, "You can talk in guild chat with:  ;Message", BrightGreen
            
            'Update player to display guild name
            Call SendPlayerData(Index)
            
        Else
            'Guild full display msg
            PlayerMsg Index, "Guild is full sorry!", BrightRed
        End If
    
        
    
End Sub
Public Function Find_Guild_Save() As Long
Dim FoundSlot As Boolean
Dim Current As Integer
FoundSlot = False
Current = 1

Do Until FoundSlot = True
    
    If Not FileExist("\Data\guilds\Guild" & Current & ".dat") Then
        Find_Guild_Save = Current
        FoundSlot = True
    Else
        Current = Current + 1
    End If
    
    'Max Guild Files check
    If Current > MAX_GUILD_SAVES Then
        'send back 0 for no slot found
        Find_Guild_Save = 0
        FoundSlot = True
    End If
    
    
Loop

End Function
Public Function FindOpenGuildSlot() As Long
    Dim i As Integer
    
    For i = 1 To MAX_PLAYERS
        If GuildData(i).In_Use = False Then
            FindOpenGuildSlot = i
            Exit Function
        End If
        
        'No slot found how?
        FindOpenGuildSlot = 0
    Next
End Function
Public Function FindOpenGuildMemberSlot(GuildSlot As Long) As Long
Dim i As Integer
    
    For i = 1 To MAX_GUILD_MEMBERS
        If GuildData(GuildSlot).Guild_Members(i).Used = False Then
            FindOpenGuildMemberSlot = i
            Exit Function
        End If
    Next
    
    'Guild is full sorry bub
    FindOpenGuildMemberSlot = 0

End Function
Public Sub ClearGuildMemberSlot(GuildSlot As Long, MembersSlot As Long)
            GuildData(GuildSlot).Guild_Members(MembersSlot).Used = False
            GuildData(GuildSlot).Guild_Members(MembersSlot).User_Login = vbNullString
            GuildData(GuildSlot).Guild_Members(MembersSlot).User_Name = vbNullString
            GuildData(GuildSlot).Guild_Members(MembersSlot).Rank = 0
            GuildData(GuildSlot).Guild_Members(MembersSlot).Comment = vbNullString
            GuildData(GuildSlot).Guild_Members(MembersSlot).Founder = False
            GuildData(GuildSlot).Guild_Members(MembersSlot).Online = False
            
            'Save guild after we remove member
            Call SaveGuild(GuildSlot)
End Sub
Public Sub LoadGuild(GuildSlot As Long, GuildFileId As Long)
Dim i As Integer

'Does this file even exist?
If Not FileExist("\Data\guilds\Guild" & GuildFileId & ".dat") Then Exit Sub

    Dim filename As String
    Dim F As Long
    
        filename = App.Path & "\data\guilds\Guild" & GuildFileId & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , GuildData(GuildSlot)
        Close #F
        
        GuildData(GuildSlot).In_Use = True
        
        'Make sure an online flag didn't manage to slip through
        For i = 1 To MAX_GUILD_MEMBERS
            If GuildData(GuildSlot).Guild_Members(i).Online = True Then
                GuildData(GuildSlot).Guild_Members(i).Online = False
            End If
        Next
        
End Sub
Public Sub SaveGuild(GuildSlot As Long)

'Dont save unless a fileid was assigned
If GuildData(GuildSlot).Guild_Fileid = 0 Then Exit Sub


    Dim filename As String
    Dim F As Long
    
    filename = App.Path & "\data\guilds\Guild" & GuildData(GuildSlot).Guild_Fileid & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , GuildData(GuildSlot)
    Close #F
    
End Sub
Public Sub UnloadGuildSlot(GuildSlot As Long)
    'Save it first
    Call SaveGuild(GuildSlot)
    'Clear and reset for next use
    Call ClearGuild(GuildSlot)
End Sub
Public Sub ClearGuilds()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        Call ClearGuild(i)
    Next
End Sub
Public Sub ClearGuild(Index As Long)
    Call ZeroMemory(ByVal VarPtr(GuildData(Index)), LenB(GuildData(Index)))
    GuildData(Index).Guild_Name = vbNullString
    GuildData(Index).Guild_Tag = vbNullString
    GuildData(Index).In_Use = False
    GuildData(Index).Guild_Fileid = 0
    GuildData(Index).Guild_Color = 1
    GuildData(Index).Guild_RecruitRank = 1
    GuildData(Index).Guild_Logo = 0
End Sub
Public Sub CheckUnloadGuild(GuildSlot As Long)
Dim i As Integer
Dim UnloadGuild As Boolean

UnloadGuild = True

If GuildData(GuildSlot).In_Use = False Then Exit Sub

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                UnloadGuild = False
                Exit For
            End If
        End If
    Next
    
    If UnloadGuild = True Then
        Call UnloadGuildSlot(GuildSlot)
    End If
End Sub
Public Sub GuildKick(GuildSlot As Long, Index As Long, playerName As String)

Dim FoundOffline As Boolean
Dim IsOnline As Boolean
Dim OnlineIndex As Long
Dim MemberSlot As Long
Dim i As Integer
    
    
    
    OnlineIndex = FindPlayer(playerName)
    
    If OnlineIndex = Index Then
        PlayerMsg Index, "Can't kick your self!", BrightRed
        Exit Sub
    End If
    
    
    
    'If OnlineIndex > 0 they are online
    If OnlineIndex > 0 Then
        IsOnline = True
        
        If Player(OnlineIndex).GuildMemberId > 0 Then
            MemberSlot = Player(OnlineIndex).GuildMemberId
        Else
            'Prevent error, rest of this code assumes this is greater than 0
            Exit Sub
        End If
        
    Else
        IsOnline = False
    End If
    
    
    'Handle kicking online user
    If IsOnline = True Then
        If Not Player(Index).GuildFileId = Player(OnlineIndex).GuildFileId Then
            PlayerMsg Index, "User must be in your guild to kick them!", BrightRed
            Exit Sub
        End If
        
        If GuildData(GuildSlot).Guild_Members(MemberSlot).Founder = True Then
            PlayerMsg Index, "You cannot kick your founder!!", BrightRed
            Exit Sub
        End If
        
        Player(OnlineIndex).GuildFileId = 0
        Player(OnlineIndex).GuildMemberId = 0
        TempPlayer(OnlineIndex).tmpGuildSlot = 0
        Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
        PlayerMsg OnlineIndex, "You have been kicked from your guild!", BrightRed
        PlayerMsg Index, "Player kicked!", BrightRed
        Call SavePlayer(OnlineIndex)
        Call SaveGuild(GuildSlot)
        Call SendGuild(True, OnlineIndex, GuildSlot)
        Call SendGuild(True, Index, GuildSlot)
        Call SendPlayerData(OnlineIndex)
        Exit Sub
    End If
    
    
    
    'Handle Kicking Offline User
    FoundOffline = False
    If IsOnline = False Then
        'Lets Try to find them in the roster
        For i = 1 To MAX_GUILD_MEMBERS
            If playerName = Trim$(GuildData(GuildSlot).Guild_Members(i).User_Name) Then
                'Found them
                FoundOffline = True
                MemberSlot = i
                Exit For
            End If
        Next
        
        If FoundOffline = True Then
        
            If MemberSlot = 0 Then Exit Sub
            
            Call ClearGuildMemberSlot(GuildSlot, MemberSlot)
            Call SaveGuild(GuildSlot)
            PlayerMsg Index, "Offline user kicked!", BrightRed
            Exit Sub
        End If
        
        If FoundOffline = False And IsOnline = False Then
            PlayerMsg Index, "Could not find " & playerName & " online or offline in your guild.", BrightRed
        End If
    
    End If
 
End Sub
Public Sub GuildLeave(Index As Long)
    'This is for the leave command only, kicking has its own sub because it handles both online and offline kicks, while this only handles online.
    
    If Not Player(Index).GuildFileId > 0 Then
        PlayerMsg Index, "You must be in a guild to leave one!", BrightRed
        Exit Sub
    End If
    
    If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).Founder = True Then
        PlayerMsg Index, "The founder cannot leave or be kicked, founder status must first be transfered.", BrightRed
        PlayerMsg Index, "Use command /founder (name) to transfer this.", BrightRed
        Exit Sub
    End If
    
    'They match so they can leave
    If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).User_Login = Player(Index).Login Then
        
        'Clear guild slot
        Call ClearGuildMemberSlot(TempPlayer(Index).tmpGuildSlot, Player(Index).GuildMemberId)
        
        'Clear player data
        Player(Index).GuildFileId = 0
        Player(Index).GuildMemberId = 0
        TempPlayer(Index).tmpGuildSlot = 0
        
        'Update user for guild name display
        '''''''''''''''''''Call SendGuild(True, OnlineIndex, GuildSlot)
        '''''''''''''''''''Call SendGuild(True, Index, GuildSlot)
        Call SendPlayerData(Index)
        
        PlayerMsg Index, "You have left the guild.", BrightRed

    Else
        'They don't match this slot remove them
        Player(Index).GuildFileId = 0
        Player(Index).GuildMemberId = 0
        TempPlayer(Index).tmpGuildSlot = 0
    End If
    
    
End Sub
Public Sub GuildLoginCheck(Index As Long)
Dim i As Long
Dim GuildSlot As Long
Dim GuildLoaded As Boolean
GuildLoaded = False


    'Not in guild
    If Player(Index).GuildFileId = 0 Then Exit Sub
    
    'Check to make sure the guild file exists
    If Not FileExist("\Data\guilds\Guild" & Player(Index).GuildFileId & ".dat") Then
        'If guild was deleted remove user from guild
        Player(Index).GuildFileId = 0
        Player(Index).GuildMemberId = 0
        TempPlayer(Index).tmpGuildSlot = 0
        Call SavePlayer(Index)
        PlayerMsg Index, "Your guild has been deleted sorry!", BrightRed
        Exit Sub
    End If
    
    'First we need to see if our guild is loaded
    For i = 1 To MAX_PLAYERS
        If GuildData(i).In_Use = True Then
            'If its already loaded set true
            If GuildData(i).Guild_Fileid = Player(Index).GuildFileId Then
                GuildLoaded = True
                GuildSlot = i
                Exit For
            End If
        End If
    Next
    
    'If the guild is not loaded we need to load it
    If GuildLoaded = False Then
        'Find open guild slot, if 0 none
        GuildSlot = FindOpenGuildSlot
        If GuildSlot > 0 Then
            'LoadGuild
            Call LoadGuild(GuildSlot, Player(Index).GuildFileId)
            
        End If
    End If
    
    'Set GuildSlot
    TempPlayer(Index).tmpGuildSlot = GuildSlot
    
    'This is to prevent errors when we look for them
    If Player(Index).GuildMemberId = 0 Then Player(Index).GuildMemberId = 1

    'Make sure user didn't get kicked or guild was replaced by a different guild, both result in removal
    If GuildCheckName(Index, Player(Index).GuildMemberId, True) = False Then
        'unload if this user is not in this guild and it was loaded for this user
        If GuildLoaded = False Then
            Call UnloadGuildSlot(GuildSlot)
            Exit Sub
        End If
    End If
    
    'Sent data and set slot if all is good
    If Player(Index).GuildFileId > 0 Then
        'Set online flag
        GuildData(GuildSlot).Guild_Members(Player(Index).GuildMemberId).Online = True

        'send
        Call SendGuild(True, Index, GuildSlot)
        
        'Display motd
        If Not Trim$(GuildData(GuildSlot).Guild_MOTD) = vbNullString Then
            PlayerMsg Index, "Guild Motd: " & Trim$(GuildData(GuildSlot).Guild_MOTD), Blue
        End If
    End If
    
    
    
End Sub
Sub DisbandGuild(GuildSlot As Long, Index As Long)
Dim i As Integer
Dim tmpGuildSlot As Long
Dim TmpGuildFileId As Long
Dim filename As String

'Set some thing we need
tmpGuildSlot = GuildSlot
TmpGuildFileId = GuildData(tmpGuildSlot).Guild_Fileid

    'They are who they say they are, and are founder
    If GuildCheckName(Index, Player(Index).GuildMemberId, False) = True And GuildData(tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).Founder = True Then
        'File exists right?
         If FileExist("\Data\Guilds\Guild" & TmpGuildFileId & ".dat") = True Then
            'We have a go for disband
            'First we take everyone online out, this will include the founder people who login later will be kicked out then
            For i = 1 To Player_HighIndex
                If IsPlaying(i) = True Then
                    If Player(i).GuildFileId = TmpGuildFileId Then
                        'remove from guild
                        Player(i).GuildFileId = 0
                        Player(i).GuildMemberId = 0
                        TempPlayer(i).tmpGuildSlot = 0
                        Call SavePlayer(i)
                        'Send player data so they don't have name over head anymore
                        Call SendPlayerData(i)
                    End If
                End If
            Next
            
            'Unload Guild from memory
            Call UnloadGuildSlot(tmpGuildSlot)
            
            filename = App.Path & "\Data\Guilds\Guild" & TmpGuildFileId & ".dat"
            Kill filename
            
            
            PlayerMsg Index, "Guild disband done!", BrightGreen
         End If
    Else
        PlayerMsg Index, "Your not allowed to do that!", BrightRed
    End If
End Sub
Sub SendDataToGuild(ByVal GuildSlot As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If Player(i).GuildFileId = GuildData(GuildSlot).Guild_Fileid Then
                Call SendDataTo(i, Data)
            End If
        End If

    Next

End Sub

Sub SendGuild(ByVal SendToWholeGuild As Boolean, ByVal Index As Long, ByVal GuildSlot As Long)
    Dim Buffer As clsBuffer
    Dim i As Integer
    Dim b As Integer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSendGuild
    
    'General data
    Buffer.WriteString GuildData(GuildSlot).Guild_Name
    Buffer.WriteString GuildData(GuildSlot).Guild_Tag
    Buffer.WriteInteger GuildData(GuildSlot).Guild_Color
    Buffer.WriteString GuildData(GuildSlot).Guild_MOTD
    Buffer.WriteInteger GuildData(GuildSlot).Guild_RecruitRank
    Buffer.WriteLong GuildData(GuildSlot).Guild_Logo
    
    'Send Members
    For i = 1 To MAX_GUILD_MEMBERS
        Buffer.WriteString GuildData(GuildSlot).Guild_Members(i).User_Name
        Buffer.WriteInteger GuildData(GuildSlot).Guild_Members(i).Rank
        Buffer.WriteString GuildData(GuildSlot).Guild_Members(i).Comment
        Buffer.WriteByte GuildData(GuildSlot).Guild_Members(i).Online
    Next
    
    'Send Ranks
    For i = 1 To MAX_GUILD_RANKS
            Buffer.WriteString GuildData(GuildSlot).Guild_Ranks(i).Name
        For b = 1 To MAX_GUILD_RANKS_PERMISSION
            Buffer.WriteByte GuildData(GuildSlot).Guild_Ranks(i).RankPermission(b)
            Buffer.WriteString Guild_Ranks_Premission_Names(b)
        Next
    Next
    
    If SendToWholeGuild = False Then
        SendDataTo Index, Buffer.ToArray()
    Else
        SendDataToGuild GuildSlot, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub
Sub ToggleGuildAdmin(ByVal Index As Long, ByVal OpenAdmin As Boolean)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAdminGuild
    
    
    If OpenAdmin = True Then
        Buffer.WriteByte 1
    Else
        Buffer.WriteByte 0
    End If

        SendDataTo Index, Buffer.ToArray()

    
    Set Buffer = Nothing
End Sub
Sub SayMsg_Guild(ByVal GuildSlot As Long, ByVal Index As Long, ByVal message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString message
    Buffer.WriteString "[" & GuildData(GuildSlot).Guild_Tag & "]"
    Buffer.WriteLong saycolour
    
    SendDataToGuild GuildSlot, Buffer.ToArray()

    
    Set Buffer = Nothing
End Sub
Public Sub HandleGuildMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next
    
    If Not Player(Index).GuildFileId > 0 Then
        PlayerMsg Index, "You need to be in a guild to talk in Guild Chat!", BrightRed
        Exit Sub
    End If
    
    s = "[" & GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Tag & "]" & GetPlayerName(Index) & ": " & Msg
    
    Call SayMsg_Guild(TempPlayer(Index).tmpGuildSlot, Index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(Msg)
    
    Set Buffer = Nothing
End Sub
Public Sub HandleGuildSave(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

Dim Buffer As clsBuffer
Dim SaveType As Integer
Dim SentIndex As Integer
Dim HoldInt As Integer
Dim i As Integer


    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    
    SaveType = Buffer.ReadInteger
    SentIndex = Buffer.ReadInteger
    
    If SaveType = 0 Or SentIndex = 0 Then Exit Sub
    
    
    Select Case SaveType
    Case 1
        'options
        If CheckGuildPermission(Index, 6) = True Then
            'Guild Name
            GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Name = Buffer.ReadString
            
            'Guild Tag
            GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Tag = Buffer.ReadString
            'Guild Color
            HoldInt = Buffer.ReadInteger
            If HoldInt > 0 Then
                GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Color = HoldInt
                HoldInt = 0
            End If
            
            'Guild Recruit rank
            HoldInt = Buffer.ReadInteger
            
            'Guild MOTD
            GuildData(TempPlayer(Index).tmpGuildSlot).Guild_MOTD = Buffer.ReadString
            
            GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Logo = Buffer.ReadLong
            
            
            'Did Recruit Rank change? Make sure they didnt set recruit rank at or above their rank
            If Not GuildData(TempPlayer(Index).tmpGuildSlot).Guild_RecruitRank = HoldInt Then
                If Not HoldInt >= GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).Rank Then
                    GuildData(TempPlayer(Index).tmpGuildSlot).Guild_RecruitRank = HoldInt
                    
                Else
                    PlayerMsg Index, "You may not set the recruit rank higher or at your rank.", BrightRed
                End If
            End If
        Else
            PlayerMsg Index, "You are not allowed to save options.", BrightRed
        End If
        HoldInt = 0
    Case 2
        'users
        If CheckGuildPermission(Index, 5) = True Then
            'Guild Member Rank
            HoldInt = Buffer.ReadInteger
            If HoldInt > 0 Then
                GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(SentIndex).Rank = HoldInt
            Else
                PlayerMsg Index, "Must set rank above 0", BrightRed
            End If
            
            'Guild Member Comment
            GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(SentIndex).Comment = Buffer.ReadString
        Else
            PlayerMsg Index, "You are not allowed to save users.", BrightRed
        End If
        
    Case 3
        'ranks
        If CheckGuildPermission(Index, 4) = True Then
            GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Ranks(SentIndex).Name = Buffer.ReadString
                For i = 1 To MAX_GUILD_RANKS_PERMISSION
                    GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Ranks(SentIndex).RankPermission(i) = Buffer.ReadByte
                Next
        Else
            PlayerMsg Index, "You are not allowed to save ranks.", BrightRed
        End If
    
    End Select
    
    Call SendGuild(True, Index, TempPlayer(Index).tmpGuildSlot)
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    Set Buffer = Nothing
End Sub
Public Sub HandleGuildCommands(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Integer
    Dim SelectedIndex As Long
    Dim SendText As String, SendText2 As String
    Dim SelectedCommand As Integer
    Dim MembersCount As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    
    SelectedCommand = Buffer.ReadInteger
    SendText = Buffer.ReadString
    SendText2 = Buffer.ReadString

    'Command 1/6/7 can be used while not in a guild
    If Player(Index).GuildFileId = 0 And Not (SelectedCommand = 1 Or SelectedCommand = 6 Or SelectedCommand = 7) Then
        PlayerMsg Index, "You must be in a guild to use this commands!", BrightRed
        Exit Sub
    End If
    
    
    
    Select Case SelectedCommand
    Case 1
        'make
        Call MakeGuild(Index, SendText, SendText2)
        PlayerMsg Index, SendText & " - " & SendText2, BrightRed
    Case 2
        'invite
        'Find user index
        SelectedIndex = 0
        
        'Try to find player
        SelectedIndex = FindPlayer(SendText)
        
        If SelectedIndex > 0 Then
            Call Guild_Invite(SelectedIndex, TempPlayer(Index).tmpGuildSlot, Index)
        Else
            PlayerMsg Index, "Could not find user " & SendText & ".", BrightRed
        End If
        
    Case 3
        'leave
        Call GuildLeave(Index)
        
    Case 4
        'admin
        If CheckGuildPermission(Index, 1) = True Then
            Call ToggleGuildAdmin(Index, True)
        Else
            PlayerMsg Index, "You are not allowed to open the admin panel.", BrightRed
        End If
    
    Case 5
        'view
        'This sets the default option
        If SendText = "" Then SendText = "online"
        MembersCount = 0
        
        Select Case SendText
        Case "online"
            PlayerMsg Index, GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).Online = True Then
                        PlayerMsg Index, GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).User_Name, Green
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next
            
            PlayerMsg Index, "Total: " & MembersCount, Green
        
        Case "all"
            PlayerMsg Index, GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    PlayerMsg Index, GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).User_Name, Green
                    MembersCount = MembersCount + 1
                End If
            Next
            
            PlayerMsg Index, "Total: " & MembersCount, Green
        
        Case "offline"
            PlayerMsg Index, GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Name & " Members List (" & UCase(SendText) & ")", Green
            For i = 1 To MAX_GUILD_MEMBERS
                If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).Used = True Then
                    If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).Online = False Then
                        PlayerMsg Index, GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(i).User_Name, Green
                        MembersCount = MembersCount + 1
                    End If
                End If
            Next
            
            PlayerMsg Index, "Total: " & MembersCount, Green
        
        End Select
    Case 6
        'accept
        If TempPlayer(Index).tmpGuildInviteSlot > 0 Then
            If GuildData(TempPlayer(Index).tmpGuildInviteSlot).In_Use = True And GuildData(TempPlayer(Index).tmpGuildInviteSlot).Guild_Fileid = TempPlayer(Index).tmpGuildInviteId Then
                Call Join_Guild(Index, TempPlayer(Index).tmpGuildInviteSlot)
                TempPlayer(Index).tmpGuildInviteSlot = 0
                TempPlayer(Index).tmpGuildInviteTimer = 0
                TempPlayer(Index).tmpGuildInviteId = 0
            Else
                PlayerMsg Index, "No one from this guild is online any more, please ask for a new invite.", BrightRed
            End If
        Else
            PlayerMsg Index, "You must get a guild invite to use this command.", BrightRed
        End If
    Case 7
        'decline
        If TempPlayer(Index).tmpGuildInviteSlot > 0 Then
            TempPlayer(Index).tmpGuildInviteSlot = 0
            TempPlayer(Index).tmpGuildInviteTimer = 0
            TempPlayer(Index).tmpGuildInviteId = 0
            PlayerMsg Index, "You declined the guild invite.", BrightRed
        Else
            PlayerMsg Index, "You must get a guild invite to use this command.", BrightRed
        End If
        
    Case 8
        'founder
        'Make sure the person who used the command is who they say they are
        If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).User_Login = Player(Index).Login Then
            'Make sure they are founder
            If GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).Founder = True Then
                'Find user index
                SelectedIndex = 0
                
                'Try to find player
                SelectedIndex = FindPlayer(SendText)
                
                If SelectedIndex > 0 Then
                    'Make sure the person getting founder is the correct person
                    If GuildData(TempPlayer(SelectedIndex).tmpGuildSlot).Guild_Members(Player(SelectedIndex).GuildMemberId).User_Login = Player(SelectedIndex).Login Then
                        GuildData(TempPlayer(Index).tmpGuildSlot).Guild_Members(Player(Index).GuildMemberId).Founder = False
                        GuildData(TempPlayer(SelectedIndex).tmpGuildSlot).Guild_Members(Player(SelectedIndex).GuildMemberId).Founder = True
                    End If
                Else
                    PlayerMsg Index, "Could not find user " & SendText & ".", BrightRed
                End If
            Else
                 PlayerMsg Index, "You must be marked as a founder to use this command.", BrightRed
            End If
        End If
    Case 9
        'kick
        Call GuildKick(TempPlayer(Index).tmpGuildSlot, Index, SendText)
    
    Case 10
        'disband
        Call DisbandGuild(TempPlayer(Index).tmpGuildSlot, Index)
        
    
    End Select
  
    Set Buffer = Nothing
End Sub

Public Sub Guild_Invite(Index As Long, GuildSlot As Long, Inviter_Index As Long)
    ' check if the person is a valid target
    If Not IsConnected(TARGET_TYPE_PLAYER) Or Not IsPlaying(TARGET_TYPE_PLAYER) Then Exit Sub
    
    If Player(Index).GuildFileId > 0 Then
        PlayerMsg Index, "You must leave your current guild before you can join " & GuildData(GuildSlot).Guild_Name & "!", BrightRed
        PlayerMsg Inviter_Index, "They are unable to join because they are already in a guild!", BrightRed
        Exit Sub
    End If
    
    If TempPlayer(Index).tmpGuildInviteSlot > 0 Then
        PlayerMsg Inviter_Index, "This user has a pending invite try again.", BrightRed
        Exit Sub
    End If
    
    'Permission 2 = Can Recruit
    If CheckGuildPermission(Inviter_Index, 2) = False Then
        PlayerMsg Inviter_Index, "Sorry your rank is not high enough!", BrightRed
        Exit Sub
    End If
    
    TempPlayer(Index).tmpGuildInviteSlot = GuildSlot
    TempPlayer(Index).tmpGuildInviteId = Player(Inviter_Index).GuildFileId
    SendGuildInvite Index, GuildSlot, Inviter_Index
    PlayerMsg Inviter_Index, "Guild invite sent!", Green
End Sub

Public Sub Guild_InviteDecline(ByVal Index As Long, ByVal TARGETPLAYER As Long)
    PlayerMsg Index, GetPlayerName(TARGETPLAYER) & " has declined to join guild", BrightRed
    PlayerMsg TARGETPLAYER, "You declined to join guild.", BrightRed
    ' clear the invitation
    TempPlayer(Index).tmpGuildInviteSlot = 0
    TempPlayer(Index).tmpGuildInviteId = 0
End Sub

Sub SendGuildInvite(ByVal Index As Long, ByVal GuildSlot As Long, ByVal Inviter_Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SGuildInvite
    Buffer.WriteString Trim$(Player(Inviter_Index).Name)
    Buffer.WriteString Trim$(GuildData(GuildSlot).Guild_Name)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
