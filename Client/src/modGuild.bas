Attribute VB_Name = "modGuild"
Public Const MAX_GUILD_MEMBERS As Long = 50
Public Const MAX_GUILD_RANKS As Long = 6
Public Const MAX_GUILD_RANKS_PERMISSION As Long = 6


Public GuildData As GuildRec

Public Type GuildRanksRec
    'General variables
    Used As Boolean
    name As String
    
    'Rank Variables
    RankPermission(1 To MAX_GUILD_RANKS_PERMISSION) As Byte
    RankPermissionName(1 To MAX_GUILD_RANKS_PERMISSION) As String
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
    Comment As String * 300
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
    Guild_Color As Long
    Guild_Logo As Long
End Type
Public Sub HandleAdminGuild(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    
    If Buffer.ReadByte = 1 Then
        If Player(MyIndex).Donator = YES Then
            frmGuildAdmin.scrlGuildLogo.Enabled = True
        Else
            frmGuildAdmin.scrlGuildLogo.Enabled = False
        End If
        frmGuildAdmin.visible = True
    Else
        frmGuildAdmin.visible = False
    End If

    
    Set Buffer = Nothing
End Sub
Public Sub HandleSendGuild(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Integer
Dim B As Integer

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes data()
    
    GuildData.Guild_Name = Buffer.ReadString
    GuildData.Guild_Tag = Buffer.ReadString
    GuildData.Guild_Color = Buffer.ReadInteger
    GuildData.Guild_MOTD = Buffer.ReadString
    GuildData.Guild_RecruitRank = Buffer.ReadInteger
    GuildData.Guild_Logo = Buffer.ReadLong
    
    'Get Members
    For i = 1 To MAX_GUILD_MEMBERS
        GuildData.Guild_Members(i).User_Name = Buffer.ReadString
        GuildData.Guild_Members(i).Rank = Buffer.ReadInteger
        GuildData.Guild_Members(i).Comment = Buffer.ReadString
        GuildData.Guild_Members(i).Online = Buffer.ReadByte
    Next i
    
    'Get Ranks
    For i = 1 To MAX_GUILD_RANKS
        GuildData.Guild_Ranks(i).name = Buffer.ReadString
        For B = 1 To MAX_GUILD_RANKS_PERMISSION
            GuildData.Guild_Ranks(i).RankPermission(B) = Buffer.ReadByte
            GuildData.Guild_Ranks(i).RankPermissionName(B) = Buffer.ReadString
        Next B
    Next i
    
    'Update Guildadmin data
    Call frmGuildAdmin.Load_Guild_Admin
    
    ' Reset Players Guild Tag/Name
    TempPlayer(MyIndex).GuildName = GuildData.Guild_Name
    TempPlayer(MyIndex).GuildTag = GuildData.Guild_Tag
    TempPlayer(MyIndex).GuildColor = GuildData.Guild_Color
    
    
    Set Buffer = Nothing
End Sub
Public Sub GuildMsg(ByVal Text As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSayGuild
    Buffer.WriteString Text
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub GuildCommand(ByVal Command As Integer, ByVal SendText As String, Optional ByVal SendText2 As String)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CGuildCommand
    Buffer.WriteInteger Command
    Buffer.WriteString SendText
    Buffer.WriteString SendText2
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub GuildSave(ByVal SaveType As Integer, ByVal Index As Integer)
Dim Buffer As clsBuffer
Dim i As Integer
'SaveType
'1=options
'2=users
'3=ranks

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSaveGuild
    
    Buffer.WriteInteger SaveType
    Buffer.WriteInteger Index
    
    Select Case SaveType
    Case 1
        'options
        Buffer.WriteString GuildData.Guild_Name
        Buffer.WriteString GuildData.Guild_Tag
        Buffer.WriteInteger GuildData.Guild_Color
        Buffer.WriteInteger GuildData.Guild_RecruitRank
        Buffer.WriteString GuildData.Guild_MOTD
        Buffer.WriteLong GuildData.Guild_Logo
    Case 2
        'users
        Buffer.WriteInteger GuildData.Guild_Members(Index).Rank
        Buffer.WriteString GuildData.Guild_Members(Index).Comment
    Case 3
        'ranks
        Buffer.WriteString GuildData.Guild_Ranks(Index).name
        For i = 1 To MAX_GUILD_RANKS_PERMISSION
            Buffer.WriteByte GuildData.Guild_Ranks(Index).RankPermission(i)
        Next i
    End Select

    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub HandleGuildInvite(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddR As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim theName As String
Dim GuildName As String
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes data()

    theName = Buffer.ReadString
    GuildName = Buffer.ReadString
    
    Dialogue "Guild Invitation", theName & " has invited you to " & GuildName & ". Would you like to join?", DIALOGUE_TYPE_GUILD, True
End Sub
