VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mHeader() As Byte
Private mVersion As Long

Private mIndex As Long
Private mName As String
Private mMusic As String

Private mRevision As Long
Private mMoral As Byte

Private mUp As Long
Private mDown As Long
Private mLeft As Long
Private mRight As Long

Private mBootMap As Long
Private mBootX As Byte
Private mBootY As Byte

Private mMaxX As Byte
Private mMaxY As Byte

Private mTile() As TileRec
Private mNPC() As Long
Private mBossNPC As Long

Private mFog As Byte
Private mFogSpeed As Byte
Private mFogOpacity As Byte

Private mRed As Byte
Private mGreen As Byte
Private mBlue As Byte
Private mAlpha As Byte

Private mPanorama As Byte
Private mDayNight As Byte
Private mNPCSpawnType() As Long

Private mMapItem() As MapItemRec
Private mMapNPC() As MapNpcRec

Public Property Get Moral() As Long
    Moral = mMoral
End Property

Public Sub Init(ByVal Index As Long)
    mIndex = Index
    ReDim mHeader(0 To 3)
    Call Move32(&HFF001020, mHeader(0))
    mVersion = 1
    
    ReDim mTile(0 To MAX_MAPX, 0 To MAX_MAPY)
    ReDim mNPC(1 To MAX_MAP_NPCS)
    ReDim mNPCSpawnType(1 To MAX_MAP_NPCS)
    ReDim mMapItem(1 To MAX_MAP_ITEMS)
    ReDim mMapNPC(1 To MAX_MAP_NPCS)
End Sub

Public Sub ClearItem(ByVal Index As Long)
    Call memzero(mMapItem(Index), LenB(mMapItem(Index)))
    mMapItem(Index).PlayerName = vbNullString
    Call SendSpawnItemToMap(mIndex, Index) ' We should create a despawn packet
End Sub

Public Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        Call ClearItem(i)
    Next
End Sub

Public Function FindOpenItemSlot() As Long
Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        If mMapItem(i).Num = 0 Then
            FindOpenItemSlot = i
            Exit Function
        End If
    Next
End Function

Public Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal PlayerName As String = vbNullString, Optional ByVal canDespawn As Boolean = True, Optional ByVal isSB As Boolean = False)
Dim i As Long

    ' Find open map item slot
    i = FindOpenItemSlot
    If i <> 0 Then
        mMapItem(i).PlayerName = PlayerName
        mMapItem(i).playerTimer = timeGetTime + ITEM_SPAWN_TIME
        mMapItem(i).canDespawn = canDespawn
        mMapItem(i).despawnTimer = timeGetTime + ITEM_DESPAWN_TIME
        mMapItem(i).Num = ItemNum
        mMapItem(i).Value = ItemVal
        mMapItem(i).X = X
        mMapItem(i).Y = Y
        mMapItem(i).Bound = isSB
        Call SendSpawnItemToMap(mIndex, i)
    End If
End Sub

Public Sub SpawnItems()
Dim X As Long, Y As Long

    ' Spawn what we have
    For X = 0 To mMaxX
        For Y = 0 To mMaxY
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If mTile(X, Y).Type = TILE_TYPE_ITEM Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(mTile(X, Y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(mTile(X, Y).Data1).Stackable = YES And mTile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(mTile(X, Y).Data1, 1, X, Y)
                Else
                    Call SpawnItem(mTile(X, Y).Data1, mTile(X, Y).Data2, X, Y)
                End If
            End If

        Next
    Next
End Sub

Public Sub ClearNPC(ByVal Index As Long)
    Call memzero(mMapNPC(Index), LenB(mMapNPC(Index)))
End Sub

Public Sub ClearMapNPCs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearNPC(i)
    Next
End Sub

Public Sub SpawnNPC(ByVal MapNPCNum As Long, Optional ForcedSpawn As Boolean = False)
Dim Buffer As clsBuffer
Dim NPCNum As Long
Dim Spawned As Boolean
Dim X As Long, Y As Long
Dim i As Long

    NPCNum = mNPC(MapNPCNum)
    If ForcedSpawn = False And mNPCSpawnType(MapNPCNum) = 1 Then Exit Sub
    
    With mMapNPC(MapNPCNum)
        .Num = NPCNum
        .target = 0
        .targetType = 0 ' clear
        .Vital(Vitals.HP) = GetNpcMaxVital(NPCNum, Vitals.HP)
        .Vital(Vitals.MP) = GetNpcMaxVital(NPCNum, Vitals.MP)
        .Dir = RAND(0, 3)
        .spellBuffer.spell = 0
        .spellBuffer.Timer = 0
        .spellBuffer.target = 0
        .spellBuffer.tType = 0
        
        'Check if theres a spawn tile for the specific npc
        For X = 0 To mMaxX
            For Y = 0 To mMaxY
                If mTile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                    If mTile(X, Y).Data1 = MapNPCNum Then
                        .X = X
                        .Y = Y
                        .Dir = mTile(X, Y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next
        Next
        
        If Not Spawned Then
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                X = Random(0, mMaxX)
                Y = Random(0, mMaxY)
                
                If X > mMaxX Then X = mMaxX
                If Y > mMaxY Then Y = mMaxY
                
                ' Check if the tile is walkable
                If NPCTileIsOpen(X, Y) Then
                    .X = X
                    .Y = Y
                    Spawned = True
                    Exit For
                End If
            Next
        End If
        
        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For X = 0 To mMaxX
                For Y = 0 To mMaxY
                    If NPCTileIsOpen(X, Y) Then
                        .X = X
                        .Y = Y
                        Spawned = True
                    End If
                Next
            Next
        End If
        
        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SSpawnNpc)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(.Num)
            Call Buffer.WriteLong(.X)
            Call Buffer.WriteLong(.Y)
            Call Buffer.WriteLong(.Dir)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        End If
        
        Call SendMapNpcVitals(mIndex, MapNPCNum)
    End With
End Sub

Public Sub SpawnNPCs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        If mNPC(i) > 0 And mNPC(i) <= MAX_NPCS Then
            If NPC(mNPC(i)).SpawnAtDay = 0 Then
                If DayTime Then Call SpawnNPC(i)
            Else
                If DayTime = False Then Call SpawnNPC(i)
            End If
        End If
    Next
End Sub

Public Sub DespawnNPC(ByVal NPCNum As Long)
Dim i As Long, Buffer As clsBuffer

    ' Reset the targets..
    mMapNPC(NPCNum).target = 0
    mMapNPC(NPCNum).targetType = TARGET_TYPE_NONE
    
    ' Set the NPC data to blank so it despawns.
    mMapNPC(NPCNum).Num = 0
    mMapNPC(NPCNum).SpawnWait = 0
    mMapNPC(NPCNum).Vital(Vitals.HP) = 0
    
    ' clear DoTs and HoTs
    For i = 1 To MAX_DOTS
        mMapNPC(NPCNum).DoT(i).spell = 0
        mMapNPC(NPCNum).DoT(i).Timer = 0
        mMapNPC(NPCNum).DoT(i).Caster = 0
        mMapNPC(NPCNum).DoT(i).StartTime = 0
        mMapNPC(NPCNum).DoT(i).Used = False
        mMapNPC(NPCNum).HoT(i).spell = 0
        mMapNPC(NPCNum).HoT(i).Timer = 0
        mMapNPC(NPCNum).HoT(i).Caster = 0
        mMapNPC(NPCNum).HoT(i).StartTime = 0
        mMapNPC(NPCNum).HoT(i).Used = False
    Next
    
    ' send death to the map
    Set Buffer = New clsBuffer
    Call Buffer.WriteLong(SNpcDead)
    Call Buffer.WriteLong(NPCNum)
    Call SendDataToMap(mIndex, Buffer.Serialize)
    
    'Loop through entire map and purge NPC from targets
    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            If Player(i).Map = mIndex Then
                If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                    If TempPlayer(i).target = NPCNum Then
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        Call SendTarget(i)
                    End If
                End If
            End If
        End If
    Next
End Sub

Private Function NPCTileIsOpen(ByVal X As Long, ByVal Y As Long) As Boolean
Dim i As Long

    NPCTileIsOpen = True
    
    If PlayersOnMap(mIndex) Then
        For i = 1 To Player_HighIndex
            If GetPlayerMap(i) = mIndex Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        NPCTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If
        Next
    End If
    
    For i = 1 To MAX_MAP_NPCS
        If mMapNPC(i).Num > 0 Then
            If mMapNPC(i).X = X Then
                If mMapNPC(i).Y = Y Then
                    NPCTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next
    
    If mTile(X, Y).Type <> TILE_TYPE_WALKABLE Then
        If mTile(X, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If mTile(X, Y).Type <> TILE_TYPE_ITEM Then
                NPCTileIsOpen = False
            End If
        End If
    End If
End Function

Private Function CanNPCMove(ByVal MapNPCNum As Long, ByVal Dir As Long) As Boolean
Dim i As Long, n As Long
Dim X As Long, Y As Long

    X = mMapNPC(MapNPCNum).X
    Y = mMapNPC(MapNPCNum).Y
    CanNPCMove = True
    
    Select Case Dir
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If Y > 0 And X > 0 Then
                n = mTile(X - 1, Y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X - 1) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X - 1) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X - 1) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y - 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_UP + 1) And isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If Y > 0 And X < mMaxX Then
                n = mTile(X + 1, Y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X + 1) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X + 1) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X + 1) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y - 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_UP + 1) And isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_DOWN_LEFT
            ' Check to make sure not outside of boundries
            If Y < mMaxY And X > 0 Then
                n = mTile(X - 1, Y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X - 1) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X - 1) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X - 1) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y + 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) And isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_DOWN_RIGHT
            ' Check to make sure not outside of boundries
            If Y < mMaxY And X < mMaxX Then
                n = mTile(X + 1, Y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X + 1) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X + 1) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X + 1) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y + 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) And isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = mTile(X, Y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y - 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_UP + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < mMaxY Then
                n = mTile(X, Y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y + 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = mTile(X - 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X - 1) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X - 1) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X - 1) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < mMaxX Then
                n = mTile(X + 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = mIndex) And (GetPlayerX(i) = mMapNPC(MapNPCNum).X + 1) And (GetPlayerY(i) = mMapNPC(MapNPCNum).Y) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                        If (GetPlayerMap(i) = mIndex) And (Player(i).Pet.X = mMapNPC(MapNPCNum).X + 1) And (Player(i).Pet.Y = mMapNPC(MapNPCNum).Y) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNPCNum) And (mMapNPC(i).Num > 0) And (mMapNPC(i).X = mMapNPC(MapNPCNum).X + 1) And (mMapNPC(i).Y = mMapNPC(MapNPCNum).Y) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
    End Select
End Function

Public Sub NPCMove(ByVal MapNPCNum As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    mMapNPC(MapNPCNum).Dir = Dir
    
    Select Case Dir
        Case DIR_UP_LEFT
            mMapNPC(MapNPCNum).Y = mMapNPC(MapNPCNum).Y - 1
            mMapNPC(MapNPCNum).X = mMapNPC(MapNPCNum).X - 1
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_UP_RIGHT
            mMapNPC(MapNPCNum).Y = mMapNPC(MapNPCNum).Y - 1
            mMapNPC(MapNPCNum).X = mMapNPC(MapNPCNum).X + 1
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_DOWN_LEFT
            mMapNPC(MapNPCNum).Y = mMapNPC(MapNPCNum).Y + 1
            mMapNPC(MapNPCNum).X = mMapNPC(MapNPCNum).X - 1
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_DOWN_RIGHT
            mMapNPC(MapNPCNum).Y = mMapNPC(MapNPCNum).Y + 1
            mMapNPC(MapNPCNum).X = mMapNPC(MapNPCNum).X + 1
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_UP
            mMapNPC(MapNPCNum).Y = mMapNPC(MapNPCNum).Y - 1
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_DOWN
            mMapNPC(MapNPCNum).Y = mMapNPC(MapNPCNum).Y + 1
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_LEFT
            mMapNPC(MapNPCNum).X = mMapNPC(MapNPCNum).X - 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_RIGHT
            mMapNPC(MapNPCNum).X = mMapNPC(MapNPCNum).X + 1
            Call Buffer.WriteLong(SNpcMove)
            Call Buffer.WriteLong(MapNPCNum)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).X)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Y)
            Call Buffer.WriteLong(mMapNPC(MapNPCNum).Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
    End Select
End Sub

Public Sub NPCDir(ByVal MapNPCNum As Long, ByVal Dir As Long)
Dim Buffer As clsBuffer

    mMapNPC(MapNPCNum).Dir = Dir
    Set Buffer = New clsBuffer
    Call Buffer.WriteLong(SNpcDir)
    Call Buffer.WriteLong(MapNPCNum)
    Call Buffer.WriteLong(Dir)
    Call SendDataToMap(mIndex, Buffer.Serialize)
End Sub

Public Sub CacheResources()
Dim X As Long, Y As Long, Resource_Count As Long

    For X = 0 To mMaxX
        For Y = 0 To mMaxY
            If mTile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(mIndex).ResourceData(0 To Resource_Count)
                ResourceCache(mIndex).ResourceData(Resource_Count).X = X
                ResourceCache(mIndex).ResourceData(Resource_Count).Y = Y
                ResourceCache(mIndex).ResourceData(Resource_Count).cur_health = Resource(mTile(X, Y).Data1).Health
            End If
        Next
    Next
    
    ResourceCache(mIndex).Resource_Count = Resource_Count
End Sub

Public Function Serialize() As Byte()
Dim Buffer As clsBuffer
Dim X As Long, Y As Long
Dim i As Long

    Set Buffer = New clsBuffer
    'Call Buffer.WriteRaw(mHeader)
    'Call Buffer.WriteLong(mVersion)
    Call Buffer.WriteString(mName)
    Call Buffer.WriteString(mMusic)
    Call Buffer.WriteLong(mRevision)
    Call Buffer.WriteByte(mMoral)
    Call Buffer.WriteLong(mUp)
    Call Buffer.WriteLong(mDown)
    Call Buffer.WriteLong(mLeft)
    Call Buffer.WriteLong(mRight)
    Call Buffer.WriteLong(mBootMap)
    Call Buffer.WriteByte(mBootX)
    Call Buffer.WriteByte(mBootY)
    Call Buffer.WriteByte(mMaxX)
    Call Buffer.WriteByte(mMaxY)
    
    For X = 0 To mMaxX
        For Y = 0 To mMaxY
            For i = 1 To MapLayer.Layer_Count
                Call Buffer.WriteLong(mTile(X, Y).Layer(i).X)
                Call Buffer.WriteLong(mTile(X, Y).Layer(i).Y)
                Call Buffer.WriteLong(mTile(X, Y).Layer(i).Tileset)
                Call Buffer.WriteByte(mTile(X, Y).Layer(i).Autotile)
            Next
            
            Call Buffer.WriteByte(mTile(X, Y).Type)
            Call Buffer.WriteLong(mTile(X, Y).Data1)
            Call Buffer.WriteLong(mTile(X, Y).Data2)
            Call Buffer.WriteLong(mTile(X, Y).Data3)
            Call Buffer.WriteLong(mTile(X, Y).Data4)
            Call Buffer.WriteByte(mTile(X, Y).DirBlock)
        Next
    Next
    
    For X = 1 To MAX_MAP_NPCS
        Call Buffer.WriteLong(mNPC(X))
    Next
    
    Call Buffer.WriteLong(mBossNPC)
    
    Call Buffer.WriteByte(mFog)
    Call Buffer.WriteByte(mFogSpeed)
    Call Buffer.WriteByte(mFogOpacity)
    
    Call Buffer.WriteByte(mRed)
    Call Buffer.WriteByte(mGreen)
    Call Buffer.WriteByte(mBlue)
    Call Buffer.WriteByte(mAlpha)
    
    Call Buffer.WriteByte(mPanorama)
    Call Buffer.WriteByte(mDayNight)
    
    For X = 1 To MAX_MAP_NPCS
        Call Buffer.WriteLong(mNPCSpawnType(X))
    Next
End Function

Public Sub Deserialize(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim Header() As Byte
Dim Version As Long

    Set Buffer = New clsBuffer
    Call Buffer.Deserialize(Data)
    Header = Buffer.ReadRaw(4)
    
    If AryCmp(Header, mHeader, 4, 4) Then
        Version = Buffer.ReadLong
        Select Case Version
            Case 1
        End Select
    Else
        Call DeserializeOld(Buffer)
    End If
    
    Call CacheResources
End Sub

Public Sub DeserializeOld(ByVal Buffer As clsBuffer)
Dim Ary() As Byte
Dim X As Long, Y As Long
Dim i As Long

    Set Buffer = New clsBuffer
    mName = StrConv(Buffer.ReadRaw(20), vbUnicode)
    mMusic = StrConv(Buffer.ReadRaw(20), vbUnicode)
    mRevision = Buffer.ReadLong
    mMoral = Buffer.ReadByte
    mUp = Buffer.ReadLong
    mDown = Buffer.ReadLong
    mLeft = Buffer.ReadLong
    mRight = Buffer.ReadLong
    mBootMap = Buffer.ReadLong
    mBootX = Buffer.ReadByte
    mBootY = Buffer.ReadByte
    mMaxX = Buffer.ReadByte
    mMaxY = Buffer.ReadByte
    
    For X = 0 To mMaxX
        For Y = 0 To mMaxY
            For i = 1 To MapLayer.Layer_Count
                mTile(X, Y).Layer(i).X = Buffer.ReadLong
                mTile(X, Y).Layer(i).Y = Buffer.ReadLong
                mTile(X, Y).Layer(i).Tileset = Buffer.ReadLong
                mTile(X, Y).Layer(i).Autotile = Buffer.ReadByte
            Next
            
            mTile(X, Y).Type = Buffer.ReadByte
            mTile(X, Y).Data1 = Buffer.ReadLong
            mTile(X, Y).Data2 = Buffer.ReadLong
            mTile(X, Y).Data3 = Buffer.ReadLong
            mTile(X, Y).Data4 = Buffer.ReadLong
            mTile(X, Y).DirBlock = Buffer.ReadByte
        Next
    Next
    
    For X = 1 To MAX_MAP_NPCS
        mNPC(X) = Buffer.ReadLong
    Next
    
    mBossNPC = Buffer.ReadLong
    
    mFog = Buffer.ReadByte
    mFogSpeed = Buffer.ReadByte
    mFogOpacity = Buffer.ReadByte
    
    mRed = Buffer.ReadByte
    mGreen = Buffer.ReadByte
    mBlue = Buffer.ReadByte
    mAlpha = Buffer.ReadByte
    
    mPanorama = Buffer.ReadByte
    mDayNight = Buffer.ReadByte
    
    For X = 1 To MAX_MAP_NPCS
        mNPCSpawnType(X) = Buffer.ReadLong
    Next
End Sub

Function CanPetMove(Index As Long, ByVal Dir As Byte) As Boolean
Dim i As Long, n As Long
Dim X As Long, Y As Long

    X = Player(Index).Pet.X
    Y = Player(Index).Pet.Y
    CanPetMove = True
    
    If TempPlayer(Index).PetspellBuffer.spell > 0 Then
        CanPetMove = False
        Exit Function
    End If
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = mTile(X, Y - 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If GetPlayerMap(i) = mIndex Then
                        If (GetPlayerX(i) = Player(Index).Pet.X + 1) And (GetPlayerY(i) = Player(Index).Pet.Y - 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (Player(i).Pet.X = Player(Index).Pet.X) And (Player(i).Pet.Y = Player(Index).Pet.Y - 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (mMapNPC(i).Num <> 0) And (mMapNPC(i).X = Player(Index).Pet.X) And (mMapNPC(i).Y = Player(Index).Pet.Y - 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_UP + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < mMaxY Then
                n = mTile(X, Y + 1).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If GetPlayerMap(i) = mIndex Then
                        If (GetPlayerX(i) = Player(Index).Pet.X) And (GetPlayerY(i) = Player(Index).Pet.Y + 1) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (Player(i).Pet.X = Player(Index).Pet.X) And (Player(i).Pet.Y = Player(Index).Pet.Y + 1) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (mMapNPC(i).Num <> 0) And (mMapNPC(i).X = Player(Index).Pet.X) And (mMapNPC(i).Y = Player(Index).Pet.Y + 1) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_DOWN + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = mTile(X - 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If GetPlayerMap(i) = mIndex Then
                        If (GetPlayerX(i) = Player(Index).Pet.X - 1) And (GetPlayerY(i) = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (Player(i).Pet.X = Player(Index).Pet.X - 1) And (Player(i).Pet.Y = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (mMapNPC(i).Num <> 0) And (mMapNPC(i).X = Player(Index).Pet.X - 1) And (mMapNPC(i).Y = Player(Index).Pet.Y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_LEFT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
        
        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < mMaxX Then
                n = mTile(X + 1, Y).Type
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanPetMove = False
                    Exit Function
                End If
                
                For i = 1 To Player_HighIndex
                    If GetPlayerMap(i) = mIndex Then
                        If (GetPlayerX(i) = Player(Index).Pet.X + 1) And (GetPlayerY(i) = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        ElseIf Player(i).Pet.Alive And (Player(i).Pet.X = Player(Index).Pet.X + 1) And (Player(i).Pet.Y = Player(Index).Pet.Y) Then
                            CanPetMove = False
                            Exit Function
                        End If
                    End If
                Next
                
                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (mMapNPC(i).Num <> 0) And (mMapNPC(i).X = Player(Index).Pet.X + 1) And (mMapNPC(i).Y = Player(Index).Pet.Y) Then
                        CanPetMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(mTile(Player(Index).Pet.X, Player(Index).Pet.Y).DirBlock, DIR_RIGHT + 1) Then
                    CanPetMove = False
                    Exit Function
                End If
            Else
                CanPetMove = False
            End If
    End Select
End Function

Public Sub PetMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
Dim Buffer As clsBuffer

    Player(Index).Pet.Dir = Dir
    
    Select Case Dir
        Case DIR_UP
            Player(Index).Pet.Y = Player(Index).Pet.Y - 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_DOWN
            Player(Index).Pet.Y = Player(Index).Pet.Y + 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_LEFT
            Player(Index).Pet.X = Player(Index).Pet.X - 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
        
        Case DIR_RIGHT
            Player(Index).Pet.X = Player(Index).Pet.X + 1
            Set Buffer = New clsBuffer
            Call Buffer.WriteLong(SPetMove)
            Call Buffer.WriteLong(Index)
            Call Buffer.WriteLong(Player(Index).Pet.X)
            Call Buffer.WriteLong(Player(Index).Pet.Y)
            Call Buffer.WriteLong(Player(Index).Pet.Dir)
            Call Buffer.WriteLong(Movement)
            Call SendDataToMap(mIndex, Buffer.Serialize)
    End Select
End Sub

Public Sub TryPetAttackNpc(ByVal Index As Long, ByVal MapNPCNum As Long)
Dim BlockAmount As Long
Dim NPCNum As Long
Dim Damage As Long
Dim RndChance As Long
Dim RndChance2 As Long

    ' Can we attack the npc?
    If CanPetAttackNPC(Index, MapNPCNum) Then
        NPCNum = mMapNPC(MapNPCNum).Num
    
        ' check if NPC can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            Call SendActionMsg(mIndex, "Dodge!", Pink, 1, mMapNPC(MapNPCNum).X * 32, mMapNPC(MapNPCNum).Y * 32)
            Exit Sub
        End If
        If CanNpcDodge(NPCNum) Then
            Call SendActionMsg(mIndex, "Dodge!", Pink, 1, mMapNPC(MapNPCNum).X * 32, mMapNPC(MapNPCNum).Y * 32)
            Exit Sub
        End If
        If CanNpcParry(NPCNum) Then
            Call SendActionMsg(mIndex, "Parry!", Pink, 1, mMapNPC(MapNPCNum).X * 32, mMapNPC(MapNPCNum).Y * 32)
            Exit Sub
        End If
        
        ' Get the damage we can do
        Damage = GetPetDamage(Index)
        
        ' if the npc blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanNpcBlock(MapNPCNum)
        Damage = Damage - BlockAmount - RndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, (NPC(NPCNum).Stat(Stats.Agility) * 2))
        ' randomise from 1 to max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if it's a crit!
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            Damage = Damage * 1.5
            Call SendActionMsg(mIndex, "Critical!", BrightCyan, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)
        End If
        
        If CanPetCrit(Index) Then
            Damage = Damage * 1.5
            Call SendActionMsg(mIndex, "Critical!", BrightCyan, 1, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32)
        End If
        
        If Damage > 0 Then
            Call PetAttackNPC(Index, MapNPCNum, Damage)
        Else
            Call PlayerMsg(Index, "Your pet's attack does nothing.", BrightRed)
        End If
    End If
End Sub

Public Function CanPetAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
Dim MapNum As Long
Dim NPCNum As Long
Dim NPCX As Long
Dim NPCY As Long
Dim Attackspeed As Long

    MapNum = GetPlayerMap(Attacker)
    NPCNum = mMapNPC(MapNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If mMapNPC(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If
    
    ' Make sure they are on the same map
    If TempPlayer(Attacker).PetspellBuffer.spell > 0 And IsSpell = False Then Exit Function
    
    ' exit out early
    If IsSpell Then
        If NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER Then
            CanPetAttackNPC = True
            Exit Function
        End If
    End If
    
    Attackspeed = 1000 'Pet cannot weild a weapon
    
    If timeGetTime > TempPlayer(Attacker).PetAttackTimer + Attackspeed Then
        ' Check if at same coordinates
        Select Case Player(Attacker).Pet.Dir
            Case DIR_UP
                NPCX = mMapNPC(MapNPCNum).X
                NPCY = mMapNPC(MapNPCNum).Y + 1
            Case DIR_DOWN
                NPCX = mMapNPC(MapNPCNum).X
                NPCY = mMapNPC(MapNPCNum).Y - 1
            Case DIR_LEFT
                NPCX = mMapNPC(MapNPCNum).X + 1
                NPCY = mMapNPC(MapNPCNum).Y
            Case DIR_RIGHT
                NPCX = mMapNPC(MapNPCNum).X - 1
                NPCY = mMapNPC(MapNPCNum).Y
        End Select
        
        If NPCX = Player(Attacker).Pet.X Then
            If NPCY = Player(Attacker).Pet.Y Then
                CanPetAttackNPC = NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_FRIENDLY And NPC(NPCNum).Behaviour <> NPC_BEHAVIOUR_SHOPKEEPER
            End If
        End If
    End If
End Function

Public Sub PetAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
Dim Exp As Long
Dim n As Integer
Dim i As Long
Dim NPCNum As Long
Dim Buffer As clsBuffer

    NPCNum = mMapNPC(MapNPCNum).Num
    
    If SpellNum = 0 Then
        ' Send this packet so they can see the pet attacking
        Set Buffer = New clsBuffer
        Call Buffer.WriteLong(SAttack)
        Call Buffer.WriteLong(Attacker)
        Call Buffer.WriteLong(1)
        Call SendDataToMap(mIndex, Buffer.Serialize)
        Set Buffer = Nothing
    End If
    
    ' set the regen timer
    TempPlayer(Attacker).PetstopRegen = True
    TempPlayer(Attacker).PetstopRegenTimer = timeGetTime

    If Damage >= mMapNPC(MapNPCNum).Vital(Vitals.HP) Then
        If MapNPCNum = mBossNPC Then
            Call SendBossMsg(NPC(mMapNPC(MapNPCNum).Num).Name & " has been slain by " & GetPlayerName(Attacker) & " in " & mName & ".", Magenta)
            Call GlobalMsg(NPC(mMapNPC(MapNPCNum).Num).Name & " has been slain by " & GetPlayerName(Attacker) & " in " & mName & ".", Magenta)
        End If
        
        Call SendActionMsg(GetPlayerMap(Attacker), "-" & mMapNPC(MapNPCNum).Vital(Vitals.HP), BrightRed, 1, mMapNPC(MapNPCNum).X * 32, mMapNPC(MapNPCNum).Y * 32)
        Call SendBlood(GetPlayerMap(Attacker), mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y)
        
        ' send the sound
        If SpellNum > 0 Then
            Call SendMapSound(Attacker, mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y, SoundEntity.seSpell, SpellNum)
        End If
        
        ' Calculate exp to give attacker
        Exp = NPC(NPCNum).Exp
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then
            Exp = 1
        End If
        
        ' in party?
        If TempPlayer(Attacker).inParty > 0 Then
            ' pass through party sharing function
            Call Party_ShareExp(TempPlayer(Attacker).inParty, Exp, Attacker)
        Else
            ' no party - keep exp for self
            Call GivePlayerEXP(Attacker, Exp)
        End If
        
        For n = 1 To MAX_NPC_DROPS
            If NPC(NPCNum).DropItem(n) = 0 Then Exit For
            If Rnd <= NPC(NPCNum).DropChance(n) Then
                Call SpawnItem(NPC(NPCNum).DropItem(n), NPC(NPCNum).DropItemValue(n), mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y, GetPlayerName(Attacker))
            End If
        Next
        
        If NPC(NPCNum).Event > 0 Then
            Call InitEvent(Attacker, NPC(NPCNum).Event)
        End If
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        mMapNPC(MapNPCNum).Num = 0
        mMapNPC(MapNPCNum).SpawnWait = timeGetTime
        mMapNPC(MapNPCNum).Vital(Vitals.HP) = 0
        
        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            mMapNPC(MapNPCNum).DoT(i).spell = 0
            mMapNPC(MapNPCNum).DoT(i).Timer = 0
            mMapNPC(MapNPCNum).DoT(i).Caster = 0
            mMapNPC(MapNPCNum).DoT(i).StartTime = 0
            mMapNPC(MapNPCNum).DoT(i).Used = False
            mMapNPC(MapNPCNum).HoT(i).spell = 0
            mMapNPC(MapNPCNum).HoT(i).Timer = 0
            mMapNPC(MapNPCNum).HoT(i).Caster = 0
            mMapNPC(MapNPCNum).HoT(i).StartTime = 0
            mMapNPC(MapNPCNum).HoT(i).Used = False
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Call Buffer.WriteLong(SNpcDead)
        Call Buffer.WriteLong(MapNPCNum)
        Call SendDataToMap(mIndex, Buffer.Serialize)
        
        'Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsConnected(i) Then
                If Player(i).Map = mIndex Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNPCNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            Call SendTarget(i)
                        End If
                    End If
                    
                    If TempPlayer(i).PetTargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).PetTarget = MapNPCNum Then
                            TempPlayer(i).PetTarget = 0
                            TempPlayer(i).PetTargetType = TARGET_TYPE_NONE
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        mMapNPC(MapNPCNum).Vital(Vitals.HP) = mMapNPC(MapNPCNum).Vital(Vitals.HP) - Damage
        
        ' Check for a weapon and say damage
        Call SendActionMsg(mIndex, "-" & Damage, BrightRed, 1, mMapNPC(MapNPCNum).X * 32, mMapNPC(MapNPCNum).Y * 32)
        Call SendBlood(GetPlayerMap(Attacker), mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y)
        
        ' send the sound
        If SpellNum > 0 Then
            Call SendMapSound(Attacker, mMapNPC(MapNPCNum).X, mMapNPC(MapNPCNum).Y, SoundEntity.seSpell, SpellNum)
        End If
        
        ' Set the NPC target to the player
        mMapNPC(MapNPCNum).targetType = TARGET_TYPE_PET ' player's pet
        mMapNPC(MapNPCNum).target = Attacker
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(mMapNPC(MapNPCNum).Num).Behaviour = NPC_BEHAVIOUR_GUARD Then
            For i = 1 To MAX_MAP_NPCS
                If mMapNPC(i).Num = mMapNPC(MapNPCNum).Num Then
                    mMapNPC(i).target = Attacker
                    mMapNPC(i).targetType = TARGET_TYPE_PET ' pet
                End If
            Next
        End If
        
        ' set the regen timer
        mMapNPC(MapNPCNum).stopRegen = True
        mMapNPC(MapNPCNum).stopRegenTimer = timeGetTime
        
        ' if stunning spell, stun the npc
        If SpellNum > 0 Then
            If spell(SpellNum).StunDuration > 0 Then Call StunNPC(MapNPCNum, mIndex, SpellNum)
            ' DoT
            If spell(SpellNum).Duration > 0 Then
                Call AddDoT_Npc(mIndex, MapNPCNum, SpellNum, Attacker)
            End If
        End If
        
        Call SendMapNpcVitals(mIndex, MapNPCNum)
    End If
    
    If SpellNum = 0 Then
        ' Reset attack timer
        TempPlayer(Attacker).PetAttackTimer = timeGetTime
    End If
End Sub

Public Sub TryNPCAttackPet(ByVal MapNPCNum As Long, ByVal Index As Long)
Dim NPCNum As Long, BlockAmount As Long, Damage As Long
Dim RndChance As Long
Dim RndChance2 As Long

    ' Can the npc attack the player?
    If CanNpcAttackPet(MapNPCNum, Index) Then
        NPCNum = mMapNPC(MapNPCNum).Num
    
        ' check if PLAYER can avoid the attack
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            Call SendActionMsg(mIndex, "Dodge!", Pink, 1, Player(Index).Pet.X * 32, Player(Index).Pet.Y * 32)
            Exit Sub
        End If
        If CanPlayerPetDodge(Index) Then
            Call SendActionMsg(mIndex, "Dodge!", Pink, 1, Player(Index).Pet.X * 32, Player(Index).Pet.Y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NPCNum)
        
        ' if the player blocks, take away the block amount
        RndChance = RAND(1, 1000)
        
        If RndChance <= 250 Then
            RndChance2 = RAND(3, 8)
        Else
            RndChance2 = 0
        End If
        
        BlockAmount = CanPlayerPetBlock(Index)
        Damage = Damage - BlockAmount - RndChance2
        
        ' take away armour
        Damage = Damage - RAND(1, Player(Index).Pet.Stat(Stats.Agility) * 2)
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' * 1.5 if crit hit
        RndChance = RAND(1, 1000)
        
        If RndChance <= 150 Then
            Damage = Damage * 1.5
            Call SendActionMsg(mIndex, "Critical!", BrightCyan, 1, mMapNPC(MapNPCNum).X * 32, mMapNPC(MapNPCNum).Y * 32)
        End If
        
        If CanNpcCrit(Index) Then
            Damage = Damage * 1.5
            Call SendActionMsg(mIndex, "Critical!", BrightCyan, 1, mMapNPC(MapNPCNum).X * 32, mMapNPC(MapNPCNum).Y * 32)
        End If

        If Damage > 0 Then
            Call NPCAttackPet(MapNPCNum, Index, Damage)
        End If
    End If
End Sub

Public Function CanNpcAttackPet(ByVal MapNPCNum As Long, ByVal Index As Long) As Boolean
Dim NPCNum As Long

    NPCNum = mMapNPC(MapNPCNum).Num

    ' Make sure the npc isn't already dead
    If mMapNPC(MapNPCNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If timeGetTime < mMapNPC(MapNPCNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then
        Exit Function
    End If

    mMapNPC(MapNPCNum).AttackTimer = timeGetTime

    ' Make sure they are on the same map
    If Player(Index).Pet.Alive Then
        ' Check if at same coordinates
        If (Player(Index).Pet.Y + 1 = mMapNPC(MapNPCNum).Y) And (Player(Index).Pet.X = mMapNPC(MapNPCNum).X) Then
            CanNpcAttackPet = True
        Else
            If (Player(Index).Pet.Y - 1 = mMapNPC(MapNPCNum).Y) And (Player(Index).Pet.X = mMapNPC(MapNPCNum).X) Then
                CanNpcAttackPet = True
            Else
                If (Player(Index).Pet.Y = mMapNPC(MapNPCNum).Y) And (Player(Index).Pet.X + 1 = mMapNPC(MapNPCNum).X) Then
                    CanNpcAttackPet = True
                Else
                    If (Player(Index).Pet.Y = mMapNPC(MapNPCNum).Y) And (Player(Index).Pet.X - 1 = mMapNPC(MapNPCNum).X) Then
                        CanNpcAttackPet = True
                    End If
                End If
            End If
        End If
    End If
End Function

Public Sub NPCAttackPet(ByVal MapNPCNum As Long, ByVal Victim As Long, ByVal Damage As Long)
Dim Buffer As clsBuffer

    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Call Buffer.WriteLong(SNpcAttack)
    Call Buffer.WriteLong(MapNPCNum)
    Call SendDataToMap(mIndex, Buffer.Serialize)
    
    If Damage <= 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    mMapNPC(MapNPCNum).stopRegen = True
    mMapNPC(MapNPCNum).stopRegenTimer = timeGetTime
    
    ' Say damage
    Call SendActionMsg(GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, Player(Victim).Pet.X * 32, Player(Victim).Pet.Y * 32)
    
    ' send the sound
    Call SendMapSound(Victim, Player(Victim).Pet.X, Player(Victim).Pet.Y, SoundEntity.seNpc, mMapNPC(MapNPCNum).Num)
    Call SendAnimation(mIndex, NPC(mMapNPC(MapNPCNum).Num).Animation, 0, 0, TARGET_TYPE_PET, Victim)
    Call SendBlood(mIndex, Player(Victim).Pet.X, Player(Victim).Pet.Y)
    
    If Damage >= Player(Victim).Pet.Health Then
        ' kill player
        Call PlayerMsg(Victim, "Your " & Player(Victim).Pet.Name & " was killed by a " & NPC(mMapNPC(MapNPCNum).Num).Name & ".", BrightRed)
        Call ReleasePet(Victim)
        
        ' Now that pet is dead, go for owner
        mMapNPC(MapNPCNum).target = Victim
        mMapNPC(MapNPCNum).targetType = TARGET_TYPE_PLAYER
    Else
        ' Player not dead, just do the damage
        Player(Victim).Pet.Health = Player(Victim).Pet.Health - Damage
        Call SendPetVital(Victim, Vitals.HP)
        
        ' set the regen timer
        TempPlayer(Victim).stopRegen = True
        TempPlayer(Victim).stopRegenTimer = timeGetTime
    End If
End Sub

Public Sub Save()
Dim Data() As Byte
Dim f As Long

    Data = Serialize
    
    f = FreeFile
    Open App.Path & "\data\maps\map" + mIndex + ".dat" For Binary As #f
        Put #f, , AryLen(Data)
        Put #f, , Data
    Close #f
End Sub

Public Sub Load()
Dim Data() As Byte
Dim Length As Long
Dim f As Long

    f = FreeFile
    Open App.Path & "\data\maps\map" + mIndex + ".dat" For Binary As #f
        Get #f, , Length
        ReDim Data(Length - 1)
        Get #f, , Data
    Close #f
    
    Call Deserialize(Data)
End Sub
