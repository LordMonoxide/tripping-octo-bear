Attribute VB_Name = "modGameEditors"
Option Explicit

Public Sub MapEditorInit()
Dim i As Long

    ' set the width
    frmEditor_Map.width = 7425
    
    ' we're in the map editor
    InMapEditor = True
    
    ' show the form
    frmEditor_Map.visible = True
    
    ' set the scrolly bars
    frmEditor_Map.scrlTileSet.Max = Count_Tileset
    frmEditor_Map.fraTileSet.Caption = "Tileset: " & 1
    frmEditor_Map.scrlTileSet.Value = 1
    
    ' set the scrollbars
    frmEditor_Map.scrlPictureY.Max = (frmEditor_Map.picBackSelect.height \ PIC_Y) - (frmEditor_Map.picBack.height \ PIC_Y)
    frmEditor_Map.scrlPictureX.Max = (frmEditor_Map.picBackSelect.width \ PIC_X) - (frmEditor_Map.picBack.width \ PIC_X)
    MapEditorTileScroll
    
    ' set shops for the shop attribute
    frmEditor_Map.cmbShop.AddItem "None"
    For i = 1 To MAX_SHOPS
        frmEditor_Map.cmbShop.AddItem i & ": " & Shop(i).name
    Next
    
    ' we're not in a shop
    frmEditor_Map.cmbShop.ListIndex = 0
    
        'set chest array
    frmEditor_Map.cmbChestindex.Clear
    For i = 1 To MAX_CHESTS
        frmEditor_Map.cmbChestindex.AddItem "Chest: " & i
    Next
End Sub

Public Sub MapEditorProperties()
Dim x As Long
Dim i As Long
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_MapProperties.lstMusic.Clear
    frmEditor_MapProperties.lstMusic.AddItem "None."
    For i = 1 To UBound(musicCache)
        frmEditor_MapProperties.lstMusic.AddItem musicCache(i)
    Next
    ' finished populating
    
    With frmEditor_MapProperties
        .scrlBoss.Max = MAX_MAP_NPCS
        .txtName.Text = Trim$(map.name)
        
        ' find the music we have set
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        ' rest of it
        .txtUp.Text = CStr(map.Up)
        .txtDown.Text = CStr(map.Down)
        .txtLeft.Text = CStr(map.Left)
        .txtRight.Text = CStr(map.Right)
        .cmbMoral.ListIndex = map.Moral
        .txtBootMap.Text = CStr(map.BootMap)
        .txtBootX.Text = CStr(map.BootX)
        .txtBootY.Text = CStr(map.BootY)
        .scrlBoss = map.BossNpc
        .ScrlFog.Value = map.Fog
        .ScrlFogSpeed.Value = map.FogSpeed
        .scrlFogOpacity.Value = map.FogOpacity
        
        .scrlR.Value = map.Red
        .scrlG.Value = map.Green
        .scrlB.Value = map.Blue
        .scrlA.Value = map.Alpha
        .cmbPanorama.ListIndex = map.Panorama
        .cmbDayNight.ListIndex = map.DayNight

        ' show the map npcs
        .lstNpcs.Clear
        For x = 1 To MAX_MAP_NPCS
            If map.NPC(x) > 0 Then
            .lstNpcs.AddItem x & ": " & Trim$(NPC(map.NPC(x)).name)
            Else
                .lstNpcs.AddItem x & ": No NPC"
            End If
        Next
        .lstNpcs.ListIndex = 0
        
        ' show the npc selection combo
        .cmbNpc.Clear
        .cmbNpc.AddItem "No NPC"
        For x = 1 To MAX_NPCS
            .cmbNpc.AddItem x & ": " & Trim$(NPC(x).name)
        Next
        
        ' set the combo box properly
        Dim tmpString() As String
        Dim npcNum As Long
        tmpString = Split(.lstNpcs.List(.lstNpcs.ListIndex))
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNpc.ListIndex = map.NPC(npcNum)
    
        ' show the current map
        .lblMap.Caption = "Current map: " & GetPlayerMap(MyIndex)
        .txtMaxX.Text = map.MaxX
        .txtMaxY.Text = map.MaxY
    End With
End Sub

Public Sub MapEditorSetTile(ByVal x As Long, ByVal y As Long, ByVal CurLayer As Long, Optional ByVal multitile As Boolean = False, Optional ByVal theAutotile As Byte = 0)
Dim X2 As Long, Y2 As Long

    If theAutotile > 0 Then
        With map.Tile(x, y)
            ' set layer
            .Layer(CurLayer).x = EditorTileX
            .Layer(CurLayer).y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = theAutotile
            cacheRenderState x, y, CurLayer
        End With
        ' do a re-init so we can see our changes
        initAutotiles
        Exit Sub
    End If

    If Not multitile Then ' single
        With map.Tile(x, y)
            ' set layer
            .Layer(CurLayer).x = EditorTileX
            .Layer(CurLayer).y = EditorTileY
            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
            .Autotile(CurLayer) = 0
            cacheRenderState x, y, CurLayer
        End With
    Else ' multitile
        Y2 = 0 ' starting tile for y axis
        For y = CurY To CurY + EditorTileHeight - 1
            X2 = 0 ' re-set x count every y loop
            For x = CurX To CurX + EditorTileWidth - 1
                If x >= 0 And x <= map.MaxX Then
                    If y >= 0 And y <= map.MaxY Then
                        With map.Tile(x, y)
                            .Layer(CurLayer).x = EditorTileX + X2
                            .Layer(CurLayer).y = EditorTileY + Y2
                            .Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                            .Autotile(CurLayer) = 0
                            cacheRenderState x, y, CurLayer
                        End With
                    End If
                End If
                X2 = X2 + 1
            Next
            Y2 = Y2 + 1
        Next
    End If
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal x As Long, ByVal y As Long, Optional ByVal movedMouse As Boolean = True)
Dim i As Long
Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    If Not isInBounds Then Exit Sub
    If Button = vbLeftButton Then
        If frmEditor_Map.optLayers.Value Then
            ' no autotiling
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then 'single tile
                MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
            Else ' multi tile!
                If frmEditor_Map.scrlAutotile.Value = 0 Then
                    MapEditorSetTile CurX, CurY, CurLayer, True
                Else
                    MapEditorSetTile CurX, CurY, CurLayer, , frmEditor_Map.scrlAutotile.Value
                End If
            End If
        ElseIf frmEditor_Map.optAttribs.Value Then
            With map.Tile(CurX, CurY)
                ' blocked tile
                If frmEditor_Map.optBlocked.Value Then
                    .Type = TILE_TYPE_BLOCKED
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' warp tile
                If frmEditor_Map.optWarp.Value Then
                    .Type = TILE_TYPE_WARP
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = 0
                End If
                ' item spawn
                If frmEditor_Map.optItem.Value Then
                    .Type = TILE_TYPE_ITEM
                    .Data1 = ItemEditorNum
                    .Data2 = ItemEditorValue
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' npc avoid
                If frmEditor_Map.optNpcAvoid.Value Then
                    .Type = TILE_TYPE_NPCAVOID
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' resource
                If frmEditor_Map.optResource.Value Then
                    .Type = TILE_TYPE_RESOURCE
                    .Data1 = ResourceEditorNum
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' npc spawn
                If frmEditor_Map.optNpcSpawn.Value Then
                    .Type = TILE_TYPE_NPCSPAWN
                    .Data1 = SpawnNpcNum
                    .Data2 = SpawnNpcDir
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' shop
                If frmEditor_Map.optShop.Value Then
                    .Type = TILE_TYPE_SHOP
                    .Data1 = EditorShop
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' bank
                If frmEditor_Map.optBank.Value Then
                    .Type = TILE_TYPE_BANK
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' heal
                If frmEditor_Map.optHeal.Value Then
                    .Type = TILE_TYPE_HEAL
                    .Data1 = MapEditorHealType
                    .Data2 = MapEditorHealAmount
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' trap
                If frmEditor_Map.optTrap.Value Then
                    .Type = TILE_TYPE_TRAP
                    .Data1 = MapEditorHealAmount
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' slide
                If frmEditor_Map.optSlide.Value Then
                    .Type = TILE_TYPE_SLIDE
                    .Data1 = MapEditorSlideDir
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' Event
                If frmEditor_Map.optEvent.Value Then
                    .Type = TILE_TYPE_EVENT
                    .Data1 = MapEditorEventIndex
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' Threshold
                If frmEditor_Map.optThreshold.Value Then
                    .Type = TILE_TYPE_THRESHOLD
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                End If
                ' Light
                If frmEditor_Map.optLight.Value Then
                    .Type = TILE_TYPE_LIGHT
                    .Data1 = MapEditorLightA
                    .Data2 = MapEditorLightR
                    .Data3 = MapEditorLightG
                    .Data4 = MapEditorLightB
                End If
                ' Arena
                If frmEditor_Map.optArena.Value Then
                    .Type = TILE_TYPE_ARENA
                    .Data1 = EditorWarpMap
                    .Data2 = EditorWarpX
                    .Data3 = EditorWarpY
                    .Data4 = 0
                End If
                If frmEditor_Map.OptChest.Value Then
                'Data1 has to go first because the renderer likes to jump in and give RTE9
                    .Data1 = frmEditor_Map.cmbChestindex.ListIndex + 1
                    .Type = TILE_TYPE_CHEST
                    .Data2 = frmEditor_Map.txtchestdata1.Text
                    .Data3 = frmEditor_Map.txtchestdata2.Text
                    .Data4 = ""
                    'Map data must be sent to the server so that any old chest could be have the tile_type removed
                    With Chest(.Data1)
                        .map = Player(MyIndex).map
                        .x = CurX
                        .y = CurY
                    End With
                End If
            End With
        ElseIf frmEditor_Map.optBlock.Value Then
            If movedMouse Then Exit Sub
            ' find what tile it is
            x = x - ((x \ 32) * 32)
            y = y - ((y \ 32) * 32)
            ' see if it hits an arrow
            For i = 1 To 4
                If x >= DirArrowX(i) And x <= DirArrowX(i) + 8 Then
                    If y >= DirArrowY(i) And y <= DirArrowY(i) + 8 Then
                        ' flip the value.
                        setDirBlock map.Tile(CurX, CurY).DirBlock, CByte(i), Not isDirBlocked(map.Tile(CurX, CurY).DirBlock, CByte(i))
                        Exit Sub
                    End If
                End If
            Next
        End If
    End If

    If Button = vbRightButton Then
        If frmEditor_Map.optLayers.Value Then
            With map.Tile(CurX, CurY)
                ' clear layer
                .Layer(CurLayer).x = 0
                .Layer(CurLayer).y = 0
                .Layer(CurLayer).Tileset = 0
                If .Autotile(CurLayer) > 0 Then
                    .Autotile(CurLayer) = 0
                    ' do a re-init so we can see our changes
                    initAutotiles
                End If
                cacheRenderState x, y, CurLayer
            End With
        ElseIf frmEditor_Map.optAttribs.Value Then
            With map.Tile(CurX, CurY)
                ' clear attribute
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
            End With
        End If
    End If

    CacheResources
End Sub

Public Sub MapEditorChooseTile(Button As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = x \ PIC_X
        EditorTileY = y \ PIC_Y
        
        shpSelectedTop = EditorTileY * PIC_Y
        shpSelectedLeft = EditorTileX * PIC_X
        
        shpSelectedWidth = PIC_X
        shpSelectedHeight = PIC_Y
    End If
End Sub

Public Sub MapEditorDrag(Button As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ' convert the pixel number to tile number
        x = (x \ PIC_X) + 1
        y = (y \ PIC_Y) + 1
        ' check it's not out of bounds
        If x < 0 Then x = 0
        If x > frmEditor_Map.picBackSelect.width / PIC_X Then x = frmEditor_Map.picBackSelect.width / PIC_X
        If y < 0 Then y = 0
        If y > frmEditor_Map.picBackSelect.height / PIC_Y Then y = frmEditor_Map.picBackSelect.height / PIC_Y
        ' find out what to set the width + height of map editor to
        If x > EditorTileX Then ' drag right
            EditorTileWidth = x - EditorTileX
        Else ' drag left
            ' TO DO
        End If
        If y > EditorTileY Then ' drag down
            EditorTileHeight = y - EditorTileY
        Else ' drag up
            ' TO DO
        End If
        shpSelectedWidth = EditorTileWidth * PIC_X
        shpSelectedHeight = EditorTileHeight * PIC_Y
    End If
End Sub

Public Sub MapEditorTileScroll()
    ' horizontal scrolling
    If frmEditor_Map.picBackSelect.width < frmEditor_Map.picBack.width Then
        frmEditor_Map.scrlPictureX.Enabled = False
    Else
        frmEditor_Map.scrlPictureX.Enabled = True
        frmEditor_Map.picBackSelect.Left = (frmEditor_Map.scrlPictureX.Value * PIC_X) * -1
    End If
    
    ' vertical scrolling
    If frmEditor_Map.picBackSelect.height < frmEditor_Map.picBack.height Then
        frmEditor_Map.scrlPictureY.Enabled = False
    Else
        frmEditor_Map.scrlPictureY.Enabled = True
        frmEditor_Map.picBackSelect.Top = (frmEditor_Map.scrlPictureY.Value * PIC_Y) * -1
    End If
End Sub

Public Sub MapEditorSend()
    Call SendMap
    InMapEditor = False
    Unload frmEditor_Map
End Sub

Public Sub MapEditorCancel()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteLong 1
    SendData Buffer.ToArray()
    InMapEditor = False
    Unload frmEditor_Map
End Sub

Public Sub MapEditorClearLayer()
Dim i As Long
Dim x As Long
Dim y As Long
Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next
    
    If CurLayer = 0 Then Exit Sub

    ' ask to clear layer
    If MsgBox("Are you sure you wish to clear this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For x = 0 To map.MaxX
            For y = 0 To map.MaxY
                map.Tile(x, y).Layer(CurLayer).x = 0
                map.Tile(x, y).Layer(CurLayer).y = 0
                map.Tile(x, y).Layer(CurLayer).Tileset = 0
                cacheRenderState x, y, CurLayer
            Next
        Next
        
        ' re-cache autos
        initAutotiles
    End If
End Sub

Public Sub MapEditorFillLayer()
Dim i As Long
Dim x As Long
Dim y As Long
Dim CurLayer As Long

    ' find which layer we're on
    For i = 1 To MapLayer.Layer_Count - 1
        If frmEditor_Map.optLayer(i).Value Then
            CurLayer = i
            Exit For
        End If
    Next

    ' Ground layer
    If MsgBox("Are you sure you wish to fill this layer?", vbYesNo, Options.Game_Name) = vbYes Then
        For x = 0 To map.MaxX
            For y = 0 To map.MaxY
                map.Tile(x, y).Layer(CurLayer).x = EditorTileX
                map.Tile(x, y).Layer(CurLayer).y = EditorTileY
                map.Tile(x, y).Layer(CurLayer).Tileset = frmEditor_Map.scrlTileSet.Value
                map.Tile(x, y).Autotile(CurLayer) = frmEditor_Map.scrlAutotile.Value
                cacheRenderState x, y, CurLayer
            Next
        Next
        
        ' now cache the positions
        initAutotiles
    End If
End Sub

Public Sub MapEditorClearAttribs()
Dim x As Long
Dim y As Long

    If MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, Options.Game_Name) = vbYes Then
        For x = 0 To map.MaxX
            For y = 0 To map.MaxY
                map.Tile(x, y).Type = 0
            Next
        Next
    End If
End Sub

Public Sub MapEditorLeaveMap()
    If InMapEditor Then
        If MsgBox("Save changes to current map?", vbYesNo) = vbYes Then
            Call MapEditorSend
        Else
            Call MapEditorCancel
        End If
    End If
End Sub

' /////////////////
' // Item Editor //
' /////////////////
Public Sub ItemEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    If frmEditor_Item.visible = False Then Exit Sub
    EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Item.cmbSound.Clear
    frmEditor_Item.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Item.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Item(EditorIndex)
        frmEditor_Item.txtName.Text = Trim$(.name)
        If .Pic > frmEditor_Item.scrlPic.Max Then .Pic = 0
        frmEditor_Item.scrlPic.Value = .Pic
        frmEditor_Item.cmbType.ListIndex = .Type
        frmEditor_Item.scrlAnim.Value = .Animation
        frmEditor_Item.txtDesc.Text = Trim$(.Desc)
        frmEditor_Item.chkStackable.Value = .Stackable
        frmEditor_Item.TxtContainerItem(0).Text = .Container(0)
        frmEditor_Item.TxtContainerItem(1).Text = .Container(1)
        frmEditor_Item.TxtContainerItem(2).Text = .Container(2)
        frmEditor_Item.TxtContainerItem(3).Text = .Container(3)
        frmEditor_Item.TxtContainerItem(4).Text = .Container(4)
        frmEditor_Item.TxtContainerChance(0).Text = .ContainerChance(0)
        frmEditor_Item.TxtContainerChance(1).Text = .ContainerChance(1)
        frmEditor_Item.TxtContainerChance(2).Text = .ContainerChance(2)
        frmEditor_Item.TxtContainerChance(3).Text = .ContainerChance(3)
        frmEditor_Item.TxtContainerChance(4).Text = .ContainerChance(4)
        
        
        ' find the sound we have set
        If frmEditor_Item.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Item.cmbSound.ListCount
                If frmEditor_Item.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Item.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Item.cmbSound.ListIndex = -1 Then frmEditor_Item.cmbSound.ListIndex = 0
        End If

        ' Type specific settings
        If (frmEditor_Item.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraEquipment.visible = True
            frmEditor_Item.scrlDamage.Value = .Data2
            frmEditor_Item.cmbTool.ListIndex = .Data3
            frmEditor_Item.chkTwoHanded.Value = .isTwoHanded
            frmEditor_Item.scrlProjectilePic.Value = .Projectile
            frmEditor_Item.scrlProjectileRange.Value = .Range
            frmEditor_Item.scrlProjectileRotation.Value = .Rotation
            frmEditor_Item.scrlProjectileAmmo.Value = .Ammo
            If .PDef > 0 Then
                frmEditor_Item.txtPDef.Text = Trim$(.PDef)
                Else
                    frmEditor_Item.txtPDef.Text = "0"
                End If
                If .MDef > 0 Then
                    frmEditor_Item.txtMDef.Text = Trim$(.MDef)
                Else
                    frmEditor_Item.txtMDef.Text = "0"
                End If
                If .RDef > 0 Then
                    frmEditor_Item.txtRDef.Text = Trim$(.RDef)
                Else
                    frmEditor_Item.txtRDef.Text = "0"
                End If

            If .Speed < 100 Then .Speed = 100
            frmEditor_Item.scrlSpeed.Value = .Speed
            
            ' loop for stats
            For i = 1 To Stats.Stat_Count - 1
                frmEditor_Item.scrlStatBonus(i).Value = .Add_Stat(i)
            Next
        Else
            frmEditor_Item.fraEquipment.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_Aura) Then
            frmEditor_Item.fraAura.visible = True
            frmEditor_Item.scrlAura.Value = .Aura
        Else
            frmEditor_Item.fraAura.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_WEAPON) Then
            frmEditor_Item.fraWeapon.visible = True
        Else
            frmEditor_Item.fraWeapon.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex > ITEM_TYPE_WEAPON) And (frmEditor_Item.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmEditor_Item.fraArmor.visible = True
        Else
            frmEditor_Item.fraArmor.visible = False
        End If
        
        If frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONSUME Then
            frmEditor_Item.fraVitals.visible = True
            frmEditor_Item.scrlAddHp.Value = .AddHP
            frmEditor_Item.scrlAddMP.Value = .AddMP
            frmEditor_Item.scrlAddExp.Value = .AddEXP
        Else
            frmEditor_Item.fraVitals.visible = False
        End If

        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_CONTAINER) Then
            frmEditor_Item.fraContainer.visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.visible = False
        End If
        

        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
            frmEditor_Item.fraSpell.visible = True
            frmEditor_Item.scrlSpell.Value = .Data1
        Else
            frmEditor_Item.fraSpell.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_PET) Then
            frmEditor_Item.fraPet.visible = True
            frmEditor_Item.scrlPet.Value = .Data1
        Else
            frmEditor_Item.fraPet.visible = False
        End If
        
        If (frmEditor_Item.cmbType.ListIndex = ITEM_TYPE_PET_STATS) Then
            frmEditor_Item.fraPetStats.visible = True
            frmEditor_Item.cmbPetStat.ListIndex = .Data1
            frmEditor_Item.optIncDec(.Data2).Value = True
            frmEditor_Item.scrlPetPercent.Value = .Data3
        Else
            frmEditor_Item.fraPetStats.visible = False
        End If
        

        ' Basic requirements
        frmEditor_Item.scrlAccessReq.Value = .AccessReq
        frmEditor_Item.scrlLevelReq.Value = .LevelReq
        
        ' loop for stats
        For i = 1 To Stats.Stat_Count - 1
            frmEditor_Item.scrlStatReq(i).Value = .Stat_Req(i)
        Next
        
        ' Info
        frmEditor_Item.txtPrice.Text = .Price
        frmEditor_Item.cmbBind.ListIndex = .BindType
        frmEditor_Item.scrlRarity.Value = .Rarity
         
        EditorIndex = frmEditor_Item.lstIndex.ListIndex + 1
    End With

    Item_Changed(EditorIndex) = True
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    For i = 1 To MAX_ITEMS
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmEditor_Item
    ClearChanged_Item
End Sub

Public Sub ItemEditorCancel()
    Unload frmEditor_Item
    ClearChanged_Item
    ClearItems
    SendRequestItems
End Sub

Public Sub ClearChanged_Item()
    ZeroMemory Item_Changed(1), MAX_ITEMS * 2 ' 2 = boolean length
End Sub

' /////////////////
' // Animation Editor //
' /////////////////
Public Sub AnimationEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    If frmEditor_Animation.visible = False Then Exit Sub
    EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Animation.cmbSound.Clear
    frmEditor_Animation.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Animation.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating

    With Animation(EditorIndex)
        frmEditor_Animation.txtName.Text = Trim$(.name)
        
        ' find the sound we have set
        If frmEditor_Animation.cmbSound.ListCount >= 0 Then
            For i = 0 To frmEditor_Animation.cmbSound.ListCount
                If frmEditor_Animation.cmbSound.List(i) = Trim$(.Sound) Then
                    frmEditor_Animation.cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or frmEditor_Animation.cmbSound.ListIndex = -1 Then frmEditor_Animation.cmbSound.ListIndex = 0
        End If
        
        For i = 0 To 1
            frmEditor_Animation.scrlSprite(i).Value = .Sprite(i)
            frmEditor_Animation.scrlFrameCount(i).Value = .Frames(i)
            frmEditor_Animation.scrlLoopCount(i).Value = .LoopCount(i)
            
            If .looptime(i) > 0 Then
                frmEditor_Animation.scrlLoopTime(i).Value = .looptime(i)
            Else
                frmEditor_Animation.scrlLoopTime(i).Value = 45
            End If
            
        Next
         
        EditorIndex = frmEditor_Animation.lstIndex.ListIndex + 1
    End With

    Animation_Changed(EditorIndex) = True
End Sub

Public Sub AnimationEditorOk()
Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Animation_Changed(i) Then
            Call SendSaveAnimation(i)
        End If
    Next
    
    Unload frmEditor_Animation
    ClearChanged_Animation
End Sub

Public Sub AnimationEditorCancel()
    Unload frmEditor_Animation
    ClearChanged_Animation
    ClearAnimations
    SendRequestAnimations
End Sub

Public Sub ClearChanged_Animation()
    ZeroMemory Animation_Changed(1), MAX_ANIMATIONS * 2 ' 2 = boolean length
End Sub

' ////////////////
' // Npc Editor //
' ////////////////
Public Sub NpcEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    If frmEditor_NPC.visible = False Then Exit Sub
    EditorIndex = frmEditor_NPC.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_NPC.cmbSound.Clear
    frmEditor_NPC.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_NPC.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_NPC
        .scrlDrop.Max = MAX_NPC_DROPS
        .scrlSpell.Max = MAX_NPC_SPELLS
        .txtName.Text = Trim$(NPC(EditorIndex).name)
        .txtAttackSay.Text = Trim$(NPC(EditorIndex).AttackSay)
        If NPC(EditorIndex).Sprite < 0 Or NPC(EditorIndex).Sprite > .scrlSprite.Max Then NPC(EditorIndex).Sprite = 0
        .scrlSprite.Value = NPC(EditorIndex).Sprite
        .txtSpawnSecs.Text = CStr(NPC(EditorIndex).SpawnSecs)
        .cmbBehaviour.ListIndex = NPC(EditorIndex).Behaviour
        .scrlRange.Value = NPC(EditorIndex).Range
        .txtHP.Text = NPC(EditorIndex).HP
        .txtExp.Text = NPC(EditorIndex).EXP
        .txtEXP_max.Text = NPC(EditorIndex).EXP_max
        .txtLevel.Text = NPC(EditorIndex).Level
        .txtDamage.Text = NPC(EditorIndex).Damage
        .chkQuest.Value = NPC(EditorIndex).Quest
        .scrlQuest.Value = NPC(EditorIndex).QuestNum
        .scrlEvent.Value = NPC(EditorIndex).Event
'        .scrlProjectilePic.Value = NPC(EditorIndex).Projectile
        .scrlProjectileRange.Value = NPC(EditorIndex).ProjectileRange
        .scrlProjectileRotation.Value = NPC(EditorIndex).Rotation
        .cmbMoral.ListIndex = NPC(EditorIndex).Moral
        .scrlA.Value = NPC(EditorIndex).A
        .scrlR.Value = NPC(EditorIndex).R
        .scrlG.Value = NPC(EditorIndex).G
        .scrlB.Value = NPC(EditorIndex).B
        .chkDay.Value = NPC(EditorIndex).SpawnAtDay
        .chkNight.Value = NPC(EditorIndex).SpawnAtNight
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(NPC(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
        
        For i = 1 To Stats.Stat_Count - 1
            .scrlStat(i).Value = NPC(EditorIndex).stat(i)
        Next
        
        ' show 1 data
        .scrlDrop.Value = 1
        .scrlSpell.Value = 1
    End With
    
    NPC_Changed(EditorIndex) = True
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    For i = 1 To MAX_NPCS
        If NPC_Changed(i) Then
            Call SendSaveNpc(i)
        End If
    Next
    
    Unload frmEditor_NPC
    ClearChanged_NPC
End Sub

Public Sub NpcEditorCancel()
    Unload frmEditor_NPC
    ClearChanged_NPC
    ClearNpcs
    SendRequestNPCS
End Sub

Public Sub ClearChanged_NPC()
    ZeroMemory NPC_Changed(1), MAX_NPCS * 2 ' 2 = boolean length
End Sub

' ////////////////
' // Resource Editor //
' ////////////////
Public Sub ResourceEditorInit()
Dim i As Long
Dim SoundSet As Boolean

    If frmEditor_Resource.visible = False Then Exit Sub
    EditorIndex = frmEditor_Resource.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Resource.cmbSound.Clear
    frmEditor_Resource.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Resource.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Resource
        .scrlExhaustedPic.Max = Count_Resource
        .scrlNormalPic.Max = Count_Resource
        .scrlAnimation.Max = MAX_ANIMATIONS
        
        .txtName.Text = Trim$(Resource(EditorIndex).name)
        .txtMessage.Text = Trim$(Resource(EditorIndex).SuccessMessage)
        .txtMessage2.Text = Trim$(Resource(EditorIndex).EmptyMessage)
        .txtExp.Text = Trim$(Resource(EditorIndex).EXP)
        .cmbType.ListIndex = Resource(EditorIndex).ResourceType
        .scrlNormalPic.Value = Resource(EditorIndex).ResourceImage
        .scrlExhaustedPic.Value = Resource(EditorIndex).ExhaustedImage
        .scrlReward.Value = Resource(EditorIndex).ItemReward
        .scrlTool.Value = Resource(EditorIndex).ToolRequired
        .txtHealth.Text = Resource(EditorIndex).Health
        .scrlRespawn.Value = Resource(EditorIndex).RespawnTime
        .scrlAnimation.Value = Resource(EditorIndex).Animation
        .scrlChance.Value = Resource(EditorIndex).Chance
        
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(Resource(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Resource_Changed(EditorIndex) = True
End Sub

Public Sub ResourceEditorOk()
Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Resource_Changed(i) Then
            Call SendSaveResource(i)
        End If
    Next
    
    Unload frmEditor_Resource
    ClearChanged_Resource
End Sub

Public Sub ResourceEditorCancel()
    Unload frmEditor_Resource
    ClearChanged_Resource
    ClearResources
    SendRequestResources
End Sub

Public Sub ClearChanged_Resource()
    ZeroMemory Resource_Changed(1), MAX_RESOURCES * 2 ' 2 = boolean length
End Sub

' /////////////////
' // Shop Editor //
' /////////////////
Public Sub ShopEditorInit()
Dim i As Long
    
    If frmEditor_Shop.visible = False Then Exit Sub
    EditorIndex = frmEditor_Shop.lstIndex.ListIndex + 1
    frmEditor_Shop.scrlShoptype.Value = Shop(EditorIndex).ShopType
    frmEditor_Shop.txtName.Text = Trim$(Shop(EditorIndex).name)
    If Shop(EditorIndex).BuyRate > 0 Then
        frmEditor_Shop.scrlBuy.Value = Shop(EditorIndex).BuyRate
    Else
        frmEditor_Shop.scrlBuy.Value = 100
    End If
    
    frmEditor_Shop.cmbItem.Clear
    frmEditor_Shop.cmbItem.AddItem "None"
    frmEditor_Shop.cmbCostItem.Clear
    frmEditor_Shop.cmbCostItem.AddItem "None"
    frmEditor_Shop.cmbCostItem2.Clear
    frmEditor_Shop.cmbCostItem2.AddItem "None"

    For i = 1 To MAX_ITEMS
        frmEditor_Shop.cmbItem.AddItem i & ": " & Trim$(Item(i).name)
        frmEditor_Shop.cmbCostItem.AddItem i & ": " & Trim$(Item(i).name)
        frmEditor_Shop.cmbCostItem2.AddItem i & ": " & Trim$(Item(i).name)
    Next

    frmEditor_Shop.cmbItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem.ListIndex = 0
    frmEditor_Shop.cmbCostItem2.ListIndex = 0
    
    UpdateShopTrade
    
    Shop_Changed(EditorIndex) = True
End Sub

Public Sub UpdateShopTrade(Optional ByVal tmpPos As Long = 0)
Dim i As Long

    frmEditor_Shop.lstTradeItem.Clear

    For i = 1 To MAX_TRADES
        With Shop(EditorIndex).TradeItem(i)
            ' if none, show as none
            If .Item = 0 And .CostItem = 0 And .CostItem2 = 0 Then
                frmEditor_Shop.lstTradeItem.AddItem "Empty Trade Slot"
            Else
                If .CostItem And .CostItem2 > 0 Then
                    frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).name) & " and " & .CostValue2 & "x " & Trim$(Item(.CostItem2).name)
                ElseIf .CostItem > 0 Then
                    frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).name) & " for " & .CostValue & "x " & Trim$(Item(.CostItem).name)
                ElseIf .CostItem2 > 0 Then
                    frmEditor_Shop.lstTradeItem.AddItem i & ": " & .ItemValue & "x " & Trim$(Item(.Item).name) & " for " & .CostValue2 & "x " & Trim$(Item(.CostItem2).name)
                End If
            End If
        End With
    Next

    frmEditor_Shop.lstTradeItem.ListIndex = tmpPos
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    For i = 1 To MAX_SHOPS
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmEditor_Shop
    ClearChanged_Shop
End Sub

Public Sub ShopEditorCancel()
    Unload frmEditor_Shop
    ClearChanged_Shop
    ClearShops
    SendRequestShops
End Sub

Public Sub ClearChanged_Shop()
    ZeroMemory Shop_Changed(1), MAX_SHOPS * 2 ' 2 = boolean length
End Sub

' //////////////////
' // Spell Editor //
' //////////////////
Public Sub SpellEditorInit()
Dim i As Long
Dim SoundSet As Boolean
    
    If frmEditor_Spell.visible = False Then Exit Sub
    EditorIndex = frmEditor_Spell.lstIndex.ListIndex + 1
    
    ' populate the cache if we need to
    If Not hasPopulated Then
        PopulateLists
    End If
    ' add the array to the combo
    frmEditor_Spell.cmbSound.Clear
    frmEditor_Spell.cmbSound.AddItem "None."
    For i = 1 To UBound(soundCache)
        frmEditor_Spell.cmbSound.AddItem soundCache(i)
    Next
    ' finished populating
    
    With frmEditor_Spell
        ' set max values
        .scrlAnimCast.Max = MAX_ANIMATIONS
        .scrlAnim.Max = MAX_ANIMATIONS
        .scrlAOE.Max = MAX_BYTE
        .scrlRange.Max = MAX_BYTE
        .scrlMap.Max = MAX_MAPS
        
        
        ' set values
        .txtName.Text = Trim$(spell(EditorIndex).name)
        .txtDesc.Text = Trim$(spell(EditorIndex).Desc)
        .cmbType.ListIndex = spell(EditorIndex).Type
        .scrlMP.Value = spell(EditorIndex).MPCost
        .scrlLevel.Value = spell(EditorIndex).LevelReq
        .scrlAccess.Value = spell(EditorIndex).AccessReq
        .scrlCast.Value = spell(EditorIndex).CastTime
        .scrlCool.Value = spell(EditorIndex).CDTime
        .scrlIcon.Value = spell(EditorIndex).Icon
        .scrlMap.Value = spell(EditorIndex).map
        .scrlX.Value = spell(EditorIndex).x
        .scrlY.Value = spell(EditorIndex).y
        .scrlDir.Value = spell(EditorIndex).dir
        .txtHPVital.Text = spell(EditorIndex).Vital(Vitals.HP)
        .txtMPVital.Text = spell(EditorIndex).Vital(Vitals.MP)
        .optHPVital(spell(EditorIndex).VitalType(Vitals.HP)).Value = True
        .optMPVital(spell(EditorIndex).VitalType(Vitals.MP)).Value = True
        .scrlDuration.Value = spell(EditorIndex).Duration
        .scrlInterval.Value = spell(EditorIndex).Interval
        .scrlRange.Value = spell(EditorIndex).Range
        If spell(EditorIndex).IsAoE Then
            .chkAOE.Value = 1
        Else
            .chkAOE.Value = 0
        End If
        .scrlAOE.Value = spell(EditorIndex).AoE
        .scrlAnimCast.Value = spell(EditorIndex).CastAnim
        .scrlAnim.Value = spell(EditorIndex).SpellAnim
        .scrlStun.Value = spell(EditorIndex).StunDuration
        .cmbBuffType.ListIndex = spell(EditorIndex).BuffType
        ' find the sound we have set
        If .cmbSound.ListCount >= 0 Then
            For i = 0 To .cmbSound.ListCount
                If .cmbSound.List(i) = Trim$(spell(EditorIndex).Sound) Then
                    .cmbSound.ListIndex = i
                    SoundSet = True
                End If
            Next
            If Not SoundSet Or .cmbSound.ListIndex = -1 Then .cmbSound.ListIndex = 0
        End If
    End With
    
    Spell_Changed(EditorIndex) = True
End Sub

Public Sub SpellEditorOk()
Dim i As Long

    For i = 1 To MAX_SPELLS
        If Spell_Changed(i) Then
            Call SendSaveSpell(i)
        End If
    Next
    
    Unload frmEditor_Spell
    ClearChanged_Spell
End Sub

Public Sub SpellEditorCancel()
    Unload frmEditor_Spell
    ClearChanged_Spell
    ClearSpells
    SendRequestSpells
End Sub

Public Sub ClearChanged_Spell()
    ZeroMemory Spell_Changed(1), MAX_SPELLS * 2 ' 2 = boolean length
End Sub
Public Sub Events_ClearChanged()
    Dim i As Long
    For i = 1 To MAX_EVENTS
        Event_Changed(i) = False
    Next i
End Sub

Public Sub EventEditorInit()
    With frmEditor_Events
        If .visible = False Then Exit Sub
        EditorIndex = .lstIndex.ListIndex + 1
        .txtName = Trim$(Events(EditorIndex).name)
        .chkPlayerSwitch.Value = Events(EditorIndex).chkSwitch
        .chkPlayerVar.Value = Events(EditorIndex).chkVariable
        .chkHasItem.Value = Events(EditorIndex).chkHasItem
        .cmbPlayerSwitch.ListIndex = Events(EditorIndex).SwitchIndex
        .cmbPlayerSwitchCompare.ListIndex = Events(EditorIndex).SwitchCompare
        .cmbPlayerVar.ListIndex = Events(EditorIndex).VariableIndex
        .cmbPlayerVarCompare.ListIndex = Events(EditorIndex).VariableCompare
        .txtPlayerVariable.Text = Events(EditorIndex).VariableCondition
        .cmbHasItem.ListIndex = Events(EditorIndex).HasItemIndex - 1
        .cmbTrigger.ListIndex = Events(EditorIndex).Trigger
        .chkWalkthrought.Value = Events(EditorIndex).WalkThrought
        .chkAnimated.Value = Events(EditorIndex).Animated
        .scrlGraphic.Value = Events(EditorIndex).Graphic(0)
        Call .PopulateSubEventList
    End With
    Event_Changed(EditorIndex) = True
End Sub

Public Sub EventEditorOk()
Dim i As Long
    For i = 1 To MAX_EVENTS
        If Event_Changed(i) Then
            Call Events_SendSaveEvent(i)
        End If
    Next i
    
    Unload frmEditor_Events
    Events_ClearChanged
End Sub

Public Sub EventEditorCancel()
    Unload frmEditor_Events
    Events_ClearChanged
    ClearEvents
    Events_SendRequestEventsData
End Sub

' *********************
' ** Event Utilities **
' *********************
Public Function GetSubEventCount(ByVal Index As Long)
    GetSubEventCount = 0
    If Index <= 0 Or Index > MAX_EVENTS Then Exit Function
    If Events(Index).HasSubEvents Then
        GetSubEventCount = UBound(Events(Index).SubEvents)
    End If
End Function

Public Sub ClearAttributeDialogue()
    frmEditor_Map.fraNpcSpawn.visible = False
    frmEditor_Map.fraResource.visible = False
    frmEditor_Map.fraMapItem.visible = False
    frmEditor_Map.fraMapWarp.visible = False
    frmEditor_Map.fraShop.visible = False
    frmEditor_Map.fraEvent.visible = False
    frmEditor_Map.fraLight.visible = False
    frmEditor_Map.fraHeal.visible = False
End Sub
