Attribute VB_Name = "modGlobals"
Option Explicit
'******************************************************
'This is base object decalaration for FMOD sound engine
Public FMOD As New clsFMOD
'******************************************************

'******************************************************
'This is base object decalaration for DirectX8 graphic engine
Public Directx8 As New clsDirectX8
Public D3DDevice8 As Direct3DDevice8
'******************************************************
Public EditorChestType As Byte

'elastic bars
Public BarWidth_GuiHP As Long
Public BarWidth_GuiSP As Long
Public BarWidth_GuiEXP As Long
Public BarWidth_NpcHP(1 To MAX_MAP_NPCS) As Long
Public BarWidth_PlayerHP(1 To MAX_PLAYERS) As Long

Public BarWidth_GuiHP_Max As Long
Public BarWidth_GuiSP_Max As Long
Public BarWidth_GuiEXP_Max As Long
Public BarWidth_NpcHP_Max(1 To MAX_MAP_NPCS) As Long
Public BarWidth_PlayerHP_Max(1 To MAX_PLAYERS) As Long

Public BarWidth_TargetHP As Long
Public BarWidth_TargetHP_Max As Long

' fog
Public fogOffsetX As Long
Public fogOffsetY As Long

' chat bubble
Public chatBubble(1 To MAX_BYTE) As ChatBubbleRec
Public chatBubbleIndex As Long

' Map animations
Public waterfallFrame As Long
Public autoTileFrame As Long

' tutorial
Public tutorialState As Byte

' NPC Chat
Public chatText As String
Public chatOptState() As Byte
Public chatContinueState As Byte
Public CurrentEventIndex As Long
Public tutOpt(1 To 4) As String
Public tutOptState(1 To 4) As Byte

' gui
Public hideGUI As Boolean
Public chatOn As Boolean
Public chatShowLine As String * 1

' map editor boxes
Public shpSelectedTop As Long
Public shpSelectedLeft As Long
Public shpSelectedHeight As Long
Public shpSelectedWidth As Long

' autotiling
Public autoInner(1 To 4) As PointRec
Public autoNW(1 To 4) As PointRec
Public autoNE(1 To 4) As PointRec
Public autoSW(1 To 4) As PointRec
Public autoSE(1 To 4) As PointRec

' fader
Public canFade As Boolean
Public faderAlpha As Long
Public faderState As Long
Public faderSpeed As Long

' menu
Public sUser As String
Public sPass As String
Public sPass2 As String
Public sChar As String
Public inMenu As Boolean
Public curMenu As Long
Public curTextbox As Long

' Cursor
Public GlobalX As Long
Public GlobalY As Long
Public GlobalX_Map As Long
Public GlobalY_Map As Long

' music & sound list cache
Public musicCache() As String
Public soundCache() As String
Public hasPopulated As Boolean

' Hotbar
Public Hotbar(1 To MAX_HOTBAR) As HotbarRec

' Amount of blood decals
Public BloodCount As Long

' Party GUI
Public Const Party_HPWidth As Long = 182
Public Const Party_SPRWidth As Long = 182

' targetting
Public myTarget As Long
Public myTargetType As Long
Public myTargetsTarget As Long

' for directional blocking
Public DirArrowX(1 To 4) As Byte
Public DirArrowY(1 To 4) As Byte

' trading
Public InTrade As Long
Public TradeYourOffer(1 To MAX_INV) As PlayerInvRec
Public TradeTheirOffer(1 To MAX_INV) As PlayerInvRec

' Cache the Resources in an array
Public MapResource() As MapResourceRec
Public Resource_Index As Long
Public Resources_Init As Boolean

' drag + drop
Public DragInvSlotNum As Long
Public DragSpell As Long

' gui
Public tmpCurrencyItem As Long
Public InShop As Long ' is the player in a shop?
Public InBank As Long
Public CurrencyMenu As Byte

' Player variables
Public myID As Long
Public myChar As clsCharacter
Public PlayerInv(1 To MAX_INV) As PlayerInvRec
Public PlayerSpells(1 To MAX_PLAYER_SPELLS) As Long
Public InventoryItemSelected As Long
Public SpellBuffer As Long
Public SpellBufferTimer As Long
Public SpellCD(1 To MAX_PLAYER_SPELLS) As Long
Public StunDuration As Long
Public TNL As Long
Public TNSL(1 To Skills.Skill_Count - 1) As Long

' Stops movement when updating a map
Public CanMoveNow As Boolean

' Debug mode
Public DEBUG_MODE As Boolean

' Game text buffer
Public MyText As String
Public RenderChatText As String
Public ChatScroll As Long
Public ChatButtonUp As Boolean
Public ChatButtonDown As Boolean
Public totalChatLines As Long

' TCP variables
Public PlayerBuffer As String

' Controls main gameloop
Public InGame As Boolean
Public isLogging As Boolean

' Game direction vars
Public ShiftDown As Boolean
Public ControlDown As Boolean
Public tabDown As Boolean
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public DirUpLeft As Boolean
Public DirUpRight As Boolean
Public DirDownLeft As Boolean
Public DirDownRight As Boolean

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if FPS needs to be drawn
Public BLoc As Boolean

' FPS and Time-based movement vars
Public ElapsedTime As Long
Public GameFPS As Long

' Mouse cursor tile location
Public CurX As Long
Public CurY As Long

' Game editors
Public EditorIndex As Long
Public AnimEditorFrame(0 To 1) As Long
Public AnimEditorTimer(0 To 1) As Long

' Used to check if in editor or not and variables for use in editor
Public InMapEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long
Public SpawnNpcNum As Long
Public SpawnNpcDir As Byte
Public EditorShop As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Map Resources
Public ResourceEditorNum As Long

' Used for map editor heal & trap & slide tiles
Public MapEditorHealType As Long
Public MapEditorHealAmount As Long
Public MapEditorSlideDir As Long
Public MapEditorEventIndex As Long
Public MapEditorLightA As Long
Public MapEditorLightR As Long
Public MapEditorLightG As Long
Public MapEditorLightB As Long

' Maximum classes
Public Camera As RECT
Public TileView As RECT

' Pinging
Public PingStart As Long
Public PingEnd As Long
Public Ping As Long

' indexing
Public ActionMsgIndex As Byte
Public BloodIndex As Byte
Public AnimationIndex As Byte

' fps lock
Public FPS_Lock As Boolean

' Editor edited items array
Public Item_Changed(1 To MAX_ITEMS) As Boolean
Public NPC_Changed(1 To MAX_NPCS) As Boolean
Public Resource_Changed(1 To MAX_RESOURCES) As Boolean
Public Animation_Changed(1 To MAX_ANIMATIONS) As Boolean
Public Spell_Changed(1 To MAX_SPELLS) As Boolean
Public Shop_Changed(1 To MAX_SHOPS) As Boolean
Public Event_Changed(1 To MAX_EVENTS) As Boolean

' New char
Public newCharSex As Long
Public newCharClothes As Long
Public newCharGear As Long
Public newCharHair As Long
Public newCharHeadgear As Long

' looping saves
Public Npc_HighIndex As Long
Public Action_HighIndex As Long

' TempStrings for rendering
Public CurrencyText As String
Public CurrencyAcceptState As Byte
Public CurrencyCloseState As Byte
Public Dialogue_ButtonVisible(1 To 3) As Boolean
Public Dialogue_ButtonState(1 To 3) As Byte
Public Dialogue_TitleCaption As String
Public Dialogue_TextCaption As String
Public TradeStatus As String
Public YourWorth As String
Public TheirWorth As String
Public RightMenuButtonState(1 To 4) As Byte

' global dialogue index
Public dialogueIndex As Long
Public dialogueData1 As Long
Public sDialogue As String

Public lastButtonSound As Long
Public lastNpcChatsound As Long

Public RenameType As Long
Public RenameIndex As Long

Public SStatus As String

Public menuAnim As Byte

Public LastProjectile As Integer
Public GME As Byte
Public CharEditState As Byte

Public Last_Dir As Long
Public SocialIcon() As String
Public SocialIconStatus() As Byte

Public CurrentFog As Byte
Public CurrentFogSpeed As Byte
Public CurrentFogOpacity As Byte
Public CurrentTintR As Byte
Public CurrentTintG As Byte
Public CurrentTintB As Byte
Public CurrentTintA As Byte
Public ParallaxX As Long
Public ParallaxY As Long
Public CurTarget As Byte
Public MouseState As Byte
Public MenuNPCAnim As Byte
Public eventAnimFrame As Byte
Public eventAnimTimer As Long
Public DayTime As Boolean
Public isLoading As Boolean
Public BFPS As Boolean
Public GuildScroll As Long

Public MaxSwearWords As Long
