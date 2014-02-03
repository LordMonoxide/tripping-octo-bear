Attribute VB_Name = "modConstants"
Option Explicit

' API
Public Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal socket As clsSocket, ByVal data As clsBuffer, ByVal wParam As Long, ByVal lParam As Long) As Long

' path constants
Public Const ADMIN_LOG As String = "admin.log"
Public Const PLAYER_LOG As String = "player.log"

Public Const MAX_LINES As Long = 500 ' Used for frmServer.txtText

' ********************************************************
' * The values below must match with the client's values *
' ********************************************************
' General constants
Public Const MAX_PLAYERS As Long = 70
Public Const MAX_ITEMS As Long = 255
Public Const MAX_NPCS As Long = 255
Public Const MAX_ANIMATIONS As Long = 255
Public Const MAX_INV As Long = 35
Public Const MAX_MAP_ITEMS As Long = 255
Public Const MAX_MAP_NPCS As Long = 30
Public Const MAX_SHOPS As Long = 50
Public Const MAX_PLAYER_SPELLS As Long = 35
Public Const MAX_SPELLS As Long = 255
Public Const MAX_TRADES As Long = 30
Public Const MAX_RESOURCES As Long = 100
Public Const MAX_LEVELS As Long = 99
Public Const MAX_BANK As Long = 99
Public Const MAX_HOTBAR As Long = 9
Public Const MAX_NPC_DROPS As Byte = 30
Public Const MAX_NPC_SPELLS As Byte = 10
Public Const MAX_EVENTS As Long = 255
Public Const MAX_SWITCHES As Long = 1000
Public Const MAX_VARIABLES As Long = 1000
'Chest constants
Public Const NumChests As Long = 300

' server-side stuff
Public Const ITEM_SPAWN_TIME As Long = 30000 ' 30 seconds
Public Const ITEM_DESPAWN_TIME As Long = 90000 ' 1:30 seconds
Public Const MAX_DOTS As Long = 30
Public Const MAX_GUILD_LOGO As Long = 10

' text color constants
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16

Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Boolean constants
Public Const NO As Byte = 0
Public Const YES As Byte = 1

' String constants
Public Const NAME_LENGTH As Byte = 20
Public Const ACCOUNT_LENGTH As Byte = 12

' Sex constants
Public Const SEX_MALE As Byte = 0
Public Const SEX_FEMALE As Byte = 1

' Map constants
Public Const MAX_MAPS As Long = 100
Public Const MAX_MAPX As Byte = 24
Public Const MAX_MAPY As Byte = 18
Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_SAFE As Byte = 1
Public Const MAP_MORAL_BOSS As Byte = 2

' Tile consants
Public Const TILE_TYPE_WALKABLE As Byte = 0
Public Const TILE_TYPE_BLOCKED As Byte = 1
Public Const TILE_TYPE_WARP As Byte = 2
Public Const TILE_TYPE_ITEM As Byte = 3
Public Const TILE_TYPE_NPCAVOID As Byte = 4
Public Const TILE_TYPE_RESOURCE As Byte = 5
Public Const TILE_TYPE_NPCSPAWN As Byte = 6
Public Const TILE_TYPE_SHOP As Byte = 7
Public Const TILE_TYPE_BANK As Byte = 8
Public Const TILE_TYPE_HEAL As Byte = 9
Public Const TILE_TYPE_TRAP As Byte = 10
Public Const TILE_TYPE_SLIDE As Byte = 11
Public Const TILE_TYPE_EVENT As Byte = 12
Public Const TILE_TYPE_THRESHOLD As Byte = 13
Public Const TILE_TYPE_LIGHT As Byte = 14
Public Const TILE_TYPE_ARENA As Byte = 15
Public Const CHEST_TYPE_STAT As Byte = 0
Public Const CHEST_TYPE_ITEM As Byte = 1
Public Const CHEST_TYPE_GOLD As Byte = 2
Public Const CHEST_TYPE_EXP As Byte = 3
Public Const MAX_CHESTS As Long = 200
Public Const TILE_TYPE_CHEST As Byte = 16

' Item constants
Public Const ITEM_TYPE_NONE As Byte = 0
Public Const ITEM_TYPE_WEAPON As Byte = 1
Public Const ITEM_TYPE_ARMOR As Byte = 2
Public Const ITEM_TYPE_Aura As Byte = 3
Public Const ITEM_TYPE_SHIELD As Byte = 4
Public Const ITEM_TYPE_CONSUME As Byte = 5
Public Const ITEM_TYPE_CURRENCY As Byte = 6
Public Const ITEM_TYPE_SPELL As Byte = 7
Public Const ITEM_TYPE_UNIQUE As Byte = 8
Public Const ITEM_TYPE_LOGO_GUILD As Byte = 9
Public Const ITEM_TYPE_CONTAINER As Byte = 10

' Direction constants
Public Const DIR_UP As Byte = 0
Public Const DIR_DOWN As Byte = 1
Public Const DIR_LEFT As Byte = 2
Public Const DIR_RIGHT As Byte = 3
Public Const DIR_UP_LEFT As Byte = 4
Public Const DIR_UP_RIGHT As Byte = 5
Public Const DIR_DOWN_LEFT As Byte = 6
Public Const DIR_DOWN_RIGHT As Byte = 7

' Constants for player movement
Public Const MOVING_WALKING As Byte = 1
Public Const MOVING_RUNNING As Byte = 2

' Admin constants
Public Const ADMIN_MONITOR As Byte = 1
Public Const ADMIN_MAPPER As Byte = 2
Public Const ADMIN_DEVELOPER As Byte = 3
Public Const ADMIN_CREATOR As Byte = 4

' NPC constants
Public Const NPC_BEHAVIOUR_ATTACKONSIGHT As Byte = 0
Public Const NPC_BEHAVIOUR_ATTACKWHENATTACKED As Byte = 1
Public Const NPC_BEHAVIOUR_FRIENDLY As Byte = 2
Public Const NPC_BEHAVIOUR_SHOPKEEPER As Byte = 3
Public Const NPC_BEHAVIOUR_GUARD As Byte = 4

' Spell constants
Public Const SPELL_TYPE_VITAL As Byte = 0
Public Const SPELL_TYPE_WARP As Byte = 1
Public Const SPELL_TYPE_BUFF As Byte = 2

'Buff Types
Public Const BUFF_NONE As Byte = 0
Public Const BUFF_ADD_HP As Byte = 1
Public Const BUFF_ADD_MP As Byte = 2
Public Const BUFF_ADD_STR As Byte = 3
Public Const BUFF_ADD_END As Byte = 4
Public Const BUFF_ADD_AGL As Byte = 5
Public Const BUFF_ADD_INT As Byte = 6
Public Const BUFF_ADD_WIL As Byte = 7
Public Const BUFF_ADD_ATK As Byte = 8
Public Const BUFF_ADD_DEF As Byte = 9
Public Const BUFF_SUB_HP As Byte = 10
Public Const BUFF_SUB_MP As Byte = 11
Public Const BUFF_SUB_STR As Byte = 12
Public Const BUFF_SUB_END As Byte = 13
Public Const BUFF_SUB_AGL As Byte = 14
Public Const BUFF_SUB_INT As Byte = 15
Public Const BUFF_SUB_WIL As Byte = 16
Public Const BUFF_SUB_ATK As Byte = 17
Public Const BUFF_SUB_DEF As Byte = 18

' Game editor constants
Public Const EDITOR_ITEM As Byte = 1
Public Const EDITOR_NPC As Byte = 2
Public Const EDITOR_SPELL As Byte = 3
Public Const EDITOR_SHOP As Byte = 4

' ********************************************
' Default starting location [Server Only]
Public Const START_MAP As Long = 1
Public Const START_X As Long = 30
Public Const START_Y As Long = 10

' Scrolling action message constants
Public Const ACTIONMSG_STATIC As Long = 0
Public Const ACTIONMSG_SCROLL As Long = 1
Public Const ACTIONMSG_SCREEN As Long = 2
