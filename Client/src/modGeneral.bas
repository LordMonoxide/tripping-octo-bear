Attribute VB_Name = "modGeneral"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Used for the 64-bit timer
Private GetSystemTimeOffset As Currency
Private Declare Sub GetSystemTime Lib "kernel32.dll" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As Currency)
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

'For Clear functions
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub Main()
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim i As Long, n As Long

  API.host = "essence.monoxidedesign.com"
  API.port = 80
  
  API.routes.auth.check.route = "api/auth/check"
  API.routes.auth.check.method = HTTP_METHOD_GET
  API.routes.auth.register.route = "api/auth/register"
  API.routes.auth.register.method = HTTP_METHOD_PUT
  API.routes.auth.login.route = "api/auth/login"
  API.routes.auth.login.method = HTTP_METHOD_POST
  API.routes.auth.logout.route = "api/auth/logout"
  API.routes.auth.logout.method = HTTP_METHOD_POST
  API.routes.auth.security.get.route = "api/auth/security"
  API.routes.auth.security.get.method = HTTP_METHOD_GET
  API.routes.auth.security.submit.route = "api/auth/security"
  API.routes.auth.security.submit.method = HTTP_METHOD_POST
  
  API.routes.storage.characters.all.route = "api/storage/characters"
  API.routes.storage.characters.all.method = HTTP_METHOD_GET
  API.routes.storage.characters.create.route = "api/storage/characters"
  API.routes.storage.characters.create.method = HTTP_METHOD_PUT
  API.routes.storage.characters.delete.route = "api/storage/characters"
  API.routes.storage.characters.delete.method = HTTP_METHOD_DELETE
  API.routes.storage.characters.use.route = "api/storage/characters"
  API.routes.storage.characters.use.method = HTTP_METHOD_POST
  
    'Set the high-resolution timer
    timeBeginPeriod 1
    
    'This MUST be called before any timeGetTime calls because it states what the
    'values of timeGetTime will be.
    InitTimeGetTime
    
    ' load gui
    GME = 1
    InitialiseGUI
    
    ' load options
    LoadOptions
    
    ' Check if the directory is there, if its not make it
    ChkDir App.path & "\", "cookies"
    ChkDir App.path & "\data files\", "graphics"
    ChkDir App.path & "\data files\graphics\", "animations"
    ChkDir App.path & "\data files\graphics\", "characters"
    ChkDir App.path & "\data files\graphics\", "items"
    ChkDir App.path & "\data files\graphics\", "resources"
    ChkDir App.path & "\data files\graphics\", "spellicons"
    ChkDir App.path & "\data files\graphics\", "tilesets"
    ChkDir App.path & "\data files\graphics\", "gui"
    ChkDir App.path & "\data files\graphics\gui\", "buttons"
    ChkDir App.path & "\data files\graphics\gui\", "designs"
    ChkDir App.path & "\data files\graphics\", "panoramas"
    ChkDir App.path & "\data files\graphics\", "projectiles"
    ChkDir App.path & "\data files\graphics\", "events"
    ChkDir App.path & "\data files\graphics\", "surfaces"
    ChkDir App.path & "\data files\graphics\", "auras"
    ChkDir App.path & "\data files\graphics\", "misc"
    ChkDir App.path & "\data files\graphics\", "fonts"
    ChkDir App.path & "\data files\graphics\", "socialicons"
    ChkDir App.path & "\data files\", "logs"
    ChkDir App.path & "\data files\", "maps"
    ChkDir App.path & "\data files\", "music"
    ChkDir App.path & "\data files\", "sound"
    
    ' load dx8
    Directx8.init
    LoadSocialicons
    
    ' initialise sound & music engines
    FMOD.init

    ' load the main game (and by extension, pre-load DD7)
    GettingMap = True
    
    ' Update the form with the game's name before it's loaded
    frmMain.Caption = Options.Game_Name
    
    ' randomize rnd's seed
    Randomize
    Call TcpInit
    Call InitMessages

    ' check if we have main-menu music
    If Len(Trim$(Options.MenuMusic)) > 0 Then FMOD.Music_Play Trim$(Options.MenuMusic)
    
    ' Reset values
    Ping = -1
    
    ' set values for directional blocking arrows
    DirArrowX(1) = 12 ' up
    DirArrowY(1) = 0
    DirArrowX(2) = 12 ' down
    DirArrowY(2) = 23
    DirArrowX(3) = 0 ' left
    DirArrowY(3) = 12
    DirArrowX(4) = 23 ' right
    DirArrowY(4) = 12
    
    ' set the main form size
    frmMain.width = 12090
    frmMain.height = 9420
    
    ' show the main menu
    frmMain.Show
  
  Call disableLogin
  Call disableChars
  
  Call showLogin
  Call menuStatus("Checking session...")
  
  Call getChars
  
  For i = 1 To 5
      MenuNPC(i).x = Rand(0, ScreenWidth)
      MenuNPC(i).y = Rand(0, ScreenHeight)
      MenuNPC(i).dir = Rand(0, 1)
  Next
  
  Call MenuLoop
End Sub

Public Sub menuStatus(Optional ByRef Text As String)
  frmMain.lblStatus.Caption = Text
End Sub

Public Sub enableLogin()
  frmMain.fraLogin.Enabled = True
  Call frmMain.txtEmail.SetFocus
End Sub

Public Sub disableLogin()
  frmMain.fraLogin.Enabled = False
End Sub

Public Sub showLogin()
  frmMain.fraLogin.Left = (frmMain.ScaleWidth - frmMain.fraLogin.width) / 2
  frmMain.fraLogin.Top = (frmMain.ScaleHeight - frmMain.fraLogin.height) / 2
  frmMain.fraLogin.visible = True
End Sub

Public Sub hideLogin()
  frmMain.fraLogin.visible = False
End Sub

Public Sub enableLoginSecurity()
  frmMain.fraLoginSecurity.Enabled = True
End Sub

Public Sub disableLoginSecurity()
  frmMain.fraLoginSecurity.Enabled = False
End Sub

Public Sub showLoginSecurity()
  frmMain.fraLoginSecurity.Left = (frmMain.ScaleWidth - frmMain.fraLoginSecurity.width) / 2
  frmMain.fraLoginSecurity.Top = (frmMain.ScaleHeight - frmMain.fraLoginSecurity.height) / 2
  frmMain.fraLoginSecurity.visible = True
End Sub

Public Sub hideLoginSecurity()
  frmMain.fraLoginSecurity.visible = False
End Sub

Public Sub clearLoginError()
  frmMain.lblEmailErr.Caption = vbNullString
  frmMain.lblPasswordErr.Caption = vbNullString
End Sub

Public Sub clearNewCharError()
  frmMain.lblNewCharNameErr.Caption = vbNullString
  frmMain.lblNewCharSexErr.Caption = vbNullString
End Sub

Public Sub enableChars()
  frmMain.fraChars.Enabled = True
End Sub

Public Sub disableChars()
  frmMain.fraChars.Enabled = False
End Sub

Public Sub showChars()
  frmMain.fraChars.Left = (frmMain.ScaleWidth - frmMain.fraChars.width) / 2
  frmMain.fraChars.Top = (frmMain.ScaleHeight - frmMain.fraChars.height) / 2
  frmMain.fraChars.visible = True
End Sub

Public Sub hideChars()
  frmMain.fraChars.visible = False
End Sub

Public Sub enableNewChar()
  frmMain.fraNewChar.Enabled = True
End Sub

Public Sub disableNewChar()
  frmMain.fraNewChar.Enabled = False
End Sub

Public Sub showNewChar()
  frmMain.fraNewChar.Left = (frmMain.ScaleWidth - frmMain.fraNewChar.width) / 2
  frmMain.fraNewChar.Top = (frmMain.ScaleHeight - frmMain.fraNewChar.height) / 2
  frmMain.fraNewChar.visible = True
End Sub

Public Sub hideNewChar()
  frmMain.fraNewChar.visible = False
End Sub

Public Sub parse401(ByVal json As Object)
  Call hideLogin
  Call hideLoginSecurity
  Call hideChars
  Call hideNewChar
  
  Select Case json("show").val
    Case "login"
      Call menuStatus(json("error").val)
      Call showLogin
      Call enableLogin
    
    Case "security"
      Call loginSecurity
  End Select
End Sub

Public Sub login(ByRef email As String, ByRef password As String)
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim pair As clsJSONPair
Dim json As Object
Dim o As Object

  Call menuStatus("Logging in...")
  Call disableLogin
  Call clearLoginError
  
  Set request = New clsHTTPRequest
  request.method = API.routes.auth.login.method
  request.route = API.routes.auth.login.route
  Call request.addHeader("Accept", "application/json")
  Call request.addData("email", email)
  Call request.addData("password", password)
  Set response = request.dispatch
  Call response.await
  
  Select Case response.responseCode
    Case 200
      Call menuStatus
      Call hideLogin
      Call getChars
    
    Case 401
      Call parse401(response.responseJSON)
    
    Case 409
      Set json = response.responseJSON
      
      For Each pair In json
        Select Case pair.key
          Case "email":    frmMain.lblEmailErr.Caption = pair.val(1).val
          Case "password": frmMain.lblPasswordErr.Caption = pair.val(1).val
        End Select
      Next
      
      Call menuStatus
      Call enableLogin
    
    Case Else
      Call MsgBox("This shouldn't happen:" & vbNewLine & response.responseBody)
  End Select
End Sub

Public Sub loginSecurity()
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim pair As clsJSONPair
Dim i As Long

  Call disableLoginSecurity
  Call showLoginSecurity
  
  Set request = New clsHTTPRequest
  request.method = API.routes.auth.security.get.method
  request.route = API.routes.auth.security.get.route
  Call request.addHeader("Accept", "application/json")
  Set response = request.dispatch
  Call response.await
  
  For i = 1 To frmMain.lblLoginSecurityQuestion.count - 1
    Call Unload(frmMain.lblLoginSecurityQuestion(i))
    Call Unload(frmMain.txtLoginSecurityAnswer(i))
  Next
  
  Select Case response.responseCode
    Case 200
      For i = 1 To response.responseJSON.count - 1
        Call Load(frmMain.lblLoginSecurityQuestion(i))
        Call Load(frmMain.txtLoginSecurityAnswer(i))
        
        frmMain.lblLoginSecurityQuestion(i).Left = frmMain.lblLoginSecurityQuestion(i - 1).Left
        frmMain.lblLoginSecurityQuestion(i).Top = frmMain.txtLoginSecurityAnswer(i - 1).Top + frmMain.txtLoginSecurityAnswer(i - 1).height + 4
        frmMain.lblLoginSecurityQuestion(i).visible = True
        frmMain.txtLoginSecurityAnswer(i).Left = frmMain.txtLoginSecurityAnswer(i - 1).Left
        frmMain.txtLoginSecurityAnswer(i).Top = frmMain.lblLoginSecurityQuestion(i).Top + frmMain.lblLoginSecurityQuestion(i).height
        frmMain.txtLoginSecurityAnswer(i).visible = True
      Next
      
      i = 0
      For Each pair In response.responseJSON
        frmMain.lblLoginSecurityQuestion(i).Caption = pair.val("question").val
        i = i + 1
      Next
      
      frmMain.picLoginSecurity.height = frmMain.txtLoginSecurityAnswer(frmMain.txtLoginSecurityAnswer.UBound).Top + frmMain.txtLoginSecurityAnswer(frmMain.txtLoginSecurityAnswer.UBound).height + 4
      frmMain.scrlLoginSecurity.Max = frmMain.picLoginSecurity.height - frmMain.picLoginSecurityCont.ScaleHeight
      
      Call enableLoginSecurity
    
    Case 409
      Call MsgBox(response.responseBody)
    
    Case Else
      Call MsgBox("This shouldn't happen:" & vbNewLine & response.responseBody)
  End Select
End Sub

Public Sub loginSecuritySubmit(ByRef answer() As String)
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim i As Long

  Call disableLoginSecurity
  
  Set request = New clsHTTPRequest
  request.method = API.routes.auth.security.submit.method
  request.route = API.routes.auth.security.submit.route
  Call request.addHeader("Accept", "application/json")
  
  For i = 0 To aryLenS(answer) - 1
    Call request.addData("answer" & i, answer(i))
  Next
  
  Set response = request.dispatch
  Call response.await
  
  Select Case response.responseCode
    Case 200
      Call hideLoginSecurity
      Call getChars
    
    Case 409
      Call MsgBox(response.responseBody)
      'Call MsgBox(response.responseJSON("error").val)
  End Select
  
  Call enableLoginSecurity
End Sub

Public Sub logout()
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim pair As clsJSONPair
Dim json As Object

  Call disableChars
  
  Set request = New clsHTTPRequest
  request.method = API.routes.auth.logout.method
  request.route = API.routes.auth.logout.route
  Call request.addHeader("Accept", "application/json")
  Set response = request.dispatch
  Call response.await
  
  Select Case response.responseCode
    Case 200
      Call hideChars
      Call showLogin
      Call enableLogin
    
    Case 409
      Set json = response.responseJSON
      
      For Each pair In json
        Call MsgBox(pair.key & ": " & pair.val(1).val)
      Next
      
      Call enableChars
    
    Case Else
      Call MsgBox("This shouldn't happen:" & vbNewLine & response.responseBody)
  End Select
End Sub

Public Sub getChars()
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim pair As clsJSONPair

  Call showChars
  Call disableChars
  
  Set request = New clsHTTPRequest
  request.method = API.routes.storage.characters.all.method
  request.route = API.routes.storage.characters.all.route
  Call request.addHeader("Accept", "application/json")
  Set response = request.dispatch
  Call response.await
  
  Call frmMain.lstChars.Clear
  
  Select Case response.responseCode
    Case 200
      For Each pair In response.responseJSON
        Call frmMain.lstChars.AddItem(pair.val("name").val & ", level " & pair.val("lvl").val & " " & pair.val("sex").val)
        frmMain.lstChars.ItemData(frmMain.lstChars.ListCount - 1) = pair.val("id").val
      Next
      
      If frmMain.lstChars.ListCount <> 0 Then
        frmMain.lstChars.ListIndex = 0
      End If
    
    Case 401
      Call parse401(response.responseJSON)
    
    Case 409
      Call hideChars
      Call showLogin
      Call menuStatus(response.responseJSON(1).val(1).val)
      Call enableLogin
    
    Case Else
      Call MsgBox("This shouldn't happen:" & vbNewLine & response.responseBody)
  End Select
  
  Call enableChars
End Sub

Public Sub delChar(ByVal id As Long)
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse

  Call disableChars
  
  Set request = New clsHTTPRequest
  request.method = API.routes.storage.characters.delete.method
  request.route = API.routes.storage.characters.delete.route
  Call request.addHeader("Accept", "application/json")
  Call request.addData("id", str$(id))
  Set response = request.dispatch
  Call response.await
  
  Select Case response.responseCode
    Case 200
      Call getChars
    
    Case 401
      Call parse401(response.responseJSON)
    
    Case 409
      MsgBox response.responseBody
    
    Case Else
      MsgBox response.responseBody
  End Select
  
  Call enableChars
End Sub

Public Sub newChar(ByRef name As String, ByRef sex As String)
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim pair As clsJSONPair
Dim json As Object

  Call disableNewChar
  Call clearNewCharError
  
  Set request = New clsHTTPRequest
  request.method = API.routes.storage.characters.create.method
  request.route = API.routes.storage.characters.create.route
  Call request.addHeader("Accept", "application/json")
  Call request.addData("name", name)
  Call request.addData("sex", sex)
  Set response = request.dispatch
  Call response.await
  
  Select Case response.responseCode
    Case 201
      Call hideNewChar
      Call getChars
    
    Case 401
      Call parse401(response.responseJSON)
    
    Case 409
      Set json = response.responseJSON
      
      For Each pair In json
        Select Case pair.key
          Case "name": frmMain.lblNewCharNameErr.Caption = pair.val(1).val
          Case "sex":  frmMain.lblNewCharSexErr.Caption = pair.val(1).val
        End Select
      Next
    
    Case Else
      Call MsgBox("This shouldn't happen:" & vbNewLine & response.responseBody)
  End Select
  
  Call enableNewChar
End Sub

Public Sub useChar(ByVal id As Long)
Dim request As clsHTTPRequest
Dim response As clsHTTPResponse
Dim buffer As clsBuffer

  Call disableChars
  
  Set request = New clsHTTPRequest
  request.method = API.routes.storage.characters.use.method
  request.route = API.routes.storage.characters.use.route
  Call request.addHeader("Accept", "application/json")
  Call request.addData("id", str$(id))
  Set response = request.dispatch
  Call response.await
  
  Select Case response.responseCode
    Case 200
      If connectToServer Then
        Set buffer = New clsBuffer
        Call sendLogin(response.responseJSON("u_id").val, response.responseJSON("c_id").val)
        Exit Sub
      Else
        Call MsgBox("Server down")
      End If
    
    Case 401
      Call parse401(response.responseJSON)
    
    Case 409
      Call MsgBox(response.responseBody)
  End Select
  
  Call enableChars
End Sub

Public Sub InitialiseGUI()

'Loading Interface.ini data
Dim FileName As String
FileName = App.path & "\data files\interface.ini"
Dim i As Long

    ' re-set chat scroll
    ChatScroll = 8

    ReDim GUIWindow(1 To GUI_Count - 1) As GUIWindowRec
    
    ' 1 - Chat
    With GUIWindow(GUI_CHAT)
        .x = val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = True
    End With
    
    ' 2 - Hotbar
    With GUIWindow(GUI_HOTBAR)
        .x = val(GetVar(FileName, "GUI_HOTBAR", "X"))
        .y = val(GetVar(FileName, "GUI_HOTBAR", "Y"))
        .height = val(GetVar(FileName, "GUI_HOTBAR", "Height"))
        .width = ((9 + 36) * (MAX_HOTBAR - 1))
    End With
    
    ' 3 - Menu
    With GUIWindow(GUI_MENU)
        .x = val(GetVar(FileName, "GUI_MENU", "X"))
        .y = val(GetVar(FileName, "GUI_MENU", "Y"))
        .width = val(GetVar(FileName, "GUI_MENU", "Width"))
        .height = val(GetVar(FileName, "GUI_MENU", "Height"))
        .visible = True
    End With
    
    ' 4 - Bars
    With GUIWindow(GUI_BARS)
        .x = val(GetVar(FileName, "GUI_BARS", "X"))
        .y = val(GetVar(FileName, "GUI_BARS", "Y"))
        .width = val(GetVar(FileName, "GUI_BARS", "Width"))
        .height = val(GetVar(FileName, "GUI_BARS", "Height"))
        .visible = True
    End With
    
    ' 5 - Inventory
    With GUIWindow(GUI_INVENTORY)
        .x = val(GetVar(FileName, "GUI_INVENTORY", "X"))
        .y = val(GetVar(FileName, "GUI_INVENTORY", "Y"))
        .width = val(GetVar(FileName, "GUI_INVENTORY", "Width"))
        .height = val(GetVar(FileName, "GUI_INVENTORY", "Height"))
        .visible = False
    End With
    
    ' 6 - Spells
    With GUIWindow(GUI_SPELLS)
        .x = val(GetVar(FileName, "GUI_SPELLS", "X"))
        .y = val(GetVar(FileName, "GUI_SPELLS", "Y"))
        .width = val(GetVar(FileName, "GUI_SPELLS", "Width"))
        .height = val(GetVar(FileName, "GUI_SPELLS", "Height"))
        .visible = False
    End With
    
    ' 7 - Character
    With GUIWindow(GUI_CHARACTER)
        .x = val(GetVar(FileName, "GUI_CHARACTER", "X"))
        .y = val(GetVar(FileName, "GUI_CHARACTER", "Y"))
        .width = val(GetVar(FileName, "GUI_CHARACTER", "Width"))
        .height = val(GetVar(FileName, "GUI_CHARACTER", "Height"))
        .visible = False
    End With
    
    ' 8 - Options
    With GUIWindow(GUI_OPTIONS)
        .x = val(GetVar(FileName, "GUI_OPTIONS", "X"))
        .y = val(GetVar(FileName, "GUI_OPTIONS", "Y"))
        .width = val(GetVar(FileName, "GUI_OPTIONS", "Width"))
        .height = val(GetVar(FileName, "GUI_OPTIONS", "Height"))
        .visible = False
    End With
    With GUIWindow(GUI_QUESTDIALOGUE)
        .x = val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With

    ' 9 - Party
    With GUIWindow(GUI_PARTY)
        .x = val(GetVar(FileName, "GUI_PARTY", "X"))
        .y = val(GetVar(FileName, "GUI_PARTY", "Y"))
        .width = val(GetVar(FileName, "GUI_PARTY", "Width"))
        .height = val(GetVar(FileName, "GUI_PARTY", "Height"))
        .visible = False
    End With
    
    ' 10 - Description
    With GUIWindow(GUI_DESCRIPTION)
        .x = val(GetVar(FileName, "GUI_DESCRIPTION", "X"))
        .y = val(GetVar(FileName, "GUI_DESCRIPTION", "Y"))
        .width = val(GetVar(FileName, "GUI_DESCRIPTION", "Width"))
        .height = val(GetVar(FileName, "GUI_DESCRIPTION", "Height"))
        .visible = False
    End With
    
        With GUIWindow(GUI_QUESTS)
        .x = 120
        .y = 140
        .width = 600
        .height = 307
        .visible = False
    End With
    
    ' 11 - Main Menu
    With GUIWindow(GUI_MAINMENU)
        .x = val(GetVar(FileName, "GUI_MAINMENU", "X"))
        .y = val(GetVar(FileName, "GUI_MAINMENU", "Y"))
        .width = val(GetVar(FileName, "GUI_MAINMENU", "Width"))
        .height = val(GetVar(FileName, "GUI_MAINMENU", "Height"))
        .visible = False
    End With
    
    ' 12 - Shop
    With GUIWindow(GUI_SHOP)
         .x = val(GetVar(FileName, "GUI_SHOP", "X"))
        .y = val(GetVar(FileName, "GUI_SHOP", "Y"))
        .width = val(GetVar(FileName, "GUI_SHOP", "Width"))
        .height = val(GetVar(FileName, "GUI_SHOP", "Height"))
        .visible = False
    End With
    
    ' 13 - Bank
    With GUIWindow(GUI_BANK)
        .x = val(GetVar(FileName, "GUI_BANK", "X"))
        .y = val(GetVar(FileName, "GUI_BANK", "Y"))
        .width = val(GetVar(FileName, "GUI_BANK", "Width"))
        .height = val(GetVar(FileName, "GUI_BANK", "Height"))
        .visible = False
    End With
    
    ' 14 - Trade
    With GUIWindow(GUI_TRADE)
        .x = val(GetVar(FileName, "GUI_TRADE", "X"))
        .y = val(GetVar(FileName, "GUI_TRADE", "Y"))
        .width = val(GetVar(FileName, "GUI_TRADE", "Width"))
        .height = val(GetVar(FileName, "GUI_TRADE", "Height"))
        .visible = False
    End With
    
    ' 15 - Currency
    With GUIWindow(GUI_CURRENCY)
        .x = val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 16 - Dialogue
    With GUIWindow(GUI_DIALOGUE)
        .x = val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 17 - Event Chat
    With GUIWindow(GUI_EVENTCHAT)
        .x = val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 18 - Tutorial
    With GUIWindow(GUI_TUTORIAL)
        .x = val(GetVar(FileName, "GUI_CHAT", "X"))
        .y = val(GetVar(FileName, "GUI_CHAT", "Y"))
        .width = val(GetVar(FileName, "GUI_CHAT", "Width"))
        .height = val(GetVar(FileName, "GUI_CHAT", "Height"))
        .visible = False
    End With
    
    ' 19 - Right-Click menu
    With GUIWindow(GUI_RIGHTMENU)
        .x = 0
        .y = 0
        .width = 110
        .height = 145
        .visible = False
    End With
    
    ' 20 - Guild Window
    With GUIWindow(GUI_GUILD)
        .x = val(GetVar(FileName, "GUI_GUILD", "X"))
        .y = val(GetVar(FileName, "GUI_GUILD", "Y"))
        .width = val(GetVar(FileName, "GUI_GUILD", "Width"))
        .height = val(GetVar(FileName, "GUI_GUILD", "Height"))
        .visible = False
    End With
    
    ' BUTTONS
    ' main - inv
    With Buttons(1)
        .state = 0 ' normal
        .x = 6
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 1
    End With
    
    ' main - skills
    With Buttons(2)
        .state = 0 ' normal
        .x = 41
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 2
    End With
    
    ' main - char
    With Buttons(3)
        .state = 0 ' normal
        .x = 76
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 3
    End With
    
    ' main - opt
    With Buttons(4)
        .state = 0 ' normal
        .x = 111
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 4
    End With
    
    ' main - trade
    With Buttons(5)
        .state = 0 ' normal
        .x = 146
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 5
    End With
    
    ' main - party
    With Buttons(6)
        .state = 0 ' normal
        .x = 181
        .y = 6
        .width = 36
        .height = 36
        .visible = True
        .PicNum = 6
    End With
    
    ' menu - login
    With Buttons(7)
        .state = 0 ' normal
        .x = 54
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 7
    End With
    
    ' menu - register
    With Buttons(8)
        .state = 0 ' normal
        .x = 154
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 8
    End With
    
    ' menu - credits
    With Buttons(9)
        .state = 0 ' normal
        .x = 254
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 9
    End With
    
    ' menu - exit
    With Buttons(10)
        .state = 0 ' normal
        .x = 354
        .y = 277
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 10
    End With
    
    ' menu - Login Accept
    With Buttons(11)
        .state = 0 ' normal
        .x = 206
        .y = 164
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Register Accept
    With Buttons(12)
        .state = 0 ' normal
        .x = 206
        .y = 169
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Accept
    With Buttons(13)
        .state = 0 ' normal
        .x = 248
        .y = 206
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' menu - Class Next
    With Buttons(14)
        .state = 0 ' normal
        .x = 348
        .y = 206
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 12
    End With
    
    ' menu - NewChar Accept
    With Buttons(15)
        .state = 0 ' normal
        .x = 205
        .y = 169
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' main - AddStats
    For i = 16 To 20
        With Buttons(i)
            .state = 0 'normal
            .width = 12
            .height = 11
            .visible = True
            .PicNum = 13
        End With
    Next
    ' set the individual spaces
    For i = 16 To 18 ' first 3
        With Buttons(i)
            .x = 80
            .y = 22 + ((i - 16) * 15)
        End With
    Next
    For i = 19 To 20
        With Buttons(i)
            .x = 165
            .y = 22 + ((i - 19) * 15)
        End With
    Next
    
    ' main - shop buy
    With Buttons(21)
        .state = 0 ' normal
        .x = 12
        .y = 276
        .width = 69
        .height = 29
        .visible = True
        .PicNum = 14
    End With
    
    ' main - shop sell
    With Buttons(22)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .width = 69
        .height = 29
        .visible = True
        .PicNum = 15
    End With
    
    ' main - shop exit
    With Buttons(23)
        .state = 0 ' normal
        .x = 90
        .y = 276
        .width = 69
        .height = 29
        .visible = True
        .PicNum = 16
    End With
    
    ' main - party invite
    With Buttons(24)
        .state = 0 ' normal
        .x = 14
        .y = 209
        .width = 79
        .height = 29
        .visible = True
        .PicNum = 17
    End With
    
    ' main - party invite
    With Buttons(25)
        .state = 0 ' normal
        .x = 101
        .y = 209
        .width = 79
        .height = 29
        .visible = True
        .PicNum = 18
    End With
    
    ' main - music on
    With Buttons(26)
        .state = 0 ' normal
        .x = 77
        .y = 14
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - music off
    With Buttons(27)
        .state = 0 ' normal
        .x = 132
        .y = 14
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - sound on
    With Buttons(28)
        .state = 0 ' normal
        .x = 77
        .y = 39
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - sound off
    With Buttons(29)
        .state = 0 ' normal
        .x = 132
        .y = 39
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - debug on
    With Buttons(30)
        .state = 0 ' normal
        .x = 77
        .y = 64
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - debug off
    With Buttons(31)
        .state = 0 ' normal
        .x = 132
        .y = 64
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - autotile on
    With Buttons(32)
        .state = 0 ' normal
        .x = 77
        .y = 89
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 19
    End With
    
    ' main - autotile off
    With Buttons(33)
        .state = 0 ' normal
        .x = 132
        .y = 89
        .width = 49
        .height = 19
        .visible = True
        .PicNum = 20
    End With
    
    ' main - scroll up
    With Buttons(34)
        .state = 0 ' normal
        .x = 340
        .y = 2
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 21
    End With
    
    ' main - scroll down
    With Buttons(35)
        .state = 0 ' normal
        .x = 340
        .y = 100
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 22
    End With
    
    ' main - Accept Trade
    With Buttons(36)
        .state = 0 'normal
        .x = GUIWindow(GUI_TRADE).x + 125
        .y = GUIWindow(GUI_TRADE).y + 335
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 11
    End With
    
    ' main - Decline Trade
    With Buttons(37)
        .state = 0 'normal
        .x = GUIWindow(GUI_TRADE).x + 265
        .y = GUIWindow(GUI_TRADE).y + 335
        .width = 89
        .height = 29
        .visible = True
        .PicNum = 10
    End With
    ' main - FPS Cap left
    With Buttons(38)
        .state = 0 'normal
        .x = 92
        .y = 112
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 23
    End With
    ' main - FPS Cap Right
    With Buttons(39)
        .state = 0 'normal
        .x = 147
        .y = 112
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 24
    End With
    ' main - Volume left
    With Buttons(40)
        .state = 0 'normal
        .x = 92
        .y = 132
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 23
    End With
    ' main - Volume Right
    With Buttons(41)
        .state = 0 'normal
        .x = 147
        .y = 132
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 24
    End With
     ' main - guild Up
    With Buttons(42)
        .state = 0 ' normal
        .x = 155
        .y = 119
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 21
    End With
    
    ' main - guild down
    With Buttons(43)
        .state = 0 ' normal
        .x = 155
        .y = 189
        .width = 19
        .height = 19
        .visible = True
        .PicNum = 22
    End With
End Sub

Public Sub logoutGame()
Dim i As Long

    isLogging = True
    InGame = False
    
    Call DestroyTCP
    
    ' destroy the animations loaded
    For i = 1 To MAX_BYTE
        ClearAnimInstance (i)
    Next
    
    ' destroy temp values
    DragInvSlotNum = 0
    MyIndex = 0
    InventoryItemSelected = 0
    SpellBuffer = 0
    SpellBufferTimer = 0
    tmpCurrencyItem = 0
    
    ' unload editors
    Unload frmEditor_Animation
    Unload frmEditor_Item
    Unload frmEditor_Map
    Unload frmEditor_MapProperties
    Unload frmEditor_NPC
    Unload frmEditor_Resource
    Unload frmEditor_Shop
    Unload frmEditor_Spell
    
    ' destroy the chat
    For i = 1 To ChatTextBufferSize
        ChatTextBuffer(i).Text = vbNullString
    Next
    
    GUIWindow(GUI_MAINMENU).visible = True
    inMenu = True
    ' Load the username + pass
    sUser = Trim$(Options.Username)
    If Options.savePass = 1 Then
        sPass = Trim$(Options.password)
    End If
    curTextbox = 1
    curMenu = MENU_LOGIN
    HideGame
    MenuLoop
End Sub

Sub GameInit()
Dim MusicFile As String

    ' hide gui
    InBank = False
    InShop = False
    InTrade = False
    
    ' get ping
    GetPing
    
    ' set values for amdin panel
    frmMain.scrlAItem.Max = MAX_ITEMS
    frmMain.scrlAItem.Value = 1
    
    GUIWindow(GUI_OPTIONS).visible = False
    
    ' play music
    MusicFile = Trim$(map.Music)
    If Not MusicFile = "None." Then
        FMOD.Music_Play MusicFile
    Else
        FMOD.Music_Stop
    End If
End Sub

Public Sub DestroyGame()
    ' break out of GameLoop
    HideGame
    Call DestroyTCP
    
    ' destroy music & sound engines
    FMOD.Destroy
    
    ' unload dx8
    Directx8.Destroy
    
    Call UnloadAllForms
    End
End Sub

Public Sub UnloadAllForms()
Dim frm As Form

    For Each frm In VB.Forms
        Unload frm
    Next
End Sub

' Used for adding text to packet debugger
Public Sub TextAdd(ByVal Txt As TextBox, msg As String, NewLine As Boolean)
    If NewLine Then
        Txt.Text = Txt.Text + msg + vbCrLf
    Else
        Txt.Text = Txt.Text + msg
    End If

    Txt.SelStart = Len(Txt.Text) - 1
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function isLoginLegal(ByVal Username As String, ByVal password As String) As Boolean
    If LenB(Trim$(Username)) >= 3 Then
        If LenB(Trim$(password)) >= 3 Then
            isLoginLegal = True
        End If
    End If
End Function

Public Function isStringLegal(ByVal sInput As String) As Boolean
Dim i As Long

    ' Prevent high ascii chars
    For i = 1 To Len(sInput)

        If Asc(Mid$(sInput, i, 1)) < vbKeySpace Or Asc(Mid$(sInput, i, 1)) > vbKeyF15 Then
            Call MsgBox("You cannot use high ASCII characters in your name, please re-enter.", vbOKOnly, Options.Game_Name)
            Exit Function
        End If

    Next

    isStringLegal = True
End Function

Public Sub resetClickedButtons()
Dim i As Long

    ' loop through entire array
    For i = 1 To MAX_BUTTONS
        Select Case i
            ' option buttons
            Case 26, 27, 28, 29, 30, 31, 32, 33
                ' Nothing in here
            ' Everything else - reset
            Case Else
                ' reset state and render
                Buttons(i).state = 0 'normal
        End Select
    Next
End Sub

Public Sub PopulateLists()
Dim strLoad As String, i As Long

    ' Cache music list
    strLoad = dir(App.path & MUSIC_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve musicCache(1 To i) As String
        musicCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
    
    ' Cache sound list
    strLoad = dir(App.path & SOUND_PATH & "*.*")
    i = 1
    Do While strLoad > vbNullString
        ReDim Preserve soundCache(1 To i) As String
        soundCache(i) = strLoad
        strLoad = dir
        i = i + 1
    Loop
End Sub

Public Sub ShowGame()
Dim i As Long

    For i = 5 To 10
        GUIWindow(i).visible = False
    Next

    For i = 1 To 4
        GUIWindow(i).visible = True
    Next
    
    InGame = True
End Sub

Public Sub HideGame()
Dim i As Long
    
    For i = 1 To 10
        GUIWindow(i).visible = False
    Next
    
    InGame = False
End Sub

Public Sub InitTimeGetTime()
'*****************************************************************
'Gets the offset time for the timer so we can start at 0 instead of
'the returned system time, allowing us to not have a time roll-over until
'the program is running for 25 days
'*****************************************************************

    'Get the initial time
    GetSystemTime GetSystemTimeOffset

End Sub

Public Function timeGetTime() As Long
'*****************************************************************
'Grabs the time from the 64-bit system timer and returns it in 32-bit
'after calculating it with the offset - allows us to have the
'"no roll-over" advantage of 64-bit timers with the RAM usage of 32-bit
'though we limit things slightly, so the rollover still happens, but after 25 days
'*****************************************************************
Dim CurrentTime As Currency

    'Grab the current time (we have to pass a variable ByRef instead of a function return like the other timers)
    GetSystemTime CurrentTime
    
    'Calculate the difference between the 64-bit times, return as a 32-bit time
    timeGetTime = CurrentTime - GetSystemTimeOffset

End Function

Public Function KeepTwoDigit(Num As Byte)
    If (Num < 10) Then
        KeepTwoDigit = "0" & Num
    Else
        KeepTwoDigit = Num
    End If
End Function
