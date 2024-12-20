Attribute VB_Name = "Protocol"
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Public Const SEPARATOR             As String * 1 = vbNullChar

Public Enum e_EditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Sex
    eo_Raza
    eo_Arma
    eo_Escudo
    eo_CASCO
    eo_Particula
    eo_Vida
    eo_Mana
    eo_Energia
    eo_MinHP
    eo_MinMP
    eo_Hit
    eo_MinHit
    eo_MaxHit
    eo_Desc
    eo_Intervalo
    eo_Hogar

End Enum

Public Enum e_FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_DIOS
    FONTTYPE_CITIZEN
    FONTTYPE_CITIZEN_ARMADA
    FONTTYPE_CRIMINAL
    FONTTYPE_CRIMINAL_CAOS
    FONTTYPE_EXP
    FONTTYPE_SUBASTA
    FONTTYPE_GLOBAL
    FONTTYPE_MP
    FONTTYPE_ROSA
    FONTTYPE_VIOLETA
    FONTTYPE_INFOIAO
    
    FONTTYPE_New_Amarillo_Oscuro
    FONTTYPE_New_Verde_Oscuro
    FONTTYPE_New_Naranja
    FONTTYPE_New_Celeste
    FONTTYPE_New_Amarillo_Verdoso
    FONTTYPE_New_Gris
    FONTTYPE_New_Blanco
    FONTTYPE_New_Rojo_Salmon
    FONTTYPE_New_DONADOR
    FONTTYPE_New_GRUPO
    FONTTYPE_New_Eventos
    
    FONTTYPE_PROMEDIO_IGUAL
    FONTTYPE_PROMEDIO_MENOR
    FONTTYPE_PROMEDIO_MAYOR
    
End Enum

Public Type t_PersonajeCuenta

    nombre As String
    nivel As Byte
    Mapa As Integer
    posX As Integer
    posY As Integer
    cuerpo As Integer
    Cabeza As Integer
    Status As Byte
    clase As Byte
    Arma As Integer
    Escudo As Integer
    Casco As Integer
    ClanIndex As Integer

End Type

Public Reader  As Network.Reader

Public Sub InitializePacketList()
    Call Protocol_Writes.InitializeAuxiliaryBuffer
End Sub

''
' Handles incoming data.
'
' @param    UserIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer, ByVal ConnectionId As Long, ByVal Message As Network.Reader) As Boolean

On Error Resume Next
    Set Reader = Message
    
    Dim PacketId As Long
    PacketId = Reader.ReadInt16
    
#If STRESSER = 1 Then
    Debug.Print "Paquete: " & PacketID
#End If

    
    Dim actual_time As Long
    Dim performance_timer As Long
    actual_time = GetTickCount()
    performance_timer = actual_time
    
    If actual_time - Mapping(ConnectionId).TimeLastReset >= 5000 Then
        Mapping(ConnectionId).TimeLastReset = actual_time
        Mapping(ConnectionId).PacketCount = 0
    End If
    
    If PacketId <> ClientPacketID.eSendPosSeguimiento Then
      '  Debug.Print PacketId
        Mapping(ConnectionId).PacketCount = Mapping(ConnectionId).PacketCount + 1
    End If
    
    If Mapping(ConnectionId).PacketCount > 100 Then
        'Lo kickeo
        If UserIndex > 0 Then
            Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Control Paquetes---> El usuario " & UserList(UserIndex).name & " | Iteración paquetes | Último paquete: " & PacketId & ".", e_FontTypeNames.FONTTYPE_FIGHT))
            Mapping(ConnectionId).PacketCount = 0
            If IsFeatureEnabled("kick_packet_overflow") Then
                Call CloseSocket(UserIndex)
            End If
        Else
            Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Control Paquetes---> Usuario desconocido | Iteración paquetes | Último paquete: " & PacketId & ".", e_FontTypeNames.FONTTYPE_FIGHT))
            Mapping(ConnectionId).PacketCount = 0
            If IsFeatureEnabled("kick_packet_overflow") Then
                Call KickConnection(ConnectionId)
            End If
        End If
        Exit Function
    End If

    If PacketId < ClientPacketID.eMinPacket Or PacketId >= ClientPacketID.PacketCount Then
        Call LogEdicionPaquete("El usuario " & UserList(UserIndex).ConnectionDetails.IP & " mando fake paquet " & PacketId)
        Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Control Paquetes---> El usuario " & UserList(UserIndex).name & " | IP: " & UserList(UserIndex).ConnectionDetails.IP & " ESTÁ ENVIANDO PAQUETES INVÁLIDOS", e_FontTypeNames.FONTTYPE_GUILD))
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    #If PYMMO = 1 Then
        'Does the packet requires a logged user??
        If Not (PacketId = ClientPacketID.eLoginExistingChar Or _
                PacketId = ClientPacketID.eLoginNewChar) Then
                   
            'Is the user actually logged?
            If Not UserList(userindex).flags.UserLogged Then
                Call CloseSocket(userindex)
                Exit Function
            
            'He is logged. Reset idle counter if id is valid.
            ElseIf PacketId <= ClientPacketID.[PacketCount] Then
                UserList(userindex).Counters.IdleCount = 0
            End If
        End If
    #ElseIf PYMMO = 0 Then
        'Does the packet requires a logged account??
        If Not (PacketId = ClientPacketID.eCreateAccount Or _
                PacketId = ClientPacketID.eLoginAccount) Then
                   
            'Is the account actually logged?
            If UserList(UserIndex).AccountID = 0 Then
                Call CloseSocket(UserIndex)
                Exit Function
            End If
            
            If Not (PacketId = ClientPacketID.eLoginExistingChar Or PacketId = ClientPacketID.eLoginNewChar) Then
                       
                'Is the user actually logged?
                If Not UserList(UserIndex).flags.UserLogged Then
                    Call CloseSocket(UserIndex)
                    Exit Function
                
                'He is logged. Reset idle counter if id is valid.
                ElseIf PacketId <= ClientPacketID.[PacketCount] Then
                    UserList(UserIndex).Counters.IdleCount = 0
                End If
            End If
        End If
    #End If
    
    Select Case PacketID
        Case ClientPacketID.eLoginExistingChar
            Call HandleLoginExistingChar(ConnectionId)
        Case ClientPacketID.eLoginNewChar
            Call HandleLoginNewChar(ConnectionId)
        Case ClientPacketID.eWalk
            Call HandleWalk(UserIndex)
        Case ClientPacketID.eAttack
            Call HandleAttack(UserIndex)
        Case ClientPacketID.eTalk
            Call HandleTalk(UserIndex)
        Case ClientPacketID.eYell
            Call HandleYell(UserIndex)
        Case ClientPacketID.eWhisper
            Call HandleWhisper(UserIndex)
        Case ClientPacketID.eRequestPositionUpdate
            Call HandleRequestPositionUpdate(UserIndex)
        Case ClientPacketID.ePickUp
            Call HandlePickUp(UserIndex)
        Case ClientPacketID.eSafeToggle
            Call HandleSafeToggle(UserIndex)
        Case ClientPacketID.ePartySafeToggle
            Call HandlePartyToggle(UserIndex)
        Case ClientPacketID.eRequestGuildLeaderInfo
            Call HandleRequestGuildLeaderInfo(UserIndex)
        Case ClientPacketID.eRequestAtributes
            Call HandleRequestAtributes(UserIndex)
        Case ClientPacketID.eRequestSkills
            Call HandleRequestSkills(UserIndex)
        Case ClientPacketID.eRequestMiniStats
            Call HandleRequestMiniStats(UserIndex)
        Case ClientPacketID.eCommerceEnd
            Call HandleCommerceEnd(UserIndex)
        Case ClientPacketID.eUserCommerceEnd
            Call HandleUserCommerceEnd(UserIndex)
        Case ClientPacketID.eBankEnd
            Call HandleBankEnd(UserIndex)
        Case ClientPacketID.eUserCommerceOk
            Call HandleUserCommerceOk(UserIndex)
        Case ClientPacketID.eUserCommerceReject
            Call HandleUserCommerceReject(UserIndex)
        Case ClientPacketID.eDrop
            Call HandleDrop(UserIndex)
        Case ClientPacketID.eCastSpell
            Call HandleCastSpell(UserIndex) ', crc)
        Case ClientPacketID.eLeftClick
            Call HandleLeftClick(UserIndex)
        Case ClientPacketID.eDoubleClick
            Call HandleDoubleClick(UserIndex)
        Case ClientPacketID.eWork
            Call HandleWork(UserIndex)
        Case ClientPacketID.eUseSpellMacro
            Call HandleUseSpellMacro(UserIndex)
        Case ClientPacketID.eUseItem
            Call HandleUseItem(UserIndex)
        Case ClientPacketID.eUseItemU
            Call HandleUseItemU(UserIndex)
        Case ClientPacketID.eCraftBlacksmith
            Call HandleCraftBlacksmith(UserIndex)
        Case ClientPacketID.eCraftCarpenter
            Call HandleCraftCarpenter(UserIndex)
        Case ClientPacketID.eWorkLeftClick
            Call HandleWorkLeftClick(UserIndex)
        Case ClientPacketID.eCreateNewGuild
            Call HandleCreateNewGuild(UserIndex)
        Case ClientPacketID.eSpellInfo
            Call HandleSpellInfo(UserIndex)
        Case ClientPacketID.eEquipItem
            Call HandleEquipItem(UserIndex)
        Case ClientPacketID.eChangeHeading
            Call HandleChange_Heading(UserIndex)
        Case ClientPacketID.eModifySkills
            Call HandleModifySkills(UserIndex)
        Case ClientPacketID.eTrain
            Call HandleTrain(UserIndex)
        Case ClientPacketID.eCommerceBuy
            Call HandleCommerceBuy(UserIndex)
        Case ClientPacketID.eBankExtractItem
            Call HandleBankExtractItem(UserIndex)
        Case ClientPacketID.eCommerceSell
            Call HandleCommerceSell(UserIndex)
        Case ClientPacketID.eBankDeposit
            Call HandleBankDeposit(UserIndex)
        Case ClientPacketID.eForumPost
            Call HandleForumPost(UserIndex)
        Case ClientPacketID.eMoveSpell
            Call HandleMoveSpell(UserIndex)
        Case ClientPacketID.eClanCodexUpdate
            Call HandleClanCodexUpdate(UserIndex)
        Case ClientPacketID.eUserCommerceOffer
            Call HandleUserCommerceOffer(UserIndex)
        Case ClientPacketID.eGuildAcceptPeace
            Call HandleGuildAcceptPeace(UserIndex)
        Case ClientPacketID.eGuildRejectAlliance
            Call HandleGuildRejectAlliance(UserIndex)
        Case ClientPacketID.eGuildRejectPeace
            Call HandleGuildRejectPeace(UserIndex)
        Case ClientPacketID.eGuildAcceptAlliance
            Call HandleGuildAcceptAlliance(UserIndex)
        Case ClientPacketID.eGuildOfferPeace
            Call HandleGuildOfferPeace(UserIndex)
        Case ClientPacketID.eGuildOfferAlliance
            Call HandleGuildOfferAlliance(UserIndex)
        Case ClientPacketID.eGuildAllianceDetails
            Call HandleGuildAllianceDetails(UserIndex)
        Case ClientPacketID.eGuildPeaceDetails
            Call HandleGuildPeaceDetails(UserIndex)
        Case ClientPacketID.eGuildRequestJoinerInfo
            Call HandleGuildRequestJoinerInfo(UserIndex)
        Case ClientPacketID.eGuildAlliancePropList
            Call HandleGuildAlliancePropList(UserIndex)
        Case ClientPacketID.eGuildPeacePropList
            Call HandleGuildPeacePropList(UserIndex)
        Case ClientPacketID.eGuildDeclareWar
            Call HandleGuildDeclareWar(UserIndex)
        Case ClientPacketID.eGuildNewWebsite
            Call HandleGuildNewWebsite(UserIndex)
        Case ClientPacketID.eGuildAcceptNewMember
            Call HandleGuildAcceptNewMember(UserIndex)
        Case ClientPacketID.eGuildRejectNewMember
            Call HandleGuildRejectNewMember(UserIndex)
        Case ClientPacketID.eGuildKickMember
            Call HandleGuildKickMember(UserIndex)
        Case ClientPacketID.eGuildUpdateNews
            Call HandleGuildUpdateNews(UserIndex)
        Case ClientPacketID.eGuildMemberInfo
            Call HandleGuildMemberInfo(UserIndex)
        Case ClientPacketID.eGuildOpenElections
            Call HandleGuildOpenElections(UserIndex)
        Case ClientPacketID.eGuildRequestMembership
            Call HandleGuildRequestMembership(UserIndex)
        Case ClientPacketID.eGuildRequestDetails
            Call HandleGuildRequestDetails(UserIndex)
        Case ClientPacketID.eOnline
            Call HandleOnline(UserIndex)
        Case ClientPacketID.eQuit
            Call HandleQuit(UserIndex)
        Case ClientPacketID.eGuildLeave
            Call HandleGuildLeave(UserIndex)
        Case ClientPacketID.eRequestAccountState
            Call HandleRequestAccountState(UserIndex)
        Case ClientPacketID.ePetStand
            Call HandlePetStand(UserIndex)
        Case ClientPacketID.ePetFollow
            Call HandlePetFollow(UserIndex)
        Case ClientPacketID.ePetLeave
            Call HandlePetLeave(UserIndex)
        Case ClientPacketID.eGrupoMsg
            Call HandleGrupoMsg(UserIndex)
        Case ClientPacketID.eTrainList
            Call HandleTrainList(UserIndex)
        Case ClientPacketID.eRest
            Call HandleRest(UserIndex)
        Case ClientPacketID.eMeditate
            Call HandleMeditate(UserIndex)
        Case ClientPacketID.eResucitate
            Call HandleResucitate(UserIndex)
        Case ClientPacketID.eHeal
            Call HandleHeal(UserIndex)
        Case ClientPacketID.eHelp
            Call HandleHelp(UserIndex)
        Case ClientPacketID.eRequestStats
            Call HandleRequestStats(UserIndex)
        Case ClientPacketID.eCommerceStart
            Call HandleCommerceStart(UserIndex)
        Case ClientPacketID.eBankStart
            Call HandleBankStart(UserIndex)
        Case ClientPacketID.eEnlist
            Call HandleEnlist(UserIndex)
        Case ClientPacketID.eInformation
            Call HandleInformation(UserIndex)
        Case ClientPacketID.eReward
            Call HandleReward(UserIndex)
        Case ClientPacketID.eRequestMOTD
            Call HandleRequestMOTD(UserIndex)
        Case ClientPacketID.eUpTime
            Call HandleUpTime(UserIndex)
        Case ClientPacketID.eGuildMessage
            Call HandleGuildMessage(UserIndex)
        Case ClientPacketID.eGuildOnline
            Call HandleGuildOnline(UserIndex)
        Case ClientPacketID.eCouncilMessage
            Call HandleCouncilMessage(UserIndex)
        Case ClientPacketID.eRoleMasterRequest
            Call HandleRoleMasterRequest(UserIndex)
        Case ClientPacketID.eChangeDescription
            Call HandleChangeDescription(UserIndex)
        Case ClientPacketID.eGuildVote
            Call HandleGuildVote(UserIndex)
        Case ClientPacketID.epunishments
            Call HandlePunishments(UserIndex)
        Case ClientPacketID.eGamble
            Call HandleGamble(UserIndex)
        Case ClientPacketID.eLeaveFaction
            Call HandleLeaveFaction(UserIndex)
        Case ClientPacketID.eBankExtractGold
            Call HandleBankExtractGold(UserIndex)
        Case ClientPacketID.eBankDepositGold
            Call HandleBankDepositGold(UserIndex)
        Case ClientPacketID.eDenounce
            Call HandleDenounce(UserIndex)
        Case ClientPacketID.eGMMessage
            Call HandleGMMessage(UserIndex)
        Case ClientPacketID.eshowName
            Call HandleShowName(UserIndex)
        Case ClientPacketID.eOnlineRoyalArmy
            Call HandleOnlineRoyalArmy(UserIndex)
        Case ClientPacketID.eOnlineChaosLegion
            Call HandleOnlineChaosLegion(UserIndex)
        Case ClientPacketID.eGoNearby
            Call HandleGoNearby(UserIndex)
        Case ClientPacketID.ecomment
            Call HandleComment(UserIndex)
        Case ClientPacketID.eserverTime
            Call HandleServerTime(UserIndex)
        Case ClientPacketID.eWhere
            Call HandleWhere(UserIndex)
        Case ClientPacketID.eCreaturesInMap
            Call HandleCreaturesInMap(UserIndex)
        Case ClientPacketID.eWarpMeToTarget
            Call HandleWarpMeToTarget(UserIndex)
        Case ClientPacketID.eWarpChar
            Call HandleWarpChar(UserIndex)
        Case ClientPacketID.eSilence
            Call HandleSilence(UserIndex)
        Case ClientPacketID.eSOSShowList
            Call HandleSOSShowList(UserIndex)
        Case ClientPacketID.eSOSRemove
            Call HandleSOSRemove(UserIndex)
        Case ClientPacketID.eGoToChar
            Call HandleGoToChar(UserIndex)
        Case ClientPacketID.einvisible
            Call HandleInvisible(UserIndex)
        Case ClientPacketID.eGMPanel
            Call HandleGMPanel(UserIndex)
        Case ClientPacketID.eRequestUserList
            Call HandleRequestUserList(UserIndex)
        Case ClientPacketID.eWorking
            Call HandleWorking(UserIndex)
        Case ClientPacketID.eHiding
            Call HandleHiding(UserIndex)
        Case ClientPacketID.eJail
            Call HandleJail(UserIndex)
        Case ClientPacketID.eKillNPC
            Call HandleKillNPC(UserIndex)
        Case ClientPacketID.eWarnUser
            Call HandleWarnUser(UserIndex)
        Case ClientPacketID.eEditChar
            Call HandleEditChar(UserIndex)
        Case ClientPacketID.eRequestCharInfo
            Call HandleRequestCharInfo(UserIndex)
        Case ClientPacketID.eRequestCharStats
            Call HandleRequestCharStats(UserIndex)
        Case ClientPacketID.eRequestCharGold
            Call HandleRequestCharGold(UserIndex)
        Case ClientPacketID.eRequestCharInventory
            Call HandleRequestCharInventory(UserIndex)
        Case ClientPacketID.eRequestCharBank
            Call HandleRequestCharBank(UserIndex)
        Case ClientPacketID.eRequestCharSkills
            Call HandleRequestCharSkills(UserIndex)
        Case ClientPacketID.eReviveChar
            Call HandleReviveChar(UserIndex)
        Case ClientPacketID.eSeguirMouse
            Call HandleSeguirMouse(UserIndex)
        Case ClientPacketID.eSendPosSeguimiento
            Call HandleSendPosMovimiento(UserIndex)
        Case ClientPacketID.eNotifyInventarioHechizos
            Call HandleNotifyInventariohechizos(UserIndex)
        Case ClientPacketID.eOnlineGM
            Call HandleOnlineGM(UserIndex)
        Case ClientPacketID.eOnlineMap
            Call HandleOnlineMap(UserIndex)
        Case ClientPacketID.eForgive
            Call HandleForgive(UserIndex)
        Case ClientPacketID.ePerdonFaccion
            Call HandlePerdonFaccion(userindex)
        Case ClientPacketID.eStartEvent
            Call HandleStartEvent(UserIndex)
        Case ClientPacketID.eCancelarEvento
            Call HandleCancelarEvento(UserIndex)
        Case ClientPacketID.eKick
            Call HandleKick(UserIndex)
        Case ClientPacketID.eExecute
            Call HandleExecute(UserIndex)
        Case ClientPacketID.eBanChar
            Call HandleBanChar(UserIndex)
        Case ClientPacketID.eUnbanChar
            Call HandleUnbanChar(UserIndex)
        Case ClientPacketID.eNPCFollow
            Call HandleNPCFollow(UserIndex)
        Case ClientPacketID.eSummonChar
            Call HandleSummonChar(UserIndex)
        Case ClientPacketID.eSpawnListRequest
            Call HandleSpawnListRequest(UserIndex)
        Case ClientPacketID.eSpawnCreature
            Call HandleSpawnCreature(UserIndex)
        Case ClientPacketID.eResetNPCInventory
            Call HandleResetNPCInventory(UserIndex)
        Case ClientPacketID.eCleanWorld
            Call HandleCleanWorld(UserIndex)
        Case ClientPacketID.eServerMessage
            Call HandleServerMessage(UserIndex)
        Case ClientPacketID.eNickToIP
            Call HandleNickToIP(UserIndex)
        Case ClientPacketID.eIPToNick
            Call HandleIPToNick(UserIndex)
        Case ClientPacketID.eGuildOnlineMembers
            Call HandleGuildOnlineMembers(UserIndex)
        Case ClientPacketID.eTeleportCreate
            Call HandleTeleportCreate(UserIndex)
        Case ClientPacketID.eTeleportDestroy
            Call HandleTeleportDestroy(UserIndex)
        Case ClientPacketID.eRainToggle
            Call HandleRainToggle(UserIndex)
        Case ClientPacketID.eSetCharDescription
            Call HandleSetCharDescription(UserIndex)
        Case ClientPacketID.eForceMIDIToMap
            Call HanldeForceMIDIToMap(UserIndex)
        Case ClientPacketID.eForceWAVEToMap
            Call HandleForceWAVEToMap(UserIndex)
        Case ClientPacketID.eRoyalArmyMessage
            Call HandleRoyalArmyMessage(UserIndex)
        Case ClientPacketID.eChaosLegionMessage
            Call HandleChaosLegionMessage(UserIndex)
        Case ClientPacketID.eTalkAsNPC
            Call HandleTalkAsNPC(UserIndex)
        Case ClientPacketID.eDestroyAllItemsInArea
            Call HandleDestroyAllItemsInArea(UserIndex)
        Case ClientPacketID.eAcceptRoyalCouncilMember
            Call HandleAcceptRoyalCouncilMember(UserIndex)
        Case ClientPacketID.eAcceptChaosCouncilMember
            Call HandleAcceptChaosCouncilMember(UserIndex)
        Case ClientPacketID.eItemsInTheFloor
            Call HandleItemsInTheFloor(UserIndex)
        Case ClientPacketID.eMakeDumb
            Call HandleMakeDumb(UserIndex)
        Case ClientPacketID.eMakeDumbNoMore
            Call HandleMakeDumbNoMore(UserIndex)
        Case ClientPacketID.eCouncilKick
            Call HandleCouncilKick(UserIndex)
        Case ClientPacketID.eSetTrigger
            Call HandleSetTrigger(UserIndex)
        Case ClientPacketID.eAskTrigger
            Call HandleAskTrigger(UserIndex)
        Case ClientPacketID.eGuildMemberList
            Call HandleGuildMemberList(UserIndex)
        Case ClientPacketID.eGuildBan
            Call HandleGuildBan(UserIndex)
        Case ClientPacketID.eCreateItem
            Call HandleCreateItem(UserIndex)
        Case ClientPacketID.eDestroyItems
            Call HandleDestroyItems(UserIndex)
        Case ClientPacketID.eChaosLegionKick
            Call HandleChaosLegionKick(UserIndex)
        Case ClientPacketID.eRoyalArmyKick
            Call HandleRoyalArmyKick(UserIndex)
        Case ClientPacketID.eForceMIDIAll
            Call HandleForceMIDIAll(UserIndex)
        Case ClientPacketID.eForceWAVEAll
            Call HandleForceWAVEAll(UserIndex)
        Case ClientPacketID.eRemovePunishment
            Call HandleRemovePunishment(UserIndex)
        Case ClientPacketID.eTileBlockedToggle
            Call HandleTile_BlockedToggle(UserIndex)
        Case ClientPacketID.eKillNPCNoRespawn
            Call HandleKillNPCNoRespawn(UserIndex)
        Case ClientPacketID.eKillAllNearbyNPCs
            Call HandleKillAllNearbyNPCs(UserIndex)
        Case ClientPacketID.eLastIP
            Call HandleLastIP(UserIndex)
        Case ClientPacketID.eChangeMOTD
            Call HandleChangeMOTD(UserIndex)
        Case ClientPacketID.eSetMOTD
            Call HandleSetMOTD(UserIndex)
        Case ClientPacketID.eSystemMessage
            Call HandleSystemMessage(UserIndex)
        Case ClientPacketID.eCreateNPC
            Call HandleCreateNPC(UserIndex)
        Case ClientPacketID.eCreateNPCWithRespawn
            Call HandleCreateNPCWithRespawn(UserIndex)
        Case ClientPacketID.eImperialArmour
            Call HandleImperialArmour(UserIndex)
        Case ClientPacketID.eChaosArmour
            Call HandleChaosArmour(UserIndex)
        Case ClientPacketID.eNavigateToggle
            Call HandleNavigateToggle(UserIndex)
        Case ClientPacketID.eServerOpenToUsersToggle
            Call HandleServerOpenToUsersToggle(UserIndex)
        Case ClientPacketID.eParticipar
            Call HandleParticipar(UserIndex)
        Case ClientPacketID.eTurnCriminal
            Call HandleTurnCriminal(UserIndex)
        Case ClientPacketID.eResetFactions
            Call HandleResetFactions(UserIndex)
        Case ClientPacketID.eRemoveCharFromGuild
            Call HandleRemoveCharFromGuild(UserIndex)
        Case ClientPacketID.eAlterName
            Call HandleAlterName(UserIndex)
        Case ClientPacketID.eDoBackUp
            Call HandleDoBackUp(UserIndex)
        Case ClientPacketID.eShowGuildMessages
            Call HandleShowGuildMessages(UserIndex)
        Case ClientPacketID.eChangeMapInfoPK
            Call HandleChangeMapInfoPK(UserIndex)
        Case ClientPacketID.eChangeMapInfoBackup
            Call HandleChangeMapInfoBackup(UserIndex)
        Case ClientPacketID.eChangeMapInfoRestricted
            Call HandleChangeMapInfoRestricted(UserIndex)
        Case ClientPacketID.eChangeMapInfoNoMagic
            Call HandleChangeMapInfoNoMagic(UserIndex)
        Case ClientPacketID.eChangeMapInfoNoInvi
            Call HandleChangeMapInfoNoInvi(UserIndex)
        Case ClientPacketID.eChangeMapInfoNoResu
            Call HandleChangeMapInfoNoResu(UserIndex)
        Case ClientPacketID.eChangeMapInfoLand
            Call HandleChangeMapInfoLand(UserIndex)
        Case ClientPacketID.eChangeMapInfoZone
            Call HandleChangeMapInfoZone(UserIndex)
        Case ClientPacketID.eChangeMapSetting
            Call HandleChangeMapSetting(UserIndex)
        Case ClientPacketID.eSaveChars
            Call HandleSaveChars(UserIndex)
        Case ClientPacketID.eCleanSOS
            Call HandleCleanSOS(UserIndex)
        Case ClientPacketID.eShowServerForm
            Call HandleShowServerForm(UserIndex)
        Case ClientPacketID.enight
            Call HandleNight(UserIndex)
        Case ClientPacketID.eKickAllChars
            Call HandleKickAllChars(UserIndex)
        Case ClientPacketID.eReloadNPCs
            Call HandleReloadNPCs(UserIndex)
        Case ClientPacketID.eReloadServerIni
            Call HandleReloadServerIni(UserIndex)
        Case ClientPacketID.eReloadSpells
            Call HandleReloadSpells(UserIndex)
        Case ClientPacketID.eReloadObjects
            Call HandleReloadObjects(UserIndex)
        Case ClientPacketID.eChatColor
            Call HandleChatColor(UserIndex)
        Case ClientPacketID.eIgnored
            Call HandleIgnored(UserIndex)
        Case ClientPacketID.eCheckSlot
            Call HandleCheckSlot(UserIndex)
        Case ClientPacketID.eSetSpeed
            Call HandleSetSpeed(UserIndex)
        Case ClientPacketID.eGlobalMessage
            Call HandleGlobalMessage(UserIndex)
        Case ClientPacketID.eGlobalOnOff
            Call HandleGlobalOnOff(UserIndex)
        Case ClientPacketID.eUseKey
            Call HandleUseKey(UserIndex)
        Case ClientPacketID.eDay
            Call HandleDay(UserIndex)
        Case ClientPacketID.eSetTime
            Call HandleSetTime(UserIndex)
        Case ClientPacketID.eDonateGold
            Call HandleDonateGold(UserIndex)
        Case ClientPacketID.ePromedio
            Call HandlePromedio(UserIndex)
        Case ClientPacketID.eGiveItem
            Call HandleGiveItem(UserIndex)
        Case ClientPacketID.eOfertaInicial
            Call HandleOfertaInicial(UserIndex)
        Case ClientPacketID.eOfertaDeSubasta
            Call HandleOfertaDeSubasta(UserIndex)
        Case ClientPacketID.eQuestionGM
            Call HandleQuestionGM(UserIndex)
        Case ClientPacketID.eCuentaRegresiva
            Call HandleCuentaRegresiva(UserIndex)
        Case ClientPacketID.ePossUser
            Call HandlePossUser(UserIndex)
        Case ClientPacketID.eDuel
            Call HandleDuel(UserIndex)
        Case ClientPacketID.eAcceptDuel
            Call HandleAcceptDuel(UserIndex)
        Case ClientPacketID.eCancelDuel
            Call HandleCancelDuel(UserIndex)
        Case ClientPacketID.eQuitDuel
            Call HandleQuitDuel(UserIndex)
        Case ClientPacketID.eNieveToggle
            Call HandleNieveToggle(UserIndex)
        Case ClientPacketID.eNieblaToggle
            Call HandleNieblaToggle(UserIndex)
        Case ClientPacketID.eTransFerGold
            Call HandleTransFerGold(UserIndex)
        Case ClientPacketID.eMoveitem
            Call HandleMoveItem(UserIndex)
        Case ClientPacketID.eGenio
            Call HandleGenio(UserIndex)
        Case ClientPacketID.eCasarse
            Call HandleCasamiento(UserIndex)
        Case ClientPacketID.eCraftAlquimista
            Call HandleCraftAlquimia(UserIndex)
        Case ClientPacketID.eFlagTrabajar
            Call HandleFlagTrabajar(UserIndex)
        Case ClientPacketID.eCraftSastre
            Call HandleCraftSastre(UserIndex)
        Case ClientPacketID.eMensajeUser
            Call HandleMensajeUser(UserIndex)
        Case ClientPacketID.eTraerBoveda
            Call HandleTraerBoveda(UserIndex)
        Case ClientPacketID.eCompletarAccion
            Call HandleCompletarAccion(UserIndex)
        Case ClientPacketID.eInvitarGrupo
            Call HandleInvitarGrupo(UserIndex)
        Case ClientPacketID.eResponderPregunta
            Call HandleResponderPregunta(UserIndex)
        Case ClientPacketID.eRequestGrupo
            Call HandleRequestGrupo(UserIndex)
        Case ClientPacketID.eAbandonarGrupo
            Call HandleAbandonarGrupo(UserIndex)
        Case ClientPacketID.eHecharDeGrupo
            Call HandleHecharDeGrupo(UserIndex)
        Case ClientPacketID.eMacroPossent
            Call HandleMacroPos(UserIndex)
        Case ClientPacketID.eSubastaInfo
            Call HandleSubastaInfo(UserIndex)
        Case ClientPacketID.eBanCuenta
            Call HandleBanCuenta(UserIndex)
        Case ClientPacketID.eUnbanCuenta
            Call HandleUnBanCuenta(UserIndex)
        Case ClientPacketID.eCerrarCliente
            Call HandleCerrarCliente(UserIndex)
        Case ClientPacketID.eEventoInfo
            Call HandleEventoInfo(UserIndex)
        Case ClientPacketID.eCrearEvento
            Call HandleCrearEvento(UserIndex)
        Case ClientPacketID.eBanTemporal
            Call HandleBanTemporal(UserIndex)
        Case ClientPacketID.eCancelarExit
            Call HandleCancelarExit(UserIndex)
        Case ClientPacketID.eCrearTorneo
            Call HandleCrearTorneo(UserIndex)
        Case ClientPacketID.eComenzarTorneo
            Call HandleComenzarTorneo(UserIndex)
        Case ClientPacketID.eCancelarTorneo
            Call HandleCancelarTorneo(UserIndex)
        Case ClientPacketID.eBusquedaTesoro
            Call HandleBusquedaTesoro(UserIndex)
        Case ClientPacketID.eCompletarViaje
            Call HandleCompletarViaje(UserIndex)
        Case ClientPacketID.eBovedaMoveItem
            Call HandleBovedaMoveItem(UserIndex)
        Case ClientPacketID.eQuieroFundarClan
            Call HandleQuieroFundarClan(UserIndex)
        Case ClientPacketID.ellamadadeclan
            Call HandleLlamadadeClan(UserIndex)
        Case ClientPacketID.eMarcaDeClanPack
            Call HandleMarcaDeClan(UserIndex)
        Case ClientPacketID.eMarcaDeGMPack
            Call HandleMarcaDeGM(UserIndex)
        Case ClientPacketID.eQuest
            Call HandleQuest(UserIndex)
        Case ClientPacketID.eQuestAccept
            Call HandleQuestAccept(UserIndex)
        Case ClientPacketID.eQuestListRequest
            Call HandleQuestListRequest(UserIndex)
        Case ClientPacketID.eQuestDetailsRequest
            Call HandleQuestDetailsRequest(UserIndex)
        Case ClientPacketID.eQuestAbandon
            Call HandleQuestAbandon(UserIndex)
        Case ClientPacketID.eSeguroClan
            Call HandleSeguroClan(UserIndex)
        Case ClientPacketID.ehome
            Call HandleHome(UserIndex)
        Case ClientPacketID.eConsulta
            Call HandleConsulta(UserIndex)
        Case ClientPacketID.eGetMapInfo
            Call HandleGetMapInfo(UserIndex)
        Case ClientPacketID.eFinEvento
            Call HandleFinEvento(UserIndex)
        Case ClientPacketID.eSeguroResu
            Call HandleSeguroResu(UserIndex)
        Case ClientPacketID.eCuentaExtractItem
            Call HandleCuentaExtractItem(UserIndex)
        Case ClientPacketID.eCuentaDeposit
            Call HandleCuentaDeposit(UserIndex)
        Case ClientPacketID.eCreateEvent
            Call HandleCreateEvent(UserIndex)
        Case ClientPacketID.eCommerceSendChatMessage
            Call HandleCommerceSendChatMessage(UserIndex)
        Case ClientPacketID.eLogMacroClickHechizo
            Call HandleLogMacroClickHechizo(UserIndex)
        Case ClientPacketID.eAddItemCrafting
            Call HandleAddItemCrafting(UserIndex)
        Case ClientPacketID.eRemoveItemCrafting
            Call HandleRemoveItemCrafting(UserIndex)
        Case ClientPacketID.eAddCatalyst
            Call HandleAddCatalyst(UserIndex)
        Case ClientPacketID.eRemoveCatalyst
            Call HandleRemoveCatalyst(UserIndex)
        Case ClientPacketID.eCraftItem
            Call HandleCraftItem(UserIndex)
        Case ClientPacketID.eCloseCrafting
            Call HandleCloseCrafting(UserIndex)
        Case ClientPacketID.eMoveCraftItem
            Call HandleMoveCraftItem(UserIndex)
        Case ClientPacketID.ePetLeaveAll
            Call HandlePetLeaveAll(UserIndex)
        Case ClientPacketID.eResetChar
            Call HandleResetChar(UserIndex)
        Case ClientPacketID.eResetearPersonaje
            Call HandleResetearPersonaje(UserIndex)
        Case ClientPacketID.eDeleteItem
            Call HandleDeleteItem(UserIndex)
        Case ClientPacketID.eFinalizarPescaEspecial
            Call HandleFinalizarPescaEspecial(UserIndex)
        Case ClientPacketID.eRomperCania
            Call HandleRomperCania(UserIndex)
        Case ClientPacketID.eRepeatMacro
            Call HandleRepeatMacro(UserIndex)
        Case ClientPacketID.eBuyShopItem
            Call HandleBuyShopItem(userindex)
        Case ClientPacketID.ePublicarPersonajeMAO
            Call HandlePublicarPersonajeMAO(userindex)
        Case ClientPacketID.eEventoFaccionario
            Call HandleEventoFaccionario(UserIndex)
        Case ClientPacketID.eRequestDebug
            Call HandleDebugRequest(UserIndex)
        Case ClientPacketID.eLobbyCommand
            Call HandleLobbyCommand(UserIndex)
        Case ClientPacketID.eFeatureToggle
            Call HandleFeatureToggle(UserIndex)
        Case ClientPacketID.eActionOnGroupFrame
            Call HandleActionOnGroupFrame(UserIndex)
        Case ClientPacketID.eSendTelemetry
            Call HandleSendTelemetry(UserIndex)
        Case ClientPacketID.eSetHotkeySlot
            Call HandleSetHotkeySlot(UserIndex)
        Case ClientPacketID.eUseHKeySlot
            Call HandleUseHKeySlot(UserIndex)
        Case ClientPacketID.eAntiCheatMessage
            Call HandleAntiCheatMessage(UserIndex)
#If PYMMO = 0 Then
        Case ClientPacketID.eCreateAccount
            Call HandleCreateAccount(ConnectionId)
        Case ClientPacketID.eLoginAccount
            Call HandleLoginAccount(ConnectionId)
        Case ClientPacketID.eDeleteCharacter
            Call HandleDeleteCharacter(ConnectionId)
#End If
        Case Else
            Err.raise -1, "Invalid Message"
    End Select
    
    If (Message.GetAvailable() > 0) Then
        Err.raise &HDEADBEEF, "HandleIncomingData", "El paquete '" & PacketID & "' se encuentra en mal estado con '" & Message.GetAvailable() & "' bytes de mas por el usuario '" & UserList(UserIndex).Name & "'"
    End If
    Call PerformTimeLimitCheck(performance_timer, "Protocol handling message " & PacketId, 100)
HandleIncomingData_Err:
    
    Set Reader = Nothing

    If Err.Number <> 0 Then
        Call TraceError(Err.Number, Err.Description & vbNewLine & "PackedID: " & PacketID & vbNewLine & IIf(UserList(UserIndex).flags.UserLogged, "UserName: " & UserList(UserIndex).name, "UserIndex: " & UserIndex), "Protocol.HandleIncomingData", Erl)
        'Call CloseSocket(UserIndex)
        
        HandleIncomingData = False
    End If
End Function

#If PYMMO = 0 Then

Private Sub HandleCreateAccount(ByVal ConnectionId As Long)
    On Error GoTo HandleCreateAccount_Err:
    
    Dim username As String
    Dim Password As String
    username = Reader.ReadString8
    Password = Reader.ReadString8
    Dim UserIndex As Integer
    UserIndex = MapConnectionToUser(ConnectionId)
    If UserIndex < 1 Then
        Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("No hay slot disponibles para el usuario."))
        Call KickConnection(ConnectionId)
        Exit Sub
    End If
    If (username = "" Or Password = "" Or LenB(Password) <= 3) Then
        Call WriteErrorMsg(userindex, "Parametros incorrectos")
        Call CloseSocket(userindex)
        Exit Sub
    End If

    Dim result As ADODB.Recordset
    Set result = Query("INSERT INTO account (email, password, salt, validate_code) VALUES (?,?,?,?)", LCase(username), Password, Password, "123")

    If (result Is Nothing) Then
        Call WriteErrorMsg(userindex, "Ya hay una cuenta asociada con ese email")
        Call CloseSocket(userindex)
        Exit Sub
    End If
    
    Set result = Query("SELECT id FROM account WHERE email=?", username)
    UserList(userindex).AccountID = result!ID
    
    Dim Personajes() As t_PersonajeCuenta
    Call WriteAccountCharacterList(userindex, Personajes, 0)

    Exit Sub
HandleCreateAccount_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateAccount", Erl)
End Sub

Private Sub HandleLoginAccount(ByVal ConnectionId As Long)
    On Error GoTo LoginAccount_Err:
    
    Dim username As String
    Dim Password As String
    username = Reader.ReadString8
    Password = Reader.ReadString8
    Dim UserIndex As Integer
    UserIndex = MapConnectionToUser(ConnectionId)
    If UserIndex < 1 Then
        Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("No hay slot disponibles para el usuario."))
        Call KickConnection(ConnectionId)
        Exit Sub
    End If
    If (username = "" Or Password = "" Or LenB(Password) <= 3) Then
        Call WriteErrorMsg(userindex, "Parametros incorrectos")
        Call CloseSocket(userindex)
        Exit Sub
    End If

    Dim result As ADODB.Recordset
    Set result = Query("SELECT * FROM account WHERE UPPER(email)=UPPER(?) AND password=?", username, Password)
    
    If (result.EOF) Then
        Call WriteErrorMsg(UserIndex, "Usuario o Contraseña erronea.")
        Call CloseSocket(userindex)
        Exit Sub
    End If
        
    UserList(userindex).AccountID = result!ID
    
    Dim Personajes(1 To 10) As t_PersonajeCuenta
    Dim Count As Long
    Count = GetPersonajesCuentaDatabase(result!ID, Personajes)
    
    Call WriteAccountCharacterList(userindex, Personajes, Count)

    Exit Sub
LoginAccount_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginAccount", Erl)
End Sub

Private Sub HandleDeleteCharacter(ByVal ConnectionId As Long)
    On Error GoTo DeleteCharacter_Err:

DeleteCharacter_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDeleteCharacter", Erl)
End Sub

Private Sub HandleLoginExistingChar(ByVal ConnectionId As Long)
        On Error GoTo ErrHandler

        Dim user_name    As String
        Dim UserIndex As Integer
        UserIndex = Mapping(ConnectionId).UserRef.ArrayIndex
        user_name = Reader.ReadString8
        Call ConnectUser(userindex, user_name)
        Exit Sub
ErrHandler:
        Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginExistingChar", Erl)
End Sub

#End If
#If PYMMO = 1 Then
''
' Handles the "LoginExistingChar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal ConnectionId As Long)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        ''Last Modification: 01/12/08 Ladder
        '***************************************************

        On Error GoTo ErrHandler

        Dim user_name    As String
        Dim CuentaEmail As String
        Dim Version     As String
        Dim MD5         As String
        Dim encrypted_session_token As String
        Dim encrypted_username As String
        
        encrypted_session_token = Reader.ReadString8
        encrypted_username = Reader.ReadString8
        Version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
        MD5 = Reader.ReadString8()

        If Len(encrypted_session_token) <> 88 Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Cliente inválido, por favor realice una actualización."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
                
        
        Dim encrypted_session_token_byte() As Byte
        Call AO20CryptoSysWrapper.Str2ByteArr(encrypted_session_token, encrypted_session_token_byte)
        
        Dim decrypted_session_token As String
        decrypted_session_token = AO20CryptoSysWrapper.DECRYPT(PrivateKey, cnvStringFromHexStr(cnvToHex(encrypted_session_token_byte)))
                
        If Not IsBase64(decrypted_session_token) Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Cliente inválido, por favor realice una actualización"))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
        
        ' Para recibir el ID del user
        Dim RS As ADODB.Recordset
        Set RS = Query("select * from tokens where decrypted_token = '" & decrypted_session_token & "'")
                
        If RS Is Nothing Or RS.RecordCount = 0 Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Sesión inválida, conéctese nuevamente."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
        
        CuentaEmail = CStr(RS!UserName)
                    
        If RS!encrypted_token <> encrypted_session_token Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Cliente inválido, por favor realice una actualización."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
        Dim UserIndex As Integer
        UserIndex = MapConnectionToUser(ConnectionId)
        If UserIndex < 1 Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("No hay slot disponibles para el usuario."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
        
        UserList(UserIndex).encrypted_session_token_db_id = RS!id
        UserList(UserIndex).encrypted_session_token = encrypted_session_token
        UserList(UserIndex).decrypted_session_token = decrypted_session_token
        UserList(UserIndex).public_key = mid$(decrypted_session_token, 1, 16)
        
        user_name = AO20CryptoSysWrapper.DECRYPT(cnvHexStrFromString(UserList(UserIndex).public_key), encrypted_username)
         
        If Not EntrarCuenta(UserIndex, CuentaEmail, MD5) Then
            Call CloseSocket(UserIndex)
            Exit Sub
        End If
        Call ConnectUser(UserIndex, user_name, False)
        Exit Sub
    
ErrHandler:
        Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginExistingChar", Erl)

End Sub

Private Sub HandleLoginNewChar(ByVal ConnectionId As Long)
        On Error GoTo ErrHandler


        Dim UserName    As String
        Dim CuentaEmail As String
        Dim Version     As String
        Dim MD5         As String
        Dim encrypted_session_token As String
        Dim encrypted_username As String
        Dim race     As e_Raza
        Dim gender   As e_Genero
        Dim Hogar    As e_Ciudad
        Dim Class As e_Class
        Dim Head        As Integer
        
         
        encrypted_session_token = Reader.ReadString8
        encrypted_username = Reader.ReadString8
        
106     Version = CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8()) & "." & CStr(Reader.ReadInt8())
114     MD5 = Reader.ReadString8()

110     race = Reader.ReadInt8()
112     gender = Reader.ReadInt8()
113     Class = Reader.ReadInt8()
116     Head = Reader.ReadInt16()
118     Hogar = Reader.ReadInt8()

        If Len(encrypted_session_token) <> 88 Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Cliente inválido, por favor realice una actualización."))
            Exit Sub
        End If

        Dim encrypted_session_token_byte() As Byte
        Call AO20CryptoSysWrapper.Str2ByteArr(encrypted_session_token, encrypted_session_token_byte)
        
        Dim decrypted_session_token As String
        decrypted_session_token = AO20CryptoSysWrapper.DECRYPT(PrivateKey, cnvStringFromHexStr(cnvToHex(encrypted_session_token_byte)))
                
        If Not IsBase64(decrypted_session_token) Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Cliente inválido, por favor realice una actualización."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
            ' Para recibir el ID del user
        Dim RS As ADODB.Recordset
        Set RS = Query("select * from tokens where decrypted_token = '" & decrypted_session_token & "'")
                
        If RS Is Nothing Or RS.RecordCount = 0 Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Sesión inválida, conectese nuevamente."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
        
        CuentaEmail = CStr(RS!UserName)
        If RS!encrypted_token <> encrypted_session_token Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("Cliente inválido, por favor realice una actualización."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
        
        Dim UserIndex As Integer
        UserIndex = MapConnectionToUser(ConnectionId)
        If UserIndex < 1 Then
            Call modSendData.SendToConnection(ConnectionId, PrepareShowMessageBox("No hay slot disponibles para el usuario."))
            Call KickConnection(ConnectionId)
            Exit Sub
        End If
        
        UserList(UserIndex).encrypted_session_token_db_id = RS!id
        UserList(UserIndex).encrypted_session_token = encrypted_session_token
        UserList(UserIndex).decrypted_session_token = decrypted_session_token
        UserList(UserIndex).public_key = mid$(decrypted_session_token, 1, 16)

        UserName = AO20CryptoSysWrapper.DECRYPT(cnvHexStrFromString(UserList(UserIndex).public_key), encrypted_username)
    
126     If PuedeCrearPersonajes = 0 Then
128         Call WriteShowMessageBox(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
130         Call CloseSocket(UserIndex)
            Exit Sub

        End If

132     If aClon.MaxPersonajes(UserList(UserIndex).ConnectionDetails.IP) Then
134         Call WriteShowMessageBox(UserIndex, "Has creado demasiados personajes.")
136         Call CloseSocket(UserIndex)
            Exit Sub

        End If

148     If EsGmChar(UserName) Then
            
150         If AdministratorAccounts(UCase$(UserName)) <> UCase$(CuentaEmail) Then
152             Call WriteShowMessageBox(UserIndex, "El nombre de usuario ingresado está siendo ocupado por un miembro del Staff.")
154             Call CloseSocket(UserIndex)
                Exit Sub

            End If
            
        End If
        UserList(userindex).AccountID = -1
        If Not EntrarCuenta(userindex, CuentaEmail, md5) Then
            Call CloseSocket(userindex)
            Exit Sub
        End If
        Debug.Assert UserList(userindex).AccountID > -1
        
        Dim num_pc As Byte
        num_pc = GetPersonajesCountByIDDatabase(UserList(UserIndex).AccountID)
        Debug.Assert num_pc > 0
        Dim user_tier As e_TipoUsuario
        user_tier = GetPatronTierFromAccountID(UserList(UserIndex).AccountID)
        Dim max_pc_for_tier As Byte
        max_pc_for_tier = MaxCharacterForTier(user_tier)
        Debug.Assert max_pc_for_tier > 0
        If num_pc >= Min(max_pc_for_tier, MAX_PERSONAJES) Then
            Call WriteShowMessageBox(UserIndex, "You need to upgrade your account to create more characters, please visit https://www.patreon.com/nolandstudios")
            Call CloseSocket(userindex)
            Exit Sub
        End If
        
        If Not ConnectNewUser(userindex, username, race, gender, Class, Head, Hogar) Then
            Call CloseSocket(userindex)
            Exit Sub
        End If
        
        
        Exit Sub
    
ErrHandler:
     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginNewChar", Erl)
End Sub

#ElseIf PYMMO = 0 Then
    

'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal userindex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

        Dim name As String
        Dim race     As e_Raza
        Dim gender   As e_Genero
        Dim Hogar    As e_Ciudad
        Dim Class As e_Class
        Dim Head        As Integer

        name = Reader.ReadString8
110     race = Reader.ReadInt()
112     gender = Reader.ReadInt()
113     Class = Reader.ReadInt()
116     Head = Reader.ReadInt()
118     Hogar = Reader.ReadInt()

126     If PuedeCrearPersonajes = 0 Then
128         Call WriteShowMessageBox(userindex, "La creacion de personajes en este servidor se ha deshabilitado.")
130         Call CloseSocket(userindex)
            Exit Sub

        End If

132     If aClon.MaxPersonajes(UserList(UserIndex).ConnectionDetails.IP) Then
134         Call WriteShowMessageBox(userindex, "Has creado demasiados personajes.")
136         Call CloseSocket(userindex)
            Exit Sub

        End If

        'Check if we reached MAX_PERSONAJES for this account after updateing the UserList(userindex).AccountID in the if above
        If GetPersonajesCountByIDDatabase(UserList(userindex).AccountID) >= MAX_PERSONAJES Then
            Call CloseSocket(userindex)
            Exit Sub
        End If
        
        If Not ConnectNewUser(userindex, name, race, gender, Class, Head, Hogar) Then
            Call CloseSocket(userindex)
            Exit Sub
        End If
        
        
        Exit Sub
    
ErrHandler:
     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLoginNewChar", Erl)
End Sub
#End If

''
' Handles the "Talk" message.
'
' @param    UserIndex The index of the user sending the message.


Private Sub HandleTalk(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim chat As String
102         chat = Reader.ReadString8()
            
            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.Talk
            
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Talk", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
                        
            '[Consejeros & GMs]
104         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call LogGM(.Name, "Dijo: " & chat)
            End If
    
       
132         If .flags.Silenciado = 1 Then
134             Call WriteLocaleMsg(UserIndex, "110", e_FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
            Else
136             If LenB(chat) <> 0 Then
                    
                    ' WyroX: Foto-denuncias - Push message
                    Dim i As Long
140                 For i = 1 To UBound(.flags.ChatHistory) - 1
142                     .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                    Next
                    
144                 .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                
                
146                 If .flags.Muerto = 1 Then
148                     Call SendData(SendTarget.ToPCDeadArea, UserIndex, PrepareMessageChatOverHead(chat, .Char.charindex, CHAT_COLOR_DEAD_CHAR))
                      
                    
                    Else
                        If Trim(chat) = "" Then
                            .Counters.timeChat = 0
                        Else
                            .Counters.timeChat = 1 + Ceil((3000 + 60 * Len(chat)) / 1000)
                        End If
                        
150                     Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageChatOverHead(chat, .Char.charindex, .flags.ChatColor, , .Pos.X, .Pos.y))
                    End If

                End If

            End If

        End With
        
        Exit Sub
        
ErrHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTalk", Erl)
154

End Sub

''
' Handles the "Yell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim chat As String
102             chat = Reader.ReadString8()
        
104         If UserList(UserIndex).flags.Muerto = 1 Then
        
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Msg77=¡¡Estás muerto!!.
            
            Else

                '[Consejeros & GMs]
108             If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
110                 Call LogGM(.Name, "Grito: " & chat)
                End If
            
                'I see you....
112             If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            
114                 .flags.Oculto = 0
116                 .Counters.TiempoOculto = 0
                
118                 If .flags.Navegando = 1 Then
                    
                        'TODO: Revisar con WyroX
120                     If .clase = e_Class.Pirat Then
                    
                            ' Pierde la apariencia de fragata fantasmal
122                         Call EquiparBarco(UserIndex)
124                         ' Msg592=¡Has recuperado tu apariencia normal!
                            Call WriteLocaleMsg(UserIndex, "592", e_FontTypeNames.FONTTYPE_INFO)
126                         Call ChangeUserChar(UserIndex, .char.body, .char.head, .char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart)
128                         Call RefreshCharStatus(UserIndex)
                        End If
    
                    Else
    
130                     If .flags.invisible = 0 Then
132                         Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(userindex).Pos.X, UserList(userindex).Pos.y))
'Msg1115= ¡Has vuelto a ser visible!
Call WriteLocaleMsg(UserIndex, "1115", e_FontTypeNames.FONTTYPE_INFO)
    
                        End If
    
                    End If

                End If
            
136             If .flags.Silenciado = 1 Then
138                 Call WriteLocaleMsg(UserIndex, "110", e_FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
        
'Msg1116= Los administradores te han impedido hablar durante los proximos ¬1
Call WriteLocaleMsg(UserIndex, "1116", e_FontTypeNames.FONTTYPE_INFO, .flags.MinutosRestantes)
                Else

140                 If LenB(chat) <> 0 Then
                        ' WyroX: Foto-denuncias - Push message
                        Dim i As Long
144                     For i = 1 To UBound(.flags.ChatHistory) - 1
146                         .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                        Next
                    
148                     .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
                        
                        If Trim(chat) = "" Then
                            .Counters.timeChat = 0
                        Else
                            .Counters.timeChat = 1 + Ceil((3000 + 60 * Len(chat)) / 1000)
                        End If
150
                        Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageChatOverHead(chat, .Char.charindex, vbRed, , .Pos.X, .Pos.y))
               
                    End If

                End If

            End If

        End With
        
        Exit Sub
        
ErrHandler:

152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleYell", Erl)
154

End Sub

''
' Handles the "Whisper" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim chat            As String
            Dim targetCharIndex As String
            Dim targetUser      As t_UserReference

102         targetCharIndex = Reader.ReadString8()
104         chat = Reader.ReadString8()
    
106         If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(targetCharIndex)) < 0 Then Exit Sub
        
108         targetUser = NameIndex(targetCharIndex)
            If UserList(UserIndex).flags.Muerto = 1 Then
'Msg1117= No puedes susurrar estando muerto.
Call WriteLocaleMsg(UserIndex, "1117", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If Not IsValidUserRef(targetUser) Then
'Msg1118= El usuario esta muy lejos o desconectado.
Call WriteLocaleMsg(UserIndex, "1118", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
114         If EstaPCarea(userIndex, targetUser.ArrayIndex) Then
                If UserList(targetUser.ArrayIndex).flags.Muerto = 1 Then
'Msg1119= No puedes susurrar a un muerto.
Call WriteLocaleMsg(UserIndex, "1119", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
116             If LenB(chat) <> 0 Then
                    Dim i As Long
120                 For i = 1 To UBound(.flags.ChatHistory) - 1
122                     .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                    Next
124                 .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
126                 Call SendData(SendTarget.ToSuperioresArea, userIndex, PrepareMessageChatOverHead(chat, .Char.charindex, RGB(157, 226, 20), , .pos.x, .pos.y))
128                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareConsoleCharText(chat, RGB(157, 226, 20), UserList(UserIndex).name, UserList(UserIndex).Faccion.Status, UserList(UserIndex).flags.Privilegios))
130                 Call SendData(SendTarget.ToIndex, TargetUser.ArrayIndex, PrepareConsoleCharText(chat, RGB(157, 226, 20), UserList(UserIndex).name, UserList(UserIndex).Faccion.Status, UserList(UserIndex).flags.Privilegios))
                End If
            Else
'Msg1120= El usuario esta muy lejos o desconectado.
Call WriteLocaleMsg(UserIndex, "1120", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:

140     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWhisper", Erl)
142

End Sub

''
' Handles the "Walk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo HandleWalk_Err

        Dim Heading As e_Heading
    
100     With UserList(UserIndex)

102         Heading = Reader.ReadInt8()
            Dim PacketCount As Long
            PacketCount = Reader.ReadInt32
            
            If .flags.Muerto = 0 Then
                If .flags.Navegando Then
                    Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.Sailing), .PacketTimers(PacketNames.Sailing), .MacroIterations(PacketNames.Sailing), UserIndex, "Sailing", PacketTimerThreshold(PacketNames.Sailing), MacroIterations(PacketNames.Sailing))
                Else
                    Call verifyTimeStamp(PacketCount, .PacketCounters(PacketNames.Walk), .PacketTimers(PacketNames.Walk), .MacroIterations(PacketNames.Walk), UserIndex, "Walk", PacketTimerThreshold(PacketNames.Walk), MacroIterations(PacketNames.Walk))
                End If
            End If
            
            If .flags.PescandoEspecial Then
                .Stats.NumObj_PezEspecial = 0
                .flags.PescandoEspecial = False
            End If
            
104         If UserMod.CanMove(.flags, .Counters) Then
                
106             If .flags.Comerciando Or .flags.Crafteando <> 0 Then Exit Sub

108             If .flags.Meditando Then
            
                    'Stop meditating, next action will start movement.
110                 .flags.Meditando = False
112                 UserList(UserIndex).Char.FX = 0
114                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))

                End If
                
                Dim CurrentTick As Long
116                 CurrentTick = GetTickCount
            
                'Prevent SpeedHack (refactored by WyroX)
118             If Not EsGM(UserIndex) And .Char.speeding > 0 Then
                    Dim ElapsedTimeStep As Long, MinTimeStep As Long, DeltaStep As Single
120                 ElapsedTimeStep = CurrentTick - .Counters.LastStep
122                 MinTimeStep = .Intervals.Caminar / .Char.speeding
124                 DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep

126                 If DeltaStep > 0 Then
                
128                     .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep
                
130                     If .Counters.SpeedHackCounter > MaximoSpeedHack Then
                            'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Administración » Posible uso de SpeedHack del usuario " & .name & ".", e_FontTypeNames.FONTTYPE_SERVER))
132                         Call WritePosUpdate(UserIndex)
                            Exit Sub

                        End If

                    Else
                
134                     .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep * 5

136                     If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0

                    End If

                End If
            
                'Move user
138             If MoveUserChar(UserIndex, Heading) Then
            
                    ' Save current step for anti-sh
140                 .Counters.LastStep = CurrentTick
                
142                 If UserList(UserIndex).Grupo.EnGrupo Then
144                     Call CompartirUbicacion(UserIndex)

                    End If
    
                    'Stop resting if needed
146                 If .flags.Descansar Then
148                     .flags.Descansar = False
                        
150                     Call WriteRestOK(UserIndex)
'Msg1121= Has dejado de descansar.
Call WriteLocaleMsg(UserIndex, "1121", e_FontTypeNames.FONTTYPE_INFO)
152                     Call WriteLocaleMsg(UserIndex, "178", e_FontTypeNames.FONTTYPE_INFO)
    
                    End If
                        
154                 Call CancelExit(UserIndex)
                        
                    'Esta usando el /HOGAR, no se puede mover
156                 If .flags.Traveling = 1 Then
158                     .flags.Traveling = 0
160                     .Counters.goHome = 0
'Msg1122= Has cancelado el viaje a casa.
Call WriteLocaleMsg(UserIndex, "1122", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                    ' Si no pudo moverse
                Else
164                 .Counters.LastStep = 0
166                 Call WritePosUpdate(UserIndex)

                End If

            Else    'paralized

168             If Not .flags.UltimoMensaje = 1 Then
170                 .flags.UltimoMensaje = 1
'Msg1123= No podes moverte porque estas paralizado.
Call WriteLocaleMsg(UserIndex, "1123", e_FontTypeNames.FONTTYPE_INFO)
172                 Call WriteLocaleMsg(UserIndex, "54", e_FontTypeNames.FONTTYPE_INFO)
                End If
                Call WritePosUpdate(UserIndex)
            End If
            
            'Can't move while hidden except he is a thief
174         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
                
176             If .clase <> e_Class.Thief And .clase <> e_Class.Bandit Then
            
178                 .flags.Oculto = 0
180                 .Counters.TiempoOculto = 0
                
182                 If .flags.Navegando = 1 Then
                        
184                     If .clase = e_Class.Pirat Then
                            ' Pierde la apariencia de fragata fantasmal
186                         Call EquiparBarco(UserIndex)
188                         ' Msg592=¡Has recuperado tu apariencia normal!
                            Call WriteLocaleMsg(UserIndex, "592", e_FontTypeNames.FONTTYPE_INFO)
190                         Call ChangeUserChar(UserIndex, .char.body, .char.head, .char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart)
192                         Call RefreshCharStatus(UserIndex)
                        End If
    
                    Else
    
                        'If not under a spell effect, show char
194                     If .flags.invisible = 0 Then
'Msg1124= Has vuelto a ser visible.
Call WriteLocaleMsg(UserIndex, "1124", e_FontTypeNames.FONTTYPE_INFO)
196                         Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
198                         Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(userindex).Pos.X, UserList(userindex).Pos.y))

                        End If
    
                    End If
    
                End If
                
            End If

        End With

        Exit Sub

HandleWalk_Err:
200     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWalk", Erl)
202
        
End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo HandleRequestPositionUpdate_Err
        If UserList(userIndex).flags.SigueUsuario.ArrayIndex > 0 Then
            Call WritePosUpdateCharIndex(userIndex, UserList(UserList(userIndex).flags.SigueUsuario.ArrayIndex).pos.x, UserList(UserList(userIndex).flags.SigueUsuario.ArrayIndex).pos.y, UserList(UserList(userIndex).flags.SigueUsuario.ArrayIndex).Char.charindex)
        Else
100         Call WritePosUpdate(UserIndex)
        End If

  
        Exit Sub

HandleRequestPositionUpdate_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandlRequestPositionUpdate", Erl)
104
        
End Sub

''
' Handles the "Attack" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)
        
        On Error GoTo HandleAttack_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************
        
100     With UserList(UserIndex)
        
        
            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.Attack
            
            
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "Attack", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
            
            'If dead, can't attack
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                'Msg77=¡¡Estás muerto!!.
                Exit Sub

            End If
        
            'If equiped weapon is ranged, can't attack this way
106         If .Invent.WeaponEqpObjIndex > 0 Then

108             If ObjData(.Invent.WeaponEqpObjIndex).Proyectil = 1 And ObjData(.Invent.WeaponEqpObjIndex).Municion > 0 Then
'Msg1125= No podés usar así esta arma.
Call WriteLocaleMsg(UserIndex, "1125", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                If IsItemInCooldown(UserList(UserIndex), .invent.Object(.invent.WeaponEqpSlot)) Then
                    Exit Sub
                End If
            End If
        
112         If .Invent.HerramientaEqpObjIndex > 0 Then
114             ' Msg694=Para atacar debes desequipar la herramienta.
                Call WriteLocaleMsg(UserIndex, "694", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub

            End If
        
116         If UserList(UserIndex).flags.Meditando Then
118             UserList(UserIndex).flags.Meditando = False
120             UserList(UserIndex).Char.FX = 0
122             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0))

            End If
        
            'If exiting, cancel
124         Call CancelExit(UserIndex)
        
            'Attack!
126         Call UsuarioAtaca(UserIndex)
        End With

        Exit Sub

HandleAttack_Err:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAttack", Erl)
154
        
End Sub

''
' Handles the "PickUp" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)
        
        On Error GoTo HandlePickUp_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

            'If dead, it can't pick up objects
102         If .flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Lower rank administrators can't pick up items
106         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
108             ' Msg695=No podés tomar ningun objeto.
                Call WriteLocaleMsg(UserIndex, "695", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
110         Call PickObj(UserIndex)

        End With
        
        Exit Sub

HandlePickUp_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePickUp", Erl)
114
        
End Sub

''
' Handles the "SafeToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSafeToggle_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
            
            Dim cambiaSeguro As Boolean
            cambiaSeguro = False
            
            If .GuildIndex > 0 And (GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Or GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA) Then
                cambiaSeguro = False
            Else
                cambiaSeguro = True
            End If
             
            If cambiaSeguro Or .flags.Seguro = 0 Then
                If esCiudadano(UserIndex) Then
102                 If .flags.Seguro Then
104                     Call WriteSafeModeOff(UserIndex)
                    Else
106                     Call WriteSafeModeOn(UserIndex)
                    End If
                    
108                 .flags.Seguro = Not .flags.Seguro
                Else
                    ' Msg696=Solo los ciudadanos pueden cambiar el seguro.
                    Call WriteLocaleMsg(UserIndex, "696", e_FontTypeNames.FONTTYPE_TALK)
                End If
            Else
                ' Msg697=Debes abandonar el clan para poder sacar el seguro.
                Call WriteLocaleMsg(UserIndex, "697", e_FontTypeNames.FONTTYPE_TALK)
            End If

        End With

        Exit Sub

HandleSafeToggle_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSafeToggle", Erl)
112
        
End Sub

' Handles the "PartySafeToggle" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandlePartyToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandlePartyToggle_Err
        
        '***************************************************
        'Author: Rapsodius
        'Creation Date: 10/10/07
        '***************************************************
100     With UserList(UserIndex)
        
102         .flags.SeguroParty = Not .flags.SeguroParty
        
104         If .flags.SeguroParty Then
106             Call WritePartySafeOn(UserIndex)
            
            Else
108             Call WritePartySafeOff(UserIndex)

            End If

        End With

        Exit Sub

HandlePartyToggle_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePartyToggle", Erl)
112
        
End Sub

Private Sub HandleSeguroClan(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSeguroClan_Err

        '***************************************************
        'Author: Ladder
        'Date: 31/10/20
        '***************************************************
100     With UserList(UserIndex)

102         .flags.SeguroClan = Not .flags.SeguroClan

104         Call WriteClanSeguro(UserIndex, .flags.SeguroClan)

        End With

        Exit Sub

HandleSeguroClan_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSeguroClan", Erl)
108
        
End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo HandleRequestGuildLeaderInfo_Err

100     Call modGuilds.SendGuildLeaderInfo(UserIndex)

        Exit Sub

HandleRequestGuildLeaderInfo_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestGuildLeaderInfo", Erl)
104
        
End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        
        On Error GoTo HandleRequestAtributes_Err

100     Call WriteAttributes(UserIndex)

        Exit Sub

HandleRequestAtributes_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestAtributes", Erl)
104
        
End Sub

''
' Handles the "RequestSkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        
        On Error GoTo HandleRequestSkills_Err

100     Call WriteSendSkills(UserIndex)

        Exit Sub

HandleRequestSkills_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestSkills", Erl)
104
        
End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        
        On Error GoTo HandleRequestMiniStats_Err

100     Call WriteMiniStats(UserIndex)

        Exit Sub

HandleRequestMiniStats_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestMiniStats", Erl)
104
        
End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo HandleCommerceEnd_Err

        'User quits commerce mode
100     If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
102         If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose <> 0 Then
104             Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose, NO_3D_SOUND, NO_3D_SOUND)
            End If

        End If

106     UserList(UserIndex).flags.Comerciando = False

108     Call WriteCommerceEnd(UserIndex)
 
        Exit Sub

HandleCommerceEnd_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceEnd", Erl)
112
        
End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUserCommerceEnd_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

        
        
        
            'Quits commerce mode with user
            
102         If IsValidUserRef(.ComUsu.DestUsu) Then
                If UserList(.ComUsu.DestUsu.ArrayIndex).ComUsu.DestUsu.ArrayIndex = userIndex Then
104                 Call WriteConsoleMsg(.ComUsu.DestUsu.ArrayIndex, .name & " ha dejado de comerciar con vos.", e_FontTypeNames.FONTTYPE_TALK)
106                 Call FinComerciarUsu(.ComUsu.DestUsu.ArrayIndex)
                
                'Send data in the outgoing buffer of the other user

                End If
            End If
        
108         Call FinComerciarUsu(UserIndex)

        End With
        
        Exit Sub

HandleUserCommerceEnd_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceEnd", Erl)
112
        
End Sub

''
' Handles the "BankEnd" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)
        
        On Error GoTo HandleBankEnd_Err
       
100      With UserList(UserIndex)
            If .flags.Comerciando Then
102             .flags.Comerciando = False
104             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("171", UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
106             Call WriteBankEnd(UserIndex)
            End If
        End With
        
        Exit Sub

HandleBankEnd_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankEnd", Erl)
110
        
End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        
        On Error GoTo HandleUserCommerceOk_Err

        'Trade accepted
100     Call AceptarComercioUsu(UserIndex)
        
        Exit Sub

HandleUserCommerceOk_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOk", Erl)
104
        
End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUserCommerceReject_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim otherUser As Integer
    
100     With UserList(UserIndex)

102         otherUser = .ComUsu.DestUsu.ArrayIndex
        
            'Offer rejected
104         If otherUser > 0 Then
106             If UserList(otherUser).flags.UserLogged Then
108                 Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", e_FontTypeNames.FONTTYPE_TALK)
110                 Call FinComerciarUsu(otherUser)
                
                    'Send data in the outgoing buffer of the other user

                End If

            End If
        
112         ' Msg698=Has rechazado la oferta del otro usuario.
            Call WriteLocaleMsg(UserIndex, "698", e_FontTypeNames.FONTTYPE_TALK)
114         Call FinComerciarUsu(UserIndex)

        End With
        
        Exit Sub

HandleUserCommerceReject_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceReject", Erl)
118
        
End Sub

''
' Handles the "Drop" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)
        
        On Error GoTo HandleDrop_Err
        
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 07/25/09
        '07/25/09: Marco - Agregue un checkeo para patear a los usuarios que tiran items mientras comercian.
        '***************************************************

        Dim Slot   As Byte
        Dim amount As Long
    
100     With UserList(UserIndex)

102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt32()
            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.Drop
            
            
            'If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), userindex, "Drop", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
            
106         If Not IntervaloPermiteTirar(UserIndex) Then Exit Sub
            If .flags.PescandoEspecial = True Then Exit Sub

108         If amount <= 0 Then Exit Sub

            'low rank admins can't drop item. Neither can the dead nor those sailing or riding a horse.
110         If .flags.Muerto = 1 Then Exit Sub
                      
            'If the user is trading, he can't drop items => He's cheating, we kick him.
112         If .flags.Comerciando Then Exit Sub
    
            'Si esta navegando y no es pirata, no dejamos tirar items al agua.
114         'If .flags.Navegando = 1 And Not .clase = e_Class.Pirat Then
'Msg1126= Solo los Piratas pueden tirar items en altamar
Call WriteLocaleMsg(userindex, "1126", e_FontTypeNames.FONTTYPE_INFO)
              '  Exit Sub

            'End If
            
118         If .flags.Montado = 1 Then
120             ' Msg699=Debes descender de tu montura para dejar objetos en el suelo.
                Call WriteLocaleMsg(UserIndex, "699", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            If UserList(UserIndex).flags.SigueUsuario.ArrayIndex > 0 Then
                ' Msg700=No podes tirar items cuando estas siguiendo a alguien.
                Call WriteLocaleMsg(UserIndex, "700", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            'Are we dropping gold or other items??
122         If Slot = FLAGORO Then
                If amount > 100000 Then amount = 100000
124             Call TirarOro(amount, UserIndex)
            
            Else
                If Slot <= getMaxInventorySlots(UserIndex) Then
                '04-05-08 Ladder
126                 If (.flags.Privilegios And e_PlayerType.Admin) <> 16 Then
128                     If EsNewbie(UserIndex) And ObjData(.Invent.Object(Slot).ObjIndex).Newbie = 1 Then
130                         ' Msg701=No se pueden tirar los objetos Newbies.
                            Call WriteLocaleMsg(UserIndex, "701", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                
132                     If ObjData(.Invent.Object(Slot).ObjIndex).Intirable = 1 And Not EsGM(UserIndex) Then
134                         ' Msg702=Acción no permitida.
                            Call WriteLocaleMsg(UserIndex, "702", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
136                     ElseIf ObjData(.Invent.Object(Slot).ObjIndex).Intirable = 1 And EsGM(UserIndex) Then
138                         If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
140                             If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
142                             Call DropObj(UserIndex, Slot, amount, .Pos.map, .Pos.X, .Pos.Y)
                            End If
                            Exit Sub
                        End If
                    
144                     If ObjData(.Invent.Object(Slot).ObjIndex).Instransferible = 1 Then
146                         ' Msg702=Acción no permitida.
                            Call WriteLocaleMsg(UserIndex, "702", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                
    
                    End If
        
148                 If ObjData(.Invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otBarcos And UserList(UserIndex).flags.Navegando Then
150                     ' Msg703=Para tirar la barca deberias estar en tierra firme.
                        Call WriteLocaleMsg(UserIndex, "703", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                Else
                    'ver de banear al usuario
                    'Call BanearIP(0, UserList(UserIndex).name, UserList(UserIndex).IP, UserList(UserIndex).Cuenta)
                    Call LogEdicionPaquete("El usuario " & UserList(UserIndex).name & " editó el slot del inventario | Valor: " & Slot & ".")
                End If
        
                '04-05-08 Ladder
        
                'Only drop valid slots
152             If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
            
154                 If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

156                 Call DropObj(UserIndex, Slot, amount, .Pos.Map, .Pos.X, .Pos.Y)

                End If

            End If

        End With
        
        Exit Sub

HandleDrop_Err:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDrop", Erl)
160
        
End Sub
Public Function verifyTimeStamp(ByVal ActualCount As Long, ByRef LastCount As Long, ByRef LastTick As Long, ByRef Iterations, ByVal UserIndex As Integer, ByVal PacketName As String, Optional ByVal DeltaThreshold As Long = 100, Optional ByVal MaxIterations As Long = 5, Optional ByVal CloseClient As Boolean = False) As Boolean
    
    Dim Ticks As Long, Delta As Long
    Ticks = GetTickCount
    
    Delta = (Ticks - LastTick)
    LastTick = Ticks

    'Controlamos secuencia para ver que no haya paquetes duplicados.
    If ActualCount <= LastCount Then
        Call SendData(SendTarget.ToGM, UserIndex, PrepareMessageConsoleMsg("Paquete grabado: " & PacketName & " | Cuenta: " & UserList(UserIndex).Cuenta & " | Ip: " & UserList(UserIndex).ConnectionDetails.IP & " (Baneado automaticamente)", e_FontTypeNames.FONTTYPE_INFOBOLD))
        Call LogEdicionPaquete("El usuario " & UserList(UserIndex).name & " editó el paquete " & PacketName & ".")
        LastCount = ActualCount
        Call CloseSocket(UserIndex)
        Exit Function
    End If
    
    'controlamos speedhack/macro
    If Delta < DeltaThreshold Then
        Iterations = Iterations + 1
        If Iterations >= MaxIterations Then
            'Call WriteShowMessageBox(UserIndex, "Relajate andá a tomarte un té con Gulfas.")
            verifyTimeStamp = False
            'Call LogMacroServidor("El usuario " & UserList(UserIndex).name & " iteró el paquete " & PacketName & " " & MaxIterations & " veces.")
            Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("Control de macro---> El usuario " & UserList(UserIndex).name & "| Revisar --> " & PacketName & " (Envíos: " & Iterations & ").", e_FontTypeNames.FONTTYPE_INFOBOLD))
            'Call WriteCerrarleCliente(UserIndex)
            'Call CloseSocket(UserIndex)
            LastCount = ActualCount
            Iterations = 0
            Debug.Print "CIERRO CLIENTE"
        End If
        'Exit Function
    Else
        Iterations = 0
    End If
        
    verifyTimeStamp = True
    LastCount = ActualCount
End Function


''
' Handles the "CastSpell" message.
' @param    UserIndex The index of the user sending the message.
Private Sub HandleCastSpell(ByVal UserIndex As Integer)
    On Error GoTo HandleCastSpell_Err
        Dim Spell As Byte
        Spell = Reader.ReadInt8()
        Dim PacketCounter As Long
        PacketCounter = Reader.ReadInt32
        Dim Packet_ID As Long
        Packet_ID = PacketNames.CastSpell
        Call UseSpellSlot(UserIndex, Spell)
        Exit Sub
HandleCastSpell_Err:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCastSpell", Erl)
End Sub

''
' Handles the "LeftClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)
        
        On Error GoTo HandleLeftClick_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)

            Dim X As Byte
            Dim Y As Byte
        
102         X = Reader.ReadInt8()
104         Y = Reader.ReadInt8()
            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.LeftClick

            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "LeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
            
106         Call LookatTile(UserIndex, .Pos.Map, X, Y)

        End With

        Exit Sub

HandleLeftClick_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLeftClick", Erl)
110
        
End Sub

''
' Handles the "DoubleClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleDoubleClick(ByVal UserIndex As Integer)
        
        On Error GoTo HandleDoubleClick_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)

            Dim X As Byte
            Dim Y As Byte
        
102         X = Reader.ReadInt8()
104         Y = Reader.ReadInt8()
        
106         Call Accion(UserIndex, .Pos.Map, X, Y)

        End With
        
        Exit Sub

HandleDoubleClick_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoubleClick", Erl)
110
        
End Sub

Private Sub HandleWork(ByVal UserIndex As Integer)
        
        On Error GoTo HandleWork_Err

100     With UserList(UserIndex)

            Dim Skill As e_Skill
102             Skill = Reader.ReadInt8()
            
            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32

104         If UserList(UserIndex).flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'If exiting, cancel
108         Call CancelExit(UserIndex)
        
110         Select Case Skill

                Case Robar, Magia, Domar
112                 Call WriteWorkRequestTarget(UserIndex, Skill)

114             Case Ocultarse
                    If Not verifyTimeStamp(PacketCounter, .PacketCounters(PacketNames.Hide), .PacketTimers(PacketNames.Hide), .MacroIterations(PacketNames.Hide), _
                                            UserIndex, "Ocultar", PacketTimerThreshold(PacketNames.Hide), MacroIterations(PacketNames.Hide)) Then Exit Sub
116                 If .flags.Montado = 1 Then

                        '[CDT 17-02-2004]
118                     If Not .flags.UltimoMensaje = 3 Then
120                         ' Msg704=No podés ocultarte si estás montado.
                            Call WriteLocaleMsg(UserIndex, "704", e_FontTypeNames.FONTTYPE_INFO)
122                         .flags.UltimoMensaje = 3

                        End If

                        '[/CDT]
                        Exit Sub

                    End If

124                 If .flags.Oculto = 1 Then

                        '[CDT 17-02-2004]
126                     If Not .flags.UltimoMensaje = 2 Then
128                         Call WriteLocaleMsg(UserIndex, "55", e_FontTypeNames.FONTTYPE_INFO)
'Msg1127= Ya estás oculto.
Call WriteLocaleMsg(UserIndex, "1127", e_FontTypeNames.FONTTYPE_INFO)
130                         .flags.UltimoMensaje = 2

                        End If

                        '[/CDT]
                        Exit Sub

                    End If
                    
132                 If .flags.EnReto Then
134                     ' Msg705=No podés ocultarte durante un reto.
                        Call WriteLocaleMsg(UserIndex, "705", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
136                 If .flags.EnConsulta Then
138                     ' Msg706=No podés ocultarte si estas en consulta.
                        Call WriteLocaleMsg(UserIndex, "706", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
                    
                    If .flags.invisible Then
139                     ' Msg707=No podés ocultarte si estás invisible.
                        Call WriteLocaleMsg(UserIndex, "707", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
                    
140                 If MapInfo(.Pos.Map).SinInviOcul Then
142                     ' Msg708=Una fuerza divina te impide ocultarte en esta zona.
                        Call WriteLocaleMsg(UserIndex, "708", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                    
144                 Call DoOcultarse(UserIndex)

            End Select

        End With
        
        Exit Sub

HandleWork_Err:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWork", Erl)
148
        
End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUseSpellMacro_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)
#If STRESSER = 1 Then
    Exit Sub
#End If
102         Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos", e_FontTypeNames.FONTTYPE_VENENO))
104         Call WriteShowMessageBox(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros")
        
106         Call CloseSocket(UserIndex)

        End With
        
        Exit Sub

HandleUseSpellMacro_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseSpellMacro", Erl)
110
        
End Sub

''
' Handles the "UseItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo HandleUseItem_Err
    
100     With UserList(UserIndex)

            Dim Slot As Byte
102         Slot = Reader.ReadInt8()

            Dim DesdeInventario As Boolean
            DesdeInventario = Reader.ReadInt8
            
            If Not DesdeInventario Then
                Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg("El usuario " & .name & " está tomando pociones con click estando en hechizos... raaaaaro, poleeeeemico. BAN?", e_FontTypeNames.FONTTYPE_INFOBOLD))
            End If
            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
            
            Dim Packet_ID As Long
            Packet_ID = PacketNames.UseItem
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItem", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
          '  Debug.Print "LLEGA PAQUETE"
104         If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
106             If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

108             Call UseInvItem(UserIndex, Slot, 1)
                
            End If

        End With

        Exit Sub

HandleUseItem_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseItem", Erl)
112
        
End Sub

''
' Handles the "UseItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUseItemU(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUseItemU_Err
    
100     With UserList(UserIndex)

            Dim Slot As Byte
102         Slot = Reader.ReadInt8()

            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
            
            Dim Packet_ID As Long
            Packet_ID = PacketNames.UseItemU
            
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "UseItemU", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
            
104         If Slot <= UserList(UserIndex).CurrentInventorySlots And Slot > 0 Then
106             If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

108             Call UseInvItem(UserIndex, Slot, 0)
                
            End If

        End With

        Exit Sub

HandleUseItemU_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseItemU", Erl)
112
        
End Sub

''
' Handles the "CraftBlacksmith" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftBlacksmith(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftBlacksmith_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

            Dim Item As Integer
102             Item = Reader.ReadInt16()
        
104         If Item < 1 Then Exit Sub
        
            ' If ObjData(Item).SkHerreria = 0 Then Exit Sub
        
106         Call HerreroConstruirItem(UserIndex, Item)

        Exit Sub

HandleCraftBlacksmith_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftBlacksmith", Erl)
110
        
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftCarpenter_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

            Dim Item As Integer
102         Item = Reader.ReadInt16()
            Dim Cantidad As Long
            Cantidad = Reader.ReadInt32()
        
104         If Item = 0 Then Exit Sub

            'Valido que haya puesto una cantidad > 0
            If Cantidad > 0 Then
                UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
                UserList(UserIndex).Trabajo.TargetSkill = e_Skill.Carpinteria
                UserList(UserIndex).Trabajo.Cantidad = Cantidad
                UserList(UserIndex).Trabajo.Item = Item
                Call WriteMacroTrabajoToggle(UserIndex, True)
            Else
            
            End If
            

'106         Call CarpinteroConstruirItem(UserIndex, Item, Cantidad)

        Exit Sub

HandleCraftCarpenter_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftCarpenter", Erl)
110
        
End Sub

Private Sub HandleCraftAlquimia(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftAlquimia_Err
        
        
            Dim Item As Integer
        
            Item = Reader.ReadInt16()
        
110         If Item < 1 Then Exit Sub
            

112         Call AlquimistaConstruirItem(UserIndex, Item)

        
        Exit Sub

HandleCraftAlquimia_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftAlquimia", Erl)
108
        
End Sub

Private Sub HandleCraftSastre(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCraftSastre_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

            Dim Item As Integer
102             Item = Reader.ReadInt16()
        
104         If Item < 1 Then Exit Sub

106         Call SastreConstruirItem(UserIndex, Item)

        Exit Sub

HandleCraftSastre_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftSastre", Erl)
110
        
End Sub
''
' Handles the "WorkLeftClick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)
        
        On Error GoTo HandleWorkLeftClick_Err
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim X        As Byte
            Dim Y        As Byte

            Dim Skill    As e_Skill
            Dim DummyInt As Integer

            Dim tU       As Integer   'Target user
            Dim tN       As Integer   'Target NPC
        
102         X = Reader.ReadInt8()
104         Y = Reader.ReadInt8()
            
106         Skill = Reader.ReadInt8()

            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.WorkLeftClick

            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "WorkLeftClick", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub

            .Trabajo.Target_X = X
            .Trabajo.Target_Y = Y
            .Trabajo.TargetSkill = Skill
            
108         If .flags.Muerto = 1 Or .flags.Descansar Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub
            If UserMod.IsStun(.flags, .Counters) Then Exit Sub
110         If Not InRangoVision(UserIndex, X, Y) Then
112             Call WritePosUpdate(UserIndex)
                Exit Sub

            End If
            
114         If .flags.Meditando Then
116             .flags.Meditando = False
118             .Char.FX = 0
120             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
            End If
        
            'If exiting, cancel
122         Call CancelExit(UserIndex)
            
124         Select Case Skill

                Dim consumirMunicion As Boolean

                Case e_Skill.Proyectiles
                    Dim WeaponData As t_ObjData
                    Dim ProjectileType As Byte
                    'Check attack interval
126                 If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub

                    'Check Magic interval
128                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub

                    'Check bow's interval
130                 If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                    'Make sure the item is valid and there is ammo equipped.
132                 With .Invent
                        If .WeaponEqpObjIndex < 1 Then Exit Sub
                        WeaponData = ObjData(.WeaponEqpObjIndex)

                        If IsItemInCooldown(UserList(UserIndex), .Object(.WeaponEqpSlot)) Then Exit Sub
                        ProjectileType = GetProjectileView(UserList(UserIndex))
                        If WeaponData.Proyectil = 1 And WeaponData.Municion = 0 Then
                            DummyInt = 0
                        ElseIf .WeaponEqpObjIndex = 0 Then
136                         DummyInt = 1
138                     ElseIf .WeaponEqpSlot < 1 Or .WeaponEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
140                         DummyInt = 1
142                     ElseIf .MunicionEqpSlot < 1 Or .MunicionEqpSlot > UserList(UserIndex).CurrentInventorySlots Then
144                         DummyInt = 1
146                     ElseIf .MunicionEqpObjIndex = 0 Then
148                         DummyInt = 1
150                     ElseIf ObjData(.WeaponEqpObjIndex).Proyectil <> 1 Then
152                         DummyInt = 2
154                     ElseIf ObjData(.MunicionEqpObjIndex).OBJType <> e_OBJType.otFlechas Then
156                         DummyInt = 1
158                     ElseIf .Object(.MunicionEqpSlot).amount < 1 Then
160                         DummyInt = 1
                        ElseIf ObjData(.MunicionEqpObjIndex).Subtipo <> WeaponData.Municion Then
161                         DummyInt = 1

                        End If
                    
162                     If DummyInt <> 0 Then
164                         If DummyInt = 1 Then
166                             ' Msg709=No tenés municiones.
                                Call WriteLocaleMsg(UserIndex, "709", e_FontTypeNames.FONTTYPE_INFO)
                            End If
168                         Call Desequipar(UserIndex, .MunicionEqpSlot)
170                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub
                        End If
                    End With
                
                    'Quitamos stamina
172                 If .Stats.MinSta >= 10 Then
174                     Call QuitarSta(UserIndex, RandomNumber(1, 10))
                    Else
180                     Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
'Msg1128= Estís muy cansado para luchar.
Call WriteLocaleMsg(UserIndex, "1128", e_FontTypeNames.FONTTYPE_INFO)
182                     Call WriteWorkRequestTarget(UserIndex, 0)
                        Exit Sub
                    End If
                
184                 Call LookatTile(UserIndex, .Pos.Map, X, Y)
186                 tU = .flags.targetUser.ArrayIndex
188                 tN = .flags.TargetNPC.ArrayIndex
190                 consumirMunicion = False
                    'Validate target
192                 If IsValidUserRef(.flags.targetUser) Then
                        'Only allow to atack if the other one can retaliate (can see us)
194                     If Abs(UserList(tU).Pos.Y - .Pos.Y) > RANGO_VISION_Y Then
                            ' Msg8=Estas demasiado lejos.
196                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
198                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub
                        End If
                    
                        'Prevent from hitting self
200                     If tU = UserIndex Then
202                         ' Msg710=¡No podés atacarte a vos mismo!
                            Call WriteLocaleMsg(UserIndex, "710", e_FontTypeNames.FONTTYPE_INFO)
204                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub
                        End If
                    
                        'Attack!
206                     If Not PuedeAtacar(UserIndex, tU) Then Exit Sub 'TODO: Por ahora pongo esto para solucionar lo anterior.
                    
                        Dim backup    As Byte
                        Dim envie     As Boolean
                        Dim Particula As Integer
                        Dim Tiempo    As Long
                        
                        If .flags.invisible > 0 Then
                            If IsFeatureEnabled("remove-inv-on-attack") Then
                                Call RemoveUserInvisibility(UserIndex)
                            End If
                        End If
208                     Call UsuarioAtacaUsuario(UserIndex, tU, Ranged)
                        Dim FX As Integer
                        If .Invent.MunicionEqpObjIndex Then
                            FX = ObjData(.Invent.MunicionEqpObjIndex).CreaFX
                        End If
210                     If FX <> 0 Then
                            UserList(tU).Counters.timeFx = 3
212                         Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageCreateFX(UserList(tU).Char.charindex, FX, 0, UserList(tU).Pos.X, UserList(tU).Pos.y))
                        End If
                        If ProjectileType > 0 And (.flags.Oculto = 0 Or Not MapInfo(.pos.Map).KeepInviOnAttack) Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y, X, y, ProjectileType))
                        End If
                        'Si no es GM invisible, le envio el movimiento del arma.
                        If UserList(UserIndex).flags.AdminInvisible = 0 Then
                            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.charindex, 1))
                        End If
                    
214                     If .Invent.MunicionEqpObjIndex > 0 Then
215                         If ObjData(.Invent.MunicionEqpObjIndex).CreaParticula <> "" Then
216                             Particula = val(ReadField(1, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
218                             Tiempo = val(ReadField(2, ObjData(.Invent.MunicionEqpObjIndex).CreaParticula, Asc(":")))
                                UserList(tU).Counters.timeFx = 3
220                             Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageParticleFX(UserList(tU).Char.charindex, Particula, Tiempo, False, , UserList(tU).Pos.X, UserList(tU).Pos.y))
                            End If
                        End If
                    
222                     consumirMunicion = True
                    
224                 ElseIf tN > 0 Then

                        'Only allow to atack if the other one can retaliate (can see us)
226                     If Abs(NpcList(tN).Pos.Y - .Pos.Y) > RANGO_VISION_Y And Abs(NpcList(tN).Pos.X - .Pos.X) > RANGO_VISION_X Then
                            ' Msg8=Estas demasiado lejos.
228                         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
230                         Call WriteWorkRequestTarget(UserIndex, 0)
                            Exit Sub
                        End If
                    
                        'Is it attackable???
232                     If NpcList(tN).Attackable <> 0 Then


                            Dim UserAttackInteractionResult As t_AttackInteractionResult
                            UserAttackInteractionResult = UserCanAttackNpc(UserIndex, tN)
                            Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
                            If UserAttackInteractionResult.CanAttack Then
                                If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
236                             Call UsuarioAtacaNpc(UserIndex, tN, Ranged)
238                             consumirMunicion = True
                                If ProjectileType > 0 And .flags.Oculto = 0 Then
                                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y, X, y, ProjectileType))
                                End If
                                'Si no es GM invisible, le envio el movimiento del arma.
                                If UserList(UserIndex).flags.AdminInvisible = 0 Then
                                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(UserList(UserIndex).Char.charindex, 1))
                                End If
                            Else
240                             consumirMunicion = False
                            End If
                        End If
                    End If
                    
242                 With .Invent
                        If WeaponData.Proyectil = 1 And WeaponData.Municion > 0 Then
244                         DummyInt = .MunicionEqpSlot
                            If ObjData(.WeaponEqpObjIndex).CreaWav > 0 Then
                                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(ObjData(.WeaponEqpObjIndex).CreaWav, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
                            End If
                            If DummyInt <> 0 Then
                                'Take 1 arrow away - we do it AFTER hitting, since if Ammo Slot is 0 it gives a rt9 and kicks players
246                             If consumirMunicion Then
248                                 Call QuitarUserInvItem(UserIndex, DummyInt, 1)
                                End If
                            
250                             If .Object(DummyInt).amount > 0 Then
                                    'QuitarUserInvItem unequipps the ammo, so we equip it again
252                                 .MunicionEqpSlot = DummyInt
254                                 .MunicionEqpObjIndex = .Object(DummyInt).objIndex
256                                 .Object(DummyInt).Equipped = 1
                                Else
258                                 .MunicionEqpSlot = 0
260                                 .MunicionEqpObjIndex = 0
                                End If
262                             Call UpdateUserInv(False, UserIndex, DummyInt)
                            End If
                        ElseIf consumirMunicion Then
                            Call UpdateCd(UserIndex, WeaponData.CdType)
                        End If
                    End With
                    '-----------------------------------
            
264             Case e_Skill.Magia
                    'Target whatever is in that tile
266                 Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                    'If it's outside range log it and exit
268                 If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
270                     Call LogSecurity("Ataque fuera de rango de " & .name & "(" & .pos.Map & "/" & .pos.x & "/" & .pos.y & ") ip: " & .ConnectionDetails.IP & " a la posicion (" & .pos.Map & "/" & x & "/" & y & ")")
                        Exit Sub
                    End If
                
                    'Check bow's interval
272                 If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                    'Check attack-spell interval
274                 If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
                    'Check Magic interval
276                 If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
                
                    'Check intervals and cast
278                 If .flags.Hechizo > 0 Then
                        .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
280                     Call LanzarHechizo(.flags.Hechizo, UserIndex)
                        If IsValidUserRef(.flags.GMMeSigue) Then
                            Call WriteNofiticarClienteCasteo(.flags.GMMeSigue.ArrayIndex, 0)
                        End If
282                     .flags.Hechizo = 0
                    Else
284                     ' Msg587=¡Primero selecciona el hechizo que quieres lanzar!
                        Call WriteLocaleMsg(UserIndex, "587", e_FontTypeNames.FONTTYPE_INFO)

                    End If
            
286             Case e_Skill.Pescar
                    If .Counters.Trabajando = 0 And .Counters.LastTrabajo = 0 Then
                        If .invent.HerramientaEqpSlot = 0 Then Exit Sub
                        If IsItemInCooldown(UserList(UserIndex), .invent.Object(.invent.HerramientaEqpSlot)) Then Exit Sub
                        Call LookatTile(UserIndex, .pos.map, X, y)
                        Call FishOrThrowNet(UserIndex)
                    End If
348             Case e_Skill.Talar
                    If .Counters.Trabajando = 0 And .Counters.LastTrabajo = 0 Then
                        Call Trabajar(UserIndex, e_Skill.Talar)
                    End If
                    
400             Case e_Skill.Alquimia
            
402                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
404                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> e_OBJType.otHerramientas Then Exit Sub
                    
                    'Check interval
406                 If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

408                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case 3  ' Herramientas de Alquimia - Tijeras

410                         If MapInfo(UserList(UserIndex).Pos.Map).Seguro = 1 Then
412                             Call WriteWorkRequestTarget(UserIndex, 0)
414                             ' Msg711=Esta prohibido cortar raices en las ciudades.
                                Call WriteLocaleMsg(UserIndex, "711", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub

                            End If
                            
416                         If MapData(.Pos.Map, X, Y).ObjInfo.amount <= 0 Then
418                             ' Msg712=El árbol ya no te puede entregar mas raices.
                                Call WriteLocaleMsg(UserIndex, "712", e_FontTypeNames.FONTTYPE_INFO)
420                             Call WriteWorkRequestTarget(UserIndex, 0)
422                             Call WriteMacroTrabajoToggle(UserIndex, False)
                                Exit Sub

                            End If
                
424                         DummyInt = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                            
426                         If DummyInt > 0 Then
                            
428                             If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
430                                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
'Msg1129= Estas demasiado lejos.
Call WriteLocaleMsg(UserIndex, "1129", e_FontTypeNames.FONTTYPE_INFO)
432                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
434                             If .Pos.X = X And .Pos.Y = Y Then
436                                 ' Msg713=No podés quitar raices allí.
                                    Call WriteLocaleMsg(UserIndex, "713", e_FontTypeNames.FONTTYPE_INFO)
438                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
                                '¡Hay un arbol donde clickeo?
440                             If ObjData(DummyInt).OBJType = e_OBJType.otPlantas Then
442                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TIJERAS, .Pos.X, .Pos.y))
444                                 Call DoRaices(UserIndex, X, Y)

                                End If

                            Else
446                             ' Msg604=No podés quitar raices allí.
                                Call WriteLocaleMsg(UserIndex, "604", e_FontTypeNames.FONTTYPE_INFO)
448                             Call WriteWorkRequestTarget(UserIndex, 0)
450                             Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If
                
                    End Select
                
452             Case e_Skill.Mineria
                    If .Counters.Trabajando = 0 And .Counters.LastTrabajo = 0 Then
                        Call Trabajar(UserIndex, e_Skill.Mineria)
                    End If
500             Case e_Skill.Robar

                    'Does the map allow us to steal here?
502                 If MapInfo(.Pos.Map).Seguro = 0 Then
                    
                        'Check interval
504                     If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub
                    
                        'Target whatever is in that tile
506                     Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
508                     tU = .flags.targetUser.ArrayIndex
                    
510                     If IsValidUserRef(.flags.targetUser) And tU <> userIndex Then

                            'Can't steal administrative players
512                         If UserList(tU).flags.Privilegios And e_PlayerType.user Then
514                             If UserList(tU).flags.Muerto = 0 Then
                                    Dim DistanciaMaxima As Integer

516                                 If .clase = e_Class.Thief Then
518                                     DistanciaMaxima = 1
                                    Else
520                                     DistanciaMaxima = 1

                                    End If

522                                 If Abs(.Pos.X - UserList(tU).Pos.X) + Abs(.Pos.Y - UserList(tU).Pos.Y) > DistanciaMaxima Then
524                                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
'Msg1130= Estís demasiado lejos.
Call WriteLocaleMsg(UserIndex, "1130", e_FontTypeNames.FONTTYPE_INFO)
526                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
                                    '17/09/02
                                    'Check the trigger
528                                 If MapData(UserList(tU).Pos.Map, UserList(tU).Pos.X, UserList(tU).Pos.Y).trigger = e_Trigger.ZonaSegura Then
530                                     ' Msg714=No podés robar aquí.
                                        Call WriteLocaleMsg(UserIndex, "714", e_FontTypeNames.FONTTYPE_WARNING)
532                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
534                                 If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = e_Trigger.ZonaSegura Then
536                                     ' Msg714=No podés robar aquí.
                                        Call WriteLocaleMsg(UserIndex, "714", e_FontTypeNames.FONTTYPE_WARNING)
538                                     Call WriteWorkRequestTarget(UserIndex, 0)
                                        Exit Sub

                                    End If
                                 
540                                 Call DoRobar(UserIndex, tU)

                                End If

                            End If

                        Else
542                         ' Msg715=No a quien robarle!
                            Call WriteLocaleMsg(UserIndex, "715", e_FontTypeNames.FONTTYPE_INFO)
544                         Call WriteWorkRequestTarget(UserIndex, 0)

                        End If

                    Else
546                     ' Msg716=¡No podés robar en zonas seguras!
                        Call WriteLocaleMsg(UserIndex, "716", e_FontTypeNames.FONTTYPE_INFO)
548                     Call WriteWorkRequestTarget(UserIndex, 0)

                    End If
                    
550             Case e_Skill.Domar
552                 Call LookatTile(UserIndex, .Pos.Map, X, Y)
556                 If IsValidNpcRef(.flags.TargetNPC) Then
                        tN = .flags.TargetNPC.ArrayIndex
558                     If NpcList(tN).flags.Domable > 0 Then
560                         If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 4 Then
562                             ' Msg8=Estas demasiado lejos.
                                Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
564                         If GetOwnedBy(tN) <> 0 Then
566                             ' Msg717=No puedes domar una criatura que esta luchando con un jugador.
                                Call WriteLocaleMsg(UserIndex, "717", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
568                         Call DoDomar(UserIndex, tN)
                        Else
570                         ' Msg718=No puedes domar a esa criatura.
                            Call WriteLocaleMsg(UserIndex, "718", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
572                     ' Msg719=No hay ninguna criatura alli!
                        Call WriteLocaleMsg(UserIndex, "719", e_FontTypeNames.FONTTYPE_INFO)
                    End If
               
574             Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            
                    'Check interval
576                 If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
                
578                 Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                    'Check there is a proper item there
580                 If .flags.TargetObj > 0 Then
582                     If ObjData(.flags.TargetObj).OBJType = e_OBJType.otFragua Then

                            'Validate other items
584                         If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > UserList(UserIndex).CurrentInventorySlots Then
                                Exit Sub

                            End If
                        
                            ''chequeamos que no se zarpe duplicando oro
586                         If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
588                             If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).amount = 0 Then
590                                 ' Msg605=No tienes más minerales
                                    Call WriteLocaleMsg(UserIndex, "605", e_FontTypeNames.FONTTYPE_INFO)
592                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                            
                                ''FUISTE
594                             Call WriteShowMessageBox(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                            
596                             Call CloseSocket(UserIndex)
                                Exit Sub

                            End If
                        
598                         Call FundirMineral(UserIndex)
                        
                        Else
                    
600                         ' Msg606=Ahí no hay ninguna fragua.
                            Call WriteLocaleMsg(UserIndex, "606", e_FontTypeNames.FONTTYPE_INFO)
602                         Call WriteWorkRequestTarget(UserIndex, 0)

604                         If UserList(UserIndex).Counters.Trabajando > 1 Then
606                             Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If

                        End If

                    Else
                
608                     ' Msg606=Ahí no hay ninguna fragua.
                        Call WriteLocaleMsg(UserIndex, "606", e_FontTypeNames.FONTTYPE_INFO)
610                     Call WriteWorkRequestTarget(UserIndex, 0)

612                     If UserList(UserIndex).Counters.Trabajando > 1 Then
614                         Call WriteMacroTrabajoToggle(UserIndex, False)

                        End If

                    End If

616             Case e_Skill.Grupo

617                 Call LookatTile(UserIndex, .Pos.map, X, Y)

                    'Target whatever is in that tile
618                 tU = .flags.targetUser.ArrayIndex
                    
620                 If IsValidUserRef(.flags.targetUser) And tU <> userIndex Then
622                     If UserList(UserIndex).Grupo.EnGrupo = False Then
624                         If UserList(tU).flags.Muerto = 0 Then
626                             If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 8 Then
628                                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
630                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub
                                End If
632                             If UserList(UserIndex).Grupo.CantidadMiembros = 0 Then
634                                 Call SetUserRef(UserList(userIndex).Grupo.Lider, userIndex)
636                                 Call SetUserRef(UserList(userIndex).Grupo.Miembros(1), userIndex)
638                                 UserList(UserIndex).Grupo.CantidadMiembros = 1
640                                 Call InvitarMiembro(UserIndex, tU)
                                Else
642                                 Call SetUserRef(UserList(userIndex).Grupo.Lider, userIndex)
644                                 Call InvitarMiembro(UserIndex, tU)
                                End If
                            Else
646                             Call WriteLocaleMsg(UserIndex, "7", e_FontTypeNames.FONTTYPE_INFO)
648                             Call WriteWorkRequestTarget(UserIndex, 0)
                            End If
                        Else
650                         If UserList(userIndex).Grupo.Lider.ArrayIndex = userIndex Then
652                             Call InvitarMiembro(UserIndex, tU)
                            Else
'Msg1131= Tu no podés invitar usuarios, debe hacerlo ¬1
Call WriteLocaleMsg(UserIndex, "1131", e_FontTypeNames.FONTTYPE_INFO, UserList(UserList(UserIndex).Grupo.Lider.ArrayIndex).name)
656                             Call WriteWorkRequestTarget(UserIndex, 0)
                            End If
                        End If
                    Else
658                     Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)
                    End If
660             Case e_Skill.MarcaDeClan
                    'Target whatever is in that tile
                    Dim clan_nivel As Byte
                
662                 If UserList(UserIndex).GuildIndex = 0 Then
664                     ' Msg720=Servidor » No perteneces a ningún clan.
                        Call WriteLocaleMsg(UserIndex, "720", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
                
666                 clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)

668                 If clan_nivel < 3 Then
670                     ' Msg721=Servidor » El nivel de tu clan debe ser 3 para utilizar esta opción.
                        Call WriteLocaleMsg(UserIndex, "721", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub
                    End If
                                
672                 Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

674                 If Not IsValidUserRef(.flags.targetUser) Then Exit Sub
676                 tU = .flags.targetUser.ArrayIndex
                    
678                 If UserList(UserIndex).GuildIndex = UserList(tU).GuildIndex Then
'Msg1132= Servidor » No podes marcar a un miembro de tu clan.
Call WriteLocaleMsg(UserIndex, "1132", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
682                 If tU > 0 And tU <> UserIndex Then

684                     If UserList(tU).flags.AdminInvisible <> 0 Then Exit Sub
                        'Can't steal administrative players
686                     If UserList(tU).flags.Muerto = 0 Then
                            'call marcar
688                         If UserList(tU).flags.invisible = 1 Or UserList(tU).flags.Oculto = 1 Then
                                UserList(userindex).Counters.timeFx = 3
690                             Call SendData(SendTarget.ToClanArea, userindex, PrepareMessageParticleFX(UserList(tU).Char.charindex, 210, 50, False, , UserList(userindex).Pos.X, UserList(userindex).Pos.y))
                            Else
                                UserList(userindex).Counters.timeFx = 3
692                             Call SendData(SendTarget.ToClanArea, userindex, PrepareMessageParticleFX(UserList(tU).Char.charindex, 210, 150, False, , UserList(userindex).Pos.X, UserList(userindex).Pos.y))
                            End If
694                         Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageConsoleMsg("Clan> [" & UserList(UserIndex).Name & "] marco a " & UserList(tU).Name & ".", e_FontTypeNames.FONTTYPE_GUILD))
                        Else
696                         Call WriteLocaleMsg(UserIndex, "7", e_FontTypeNames.FONTTYPE_INFO)
698                         Call WriteWorkRequestTarget(UserIndex, 0)
                        End If
                    Else
700                     Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)
                    End If
702             Case e_Skill.MarcaDeGM
704                 Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
706                 tU = .flags.targetUser.ArrayIndex
708                 If IsValidUserRef(.flags.targetUser) Then
'Msg1133= Servidor » [¬1
Call WriteLocaleMsg(UserIndex, "1133", e_FontTypeNames.FONTTYPE_INFO, UserList(tU).name)
                    Else
712                     Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)

                    End If
                Case e_Skill.TargetableItem
                    If .Stats.MinSta < ObjData(.invent.Object(.flags.TargetObjInvSlot).objIndex).MinSta Then
                        Call WriteLocaleMsg(UserIndex, MsgNotEnoughtStamina, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    Call LookatTile(UserIndex, UserList(UserIndex).pos.map, X, y)
                    Call UserTargetableItem(UserIndex, X, y)
            End Select

        End With
        
        Exit Sub

HandleWorkLeftClick_Err:
714     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWorkLeftClick", Erl)
716
        
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo ErrHandler

100 With UserList(UserIndex)
        
        Dim Desc       As String
        Dim GuildName  As String
        Dim errorStr   As String
        Dim Alineacion As Byte
        
102     Desc = Reader.ReadString8()
104     GuildName = Reader.ReadString8()
106     Alineacion = Reader.ReadInt8()
        
108     If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Alineacion, errorStr) Then

110         Call QuitarObjetos(407, 1, UserIndex)
112         Call QuitarObjetos(408, 1, UserIndex)
114         Call QuitarObjetos(409, 1, UserIndex)
116         Call QuitarObjetos(412, 1, UserIndex)
            
            
                
118             Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageConsoleMsg(.name & " ha fundado el clan <" & GuildName & "> de alineación " & GuildAlignment(.GuildIndex) & ".", e_FontTypeNames.FONTTYPE_GUILD))
120             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
                'Update tag
122             Call RefreshCharStatus(UserIndex)
            Else
124             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

        End If

    End With
        
    Exit Sub
        
ErrHandler:

126 Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNewGuild", Erl)
128

End Sub

''
' Handles the "SpellInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSpellInfo_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)
        
            Dim spellSlot As Byte
            Dim Spell     As Integer
        
102         spellSlot = Reader.ReadInt8()
        
            'Validate slot
104         If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
'Msg1134= ¡Primero selecciona el hechizo!
Call WriteLocaleMsg(UserIndex, "1134", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate spell in the slot
108         Spell = .Stats.UserHechizos(spellSlot)

110         If Spell > 0 And Spell < NumeroHechizos + 1 Then

112             With Hechizos(Spell)
                    'Send information
                    Call WriteConsoleMsg(UserIndex, "HECINF*" & Spell, e_FontTypeNames.FONTTYPE_INFO)

                End With

            End If

        End With
        
        Exit Sub

HandleSpellInfo_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpellInfo", Erl)
118
        
End Sub

''
' Handles the "EquipItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)
        
        On Error GoTo HandleEquipItem_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim itemSlot As Byte
102         itemSlot = Reader.ReadInt8()
                
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.EquipItem
            
            'If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), userindex, "EquipItem", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
            
            'Dead users can't equip items
104         If .flags.Muerto = 1 Then
'Msg1136= ¡¡Estás muerto!! Sólo podés usar items cuando estás vivo.
Call WriteLocaleMsg(UserIndex, "1136", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate item slot
108         If itemSlot > UserList(UserIndex).CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
110         If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
112         Call EquiparInvItem(UserIndex, itemSlot)

        End With
        
        Exit Sub

HandleEquipItem_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEquipItem", Erl)
116
        
End Sub

''
' Handles the "Change_Heading" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChange_Heading(ByVal UserIndex As Integer)
        
        On Error GoTo HandleChange_Heading_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 10/01/08
        'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
        ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Heading As e_Heading
102             Heading = Reader.ReadInt8()
                            
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.ChangeHeading
            
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "ChangeHeading", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
        
            'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
104         If Heading > 0 And Heading < 5 Then
106             .Char.Heading = Heading
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterChange(.Char.body, .Char.head, .Char.Heading, .Char.charindex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CartAnim, .Char.FX, .Char.loops, .Char.CascoAnim, False, .flags.Navegando))

            End If

        End With

        Exit Sub

HandleChange_Heading_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChange_Heading", Erl)
112
        
End Sub

''
' Handles the "ModifySkills" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleModifySkills(ByVal UserIndex As Integer)
        
        On Error GoTo HandleModifySkills_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)

            Dim i                      As Long
            Dim Count                  As Integer
            Dim points(1 To NUMSKILLS) As Byte
        
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
102         For i = 1 To NUMSKILLS
104             points(i) = Reader.ReadInt8()
            
106             If points(i) < 0 Then
108                 Call LogSecurity(.name & " IP:" & .ConnectionDetails.IP & " trató de hackear los skills.")
110                 .Stats.SkillPts = 0
112                 Call CloseSocket(UserIndex)
                    Exit Sub

                End If
            
114             Count = Count + points(i)
116         Next i
        
118         If Count > .Stats.SkillPts Then
120             Call LogSecurity(.name & " IP:" & .ConnectionDetails.IP & " trató de hackear los skills.")
122             Call CloseSocket(UserIndex)
                Exit Sub

            End If

            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        
124         With .Stats

126             For i = 1 To NUMSKILLS
128                 .SkillPts = .SkillPts - points(i)
                    
                    If .UserSkills(i) <> .UserSkills(i) + points(i) Then
130                     .UserSkills(i) = .UserSkills(i) + points(i)
                    
                        'Client should prevent this, but just in case...
132                     If .UserSkills(i) > 100 Then
134                         .SkillPts = .SkillPts + .UserSkills(i) - 100
136                         .UserSkills(i) = 100
                        End If
                        
                        UserList(UserIndex).flags.ModificoSkills = True
                    End If
138             Next i

            End With

        End With
        
        Exit Sub

HandleModifySkills_Err:
140     Call TraceError(Err.Number, Err.Description, "Protocol.HandleModifySkills", Erl)
142
        
End Sub

''
' Handles the "Train" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)
        
        On Error GoTo HandleTrain_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim SpawnedNpc As Integer
            Dim PetIndex   As Byte
        
102         PetIndex = Reader.ReadInt8()
        
104         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
106         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Entrenador Then Exit Sub
108         If NpcList(.flags.TargetNPC.ArrayIndex).Mascotas < MAXMASCOTASENTRENADOR Then
        
110             If PetIndex > 0 And PetIndex < NpcList(.flags.TargetNPC.ArrayIndex).NroCriaturas + 1 Then
                    'Create the creature
112                 SpawnedNpc = SpawnNpc(NpcList(.flags.TargetNPC.ArrayIndex).Criaturas(PetIndex).NpcIndex, NpcList(.flags.TargetNPC.ArrayIndex).Pos, True, False)
                
114                 If SpawnedNpc > 0 Then
116                     NpcList(SpawnedNpc).MaestroNPC = .flags.TargetNPC
118                     NpcList(.flags.TargetNPC.ArrayIndex).Mascotas = NpcList(.flags.TargetNPC.ArrayIndex).Mascotas + 1
                    End If
                End If
            Else
120             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead("No puedo traer más criaturas, mata las existentes!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
            End If

        End With
        
        Exit Sub

HandleTrain_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTrain", Erl)
124
        
End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCommerceBuy_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Slot   As Byte
            Dim amount As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
        
            'Dead people can't commerce...
106         If .flags.Muerto = 1 Then
108             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'El target es un NPC valido?
110         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
            
            'íEl NPC puede comerciar?
112         If NpcList(.flags.TargetNPC.ArrayIndex).Comercia = 0 Then
114             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
                Exit Sub

            End If
        
            'Only if in commerce mode....
116         If Not .flags.Comerciando Then
'Msg1137= No estás comerciando
Call WriteLocaleMsg(UserIndex, "1137", e_FontTypeNames.FONTTYPE_INFO)
120             Call WriteCommerceEnd(UserIndex)
                Exit Sub

            End If
        
            'User compra el item
122         Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC.ArrayIndex, Slot, amount)

        End With

        Exit Sub

HandleCommerceBuy_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceBuy", Erl)
126
        
End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)
        
        On Error GoTo HandleBankExtractItem_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)

            Dim Slot        As Byte
            Dim slotdestino As Byte
            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
106         slotdestino = Reader.ReadInt8()
        
            'Dead people can't commerce
108         If .flags.Muerto = 1 Then
110             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            '¿El target es un NPC valido?
112         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        
            '¿Es el banquero?
114         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub

            'User retira el item del slot
116         Call UserRetiraItem(UserIndex, Slot, amount, slotdestino)

        End With

        Exit Sub

HandleBankExtractItem_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankExtractItem", Erl)
120
        
End Sub

''
' Handles the "CommerceSell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCommerceSell_Err
        
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)

            Dim Slot   As Byte
            Dim amount As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
        
            'Dead people can't commerce...
106         If .flags.Muerto = 1 Then
108             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'íEl target es un NPC valido?
110         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        
            'íEl NPC puede comerciar?
112         If NpcList(.flags.TargetNPC.ArrayIndex).Comercia = 0 Then
114             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageChatOverHead("No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
                Exit Sub

            End If
        
            'User compra el item del slot
116         Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC.ArrayIndex, Slot, amount)

        End With

        Exit Sub

HandleCommerceSell_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceSell", Erl)
120
        
End Sub

''
' Handles the "BankDeposit" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)
        
        On Error GoTo HandleBankDeposit_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)
        
            Dim Slot        As Byte
            Dim slotdestino As Byte
            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
106         slotdestino = Reader.ReadInt8()
        
            'Dead people can't commerce...
108         If .flags.Muerto = 1 Then
110             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'íEl target es un NPC valido?
112         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
            'íEl NPC puede comerciar?
114         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then
                Exit Sub
            End If
            
116         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
118             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'User deposita el item del slot rdata
120         Call UserDepositaItem(UserIndex, Slot, amount, slotdestino)

        End With
        
        Exit Sub

HandleBankDeposit_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankDeposit", Erl)
124
        
End Sub

''
' Handles the "ForumPost" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim File     As String
            Dim title    As String
            Dim Msg      As String
            Dim postFile As String
            Dim handle   As Integer
            Dim i        As Long
            Dim Count    As Integer
        
102         title = Reader.ReadString8()
104         Msg = Reader.ReadString8()
        
106         If .flags.TargetObj > 0 Then
108             File = App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
            
110             If FileExist(File, vbNormal) Then
112                 Count = val(GetVar(File, "INFO", "CantMSG"))
                
                    'If there are too many messages, delete the forum
114                 If Count > MAX_MENSAJES_FORO Then

116                     For i = 1 To Count
118                         Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & i & ".for"
120                     Next i

122                     Kill App.Path & "\foros\" & UCase$(ObjData(.flags.TargetObj).ForoID) & ".for"
124                     Count = 0

                    End If

                Else
                    'Starting the forum....
126                 Count = 0

                End If
            
128             handle = FreeFile()
130             postFile = Left$(File, Len(File) - 4) & CStr(Count + 1) & ".for"
            
                'Create file
132             Open postFile For Output As handle
134             Print #handle, title
136             Print #handle, Msg
138             Close #handle
            
                'Update post count
140             Call WriteVar(File, "INFO", "CantMSG", Count + 1)

            End If

        End With
        
        Exit Sub
        
ErrHandler:
142     Close #handle
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForumPost", Erl)
146

End Sub

''
' Handles the "MoveSpell" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)
        
        On Error GoTo HandleMoveSpell_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

            Dim dir As Integer
        
102         If Reader.ReadBool() Then
104             dir = 1
            Else
106             dir = -1

            End If
        
108         Call DesplazarHechizo(UserIndex, dir, Reader.ReadInt8())

        Exit Sub

HandleMoveSpell_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
112
        
End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Desc As String
        
102         Desc = Reader.ReadString8()
        
104         Call modGuilds.ChangeCodexAndDesc(Desc, .GuildIndex)

        End With
        
        Exit Sub
        
ErrHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveSpell", Erl)
108

End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUserCommerceOffer_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
100     With UserList(UserIndex)

            Dim tUser  As Integer
            Dim Slot   As Byte
            Dim amount As Long
            
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt32()
            
            
            'Is the commerce attempt valid??
            If Not IsValidUserRef(.ComUsu.DestUsu) Then
                Call FinComerciarUsu(userIndex)
                Exit Sub
            End If
            'Get the other player
106         tUser = .ComUsu.DestUsu.ArrayIndex
            If UserList(tUser).ComUsu.DestUsu.ArrayIndex <> UserIndex Then
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            End If
            'If Amount is invalid, or slot is invalid and it's not gold, then ignore it.
108         If ((Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Or amount <= 0 Then Exit Sub
        
            'Is the other player valid??
110         If tUser < 1 Or tUser > MaxUsers Then Exit Sub
        
            'Is he still logged??
116         If Not UserList(tUser).flags.UserLogged Then
118             Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else

                'Is he alive??
120             If UserList(tUser).flags.Muerto = 1 Then
122                 Call FinComerciarUsu(UserIndex)
                    Exit Sub

                End If
            
                'Has he got enough??
124             If Slot = FLAGORO Then

                    'gold
126                 If amount > .Stats.GLD Then
'Msg1138= No tienes esa cantidad.
Call WriteLocaleMsg(UserIndex, "1138", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                Else

                    'inventory
130                 If amount > .Invent.Object(Slot).amount Then
'Msg1139= No tienes esa cantidad.
Call WriteLocaleMsg(UserIndex, "1139", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If
                
134                 If .Invent.Object(Slot).ObjIndex > 0 Then
136                     If ObjData(.Invent.Object(Slot).ObjIndex).Instransferible = 1 Then
'Msg1140= Este objeto es intransferible, no podés venderlo.
Call WriteLocaleMsg(UserIndex, "1140", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
    
                        End If
                    
140                     If ObjData(.Invent.Object(Slot).ObjIndex).Newbie = 1 Then
'Msg1141= No puedes comerciar objetos newbie.
Call WriteLocaleMsg(UserIndex, "1141", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
    
                    End If

                End If
            
                'Prevent offer changes (otherwise people would ripp off other players)
                'If .ComUsu.Objeto > 0 Then
'Msg1142= No podés cambiar tu oferta.
Call WriteLocaleMsg(UserIndex, "1142", e_FontTypeNames.FONTTYPE_INFO)
                '     Exit Sub

                '  End If
            
                'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
144             If .flags.Navegando = 1 Then
146                 If .Invent.BarcoSlot = Slot Then
'Msg1143= No podés vender tu barco mientras lo estás usando.
Call WriteLocaleMsg(UserIndex, "1143", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
            
150             If .flags.Montado = 1 Then
152                 If .Invent.MonturaSlot = Slot Then
'Msg1144= No podés vender tu montura mientras la estás usando.
Call WriteLocaleMsg(UserIndex, "1144", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                End If
            
156             .ComUsu.Objeto = Slot
158             .ComUsu.cant = amount
            
                'If the other one had accepted, we turn that back and inform of the new offer (just to be cautious).
160             If UserList(tUser).ComUsu.Acepto Then
162                 UserList(tUser).ComUsu.Acepto = False
164                 Call WriteConsoleMsg(tUser, .Name & " ha cambiado su oferta.", e_FontTypeNames.FONTTYPE_TALK)

                End If
            
                Dim ObjAEnviar As t_Obj
                
166             ObjAEnviar.amount = amount

                'Si no es oro tmb le agrego el objInex
168             If Slot <> FLAGORO Then ObjAEnviar.ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
                'Llamos a la funcion
170             Call EnviarObjetoTransaccion(tUser, UserIndex, ObjAEnviar)

            End If

        End With

        Exit Sub

HandleUserCommerceOffer_Err:
172     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUserCommerceOffer", Erl)
174
        
End Sub

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("No se pueden actualizar relaciones.", e_FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD))
            End If
        End With
        Exit Sub
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptPeace", Erl)
End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Relaciones de clan desactivadas por el momento.", e_FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub
        
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRejectAlliance", Erl)
116

End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
        
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Relaciones de clan desactivadas por el momento.", e_FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub
        
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRejectPeace", Erl)
116

End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild          As String
            Dim errorStr       As String
            Dim otherClanIndex As String
        
102         guild = Reader.ReadString8()
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Relaciones de clan desactivadas por el momento.", e_FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        
106         If otherClanIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex), e_FontTypeNames.FONTTYPE_GUILD))

            End If

        End With
        
        Exit Sub
        
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptAlliance", Erl)
116

End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim proposal As String
            Dim errorStr As String
        
102         guild = Reader.ReadString8()
104         proposal = Reader.ReadString8()
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Relaciones de clan desactivadas por el momento.", e_FontTypeNames.FONTTYPE_GUILD))
            Exit Sub
        End With
        
        Exit Sub
        
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
114

End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim proposal As String
            Dim errorStr As String
        
102         guild = Reader.ReadString8()
104         proposal = Reader.ReadString8()
'Msg1145= Relaciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1145", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End With
        
        Exit Sub
        
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
114

End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim errorStr As String
            Dim details  As String
        
102         guild = Reader.ReadString8()
            
'Msg1146= Relaciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1146", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
106         If LenB(details) = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
            Else
110             Call WriteOfferDetails(UserIndex, details)

            End If

        End With
        
        Exit Sub
        
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOfferPeace", Erl)
114

End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild    As String
            Dim errorStr As String
            Dim details  As String
        
102         guild = Reader.ReadString8()
'Msg1147= Relaciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1147", e_FontTypeNames.FONTTYPE_INFO)
            
        End With
        
        Exit Sub
        
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeaceDetails", Erl)
114

End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim user    As String
            Dim details As String
        
102         user = Reader.ReadString8()
        
104         details = modGuilds.a_DetallesAspirante(UserIndex, user)
        
106         If LenB(details) = 0 Then
'Msg1148= El personaje no ha mandado solicitud, o no estás habilitado para verla.
Call WriteLocaleMsg(UserIndex, "1148", e_FontTypeNames.FONTTYPE_INFO)
            Else
110             Call WriteShowUserRequest(UserIndex, details)

            End If

        End With
        
        Exit Sub
        
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestJoinerInfo", Erl)
114

End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo HandleGuildAlliancePropList_Err
    
'Msg1149= Relaciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1149", e_FontTypeNames.FONTTYPE_INFO)
        
        Exit Sub

HandleGuildAlliancePropList_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAlliancePropList", Erl)
104
        
End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
      
        On Error GoTo HandleGuildPeacePropList_Err

'Msg1150= Relaciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1150", e_FontTypeNames.FONTTYPE_INFO)
        
        Exit Sub

HandleGuildPeacePropList_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
104
        
End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild           As String
            Dim errorStr        As String
            Dim otherGuildIndex As Integer
        
102         guild = Reader.ReadString8()
        
'Msg1151= Relaciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1151", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
106         If otherGuildIndex = 0 Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
                'WAR shall be!
110             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & guild, e_FontTypeNames.FONTTYPE_GUILD))
112             Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN", e_FontTypeNames.FONTTYPE_GUILD))
114             Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
116             Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

            End If

        End With
        
        Exit Sub
        
ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildPeacePropList", Erl)
120

End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo ErrHandler

100     Call modGuilds.ActualizarWebSite(UserIndex, Reader.ReadString8())

        Exit Sub
        
ErrHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildNewWebsite", Erl)
104

End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim errorStr As String
            Dim UserName As String
            Dim tUser    As t_UserReference
        
102         UserName = Reader.ReadString8()
        
108         tUser = NameIndex(username)

            If IsValidUserRef(tUser) Then
104             If Not modGuilds.a_AceptarAspirante(UserIndex, username, errorStr) Then
106                 Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
                Else
112                 Call modGuilds.m_ConectarMiembroAClan(tUser.ArrayIndex, .GuildIndex)
114                 Call RefreshCharStatus(tUser.ArrayIndex)
116                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("[" & username & "] ha sido aceptado como miembro del clan.", e_FontTypeNames.FONTTYPE_GUILD))
118                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))
                End If
            Else
                If Not modGuilds.a_AceptarAspirante(UserIndex, username, errorStr) Then
                    Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
                Else
124                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("[" & username & "] ha sido aceptado como miembro del clan.", e_FontTypeNames.FONTTYPE_GUILD))
                End If
            End If

        End With
        
        Exit Sub
        
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
122

End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/08/07
        'Last Modification by: (liquid)
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim errorStr As String
            Dim UserName As String
            Dim Reason   As String
            Dim tUser    As t_UserReference
        
102         UserName = Reader.ReadString8()
104         Reason = Reader.ReadString8()
        
106         If Not modGuilds.a_RechazarAspirante(UserIndex, UserName, errorStr) Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)
            Else
110             tUser = NameIndex(UserName)
112             If IsValidUserRef(tUser) Then
114                 Call WriteConsoleMsg(tUser.ArrayIndex, errorStr & " : " & Reason, e_FontTypeNames.FONTTYPE_GUILD)
                Else
                    'hay que grabar en el char su rechazo
116                 Call modGuilds.a_RechazarAspiranteChar(UserName, .GuildIndex, Reason)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildAcceptNewMember", Erl)
120

End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim UserName   As String
            Dim GuildIndex As Integer
        
102         UserName = Reader.ReadString8()
        Dim CharId As Long
        CharId = GetCharacterIdWithName(username)
        If CharId <= 0 Then
            Exit Sub
        End If
104         GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, CharId)
        
106         If GuildIndex > 0 Then
                Dim expulsado As t_UserReference
108             expulsado = NameIndex(username)
'Msg1152= Has sido expulsado del clan.
Call WriteLocaleMsg(expulsado.ArrayIndex, "1152", e_FontTypeNames.FONTTYPE_INFO)
112             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " fue expulsado del clan.", e_FontTypeNames.FONTTYPE_GUILD))
114             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Else
'Msg1153= No podés expulsar ese personaje del clan.
Call WriteLocaleMsg(UserIndex, "1153", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
        
ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildKickMember", Erl)
120

End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     Call modGuilds.ActualizarNoticias(UserIndex, Reader.ReadString8())

        Exit Sub
        
ErrHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildUpdateNews", Erl)
104

End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     Call modGuilds.SendDetallesPersonaje(UserIndex, Reader.ReadString8())

        Exit Sub
        
ErrHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMemberInfo", Erl)
104

End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)
        
        On Error GoTo HandleGuildOpenElections_Err
        
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

            Dim Error As String
'Msg1154= Elecciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1154", e_FontTypeNames.FONTTYPE_INFO)
        End With
        
        Exit Sub

HandleGuildOpenElections_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOpenElections", Erl)
110
        
End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild       As String
            Dim application As String
            Dim errorStr    As String
        
102         guild = Reader.ReadString8()
104         application = Reader.ReadString8()
        
106         If Not modGuilds.a_NuevoAspirante(UserIndex, guild, application, errorStr) Then
108             Call WriteConsoleMsg(UserIndex, errorStr, e_FontTypeNames.FONTTYPE_GUILD)

            Else
'Msg1155= Tu solicitud ha sido enviada. Espera prontas noticias del líder de ¬1
Call WriteLocaleMsg(UserIndex, "1155", e_FontTypeNames.FONTTYPE_INFO, guild)

            End If

        End With
        
        Exit Sub
        
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestMembership", Erl)
114

End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler
 
100     Call modGuilds.SendGuildDetails(UserIndex, Reader.ReadString8())

        Exit Sub
        
ErrHandler:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildRequestDetails", Erl)
104

End Sub



''
' Handles the "Quit" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuit_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/15/2008 (NicoNZ)
        'If user is invisible, it automatically becomes
        'visible before doing the countdown to exit
        '04/15/2008 - No se reseteaban los contadores de invi ni de ocultar. (NicoNZ)
        '***************************************************
        
    Dim tUser        As Integer
    
100     With UserList(UserIndex)

102         If .flags.Paralizado = 1 Then
'Msg1156= No podés salir estando paralizado.
Call WriteLocaleMsg(UserIndex, "1156", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'exit secure commerce
106         If .ComUsu.DestUsu.ArrayIndex > 0 Then
108             tUser = .ComUsu.DestUsu.ArrayIndex
            
110             If IsValidUserRef(.ComUsu.DestUsu) And UserList(tUser).flags.UserLogged Then
            
112                 If UserList(tUser).ComUsu.DestUsu.ArrayIndex = userIndex Then
'Msg1157= Comercio cancelado por el otro usuario
Call WriteLocaleMsg(tUser, "1157", e_FontTypeNames.FONTTYPE_INFO)
116                     Call FinComerciarUsu(tUser)

                    End If

                End If
            
'Msg1158= Comercio cancelado. 
Call WriteLocaleMsg(UserIndex, "1158", e_FontTypeNames.FONTTYPE_INFO)
120             Call FinComerciarUsu(UserIndex)

        End If

138         Call Cerrar_Usuario(UserIndex)

        End With

        Exit Sub

HandleQuit_Err:
140     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuit", Erl)
142
        
End Sub

''
' Handles the "GuildLeave" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)
        
        On Error GoTo HandleGuildLeave_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim GuildIndex As Integer
    
100     With UserList(UserIndex)

            'obtengo el guildindex
102         GuildIndex = m_EcharMiembroDeClan(UserIndex, .id)
        
104         If GuildIndex > 0 Then
'Msg1159= Dejas el clan.
Call WriteLocaleMsg(UserIndex, "1159", e_FontTypeNames.FONTTYPE_INFO)
108             Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", e_FontTypeNames.FONTTYPE_GUILD))
            Else
'Msg1160= Tu no puedes salir de ningún clan.
Call WriteLocaleMsg(UserIndex, "1160", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        Exit Sub

HandleGuildLeave_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildLeave", Erl)
114
        
End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)
        
        On Error GoTo HandleRequestAccountState_Err
        
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        Dim earnings   As Integer
        Dim percentage As Integer
    
100     With UserList(UserIndex)

            'Dead people can't check their accounts
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If Not IsValidNpcRef(.flags.TargetNPC) Then
'Msg1161= Primero tenes que seleccionar un personaje, hace click izquierdo sobre él.
Call WriteLocaleMsg(UserIndex, "1161", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
110         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 3 Then
112             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
114         Select Case NpcList(.flags.TargetNPC.ArrayIndex).npcType
                Case e_NPCType.Banquero
116                 Call WriteChatOverHead(UserIndex, "Tenes " & PonerPuntos(.Stats.Banco) & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
            
118             Case e_NPCType.Timbero
120                 If Not .flags.Privilegios And e_PlayerType.user Then
122                     earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
124                     If earnings >= 0 And Apuestas.Ganancias <> 0 Then
126                         percentage = Int(earnings * 100 / Apuestas.Ganancias)
                        End If
                    
128                     If earnings < 0 And Apuestas.Perdidas <> 0 Then
130                         percentage = Int(earnings * 100 / Apuestas.Perdidas)
                        End If
                    
'Msg1162= Entradas: ¬1
Call WriteLocaleMsg(UserIndex, "1162", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Apuestas.Ganancias))
                    End If
            End Select
        End With
        Exit Sub

HandleRequestAccountState_Err:
134     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestAccountState", Erl)
136
        
End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo HandlePetStand_Err
        
100     With UserList(UserIndex)

            'Dead people can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If Not IsValidNpcRef(.flags.TargetNPC) Then
108             ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "757", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
110         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
'Msg1163= Estás demasiado lejos.
Call WriteLocaleMsg(UserIndex, "1163", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Make sure it's his pet
114         If Not IsValidUserRef(NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser) Or NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser.ArrayIndex <> UserIndex Then Exit Sub
        
            'Do it!
116         Call SetMovement(.flags.TargetNPC.ArrayIndex, e_TipoAI.Estatico)
118         Call Expresar(.flags.TargetNPC.ArrayIndex, UserIndex)
        End With
        
        Exit Sub

HandlePetStand_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetStand", Erl)
122
        
End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo HandlePetFollow_Err
        
100     With UserList(UserIndex)

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If Not IsValidNpcRef(.flags.TargetNPC) Then
108             ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "757", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make sure it's close enough
110         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
'Msg1164= Estás demasiado lejos.
Call WriteLocaleMsg(UserIndex, "1164", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make usre it's the user's pet
114         If Not IsValidUserRef(NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser) Or NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser.ArrayIndex <> UserIndex Then Exit Sub
        
            'Do it
116         Call FollowAmo(.flags.TargetNPC.ArrayIndex)
        
118         Call Expresar(.flags.TargetNPC.ArrayIndex, UserIndex)
        End With
        
        Exit Sub

HandlePetFollow_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetFollow", Erl)
122
        
End Sub

''
' Handles the "PetLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetLeave(ByVal UserIndex As Integer)
        '***************************************************
        
        On Error GoTo HandlePetLeave_Err
        
100     With UserList(UserIndex)

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
106         If Not IsValidNpcRef(.flags.TargetNPC) Then
108             ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "757", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Make usre it's the user's pet
110         If Not IsValidUserRef(NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser) Or NpcList(.flags.TargetNPC.ArrayIndex).MaestroUser.ArrayIndex <> UserIndex Then Exit Sub

112         Call QuitarNPC(.flags.TargetNPC.ArrayIndex, e_DeleteSource.ePetLeave)

        End With
        
        Exit Sub

HandlePetLeave_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetLeave", Erl)
116
        
End Sub

''
' Handles the "GrupoMsg" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGrupoMsg(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim chat As String
102         chat = Reader.ReadString8()
104         If LenB(chat) <> 0 Then

108             If .Grupo.EnGrupo = True Then

                    Dim i As Byte
         
110                 For i = 1 To UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros
112                     Call WriteConsoleMsg(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex, .name & "> " & chat, e_FontTypeNames.FONTTYPE_New_Amarillo_Verdoso)
114                     Call WriteChatOverHead(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(i).ArrayIndex, "NOCONSOLA*" & chat, UserList(userIndex).Char.charindex, &HFF8000)
116                 Next i
                Else
118                 ' Msg758=Grupo> No estas en ningun grupo.
                    Call WriteLocaleMsg(UserIndex, "758", e_FontTypeNames.FONTTYPE_New_GRUPO)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGrupoMsg", Erl)
122

End Sub

Private Sub HandleTrainList(ByVal UserIndex As Integer)
        On Error GoTo HandleTrainList_Err
        
100     With UserList(UserIndex)
            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Validate target NPC
106         If Not IsValidNpcRef(.flags.TargetNPC) Then
108             ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "757", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Make sure it's close enough
110         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
112             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Make sure it's the trainer
114         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Entrenador Then Exit Sub
116         Call WriteTrainerCreatureList(UserIndex, .flags.TargetNPC.ArrayIndex)

        End With
        Exit Sub

HandleTrainList_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTrainList", Erl)
120
        
End Sub

''
' Handles the "Rest" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleRest_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

            'Dead users can't use pets
102         If .flags.Muerto = 1 Then
104             ' Msg752=¡¡Estás muerto!! Solo podés usar items cuando estás vivo.
                Call WriteLocaleMsg(UserIndex, "752", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
106         If HayOBJarea(.Pos, FOGATA) Then
108             Call WriteRestOK(UserIndex)
            
110             If Not .flags.Descansar Then
112                 ' Msg753=Te acomodás junto a la fogata y comenzas a descansar.
                    Call WriteLocaleMsg(UserIndex, "753", e_FontTypeNames.FONTTYPE_INFO)
                Else
114                 ' Msg754=Te levantas.
                    Call WriteLocaleMsg(UserIndex, "754", e_FontTypeNames.FONTTYPE_INFO)

                End If
            
116             .flags.Descansar = Not .flags.Descansar
            Else

118             If .flags.Descansar Then
120                 Call WriteRestOK(UserIndex)
122                 ' Msg754=Te levantas.
                    Call WriteLocaleMsg(UserIndex, "754", e_FontTypeNames.FONTTYPE_INFO)
                
124                 .flags.Descansar = False
                    Exit Sub

                End If
            
126             ' Msg755=No hay ninguna fogata junto a la cual descansar.
                Call WriteLocaleMsg(UserIndex, "755", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With

        Exit Sub

HandleRest_Err:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRest", Erl)
130
        
End Sub

''
' Handles the "Meditate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)
        
        On Error GoTo HandleMeditate_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 04/15/08 (NicoNZ)
        'Arreglí un bug que mandaba un index de la meditacion diferente
        'al que decia el server.
        '***************************************************
        
100     With UserList(UserIndex)

            'Si ya tiene el mana completo, no lo dejamos meditar.
102         If .Stats.MinMAN = .Stats.MaxMAN Then Exit Sub
                           
            'Las clases NO MAGICAS no meditan...
104         If .clase = e_Class.Hunter Or .clase = e_Class.Trabajador Or .clase = e_Class.Warrior Or .clase = e_Class.Pirat Or .clase = e_Class.Thief Then Exit Sub

106         If .flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
108             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
110         If .flags.Montado = 1 Then
112             ' Msg756=No podes meditar estando montado.
                Call WriteLocaleMsg(UserIndex, "756", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

114         .flags.Meditando = Not .flags.Meditando

116         If .flags.Meditando Then

118             .Counters.TimerMeditar = 0
                .Counters.TiempoInicioMeditar = 0
                Dim customEffect As Integer
                Dim Index As Integer
                Dim obj As t_ObjData
                For Index = 1 To UBound(.Invent.Object)
                    If .Invent.Object(Index).objIndex > 0 Then
                        If .Invent.Object(Index).objIndex > 0 Then
                            obj = ObjData(.Invent.Object(Index).objIndex)
                            If obj.OBJType = OtDonador And obj.Subtipo = 4 And .Invent.Object(Index).Equipped Then
                               customEffect = obj.HechizoIndex
                               Exit For
                            End If
                        End If
                    End If
                Next Index
                If customEffect > 0 Then
                    .Char.FX = customEffect
                Else
120                 Select Case .Stats.ELV
    
                        Case 1 To 14
122                         .Char.FX = e_Meditaciones.MeditarInicial
                          'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 37, -1, False))
    
124                     Case 15 To 24
                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 38, -1, False))
126                         .Char.FX = e_Meditaciones.MeditarMayor15
    
128                     Case 25 To 35
130                         .Char.FX = e_Meditaciones.MeditarMayor30
                            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 39, -1, False))
    
132                     Case 35 To 44
134                         .Char.FX = e_Meditaciones.MeditarMayor40
    
136                     Case 45 To 46
138                         .Char.FX = e_Meditaciones.MeditarMayor45
    
140                     Case Else
142                         .Char.FX = e_Meditaciones.MeditarMayor47
    
                    End Select
                End If

            Else
144             .Char.FX = 0

                'Call WriteLocaleMsg(UserIndex, "123", e_FontTypeNames.FONTTYPE_INFO)
            End If

146         Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageMeditateToggle(.Char.charindex, .Char.FX, .Pos.X, .Pos.y))

        End With
        
        Exit Sub

HandleMeditate_Err:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMeditate", Erl)
150
        
End Sub

''
' Handles the "Resucitate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)
        
        On Error GoTo HandleResucitate_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

            'Se asegura que el target es un npc
102         If Not IsValidNpcRef(.flags.TargetNPC) Then
104             ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "757", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate NPC and make sure player is dead
106         If (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor And (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 0 Then Exit Sub
        
            'Make sure it's close enough
108         If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 10 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                'Msg8=Estás muy lejos.
                Exit Sub

            End If
        
112         Call RevivirUsuario(UserIndex)
            UserList(userindex).Counters.timeFx = 3
114         Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageParticleFX(UserList(userindex).Char.charindex, e_ParticulasIndex.Curar, 100, False, , UserList(userindex).Pos.X, UserList(userindex).Pos.y))
116         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("104", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
118         ' Msg585=¡Has sido resucitado!
            Call WriteLocaleMsg(UserIndex, "585", e_FontTypeNames.FONTTYPE_INFO)

        End With
        
        Exit Sub

HandleResucitate_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResucitate", Erl)
122
        
End Sub

Private Sub HandleHeal(ByVal UserIndex As Integer)
        On Error GoTo HandleHeal_Err
100     With UserList(UserIndex)
            'Se asegura que el target es un npc
102         If Not IsValidNpcRef(.flags.TargetNPC) Then
104             ' Msg757=Primero tenés que seleccionar un personaje, hace click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "757", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor And NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.ResucitadorNewbie) Or .flags.Muerto <> 0 Then Exit Sub
        
108         If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 10 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         .Stats.MinHp = .Stats.MaxHp
114         Call WriteUpdateHP(UserIndex)
            'Msg496=¡¡Hás sido curado!!
116         Call WriteLocaleMsg(UserIndex, "496", e_FontTypeNames.FONTTYPE_INFO)
        End With
        
        Exit Sub
HandleHeal_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHeal", Erl)
End Sub


''
' Handles the "CommerceStart" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCommerceStart_Err
        
        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

            'Dead people can't commerce
102         If .flags.Muerto = 1 Then
                ''Msg77=¡¡Estás muerto!!.)
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Is it already in commerce mode??
106         If .flags.Comerciando Then
108             ' Msg759=Ya estás comerciando
                Call WriteLocaleMsg(UserIndex, "759", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            'Validate target NPC
110         If IsValidNpcRef(.flags.TargetNPC) Then
                
                'VOS, como GM, NO podes COMERCIAR con NPCs. (excepto Admins)
112             If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin)) = 0 Then
114                 ' Msg767=No podés vender items.
                    Call WriteLocaleMsg(UserIndex, "767", e_FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
                
                'Does the NPC want to trade??
116             If NpcList(.flags.TargetNPC.ArrayIndex).Comercia = 0 Then
118                 If LenB(NpcList(.flags.TargetNPC.ArrayIndex).Desc) <> 0 Then
120                     Call WriteChatOverHead(UserIndex, "No tengo ningún interés en comerciar.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    End If
                    Exit Sub
                End If
            
122             If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 3 Then
124                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'Start commerce....
126             Call IniciarComercioNPC(UserIndex)
128         ElseIf IsValidUserRef(.flags.targetUser) Then

                ' **********************  Comercio con Usuarios  *********************
                
                'VOS, como GM, NO podes COMERCIAR con usuarios. (excepto  Admins)
130             If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin)) = 0 Then
132                 ' Msg767=No podés vender items.
                    Call WriteLocaleMsg(UserIndex, "767", e_FontTypeNames.FONTTYPE_WARNING)
                    Exit Sub

                End If
                
                'NO podes COMERCIAR CON un GM. (excepto  Admins)
134             If (UserList(.flags.targetUser.ArrayIndex).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin)) = 0 Then
'Msg1165= No podés vender items a este usuario.
Call WriteLocaleMsg(UserIndex, "1165", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                'Is the other one dead??
138             If UserList(.flags.targetUser.ArrayIndex).flags.Muerto = 1 Then
                    Call FinComerciarUsu(.flags.targetUser.ArrayIndex, True)
'Msg1166= ¡¡No podés comerciar con los muertos!!
Call WriteLocaleMsg(UserIndex, "1166", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Is it me??
142             If .flags.targetUser.ArrayIndex = userIndex Then
'Msg1167= No podés comerciar con vos mismo...
Call WriteLocaleMsg(UserIndex, "1167", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Check distance
146             If .pos.map <> UserList(.flags.targetUser.ArrayIndex).pos.map Or Distancia(UserList(.flags.targetUser.ArrayIndex).pos, .pos) > 3 Then
                    Call FinComerciarUsu(.flags.targetUser.ArrayIndex, True)
148                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
 
                'Check if map is not safe
                If MapInfo(.Pos.Map).Seguro = 0 Then
                    Call FinComerciarUsu(.flags.targetUser.ArrayIndex, True)
'Msg1168= No se puede usar el comercio seguro en zona insegura.
Call WriteLocaleMsg(UserIndex, "1168", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                'Is he already trading?? is it with me or someone else??
150             If UserList(.flags.targetUser.ArrayIndex).flags.Comerciando = True Then
                    Call FinComerciarUsu(.flags.targetUser.ArrayIndex, True)
'Msg1169= No podés comerciar con el usuario en este momento.
Call WriteLocaleMsg(UserIndex, "1169", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                'Initialize some variables...
154             .ComUsu.DestUsu = .flags.targetUser
156             .ComUsu.DestNick = UserList(.flags.targetUser.ArrayIndex).name
158             .ComUsu.cant = 0
160             .ComUsu.Objeto = 0
162             .ComUsu.Acepto = False
            
                'Rutina para comerciar con otro usuario
164             Call IniciarComercioConUsuario(userIndex, .flags.targetUser.ArrayIndex)

            Else
166             ' Msg760=Primero haz click izquierdo sobre el personaje.
                Call WriteLocaleMsg(UserIndex, "760", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleCommerceStart_Err:
168     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceStart", Erl)
170
        
End Sub

Private Sub HandleBankStart(ByVal UserIndex As Integer)
        On Error GoTo HandleBankStart_Err
100     With UserList(UserIndex)
            'Dead people can't commerce
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If .flags.Comerciando Then
108             ' Msg759=Ya estás comerciando
                Call WriteLocaleMsg(UserIndex, "759", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Validate target NPC
110         If IsValidNpcRef(.flags.TargetNPC) Then
112             If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 6 Then
114                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                'If it's the banker....
116             If NpcList(.flags.TargetNPC.ArrayIndex).npcType = e_NPCType.Banquero Then
118                 Call IniciarDeposito(UserIndex)
                End If
            Else
120             ' Msg760=Primero haz click izquierdo sobre el personaje.
                Call WriteLocaleMsg(UserIndex, "760", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
HandleBankStart_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankStart", Erl)
End Sub

Private Sub HandleEnlist(ByVal UserIndex As Integer)
        On Error GoTo HandleEnlist_Err
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
            'Validate target NPC
104         If Not IsValidNpcRef(.flags.TargetNPC) Then
106             ' Msg761=Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "761", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
110         If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 4 Then
'Msg1170= Debes acercarte más.
Call WriteLocaleMsg(UserIndex, "1170", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
114         If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
116             Call EnlistarArmadaReal(UserIndex)
            Else
118             Call EnlistarCaos(UserIndex)
            End If
        End With
        Exit Sub
HandleEnlist_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEnlist", Erl)
End Sub

Private Sub HandleInformation(ByVal UserIndex As Integer)
        On Error GoTo HandleInformation_Err
100     With UserList(UserIndex)
            'Validate target NPC
102         If Not IsValidNpcRef(.flags.TargetNPC) Then
104             ' Msg761=Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "761", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
        
108         If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 4 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
114             If .Faccion.Status <> e_Facciones.Armada Or .Faccion.Status <> e_Facciones.consejo Then
116                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    Exit Sub
                End If

118             Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, cada 100 criminales que derrotes te darí una recompensa.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
            
            Else

120             If .Faccion.Status <> e_Facciones.Caos Or .Faccion.Status <> e_Facciones.concilio Then
122                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    Exit Sub

                End If

124             Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te darí una recompensa.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)

            End If

        End With
        
        Exit Sub

HandleInformation_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInformation", Erl)
128
        
End Sub

''
' Handles the "Reward" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)
        
        On Error GoTo HandleReward_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

            'Validate target NPC
102         If Not IsValidNpcRef(.flags.TargetNPC) Then
104             ' Msg761=Primero tenés que seleccionar un personaje, hacé click izquierdo sobre él.
                Call WriteLocaleMsg(UserIndex, "761", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Enlistador Or .flags.Muerto <> 0 Then Exit Sub
108         If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 4 Then
110             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
114             If .Faccion.Status <> e_Facciones.Armada And .Faccion.Status <> e_Facciones.consejo Then
116                 Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    Exit Sub
                End If
118             Call RecompensaArmadaReal(UserIndex)
            Else
120             If .Faccion.Status <> e_Facciones.Caos And .Faccion.Status <> e_Facciones.concilio Then
122                 Call WriteChatOverHead(UserIndex, "No perteneces a la legión oscura!!!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    Exit Sub
                End If
124             Call RecompensaCaos(UserIndex)
            End If
        End With
        
        Exit Sub

HandleReward_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReward", Erl)
128
        
End Sub

''
' Handles the "GuildMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim chat As String
102         chat = Reader.ReadString8()
               
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
                        
            Dim Packet_ID As Long
            Packet_ID = PacketNames.GuildMessage
            
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "GuildMessage", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub
       
104         If LenB(chat) <> 0 Then
                ' WyroX: Foto-denuncias - Push message
                Dim i As Integer

108             For i = 1 To UBound(.flags.ChatHistory) - 1
110                 .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                Next
                
112             .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
114             If .GuildIndex > 0 Then
                    'HarThaoS: si es leade mando un 10 para el status del color(medio villero pero me dio paja)
116                 If LCase(GuildLeader(.GuildIndex)) = .Name Then
118                     Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & chat, 10))
                    Else
                        .Counters.timeGuildChat = 1 + Ceil((3000 + 60 * Len(chat)) / 1000)
                        
120                     Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & chat, .Faccion.Status))
                        Call SendData(SendTarget.ToClanArea, userindex, PrepareMessageChatOverHead("NOCONSOLA*< " & chat & " >", .Char.charindex, RGB(255, 255, 0), , .Pos.X, .Pos.y))
                    End If
                    'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                   ' Call SendData(SendTarget.ToAll, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "í< " & chat & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
                End If

            End If

        End With
        
        Exit Sub
        
ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMessage", Erl)
124

End Sub

''
' Handles the "GuildOnline" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)
        
        On Error GoTo HandleGuildOnline_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

            Dim onlineList As String
102             onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
104         If .GuildIndex <> 0 Then
'Msg1171= Compañeros de tu clan conectados: ¬1
Call WriteLocaleMsg(UserIndex, "1171", e_FontTypeNames.FONTTYPE_INFO, onlineList)
            
            Else
108             ' Msg762=No pertences a ningún clan.
                Call WriteLocaleMsg(UserIndex, "762", e_FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End With
        
        Exit Sub

HandleGuildOnline_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOnline", Erl)
112
        
End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim chat As String
102             chat = Reader.ReadString8()
        
104         If LenB(chat) <> 0 Then
                ' WyroX: Foto-denuncias - Push message
                Dim i As Long
108             For i = 1 To UBound(.flags.ChatHistory) - 1
110                 .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                Next
                
112             .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
            
114             If .Faccion.Status = e_Facciones.consejo Then
116                 Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejo) " & .name & "> " & chat, e_FontTypeNames.FONTTYPE_CONSEJO))

118             ElseIf .Faccion.Status = e_Facciones.concilio Then
120                 Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Concilio) " & .name & "> " & chat, e_FontTypeNames.FONTTYPE_CONSEJOCAOS))

                End If

            End If

        End With
        
        Exit Sub
        
ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCouncilMessage", Erl)
124

End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Description As String
102             Description = Reader.ReadString8()
        
104         If .flags.Muerto = 1 Then
106             ' Msg763=No podés cambiar la descripción estando muerto.
                Call WriteLocaleMsg(UserIndex, "763", e_FontTypeNames.FONTTYPE_INFOIAO)

            Else
            
108             If Len(Description) > 128 Then
110                 ' Msg764=La descripción es muy larga.
                    Call WriteLocaleMsg(UserIndex, "764", e_FontTypeNames.FONTTYPE_INFOIAO)

112             ElseIf Not DescripcionValida(Description) Then
114                 ' Msg765=La descripción tiene carácteres inválidos.
                    Call WriteLocaleMsg(UserIndex, "765", e_FontTypeNames.FONTTYPE_INFOIAO)
                
                Else
116                 .Desc = Trim$(Description)
118                 ' Msg766=La descripción a cambiado.
                    Call WriteLocaleMsg(UserIndex, "766", e_FontTypeNames.FONTTYPE_INFOIAO)

                End If

            End If

        End With
        
        Exit Sub
        
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeDescription", Erl)
122

End Sub

''
' Handles the "GuildVote" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim vote     As String
            Dim errorStr As String
        
102         vote = Reader.ReadString8()
'Msg1172= Elecciones de clan desactivadas por el momento.
Call WriteLocaleMsg(UserIndex, "1172", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End With
        
        Exit Sub
        
ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildVote", Erl)
112

End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)
        
        On Error GoTo HandleBankExtractGold_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

100     With UserList(UserIndex)

            Dim amount As Long
102             amount = Reader.ReadInt32()
        
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Validate target NPC
108         If Not IsValidNpcRef(.flags.TargetNPC) Then
'Msg1173= Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.
Call WriteLocaleMsg(UserIndex, "1173", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub
        
114         If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 10 Then
116             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
118         If amount > 0 And amount <= .Stats.Banco Then
120             .Stats.Banco = .Stats.Banco - amount
122             .Stats.GLD = .Stats.GLD + amount
                Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
124             Call WriteUpdateGold(UserIndex)
                Call WriteUpdateBankGld(UserIndex)
            Else
128             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
            End If
        End With

        Exit Sub

HandleBankExtractGold_Err:
130     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankExtractGold", Erl)
132
        
End Sub

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
        On Error GoTo HandleLeaveFaction_Err
100     With UserList(UserIndex)
            'Dead people can't leave a faction.. they can't talk...
102         If .flags.Muerto = 1 Then
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If .Faccion.Status = e_Facciones.Ciudadano Then
108             If .Faccion.Status = 1 Then
110                 Call VolverCriminal(UserIndex)
'Msg1174= Ahora sos un criminal.
Call WriteLocaleMsg(UserIndex, "1174", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        
            'Validate target NPC
114         If Not IsValidNpcRef(.flags.TargetNPC) Then
116             If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
'Msg1175= Para salir del ejercito debes ir a visitar al rey.
Call WriteLocaleMsg(UserIndex, "1175", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
120             ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
'Msg1176= Para salir de la legion debes ir a visitar al diablo.
Call WriteLocaleMsg(UserIndex, "1176", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                Exit Sub
            End If
        
124         If NpcList(.flags.TargetNPC.ArrayIndex).npcType = e_NPCType.Enlistador Then
                'Quit the Royal Army?
126             If .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
128                 If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 0 Then
                        'Si tiene clan
130                     If .GuildIndex > 0 Then
                            'Y no es leader
132                         If Not PersonajeEsLeader(.Id) Then
                                'Me fijo de que alineación es el clan, si es ARMADA, lo hecho
                                If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA Then
                                    Call m_EcharMiembroDeClan(UserIndex, .name)
'Msg1177= Has dejado el clan.
Call WriteLocaleMsg(UserIndex, "1177", e_FontTypeNames.FONTTYPE_INFO)
                                End If
                            Else
                                'Me fijo si está en un clan armada, en ese caso no lo dejo salir de la facción
                                If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA Then
138                                 Call WriteChatOverHead(UserIndex, "Para dejar la facción primero deberás ceder el liderazgo del clan", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                                    Exit Sub
                                End If
                            End If
                        End If
                    
140                     Call ExpulsarFaccionReal(UserIndex)
142                     Call WriteChatOverHead(UserIndex, "Serás bienvenido a las fuerzas imperiales si deseas regresar.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                        Exit Sub
                    Else
144                     Call WriteChatOverHead(UserIndex, "¡¡¡Sal de aquí bufón!!!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    End If

                    'Quit the Chaos Legion??
146             ElseIf .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.concilio Then
148                 If NpcList(.flags.TargetNPC.ArrayIndex).flags.Faccion = 2 Then
                        'Si tiene clan
                         If .GuildIndex > 0 Then
                            'Y no es leader
                            If Not PersonajeEsLeader(.Id) Then
                                'Me fijo de que alineación es el clan, si es CAOS, lo hecho
                                If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
                                    Call m_EcharMiembroDeClan(UserIndex, .name)
'Msg1178= Has dejado el clan.
Call WriteLocaleMsg(UserIndex, "1178", e_FontTypeNames.FONTTYPE_INFO)
                                End If
                            Else
                                'Me fijo si está en un clan CAOS, en ese caso no lo dejo salir de la facción
                                If GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
                                    Call WriteChatOverHead(UserIndex, "Para dejar la facción primero deberás ceder el liderazgo del clan", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                                    Exit Sub
                                End If
                            End If
                        End If
                    
160                     Call ExpulsarFaccionCaos(UserIndex)
162                     Call WriteChatOverHead(UserIndex, "Ya volverás arrastrandote.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    Else
164                     Call WriteChatOverHead(UserIndex, "Sal de aquí maldito criminal", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                    End If
                Else
166                 Call WriteChatOverHead(UserIndex, "¡No perteneces a ninguna facción!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                End If
            End If
        End With
        Exit Sub
HandleLeaveFaction_Err:
168     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLeaveFaction", Erl)
End Sub

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)
        On Error GoTo HandleBankDepositGold_Err
100     With UserList(UserIndex)

            Dim amount As Long
102         amount = Reader.ReadInt32()
            'Dead people can't leave a faction.. they can't talk...
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'Validate target NPC
108         If Not IsValidNpcRef(.flags.TargetNPC) Then
'Msg1179= Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.
Call WriteLocaleMsg(UserIndex, "1179", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
112         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub
        
114         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
116             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
118         If amount > 0 And amount <= .Stats.GLD Then
                'substract first in case there is overflow we don't dup gold
                .Stats.GLD = .Stats.GLD - amount
                .Stats.Banco = .Stats.Banco + amount
                Call WriteChatOverHead(UserIndex, "Tenés " & .Stats.Banco & " monedas de oro en tu cuenta.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
124             Call WriteUpdateGold(UserIndex)
                Call WriteUpdateBankGld(UserIndex)
            Else
128             Call WriteChatOverHead(UserIndex, "No tenés esa cantidad.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
            End If
        End With
        Exit Sub

HandleBankDepositGold_Err:
130     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBankDepositGold", Erl)
132
        
End Sub

' @param    UserIndex The index of the user sending the message.
Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild       As String
            Dim memberCount As Integer
            Dim i           As Long
            Dim UserName    As String
        
102         guild = Reader.ReadString8()
104         If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
106             If (InStrB(guild, "\") <> 0) Then
108                 guild = Replace(guild, "\", "")
                End If

110             If (InStrB(guild, "/") <> 0) Then
112                 guild = Replace(guild, "/", "")
                End If
                If Not modGuilds.YaExiste(guild) Then
'Msg1180= No existe el clan: ¬1
Call WriteLocaleMsg(UserIndex, "1180", e_FontTypeNames.FONTTYPE_INFO, guild)
                    Exit Sub
                End If
                Dim MembersId() As Long
                MembersId = GetGuildMemberList(guild)
                For i = LBound(MembersId) To UBound(MembersId)
                    Call WriteConsoleMsg(UserIndex, GetUserName(MembersId(i)) & "<" & guild & ">", e_FontTypeNames.FONTTYPE_INFO)
                Next i
            End If
        End With
        Exit Sub
ErrHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildMemberList", Erl)
End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)
        
        On Error GoTo HandleOnlineRoyalArmy_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
    
            Dim i    As Long
            Dim list As String

104         For i = 1 To LastUser

106             If UserList(i).ConnectionDetails.ConnIDValida Then
108                 If UserList(i).Faccion.Status = e_Facciones.Armada Or UserList(i).Faccion.Status = e_Facciones.consejo Then
110                     If UserList(i).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Or .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
112                         list = list & UserList(i).Name & ", "

                        End If

                    End If

                End If

114         Next i

        End With
    
116     If Len(list) > 0 Then
'Msg1289= Armadas conectados: ¬1
Call WriteLocaleMsg(UserIndex, "1289", e_FontTypeNames.FONTTYPE_INFO, Left$(list, Len(list) - 2))
        Else
'Msg1182= No hay Armadas conectados
Call WriteLocaleMsg(UserIndex, "1182", e_FontTypeNames.FONTTYPE_INFO)

        End If
        
        Exit Sub

HandleOnlineRoyalArmy_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineRoyalArmy", Erl)
124
        
End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)
        
        On Error GoTo HandleOnlineChaosLegion_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
    
            Dim i    As Long
            Dim list As String

104         For i = 1 To LastUser

106             If UserList(i).ConnectionDetails.ConnIDValida Then
108                 If UserList(i).Faccion.Status = e_Facciones.Caos Or UserList(i).Faccion.Status = e_Facciones.concilio Then
110                     If UserList(i).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Or .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
112                         list = list & UserList(i).Name & ", "

                        End If

                    End If

                End If

114         Next i

        End With

116     If Len(list) > 0 Then
'Msg1290= Caos conectados: ¬1
Call WriteLocaleMsg(UserIndex, "1290", e_FontTypeNames.FONTTYPE_INFO, Left$(list, Len(list) - 2))
        
        Else
'Msg1184= No hay Caos conectados
Call WriteLocaleMsg(UserIndex, "1184", e_FontTypeNames.FONTTYPE_INFO)

        End If
        
        Exit Sub

HandleOnlineChaosLegion_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineChaosLegion", Erl)
124
        
End Sub

''
' Handles the "Comment" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim comment As String
102             comment = Reader.ReadString8()
        
104         If Not .flags.Privilegios And e_PlayerType.user Then
106             Call LogGM(.Name, "Comentario: " & comment)
'Msg1185= Comentario salvado...
Call WriteLocaleMsg(UserIndex, "1185", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub
        
ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleComment", Erl)
112

End Sub

''
' Handles the "ServerTime" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)
        
        On Error GoTo HandleServerTime_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/08/07
        'Last Modification by: (liquid)
        '***************************************************
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
    
104         Call LogGM(.Name, "Hora.")

        End With
    
106     Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Hora: " & Time & " " & Date, e_FontTypeNames.FONTTYPE_INFO))
        
        Exit Sub

HandleServerTime_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerTime", Erl)
110
        
End Sub

Private Sub HandleUseKey(ByVal UserIndex As Integer)
        
        On Error GoTo HandleUseKey_Err

100     With UserList(UserIndex)

            Dim Slot As Byte
102             Slot = Reader.ReadInt8

104         Call UsarLlave(UserIndex, Slot)
                
        End With
        
        Exit Sub

HandleUseKey_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseKey", Erl)
108
        
End Sub

Private Sub HandleMensajeUser(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim Mensaje  As String
            Dim tUser    As t_UserReference
        
102         UserName = Reader.ReadString8()
104         Mensaje = Reader.ReadString8()
        
106         If EsGM(UserIndex) Then
        
108             If LenB(UserName) = 0 Or LenB(Mensaje) = 0 Then
'Msg1186= Utilice /MENSAJEINFORMACION nick@mensaje
Call WriteLocaleMsg(UserIndex, "1186", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
112                 tUser = NameIndex(UserName)
                
114                 If IsValidUserRef(tUser) Then
'Msg1187= Mensaje recibido de ¬1
Call WriteLocaleMsg(tUser.ArrayIndex, "1187", e_FontTypeNames.FONTTYPE_INFO, .name)
118                     Call WriteConsoleMsg(tUser.ArrayIndex, mensaje, e_FontTypeNames.FONTTYPE_New_DONADOR)
                    Else
120                     If PersonajeExiste(UserName) Then
122                         Call SetMessageInfoDatabase(UserName, "Mensaje recibido de " & .Name & " [Game Master]: " & vbNewLine & Mensaje & vbNewLine)
                        End If
                    End If

'Msg1188= Mensaje enviado a ¬1
Call WriteLocaleMsg(UserIndex, "1188", e_FontTypeNames.FONTTYPE_INFO, UserName)
126                 Call LogGM(.name, "Envió mensaje como GM a " & username & ": " & mensaje)

                End If

            End If

        End With
    
        Exit Sub

ErrHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMensajeUser", Erl)
130

End Sub

Private Sub HandleTraerBoveda(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Ladder
        'Last Modification: 04/jul/2014
        '
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

102         Call UpdateUserHechizos(True, UserIndex, 0)
       
104         Call UpdateUserInv(True, UserIndex, 0)

        End With
    
        Exit Sub

ErrHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTraerBoveda", Erl)
108

End Sub


'HarThaoS: Agrego perdón faccionario.


' Handles the "SendPosMovimiento" message.

Private Sub HandleSendPosMovimiento(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Martín Trionfetti - HarThaoS
        'Last Modification: 6/4/2022
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
        
            Dim PosX As Integer
            Dim PosY As Integer
            Dim tUser As Integer
        
102         PosX = Reader.ReadString16()
103         PosY = Reader.ReadString16()

            If IsValidUserRef(.flags.GMMeSigue) Then
                Call WriteRecievePosSeguimiento(.flags.GMMeSigue.ArrayIndex, posX, posY)
                'CUANDO DESCONECTA SEGUIDOR Y SEGUIDO VER FLAGS
            End If
            
        End With

        Exit Sub

ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
140

End Sub

' Handles the "SendPosMovimiento" message.

Private Sub HandleNotifyInventariohechizos(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Martín Trionfetti - HarThaoS
        'Last Modification: 6/6/2022
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
        
            Dim Value As Byte
            Dim hechiSel As Byte
            Dim scrollSel As Byte
        
102         Value = Reader.ReadInt8()
            hechiSel = Reader.ReadInt8()
            scrollSel = Reader.ReadInt8()

            If IsValidUserRef(.flags.GMMeSigue) Then
                Call WriteGetInventarioHechizos(.flags.GMMeSigue.ArrayIndex, value, hechiSel, scrollSel)
            End If
            
        End With

        Exit Sub

ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
140

End Sub
'HarThaoS: Agrego perdón faccionario.


'Lee abajo
'Lee arriba
Private Sub HandlePerdonFaccion(ByVal userindex As Integer)

        On Error GoTo ErrHandler

100     With UserList(userindex)
        
            Dim username As String
            Dim tUser    As t_UserReference
            Dim LoopC    As Byte
        
102         username = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             If UCase$(username) <> "YO" Then
108                 tUser = NameIndex(username)
                Else
110                 Call SetUserRef(tUser, userIndex)
                End If
                
                If Not IsValidUserRef(tUser) Then
                    ' Msg743=Usuario offline.
                    Call WriteLocaleMsg(UserIndex, "743", e_FontTypeNames.FONTTYPE_INFO)
                End If
                
                If UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Armada Or UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Caos Or UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.consejo Or UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.concilio Then
'Msg1189= No puedes perdonar a alguien que ya pertenece a una facción
Call WriteLocaleMsg(UserIndex, "1189", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
                'Si es ciudadano aparte de quitarle las reenlistadas le saco los ciudadanos matados.
                If UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Ciudadano Then
                    If UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados > 0 Or UserList(tUser.ArrayIndex).Faccion.Reenlistadas > 0 Then
                        UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados = 0
                        UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 0
                        UserList(tUser.ArrayIndex).Faccion.RecibioArmaduraReal = 0
'Msg1190= Has sido perdonado.
Call WriteLocaleMsg(tUser.ArrayIndex, "1190", e_FontTypeNames.FONTTYPE_INFO)
'Msg1191= Has perdonado a ¬1
Call WriteLocaleMsg(userIndex, "1191", e_FontTypeNames.FONTTYPE_INFO, UserList(tUser.ArrayIndex).name)
                    Else
'Msg1192= No necesitas ser perdonado.
Call WriteLocaleMsg(tUser.ArrayIndex, "1192", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                ElseIf UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Criminal Then
                    If UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 0 Then
'Msg1193= No necesitas ser perdonado.
Call WriteLocaleMsg(tUser.ArrayIndex, "1193", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    Else
                        UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 0
                        UserList(tUser.ArrayIndex).Faccion.RecibioArmaduraCaos = 0
'Msg1194= Has sido perdonado.
Call WriteLocaleMsg(tUser.ArrayIndex, "1194", e_FontTypeNames.FONTTYPE_INFO)
'Msg1195= Has perdonado a ¬1
Call WriteLocaleMsg(userIndex, "1195", e_FontTypeNames.FONTTYPE_INFO, UserList(tUser.ArrayIndex).name)
                    End If
                End If
            Else
136             Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePerdonFaccion", Erl)
140

End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim GuildName As String
            Dim tGuild    As Integer
        
102         GuildName = Reader.ReadString8()
        
104         If (InStrB(GuildName, "+") <> 0) Then
106             GuildName = Replace(GuildName, "+", " ")
            End If
        
108         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
110             tGuild = GuildIndex(GuildName)
            
112             If tGuild > 0 Then
'Msg1196= Clan ¬1
Call WriteLocaleMsg(userindex, "1196", e_FontTypeNames.FONTTYPE_INFO, UCase$(GuildName))
                End If
            Else
116             Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With

        Exit Sub

ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildOnlineMembers", Erl)
120

End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Or .Faccion.Status = e_Facciones.consejo Then
106             Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("[ARMADA REAL] " & UserList(UserIndex).name & "> " & message, e_FontTypeNames.FONTTYPE_CONSEJO))
            End If

        End With

        Exit Sub

ErrHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoyalArmyMessage", Erl)
110

End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/29/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Or .Faccion.Status = e_Facciones.concilio Then
106             Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("[FUERZAS DEL CAOS] " & UserList(UserIndex).name & "> " & message, e_FontTypeNames.FONTTYPE_CONSEJOCAOS))
            End If

        End With

        Exit Sub

ErrHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosLegionMessage", Erl)
110

End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
    
            Dim UserName As String
            Dim tUser    As t_UserReference
            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             tUser = NameIndex(UserName)

108             If Not IsValidUserRef(tUser) Then
'Msg1197= Usuario offline
Call WriteLocaleMsg(UserIndex, "1197", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                Else
                    If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                        If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) <> e_ALINEACION_GUILD.ALINEACION_ARMADA Then
'Msg1198= El miembro no puede ingresar al consejo porque forma parte de un clan que no es de la armada.
Call WriteLocaleMsg(UserIndex, "1198", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
            
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue aceptado en el honorable Consejo Real de Banderbill.", e_FontTypeNames.FONTTYPE_CONSEJO))

114                 With UserList(tUser.ArrayIndex)
                        .Faccion.Status = e_Facciones.consejo
120                     Call WarpUserChar(tUser.ArrayIndex, .pos.map, .pos.x, .pos.y, False)

                    End With
                End If
            End If
        End With
        Exit Sub
ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptRoyalCouncilMember", Erl)
124

End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As t_UserReference
            Dim LoopC    As Byte
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             tUser = NameIndex(UserName)

108             If Not IsValidUserRef(tUser) Then
'Msg1199= Usuario offline
Call WriteLocaleMsg(UserIndex, "1199", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                        If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) <> e_ALINEACION_GUILD.ALINEACION_CAOTICA Then
'Msg1200= El miembro no puede ingresar al concilio porque forma parte de un clan que no es caótico.
Call WriteLocaleMsg(UserIndex, "1200", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
                    
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(username & " fue aceptado en el Consejo de la Legión Oscura.", e_FontTypeNames.FONTTYPE_CONSEJOCAOS))
                
114                 With UserList(tUser.ArrayIndex)
                        .Faccion.Status = e_Facciones.concilio
120                     Call WarpUserChar(tUser.ArrayIndex, .pos.map, .pos.x, .pos.y, False)

                    End With

                End If

            End If

        End With

        Exit Sub

ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptChaosCouncilMember", Erl)
124

End Sub

''
' Handles the "CouncilKick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
  
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
 
            Dim UserName As String
            Dim tUser    As t_UserReference
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             tUser = NameIndex(UserName)

108             If Not IsValidUserRef(tUser) Then
110                 If PersonajeExiste(UserName) Then
'Msg1201= Usuario offline, echando de los consejos
Call WriteLocaleMsg(UserIndex, "1201", e_FontTypeNames.FONTTYPE_INFO)
                        Dim Status As Integer
                        
                        Status = GetDBValue("user", "status", "name", username)
116                     Call EcharConsejoDatabase(username, IIf(Status = 4, 2, 3))
'Msg1202= Usuario ¬1
Call WriteLocaleMsg(userindex, "1202", e_FontTypeNames.FONTTYPE_INFO, username)
                    Else
'Msg1203= No existe el personaje.
Call WriteLocaleMsg(UserIndex, "1203", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
124                 With UserList(tUser.ArrayIndex)
                        If .Faccion.status = e_Facciones.consejo Then
'Msg1204= Has sido echado del consejo de Banderbill
Call WriteLocaleMsg(tUser.ArrayIndex, "1204", e_FontTypeNames.FONTTYPE_INFO)
130                         .Faccion.status = e_Facciones.Armada
132                         Call WarpUserChar(tUser.ArrayIndex, .pos.map, .pos.x, .pos.y)
134                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserName & " fue expulsado del consejo de Banderbill", e_FontTypeNames.FONTTYPE_CONSEJO))
                        End If
                    
                        If .Faccion.Status = e_Facciones.concilio Then
'Msg1205= Has sido echado del consejo de la Legión Oscura
Call WriteLocaleMsg(tUser.ArrayIndex, "1205", e_FontTypeNames.FONTTYPE_INFO)
140                        .Faccion.Status = e_Facciones.Caos
142                         Call WarpUserChar(tUser.ArrayIndex, .pos.map, .pos.x, .pos.y)
144                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(username & " fue expulsado del consejo de la Legión Oscura", e_FontTypeNames.FONTTYPE_CONSEJOCAOS))
                        End If
                        Call RefreshCharStatus(tUser.ArrayIndex)
                    End With
                End If
            End If
        End With
        Exit Sub
ErrHandler:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCouncilKick", Erl)
148

End Sub

''
' Handles the "GuildBan" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)
 
            Dim GuildName   As String
            Dim cantMembers As Integer
            Dim LoopC       As Long
            Dim member      As String
            Dim Count       As Byte
            Dim tUser       As t_UserReference
            Dim tFile       As String
        
102         GuildName = Reader.ReadString8()
        
104         If (Not .flags.Privilegios And e_PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
106             tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
108             If Not FileExist(tFile) Then
'Msg1206= No existe el clan: ¬1
Call WriteLocaleMsg(UserIndex, "1206", e_FontTypeNames.FONTTYPE_INFO, GuildName)
                Else
112                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.Name & " banned al clan " & UCase$(GuildName), e_FontTypeNames.FONTTYPE_FIGHT))
                    'baneamos a los miembros
114                 Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
116                 cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
118                 For LoopC = 1 To cantMembers
                        'member es la victima
120                     member = GetVar(tFile, "Members", "Member" & LoopC)
122                     Call Ban(member, "Administracion del servidor", "Clan Banned")
124                     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", e_FontTypeNames.FONTTYPE_FIGHT))
126                     tUser = NameIndex(member)
128                     If IsValidUserRef(tUser) Then
                            'esta online
130                         UserList(tUser.ArrayIndex).flags.Ban = 1
132                         Call CloseSocket(tUser.ArrayIndex)
                        End If
136                     Call SaveBanDatabase(member, .Name & " - BAN AL CLAN: " & GuildName & ". " & Date & " " & Time, .Name)
150                 Next LoopC
                End If
            End If
        End With
        Exit Sub

ErrHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGuildBan", Erl)
154

End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
 
            Dim UserName As String
            Dim tUser    As t_UserReference
        
102         UserName = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             If (InStrB(UserName, "\") <> 0) Then
108                 UserName = Replace(UserName, "\", "")

                End If

110             If (InStrB(UserName, "/") <> 0) Then
112                 UserName = Replace(UserName, "/", "")

                End If

114             tUser = NameIndex(UserName)
            
116             Call LogGM(.Name, "ECHO DEL CAOS A: " & UserName)
    
118             If IsValidUserRef(tUser) Then
                    If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                        If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA Then
'Msg1207= El usuario ¬1
Call WriteLocaleMsg(UserIndex, "1207", e_FontTypeNames.FONTTYPE_INFO, username)
                            Exit Sub
                        End If
                    Else
122                     UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 2
                        UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Criminal
124                     Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas del caos y prohibida la reenlistada", e_FontTypeNames.FONTTYPE_INFO)
126                     Call WriteConsoleMsg(tUser.ArrayIndex, .name & " te ha expulsado en forma definitiva de las fuerzas del caos.", e_FontTypeNames.FONTTYPE_FIGHT)
                    End If
                Else
                    If PersonajeExiste(username) Then
'Msg1208= Usuario offline, echando de la facción
Call WriteLocaleMsg(UserIndex, "1208", e_FontTypeNames.FONTTYPE_INFO)
                        
                        
                        Dim Status As Integer
                        Status = GetDBValue("user", "status", "name", username)
                        
                        If Status = e_Facciones.Caos Then
                            Call EcharLegionDatabase(username)
'Msg1209= Usuario ¬1
Call WriteLocaleMsg(userindex, "1209", e_FontTypeNames.FONTTYPE_INFO, username)
                        Else
'Msg1210= El personaje no pertenece a la legión.
Call WriteLocaleMsg(UserIndex, "1210", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                        
                    Else
'Msg1211= No existe el personaje.
Call WriteLocaleMsg(userindex, "1211", e_FontTypeNames.FONTTYPE_INFO)

                    End If
                End If

            End If

        End With

        Exit Sub

ErrHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosLegionKick", Erl)
146

End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Last Modification: 12/30/06
        '
        '***************************************************
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As t_UserReference
        
102         UserName = Reader.ReadString8()
            'HarThaoS: Comando roto / revisar.
            'Exit Sub
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             If (InStrB(UserName, "\") <> 0) Then
108                 UserName = Replace(UserName, "\", "")
                End If

110             If (InStrB(UserName, "/") <> 0) Then
112                 UserName = Replace(UserName, "/", "")
                End If

114             tUser = NameIndex(UserName)
            
116             Call LogGM(.Name, "ECHO DE LA REAL A: " & UserName)
            
118             If IsValidUserRef(tUser) Then
                    If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                        If GuildAlignmentIndex(UserList(tUser.ArrayIndex).GuildIndex) = e_ALINEACION_GUILD.ALINEACION_ARMADA Then
'Msg1212= El usuario ¬1
Call WriteLocaleMsg(UserIndex, "1212", e_FontTypeNames.FONTTYPE_INFO, username)
                            Exit Sub
                        End If
                    Else
122                     UserList(tUser.ArrayIndex).Faccion.Reenlistadas = 2
                        UserList(tUser.ArrayIndex).Faccion.Status = e_Facciones.Ciudadano
124                     Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas reales y prohibida la reenlistada", e_FontTypeNames.FONTTYPE_INFO)
126                     Call WriteConsoleMsg(tUser.ArrayIndex, .name & " te ha expulsado en forma definitiva de las fuerzas reales.", e_FontTypeNames.FONTTYPE_FIGHT)
                    End If

                Else
                    If PersonajeExiste(username) Then
'Msg1213= Usuario offline, echando de la facción
Call WriteLocaleMsg(UserIndex, "1213", e_FontTypeNames.FONTTYPE_INFO)
                                                
                        Dim Status As Integer
                        Status = GetDBValue("user", "status", "name", username)
                        
                        If Status = e_Facciones.Armada Then
                            Call EcharArmadaDatabase(username)
'Msg1214= Usuario ¬1
Call WriteLocaleMsg(userindex, "1214", e_FontTypeNames.FONTTYPE_INFO, username)
                        Else
'Msg1215= El personaje no pertenece a la armada.
Call WriteLocaleMsg(userindex, "1215", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
'Msg1216= No existe el personaje.
Call WriteLocaleMsg(userindex, "1216", e_FontTypeNames.FONTTYPE_INFO)

                    End If
                End If

            End If

        End With

        Exit Sub

ErrHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoyalArmyKick", Erl)
146

End Sub

''
' Handles the "ChatColor" message.
'
' @param    UserIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)
        
        On Error GoTo HandleChatColor_Err

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/23/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Change the user`s chat color
        '***************************************************

100     With UserList(UserIndex)

            Dim Color As Long
102             Color = RGB(Reader.ReadInt8(), Reader.ReadInt8(), Reader.ReadInt8())
        
104         If EsGM(UserIndex) Then
106             .flags.ChatColor = Color
            End If

        End With
        
        Exit Sub

HandleChatColor_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChatColor", Erl)
110
        
End Sub



Public Sub HandleDonateGold(ByVal UserIndex As Integer)
        
        On Error GoTo handle

100     With UserList(UserIndex)
        
        

            Dim Oro As Long
102         Oro = Reader.ReadInt32

104         If Oro <= 0 Then Exit Sub
        
            'Se asegura que el target es un npc
106         If Not IsValidNpcRef(.flags.TargetNPC) Then
'Msg1217= Primero tenés que seleccionar al sacerdote.
Call WriteLocaleMsg(UserIndex, "1217", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim priest As t_Npc
110         priest = NpcList(.flags.TargetNPC.ArrayIndex)

            'Validate NPC is an actual priest and the player is not dead
112         If (priest.NPCtype <> e_NPCType.Revividor And (priest.NPCtype <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub

            'Make sure it's close enough
114         If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 3 Then
116             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

118         If .Faccion.Status = e_Facciones.Ciudadano Or .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Or .Faccion.Status = e_Facciones.concilio Or .Faccion.Status = e_Facciones.Caos Then
120             Call WriteChatOverHead(UserIndex, "No puedo aceptar tu donación en este momento...", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                Exit Sub
            End If

122         If .GuildIndex <> 0 Then
124             If modGuilds.Alineacion(.GuildIndex) = 1 Then
126                 Call WriteChatOverHead(UserIndex, "Te encuentras en un clan criminal... no puedo aceptar tu donación.", priest.Char.charindex, vbWhite)
                    Exit Sub
                End If
            End If

128         If .Stats.GLD < Oro Then
'Msg1218= No tienes suficiente dinero.
Call WriteLocaleMsg(UserIndex, "1218", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            Dim Donacion As Long
            If .Faccion.ciudadanosMatados > 0 Then
132             Donacion = .Faccion.ciudadanosMatados * OroMult * CostoPerdonPorCiudadano
            Else
                Donacion = CostoPerdonPorCiudadano / 2
            End If
            
134         If Oro < Donacion Then
136             Call WriteChatOverHead(UserIndex, "Dios no puede perdonarte si eres una persona avara.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                Exit Sub
            End If

138         .Stats.GLD = .Stats.GLD - Oro
140         Call WriteUpdateGold(UserIndex)
'Msg1219= Has donado ¬1
Call WriteLocaleMsg(UserIndex, "1219", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Oro))
144         Call WriteChatOverHead(UserIndex, "¡Gracias por tu generosa donación! Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!", NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbYellow)
146         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, "80", 100, False))
148         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
150         Call VolverCiudadano(UserIndex)
        End With
        Exit Sub

handle:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDonateGold", Erl)
154
        
End Sub

Public Sub HandlePromedio(ByVal UserIndex As Integer)
        
        On Error GoTo handle

100     With UserList(UserIndex)

102         Call WriteConsoleMsg(UserIndex, ListaClases(.clase) & " " & ListaRazas(.raza) & " nivel " & .Stats.ELV & ".", FONTTYPE_INFOBOLD)
            
            Dim Promedio As Double, Vida As Long
        
104         Promedio = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
106         Vida = 18 + ModRaza(.raza).Constitucion + Promedio * (.Stats.ELV - 1)

'Msg1220= Vida esperada: ¬1
Call WriteLocaleMsg(UserIndex, "1220", e_FontTypeNames.FONTTYPE_INFO, Vida)

110         Promedio = CalcularPromedioVida(UserIndex)

            Dim Diff As Long, Color As e_FontTypeNames, Signo As String
            
112         Diff = .Stats.MaxHp - Vida
            
114         If Diff < 0 Then
116             Color = FONTTYPE_PROMEDIO_MENOR
118             Signo = "-"

120         ElseIf Diff > 0 Then
122             Color = FONTTYPE_PROMEDIO_MAYOR
124             Signo = "+"

            Else
126             Color = FONTTYPE_PROMEDIO_IGUAL
128             Signo = "+"
                
            End If

'Msg1221= Vida actual: ¬1
Call WriteLocaleMsg(UserIndex, "1221", e_FontTypeNames.FONTTYPE_INFO, .Stats.MaxHp)

        End With
        
        Exit Sub

handle:
132     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePromedio", Erl)
134
        
End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
        'Allows admins to read guild messages
        '***************************************************
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim guild As String
102             guild = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
106             Call modGuilds.GMEscuchaClan(UserIndex, guild)
                Call LogGM(.name, .name & " espia a " & guild)
            End If

        End With

        Exit Sub

ErrHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
110

End Sub

''
' Handle the "DoBackUp" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
        
        On Error GoTo HandleDoBackUp_Err

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/24/06
        'Show guilds messages
        '***************************************************
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub

104         Call LogGM(.Name, .Name & " ha hecho un backup")
        
106         Call ES.DoBackUp 'Sino lo confunde con la id del paquete

        End With
        
        Exit Sub

HandleDoBackUp_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoBackUp", Erl)
110
        
End Sub

''
' Handle the "NavigateToggle" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleNavigateToggle_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/12/07
        '
        '***************************************************
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             'Msg528=Servidor » Comando deshabilitado para tu cargo.
                Call WriteLocaleMsg(UserIndex, "528", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        
106         If .flags.Navegando = 1 Then
108             .flags.Navegando = 0
            
            Else
110             .flags.Navegando = 1

            End If
        
            'Tell the client that we are navigating.
112         Call WriteNavigateToggle(UserIndex, .flags.Navegando)

        End With
        
        Exit Sub

HandleNavigateToggle_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNavigateToggle", Erl)
116
        
End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)
        
        On Error GoTo HandleServerOpenToUsersToggle_Err

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/24/06
        '
        '***************************************************
100     With UserList(UserIndex)
        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         If ServerSoloGMs > 0 Then
'Msg1222= Servidor habilitado para todos.
Call WriteLocaleMsg(UserIndex, "1222", e_FontTypeNames.FONTTYPE_INFO)
108             ServerSoloGMs = 0
            
            Else
'Msg1223= Servidor restringido a administradores.
Call WriteLocaleMsg(UserIndex, "1223", e_FontTypeNames.FONTTYPE_INFO)
112             ServerSoloGMs = 1

            End If

        End With
        
        Exit Sub

HandleServerOpenToUsersToggle_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerOpenToUsersToggle", Erl)
116
        
End Sub

Public Sub HandleParticipar(ByVal UserIndex As Integer)
        On Error GoTo HandleParticipar_Err

        Dim handle As Integer
        Dim RoomId As Integer
        Dim Password As String
100     RoomId = Reader.ReadInt16
102     Password = Reader.ReadString8
    
104     With UserList(UserIndex)
106         If RoomId = -1 Then
108             If CurrentActiveEventType = CaptureTheFlag Then
110                 If Not InstanciaCaptura Is Nothing Then
112                     Call InstanciaCaptura.inscribirse(UserIndex)
                        Exit Sub
                    End If
                Else
114                 RoomId = GlobalLobbyIndex
                End If
            End If
        
116         If LobbyList(RoomId).State = AcceptingPlayers Then
118             If LobbyList(RoomId).IsPublic Then
                    Dim addPlayerResult As t_response
120                 addPlayerResult = ModLobby.AddPlayerOrGroup(LobbyList(RoomId), UserIndex, Password)
122                 Call WriteLocaleMsg(UserIndex, addPlayerResult.Message, e_FontTypeNames.FONTTYPE_INFO)
                Else
124                 Call WriteLocaleMsg(UserIndex, MsgCantJoinPrivateLobby, e_FontTypeNames.FONTTYPE_INFO)
                End If
                Exit Sub
            End If
        End With
        Exit Sub
HandleParticipar_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleParticipar", Erl)
End Sub

''
' Handle the "ResetFactions" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/26/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)
 
            Dim UserName As String
            Dim tUser    As t_UserReference
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             Call LogGM(.Name, "/RAJAR " & UserName)
            
108             tUser = NameIndex(UserName)
            
110             If IsValidUserRef(tUser) Then Call ResetFacciones(tUser.ArrayIndex)

            End If

        End With

        Exit Sub

ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetFactions", Erl)
114

End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 12/26/06
        '
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)
 
            Dim UserName   As String
            Dim GuildIndex As Integer
        
102         UserName = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, "/RAJARCLAN " & UserName)
            
108             GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, UserName)
            
110             If GuildIndex = 0 Then
'Msg1224= No pertenece a ningún clan o es fundador.
Call WriteLocaleMsg(UserIndex, "1224", e_FontTypeNames.FONTTYPE_INFO)
                Else
'Msg1225= Expulsado.
Call WriteLocaleMsg(UserIndex, "1225", e_FontTypeNames.FONTTYPE_INFO)
116                 Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(UserName & " ha sido expulsado del clan por los administradores del servidor.", e_FontTypeNames.FONTTYPE_GUILD))

                End If

            End If

        End With

        Exit Sub

ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveCharFromGuild", Erl)
120

End Sub

''
' Handle the "SystemMessage" message
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last Modification: 12/29/06
        'Send a message to all the users
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Message As String
102             Message = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.Name, "Mensaje de sistema:" & Message)
            
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageShowMessageBox(Message))

            End If

        End With

        Exit Sub

ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSystemMessage", Erl)
112

End Sub

Private Sub HandleOfertaInicial(ByVal UserIndex As Integer)
        
        On Error GoTo HandleOfertaInicial_Err
    
100     With UserList(UserIndex)

            Dim Oferta As Long
102             Oferta = Reader.ReadInt32()
        
104         If UserList(UserIndex).flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

108         If Not IsValidNpcRef(.flags.TargetNPC) Then
'Msg1226= Primero tenés que hacer click sobre el subastador.
Call WriteLocaleMsg(UserIndex, "1226", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

112         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Subastador Then
'Msg1227= Primero tenés que hacer click sobre el subastador.
Call WriteLocaleMsg(UserIndex, "1227", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
116         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 2 Then
118             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
120         If .flags.Subastando = False Then
122             Call WriteChatOverHead(UserIndex, "Oye amigo, tu no podés decirme cual es la oferta inicial.", NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                Exit Sub
            End If
        
124         If Subasta.HaySubastaActiva = False And .flags.Subastando = False Then
'Msg1228= No hay ninguna subasta en curso.
Call WriteLocaleMsg(UserIndex, "1228", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
128         If .flags.Subastando = True Then
130             UserList(UserIndex).Counters.TiempoParaSubastar = 0
132             Subasta.OfertaInicial = Oferta
134             Subasta.MejorOferta = 0
136             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " está subastando: " & ObjData(Subasta.ObjSubastado).name & " (Cantidad: " & Subasta.ObjSubastadoCantidad & " ) - con un precio inicial de " & PonerPuntos(Subasta.OfertaInicial) & " monedas. Escribe /OFERTAR (cantidad) para participar.", e_FontTypeNames.FONTTYPE_SUBASTA))
138             .flags.Subastando = False
140             Subasta.HaySubastaActiva = True
142             Subasta.Subastador = .Name
144             Subasta.MinutosDeSubasta = 5
146             Subasta.TiempoRestanteSubasta = 300
148             Call LogearEventoDeSubasta("#################################################################################################################################################################################################")
150             Call LogearEventoDeSubasta("El dia: " & Date & " a las " & Time)
152             Call LogearEventoDeSubasta(.Name & ": Esta subastando el item numero " & Subasta.ObjSubastado & " con una cantidad de " & Subasta.ObjSubastadoCantidad & " y con un precio inicial de " & PonerPuntos(Subasta.OfertaInicial) & " monedas.")
154             frmMain.SubastaTimer.Enabled = True
156             Call WarpUserChar(UserIndex, 14, 27, 64, True)

                'lalala toda la bola de los timerrr
            End If

        End With
        
        Exit Sub

HandleOfertaInicial_Err:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOfertaInicial", Erl)
160
        
End Sub

Private Sub HandleOfertaDeSubasta(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Oferta   As Long
            Dim ExOferta As t_UserReference
        
102         Oferta = Reader.ReadInt32()
        
104         If Subasta.HaySubastaActiva = False Then
'Msg1229= No hay ninguna subasta en curso.
Call WriteLocaleMsg(UserIndex, "1229", e_FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
               
108         If Oferta < Subasta.MejorOferta + 100 Then
'Msg1230= Debe haber almenos una diferencia de 100 monedas a la ultima oferta!
Call WriteLocaleMsg(UserIndex, "1230", e_FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
        
112         If .Name = Subasta.Subastador Then
'Msg1231= No podés auto ofertar en tus subastas. La proxima vez iras a la carcel...
Call WriteLocaleMsg(UserIndex, "1231", e_FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
        
116         If .Stats.GLD >= Oferta Then

                'revisar que pasa si el usuario que oferto antes esta offline
                'Devolvemos el oro al usuario que oferto antes...(si es que hubo oferta)
118             If Subasta.HuboOferta = True Then
120                 ExOferta = NameIndex(Subasta.Comprador)
122                 UserList(ExOferta.ArrayIndex).Stats.GLD = UserList(ExOferta.ArrayIndex).Stats.GLD + Subasta.MejorOferta
124                 Call WriteUpdateGold(ExOferta.ArrayIndex)
                End If
            
126             Subasta.MejorOferta = Oferta
128             Subasta.Comprador = .Name
            
130             .Stats.GLD = .Stats.GLD - Oferta
132             Call WriteUpdateGold(UserIndex)
            
134             If Subasta.TiempoRestanteSubasta < 60 Then
136                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Oferta mejorada por: " & .name & " (Ofrece " & PonerPuntos(Oferta) & " monedas de oro) - Tiempo Extendido. Escribe /SUBASTA para mas información.", e_FontTypeNames.FONTTYPE_SUBASTA))
138                 Call LogearEventoDeSubasta(.Name & ": Mejoro la oferta en el ultimo minuto ofreciendo " & PonerPuntos(Oferta) & " monedas.")
140                 Subasta.TiempoRestanteSubasta = Subasta.TiempoRestanteSubasta + 30
                Else
144                 Call LogearEventoDeSubasta(.Name & ": Mejoro la oferta ofreciendo " & PonerPuntos(Oferta) & " monedas.")
146                 Subasta.HuboOferta = True
148                 Subasta.PosibleCancelo = False

                End If

            Else
'Msg1232= No posees esa cantidad de oro.
Call WriteLocaleMsg(UserIndex, "1232", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
    
        Exit Sub

ErrHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
154

End Sub

Public Sub HandleDuel(ByVal UserIndex As Integer)
    
        On Error GoTo ErrHandler
        
        Dim Players         As String
        Dim Bet             As Long
        Dim PocionesMaximas As Integer
        Dim CaenItems       As Boolean

100     With UserList(UserIndex)

102         Players = Reader.ReadString8
104         Bet = Reader.ReadInt32
106         PocionesMaximas = Reader.ReadInt16
108         CaenItems = Reader.ReadBool
'Msg1233= No puedes realizar un reto en este momento.
Call WriteLocaleMsg(UserIndex, "1233", e_FontTypeNames.FONTTYPE_INFO)
           'Exit Sub
110         Call CrearReto(UserIndex, Players, Bet, PocionesMaximas, CaenItems)

        End With
    
        Exit Sub
    
ErrHandler:

112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDuel", Erl)
114

End Sub

Private Sub HandleAcceptDuel(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
        
        Dim Offerer As String

100     With UserList(UserIndex)

102         Offerer = Reader.ReadString8

104         Call AceptarReto(UserIndex, Offerer)

        End With
    
        Exit Sub
    
ErrHandler:

106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAcceptDuel", Erl)
108

End Sub

Private Sub HandleCancelDuel(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         Reader.ReadInt16

104         If .flags.SolicitudReto.estado <> e_SolicitudRetoEstado.Libre Then
106             Call CancelarSolicitudReto(UserIndex, .Name & " ha cancelado la solicitud.")

108         ElseIf IsValidUserRef(.flags.AceptoReto) Then
110             Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .name & " ha cancelado su admisión.")

            End If

        End With

End Sub

Private Sub HandleQuitDuel(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If .flags.EnReto Then
104             Call AbandonarReto(UserIndex)
            End If

        End With

End Sub

Private Sub HandleTransFerGold(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim Cantidad As Long
            Dim tUser    As t_UserReference
        
102         Cantidad = Reader.ReadInt32()
104         UserName = Reader.ReadString8()

            ' WyroX: Chequeos de seguridad... Estos chequeos ya se hacen en el cliente, pero si no se hacen se puede duplicar oro...

            ' Cantidad válida?
106         If Cantidad <= 0 Then Exit Sub

            ' Tiene el oro?
108         If .Stats.Banco < Cantidad Then Exit Sub
            
110         If .flags.Muerto = 1 Then
112             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                ''Msg77=¡¡Estás muerto!!.)
                Exit Sub

            End If
        
            'Validate target NPC
114         If Not IsValidNpcRef(.flags.TargetNPC) Then
'Msg1234= Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.
Call WriteLocaleMsg(UserIndex, "1234", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

118         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then Exit Sub
            
120         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
122             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
124         tUser = NameIndex(UserName)
            ' Enviar a vos mismo?
126         If tUser.ArrayIndex = userIndex Then
128             Call WriteChatOverHead(UserIndex, "¡No puedo enviarte oro a vos mismo!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                Exit Sub
            End If
    
130         If Not EsGM(userindex) Then
132             If Not IsValidUserRef(tUser) Then
                    If GetTickCount() - .Counters.LastTransferGold >= 10000 Then
                        If PersonajeExiste(username) Then
136                         If Not AddOroBancoDatabase(username, Cantidad) Then
138                             Call WriteChatOverHead(UserIndex, "Error al realizar la operación.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                                Exit Sub
                            Else
150                             UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
                            End If
                            .Counters.LastTransferGold = GetTickCount()
                        Else
                            Call WriteChatOverHead(UserIndex, "El usuario no existe.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                            Exit Sub
                        End If
                    Else
                        Call WriteChatOverHead(UserIndex, "Espera un momento.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                        Exit Sub
                    End If
                Else
                 UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(Cantidad) 'Quitamos el oro al usuario
                 UserList(tUser.ArrayIndex).Stats.Banco = UserList(tUser.ArrayIndex).Stats.Banco + val(Cantidad) 'Se lo damos al otro.
                End If
152             Call WriteChatOverHead(UserIndex, "¡El envío se ha realizado con éxito! Gracias por utilizar los servicios de Finanzas Goliath", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
            Else
156             Call WriteChatOverHead(UserIndex, "Los administradores no pueden transferir oro.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
158             Call LogGM(.Name, "Quizo transferirle oro a: " & UserName)
            End If
        End With
        Exit Sub

ErrHandler:
160     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
162

End Sub

Private Sub HandleMoveItem(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim SlotViejo As Byte
            Dim SlotNuevo As Byte
        
102         SlotViejo = Reader.ReadInt8()
104         SlotNuevo = Reader.ReadInt8()
        
            Dim Objeto    As t_Obj
            Dim Equipado  As Boolean
            Dim Equipado2 As Boolean
            Dim Equipado3 As Boolean
            Dim ObjCania As t_Obj
            'HarThaoS: Si es un hilo de pesca y lo estoy arrastrando en una caña rota borro del slot viejo y en el nuevo pongo la caña correspondiente
             If SlotViejo > getMaxInventorySlots(UserIndex) Or SlotNuevo > getMaxInventorySlots(UserIndex) Or SlotViejo <= 0 Or SlotNuevo <= 0 Then Exit Sub
            
            If .Invent.Object(SlotViejo).ObjIndex = 2183 Then
            
                Select Case .Invent.Object(SlotNuevo).ObjIndex
                     Case 3457
                        ObjCania.ObjIndex = 881
                    Case 3456
                        ObjCania.ObjIndex = 2121
                    Case 3459
                        ObjCania.ObjIndex = 2132
                    Case 3458
                        ObjCania.ObjIndex = 2133
                End Select
                ObjCania.amount = 1
                'si el objeto que estaba pisando era una caña rota.
                If ObjCania.ObjIndex > 0 Then
                    'Quitamos del inv el item
                    Call QuitarUserInvItem(UserIndex, SlotViejo, 1)
                    Call UpdateUserInv(False, UserIndex, SlotViejo)
                    Call QuitarUserInvItem(UserIndex, SlotNuevo, 1)
                    Call UpdateUserInv(False, UserIndex, SlotNuevo)
                    Call MeterItemEnInventario(UserIndex, ObjCania)
                    Exit Sub
                End If
                
            End If
            
        
106         If (SlotViejo > .CurrentInventorySlots) Or (SlotNuevo > .CurrentInventorySlots) Then
'Msg1235= Espacio no desbloqueado.
Call WriteLocaleMsg(UserIndex, "1235", e_FontTypeNames.FONTTYPE_INFO)
            
            Else
    
110             If .Invent.Object(SlotNuevo).ObjIndex = .Invent.Object(SlotViejo).ObjIndex Then
112                 .Invent.Object(SlotNuevo).amount = .Invent.Object(SlotNuevo).amount + .Invent.Object(SlotViejo).amount
                    
                    Dim Excedente As Integer
114                 Excedente = .Invent.Object(SlotNuevo).amount - MAX_INVENTORY_OBJS

116                 If Excedente > 0 Then
118                     .Invent.Object(SlotViejo).amount = Excedente
120                     .Invent.Object(SlotNuevo).amount = MAX_INVENTORY_OBJS
                    Else

122                     If .Invent.Object(SlotViejo).Equipped = 1 Then
124                         .Invent.Object(SlotNuevo).Equipped = 1

                        End If
                    
126                     .Invent.Object(SlotViejo).ObjIndex = 0
128                     .Invent.Object(SlotViejo).amount = 0
130                     .Invent.Object(SlotViejo).Equipped = 0
                    
                        'Cambiamos si alguno es un anillo
132                     If .Invent.DañoMagicoEqpSlot = SlotViejo Then
134                         .Invent.DañoMagicoEqpSlot = SlotNuevo

                        End If

136                     If .Invent.ResistenciaEqpSlot = SlotViejo Then
138                         .Invent.ResistenciaEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un armor
140                     If .Invent.ArmourEqpSlot = SlotViejo Then
142                         .Invent.ArmourEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un barco
144                     If .Invent.BarcoSlot = SlotViejo Then
146                         .Invent.BarcoSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es una montura
148                     If .Invent.MonturaSlot = SlotViejo Then
150                         .Invent.MonturaSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un casco
152                     If .Invent.CascoEqpSlot = SlotViejo Then
154                         .Invent.CascoEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un escudo
156                     If .Invent.EscudoEqpSlot = SlotViejo Then
158                         .Invent.EscudoEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es munición
160                     If .Invent.MunicionEqpSlot = SlotViejo Then
162                         .Invent.MunicionEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un arma
164                     If .Invent.WeaponEqpSlot = SlotViejo Then
166                         .Invent.WeaponEqpSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es un magico
172                     If .Invent.MagicoSlot = SlotViejo Then
174                         .Invent.MagicoSlot = SlotNuevo

                        End If
                        
                        'Cambiamos si alguno es una herramienta
176                     If .Invent.HerramientaEqpSlot = SlotViejo Then
178                         .Invent.HerramientaEqpSlot = SlotNuevo

                        End If

                    End If
                
                Else

180                 If .Invent.Object(SlotNuevo).ObjIndex <> 0 Then
182                     Objeto.amount = .Invent.Object(SlotViejo).amount
184                     Objeto.ObjIndex = .Invent.Object(SlotViejo).ObjIndex
                    
186                     If .Invent.Object(SlotViejo).Equipped = 1 Then
188                         Equipado = True
    
                        End If
                    
190                     If .Invent.Object(SlotNuevo).Equipped = 1 Then
192                         Equipado2 = True
    
                        End If
                    
                        '  If .Invent.Object(SlotNuevo).Equipped = 1 And .Invent.Object(SlotViejo).Equipped = 1 Then
                        '     Equipado3 = True
                        ' End If
                    
194                     .Invent.Object(SlotViejo).ObjIndex = .Invent.Object(SlotNuevo).ObjIndex
196                     .Invent.Object(SlotViejo).amount = .Invent.Object(SlotNuevo).amount
                    
198                     .Invent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
200                     .Invent.Object(SlotNuevo).amount = Objeto.amount
                    
202                     If Equipado Then
204                         .Invent.Object(SlotNuevo).Equipped = 1
                        Else
206                         .Invent.Object(SlotNuevo).Equipped = 0
    
                        End If
                                    
208                     If Equipado2 Then
210                         .Invent.Object(SlotViejo).Equipped = 1
                        Else
212                         .Invent.Object(SlotViejo).Equipped = 0
    
                        End If
    
                    End If
    
                    'Cambiamos si alguno es un anillo
214                 If .Invent.DañoMagicoEqpSlot = SlotViejo Then
216                     .Invent.DañoMagicoEqpSlot = SlotNuevo
218                 ElseIf .Invent.DañoMagicoEqpSlot = SlotNuevo Then
220                     .Invent.DañoMagicoEqpSlot = SlotViejo

                    End If

222                 If .Invent.ResistenciaEqpSlot = SlotViejo Then
224                     .Invent.ResistenciaEqpSlot = SlotNuevo
226                 ElseIf .Invent.ResistenciaEqpSlot = SlotNuevo Then
228                     .Invent.ResistenciaEqpSlot = SlotViejo

                    End If
                    
                    'Cambiamos si alguno es un armor
230                 If .Invent.ArmourEqpSlot = SlotViejo Then
232                     .Invent.ArmourEqpSlot = SlotNuevo
234                 ElseIf .Invent.ArmourEqpSlot = SlotNuevo Then
236                     .Invent.ArmourEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un barco
238                 If .Invent.BarcoSlot = SlotViejo Then
240                     .Invent.BarcoSlot = SlotNuevo
242                 ElseIf .Invent.BarcoSlot = SlotNuevo Then
244                     .Invent.BarcoSlot = SlotViejo
    
                    End If
                     
                    'Cambiamos si alguno es una montura
246                 If .Invent.MonturaSlot = SlotViejo Then
248                     .Invent.MonturaSlot = SlotNuevo
250                 ElseIf .Invent.MonturaSlot = SlotNuevo Then
252                     .Invent.MonturaSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un casco
254                 If .Invent.CascoEqpSlot = SlotViejo Then
256                     .Invent.CascoEqpSlot = SlotNuevo
258                 ElseIf .Invent.CascoEqpSlot = SlotNuevo Then
260                     .Invent.CascoEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un escudo
262                 If .Invent.EscudoEqpSlot = SlotViejo Then
264                     .Invent.EscudoEqpSlot = SlotNuevo
266                 ElseIf .Invent.EscudoEqpSlot = SlotNuevo Then
268                     .Invent.EscudoEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es munición
270                 If .Invent.MunicionEqpSlot = SlotViejo Then
272                     .Invent.MunicionEqpSlot = SlotNuevo
274                 ElseIf .Invent.MunicionEqpSlot = SlotNuevo Then
276                     .Invent.MunicionEqpSlot = SlotViejo
    
                    End If
                    
                    'Cambiamos si alguno es un arma
278                 If .Invent.WeaponEqpSlot = SlotViejo Then
280                     .Invent.WeaponEqpSlot = SlotNuevo
282                 ElseIf .Invent.WeaponEqpSlot = SlotNuevo Then
284                     .Invent.WeaponEqpSlot = SlotViejo
    
                    End If
                     
                    'Cambiamos si alguno es un magico
294                 If .Invent.MagicoSlot = SlotViejo Then
296                     .Invent.MagicoSlot = SlotNuevo
298                 ElseIf .Invent.MagicoSlot = SlotNuevo Then
300                     .Invent.MagicoSlot = SlotViejo
    
                    End If
                     
                    'Cambiamos si alguno es una herramienta
302                 If .Invent.HerramientaEqpSlot = SlotViejo Then
304                     .Invent.HerramientaEqpSlot = SlotNuevo
306                 ElseIf .Invent.HerramientaEqpSlot = SlotNuevo Then
308                     .Invent.HerramientaEqpSlot = SlotViejo
    
                    End If
                
310                 If Objeto.ObjIndex = 0 Then
312                     .Invent.Object(SlotNuevo).ObjIndex = .Invent.Object(SlotViejo).ObjIndex
314                     .Invent.Object(SlotNuevo).amount = .Invent.Object(SlotViejo).amount
316                     .Invent.Object(SlotNuevo).Equipped = .Invent.Object(SlotViejo).Equipped
                            
318                     .Invent.Object(SlotViejo).ObjIndex = 0
320                     .Invent.Object(SlotViejo).amount = 0
322                     .Invent.Object(SlotViejo).Equipped = 0
    
                    End If
                    
                End If
                
324             Call UpdateUserInv(False, UserIndex, SlotViejo)
326             Call UpdateUserInv(False, UserIndex, SlotNuevo)

            End If
            
            If IsValidUserRef(.flags.GMMeSigue) Then
                UserList(.flags.GMMeSigue.ArrayIndex).Invent = UserList(UserIndex).Invent
                Call UpdateUserInv(True, UserIndex, 1)
            End If


        End With
    
        Exit Sub

ErrHandler:
328     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveItem", Erl)
330

End Sub

Private Sub HandleBovedaMoveItem(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim SlotViejo As Byte
            Dim SlotNuevo As Byte
        
102         SlotViejo = Reader.ReadInt8()
104         SlotNuevo = Reader.ReadInt8()
        
            Dim Objeto    As t_Obj
            Dim Equipado  As Boolean
            Dim Equipado2 As Boolean
            Dim Equipado3 As Boolean
        
            If SlotViejo > MAX_BANCOINVENTORY_SLOTS Or SlotNuevo > MAX_BANCOINVENTORY_SLOTS Or SlotViejo <= 0 Or SlotNuevo <= 0 Then Exit Sub
106         Objeto.ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex
108         Objeto.amount = UserList(UserIndex).BancoInvent.Object(SlotViejo).amount
        
110         UserList(UserIndex).BancoInvent.Object(SlotViejo).ObjIndex = UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex
112         UserList(UserIndex).BancoInvent.Object(SlotViejo).amount = UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount
         
114         UserList(UserIndex).BancoInvent.Object(SlotNuevo).ObjIndex = Objeto.ObjIndex
116         UserList(UserIndex).BancoInvent.Object(SlotNuevo).amount = Objeto.amount
    
            'Actualizamos el banco
118         Call UpdateBanUserInv(False, UserIndex, SlotViejo, "HandleBovedaMoveItem - slot viejo")
120         Call UpdateBanUserInv(False, UserIndex, SlotNuevo, "HandleBovedaMoveItem - slot nuevo")

        End With
    
        Exit Sub
    
        Exit Sub

ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBovedaMoveItem", Erl)
124

End Sub

Private Sub HandleQuieroFundarClan(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

102         If UserList(UserIndex).flags.Privilegios And e_PlayerType.Consejero Then Exit Sub

104         If UserList(UserIndex).GuildIndex > 0 Then
'Msg1236= Ya perteneces a un clan, no podés fundar otro.
Call WriteLocaleMsg(UserIndex, "1236", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

108         If UserList(UserIndex).Stats.ELV < 23 Or UserList(UserIndex).Stats.UserSkills(e_Skill.liderazgo) < 50 Then
'Msg1237= Para fundar un clan debes ser Nivel 23, tener 50 en liderazgo y tener en tu inventario las 4 Gemas de Fundación: Gema Verde, Gema Roja, Gema Azul y Gema Polar.
Call WriteLocaleMsg(UserIndex, "1237", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

112         If Not TieneObjetos(407, 1, UserIndex) Or Not TieneObjetos(408, 1, UserIndex) Or Not TieneObjetos(409, 1, UserIndex) Or Not TieneObjetos(412, 1, UserIndex) Then
'Msg1238= Para fundar un clan debes tener en tu inventario las 4 Gemas de Fundación: Gema Verde, Gema Roja, Gema Azul y Gema Polar.
Call WriteLocaleMsg(UserIndex, "1238", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

'Msg1239= Servidor » ¡Comenzamos a fundar el clan! Ingresa todos los datos solicitados.
Call WriteLocaleMsg(UserIndex, "1239", e_FontTypeNames.FONTTYPE_INFO)
        
118         Call WriteShowFundarClanForm(UserIndex)

        End With
    
        Exit Sub
    
        Exit Sub

ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuieroFundarClan", Erl)
122

End Sub

Private Sub HandleLlamadadeClan(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim refError   As String
            Dim clan_nivel As Byte

102         If .GuildIndex <> 0 Then
104             clan_nivel = modGuilds.NivelDeClan(.GuildIndex)

106             If clan_nivel >= 2 Then
108                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Clan> [" & .name & "] solicita apoyo de su clan en " & get_map_name(.pos.Map) & " (" & .pos.Map & "-" & .pos.x & "-" & .pos.y & "). Puedes ver su ubicación en el mapa del mundo.", e_FontTypeNames.FONTTYPE_GUILD))
110                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave("43", NO_3D_SOUND, NO_3D_SOUND))
112                 Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageUbicacionLlamada(.Pos.Map, .Pos.X, .Pos.Y))

                Else
'Msg1240= Servidor » El nivel de tu clan debe ser 2 para utilizar esta opción.
Call WriteLocaleMsg(UserIndex, "1240", e_FontTypeNames.FONTTYPE_INFO)

                End If
            End If

        End With
    
        Exit Sub

ErrHandler:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLlamadadeClan", Erl)
118

End Sub

Private Sub HandleCasamiento(ByVal UserIndex As Integer)

        'Author: Pablo Mercavides

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim UserName As String
            Dim tUser    As t_UserReference

102         UserName = Reader.ReadString8()
104         tUser = NameIndex(UserName)
            If Not IsValidUserRef(tUser) Then
                ' Msg743=Usuario offline.
                Call WriteLocaleMsg(UserIndex, "743", e_FontTypeNames.FONTTYPE_INFO)
            End If
106         If IsValidNpcRef(.flags.TargetNPC) Then
108             If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor Then
110                 ' Msg744=Primero haz click sobre un sacerdote.
                    Call WriteLocaleMsg(UserIndex, "744", e_FontTypeNames.FONTTYPE_INFO)
                Else
112                 If Distancia(.Pos, NpcList(.flags.TargetNPC.ArrayIndex).Pos) > 10 Then
114                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                    Else
116                     If tUser.ArrayIndex = userIndex Then
118                         ' Msg745=No podés casarte contigo mismo.
                            Call WriteLocaleMsg(UserIndex, "745", e_FontTypeNames.FONTTYPE_INFO)
120                     ElseIf .flags.Casado = 1 Then
122                         ' Msg746=¡Ya estás casado! Debes divorciarte de tu actual pareja para casarte nuevamente.
                            Call WriteLocaleMsg(UserIndex, "746", e_FontTypeNames.FONTTYPE_INFO)
124                     ElseIf UserList(tUser.ArrayIndex).flags.Casado = 1 Then
126                         ' Msg747=Tu pareja debe divorciarse antes de tomar tu mano en matrimonio.
                            Call WriteLocaleMsg(UserIndex, "747", e_FontTypeNames.FONTTYPE_INFO)
                        Else
132                         If UserList(tUser.ArrayIndex).flags.Candidato.ArrayIndex = userIndex Then
134                             UserList(tUser.ArrayIndex).flags.Casado = 1
136                             UserList(tUser.ArrayIndex).flags.SpouseId = UserList(UserIndex).id
138                             .flags.Casado = 1
140                             .flags.SpouseId = UserList(tUser.ArrayIndex).id
142                             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(e_FXSound.Casamiento_sound, NO_3D_SOUND, NO_3D_SOUND))
144                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El sacerdote de " & get_map_name(.pos.map) & " celebra el casamiento entre " & UserList(userIndex).name & " y " & UserList(tUser.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_WARNING))
146                             Call WriteChatOverHead(UserIndex, "Los declaro unidos en legal matrimonio ¡Felicidades!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
148                             Call WriteChatOverHead(tUser.ArrayIndex, "Los declaro unidos en legal matrimonio ¡Felicidades!", NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                            Else
150                             Call WriteChatOverHead(UserIndex, "La solicitud de casamiento a sido enviada a " & username & ".", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
152                             Call WriteConsoleMsg(tUser.ArrayIndex, .name & " desea casarse contigo, para permitirlo haz click en el sacerdote y escribe /PROPONER " & .name & ".", e_FontTypeNames.FONTTYPE_TALK)
154                             .flags.Candidato = tUser
                            End If
                        End If
                    End If
                End If
            Else
156             ' Msg748=Primero haz click sobre el sacerdote.
                Call WriteLocaleMsg(UserIndex, "748", e_FontTypeNames.FONTTYPE_INFO)

            End If
        End With
        Exit Sub
ErrHandler:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCasamiento", Erl)
160

End Sub



Private Sub HandleComenzarTorneo(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

102         If EsGM(UserIndex) Then

104             Call ComenzarTorneoOk

            End If

        End With
    
        Exit Sub

ErrHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
108

End Sub



Private Sub HandleBusquedaTesoro(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Tipo As Byte
102             Tipo = Reader.ReadInt8()
  
104         If (.flags.Privilegios And Not (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.user)) Then

106             Select Case Tipo

                    Case 0

108                     If Not BusquedaTesoroActiva And BusquedaRegaloActiva = False And BusquedaNpcActiva = False Then
110                         Call PerderTesoro
                        Else

112                         If BusquedaTesoroActiva Then
114                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavia nadie fue capaz de encontar el tesoro, recorda que se encuentra en " & get_map_name(TesoroNumMapa) & "(" & TesoroNumMapa & "). ¿Quien sera el valiente que lo encuentre?", e_FontTypeNames.FONTTYPE_TALK))
'Msg1241= Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: ¬1
Call WriteLocaleMsg(UserIndex, "1241", e_FontTypeNames.FONTTYPE_INFO, TesoroNumMapa)
                            Else
118                             ' Msg734=Ya hay una busqueda del tesoro activa.
                                Call WriteLocaleMsg(UserIndex, "734", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

120                 Case 1

122                     If Not BusquedaRegaloActiva And BusquedaTesoroActiva = False And BusquedaNpcActiva = False Then
124                         Call PerderRegalo
                        Else

126                         If BusquedaRegaloActiva Then
128                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Ningún valiente fue capaz de encontrar el item misterioso, recuerda que se encuentra en " & get_map_name(RegaloNumMapa) & "(" & RegaloNumMapa & "). ¡Ten cuidado!", e_FontTypeNames.FONTTYPE_TALK))
'Msg1242= Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: ¬1
Call WriteLocaleMsg(UserIndex, "1242", e_FontTypeNames.FONTTYPE_INFO, RegaloNumMapa)
                            Else
132                             ' Msg734=Ya hay una busqueda del tesoro activa.
                                Call WriteLocaleMsg(UserIndex, "734", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

134                 Case 2

136                     If Not BusquedaNpcActiva And BusquedaTesoroActiva = False And BusquedaRegaloActiva = False Then
                            Dim Pos As t_WorldPos
138                         Pos.Map = TesoroNPCMapa(RandomNumber(1, UBound(TesoroNPCMapa)))
140                         Pos.Y = 50
142                         Pos.X = 50
144                         npc_index_evento = SpawnNpc(TesoroNPC(RandomNumber(1, UBound(TesoroNPC))), Pos, True, False, True)
146                         BusquedaNpcActiva = True
                        Else

148                         If BusquedaNpcActiva Then
150                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavía nadie logró matar el NPC que se encuentra en el mapa " & NpcList(npc_index_evento).pos.Map & ".", e_FontTypeNames.FONTTYPE_TALK))
'Msg1243= Ya hay una busqueda de npc activo. El tesoro se encuentra en: ¬1
Call WriteLocaleMsg(UserIndex, "1243", e_FontTypeNames.FONTTYPE_INFO, NpcList(npc_index_evento).Pos.Map)
                            Else
154                             ' Msg734=Ya hay una busqueda del tesoro activa.
                                Call WriteLocaleMsg(UserIndex, "734", e_FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                End Select
            Else
156             ' Msg735=Servidor » No estas habilitado para hacer Eventos.
                Call WriteLocaleMsg(UserIndex, "735", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

ErrHandler:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBusquedaTesoro", Erl)
160

End Sub

Private Sub HandleFlagTrabajar(ByVal UserIndex As Integer)
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

102         .Counters.Trabajando = 0
104         .flags.UsandoMacro = False
106         .flags.TargetObj = 0 ' Sacamos el targer del objeto
108         .flags.UltimoMensaje = 0

        End With
    
        Exit Sub

ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
112

End Sub

Private Sub HandleCompletarAccion(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Accion As Byte
102             Accion = Reader.ReadInt8()
        
104         If .Accion.AccionPendiente = True Then
106             If .Accion.TipoAccion = Accion Then
108                 Call CompletarAccionFin(UserIndex)
                Else
110                 ' Msg749=Servidor » La acción que solicitas no se corresponde.
                    Call WriteLocaleMsg(UserIndex, "749", e_FontTypeNames.FONTTYPE_SERVER)

                End If

            Else
112             ' Msg750=Servidor » Tu no tenias ninguna acción pendiente.
                Call WriteLocaleMsg(UserIndex, "750", e_FontTypeNames.FONTTYPE_SERVER)

            End If

        End With
    
        Exit Sub

ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
116

End Sub

Private Sub HandleInvitarGrupo(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If .flags.Muerto = 1 Then
                'Msg77=¡¡Estás muerto!!.
104             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
            
            Else
            
106             If .Grupo.CantidadMiembros <= UBound(.Grupo.Miembros) Then
108                 Call WriteWorkRequestTarget(UserIndex, e_Skill.Grupo)
                Else
110                 ' Msg751=¡No podés invitar a más personas!
                    Call WriteLocaleMsg(UserIndex, "751", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With
        
        Exit Sub

HandleInvitarGrupo_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInvitarGrupo", Erl)
114
    
End Sub

Private Sub HandleMarcaDeClan(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleMarcaDeClan_Err

100     With UserList(UserIndex)
            'Exit sub para anular marca de clan
            Exit Sub
102         If UserList(UserIndex).GuildIndex = 0 Then
                Exit Sub
            End If

104         If .flags.Muerto = 1 Then
                ''Msg77=¡¡Estás muerto!!.
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim clan_nivel As Byte

108         clan_nivel = modGuilds.NivelDeClan(UserList(UserIndex).GuildIndex)

110         If clan_nivel > 20 Then
112             ' Msg721=Servidor » El nivel de tu clan debe ser 3 para utilizar esta opción.
                Call WriteLocaleMsg(UserIndex, "721", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
       
114         Call WriteWorkRequestTarget(UserIndex, e_Skill.MarcaDeClan)
        
        End With
        
        Exit Sub

HandleMarcaDeClan_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMarcaDeClan", Erl)
118
End Sub

Private Sub HandleResponderPregunta(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim respuesta As Boolean
            Dim DeDonde   As String

102         respuesta = Reader.ReadBool()
        
            Dim Log As String

104         Log = "Repuesta "
            UserList(UserIndex).flags.RespondiendoPregunta = False
106         If respuesta Then
        
108             Select Case UserList(UserIndex).flags.pregunta

                    Case 1
110                     Log = "Repuesta Afirmativa 1"
                        If UserList(UserIndex).Grupo.EnGrupo Then
                            Call WriteLocaleMsg(UserIndex, MsgYouAreAlreadyInGroup, e_FontTypeNames.FONTTYPE_INFOIAO)
                            Exit Sub
                        End If
112                     If IsValidUserRef(UserList(userIndex).Grupo.PropuestaDe) Then
114                         If UserList(UserList(userIndex).Grupo.PropuestaDe.ArrayIndex).Grupo.Lider.ArrayIndex <> UserList(userIndex).Grupo.PropuestaDe.ArrayIndex Then
116                             ' Msg722=¡El lider del grupo ha cambiado, imposible unirse!
                                Call WriteLocaleMsg(UserIndex, "722", e_FontTypeNames.FONTTYPE_INFOIAO)
                            Else
118                             Log = "Repuesta Afirmativa 1-1 "
120                             If Not IsValidUserRef(UserList(UserList(userIndex).Grupo.PropuestaDe.ArrayIndex).Grupo.Lider) Then
122                                 ' Msg723=¡El grupo ya no existe!
                                    Call WriteLocaleMsg(UserIndex, "723", e_FontTypeNames.FONTTYPE_INFOIAO)
                                Else
124                                 Log = "Repuesta Afirmativa 1-2 "
126                                 If UserList(UserList(userIndex).Grupo.PropuestaDe.ArrayIndex).Grupo.CantidadMiembros = 1 Then
128                                     Call GroupCreateSuccess(UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex)
132                                     Log = "Repuesta Afirmativa 1-3 "
                                    End If
134                                 Call AddUserToGRoup(UserIndex, UserList(UserIndex).Grupo.PropuestaDe.ArrayIndex)
                                End If
                            End If
                        Else
166                         ' Msg724=Servidor » Solicitud de grupo invalida, reintente...
                            Call WriteLocaleMsg(UserIndex, "724", e_FontTypeNames.FONTTYPE_SERVER)
                        End If

                        'unirlo
168                 Case 2
170                     Log = "Repuesta Afirmativa 2"
172                     ' Msg725=¡Ahora sos un ciudadano!
                        Call WriteLocaleMsg(UserIndex, "725", e_FontTypeNames.FONTTYPE_INFOIAO)
174                     Call VolverCiudadano(UserIndex)
                    
176                 Case 3
178                     Log = "Repuesta Afirmativa 3"
                    
180                     UserList(UserIndex).Hogar = UserList(UserIndex).PosibleHogar

182                     Select Case UserList(UserIndex).Hogar

                            Case e_Ciudad.cUllathorpe
184                             DeDonde = "Ullathorpe"
                            
186                         Case e_Ciudad.cNix
188                             DeDonde = "Nix"
                
190                         Case e_Ciudad.cBanderbill
192                             DeDonde = "Banderbill"
                        
194                         Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
196                             DeDonde = "Lindos"
                            
198                         Case e_Ciudad.cArghal
200                             DeDonde = " Arghal"

                            Case e_Ciudad.cForgat
                                DeDonde = " Forgat"
                            
202                         Case e_Ciudad.cArkhein
204                             DeDonde = " Arkhein"
                            
206                         Case Else
208                             DeDonde = "Ullathorpe"

                        End Select
                    
210                     If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
212                         Call WriteChatOverHead(UserIndex, "¡Gracias " & UserList(UserIndex).name & "! Ahora perteneces a la ciudad de " & DeDonde & ".", NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                        Else
'Msg1244= ¡Gracias ¬1
Call WriteLocaleMsg(UserIndex, "1244", e_FontTypeNames.FONTTYPE_INFO, UserList(UserIndex).name)

                        End If
216                 Case 4
218                     Log = "Repuesta Afirmativa 4"
                
220                     If IsValidUserRef(UserList(userIndex).flags.targetUser) Then
                
222                         UserList(userIndex).ComUsu.DestUsu = UserList(userIndex).flags.targetUser
224                         UserList(userIndex).ComUsu.DestNick = UserList(UserList(userIndex).flags.targetUser.ArrayIndex).name
226                         UserList(UserIndex).ComUsu.cant = 0
228                         UserList(UserIndex).ComUsu.Objeto = 0
230                         UserList(UserIndex).ComUsu.Acepto = False
                    
                            'Rutina para comerciar con otro usuario
232                         Call IniciarComercioConUsuario(userIndex, UserList(userIndex).flags.targetUser.ArrayIndex)
                        Else
234                         ' Msg726=Servidor » Solicitud de comercio invalida, reintente...
                            Call WriteLocaleMsg(UserIndex, "726", e_FontTypeNames.FONTTYPE_SERVER)
                        End If
                
                    Case 5
                        Dim i As Integer, j As Integer
                        
                        With UserList(UserIndex)
                            For i = 1 To MAX_INVENTORY_SLOTS
                                For j = 1 To UBound(PecesEspeciales)
                                    If .Invent.Object(i).ObjIndex = PecesEspeciales(j).ObjIndex Then
                                        .Stats.PuntosPesca = .Stats.PuntosPesca + (ObjData(.Invent.Object(i).ObjIndex).PuntosPesca * .Invent.Object(i).amount)
                                        .Stats.GLD = .Stats.GLD + (ObjData(.Invent.Object(i).ObjIndex).Valor * .Invent.Object(i).amount * 1.2)
                                        Call WriteUpdateGold(userindex)
                                        Call QuitarUserInvItem(UserIndex, i, .Invent.Object(i).amount)
                                        Call UpdateUserInv(False, UserIndex, i)
                                    End If
                                Next j
                            Next i
                            Dim charindexstr As Integer
                            charIndexStr = str(NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex)
                            If charindexstr > 0 Then
                                Call WriteChatOverHead(UserIndex, "Felicitaciones! Ahora tienes un total de " & .Stats.PuntosPesca & " puntos de pesca.", charindexstr, &HFFFF00)
                            End If
                            .flags.pregunta = 0
                        End With
                                                
236
262                 Case Else
264                     ' Msg727=No tienes preguntas pendientes.
                        Call WriteLocaleMsg(UserIndex, "727", e_FontTypeNames.FONTTYPE_INFOIAO)

                        
                End Select
        
            Else
266             Log = "Repuesta negativa"
        
268             Select Case UserList(UserIndex).flags.pregunta

                    Case 1
270                     Log = "Repuesta negativa 1"
272                     If IsValidUserRef(UserList(userIndex).Grupo.PropuestaDe) Then
'Msg1245= El usuario no esta interesado en formar parte del grupo.
Call WriteLocaleMsg(UserList(userIndex).Grupo.PropuestaDe.ArrayIndex, "1245", e_FontTypeNames.FONTTYPE_INFO)
                        End If

276                     Call SetUserRef(UserList(userIndex).Grupo.PropuestaDe, 0)
'Msg1246= Has rechazado la propuesta.
Call WriteLocaleMsg(UserIndex, "1246", e_FontTypeNames.FONTTYPE_INFO)
                
280                 Case 2
282                     Log = "Repuesta negativa 2"
'Msg1247= ¡Continuas siendo neutral!
Call WriteLocaleMsg(UserIndex, "1247", e_FontTypeNames.FONTTYPE_INFO)
286                     Call VolverCriminal(UserIndex)

288                 Case 3
290                     Log = "Repuesta negativa 3"
                    
292                     Select Case UserList(UserIndex).PosibleHogar

                            Case e_Ciudad.cUllathorpe
294                             DeDonde = "Ullathorpe"
                            
296                         Case e_Ciudad.cNix
298                             DeDonde = "Nix"
                
300                         Case e_Ciudad.cBanderbill
302                             DeDonde = "Banderbill"
                        
304                         Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
306                             DeDonde = "Lindos"
                            
308                         Case e_Ciudad.cArghal
310                             DeDonde = " Arghal"

                            Case e_Ciudad.cForgat
                                DeDonde = " Forgat"
                            
312                         Case e_Ciudad.cArkhein
314                             DeDonde = " Arkhein"
                            
316                         Case Else
318                             DeDonde = "Ullathorpe"

                        End Select
                    
320                     If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
322                         Call WriteChatOverHead(UserIndex, "¡No hay problema " & UserList(UserIndex).name & "! Sos bienvenido en " & DeDonde & " cuando gustes.", NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                        End If
324                     UserList(UserIndex).PosibleHogar = UserList(UserIndex).Hogar
326                 Case 4
328                     Log = "Repuesta negativa 4"
                    
330                     If IsValidUserRef(UserList(userIndex).flags.targetUser) Then
'Msg1248= El usuario no desea comerciar en este momento.
Call WriteLocaleMsg(UserList(userIndex).flags.targetUser.ArrayIndex, "1248", e_FontTypeNames.FONTTYPE_INFO)
                        End If

334                 Case 5
336                     Log = "Repuesta negativa 5"
338                 Case Else
340                     ' Msg727=No tienes preguntas pendientes.
                        Call WriteLocaleMsg(UserIndex, "727", e_FontTypeNames.FONTTYPE_INFOIAO)
                End Select
            End If
        End With
        Exit Sub
    
ErrHandler:
342     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResponderPregunta", Erl)
344

End Sub

Private Sub HandleRequestGrupo(ByVal UserIndex As Integer)

        On Error GoTo hErr

        'Author: Pablo Mercavides

100     Call WriteDatosGrupo(UserIndex)
    
        Exit Sub
    
hErr:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestGrupo", Erl)
104

End Sub

Private Sub HandleAbandonarGrupo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleAbandonarGrupo_Err

100     With UserList(UserIndex)
        
102         Call Reader.ReadInt16
        
104         If UserList(userIndex).Grupo.Lider.ArrayIndex = userIndex Then
106             Call FinalizarGrupo(UserIndex)
            Else
126             Call SalirDeGrupo(UserIndex)
            End If

        End With
        Exit Sub

HandleAbandonarGrupo_Err:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAbandonarGrupo", Erl)
130
    
End Sub

Private Sub HandleHecharDeGrupo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleHecharDeGrupo_Err

100     With UserList(UserIndex)

            Dim Indice As Byte

102         Indice = Reader.ReadInt8()
        
104         Call EcharMiembro(UserIndex, Indice)

        End With
        
        Exit Sub

HandleHecharDeGrupo_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHecharDeGrupo", Erl)
108
    
End Sub

Private Sub HandleMacroPos(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleMacroPos_Err

100     With UserList(UserIndex)

102         .ChatCombate = Reader.ReadInt8()
104         .ChatGlobal = Reader.ReadInt8()

        End With
        
        Exit Sub

HandleMacroPos_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMacroPos", Erl)
108
    
End Sub

Private Sub HandleSubastaInfo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleSubastaInfo_Err

100     With UserList(UserIndex)

102         If Subasta.HaySubastaActiva Then

'Msg1249= Subastador: ¬1
Call WriteLocaleMsg(UserIndex, "1249", e_FontTypeNames.FONTTYPE_INFO, Subasta.Subastador)
'Msg1250= Objeto: ¬1
Call WriteLocaleMsg(UserIndex, "1250", e_FontTypeNames.FONTTYPE_INFO, ObjData(Subasta.ObjSubastado).Name)

108             If Subasta.HuboOferta Then
'Msg1251= Mejor oferta: ¬1
Call WriteLocaleMsg(UserIndex, "1251", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.MejorOferta))
'Msg1252= Podes realizar una oferta escribiendo /OFERTAR ¬1
Call WriteLocaleMsg(UserIndex, "1252", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.MejorOferta + 100))
                Else
'Msg1253= Oferta inicial: ¬1
Call WriteLocaleMsg(UserIndex, "1253", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.OfertaInicial))
'Msg1254= Podes realizar una oferta escribiendo /OFERTAR ¬1
Call WriteLocaleMsg(UserIndex, "1254", e_FontTypeNames.FONTTYPE_INFO, PonerPuntos(Subasta.OfertaInicial + 100))

                End If

'Msg1255= Tiempo Restante de subasta:  ¬1
Call WriteLocaleMsg(UserIndex, "1255", e_FontTypeNames.FONTTYPE_INFO, SumarTiempo(Subasta.TiempoRestanteSubasta))
            
            Else
120             ' Msg728=No hay ninguna subasta activa en este momento.
                Call WriteLocaleMsg(UserIndex, "728", e_FontTypeNames.FONTTYPE_SUBASTA)

            End If

        End With
        
        Exit Sub

HandleSubastaInfo_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSubastaInfo", Erl)
124
End Sub

Private Sub HandleCancelarExit(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleCancelarExit_Err

100     Call CancelExit(UserIndex)
        
        Exit Sub

HandleCancelarExit_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCancelarExit", Erl)
104
        
End Sub

Private Sub HandleEventoInfo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        
        On Error GoTo HandleEventoInfo_Err

100     With UserList(UserIndex)

102         If EventoActivo Then
104             Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", e_FontTypeNames.FONTTYPE_New_Eventos)
            Else
106             ' Msg729=Eventos> Actualmente no hay ningún evento en curso.
                Call WriteLocaleMsg(UserIndex, "729", e_FontTypeNames.FONTTYPE_New_Eventos)

            End If
        
            Dim i           As Byte
            Dim encontre    As Boolean
            Dim HoraProximo As Byte
   
108         If Not HoraEvento + 1 >= 24 Then
   
110             For i = HoraEvento + 1 To 23

112                 If Evento(i).Tipo <> 0 Then
114                     encontre = True
116                     HoraProximo = i
                        Exit For

                    End If

118             Next i

            End If
        
120         If encontre = False Then

122             For i = 0 To HoraEvento

124                 If Evento(i).Tipo <> 0 Then
126                     encontre = True
128                     HoraProximo = i
                        Exit For

                    End If

130             Next i

            End If
        
132         If encontre Then
'Msg1256= Eventos> El proximo evento ¬1
Call WriteLocaleMsg(UserIndex, "1256", e_FontTypeNames.FONTTYPE_INFO, DescribirEvento(HoraProximo))
            Else
136             ' Msg730=Eventos> No hay eventos próximos.
                Call WriteLocaleMsg(UserIndex, "730", e_FontTypeNames.FONTTYPE_New_Eventos)

            End If

        End With
        
        Exit Sub

HandleEventoInfo_Err:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEventoInfo", Erl)
140
End Sub

Private Sub HandleCrearEvento(ByVal UserIndex As Integer)

        '***************************************************
        'Author: Pablo Mercavides
        '***************************************************

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Tipo           As Byte
            Dim Duracion       As Byte
            Dim multiplicacion As Byte
        
102         Tipo = Reader.ReadInt8()
104         Duracion = Reader.ReadInt8()
106         multiplicacion = Reader.ReadInt8()

108         If multiplicacion > 5 Then 'no superar este multiplicador
110             multiplicacion = 2
            End If
        
            '/ dejar solo Administradores
112         If .flags.Privilegios >= e_PlayerType.Admin Then
114             If EventoActivo = False Then
116                 If LenB(Tipo) = 0 Or LenB(Duracion) = 0 Or LenB(multiplicacion) = 0 Then
118                     ' Msg731=Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.
                        Call WriteLocaleMsg(UserIndex, "731", e_FontTypeNames.FONTTYPE_New_Eventos)
                    Else
                
120                     Call ForzarEvento(Tipo, Duracion, multiplicacion, UserList(UserIndex).Name)
                  
                    End If

                Else
122                 ' Msg732=Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.
                    Call WriteLocaleMsg(UserIndex, "732", e_FontTypeNames.FONTTYPE_New_Eventos)

                End If
            Else
124             ' Msg733=Servidor » Solo Administradores pueder crear estos eventos.
                Call WriteLocaleMsg(UserIndex, "733", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
    
        Exit Sub

ErrHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
128

End Sub

Private Sub HandleCompletarViaje(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides

        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim Destino As Byte

            Dim costo   As Long

102         Destino = Reader.ReadInt8()
104         costo = Reader.ReadInt32()

            ' WyroX: WTF el costo lo decide el cliente... Desactivo....
            Exit Sub

106         If costo <= 0 Then Exit Sub

            Dim DeDonde As t_CityWorldPos

108         If UserList(UserIndex).Stats.GLD < costo Then
'Msg1257= No tienes suficiente dinero.
Call WriteLocaleMsg(UserIndex, "1257", e_FontTypeNames.FONTTYPE_INFO)
            
            Else

112             Select Case Destino

                    Case e_Ciudad.cUllathorpe
114                     DeDonde = CityUllathorpe
                        
116                 Case e_Ciudad.cNix
118                     DeDonde = CityNix
            
120                 Case e_Ciudad.cBanderbill
122                     DeDonde = CityBanderbill
                    
124                 Case e_Ciudad.cLindos 'Vamos a tener que ir por todo el desierto... uff!
126                     DeDonde = CityLindos
                        
128                 Case e_Ciudad.cArghal
130                     DeDonde = CityArghal

                    Case e_Ciudad.cForgat
                        DeDonde = CityForgat
                        
132                 Case e_Ciudad.cArkhein
134                     DeDonde = CityArkhein
                        
136                 Case Else
138                     DeDonde = CityUllathorpe

                End Select
        
140             If DeDonde.NecesitaNave > 0 Then
142                 If UserList(UserIndex).Stats.UserSkills(e_Skill.Navegacion) < 80 Then
'Msg1258= Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.
Call WriteLocaleMsg(UserIndex, "1258", e_FontTypeNames.FONTTYPE_INFO)
'Msg1259= Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.
Call WriteLocaleMsg(UserIndex, "1259", e_FontTypeNames.FONTTYPE_INFO)
                    Else

146                     If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
148                         If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose <> 0 Then
150                             Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose, NO_3D_SOUND, NO_3D_SOUND)
                            End If
                        End If

152                     Call WarpToLegalPos(UserIndex, DeDonde.MapaViaje, DeDonde.ViajeX, DeDonde.ViajeY, True)
'Msg1260= Has viajado por varios días, te sientes exhausto!
Call WriteLocaleMsg(UserIndex, "1260", e_FontTypeNames.FONTTYPE_INFO)
156                     UserList(UserIndex).Stats.MinAGU = 0
158                     UserList(UserIndex).Stats.MinHam = 0
                    
164                     UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
166                     Call WriteUpdateHungerAndThirst(UserIndex)
168                     Call WriteUpdateUserStats(UserIndex)

                    End If

                Else
            
                    Dim Map As Integer

                    Dim X   As Byte

                    Dim Y   As Byte
            
170                 Map = DeDonde.MapaViaje
172                 X = DeDonde.ViajeX
174                 Y = DeDonde.ViajeY

176                 If IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then
178                     If NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose <> 0 Then
180                         Call WritePlayWave(UserIndex, NpcList(UserList(UserIndex).flags.TargetNPC.ArrayIndex).SoundClose, NO_3D_SOUND, NO_3D_SOUND)

                        End If

                    End If
                
182                 Call WarpUserChar(UserIndex, Map, X, Y, True)
'Msg1261= Has viajado por varios días, te sientes exhausto!
Call WriteLocaleMsg(UserIndex, "1261", e_FontTypeNames.FONTTYPE_INFO)
186                 UserList(UserIndex).Stats.MinAGU = 0
188                 UserList(UserIndex).Stats.MinHam = 0
                
194                 UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - costo
196                 Call WriteUpdateHungerAndThirst(UserIndex)
198                 Call WriteUpdateUserStats(UserIndex)
        
                End If

            End If

        End With
    
        Exit Sub

ErrHandler:
200     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCompletarViaje", Erl)
202

End Sub

Public Sub HandleQuest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuest_Err

100     If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then Exit Sub
        Dim NpcIndex As Integer
        Dim tmpByte  As Byte
102     NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
        'Esta el personaje en la distancia correcta?
104     If Distancia(UserList(UserIndex).Pos, NpcList(NpcIndex).Pos) > 5 Then
106         ' Msg8=Estas demasiado lejos.
            Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'El NPC hace quests?
108     If NpcList(NpcIndex).NumQuest = 0 Then
110         Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageChatOverHead("No tengo ninguna misión para ti.", NpcList(NpcIndex).Char.charindex, vbWhite))
            Exit Sub

        End If
    
112     Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", NpcList(NpcIndex).Char.CharIndex, vbWhite))

        Exit Sub

HandleQuest_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuest", Erl)
116
        
End Sub

Public Sub HandleQuestAccept(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuestAccept_Err

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el evento de aceptar una quest.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim NpcIndex  As Integer
        Dim QuestSlot As Byte
        Dim Indice    As Byte
        
100     Indice = Reader.ReadInt8
102     If Not IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) And UserList(UserIndex).flags.QuestOpenByObj = False Then Exit Sub

104     NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
106     If NpcIndex > 0 Then
108         If QuestList(NpcList(NpcIndex).QuestNumber(Indice)).Trabajador And UserList(UserIndex).clase <> e_Class.Trabajador Then
'Msg1262= La quest es solo para trabajadores.
Call WriteLocaleMsg(UserIndex, "1262", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            
112         If Distancia(UserList(UserIndex).pos, NpcList(NpcIndex).pos) > 5 Then
114             ' Msg8=Estas demasiado lejos.
                Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            
116         If TieneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(Indice)) Then
'Msg1263= La quest ya esta en curso.
Call WriteLocaleMsg(UserIndex, "1263", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            'El personaje completo la quest que requiere?
120         If QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest > 0 Then
122             If Not UserDoneQuest(UserIndex, QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest) Then
124                 Call WriteChatOverHead(UserIndex, "Debes completar la quest " & QuestList(QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredQuest).nombre & " para emprender esta misión.", NpcList(NpcIndex).Char.charindex, vbYellow)
                    Exit Sub
    
                End If
    
            End If
    
            'El personaje tiene suficiente nivel?
126         If UserList(UserIndex).Stats.ELV < QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredLevel Then
128             Call WriteChatOverHead(UserIndex, "Debes ser por lo menos nivel " & QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredLevel & " para emprender esta misión.", NpcList(NpcIndex).Char.charindex, vbYellow)
                Exit Sub
            End If
            
130         If QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredSkill.SkillType > 0 Then
132             If UserList(UserIndex).Stats.UserSkills(QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredSkill.SkillType) < QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredSkill.RequiredValue Then
134                 Call WriteLocaleChatOverHead(UserIndex, MsgRequiredSkill, SkillsNames(QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredSkill.SkillType), NpcList(NpcIndex).Char.charindex, vbYellow)
                    Exit Sub
                End If
            End If
            
            'El personaje no es la clase requerida?
136         If UserList(UserIndex).clase <> QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredClass And _
                QuestList(NpcList(npcIndex).QuestNumber(Indice)).RequiredClass > 0 Then
138              Call WriteChatOverHead(UserIndex, "Debes ser " & ListaClases(QuestList(NpcList(NpcIndex).QuestNumber(Indice)).RequiredClass) & " para emprender esta misión.", NpcList(NpcIndex).Char.charindex, vbYellow)
                Exit Sub

            End If
            'La quest no es repetible?
140         If QuestList(NpcList(NpcIndex).QuestNumber(Indice)).Repetible = 0 Then
                'El personaje ya hizo la quest?
142             If UserDoneQuest(UserIndex, NpcList(NpcIndex).QuestNumber(Indice)) Then
144                 Call WriteChatOverHead(UserIndex, "QUESTNEXT*" & NpcList(NpcIndex).QuestNumber(Indice), NpcList(NpcIndex).Char.charindex, vbYellow)
                    Exit Sub
        
                End If
            End If
        
146         QuestSlot = FreeQuestSlot(UserIndex)
    
148         If QuestSlot = 0 Then
150             Call WriteChatOverHead(UserIndex, "Debes completar las misiones en curso para poder aceptar más misiones.", NpcList(NpcIndex).Char.charindex, vbYellow)
                Exit Sub
    
            End If
        
            'Agregamos la quest.
152         With UserList(UserIndex).QuestStats.Quests(QuestSlot)
                
154             .QuestIndex = NpcList(NpcIndex).QuestNumber(Indice)
                '.QuestIndex = UserList(UserIndex).flags.QuestNumber
            
156             If QuestList(.QuestIndex).RequiredNPCs Then ReDim .NPCsKilled(1 To QuestList(.QuestIndex).RequiredNPCs)
158             If QuestList(.QuestIndex).RequiredTargetNPCs Then ReDim .NPCsTarget(1 To QuestList(.QuestIndex).RequiredTargetNPCs)
160             UserList(UserIndex).flags.ModificoQuests = True
                
'Msg1264= Has aceptado la misión ¬1
Call WriteLocaleMsg(UserIndex, "1264", e_FontTypeNames.FONTTYPE_INFO, Chr(34))

164             If (FinishQuestCheck(UserIndex, .QuestIndex, QuestSlot)) Then
166                 Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 3)
                Else
168                 Call WriteUpdateNPCSimbolo(UserIndex, NpcIndex, 4)
                End If
                
            End With
        Else
            
        End If
        Exit Sub

HandleQuestAccept_Err:
170     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestAccept", Erl)

        
End Sub

Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
        
        On Error GoTo HandleQuestDetailsRequest_Err

        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestInfoRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        Dim QuestSlot As Byte

100     QuestSlot = Reader.ReadInt8
        If QuestSlot <= MAXUSERQUESTS And QuestSlot > 0 Then
            If UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex > 0 Then
102             Call WriteQuestDetails(UserIndex, UserList(UserIndex).QuestStats.Quests(QuestSlot).QuestIndex, QuestSlot)
            End If
        End If
        Exit Sub

HandleQuestDetailsRequest_Err:
104     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestDetailsRequest", Erl)
106
        
End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestAbandon.
        'Last modified: 31/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

        On Error GoTo HandleQuestAbandon_Err
        
        With UserList(UserIndex)
        
            Dim Slot As Byte
            Slot = Reader.ReadInt8
            
            If Slot > MAXUSERQUESTS Then Exit Sub
            
            With .QuestStats.Quests(Slot)
                ' Le quitamos los objetos de quest que no puede tirar
                If QuestList(.QuestIndex).RequiredOBJs Then
                
                    Dim ObjIndex As Integer, i As Integer
                    
                    For i = 1 To QuestList(.QuestIndex).RequiredOBJs
                        ObjIndex = QuestList(.QuestIndex).RequiredOBJ(i).ObjIndex
                        
                        If ObjData(ObjIndex).Intirable = 1 And ObjData(ObjIndex).Instransferible Then
                        
                            ' Revisamos que ninguna otra quest que tenga activa le pida el mismo item
                            Dim q As Integer, j As Byte, K As Byte, QuitarItem As Boolean

                            QuitarItem = True
                            
                            For j = 1 To MAXUSERQUESTS
                                q = UserList(UserIndex).QuestStats.Quests(j).QuestIndex
                                If q <> 0 And q <> .QuestIndex Then
                                    For K = 1 To QuestList(q).RequiredOBJs
                                        If QuestList(q).RequiredOBJ(K).ObjIndex = ObjIndex Then
                                            QuitarItem = False
                                            Exit For
                                        End If
                                    Next
                                End If
                                If Not QuitarItem Then Exit For
                            Next
                            
                            If QuitarItem Then
                                Call QuitarObjetos(ObjIndex, MAX_INVENTORY_OBJS, UserIndex)
                            End If
                        End If
                    Next i
                
                End If
            End With
    
            'Borramos la quest.
100         Call CleanQuestSlot(UserIndex, Slot)
        
            'Ordenamos la lista de quests del usuario.
102         Call ArrangeUserQuests(UserIndex)
        
            'Enviamos la lista de quests actualizada.
104         Call WriteQuestListSend(UserIndex)

        End With
        
        Exit Sub

HandleQuestAbandon_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestAbandon", Erl)
108
        
End Sub

Public Sub HandleQuestListRequest(ByVal UserIndex As Integer)
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        'Maneja el paquete QuestListRequest.
        'Last modified: 30/01/2010 by Amraphen
        '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
        
        On Error GoTo HandleQuestListRequest_Err

100     Call WriteQuestListSend(UserIndex)
        
        Exit Sub

HandleQuestListRequest_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleQuestListRequest", Erl)
104
        
End Sub

''
' Handles the "Consulta" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsulta(ByVal UserIndex As Integer)
        '***************************************************
        'Author: ZaMa
        'Last Modification: 01/05/2010
        'Habilita/Deshabilita el modo consulta.
        '01/05/2010: ZaMa - Agrego validaciones.
        '16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
        '***************************************************

        Dim UserConsulta As t_UserReference
    
100     With UserList(UserIndex)
 
            Dim Nick As String
102         Nick = Reader.ReadString8

            ' Comando exclusivo para gms
104         If Not EsGM(UserIndex) Then Exit Sub
        
106         If Len(Nick) <> 0 Then
108             UserConsulta = NameIndex(Nick)
                'Se asegura que el target exista
110             If Not IsValidUserRef(UserConsulta) Then
'Msg1265= El usuario se encuentra offline.
Call WriteLocaleMsg(UserIndex, "1265", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
114             Call SetUserRef(UserConsulta, .flags.targetUser.ArrayIndex)
                'Se asegura que el target exista
116             If IsValidUserRef(UserConsulta) Then
'Msg1266= Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.
Call WriteLocaleMsg(UserIndex, "1266", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            ' No podes ponerte a vos mismo en modo consulta.
120         If UserConsulta.ArrayIndex = userIndex Then Exit Sub
            ' No podes estra en consulta con otro gm
122         If EsGM(UserConsulta.ArrayIndex) Then
'Msg1267= No puedes iniciar el modo consulta con otro administrador.
Call WriteLocaleMsg(UserIndex, "1267", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

            ' Si ya estaba en consulta, termina la consulta
126         If UserList(UserConsulta.ArrayIndex).flags.EnConsulta Then
'Msg1268= Has terminado el modo consulta con ¬1
Call WriteLocaleMsg(userIndex, "1268", e_FontTypeNames.FONTTYPE_INFO, UserList(UserConsulta.ArrayIndex).name)
'Msg1269= Has terminado el modo consulta.
Call WriteLocaleMsg(UserConsulta.ArrayIndex, "1269", e_FontTypeNames.FONTTYPE_INFO)
            
132             Call LogGM(.name, "Termino consulta con " & UserList(UserConsulta.ArrayIndex).name)
            
134             UserList(UserConsulta.ArrayIndex).flags.EnConsulta = False
        
                ' Sino la inicia
            Else
        
'Msg1270= Has iniciado el modo consulta con ¬1
Call WriteLocaleMsg(userIndex, "1270", e_FontTypeNames.FONTTYPE_INFO, UserList(UserConsulta.ArrayIndex).name)
'Msg1271= Has iniciado el modo consulta.
Call WriteLocaleMsg(UserConsulta.ArrayIndex, "1271", e_FontTypeNames.FONTTYPE_INFO)
            
140             Call LogGM(.name, "Inicio consulta con " & UserList(UserConsulta.ArrayIndex).name)
            
142             With UserList(UserConsulta.ArrayIndex)

144                 If Not EstaPCarea(userIndex, UserConsulta.ArrayIndex) Then
                        Dim X As Byte
                        Dim Y As Byte
                        
146                     X = .Pos.X
148                     Y = .Pos.Y
150                     Call FindLegalPos(UserIndex, .Pos.Map, X, Y)
152                     Call WarpUserChar(UserIndex, .Pos.Map, X, Y, True)
                        
                    End If
            
154                 If UserList(UserIndex).flags.AdminInvisible = 1 Then
156                     Call DoAdminInvisible(UserIndex)

                    End If

158                 .flags.EnConsulta = True
                
                    ' Pierde invi u ocu
160                 If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                
162                     .flags.Oculto = 0
164                     .flags.invisible = 0
166                     .Counters.TiempoOculto = 0
168                     .Counters.Invisibilidad = 0
                        .Counters.DisabledInvisibility = 0
                    
170                     If UserList(UserConsulta.ArrayIndex).flags.Navegando = 0 Then
                            
172                         Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(userindex).Pos.X, UserList(userindex).Pos.y))

                        End If

                    End If

                End With

            End If
        
174         Call SetModoConsulta(UserConsulta.ArrayIndex)

        End With
    
        Exit Sub
    
ErrHandler:
176     Call TraceError(Err.Number, Err.Description, "Protocol.HandleConsulta", Erl)
178

End Sub

Private Sub HandleGetMapInfo(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         If EsGM(UserIndex) Then
            
                Dim Response As String
            
104             Response = "[Info de mapa " & .Pos.Map & "]" & vbNewLine
106             Response = Response & "Nombre = " & MapInfo(.Pos.Map).map_name & vbNewLine
108             Response = Response & "Seguro = " & MapInfo(.Pos.Map).Seguro & vbNewLine
110             Response = Response & "Newbie = " & MapInfo(.Pos.Map).Newbie & vbNewLine
112             Response = Response & "Nivel = " & MapInfo(.Pos.Map).MinLevel & "/" & MapInfo(.Pos.Map).MaxLevel & vbNewLine
114             Response = Response & "SinInviOcul = " & MapInfo(.Pos.Map).SinInviOcul & vbNewLine
116             Response = Response & "SinMagia = " & MapInfo(.Pos.Map).SinMagia & vbNewLine
118             Response = Response & "SoloClanes = " & MapInfo(.Pos.Map).SoloClanes & vbNewLine
120             Response = Response & "NoPKs = " & MapInfo(.Pos.Map).NoPKs & vbNewLine
122             Response = Response & "NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos & vbNewLine
124             Response = Response & "Salida = " & MapInfo(.Pos.Map).Salida.Map & "-" & MapInfo(.Pos.Map).Salida.X & "-" & MapInfo(.Pos.Map).Salida.Y & vbNewLine
126             Response = Response & "Terreno = " & MapInfo(.Pos.Map).terrain & vbNewLine
128             Response = Response & "NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos & vbNewLine
130             Response = Response & "Zona = " & MapInfo(.Pos.Map).zone & vbNewLine
            
132             Call WriteConsoleMsg(UserIndex, Response, e_FontTypeNames.FONTTYPE_INFO)
        
            End If
    
        End With

End Sub


Private Sub HandleSeguroResu(ByVal UserIndex As Integer)

100     With UserList(UserIndex)

102         .flags.SeguroResu = Not .flags.SeguroResu
        
104         Call WriteSeguroResu(UserIndex, .flags.SeguroResu)
    
        End With

End Sub

Private Sub HandleCuentaExtractItem(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCuentaExtractItem_Err
100     With UserList(UserIndex)
            Dim Slot        As Byte
            Dim slotdestino As Byte
            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
106         slotdestino = Reader.ReadInt8()
        
108         If .flags.Muerto = 1 Then
110             'Msg77=¡¡Estás muerto!!.
                Exit Sub
            End If
        
112         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
114         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then
                Exit Sub
            End If
        End With
        Exit Sub

HandleCuentaExtractItem_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaExtractItem", Erl)
118
        
End Sub

Private Sub HandleCuentaDeposit(ByVal UserIndex As Integer)
        On Error GoTo HandleCuentaDeposit_Err
100     With UserList(UserIndex)

            Dim Slot        As Byte

            Dim slotdestino As Byte

            Dim amount      As Integer
        
102         Slot = Reader.ReadInt8()
104         amount = Reader.ReadInt16()
106         slotdestino = Reader.ReadInt8()
        
            'Dead people can't commerce...
108         If .flags.Muerto = 1 Then
110             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            'íEl target es un NPC valido?
112         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
        
            'íEl NPC puede comerciar?
114         If NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Banquero Then
                Exit Sub
            End If
            
116         If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
118             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End With
        Exit Sub
HandleCuentaDeposit_Err:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaDeposit", Erl)
End Sub

Private Sub HandleCommerceSendChatMessage(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)

            Dim chatMessage As String
        
102         chatMessage = "[" & UserList(UserIndex).Name & "] " & Reader.ReadString8
        
            'El mensaje se lo envío al destino
            If Not IsValidUserRef(UserList(userIndex).ComUsu.DestUsu) Then Exit Sub
104         Call WriteCommerceRecieveChatMessage(UserList(userIndex).ComUsu.DestUsu.ArrayIndex, chatMessage)
        
            'y tambien a mi mismo
106         Call WriteCommerceRecieveChatMessage(UserIndex, chatMessage)

        End With
    
        Exit Sub
    
ErrHandler:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCommerceSendChatMessage", Erl)
110
    
End Sub

Private Sub HandleLogMacroClickHechizo(ByVal UserIndex As Integer)

100     With UserList(UserIndex)
            Dim tipoMacro As Byte
            Dim mensaje As String
            Dim clicks As Long
            tipoMacro = Reader.ReadInt8
            clicks = Reader.ReadInt32
            
            Select Case tipoMacro
            
                Case tMacro.Coordenadas
102                 mensaje = "Control AntiCheat--> El usuario " & UserList(UserIndex).name & "| está utilizando macro de COORDENADAS."
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(mensaje, e_FontTypeNames.FONTTYPE_INFO))
                Case tMacro.dobleclick
                    mensaje = "Control AntiCheat--> El usuario " & UserList(UserIndex).name & "| está utilizando macro de DOBLE CLICK (CANTIDAD DE CLICKS: " & clicks & " )."
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(mensaje, e_FontTypeNames.FONTTYPE_INFO))
                Case tMacro.inasistidoPosFija
                    mensaje = "Control AntiCheat--> El usuario " & UserList(UserIndex).name & "| está utilizando macro de INASISTIDO."
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(mensaje, e_FontTypeNames.FONTTYPE_INFO))
                Case tMacro.borrarCartel
                    mensaje = "Control AntiCheat--> El usuario " & UserList(UserIndex).name & "| está utilizando macro de CARTELEO."
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(mensaje, e_FontTypeNames.FONTTYPE_INFO))
            End Select
            
            

        End With

End Sub



Private Sub HandleHome(ByVal UserIndex As Integer)
        
        On Error GoTo HandleHome_Err
    
        

        '***************************************************
        'Author: Budi
        'Creation Date: 06/01/2010
        'Last Modification: 05/06/10
        'Pato - 05/06/10: Add the UCase$ to prevent problems.
        '***************************************************
    
100     With UserList(UserIndex)

104         If .flags.Muerto = 0 Then
'Msg1272= Debes estar muerto para utilizar este comando.
Call WriteLocaleMsg(UserIndex, "1272", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
                
            'Si el mapa tiene alguna restriccion (newbie, dungeon, etc...), no lo dejamos viajar.
108         If MapInfo(.Pos.Map).zone = "NEWBIE" Or MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
'Msg1273= No pueder viajar a tu hogar desde este mapa.
Call WriteLocaleMsg(UserIndex, "1273", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            
            End If
        
            'Si es un mapa comun y no esta en cana
112         If .Counters.Pena <> 0 Then
'Msg1274= No puedes usar este comando en prisión.
Call WriteLocaleMsg(UserIndex, "1274", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
116         If .flags.EnReto Then
'Msg1275= No podés regresar desde un reto. Usa /ABANDONAR para admitir la derrota y volver a la ciudad.
Call WriteLocaleMsg(UserIndex, "1275", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

120         If .flags.Traveling = 0 Then
            
122             If .Pos.Map <> Ciudades(.Hogar).Map Then
124                 Call goHome(UserIndex)
                
                Else
'Msg1276= Ya te encuentras en tu hogar.
Call WriteLocaleMsg(UserIndex, "1276", e_FontTypeNames.FONTTYPE_INFO)

                End If

            Else

128             .flags.Traveling = 0
130             .Counters.goHome = 0
            
'Msg1277= Ya hay un viaje en curso.
Call WriteLocaleMsg(UserIndex, "1277", e_FontTypeNames.FONTTYPE_INFO)
            
            End If
        
        End With

        
        Exit Sub

HandleHome_Err:
134     Call TraceError(Err.Number, Err.Description, "Hogar.HandleHome", Erl)

        
End Sub

Private Sub HandleAddItemCrafting(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)
    
            Dim InvSlot As Byte, CraftSlot As Byte
102         InvSlot = Reader.ReadInt8
104         CraftSlot = Reader.ReadInt8
        
106         If .flags.Crafteando = 0 Then Exit Sub
        
108         If InvSlot < 1 Or InvSlot > .CurrentInventorySlots Then Exit Sub

110         If .Invent.Object(InvSlot).ObjIndex = 0 Then Exit Sub

112         If CraftSlot < 1 Then
114             For CraftSlot = 1 To MAX_SLOTS_CRAFTEO
116                 If .CraftInventory(CraftSlot) = 0 Then
                        Exit For
                    End If
                Next
            End If

118         If CraftSlot > MAX_SLOTS_CRAFTEO Then
                Exit Sub
            End If

120         If .CraftInventory(CraftSlot) <> 0 Then Exit Sub

122         .CraftInventory(CraftSlot) = .Invent.Object(InvSlot).ObjIndex
    
124         Call QuitarUserInvItem(UserIndex, InvSlot, 1)
126         Call UpdateUserInv(False, UserIndex, InvSlot)

128         Call WriteCraftingItem(UserIndex, CraftSlot, .CraftInventory(CraftSlot))

            Dim Result As clsCrafteo
130         Set Result = CheckCraftingResult(UserIndex)
        
132         If Not Result Is .CraftResult Then
134             Set .CraftResult = Result
136             If Not .CraftResult Is Nothing Then
138                 Call WriteCraftingResult(UserIndex, .CraftResult.Resultado, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad), .CraftResult.Precio)
                Else
140                 Call WriteCraftingResult(UserIndex, 0)
                End If
            End If

        End With
    
        Exit Sub

ErrHandler:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAddItemCrafting", Erl)
144
End Sub

Private Sub HandleRemoveItemCrafting(ByVal UserIndex As Integer)
    
        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)
    
            Dim InvSlot As Byte, CraftSlot As Byte
102         CraftSlot = Reader.ReadInt8
104         InvSlot = Reader.ReadInt8
        
106         If .flags.Crafteando = 0 Then Exit Sub

108         If CraftSlot < 1 Or CraftSlot > MAX_SLOTS_CRAFTEO Then Exit Sub

110         If .CraftInventory(CraftSlot) = 0 Then Exit Sub

112         If InvSlot < 1 Then
                Dim TmpObj As t_Obj
114             TmpObj.ObjIndex = .CraftInventory(CraftSlot)
116             TmpObj.amount = 1
             
118             If Not MeterItemEnInventario(UserIndex, TmpObj) Then Exit Sub

120         ElseIf InvSlot <= .CurrentInventorySlots Then
122             If .Invent.Object(InvSlot).ObjIndex = 0 Then
124                 .Invent.Object(InvSlot).ObjIndex = .CraftInventory(CraftSlot)
            
126             ElseIf .Invent.Object(InvSlot).ObjIndex <> .CraftInventory(CraftSlot) Then
                    Exit Sub
                End If

128             .Invent.Object(InvSlot).amount = .Invent.Object(InvSlot).amount + 1
130             Call UpdateUserInv(False, UserIndex, InvSlot)
            End If

132         .CraftInventory(CraftSlot) = 0
134         Call WriteCraftingItem(UserIndex, CraftSlot, 0)
        
            Dim Result As clsCrafteo
136         Set Result = CheckCraftingResult(UserIndex)
        
138         If Not Result Is .CraftResult Then
140             Set .CraftResult = Result
142             If Not .CraftResult Is Nothing Then
144                 Call WriteCraftingResult(UserIndex, .CraftResult.Resultado, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad), .CraftResult.Precio)
                Else
146                 Call WriteCraftingResult(UserIndex, 0)
                End If
            End If

        End With
    
        Exit Sub
    
ErrHandler:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveItemCrafting", Erl)
150
End Sub

Private Sub HandleAddCatalyst(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)
    
            Dim Slot As Byte
102         Slot = Reader.ReadInt8
        
104         If .flags.Crafteando = 0 Then Exit Sub
        
106         If Slot < 1 Or Slot > .CurrentInventorySlots Then Exit Sub

108         If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub
        
110         If ObjData(.Invent.Object(Slot).ObjIndex).CatalizadorTipo = 0 Then Exit Sub

112         If .CraftCatalyst.ObjIndex <> 0 Then Exit Sub

114         .CraftCatalyst.ObjIndex = .Invent.Object(Slot).ObjIndex
116         .CraftCatalyst.amount = .Invent.Object(Slot).amount

118         Call QuitarUserInvItem(UserIndex, Slot, MAX_INVENTORY_OBJS)
120         Call UpdateUserInv(False, UserIndex, Slot)

122         If .CraftResult Is Nothing Then
124             Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, 0)
            Else
126             Call WriteCraftingCatalyst(UserIndex, .CraftCatalyst.ObjIndex, .CraftCatalyst.amount, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad))
            End If

        End With
    
        Exit Sub
    
ErrHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAddCatalyst", Erl)
130
End Sub

Private Sub HandleRemoveCatalyst(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)
    
            Dim Slot As Byte
102         Slot = Reader.ReadInt8
        
104         If .flags.Crafteando = 0 Then Exit Sub

106         If .CraftCatalyst.ObjIndex = 0 Then Exit Sub

108         If Slot < 1 Then
110             If Not MeterItemEnInventario(UserIndex, .CraftCatalyst) Then Exit Sub

112         ElseIf Slot <= .CurrentInventorySlots Then
114             If .Invent.Object(Slot).ObjIndex = 0 Then
116                 .Invent.Object(Slot).ObjIndex = .CraftCatalyst.ObjIndex

118             ElseIf .Invent.Object(Slot).ObjIndex <> .CraftCatalyst.ObjIndex Then
                    Exit Sub
                End If

120             .Invent.Object(Slot).amount = .Invent.Object(Slot).amount + .CraftCatalyst.amount
122             Call UpdateUserInv(False, UserIndex, Slot)
            End If

124         .CraftCatalyst.ObjIndex = 0
126         .CraftCatalyst.amount = 0
        
128         If .CraftResult Is Nothing Then
130             Call WriteCraftingCatalyst(UserIndex, 0, 0, 0)
            Else
132             Call WriteCraftingCatalyst(UserIndex, 0, 0, CalculateCraftProb(UserIndex, .CraftResult.Probabilidad))
            End If

        End With
    
        Exit Sub
    
ErrHandler:
134     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemoveCatalyst", Erl)
136
End Sub

Sub HandleCraftItem(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler

100     If UserList(UserIndex).flags.Crafteando = 0 Then Exit Sub

102     Call DoCraftItem(UserIndex)
    
        Exit Sub

ErrHandler:
104     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCraftItem", Erl)
106
End Sub

Private Sub HandleCloseCrafting(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
100     If UserList(UserIndex).flags.Crafteando = 0 Then Exit Sub

102     Call ReturnCraftingItems(UserIndex)
    
104     UserList(UserIndex).flags.Crafteando = 0
    
        Exit Sub
    
ErrHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCloseCrafting", Erl)
108
End Sub

Private Sub HandleMoveCraftItem(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)
    
            Dim Drag As Byte, Drop As Byte
102         Drag = Reader.ReadInt8
104         Drop = Reader.ReadInt8
        
106         If .flags.Crafteando = 0 Then Exit Sub
        
108         If Drag < 1 Or Drag > MAX_SLOTS_CRAFTEO Then Exit Sub
110         If Drop < 1 Or Drop > MAX_SLOTS_CRAFTEO Then Exit Sub
112         If Drag = Drop Then Exit Sub

114         If .CraftInventory(Drag) = 0 Then Exit Sub
116         If .CraftInventory(Drag) = .CraftInventory(Drop) Then Exit Sub

            Dim aux As Integer
118         aux = .CraftInventory(Drop)
120         .CraftInventory(Drop) = .CraftInventory(Drag)
122         .CraftInventory(Drag) = aux

124         Call WriteCraftingItem(UserIndex, Drag, .CraftInventory(Drag))
126         Call WriteCraftingItem(UserIndex, Drop, .CraftInventory(Drop))

        End With
    
        Exit Sub
    
ErrHandler:
128     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMoveCraftItem", Erl)
130
End Sub

Private Sub HandlePetLeaveAll(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
    
            Dim AlmenosUna As Boolean, i As Integer
    
102         For i = 1 To MAXMASCOTAS
104             If IsValidNpcRef(.MascotasIndex(i)) Then
106                 If NpcList(.MascotasIndex(i).ArrayIndex).flags.NPCActive Then
108                     Call QuitarNPC(.MascotasIndex(i).ArrayIndex, e_DeleteSource.ePetLeave)
110                     AlmenosUna = True
                    End If
                End If
112         Next i
114         If AlmenosUna Then
                .flags.ModificoMascotas = True
'Msg1278= Liberaste a tus mascotas.
Call WriteLocaleMsg(UserIndex, "1278", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePetLeaveAll", Erl)
End Sub


Private Sub HandleResetChar(ByVal UserIndex As Integer)
        On Error GoTo HandleResetChar_Err:
        
100     Dim Nick As String: Nick = Reader.ReadString8()

        #If DEBUGGING = 1 Then

            If UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin Then
                Dim user As t_UserReference
                user = NameIndex(Nick)
                
                If Not IsValidUserRef(user) Then
'Msg1279= Usuario offline o inexistente.
Call WriteLocaleMsg(UserIndex, "1279", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                
                With UserList(user.ArrayIndex)
                    .Stats.ELV = 1
                    .Stats.Exp = 0
                    .Stats.SkillPts = 10
                    
                    Dim i As Integer
                    For i = 1 To NUMSKILLS
                        .Stats.UserSkills(i) = 0
                    Next

                
                    .Stats.MaxAGU = 100
                    .Stats.MinAGU = 100
                
                    .Stats.MaxHam = 100
                    .Stats.MinHam = 100
            
                    .Stats.MaxHit = 2
                    .Stats.MinHIT = 1
                    .Stats.MaxMAN = UserMod.GetMaxMana(UserIndex)
                    .Stats.MaxSta = UserMod.GetMaxStamina(UserIndex)
                    .Stats.MaxHp = .Stats.UserAtributos(e_Atributos.Constitucion)
                    
                    .Stats.MinHp = .Stats.MaxHp
                    .Stats.MinMAN = .Stats.MaxMAN
                    .Stats.MinSta = .Stats.MaxSta
                    
                    .flags.ModificoSkills = True
                    
                    Call WriteUpdateUserStats(user.ArrayIndex)
                    Call WriteLevelUp(user.ArrayIndex, .Stats.SkillPts)
                End With
                
'Msg1280= Personaje reseteado a nivel 1.
Call WriteLocaleMsg(UserIndex, "1280", e_FontTypeNames.FONTTYPE_INFO)
            End If
        
        #End If
        
        Exit Sub

HandleResetChar_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetChar", Erl)
End Sub
Private Sub HandleResetearPersonaje(ByVal UserIndex As Integer)
    On Error GoTo HandleResetearPersonaje_Err:

   ' Call resetPj(UserIndex)

    Exit Sub

HandleResetearPersonaje_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetearPersonaje", Erl)
End Sub

Private Sub HandleRomperCania(ByVal UserIndex As Integer)

    On Error GoTo HandleRomperCania_Err:
    
    Dim LoopC As Integer
    Dim obj As t_Obj
    Dim caniaOld As Integer
    With UserList(UserIndex)
    
    obj.ObjIndex = .Invent.HerramientaEqpObjIndex
    caniaOld = .Invent.HerramientaEqpObjIndex
    obj.amount = 1
    For LoopC = 1 To MAX_INVENTORY_SLOTS
            
        'Rastreo la caña que está usando en el inventario y se la rompo
        If .Invent.Object(LoopC).ObjIndex = .Invent.HerramientaEqpObjIndex Then
            'Le quito una caña
            Call QuitarUserInvItem(UserIndex, LoopC, 1)
            Call UpdateUserInv(False, UserIndex, LoopC)
            Select Case caniaOld
                Case 881
                    obj.ObjIndex = 3457
                Case 2121
                    obj.ObjIndex = 3456
                Case 2132
                    obj.ObjIndex = 3459
                Case 2133
                    obj.ObjIndex = 3458
            End Select
            
            Call MeterItemEnInventario(UserIndex, obj)
            
            
            Exit Sub
            
        End If

262 Next LoopC

    End With
    
     'UserList(UserIndex).Invent.HerramientaEqpObjIndex
    
HandleRomperCania_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRomperCania", Erl)
End Sub
Private Sub HandleFinalizarPescaEspecial(ByVal UserIndex As Integer)

    On Error GoTo HandleFinalizarPescaEspecial_Err:
    
    Call EntregarPezEspecial(UserIndex)
    
    Exit Sub

HandleFinalizarPescaEspecial_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleFinalizarPescaEspecial", Erl)
End Sub

Private Sub HandleRepeatMacro(ByVal UserIndex As Integer)

    On Error GoTo HandleRepeatMacro_Err:
    'Call LogMacroCliente("El usuario " & UserList(UserIndex).name & " iteró el paquete click o u." & GetTickCount)
    Exit Sub

HandleRepeatMacro_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRepeatMacro", Erl)
End Sub

Private Sub HandleBuyShopItem(ByVal userindex As Integer)

    On Error GoTo HandleBuyShopItem_Err:
    Dim obj_to_buy As Long
        
    obj_to_buy = Reader.ReadInt32
    
    Call ModShopAO20.init_transaction(obj_to_buy, userindex)
    
    Exit Sub

HandleBuyShopItem_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBuyShopItem", Erl)
End Sub

Private Sub HandlePublicarPersonajeMAO(ByVal UserIndex As Integer)

    On Error GoTo HandlePublicarPersonajeMAO_Err:
    Dim Valor As Long
        
    Valor = Reader.ReadInt32
    
    If Valor <= MinimumPriceMao Then
'Msg1281= El valor de venta del personaje debe ser mayor que $¬1
Call WriteLocaleMsg(UserIndex, "1281", e_FontTypeNames.FONTTYPE_INFO, MinimumPriceMao)
        Exit Sub
    End If
    
    With UserList(UserIndex)
        ' Para recibir el ID del user
        Dim RS As ADODB.Recordset
        Set RS = Query("select is_locked_in_mao from user where id = ?;", .ID)
                    
        If EsGM(UserIndex) Then
'Msg1282= No podes vender un gm.
Call WriteLocaleMsg(UserIndex, "1282", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If CBool(RS!is_locked_in_mao) Then
'Msg1283= El personaje ya está publicado.
Call WriteLocaleMsg(UserIndex, "1283", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Stats.ELV < MinimumLevelMao Then
'Msg1284= No puedes publicar un personaje menor a nivel ¬1
Call WriteLocaleMsg(UserIndex, "1284", e_FontTypeNames.FONTTYPE_INFO, MinimumLevelMao)
            Exit Sub
        End If
        
        If .Stats.GLD < GoldPriceMao Then
'Msg1291= El costo para vender tu personajes es de ¬1 monedas de oro, no tienes esa cantidad.
Call WriteLocaleMsg(UserIndex, "1291", e_FontTypeNames.FONTTYPE_INFOBOLD, GoldPriceMao)
            Exit Sub
        Else
            .Stats.GLD = .Stats.GLD - GoldPriceMao
            Call WriteUpdateGold(UserIndex)
        End If
        Call Execute("update user set price_in_mao = ?, is_locked_in_mao = 1 where id = ?;", Valor, .ID)
        Call modNetwork.Kick(UserList(UserIndex).ConnectionDetails.ConnID, "El personaje fue publicado.")
    End With
        
    Exit Sub

HandlePublicarPersonajeMAO_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePublicarPersonajeMAO", Erl)
End Sub

Private Sub HandleDeleteItem(ByVal UserIndex As Integer)
    On Error GoTo HandleDeleteItem_Err:

    Dim Slot As Byte

    Slot = Reader.ReadInt8()

    With UserList(UserIndex)
        If Slot >= getMaxInventorySlots(UserIndex) Or Slot <= 0 Then Exit Sub
        
        If MapInfo(UserList(UserIndex).pos.Map).Seguro = 0 Or EsMapaEvento(.pos.Map) Then
'Msg1285= Solo puedes eliminar items en zona segura.
Call WriteLocaleMsg(UserIndex, "1285", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then
'Msg1286= No puede eliminar items cuando estas muerto.
Call WriteLocaleMsg(UserIndex, "1286", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Invent.Object(Slot).Equipped = 0 Then
            UserList(UserIndex).Invent.Object(Slot).amount = 0
            UserList(UserIndex).Invent.Object(Slot).Equipped = 0
            UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0
            Call UpdateUserInv(False, UserIndex, Slot)
'Msg1287= Objeto eliminado correctamente.
Call WriteLocaleMsg(UserIndex, "1287", e_FontTypeNames.FONTTYPE_INFO)
        Else
'Msg1288= No puedes eliminar un objeto estando equipado.
Call WriteLocaleMsg(UserIndex, "1288", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End With

    Exit Sub

HandleDeleteItem_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDeleteItem", Erl)
End Sub

Public Sub HandleActionOnGroupFrame(ByVal UserIndex As Integer)
On Error GoTo HandleActionOnGroupFrame_Err:
    Dim TargetGroupMember As Byte
    TargetGroupMember = Reader.ReadInt8
    
    With UserList(UserIndex)
        If Not .Grupo.EnGrupo Then Exit Sub
        If Not IsFeatureEnabled("target_group_frames") Then Exit Sub
        If UserList(.Grupo.Lider.ArrayIndex).Grupo.CantidadMiembros < TargetGroupMember Then Exit Sub
        If Not IsValidUserRef(UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember)) Then Exit Sub
        If UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember).ArrayIndex = UserIndex Then Exit Sub
        If UserMod.IsStun(.flags, .Counters) Then Exit Sub
        If .flags.Muerto = 1 Or .flags.Descansar Then Exit Sub
        Dim TargetUserIndex As Integer
        TargetUserIndex = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember).ArrayIndex
        If Abs(.pos.Map <> UserList(TargetUserIndex).pos.Map) Then Exit Sub
        If Abs(.pos.x - UserList(TargetUserIndex).pos.x) > RANGO_VISION_X Or Abs(.pos.y - UserList(TargetUserIndex).pos.y) > RANGO_VISION_Y Then Exit Sub
        If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
        If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
        If .flags.Meditando Then
            .flags.Meditando = False
            .Char.FX = 0
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
        End If
        .flags.targetUser = UserList(.Grupo.Lider.ArrayIndex).Grupo.Miembros(TargetGroupMember)
        If .flags.Hechizo > 0 Then
            If Not IsSet(Hechizos(UserList(UserIndex).Stats.UserHechizos(.flags.Hechizo)).SpellRequirementMask, e_SpellRequirementMask.eIsBindable) Then
                Call WriteLocaleMsg(UserIndex, MsgBindableHotkeysOnly, e_FontTypeNames.FONTTYPE_INFO)
                Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
            End If
            .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales + 1
            Call LanzarHechizo(.flags.Hechizo, UserIndex)
            Call WriteWorkRequestTarget(UserIndex, 0)
            If IsValidUserRef(.flags.GMMeSigue) Then
                Call WriteNofiticarClienteCasteo(.flags.GMMeSigue.ArrayIndex, 0)
            End If
            .flags.Hechizo = 0
        Else
            ' Msg587=¡Primero selecciona el hechizo que quieres lanzar!
            Call WriteLocaleMsg(UserIndex, "587", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
HandleActionOnGroupFrame_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleActionOnGroupFrame UserId:" & UserIndex, Erl)
End Sub

Public Sub HandleSendTelemetry(ByVal UserIndex As Integer)
On Error GoTo HandleSendTelemetry_Err:
    Dim TelemetryData(256) As Byte
    Dim TelemetrySize As Long
    Dim TelemetryIndex As Long
    TelemetrySize = Reader.ReadInt32
    Dim i As Long
    For i = 0 To TelemetrySize - 1
        TelemetryData(i) = Reader.ReadInt8
    Next i
    Dim TelemetryErrors(512) As Byte
    TelemetrySize = AOT_GetTelemetryResult(TelemetryData(0), TelemetrySize, UserList(UserIndex).ID, TelemetryErrors(0), 512)
    If TelemetrySize > 0 Then
        Call AddLogToCircularBuffer(StrConv(TelemetryErrors, vbUnicode))
    End If
    Exit Sub
HandleSendTelemetry_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSendTelemetry", Erl)
End Sub

Public Sub HandleSetHotkeySlot(ByVal UserIndex As Integer)
On Error GoTo HandleSetHotkeySlot_Err:
    With UserList(UserIndex)
        Dim SlotIndex As Byte
        Dim TargetIndex As Integer
        Dim LastKnownSlot As Integer
        Dim HkType As Byte
        SlotIndex = Reader.ReadInt8
        TargetIndex = Reader.ReadInt16
        LastKnownSlot = Reader.ReadInt16
        HkType = Reader.ReadInt8
        
        .HotkeyList(SlotIndex).Index = TargetIndex
        .HotkeyList(SlotIndex).LastKnownSlot = LastKnownSlot
        .HotkeyList(SlotIndex).Type = HkType
    End With
    Exit Sub
HandleSetHotkeySlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetHotkeySlot", Erl)
End Sub

Public Sub HandleUseHKeySlot(ByVal UserIndex As Integer)
On Error GoTo HandleUseHKeySlot_Err:
    Dim SlotIndex As Byte
    SlotIndex = Reader.ReadInt8
    If Not IsFeatureEnabled("hotokey-enabled") Then Exit Sub
    Dim CurrentSlotIndex As Integer
    Dim i As Integer
    With UserList(UserIndex)
        If .HotkeyList(SlotIndex).Index > 0 Then
            If .HotkeyList(SlotIndex).Type = Item Then
            ElseIf .HotkeyList(SlotIndex).Type = Spell Then
                If .HotkeyList(SlotIndex).LastKnownSlot > 0 And .HotkeyList(SlotIndex).LastKnownSlot < UBound(.Stats.UserHechizos) Then
                    If .Stats.UserHechizos(.HotkeyList(SlotIndex).LastKnownSlot) = .HotkeyList(SlotIndex).Index Then
                        CurrentSlotIndex = .HotkeyList(SlotIndex).LastKnownSlot
                    End If
                End If
                If CurrentSlotIndex = 0 Then
                    For i = LBound(.Stats.UserHechizos) To UBound(.Stats.UserHechizos)
                        If .Stats.UserHechizos(i) = .HotkeyList(SlotIndex).Index Then
                            CurrentSlotIndex = i
                            Exit For
                        End If
                    Next i
                End If
                If CurrentSlotIndex > 0 Then
                    If .Stats.UserHechizos(CurrentSlotIndex) > 0 Then
                        If IsSet(Hechizos(UserList(UserIndex).Stats.UserHechizos(CurrentSlotIndex)).SpellRequirementMask, e_SpellRequirementMask.eIsBindable) Then
                            Call UseSpellSlot(UserIndex, CurrentSlotIndex)
                        End If
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
HandleUseHKeySlot_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HandleUseHKeySlot", Erl)
End Sub


Public Sub HandleAntiCheatMessage(ByVal UserIndex As Integer)
On Error GoTo AntiCheatMessage_Err:
    Dim Data() As Byte
    Call Reader.ReadSafeArrayInt8(Data)
    Call HandleAntiCheatServerMessage(UserIndex, Data)
    Exit Sub
AntiCheatMessage_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.AntiCheatMessage", Erl)
End Sub

Public Sub HendleRequestLobbyList(ByVal UserIndex As Integer)
On Error GoTo HendleRequestLobbyList_Err:
    Call WriteUpdateLobbyList(UserIndex)
    Exit Sub
HendleRequestLobbyList_Err:
    Call TraceError(Err.Number, Err.Description, "Protocol.HendleRequestLobbyList", Erl)
End Sub
