Attribute VB_Name = "Extra"
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

Public m_NameIndex As New Dictionary
Public Function esCiudad(ByVal map As Integer) As Boolean

    Dim i As Byte
    
    For i = 0 To UBound(TotalMapasCiudades)
        If TotalMapasCiudades(i) = map Then
            esCiudad = True
            Exit Function
        End If
    Next i
    
End Function

Public Sub AgregarAConsola(ByVal Text As String)
        On Error GoTo AgregarAConsola_Err
        frmMain.List1.AddItem (Text)
        Exit Sub
AgregarAConsola_Err:
        Call TraceError(Err.Number, Err.Description, "ModLadder.AgregarAConsola", Erl)
End Sub

Public Function NameIndex(ByRef username As String) As t_UserReference
        Dim UserRef As t_UserReference
        Dim key As String
        Dim wrapper As clsUserRefWrapper
        key = UCase$(Replace(username, "+", " "))
        If m_NameIndex.Exists(key) Then
            Set wrapper = m_NameIndex(key)
            UserRef.ArrayIndex = wrapper.PlayerIndex
            UserRef.VersionId = wrapper.VersionId
        Else
            Call SetUserRef(UserRef, 0)
        End If
        NameIndex = UserRef
End Function

Public Sub FindLegalPos(ByVal UserIndex As Integer, ByVal Map As Integer, ByRef X As Byte, ByRef Y As Byte)
        '***************************************************
        'Autor: ZaMa
        'Last Modification: 26/03/2009
        'Search for a Legal pos for the user who is being teleported.
        '***************************************************
        
        On Error GoTo FindLegalPos_Err

100     If MapData(Map, X, Y).UserIndex <> 0 Or MapData(Map, X, Y).NpcIndex <> 0 Then
                    
            ' Se teletransporta a la misma pos a la que estaba
102         If MapData(Map, X, Y).UserIndex = UserIndex Then Exit Sub
                            
            Dim FoundPlace     As Boolean

            Dim tX             As Long

            Dim tY             As Long

            Dim Rango          As Long

            Dim OtherUserIndex As Integer
    
104         For Rango = 0 To 5
106             For tY = Y - Rango To Y + Rango
108                 For tX = X - Rango To X + Rango

                        'Reviso que no haya User ni NPC
110                     If MapData(Map, tX, tY).UserIndex = 0 And MapData(Map, tX, tY).NpcIndex = 0 Then
                        
112                         If InMapBounds(Map, tX, tY) Then FoundPlace = True
                        
                            Exit For

                        End If

114                 Next tX
        
116                 If FoundPlace Then Exit For
118             Next tY
            
120             If FoundPlace Then Exit For
122         Next Rango
    
124         If FoundPlace Then 'Si encontramos un lugar, listo, nos quedamos ahi
126             X = tX
128             Y = tY
            Else
                'Muy poco probable, pero..
                'Si no encontramos un lugar, sacamos al usuario que tenemos abajo, y si es un NPC, lo pisamos.
130             OtherUserIndex = MapData(Map, X, Y).UserIndex

132             If OtherUserIndex <> 0 Then

                    'Si no encontramos lugar, y abajo teniamos a un usuario, lo pisamos y cerramos su comercio seguro
134                 If IsValidUserRef(UserList(OtherUserIndex).ComUsu.DestUsu) Then

                        'Le avisamos al que estaba comerciando que se tuvo que ir.
136                     If UserList(UserList(OtherUserIndex).ComUsu.DestUsu.ArrayIndex).flags.UserLogged Then
138                         Call FinComerciarUsu(UserList(OtherUserIndex).ComUsu.DestUsu.ArrayIndex)
140                         Call WriteConsoleMsg(UserList(OtherUserIndex).ComUsu.DestUsu.ArrayIndex, "Comercio cancelado. El otro usuario se ha desconectado.", e_FontTypeNames.FONTTYPE_TALK)
                        

                        End If

                        'Lo sacamos.
142                     If UserList(OtherUserIndex).flags.UserLogged Then
144                         Call FinComerciarUsu(OtherUserIndex)
146                         Call WriteShowMessageBox(OtherUserIndex, "Alguien se ha conectado donde te encontrabas, por favor reconéctate...")
                        

                        End If

                    End If
            
148                 Call CloseSocket(OtherUserIndex)

                End If

            End If

        End If

        
        Exit Sub

FindLegalPos_Err:
150     Call TraceError(Err.Number, Err.Description, "Extra.FindLegalPos", Erl)

        
End Sub

Public Function EsNewbie(ByVal UserIndex As Integer) As Boolean
        On Error GoTo EsNewbie_Err
        
100     If UserIndex > 0 Then EsNewbie = UserList(UserIndex).Stats.ELV <= LimiteNewbie

        Exit Function

EsNewbie_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.EsNewbie", Erl)

        
End Function

Public Function esCiudadano(ByVal UserIndex As Integer) As Boolean
        On Error GoTo esCiudadano_Err
        
100     If UserIndex > 0 Then esCiudadano = Status(UserIndex) = Ciudadano
        
        Exit Function

esCiudadano_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.esCiudadano", Erl)

End Function

Public Function esCriminal(ByVal UserIndex As Integer) As Boolean
        On Error GoTo esCriminal_Err
        
100     If UserIndex > 0 Then esCriminal = Status(UserIndex) = Criminal
        
        Exit Function

esCriminal_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.esCriminal", Erl)

End Function

Public Function esArmada(ByVal UserIndex As Integer) As Boolean
        On Error GoTo esArmada_Err
        
100     If UserIndex > 0 Then esArmada = (UserList(UserIndex).Faccion.status = e_Facciones.Armada Or UserList(UserIndex).Faccion.status = e_Facciones.consejo)
        
        Exit Function

esArmada_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.esArmada", Erl)

        
End Function

Public Function esCaos(ByVal UserIndex As Integer) As Boolean
    On Error GoTo esCaos_Err
100 If UserIndex > 0 Then esCaos = (UserList(UserIndex).Faccion.Status = e_Facciones.Caos Or UserList(UserIndex).Faccion.Status = e_Facciones.concilio)
    Exit Function
esCaos_Err:
    Call TraceError(Err.Number, Err.Description, "Extra.esCaos", Erl)
End Function

Public Function FactionCanAttackFaction(ByVal AttackerFaction As e_Facciones, ByVal TargetFaction As e_Facciones)
    Select Case AttackerFaction
        Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo
            If TargetFaction = e_Facciones.Armada Or TargetFaction = e_Facciones.Ciudadano Or TargetFaction = e_Facciones.consejo Then
                Exit Function
            End If
        Case e_Facciones.Caos, e_Facciones.concilio
            If TargetFaction = e_Facciones.Caos Or TargetFaction = e_Facciones.concilio Then
                Exit Function
            End If
    End Select
    FactionCanAttackFaction = True
End Function

Public Function FactionCanHelpFaction(ByVal SourceFaction As e_Facciones, ByVal TargetFaction As e_Facciones) As e_InteractionResult
Select Case SourceFaction
    Case e_Facciones.Ciudadano, e_Facciones.Armada, e_Facciones.consejo
        If TargetFaction = e_Facciones.Caos Or TargetFaction = e_Facciones.concilio Then
            FactionCanHelpFaction = eOposingFaction
            Exit Function
        ElseIf TargetFaction = e_Facciones.Criminal Then
            FactionCanHelpFaction = eCantHelpCriminal
            Exit Function
        End If
    Case e_Facciones.Caos, e_Facciones.concilio
        If TargetFaction = e_Facciones.Armada Or TargetFaction = e_Facciones.consejo Or TargetFaction = e_Facciones.Ciudadano Then
            FactionCanHelpFaction = eOposingFaction
            Exit Function
        End If
    Case Else
End Select
    FactionCanHelpFaction = eInteractionOk
End Function

Public Function EsGM(ByVal UserIndex As Integer) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste)
        'Last Modification: 23/01/2007
        '***************************************************
        
        On Error GoTo EsGM_Err
        
100     If UserIndex = 0 Then
102         EsGM = False
            Exit Function
        End If
        
104     EsGM = (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) <> 0
        
        Exit Function

EsGM_Err:
106     Call TraceError(Err.Number, Err.Description, "Extra.EsGM", Erl)

        
End Function

Private Function CheckMapRestrictions(ByVal UserIndex As Integer, ByVal Map As Integer) As Boolean

100     With UserList(UserIndex)

102         If EsGM(UserIndex) Then
104             CheckMapRestrictions = True
                Exit Function
            End If


            Dim ZonaId As Integer
            ZonaId = MapData(Pos.Map).Tile(Pos.X, Pos.Y).ZonaId
            
106         If MapInfo(Map).Newbie Or Zona(ZonaId).Newbie And Not EsNewbie(UserIndex) Then
108             If .flags.UltimoMensaje <> 101 Then
110                 Call WriteConsoleMsg(UserIndex, "Sólo los newbies pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
112                 .flags.UltimoMensaje = 101
                End If
                Exit Function
            End If

114         If MapInfo(Map).NoPKs And (Status(UserIndex) = 0 Or Status(UserIndex) = 2) Then
116             If .flags.UltimoMensaje <> 102 Then
118                 Call WriteConsoleMsg(UserIndex, "Sólo los ciudadanos pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
120                 .flags.UltimoMensaje = 102
                End If
                Exit Function
            End If

122         If MapInfo(Map).NoCiudadanos And (Status(UserIndex) = 1 Or Status(UserIndex) = 3) Then
124             If .flags.UltimoMensaje <> 103 Then
126                 Call WriteConsoleMsg(UserIndex, "Sólo los criminales pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
128                 .flags.UltimoMensaje = 103
                End If
                Exit Function
            End If

130         If MapInfo(Map).SoloClanes And .GuildIndex <= 0 Then
132             If .flags.UltimoMensaje <> 104 Then
134                 Call WriteConsoleMsg(UserIndex, "Necesitas pertenecer a un clan para entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
136                 .flags.UltimoMensaje = 104
                End If
                Exit Function
            End If

138         If MapInfo(Map).MinLevel <> 0 And .Stats.ELV < MapInfo(Map).MinLevel Then
140             If .flags.UltimoMensaje <> 105 Then
142                 Call WriteConsoleMsg(UserIndex, "Necesitas ser al menos nivel " & MapInfo(Map).MinLevel & " para entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
144                 .flags.UltimoMensaje = 105
                End If
                Exit Function
            End If

146         If MapInfo(Map).MaxLevel <> 0 And .Stats.ELV >= MapInfo(Map).MaxLevel Then
148             If .flags.UltimoMensaje <> 106 Then
150                 Call WriteConsoleMsg(UserIndex, "Sólo los personajes inferiores a nivel " & MapInfo(Map).MaxLevel & " pueden entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
152                 .flags.UltimoMensaje = 106
                End If
                Exit Function
            End If
            
153         If MapInfo(Map).OnlyGroups And Not .Grupo.EnGrupo Then
154             If .flags.UltimoMensaje <> 107 Then
155                 Call WriteConsoleMsg(UserIndex, "Necesitas pertenecer a un grupo para entrar a este mapa.", e_FontTypeNames.FONTTYPE_INFO)
156                 .flags.UltimoMensaje = 107
                End If
                Exit Function
            End If

157         CheckMapRestrictions = True

        End With

End Function

Public Function GetTransportNextIndex(ByVal Map As Integer, ByVal PosX As Byte, ByVal PosY As Byte) As Integer
    Dim i As Integer
    With MapInfo(Map)
    For i = 0 To UBound(.TransportNetwork)
        If .TransportNetwork(i).TileX = PosX And .TransportNetwork(i).TileY = PosY Then
            GetTransportNextIndex = i
            Exit Function
        End If
    Next i
    End With
    GetTransportNextIndex = -1
End Function

Public Function GetExitTransport(ByVal Map As Integer, ByVal ExcludeIndex As Integer) As Integer
    Dim output As Integer
    Do
        output = RandomNumber(0, UBound(MapInfo(Map).TransportNetwork))
    Loop While output = excludeIndex
    GetExitTransport = output
End Function

Public Sub DoTileEvents(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)

        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Handles the Map passage of Users. Allows the existance
        'of exclusive maps for Newbies, Royal Army and Caos Legion members
        'and enables GMs to enter every map without restriction.
        'Uses: Mapinfo(map).Restringir = "NEWBIE" (newbies), "ARMADA", "CAOS", "FACCION" or "NO".
        '***************************************************
        On Error GoTo ErrHandler

        Dim nPos   As t_WorldPos

        Dim EsTeleport As Boolean
        Dim TelepRadio As Byte
        Dim aN As Integer
        Dim destPos As t_WorldPos
100     With UserList(UserIndex)
            'Controla las salidas
102         If InMapBounds(Map, X, Y) Then
                If MapData(Map, X, Y).trigger = AUTORESU Then
                    Call ResucitarOCurar(UserIndex)
                End If
104             If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
106                 EsTeleport = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = e_OBJType.otTeleport
                End If
                If Not MapData(map, X, y).Trap Is Nothing Then
                    Call ModMap.ActivateTrap(UserIndex, eUser, map, X, y)
                End If
                If EsTeleport Then
108                 If ObjData(MapData(map, X, y).ObjInfo.objIndex).Subtipo = e_TeleportSubType.eTransportNetwork Then
110                     Dim StartTransportIndex As Integer
112                     Dim ExitPortal As Integer
114                     StartTransportIndex = GetTransportNextIndex(map, X, y)
116                     If .LastTransportNetwork.map = map And .LastTransportNetwork.ExitIndex = StartTransportIndex Then
118                         ExitPortal = .LastTransportNetwork.StartIdex
                        Else
120                         ExitPortal = GetExitTransport(map, StartTransportIndex)
                        End If
122                     destPos = MapData(map, MapInfo(map).TransportNetwork(ExitPortal).TileX, MapInfo(map).TransportNetwork(ExitPortal).TileY).TileExit
124                     If destPos.map > 0 And destPos.map <= NumMaps Then
126                         .LastTransportNetwork.map = map
128                         .LastTransportNetwork.StartIdex = StartTransportIndex
130                         .LastTransportNetwork.ExitIndex = ExitPortal
132                         Call WarpUserChar(UserIndex, destPos.map, destPos.X, destPos.y, EsTeleport)
                        Else
134                         Call LogError("Invalid teleport at map: " & map & "(" & X & ", " & y & ")")
                        End If
                        Exit Sub
                    End If
                End If
                
    
136             If (MapData(map, X, y).TileExit.map > 0) And (MapData(map, X, y).TileExit.map <= NumMaps) Then
    
                    ' WyroX: Restricciones de mapas
138                 If CheckMapRestrictions(UserIndex, MapData(map, X, y).TileExit.map) Then
140                     If EsMapaInterdimensional(MapData(map, X, y).TileExit.map) And Not EsMapaInterdimensional(.pos.map) Then
142                         .flags.ReturnPos = .pos
                        End If
                        
144                     destPos.map = MapData(map, X, y).TileExit.map
                        If EsTeleport Then
146                         destPos.X = RandomNumber(MapData(map, X, y).TileExit.X - ObjData(MapData(map, X, y).ObjInfo.objIndex).Radio, MapData(map, X, y).TileExit.X + ObjData(MapData(map, X, y).ObjInfo.objIndex).Radio)
148                         destPos.y = RandomNumber(MapData(map, X, y).TileExit.y - ObjData(MapData(map, X, y).ObjInfo.objIndex).Radio, MapData(map, X, y).TileExit.y + ObjData(MapData(map, X, y).ObjInfo.objIndex).Radio)
                        Else
150                         destPos.X = MapData(map, X, y).TileExit.X
152                         destPos.y = MapData(map, X, y).TileExit.y
                        End If
                        
                        If .flags.Navegando Then
154                         Call ClosestLegalPos(destPos, nPos, True)
                        Else
156                         Call ClosestLegalPos(destPos, nPos)
                        End If
                        
158                     If nPos.X <> 0 And nPos.y <> 0 Then
160                         Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.y, EsTeleport)
                        End If
                    End If
    
                    'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
162                 Call ClearAttackerNpc(UserIndex)
    
164             ElseIf MapData(map, X, y).TileExit.map < 0 Then
166                 If .flags.ReturnPos.map <> 0 Then
168                     If LegalPos(.flags.ReturnPos.map, .flags.ReturnPos.X, .flags.ReturnPos.y, .flags.Navegando = 1, , , False) Then
170                         Call WarpUserChar(UserIndex, .flags.ReturnPos.map, .flags.ReturnPos.X, .flags.ReturnPos.y, False)
                        
                        Else
172                         Call ClosestLegalPos(.flags.ReturnPos, nPos)
                        
174                         If nPos.X <> 0 And nPos.y <> 0 Then
176                             Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.y, EsTeleport)
                            End If
                        End If
                        
                        'Te fusite del mapa. La criatura ya no es más tuya ni te reconoce como que vos la atacaste.
178                     Call ClearAttackerNpc(UserIndex)
                    End If
                End If
            End If

        End With

        Exit Sub

ErrHandler:
    Call TraceError(Err.Number, Err.Description, ".DotileEvents", Erl)
End Sub

Public Sub ClearAttackerNpc(ByVal UserIndex As Integer)
On Error GoTo ClearAttackerNpc_err
    With UserList(UserIndex)
        Dim aN As Integer
        If Not IsValidNpcRef(.flags.AtacadoPorNpc) Then
            Call ClearNpcRef(.flags.AtacadoPorNpc)
        Else
100         aN = .flags.AtacadoPorNpc.ArrayIndex

102         If aN > 0 Then
104             If IsValidUserRef(NpcList(aN).TargetUser) And NpcList(aN).TargetUser.ArrayIndex = UserIndex Then
106                 Call SetMovement(aN, NpcList(aN).flags.OldMovement)
108                 NpcList(aN).Hostile = NpcList(aN).flags.OldHostil
110                 NpcList(aN).flags.AttackedBy = vbNullString
112                 Call SetUserRef(NpcList(aN).TargetUser, 0)
114             End If
116         End If
        End If


        If Not IsValidNpcRef(.flags.NPCAtacado) Then
            Call ClearNpcRef(.flags.NPCAtacado)
        Else
118         aN = .flags.NPCAtacado.ArrayIndex
120         If aN > 0 Then
122             If NpcList(aN).flags.AttackedFirstBy = .Name Then
124                 NpcList(aN).flags.AttackedFirstBy = vbNullString
126             End If
128         End If
        End If
130        Call ClearNpcRef(.flags.AtacadoPorNpc)
132        Call ClearNpcRef(.flags.NPCAtacado)
    End With
    Exit Sub
ClearAttackerNpc_err:
    Call TraceError(Err.Number, Err.Description, "Extra.ClearAttackerNpc failed to clear userindex" & userIndex & "(" & UserList(userIndex).VersionId & ")", Erl)
End Sub

Function InRangoVision(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo InRangoVision_Err
        

100     If X > UserList(UserIndex).Pos.X - MinXBorder And X < UserList(UserIndex).Pos.X + MinXBorder Then
102         If Y > UserList(UserIndex).Pos.Y - MinYBorder And Y < UserList(UserIndex).Pos.Y + MinYBorder Then
104             InRangoVision = True
                Exit Function

            End If

        End If

106     InRangoVision = False

        
        Exit Function

InRangoVision_Err:
108     Call TraceError(Err.Number, Err.Description, "Extra.InRangoVision", Erl)

        
End Function

Function InRangoVisionNPC(ByVal NpcIndex As Integer, ByVal x As Integer, ByVal y As Integer) As Boolean
        
        On Error GoTo InRangoVisionNPC_Err
        

100     If X > NpcList(NpcIndex).Pos.X - MinXBorder And X < NpcList(NpcIndex).Pos.X + MinXBorder Then
102         If Y > NpcList(NpcIndex).Pos.Y - MinYBorder And Y < NpcList(NpcIndex).Pos.Y + MinYBorder Then
104             InRangoVisionNPC = True
                Exit Function

            End If

        End If

106     InRangoVisionNPC = False

        
        Exit Function

InRangoVisionNPC_Err:
108     Call TraceError(Err.Number, Err.Description, "Extra.InRangoVisionNPC", Erl)

        
End Function

Function InMapBounds(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo InMapBounds_Err
        
            
100     If (Map <= 0 Or Map > NumMaps) Or X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
102         InMapBounds = False
        Else
104         InMapBounds = True
        End If

        
        Exit Function

InMapBounds_Err:
106     Call TraceError(Err.Number, Err.Description, "Extra.InMapBounds", Erl)

        
End Function

Function ClosestLegalPosNPC(ByVal NpcIndex As Integer, ByVal MaxRange As Integer, Optional ByVal IgnoreUsers As Boolean, Optional ByVal IgnoreDeadUsers As Boolean) As t_WorldPos

        On Error GoTo ErrHandler

        Dim LoopC    As Integer
        Dim tX       As Integer
        Dim tY       As Integer
    
100     With NpcList(NpcIndex)

            Do
102             tY = .Pos.Y - LoopC
104             For tX = .Pos.X - LoopC To .Pos.X + LoopC
106                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                        Exit Function
                    End If
                Next

108             tX = .Pos.X - LoopC
110             For tY = .Pos.Y - LoopC + 1 To .Pos.Y + LoopC - 1
112                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                        Exit Function
                    End If
                Next

114             tX = .Pos.X + LoopC
116             For tY = .Pos.Y - LoopC + 1 To .Pos.Y + LoopC - 1
118                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                        Exit Function
                    End If
                Next

120             tY = .Pos.Y + LoopC
122             For tX = .Pos.X - LoopC To .Pos.X + LoopC
124                 If ValidNPCSpawnPos(ClosestLegalPosNPC, .Pos.map, tX, tY, .flags.AguaValida = 1, .flags.TierraInvalida = 0, IgnoreUsers, IgnoreDeadUsers) Then
                        Exit Function
                    End If
                Next

126             LoopC = LoopC + 1

128         Loop While LoopC <= MaxRange

        End With

        Exit Function

ErrHandler:
130     Call TraceError(Err.Number, Err.Description, "Extra.ClosestLegalPosNPC")

        
End Function

Private Function ValidNPCSpawnPos(OutPos As t_WorldPos, ByVal map As Integer, ByVal X As Integer, ByVal y As Integer, ByVal AguaValida As Boolean, ByVal TierraValida As Boolean, ByVal IgnoreUsers As Boolean, ByVal IgnoreDeadUsers As Boolean) As Boolean

100     If LegalPos(Map, X, Y, AguaValida, TierraValida, , False) Then
102         If TestSpawnTrigger(Map, X, Y) Then
104             If Not HayPCarea(map, X, y, IgnoreDeadUsers) Or IgnoreUsers Then
106                 ValidNPCSpawnPos = True
108                 OutPos.Map = Map
110                 OutPos.X = X
112                 OutPos.Y = Y
                    Exit Function
                End If
            End If
        End If

End Function

Sub ClosestLegalPos(Pos As t_WorldPos, ByRef nPos As t_WorldPos, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True)
        '*****************************************************************
        'Author: Unknown (original version)
        'Last Modification: 24/01/2007 (ToxicWaste)
        'Encuentra la posicion legal mas cercana y la guarda en nPos
        '*****************************************************************
        
        On Error GoTo ClosestLegalPos_Err
        

        Dim Notfound As Boolean

        Dim LoopC    As Integer

        Dim tX       As Integer

        Dim tY       As Integer

100     nPos.Map = Pos.Map

102     Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y, PuedeAgua, PuedeTierra, , False)

104         If LoopC > 12 Then
106             Notfound = True
                Exit Do

            End If
    
108         For tY = Pos.Y - LoopC To Pos.Y + LoopC
110             For tX = Pos.X - LoopC To Pos.X + LoopC
            
112                 If LegalPos(nPos.Map, tX, tY, PuedeAgua, PuedeTierra, , False) Then
114                     nPos.X = tX
116                     nPos.Y = tY
                        Exit Sub
                    End If
        
122             Next tX
124         Next tY
    
126         LoopC = LoopC + 1
    
        Loop

128     If Notfound = True Then
130         nPos.X = 0
132         nPos.Y = 0

        End If

        
        Exit Sub

ClosestLegalPos_Err:
134     Call TraceError(Err.Number, Err.Description, "Extra.ClosestLegalPos", Erl)

        
End Sub

Sub ClosestStablePos(Pos As t_WorldPos, ByRef nPos As t_WorldPos)
        '*****************************************************************
        'Encuentra la posicion legal mas cercana que no sea un portal y la guarda en nPos
        '*****************************************************************
        
        On Error GoTo ClosestStablePos_Err
        

        Dim Notfound As Boolean

        Dim LoopC    As Integer

        Dim tX       As Integer

        Dim tY       As Integer

100     nPos.Map = Pos.Map

102     Do While Not LegalPos(Pos.Map, nPos.X, nPos.Y)

104         If LoopC > 12 Then
106             Notfound = True
                Exit Do

            End If
    
108         For tY = Pos.Y - LoopC To Pos.Y + LoopC
110             For tX = Pos.X - LoopC To Pos.X + LoopC
            
112                 If LegalPos(nPos.Map, tX, tY) And MapData(nPos.Map, tX, tY).TileExit.Map = 0 Then
114                     nPos.X = tX
116                     nPos.Y = tY
                        '¿Hay objeto?
                
118                     tX = Pos.X + LoopC
120                     tY = Pos.Y + LoopC
  
                    End If
        
122             Next tX
124         Next tY
    
126         LoopC = LoopC + 1
    
        Loop

128     If Notfound = True Then
130         nPos.X = 0
132         nPos.Y = 0

        End If

        
        Exit Sub

ClosestStablePos_Err:
134     Call TraceError(Err.Number, Err.Description, "Extra.ClosestStablePos", Erl)

        
End Sub

Function IP_Index(ByVal inIP As String) As Integer
        
        On Error GoTo IP_Index_Err
        
 
        Dim UserIndex As Integer

        '¿Nombre valido?
100     If LenB(inIP) = 0 Then
102         IP_Index = 0
            Exit Function

        End If
  
104     UserIndex = 1

106     Do Until UserList(UserIndex).ConnectionDetails.IP = inIP
    
108         UserIndex = UserIndex + 1
    
110         If UserIndex > MaxUsers Then
112             IP_Index = 0
                Exit Function

            End If
    
        Loop
 
114     IP_Index = UserIndex

        Exit Function

        
        Exit Function

IP_Index_Err:
116     Call TraceError(Err.Number, Err.Description, "Extra.IP_Index", Erl)

        
End Function


Sub HeadtoPos(ByVal Head As e_Heading, ByRef Pos As t_WorldPos)
        On Error GoTo HeadtoPos_Err
        '*****************************************************************
        'Toma una posicion y se mueve hacia donde esta perfilado
        '*****************************************************************
        Dim X  As Integer
        Dim Y  As Integer
        Dim nX As Integer
        Dim nY As Integer
100     X = Pos.X
102     Y = Pos.Y
104     If Head = e_Heading.NORTH Then
106         nX = X
108         nY = Y - 1
        End If
110     If Head = e_Heading.SOUTH Then
112         nX = X
114         nY = Y + 1
        End If
116     If Head = e_Heading.EAST Then
118         nX = X + 1
120         nY = Y
        End If
122     If Head = e_Heading.WEST Then
124         nX = X - 1
126         nY = Y
        End If
        'Devuelve valores
128     Pos.X = nX
130     Pos.Y = nY
        Exit Sub
HeadtoPos_Err:
132     Call TraceError(Err.Number, Err.Description, "Extra.HeadtoPos", Erl)
End Sub

'Returns the front left positon from current heading
Public Sub GetHeadingLeft(ByVal head As e_Heading, ByRef pos As t_WorldPos)
        Dim X, Y, nX, nY As Integer
100     X = pos.X
102     Y = pos.Y
104     If head = e_Heading.NORTH Then
106         nX = X - 1
108         nY = Y - 1
        End If
110     If head = e_Heading.SOUTH Then
112         nX = X + 1
114         nY = Y + 1
        End If
116     If head = e_Heading.EAST Then
118         nX = X + 1
120         nY = Y - 1
        End If
122     If head = e_Heading.WEST Then
124         nX = X - 1
126         nY = Y + 1
        End If
        'Devuelve valores
128     pos.X = nX
130     pos.Y = nY
End Sub

'Returns the front right positon from current heading
Sub GetHeadingRight(ByVal head As e_Heading, ByRef pos As t_WorldPos)
        Dim X, Y, nX, nY As Integer
100     X = pos.X
102     Y = pos.Y
104     If head = e_Heading.NORTH Then
106         nX = X + 1
108         nY = Y - 1
        End If
110     If head = e_Heading.SOUTH Then
112         nX = X - 1
114         nY = Y + 1
        End If
116     If head = e_Heading.EAST Then
118         nX = X + 1
120         nY = Y + 1
        End If
122     If head = e_Heading.WEST Then
124         nX = X - 1
126         nY = Y - 1
        End If
        'Devuelve valores
128     pos.X = nX
130     pos.Y = nY
End Sub

' Autor: WyroX - 20/01/2021
' Retorna el heading recibo como parámetro pero rotado, según el valor R.
' Si R es 1, rota en sentido horario. Si R es -1, en sentido antihorario.
Function Rotate_Heading(ByVal Heading As e_Heading, ByVal R As Integer) As e_Heading
    
100     Rotate_Heading = (Heading + R + 3) Mod 4 + 1
    
End Function

Function LegalPos(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True, Optional ByVal PuedeBloqueoParcial As Boolean = True) As Boolean
        '***************************************************
        'Autor: Pablo (ToxicWaste) & Unknown (orginal version)
        'Last Modification: 23/01/2007
        'Checks if the position is Legal.
        '***************************************************
        '¿Es un mapa valido?
        
        On Error GoTo LegalPos_Err
        
100     If Map <= 0 Or Map > NumMaps Then Exit Function
        
102     If X < MinXBorder Or X > MaxXBorder Then Exit Function
        
104     If Y < MinYBorder Or Y > MaxYBorder Then Exit Function
        
106     With MapData(Map, X, Y)
        
108         If .NpcIndex <> 0 Then Exit Function

110         If .UserIndex <> 0 Then Exit Function
            
112         If Not PuedeTraslado Then
114             If .TileExit.Map > 0 Then Exit Function
            End If
            
116         If Not PuedeAgua Then
118             If (.Blocked And FLAG_AGUA) <> 0 Then Exit Function
            End If
            
120         If Not PuedeTierra Then
122             If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
            End If
            
            If PuedeBloqueoParcial Then
124             If (.Blocked And e_Block.ALL_SIDES) = e_Block.ALL_SIDES Then Exit Function
            Else
                If (.Blocked And e_Block.ALL_SIDES) > 0 Then Exit Function
            End If
            
        End With
        
        
126     LegalPos = True

        Exit Function

LegalPos_Err:
128     Call TraceError(Err.Number, Err.Description, "Extra.LegalPos", Erl)
        
End Function

Function LegalPosDestrabar(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True, Optional ByVal PuedeBloqueoParcial As Boolean = True) As Boolean
        On Error GoTo LegalPosDestrabar_Err
        
100     If Map <= 0 Or Map > NumMaps Then Exit Function
        
102     If X < MinXBorder Or X > MaxXBorder Then Exit Function
        
104     If Y < MinYBorder Or Y > MaxYBorder Then Exit Function
        
106     With MapData(Map, X, Y)
        
108         If .NpcIndex <> 0 Then Exit Function

            
112         If Not PuedeTraslado Then
114             If .TileExit.Map > 0 Then Exit Function
            End If
            
116         If Not PuedeAgua Then
118             If (.Blocked And FLAG_AGUA) <> 0 Then Exit Function
            End If
            
120         If Not PuedeTierra Then
122             If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
            End If
            
            If PuedeBloqueoParcial Then
124             If (.Blocked And e_Block.ALL_SIDES) = e_Block.ALL_SIDES Then Exit Function
            Else
                If (.Blocked And e_Block.ALL_SIDES) > 0 Then Exit Function
            End If
            
        End With
        
        
126     LegalPosDestrabar = True

        Exit Function
        
LegalPosDestrabar_Err:
128     Call TraceError(Err.Number, Err.Description, "Extra.LegalPosDestrabar", Erl)
        
End Function

Function LegalWalk(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As e_Heading, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal Montado As Boolean = False, Optional ByVal PuedeTraslado As Boolean = True, Optional ByVal WalkerIndex As Integer) As Boolean
        On Error GoTo LegalWalk_Err
        

100     If Map <= 0 Or Map > NumMaps Then Exit Function
        
102     If X < MinXBorder Or X > MaxXBorder Then Exit Function
        
104     If Y < MinYBorder Or Y > MaxYBorder Then Exit Function
        
106     With MapData(Map, X, Y)
        
108         If .NpcIndex <> 0 Then
              If Not IsSet(NpcList(.NpcIndex).flags.StatusMask, e_StatusMask.eDontBlockTile) Then Exit Function
            End If

110         If .UserIndex <> 0 Then
112             If UserList(.UserIndex).flags.AdminInvisible = 0 And UserList(.UserIndex).flags.Muerto = 0 Then
                    Exit Function
                End If
            End If
            
114         If Not PuedeTraslado Then
116             If .TileExit.Map > 0 Then Exit Function
            End If
            
118         If Not PuedeAgua Then
120             If (.Blocked And FLAG_AGUA) <> 0 And Not .trigger = e_Trigger.VALIDOPUENTE Then Exit Function
            End If
            
122         If Not PuedeTierra Then
124             If (.Blocked And FLAG_AGUA) = 0 Then Exit Function
            End If
            
            If .trigger = WORKERONLY Then
                If Not UserList(WalkerIndex).clase = Trabajador Then Exit Function
            End If
                       
            If (.Blocked And 2 ^ (Heading - 1)) <> 0 Then Exit Function
            
        
        End With
        
128     LegalWalk = True

        
        Exit Function

LegalWalk_Err:
130     Call TraceError(Err.Number, Err.Description, "Extra.LegalWalk", Erl)

        
End Function

Function LegalPosNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal AguaValida As Byte, Optional ByVal IsPet As Boolean = False) As Boolean
        
        On Error GoTo LegalPosNPC_Err
        

100     If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
102         LegalPosNPC = False
        
104     ElseIf MapData(Map, X, Y).TileExit.Map > 0 Then
106         LegalPosNPC = False
        
        Else

108         If AguaValida = 0 Then
110             LegalPosNPC = (MapData(Map, X, Y).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).trigger <> e_Trigger.POSINVALIDA Or IsPet) And (MapData(Map, X, Y).Blocked And FLAG_AGUA) = 0
            Else
112             LegalPosNPC = (MapData(Map, X, Y).Blocked And e_Block.ALL_SIDES) <> e_Block.ALL_SIDES And (MapData(Map, X, Y).UserIndex = 0) And (MapData(Map, X, Y).NpcIndex = 0) And (MapData(Map, X, Y).trigger <> e_Trigger.POSINVALIDA Or IsPet)
            End If
 
        End If

        
        Exit Function

LegalPosNPC_Err:
114     Call TraceError(Err.Number, Err.Description, "Extra.LegalPosNPC", Erl)

        
End Function

Function LegalWalkNPC(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Heading As e_Heading, Optional ByVal PuedeAgua As Boolean = False, Optional ByVal PuedeTierra As Boolean = True, Optional ByVal IgnoraInvalida As Boolean = False, Optional ByVal PuedePisar As Boolean, Optional ByVal esGuardia As Boolean = False) As Boolean
    ' Reescrito por WyroX

    On Error GoTo LegalWalkNPC_Err

100 If (Map <= 0 Or Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then
        Exit Function
    End If

102 With MapData(Map, X, Y)

104     If .TileExit.Map Then Exit Function

106     If Not PuedeAgua Then
108         If .Blocked And FLAG_AGUA Then
                Exit Function
            End If
        End If

110     If Not PuedeTierra Then
112         If (.Blocked And FLAG_AGUA) = 0 Then
                Exit Function
            End If
        End If

114     If Not PuedePisar Then
116         If .UserIndex Then
118             If UserList(.UserIndex).flags.AdminInvisible = 0 And UserList(.UserIndex).flags.Muerto = 0 Then
                    Exit Function
                End If
120         ElseIf .NpcIndex Then
                Exit Function
            End If
        End If
        
122     If Not IgnoraInvalida Then
124         If .trigger = e_Trigger.POSINVALIDA Then
                Exit Function
            End If
        End If
        
126     If Not esGuardia Then
128         If .Blocked And 2 ^ (Heading - 1) Then
                Exit Function
            End If
        Else
130         If (.Blocked And 2 ^ (Heading - 1)) And Not HayPuerta(Map, X + 1, Y) And Not HayPuerta(Map, X, Y) And Not HayPuerta(Map, X + 1, Y - 1) And Not HayPuerta(Map, X, Y - 1) Then Exit Function
        End If


    End With
    
132 LegalWalkNPC = True
    
    Exit Function

LegalWalkNPC_Err:
134     Call TraceError(Err.Number, Err.Description, "Extra.LegalWalkNPC", Erl)

        
End Function

Sub SendHelp(ByVal Index As Integer)
        
        On Error GoTo SendHelp_Err
        

        Dim NumHelpLines As Integer

        Dim LoopC        As Integer

100     NumHelpLines = val(GetVar(DatPath & "Help.dat", "INIT", "NumLines"))

102     For LoopC = 1 To NumHelpLines
104         Call WriteConsoleMsg(Index, GetVar(DatPath & "Help.dat", "Help", "Line" & LoopC), e_FontTypeNames.FONTTYPE_INFO)
106     Next LoopC

        
        Exit Sub

SendHelp_Err:
108     Call TraceError(Err.Number, Err.Description, "Extra.SendHelp", Erl)

        
End Sub

Public Sub Expresar(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo Expresar_Err
        

100     If NpcList(NpcIndex).NroExpresiones > 0 Then

            Dim randomi

102         randomi = RandomNumber(1, NpcList(NpcIndex).NroExpresiones)
104         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(NpcList(NpcIndex).Expresiones(randomi), NpcList(NpcIndex).Char.CharIndex, vbWhite))

        End If

        
        Exit Sub

Expresar_Err:
106     Call TraceError(Err.Number, Err.Description, "Extra.Expresar", Erl)

        
End Sub

Sub LookatTile(ByVal UserIndex As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo LookatTile_Err
        

        'Responde al click del usuario sobre el mapa
        Dim FoundChar      As Byte

        Dim FoundSomething As Byte

        Dim TempCharIndex  As Integer

        Dim Stat           As String

        Dim ft             As e_FontTypeNames

        '¿Rango Visión? (ToxicWaste)
100     If (Abs(UserList(UserIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(UserIndex).Pos.X - X) > RANGO_VISION_X) Then
            Exit Sub

        End If
        


        '¿Posicion valida?
102     If InMapBounds(Map, X, Y) Then
104         UserList(UserIndex).flags.TargetMap = Map
106         UserList(UserIndex).flags.TargetX = X
108         UserList(UserIndex).flags.TargetY = Y

            '¿Es un obj?
110         If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
                'Informa el nombre
112             UserList(UserIndex).flags.TargetObjMap = Map
114             UserList(UserIndex).flags.TargetObjX = X
116             UserList(UserIndex).flags.TargetObjY = Y
118             FoundSomething = 1
120         ElseIf MapData(Map, X + 1, Y).ObjInfo.ObjIndex > 0 Then

                'Informa el nombre
122             If ObjData(MapData(Map, X + 1, Y).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas Then
124                 UserList(UserIndex).flags.TargetObjMap = Map
126                 UserList(UserIndex).flags.TargetObjX = X + 1
128                 UserList(UserIndex).flags.TargetObjY = Y
130                 FoundSomething = 1

                End If

132         ElseIf MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then

134             If ObjData(MapData(Map, X + 1, Y + 1).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas Then
                    'Informa el nombre
136                 UserList(UserIndex).flags.TargetObjMap = Map
138                 UserList(UserIndex).flags.TargetObjX = X + 1
140                 UserList(UserIndex).flags.TargetObjY = Y + 1
142                 FoundSomething = 1

                End If

144         ElseIf MapData(Map, X, Y + 1).ObjInfo.ObjIndex > 0 Then

146             If ObjData(MapData(Map, X, Y + 1).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas Then
                    'Informa el nombre
148                 UserList(UserIndex).flags.TargetObjMap = Map
150                 UserList(UserIndex).flags.TargetObjX = X
152                 UserList(UserIndex).flags.TargetObjY = Y + 1
154                 FoundSomething = 1

                End If

            End If
    
156         If FoundSomething = 1 Then
158             UserList(UserIndex).flags.TargetObj = MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex

160             If MostrarCantidad(UserList(UserIndex).flags.TargetObj) Then

162                 Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "* - " & MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & "", e_FontTypeNames.FONTTYPE_INFO)
                Else

164                 If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otYacimiento Then
166                     Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
168                     Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).Name & " - (Minerales disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & ")", e_FontTypeNames.FONTTYPE_INFO)

170                 ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otArboles Then
172                     Call ActualizarRecurso(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY)
174                     Call WriteConsoleMsg(UserIndex, ObjData(UserList(UserIndex).flags.TargetObj).Name & " - (Recursos disponibles: " & MapData(Map, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.amount & ")", e_FontTypeNames.FONTTYPE_INFO)
                    
176                 ElseIf ObjData(UserList(UserIndex).flags.TargetObj).OBJType = e_OBJType.otTeleport Then
178                     If MapData(Map, X, Y).TileExit.Map > 0 Then
180                         If LenB(MapInfo(MapData(Map, X, Y).TileExit.Map).map_name) <> 0 Then
182                             Call WriteConsoleMsg(UserIndex, "Portal a " & MapInfo(MapData(Map, X, Y).TileExit.Map).map_name, e_FontTypeNames.FONTTYPE_INFO)
                            Else
184                             Call WriteConsoleMsg(UserIndex, "Portal a un mapa desconocido...", e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        Else
186                         Call WriteConsoleMsg(UserIndex, "Portal hacia la quinta dimensión", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
188                     Call WriteConsoleMsg(UserIndex, "O*" & UserList(UserIndex).flags.TargetObj & "*", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
    
            End If
            Dim yy As Byte
            '¿Es un personaje?
190         If Y + 1 <= YMaxMapSize Then
192             If MapData(Map, X, Y + 1).UserIndex > 0 Then
194                 TempCharIndex = MapData(Map, X, Y + 1).UserIndex
                    yy = y + 1
196                 FoundChar = 1

                End If

198             If MapData(Map, X, Y + 1).NpcIndex > 0 Then
200                 TempCharIndex = MapData(Map, X, Y + 1).NpcIndex
202                 FoundChar = 2

                End If

            End If

            '¿Es un personaje?
204         If FoundChar = 0 Then
206             If MapData(Map, X, Y).UserIndex > 0 Then
208                 TempCharIndex = MapData(Map, X, Y).UserIndex
                    yy = y
210                 FoundChar = 1

                End If

212             If MapData(Map, X, Y).NpcIndex > 0 Then
214                 TempCharIndex = MapData(Map, X, Y).NpcIndex
216                 FoundChar = 2

                End If

            End If
    
            'Reaccion al personaje
218         If FoundChar = 1 Then '  ¿Encontro un Usuario?
                
                If UserList(TempCharIndex).Pos.X <> X Or UserList(TempCharIndex).Pos.y <> yy Then
                    'Mapblock bugeado
                    MapData(map, X, yy).UserIndex = 0
                    Call TraceError(999, "Hay un mapblock con un usuario y el usuario tiene otra posicion. Index: " & TempCharIndex & ", user: " & UserList(TempCharIndex).name & ", mapblock: " & map & "," & X & "," & y & ", userpos: " & UserList(TempCharIndex).Pos.map & "," & UserList(TempCharIndex).Pos.X & "," & UserList(TempCharIndex).Pos.y, "LookatTile")
                End If
            
            
220             If UserList(TempCharIndex).flags.AdminInvisible = 0 Or CompararPrivilegiosUser(UserIndex, TempCharIndex) >= 0 Then
                    
222                 If UserList(TempCharIndex).showName Or CompararPrivilegiosUser(UserIndex, TempCharIndex) >= 0 Then
                
224                     If UserList(TempCharIndex).flags.Privilegios = user Then
                    
                            Dim Fragsnick As String
    
226                         If EsGM(UserIndex) Then
228                             Stat = Stat & " <" & ListaClases(UserList(TempCharIndex).clase) & " " & ListaRazas(UserList(TempCharIndex).raza) & " Nivel: " & UserList(TempCharIndex).Stats.ELV & ">"
                                Stat = Stat & " (ELO " & UserList(TempCharIndex).Stats.ELO & ")"
                            End If

232                         If EsNewbie(TempCharIndex) Then
234                             Stat = Stat & " <Newbie>"
                            End If
    
236                         If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) > 49 Then
238                             If UserList(TempCharIndex).flags.Envenenado > 0 Then
240                                 Fragsnick = " | Envenenado"
                                End If
    
242                             If UserList(TempCharIndex).flags.Ceguera = 1 Then
244                                 Fragsnick = Fragsnick & " | Ciego"
                                End If
    
246                             If UserList(TempCharIndex).flags.Incinerado = 1 Then
248                                 Fragsnick = Fragsnick & " | Incinerado"
                                End If
    
250                             If UserList(TempCharIndex).flags.Paralizado = 1 Then
252                                 Fragsnick = Fragsnick & " | Paralizado"
                                End If
    
254                             If UserList(TempCharIndex).flags.Inmovilizado = 1 Then
256                                 Fragsnick = Fragsnick & " | Inmovilizado"
                                End If
    
258                             If UserList(TempCharIndex).Counters.Trabajando > 0 Then
260                                 Fragsnick = Fragsnick & " | Trabajando"
                                End If
    
262                             If UserList(TempCharIndex).flags.invisible = 1 Then
264                                 Fragsnick = Fragsnick & " | Invisible"
                                End If
    
266                             If UserList(TempCharIndex).flags.Oculto = 1 Then
268                                 Fragsnick = Fragsnick & " | Oculto"
                                End If
    
270                             If UserList(TempCharIndex).flags.Estupidez = 1 Then
272                                 Fragsnick = Fragsnick & " | Estupido"
                                End If
    
274                             If UserList(TempCharIndex).flags.Maldicion = 1 Then
276                                 Fragsnick = Fragsnick & " | Maldito"
                                End If
    
278                             If UserList(TempCharIndex).flags.Silenciado = 1 Then
280                                 Fragsnick = Fragsnick & " | Silenciado"
                                End If
    
282                             If UserList(TempCharIndex).flags.Comerciando = True Then
284                                 Fragsnick = Fragsnick & " | Comerciando"
                                End If
    
286                             If UserList(TempCharIndex).flags.Descansar = 1 Then
288                                 Fragsnick = Fragsnick & " | Descansando"
                                End If
    
290                             If UserList(TempCharIndex).flags.Meditando Then
292                                 Fragsnick = Fragsnick & " | Concentrado"
                                End If
                                
294                             If UserList(TempCharIndex).Stats.MinHp = 0 Then
296                                 Stat = Stat & " <Muerto>"
298                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.1) Then
300                                 Stat = Stat & " <Casi muerto" & Fragsnick & ">"
302                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.5) Then
304                                 Stat = Stat & " <Malherido" & Fragsnick & ">"
306                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.75) Then
308                                 Stat = Stat & " <Herido" & Fragsnick & ">"
310                             ElseIf UserList(TempCharIndex).Stats.MinHp < (UserList(TempCharIndex).Stats.MaxHp * 0.99) Then
312                                 Stat = Stat & " <Levemente herido" & Fragsnick & ">"
                                Else
314                                 Stat = Stat & " <Intacto" & Fragsnick & ">"
                                End If
    
                            End If
                    
316                         If UserList(TempCharIndex).Faccion.status = e_Facciones.Armada Or UserList(TempCharIndex).Faccion.status = e_Facciones.consejo Then
318                             Stat = Stat & " <" & TituloReal(TempCharIndex) & ">"

320                             ft = e_FontTypeNames.FONTTYPE_CONSEJOVesA
322                         ElseIf UserList(TempCharIndex).Faccion.status = e_Facciones.Caos Or UserList(TempCharIndex).Faccion.status = e_Facciones.concilio Then
324                             Stat = Stat & " <" & TituloCaos(TempCharIndex) & ">"
326                             ft = e_FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
                            End If
                    
328                         If UserList(TempCharIndex).GuildIndex > 0 Then
330                             Stat = Stat & " <" & modGuilds.GuildName(UserList(TempCharIndex).GuildIndex) & ">"
                            End If
    
                        End If
                                
332                     If EsGM(TempCharIndex) Then
334                         ft = e_FontTypeNames.FONTTYPE_GM
                        
336                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.Consejero Then
338                             Stat = Stat & " <Consejero>"
                            End If
    
340                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.SemiDios Then
342                             Stat = Stat & " <Semi-Dios>"
                            End If
    
344                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.Dios Then
346                             Stat = Stat & " <Dios>"
                            End If
                            
348                         If UserList(TempCharIndex).flags.Privilegios = e_PlayerType.Admin Then
350                             Stat = Stat & " <Administrador>"
    
                            End If
                        'si es pk
                        ElseIf UserList(TempCharIndex).Faccion.Status = e_Facciones.concilio Then
                            Stat = Stat + " <Concilio Del Caos>"
                            ft = e_FontTypeNames.FONTTYPE_CONSEJOCAOS
                        ElseIf UserList(TempCharIndex).Faccion.Status = e_Facciones.Caos Then
                            ft = e_FontTypeNames.FONTTYPE_CRIMINAL_CAOS
                        ElseIf UserList(TempCharIndex).Faccion.Status = e_Facciones.Criminal Then
                            ft = e_FontTypeNames.FONTTYPE_CRIMINAL
                        ElseIf UserList(TempCharIndex).Faccion.Status = e_Facciones.consejo Then
                            ft = e_FontTypeNames.FONTTYPE_CONSEJO
                            Stat = Stat + " <Consejo de Banderbill>"
                        ElseIf UserList(TempCharIndex).Faccion.Status = e_Facciones.armada Then
                            ft = e_FontTypeNames.FONTTYPE_CITIZEN_ARMADA
                        ElseIf UserList(TempCharIndex).Faccion.Status = e_Facciones.Ciudadano Then
                            ft = e_FontTypeNames.FONTTYPE_CITIZEN
                        End If
                        
360                     If UserList(TempCharIndex).flags.Casado = 1 Then
362                         Stat = Stat & " <Pareja de " & GetUserSpouse(TempCharIndex) & ">"
                        End If
                        
364                     If Len(UserList(TempCharIndex).Desc) > 0 Then
366                         Stat = "Ves a " & UserList(TempCharIndex).Name & Stat & " - " & UserList(TempCharIndex).Desc
                        Else
368                         Stat = "Ves a " & UserList(TempCharIndex).Name & Stat
                        End If
                
370                     If LenB(Stat) > 0 Then
                            If UserList(UserIndex).flags.Muerto = 0 Or (UserList(UserIndex).GuildIndex > 0 And UserList(UserIndex).GuildIndex = UserList(TempCharIndex).GuildIndex) Or UserIndex = TempCharIndex Then
372                             If UserList(TempCharIndex).flags.Muerto Then
374                                 Call WriteConsoleMsg(UserIndex, Stat, e_FontTypeNames.FONTTYPE_New_Gris)
                                Else
376                                 Call WriteConsoleMsg(UserIndex, Stat, ft)
                                End If
                            End If
                        End If
                
378                     FoundSomething = 1
380                     Call SetUserRef(UserList(userIndex).flags.targetUser, TempCharIndex)
382                     Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
384                     UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun

                    Else
386                     Call WriteConsoleMsg(UserIndex, "Ves a ??? <Game Master>", e_FontTypeNames.FONTTYPE_GM)
    
388                     Call SetUserRef(UserList(userIndex).flags.targetUser, TempCharIndex)
390                     Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
392                     UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
    
                    End If
                End If
            End If

394         If FoundChar = 2 Then '¿Encontro un NPC?

                Dim estatus As String

396             If EsGM(UserIndex) Or UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 75 Then
398                 estatus = "<" & NpcList(TempCharIndex).Stats.MinHp & "/" & NpcList(TempCharIndex).Stats.MaxHp
                    
400             ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 50 Then
402                 If NpcList(TempCharIndex).Stats.MaxHp <> 0 Then
404                     estatus = "<" & Round((NpcList(TempCharIndex).Stats.MinHp / NpcList(TempCharIndex).Stats.MaxHp) * 100#, 0) & "%"
                    End If
 
406             ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 25 Then
                
408                 If NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.1) Then
410                     estatus = "<Agonizando"
412                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.2) Then
414                     estatus = "<Casi muerto"
416                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.5) Then
418                     estatus = "<Malherido"
420                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.7) Then
422                     estatus = "<Herido"
424                 ElseIf NpcList(TempCharIndex).Stats.MinHp < (NpcList(TempCharIndex).Stats.MaxHp * 0.95) Then
426                     estatus = "<Levemente herido"
                    Else
428                     estatus = "<Intacto"
                    End If
                    
                Else
430                 If NpcList(TempCharIndex).Stats.MinHp < NpcList(TempCharIndex).Stats.MaxHp Then
432                     estatus = "<Herido"
                    Else
434                     estatus = "<Intacto"
                    End If
                        
                End If
                        
436             If NpcList(TempCharIndex).flags.Envenenado > 0 Then
438                 estatus = estatus & " | Envenenado"
                End If
                        
440             If NpcList(TempCharIndex).flags.Paralizado = 1 Then
442                 If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 100 Then
444                     estatus = estatus & " | Paralizado (" & CInt(NpcList(TempCharIndex).Contadores.Paralisis / 6.5) & "s)"
                    Else
446                     estatus = estatus & " | Paralizado"
                    End If
                End If
                        
448             If NpcList(TempCharIndex).flags.Inmovilizado = 1 Then
450                 If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 100 Then
452                     estatus = estatus & " | Inmovilizado (" & CInt(NpcList(TempCharIndex).Contadores.Inmovilizado / 6.5) & "s)"
                    Else
454                     estatus = estatus & " | Inmovilizado"

                    End If

                End If

                If EsGM(UserIndex) And GetOwnedBy(TempCharIndex) <> 0 Then
                    estatus = estatus & " | Owned by " & NpcList(TempCharIndex).flags.AttackedBy
                End If
                        
456             estatus = estatus & ">"
    
                'End If
            
458             If Len(NpcList(TempCharIndex).Desc) > 1 Then
                    ' WyroX: Hacemos que se detenga a hablar un momento :P
460                 If NpcList(TempCharIndex).Movement = Caminata Then
462                     NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 5000 + Len(NpcList(TempCharIndex).Desc) * 50 - NpcList(TempCharIndex).IntervaloMovimiento ' 5 segundos + 1 segundo cada 20 caracteres
                    End If
                    
                    If UserList(UserIndex).flags.Muerto = 0 Or (UserList(UserIndex).flags.Muerto = 1 And NpcList(TempCharIndex).NPCtype = e_NPCType.Revividor) Then
                        If NpcList(TempCharIndex).NPCtype = e_NPCType.Quest Then
                            If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y, NpcList(TempCharIndex).Pos.X, NpcList(TempCharIndex).Pos.y) < 3 Then
                                If NpcList(TempCharIndex).Movement = Caminata Then
                                    NpcList(TempCharIndex).Contadores.IntervaloMovimiento = GetTickCount + 15000 - NpcList(TempCharIndex).IntervaloMovimiento ' 15 segundos
                                End If
                                
                                If NpcList(TempCharIndex).SoundOpen <> 0 Then
                                    Call WritePlayWave(UserIndex, NpcList(TempCharIndex).SoundOpen, NpcList(TempCharIndex).Pos.X, NpcList(TempCharIndex).Pos.y, 1)
                                End If
                                Call WriteChatOverHead(UserIndex, "NPCDESC*" & NpcList(TempCharIndex).Numero, NpcList(TempCharIndex).Char.charindex, vbWhite)
                            End If
                            
                            
                        Else
                            'Optimizacion de protocolo por Ladder
464                         Call WriteChatOverHead(UserIndex, "NPCDESC*" & NpcList(TempCharIndex).Numero, NpcList(TempCharIndex).Char.charindex, vbWhite)
                        End If
                    End If
470             ElseIf IsValidUserRef(NpcList(TempCharIndex).MaestroUser) Then
                    If UserList(UserIndex).flags.Muerto = 0 Then
472                     Call WriteConsoleMsg(userIndex, "NPCNAME*" & NpcList(TempCharIndex).Numero & "* es mascota de " & UserList(NpcList(TempCharIndex).MaestroUser.ArrayIndex).name & " " & estatus, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
                    If UserList(UserIndex).flags.Muerto = 0 Then
474                     Call WriteConsoleMsg(UserIndex, "NPCNAME*" & NpcList(TempCharIndex).Numero & "*" & " " & estatus, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
               ' End If
                
476             FoundSomething = 1
478             UserList(UserIndex).flags.TargetNpcTipo = NpcList(TempCharIndex).NPCtype
480             Call SetNpcRef(UserList(UserIndex).flags.TargetNPC, TempCharIndex)
482             Call SetUserRef(UserList(userIndex).flags.targetUser, 0)
484             UserList(UserIndex).flags.TargetObj = 0


            
            Dim i As Long, j As Long
                    
486                 For i = 1 To MAXUSERQUESTS
                    
488                     With UserList(UserIndex).QuestStats.Quests(i)
                    
490                         If .QuestIndex Then
492                             If QuestList(.QuestIndex).RequiredTargetNPCs Then
                    
494                                 For j = 1 To QuestList(.QuestIndex).RequiredTargetNPCs
                    
496                                     If QuestList(.QuestIndex).RequiredTargetNPC(j).NpcIndex = NpcList(TempCharIndex).Numero Then
498                                         If QuestList(.QuestIndex).RequiredTargetNPC(j).amount > .NPCsTarget(j) Then
500                                             .NPCsTarget(j) = .NPCsTarget(j) + 1
                    
                                            End If
                                                
502                                         If QuestList(.QuestIndex).RequiredTargetNPC(j).amount = .NPCsTarget(j) Then
504                                             Call FinishQuest(UserIndex, .QuestIndex, i)
506                                             Call WriteUpdateNPCSimbolo(UserIndex, TempCharIndex, 1)
508                                             Call WriteChatOverHead(UserIndex, "¡Quest Finalizada!", NpcList(TempCharIndex).Char.charindex, vbYellow)
510                                             Call WriteConsoleMsg(UserIndex, "Quest Finalizada!", e_FontTypeNames.FONTTYPE_INFOIAO)
                                            End If
                    
                                        End If
                    
512                                 Next j
                    
                                End If
                    
                            End If
                    
                        End With
                    
514                 Next i
        
            End If
    
516         If FoundChar = 0 Then
518             Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
520             UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
522             Call SetUserRef(UserList(userIndex).flags.targetUser, 0)

            End If
    
            '*** NO ENCONTRO NADA ***
524         If FoundSomething = 0 Then
526             Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
528             UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
530             Call SetUserRef(UserList(userIndex).flags.targetUser, 0)
532             UserList(UserIndex).flags.TargetObj = 0
534             UserList(UserIndex).flags.TargetObjMap = 0
536             UserList(UserIndex).flags.TargetObjX = 0
538             UserList(UserIndex).flags.TargetObjY = 0

                ' Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        Else

540         If FoundSomething = 0 Then
542             Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
544             UserList(UserIndex).flags.TargetNpcTipo = e_NPCType.Comun
546             Call SetUserRef(UserList(userIndex).flags.targetUser, 0)
548             UserList(UserIndex).flags.TargetObj = 0
550             UserList(UserIndex).flags.TargetObjMap = 0
552             UserList(UserIndex).flags.TargetObjX = 0
554             UserList(UserIndex).flags.TargetObjY = 0

                '  Call WriteConsoleMsg(UserIndex, "No ves nada interesante.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End If

        
        Exit Sub

LookatTile_Err:
556         Call TraceError(Err.Number, Err.Description & " Pos: " & Map & "-" & X & "-" & Y & " Usuario: " & UserList(UserIndex).Name & " Extra.LookatTile", Erl)
558     If FoundChar = 2 Then
560         Call TraceError(Err.Number, Err.Description & " Pos: " & Map & "-" & X & "-" & Y & "Npc: " & NpcList(TempCharIndex).Numero & " Usuario: " & UserList(UserIndex).Name & "Extra.LookatTile X LADDER", Erl)
        End If

        
End Sub

Function FindDirection(Pos As t_WorldPos, Target As t_WorldPos) As e_Heading
        
        On Error GoTo FindDirection_Err
        

        '*****************************************************************
        'Devuelve la direccion en la cual el target se encuentra
        'desde pos, 0 si la direc es igual
        '*****************************************************************
        Dim X As Integer

        Dim Y As Integer

100     X = Pos.X - Target.X
102     Y = Pos.Y - Target.Y

        'NE
104     If Sgn(X) = -1 And Sgn(Y) = 1 Then
106         FindDirection = IIf(RandomNumber(0, 1), e_Heading.NORTH, e_Heading.EAST)
            Exit Function

        End If

        'NW
108     If Sgn(X) = 1 And Sgn(Y) = 1 Then
110         FindDirection = IIf(RandomNumber(0, 1), e_Heading.WEST, e_Heading.NORTH)
            Exit Function

        End If

        'SW
112     If Sgn(X) = 1 And Sgn(Y) = -1 Then
114         FindDirection = IIf(RandomNumber(0, 1), e_Heading.WEST, e_Heading.SOUTH)
            Exit Function

        End If

        'SE
116     If Sgn(X) = -1 And Sgn(Y) = -1 Then
118         FindDirection = IIf(RandomNumber(0, 1), e_Heading.SOUTH, e_Heading.EAST)
            Exit Function

        End If

        'Sur
120     If Sgn(X) = 0 And Sgn(Y) = -1 Then
122         FindDirection = e_Heading.SOUTH
            Exit Function

        End If

        'norte
124     If Sgn(X) = 0 And Sgn(Y) = 1 Then
126         FindDirection = e_Heading.NORTH
            Exit Function

        End If

        'oeste
128     If Sgn(X) = 1 And Sgn(Y) = 0 Then
130         FindDirection = e_Heading.WEST
            Exit Function

        End If

        'este
132     If Sgn(X) = -1 And Sgn(Y) = 0 Then
134         FindDirection = e_Heading.EAST
            Exit Function

        End If

        'misma
136     If Sgn(X) = 0 And Sgn(Y) = 0 Then
138         FindDirection = 0
            Exit Function

        End If

        
        Exit Function

FindDirection_Err:
140     Call TraceError(Err.Number, Err.Description, "Extra.FindDirection", Erl)

        
End Function

'[Barrin 30-11-03]
Public Function ItemNoEsDeMapa(ByVal Index As Integer) As Boolean
        
        On Error GoTo ItemNoEsDeMapa_Err
        

100     ItemNoEsDeMapa = ObjData(Index).OBJType <> e_OBJType.otPuertas And ObjData(Index).OBJType <> e_OBJType.otCarteles And ObjData(Index).OBJType <> e_OBJType.otArboles And ObjData(Index).OBJType <> e_OBJType.otYacimiento And ObjData(Index).OBJType <> e_OBJType.otTeleport And ObjData(Index).OBJType <> e_OBJType.OtCorreo And ObjData(Index).OBJType <> e_OBJType.OtDecoraciones

        
        Exit Function

ItemNoEsDeMapa_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.ItemNoEsDeMapa", Erl)

        
End Function

'[/Barrin 30-11-03]

Public Function MostrarCantidad(ByVal Index As Integer) As Boolean
        
        On Error GoTo MostrarCantidad_Err
        
100     MostrarCantidad = ObjData(Index).OBJType <> e_OBJType.otPuertas And ObjData(Index).OBJType <> e_OBJType.otCarteles And ObjData(Index).OBJType <> e_OBJType.otYacimiento And ObjData(Index).OBJType <> e_OBJType.otArboles And ObjData(Index).OBJType <> e_OBJType.OtCorreo And ObjData(Index).OBJType <> e_OBJType.otTeleport

        
        Exit Function

MostrarCantidad_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.MostrarCantidad", Erl)

        
End Function

Public Function EsObjetoFijo(ByVal OBJType As e_OBJType) As Boolean
        
        On Error GoTo EsObjetoFijo_Err
        

100     EsObjetoFijo = OBJType = OBJType = e_OBJType.otCarteles Or OBJType = e_OBJType.otArboles Or OBJType = e_OBJType.otYacimiento Or OBJType = e_OBJType.OtDecoraciones

        
        Exit Function

EsObjetoFijo_Err:
102     Call TraceError(Err.Number, Err.Description, "Extra.EsObjetoFijo", Erl)

        
End Function

Public Function HayPuerta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
100     If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then
102         HayPuerta = (ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).OBJType = e_OBJType.otPuertas) And ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 And (ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Llave = 0)
        End If
End Function

Public Sub CargarMapasEspeciales()
        
        Dim File As clsIniManager
100     Set File = New clsIniManager
    
102     Call File.Initialize(DatPath & "MapasEspeciales.dat")
    
        Dim Cantidad As Integer
104     Cantidad = val(File.GetValue("MapasInterdimensionales", "Cantidad"))
    
        Dim i As Integer
    
106     If Cantidad > 0 Then
108         ReDim MapasInterdimensionales(1 To Cantidad)
        
110         For i = 1 To Cantidad
112             MapasInterdimensionales(i) = val(File.GetValue("MapasInterdimensionales", "Mapa" & i))
            Next
        Else
114         ReDim MapasInterdimensionales(0)
        End If
    
        
117     Cantidad = val(File.GetValue("MapasEventos", "Cantidad"))
    
128     If Cantidad > 0 Then
130         ReDim MapasEventos(1 To Cantidad)
        
132         For i = 1 To Cantidad
134             MapasEventos(i) = val(File.GetValue("MapasEventos", "Mapa" & i))
            Next
        Else
136         ReDim MapasEventos(0)
        End If
        
    
138     Set File = Nothing

End Sub

Public Function EsMapaEvento(ByVal destMap As Long) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(MapasEventos)
        If MapasEventos(i) = destMap Then
            EsMapaEvento = True
            Exit Function
        End If
    Next i
    EsMapaEvento = False
    
End Function

Public Sub resetPj(ByVal UserIndex As Integer, Optional ByVal borrarHechizos As Boolean = False)


100     With UserList(UserIndex)

140         .flags.Muerto = False
142         .flags.Escondido = 0

144         .flags.Casado = 0
146         .flags.SpouseId = 0

            '%%%%%%%%%%%%% PREVENIR HACKEO DE LOS SKILLS %%%%%%%%%%%%%
            .Stats.SkillPts = 10
960         Call WriteLevelUp(UserIndex, 10)


164         Call DarCuerpo(UserIndex) 'Ladder REVISAR

166         .OrigChar = .Char
            Dim i As Long
            For i = 1 To NUMSKILLS
                .Stats.UserSkills(i) = 100
            Next i

168         .Char.WeaponAnim = NingunArma
170         .Char.ShieldAnim = NingunEscudo
172         .Char.CascoAnim = NingunCasco
173         .char.CartAnim = NoCart

            ' WyroX: Vida inicial
174         .Stats.MaxHp = .Stats.UserAtributos(e_Atributos.Constitucion)
176         .Stats.MinHp = .Stats.MaxHp

            ' WyroX: Maná inicial
178         .Stats.MaxMAN = .Stats.UserAtributos(e_Atributos.Inteligencia) * ModClase(.clase).ManaInicial
180         .Stats.MinMAN = .Stats.MaxMAN

            Dim MiInt As Integer
182         MiInt = RandomNumber(1, .Stats.UserAtributos(e_Atributos.Agilidad) \ 6)

184         If MiInt = 1 Then MiInt = 2

186         .Stats.MaxSta = 20 * MiInt
188         .Stats.MinSta = 20 * MiInt

190         .Stats.MaxAGU = 100
192         .Stats.MinAGU = 100

194         .Stats.MaxHam = 100
196         .Stats.MinHam = 100

202         .flags.VecesQueMoriste = 0
204         .flags.Montado = 0

206         .Stats.MaxHit = 2
208         .Stats.MinHIT = 1

212         .Stats.Exp = 0
214         .Stats.ELV = 1

            .Stats.GLD = 0
            .Stats.Banco = 0
            

             If .flags.TomoPocion Then

                For i = 1 To 4
                    .Stats.UserAtributos(i) = .Stats.UserAtributosBackUP(i)
                Next i

198             Call WriteFYA(UserIndex)

            End If

            .flags.DuracionEfecto = 0

            Call VaciarInventario(UserIndex)
            Call ResetCd(UserList(UserIndex))

216         Call RellenarInventario(UserIndex)
            'Agrego la poción

            Dim slot_libre As Byte

            For i = 1 To MAX_INVENTORY_SLOTS
                If .Invent.Object(i).amount = 0 Then
                    slot_libre = i
                    Exit For
                End If
            Next i
            
            For i = 1 To MAX_BANCOINVENTORY_SLOTS
102             .BancoInvent.Object(i).amount = 0
104             .BancoInvent.Object(i).Equipped = 0
106             .BancoInvent.Object(i).ObjIndex = 0
            Next i

            .Invent.Object(slot_libre).ObjIndex = POCION_RESET
            .Invent.Object(slot_libre).amount = 1

            'Valores Default de facciones al Activar nuevo usuario
222         Call ResetFacciones(UserIndex)

224         .Faccion.Status = 1
            
            If borrarHechizos Then
                Call ResetUserSpells(UserIndex)
575         '    Call UpdateUserHechizos(True, UserIndex, 0)
            End If
            Call UpdateUserHechizos(True, UserIndex, 0)
900         Call WriteUpdateUserStats(UserIndex)
905         Call WriteUpdateHungerAndThirst(UserIndex)
570         Call UpdateUserInv(True, UserIndex, 0)

            Call Execute("update user set is_reset = 1 where id = ?;", .ID)
            Call Execute("update quest set quest_id = 0, npcs = 0, npcstarget = 0 where user_id = ?;", .ID)
            Call Execute("delete from quest_done where user_id = ?;", .ID)
            
            Call ResetQuestStats(UserIndex)
            
            Call WarpUserChar(UserIndex, .Pos.Map, .Pos.X, .Pos.Y, True)
        End With

    'Call WarpUserChar(UserIndex, 1, 55, 45, True)
End Sub


Public Sub ResucitarOCurar(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.Muerto = 1 Then
        Call RevivirUsuario(UserIndex)
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
        UserList(UserIndex).flags.Envenenado = False
            
        Call WriteUpdateHP(UserIndex)
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(20, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
        UserList(UserIndex).Counters.timeFx = 3
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 35, 1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
        
        Call WriteConsoleMsg(UserIndex, "¡¡Hás sido resucitado!!", e_FontTypeNames.FONTTYPE_INFO)
    ElseIf UserList(UserIndex).Stats.MinHp < UserList(UserIndex).Stats.MaxHp Then
        UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
        UserList(UserIndex).flags.Envenenado = False
            
        Call WriteUpdateHP(UserIndex)
        
        UserList(UserIndex).Counters.timeFx = 3
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, 9, 1, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(18, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
            
        Call WriteConsoleMsg(UserIndex, "¡¡Hás sido curado!!", e_FontTypeNames.FONTTYPE_INFO)
    End If
End Sub

Public Function ByteArr2String(ByRef arr() As Byte) As String
    
    Dim str As String
    Dim i As Long
    For i = 0 To UBound(arr)
        str = str + Chr$(arr(i))
    Next i
    
    ByteArr2String = str
    
End Function


Public Function PacketIdToString(ByVal PacketId As Long) As String

    Select Case PacketId
        Case 1
            PacketIdToString = "CastSpell"
            Exit Function
        Case 2
            PacketIdToString = "WorkLeftClick"
            Exit Function
        Case 3
            PacketIdToString = "LeftClick"
            Exit Function
        Case 4
            PacketIdToString = "UseItem"
            Exit Function
        Case 5
            PacketIdToString = "UseItemU"
            Exit Function
        Case 6
            PacketIdToString = "Walk"
            Exit Function
        Case 7
            PacketIdToString = "Sailing"
            Exit Function
        Case 8
            PacketIdToString = "Talk"
            Exit Function
        Case 9
            PacketIdToString = "Attack"
            Exit Function
        Case 10
            PacketIdToString = "Drop"
            Exit Function
        Case 11
            PacketIdToString = "Work"
            Exit Function
        Case 12
            PacketIdToString = "EquipItem"
            Exit Function
        Case 13
            PacketIdToString = "GuildMessage"
            Exit Function
        Case 14
            PacketIdToString = "QuestionGM"
            Exit Function
        Case 15
            PacketIdToString = "ChangeHeading"
            Exit Function
        Case 16
            PacketIdToString = "Hide"
            Exit Function
    End Select
    
End Function

Public Sub TimerQuestOrco()
    Dim UserIndex As Integer
    
        For UserIndex = 1 To LastUser
        If UserIndex > 0 Then
        With UserList(UserIndex)
            If .flags.UserLogged Then
                Dim prob As Long, estimatedProb As Long
                
                prob = RandomNumber(1, LastUser)
                estimatedProb = LastUser / 3
                If Prob < estimatedProb And Not EsGM(UserIndex) Then
                    UserList(UserIndex).Stats.MinHam = 0
                    UserList(UserIndex).Stats.MinAGU = 0
122                 Call WriteUpdateHungerAndThirst(UserIndex)
202                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 20, 50, False))
                End If
            End If
        End With
        End If
    Next UserIndex
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(156, NO_3D_SOUND, NO_3D_SOUND))
End Sub

Public Function TestRequiredEquipedItem(ByRef inventory As t_Inventario, ByVal RequiredItemsFlag As Long, ByVal RequiredWeaponMask As Integer) As e_SpellRequirementMask
    If IsSet(requiredItemsFlag, e_SpellRequirementMask.eArmor) And _
      inventory.ArmourEqpObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eArmor
        Exit Function
    End If
    If IsSet(requiredItemsFlag, e_SpellRequirementMask.eHelm) And _
      inventory.CascoEqpObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eHelm
        Exit Function
    End If
    If IsSet(requiredItemsFlag, e_SpellRequirementMask.eMagicItem) And _
      inventory.MagicoObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eMagicItem
        Exit Function
    End If
    If IsSet(requiredItemsFlag, e_SpellRequirementMask.eProjectile) And _
      inventory.MunicionEqpObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eProjectile
        Exit Function
    End If
    If IsSet(requiredItemsFlag, e_SpellRequirementMask.eShield) And _
      inventory.EscudoEqpObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eShield
        Exit Function
    End If
    If IsSet(requiredItemsFlag, e_SpellRequirementMask.eShip) And _
      inventory.BarcoObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eShip
        Exit Function
    End If
    If IsSet(requiredItemsFlag, e_SpellRequirementMask.eTool) And _
      inventory.HerramientaEqpObjIndex = 0 Then
        TestRequiredEquipedItem = e_SpellRequirementMask.eTool
        Exit Function
    End If
    If IsSet(RequiredItemsFlag, e_SpellRequirementMask.eWeapon) Then
        If inventory.WeaponEqpObjIndex = 0 Then
            If Not IsIntSet(RequiredWeaponMask, e_WeaponType.eFist) Then
                TestRequiredEquipedItem = e_SpellRequirementMask.eWeapon
                Exit Function
            End If
        ElseIf RequiredWeaponMask > 0 Then
            If Not IsIntSet(RequiredWeaponMask, ShiftLeft(1, ObjData(inventory.WeaponEqpObjIndex).WeaponType)) Then
                TestRequiredEquipedItem = e_SpellRequirementMask.eWeapon
                Exit Function
            End If
        End If
    End If
    TestRequiredEquipedItem = e_SpellRequirementMask.eNone
End Function

Public Sub SendrequiredItemMessage(ByVal UserIndex As Integer, ByVal itemMask As e_SpellRequirementMask, ByVal Message As String)
    Select Case itemMask
        Case e_SpellRequirementMask.eArmor
             Call WriteConsoleMsg(UserIndex, "Necesitás una armadura " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eHelm
             Call WriteConsoleMsg(UserIndex, "Necesitás un casco " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eKnucle
             Call WriteConsoleMsg(UserIndex, "Necesitás unos nudillos " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eMagicItem
             Call WriteConsoleMsg(UserIndex, "Necesitás un objeto magico " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eProjectile
             Call WriteConsoleMsg(UserIndex, "Necesitás municiones " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eShield
             Call WriteConsoleMsg(UserIndex, "Necesitás un escudo " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eShip
             Call WriteConsoleMsg(UserIndex, "Necesitás un barco " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eTool
             Call WriteConsoleMsg(UserIndex, "Necesitás una herramienta " + Message, e_FontTypeNames.FONTTYPE_INFO)
        Case e_SpellRequirementMask.eWeapon
             Call WriteConsoleMsg(UserIndex, "Necesitás un arma " + Message, e_FontTypeNames.FONTTYPE_INFO)
    End Select
End Sub

Public Sub SendHelpInteractionMessage(ByVal UserIndex As Integer, ByVal CanHelpResult As e_InteractionResult)
    Select Case CanHelpResult
            Case eOposingFaction
108             Call WriteLocaleMsg(UserIndex, MsgFactionForbidAction, e_FontTypeNames.FONTTYPE_INFO)
            Case eCantHelpCriminal
110             Call WriteLocaleMsg(UserIndex, MsgDisableAttackGuardToContinue, e_FontTypeNames.FONTTYPE_INFO)
            Case eCantHelpCriminalClanRules
112             Call WriteLocaleMsg(UserIndex, MsgClanForbidAction, e_FontTypeNames.FONTTYPE_INFO)
            Case Else
                Debug.Assert "Invalid message"
        End Select
End Sub

Public Sub SendAttackInteractionMessage(ByVal UserIndex As Integer, ByVal CanAttackResult As e_AttackInteractionResult)
    Select Case CanAttackResult
        Case e_AttackInteractionResult.eCanAttack
            'Do nothing
        Case e_AttackInteractionResult.eDeathAttacker
            Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eFightActive
            Call WriteConsoleMsg(UserIndex, "No podés atacar en este momento.", e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eDeathTarget
            Call WriteConsoleMsg(UserIndex, "No podés atacar a un espiritu.", e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eSameGroup
            Call WriteConsoleMsg(UserIndex, "No podés atacar a un miembro de tu grupo.", e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eTalkWithMaster
            Call WriteConsoleMsg(UserIndex, "No podés atacar usuarios mientras estás en consulta.", e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eAttackerIsCursed
            Call WriteConsoleMsg(UserIndex, "No podés atacar usuarios mientras estás en consulta.", e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eMounted
            Call WriteConsoleMsg(UserIndex, "No podés atacar usando una montura.", e_FontTypeNames.FONTTYPE_INFO)
        Case e_AttackInteractionResult.eSameTeam
            Call WriteConsoleMsg(UserIndex, "No podes atacar un miembro de tu equipo.", e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eNotEnougthPrivileges
            Call WriteConsoleMsg(UserIndex, "El ser es demasiado poderoso", e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eSameClan
            Call WriteConsoleMsg(UserIndex, "No podes atacar a un miembro de tu clan.", e_FontTypeNames.FONTTYPE_INFOIAO)
        Case e_AttackInteractionResult.eSameFaction
            Call WriteLocaleMsg(UserIndex, MsgFacctionForbidAttack, e_FontTypeNames.FONTTYPE_INFOIAO)
        Case e_AttackInteractionResult.eRemoveSafe
            Call WriteLocaleMsg(UserIndex, MsgRemoveSafeToAttack, e_FontTypeNames.FONTTYPE_INFOIAO)
        Case e_AttackInteractionResult.eSafeArea
            Call WriteConsoleMsg(UserIndex, "Esta es una zona segura, aqui no podes atacar otros usuarios.", e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eCantAttackYourself
            Call WriteLocaleMsg(UserIndex, MsgCantAttackYourself, e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eAttackSameFaction
            Call WriteLocaleMsg(UserIndex, "¡Atacaste un ciudadano! Te has convertido en un Criminal.", e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eAttackPetSameFaction
            Call WriteLocaleMsg(UserIndex, "¡Atacaste una mascota de un ciudadano! Te has convertido en un Criminal.", e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eRemoveSafeCitizenNpc
            Call WriteConsoleMsg(UserIndex, "Debes quitar el seguro para poder atacar la criatura que esta luchando con otro usuario.", e_FontTypeNames.FONTTYPE_WARNING)
        Case e_AttackInteractionResult.eAttackCitizenNpc
            Call WriteConsoleMsg(UserIndex, "¡Atacaste una criatura de otro usuario! Te has convertido en un Criminal.", e_FontTypeNames.FONTTYPE_WARNING)
        Case Else
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFOIAO)
    End Select
End Sub

Public Function PreferedTileForDirection(ByRef Direction As t_Vector, ByRef CurrentPosition As t_WorldPos) As t_WorldPos
    Dim Ret As t_WorldPos
    Dim Normal As t_Vector
    Ret.Map = CurrentPosition.Map
    Normal = GetNormal(Direction)
    If Abs(Normal.x) > Abs(Normal.y) Then
        Ret.x = CurrentPosition.x + 1 * Sgn(Normal.x)
        Ret.y = CurrentPosition.y
    ElseIf Abs(Normal.x) < Abs(Normal.y) Then
        Ret.x = CurrentPosition.x
        Ret.y = CurrentPosition.y + 1 * Sgn(Normal.y)
    Else
        Ret.y = CurrentPosition.y + 1 * Sgn(Normal.y)
        Ret.x = CurrentPosition.x + 1 * Sgn(Normal.x)
    End If
    If Not LegalPos(Ret.Map, Ret.x, Ret.y, False, True) Then
        Call ClosestStablePos(Ret, Ret)
    End If
    If Ret.x = 0 And Ret.y = 0 Then
        PreferedTileForDirection = CurrentPosition
    Else
        PreferedTileForDirection = Ret
    End If
End Function
