Attribute VB_Name = "Protocol_GmCommands"
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

' Handles the "Online" message.
Public Sub HandleOnline(ByVal UserIndex As Integer)
        
        On Error GoTo HandleOnline_Err
        'Ladder 17/12/20 : Envio records de usuarios y uptime
        
        Dim i         As Long
        Dim count     As Long
        Dim Time      As Long
        Dim UpTimeStr As String
    
100     With UserList(UserIndex)
102         For i = 1 To LastUser

104             If UserList(i).flags.UserLogged Then
110                 count = count + 1

                End If

112         Next i
        
            'Get total time in seconds
114         Time = ((GetTickCount()) - tInicioServer) \ 1000
        
            'Get times in dd:hh:mm:ss format
116         UpTimeStr = (Time Mod 60) & " segundos."
118         Time = Time \ 60
        
120         UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
122         Time = Time \ 60
        
124         UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
126         Time = Time \ 24
        
128         If Time = 1 Then
130             UpTimeStr = Time & " día, " & UpTimeStr
            Else
132             UpTimeStr = Time & " días, " & UpTimeStr
    
            End If
    
134         Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, e_FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Número de usuarios: " & CStr(Count) & " conectados.", e_FontTypeNames.FONTTYPE_INFOIAO)
            Call WriteConsoleMsg(UserIndex, "Tiempo en línea: " & UpTimeStr & " Record de usuarios en simultaneo: " & RecordUsuarios & ".", e_FontTypeNames.FONTTYPE_INFOIAO)
        End With
        
        Exit Sub

HandleOnline_Err:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnline", Erl)
148
        
End Sub

' Handles the "Help" message.
Public Sub HandleHelp(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo HandleHelp_Err

100     Call SendHelp(UserIndex)
        
        Exit Sub

HandleHelp_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHelp", Erl)
104
        
End Sub

' Handles the "RequestStats" message.
Public Sub HandleRequestStats(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo HandleRequestStats_Err

100     Call SendUserStatsTxt(UserIndex, UserIndex)
        
        Exit Sub

HandleRequestStats_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestStats", Erl)
104
        
End Sub

' Handles the "RequestMOTD" message.
Public Sub HandleRequestMOTD(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo HandleRequestMOTD_Err

100     Call SendMOTD(UserIndex)
        
        Exit Sub

HandleRequestMOTD_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestMOTD", Erl)
104
        
End Sub


' Handles the "UpTime" message.
Public Sub HandleUpTime(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo HandleUpTime_Err

        Dim Time      As Long
        Dim UpTimeStr As String
    
        'Get total time in seconds
100     Time = ((GetTickCount()) - tInicioServer) \ 1000
    
        'Get times in dd:hh:mm:ss format
102     UpTimeStr = (Time Mod 60) & " segundos."
104     Time = Time \ 60
    
106     UpTimeStr = (Time Mod 60) & " minutos, " & UpTimeStr
108     Time = Time \ 60
    
110     UpTimeStr = (Time Mod 24) & " horas, " & UpTimeStr
112     Time = Time \ 24
    
114     If Time = 1 Then
116         UpTimeStr = Time & " día, " & UpTimeStr
        Else
118         UpTimeStr = Time & " días, " & UpTimeStr

        End If
    
120     Call WriteConsoleMsg(UserIndex, "Server Online: " & UpTimeStr, e_FontTypeNames.FONTTYPE_INFO)
        
        Exit Sub

HandleUpTime_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUpTime", Erl)
124
        
End Sub

' Handles the "RoleMasterRequest" message.
Public Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim request As String
102             request = Reader.ReadString8()
        
104         If LenB(request) <> 0 Then
106             Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada", e_FontTypeNames.FONTTYPE_INFO)
108             Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.name & " PREGUNTA ROL: " & request, e_FontTypeNames.FONTTYPE_GUILDMSG))

            End If

        End With
    
        Exit Sub
        
ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRoleMasterRequest", Erl)
112

End Sub

Public Sub HandlePunishments(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim name As String
102             name = Reader.ReadString8()

            ' Si un GM usa este comando, me fijo que me haya dado el nick del PJ a analizar.
104         If LenB(name) = 0 Then Exit Sub

            If UserList(UserIndex).name <> name Then
                If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
        
            Dim count As Integer

106         If (InStrB(name, "\") <> 0) Then
108             name = Replace(name, "\", vbNullString)

            End If

110         If (InStrB(name, "/") <> 0) Then
112             name = Replace(name, "/", vbNullString)

            End If

114         If (InStrB(name, ":") <> 0) Then
116             name = Replace(name, ":", vbNullString)

            End If

118         If (InStrB(name, "|") <> 0) Then
120             name = Replace(name, "|", vbNullString)

            End If
           
            Dim TargetUserName As String

122         If EsGM(UserIndex) Then
        
124             If PersonajeExiste(name) Then
126                 TargetUserName = name
                
                Else
128                 Call WriteConsoleMsg(UserIndex, "El personaje " & TargetUserName & " no existe.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
            Else
        
130             TargetUserName = .name
            
            End If

134         count = GetUserAmountOfPunishmentsDatabase(TargetUserName)


138         If count = 0 Then
140             Call WriteConsoleMsg(UserIndex, "Sin prontuario..", e_FontTypeNames.FONTTYPE_INFO)
            Else
144             Call SendUserPunishmentsDatabase(UserIndex, TargetUserName)
            End If

        End With
        Exit Sub
ErrHandler:
152     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePunishments", Erl)
154

End Sub

Public Sub HandleGamble(ByVal UserIndex As Integer)
        
        On Error GoTo HandleGamble_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

            Dim amount As Integer
102             amount = Reader.ReadInt16()
        
104         If .flags.Muerto = 1 Then
106             Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                
108         ElseIf Not IsValidNpcRef(.flags.TargetNPC) Then
                'Validate target NPC
110             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar un personaje, haz click izquierdo sobre él.", e_FontTypeNames.FONTTYPE_INFO)

112         ElseIf Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 10 Then
114             Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                
116         ElseIf NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Timbero Then
118             Call WriteChatOverHead(UserIndex, "No tengo ningún interés en apostar.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)

120         ElseIf amount < 1 Then
122             Call WriteChatOverHead(UserIndex, "El mínimo de apuesta es 1 moneda.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)

124         ElseIf amount > 10000 Then
126             Call WriteChatOverHead(UserIndex, "El máximo de apuesta es 10.000 monedas.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)

128         ElseIf .Stats.GLD < amount Then
130             Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)

            Else
132             If RandomNumber(1, 100) <= 45 Then
134                 .Stats.GLD = .Stats.GLD + amount
136                 Call WriteChatOverHead(UserIndex, "¡Felicidades! Has ganado " & PonerPuntos(amount) & " monedas de oro!", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                
138                 Apuestas.Perdidas = Apuestas.Perdidas + amount
140                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
142                 .Stats.GLD = .Stats.GLD - amount
144                 Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & PonerPuntos(amount) & " monedas de oro.", NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite)
                
146                 Apuestas.Ganancias = Apuestas.Ganancias + amount
148                 Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

                End If
            
150             Apuestas.Jugadas = Apuestas.Jugadas + 1
152             Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
154             Call WriteUpdateGold(UserIndex)

            End If

        End With

        Exit Sub

HandleGamble_Err:
156     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGamble", Erl)
158
        
End Sub
 
Public Sub HandleDenounce(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)

            Dim name As String
102         name = Reader.ReadString8()

104         If LenB(name) = 0 Then Exit Sub

106         If EsGmChar(name) Then
108             Call WriteConsoleMsg(UserIndex, "No podés denunciar a un administrador.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            Dim tUser As t_UserReference
110         tUser = NameIndex(name)
        
112         If Not IsValidUserRef(tUser) Then 'user abuse of this system to know if someone is online
114             Call WriteConsoleMsg(UserIndex, "Tu denuncia fue recibida por el equipo de soporte.", e_FontTypeNames.FONTTYPE_INFOIAO)
                Exit Sub
            End If
        
            Dim Denuncia As String, HayChat As Boolean
116         Denuncia = "[Últimos mensajes de " & UserList(tUser.ArrayIndex).Name & "]" & vbNewLine
            Dim i As Integer

118         For i = 1 To UBound(UserList(tUser.ArrayIndex).flags.ChatHistory)

120             If LenB(UserList(tUser.ArrayIndex).flags.ChatHistory(i)) <> 0 Then
122                 Denuncia = Denuncia & UserList(tUser.ArrayIndex).flags.ChatHistory(i) & vbNewLine
124                 HayChat = True
                End If
            Next
        
126         If Not HayChat Then
128             Call WriteConsoleMsg(UserIndex, "El usuario no ha escrito nada. Recordá que las denuncias inválidas pueden ser motivo de advertencia.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

136         Call Ayuda.Push(.name, Denuncia, "Denuncia a " & UserList(tUser.ArrayIndex).name)
138         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido una nueva denuncia de parte de " & .name & ".", e_FontTypeNames.FONTTYPE_SERVER))
140         Call WriteConsoleMsg(UserIndex, "Tu denuncia fue recibida por el equipo de soporte.", e_FontTypeNames.FONTTYPE_INFOIAO)
142         Call LogConsulta(.name & " (Denuncia a " & UserList(tUser.ArrayIndex).name & ")" & vbNewLine & Denuncia)
        End With
        Exit Sub
ErrHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
End Sub

Public Sub HandleGMMessage(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim Message As String
102         Message = Reader.ReadString8()
104         If EsGM(UserIndex) Then
106             Call LogGM(.name, "Mensaje a Gms: " & Message)
108             If LenB(Message) <> 0 Then
112                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " » " & Message, e_FontTypeNames.FONTTYPE_GMMSG))
                End If

            End If

        End With

        Exit Sub
    
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGMMessage", Erl)
116

End Sub

Public Sub HandleShowName(ByVal UserIndex As Integer)
        
        On Error GoTo HandleShowName_Err

        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
        
104             .showName = Not .showName 'Show / Hide the name
            
106             Call RefreshCharStatus(UserIndex)

            End If

        End With
        
        Exit Sub

HandleShowName_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleShowName", Erl)
110
        
End Sub

Public Sub HandleGoNearby(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim username As String
102             username = Reader.ReadString8()
        
            Dim tUser As t_UserReference

            Dim X      As Long
            Dim y      As Long

            Dim i      As Long
            
            Dim Found  As Boolean
        
104         If Not EsGM(UserIndex) Then Exit Sub
        
            'Check the user has enough powers
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Or Ayuda.Existe(username) Then
108             tUser = NameIndex(username)
            
110             If Not IsValidUserRef(tUser) Then
                    ' Si está offline, comparamos privilegios offline, para no revelar si está el gm conectado
112                 If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(username)) >= 0 Then
114                     Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
116                     Call WriteConsoleMsg(UserIndex, "No podés ir cerca de un GM de mayor jerarquía.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
118                 If CompararPrivilegiosUser(userIndex, tUser.ArrayIndex) >= 0 Then
120                     For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
122                         For x = UserList(tUser.ArrayIndex).pos.x - i To UserList(tUser.ArrayIndex).pos.x + i
124                             For y = UserList(tUser.ArrayIndex).pos.y - i To UserList(tUser.ArrayIndex).pos.y + i
126                                 If MapData(UserList(tUser.ArrayIndex).pos.map, x, y).userIndex = 0 Then
128                                     If LegalPos(UserList(tUser.ArrayIndex).pos.map, x, y, True, True) Then
130                                         Call WriteConsoleMsg(userIndex, "Te teletransportaste cerca de " & UserList(tUser.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
132                                         Call WarpUserChar(userIndex, UserList(tUser.ArrayIndex).pos.map, x, y, True)
134                                         Found = True
                                            Exit For
                                        End If
                                    End If
136                             Next y
138                             If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
140                         Next X
142                         If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
144                     Next i
                        'No space found??
146                     If Not Found Then
148                         Call WriteConsoleMsg(UserIndex, "Todos los lugares están ocupados.", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    Else
150                     Call WriteConsoleMsg(UserIndex, "No podés ir cerca de un GM de mayor jerarquía.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
152             Call WriteConsoleMsg(UserIndex, "Servidor » No podés ir cerca de ningun Usuario si no pidio SOS.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
154     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGoNearby", Erl)
End Sub

Public Sub HandleWhere(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.RoleMaster Or e_PlayerType.Consejero Or e_PlayerType.user Or e_PlayerType.SemiDios)) = 0 Then
106             tUser = NameIndex(username)
108             If Not IsValidUserRef(tUser) Then
110                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)
                Else
112                 If CompararPrivilegiosUser(userIndex, tUser.ArrayIndex) >= 0 Then
114                     Call WriteConsoleMsg(UserIndex, "Ubicación  " & username & ": " & UserList(tUser.ArrayIndex).Pos.Map & ", " & UserList(tUser.ArrayIndex).Pos.X & ", " & UserList(tUser.ArrayIndex).Pos.Y & ".", e_FontTypeNames.FONTTYPE_INFO)
116                     Call LogGM(.name, "/Donde " & username)
                    End If
                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWhere", Erl)
End Sub

Public Sub HandleCreaturesInMap(ByVal UserIndex As Integer)
        
        On Error GoTo HandleCreaturesInMap_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

            Dim map As Integer
            Dim i, j As Long
            Dim NPCcount1, NPCcount2 As Integer
            Dim NPCcant1() As Integer
            Dim NPCcant2() As Integer
            Dim List1()    As String
            Dim List2()    As String
        
102         map = Reader.ReadInt16()
        
104         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
        
106         If MapaValido(map) Then

108             For i = 1 To LastNPC
                    'VB isn't lazzy, so we put more restrictive condition first to speed up the process
110                 If NpcList(i).Pos.map = map Then
                        'íesta vivo?
112                     If NpcList(i).flags.NPCActive And NpcList(i).Hostile = 1 Then
114                         If NPCcount1 = 0 Then
116                             ReDim List1(0) As String
118                             ReDim NPCcant1(0) As Integer
120                             NPCcount1 = 1
122                             List1(0) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.y & ")"
124                             NPCcant1(0) = 1
                            Else
126                             For j = 0 To NPCcount1 - 1

128                                 If Left$(List1(j), Len(NpcList(i).name)) = NpcList(i).name Then
130                                     List1(j) = List1(j) & ", (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.y & ")"
132                                     NPCcant1(j) = NPCcant1(j) + 1
                                        Exit For
                                    End If
134                             Next j

136                             If j = NPCcount1 Then
138                                 ReDim Preserve List1(0 To NPCcount1) As String
140                                 ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
142                                 NPCcount1 = NPCcount1 + 1
144                                 List1(j) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.y & ")"
146                                 NPCcant1(j) = 1
                                End If
                            End If
                        Else
148                         If NPCcount2 = 0 Then
150                             ReDim List2(0) As String
152                             ReDim NPCcant2(0) As Integer
154                             NPCcount2 = 1
156                             List2(0) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.y & ")"
158                             NPCcant2(0) = 1
                            Else
160                             For j = 0 To NPCcount2 - 1
162                                 If Left$(List2(j), Len(NpcList(i).name)) = NpcList(i).name Then
164                                     List2(j) = List2(j) & ", (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.y & ")"
166                                     NPCcant2(j) = NPCcant2(j) + 1
                                        Exit For
                                    End If
168                             Next j

170                             If j = NPCcount2 Then
172                                 ReDim Preserve List2(0 To NPCcount2) As String
174                                 ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
176                                 NPCcount2 = NPCcount2 + 1
178                                 List2(j) = NpcList(i).name & ": (" & NpcList(i).Pos.X & "," & NpcList(i).Pos.y & ")"
180                                 NPCcant2(j) = 1
                                End If
                            End If
                        End If
                    End If
182             Next i
184             Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en mapa: ", e_FontTypeNames.FONTTYPE_WARNING)
186             If NPCcount1 = 0 Then
188                 Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles", e_FontTypeNames.FONTTYPE_INFO)
                Else
190                 For j = 0 To NPCcount1 - 1
192                     Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), e_FontTypeNames.FONTTYPE_INFO)
194                 Next j
                End If
196             Call WriteConsoleMsg(UserIndex, "Otros Npcs en mapa: ", e_FontTypeNames.FONTTYPE_WARNING)
198             If NPCcount2 = 0 Then
200                 Call WriteConsoleMsg(UserIndex, "No hay más NPCS", e_FontTypeNames.FONTTYPE_INFO)
                Else
202                 For j = 0 To NPCcount2 - 1
204                     Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), e_FontTypeNames.FONTTYPE_INFO)
206                 Next j
                End If
208             Call LogGM(.name, "Numero enemigos en mapa " & map)
            End If
        End With
        Exit Sub
HandleCreaturesInMap_Err:
210     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreaturesInMap", Erl)
End Sub

Public Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)
        
        On Error GoTo HandleWarpMeToTarget_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
        
104         Call WarpUserChar(UserIndex, .flags.TargetMap, .flags.TargetX, .flags.TargetY, True)
        
106         Call LogGM(.name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.map)

        End With
        
        Exit Sub

HandleWarpMeToTarget_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarpMeToTarget", Erl)
110
        
End Sub

Public Sub HandleWarpChar(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim username As String
            Dim map      As Integer
            Dim X        As Byte
            Dim y        As Byte
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         map = Reader.ReadInt16()
106         X = Reader.ReadInt8()
108         y = Reader.ReadInt8()

110         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
            
112         If .flags.Privilegios And e_PlayerType.Consejero Then
        
114             If MapInfo(map).Seguro = 0 Then
116                 Call WriteConsoleMsg(UserIndex, "Solo puedes transportarte a ciudades.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                    'Si manda yo o su propio nombre
118             ElseIf LCase$(username) <> LCase$(UserList(UserIndex).name) And UCase$(username) <> "YO" Then
120                 Call WriteConsoleMsg(UserIndex, "Solo puedes transportarte a ti mismo.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
            
            '¿Para que te vas a transportar a la misma posicion?
122         If .Pos.map = map And .Pos.X = X And .Pos.y = y Then Exit Sub
            
124         If MapaValido(map) And LenB(username) <> 0 Then

126             If UCase$(username) <> "YO" Then
128                 tUser = NameIndex(username)
                Else
130                 Call SetUserRef(tUser, userIndex)
                End If
            
132             If Not IsValidUserRef(tUser) Then
134                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)

136             ElseIf InMapBounds(map, X, y) Then
                    'no permitimos que se use el telep para llevas User a casas privadas.
                    If UCase$(username) <> "YO" Then
                        If .flags.Privilegios And e_PlayerType.Consejero Or e_PlayerType.SemiDios Then
                            If Not EsMapaEvento(map) Then
                                Call WriteConsoleMsg(UserIndex, "Solamente puedes teletransportar gente a mapas de evento.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        End If
                        Call WarpToLegalPos(tUser.ArrayIndex, map, x, y, True, True)
                    Else
138                     Call FindLegalPos(tUser.ArrayIndex, map, x, y)
140                     Call WarpUserChar(tUser.ArrayIndex, map, x, y, True)
                    End If
142                 If tUser.ArrayIndex <> userIndex Then
144                     Call LogGM(.Name, "Transportó a " & UserList(tUser.ArrayIndex).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
                    End If
                End If
            End If
        End With
        Exit Sub
ErrHandler:
146     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarpChar", Erl)
148

End Sub

Public Sub HandleSilence(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim username As String
            Dim minutos  As Integer
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         minutos = Reader.ReadInt16()

106         If EsGM(UserIndex) Then
108             tUser = NameIndex(username)
        
110             If Not IsValidUserRef(tUser) Then

112                 If PersonajeExiste(username) Then

114                     If CompararPrivilegios(.flags.Privilegios, UserDarPrivilegioLevel(username)) > 0 Then

116                         If minutos > 0 Then
118                             Call SilenciarUserDatabase(username, minutos)
120                             Call SavePenaDatabase(username, .name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
122                             Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha silenciado a " & username & "(offline) por " & minutos & " minutos.", e_FontTypeNames.FONTTYPE_GM))
124                             Call LogGM(.name, "Silenciar a " & UserList(tUser.ArrayIndex).name & " por " & minutos & " minutos.")
                            Else
126                             Call DesilenciarUserDatabase(username)
128                             Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha desilenciado a " & username & "(offline).", e_FontTypeNames.FONTTYPE_GM))
130                             Call LogGM(.name, "Desilenciar a " & UserList(tUser.ArrayIndex).name & ".")

                            End If
                            
                        Else
                        
132                         Call WriteConsoleMsg(UserIndex, "No puedes silenciar a un administrador de mayor o igual rango.", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                    
134                     Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_INFO)

                    End If
                
136             ElseIf CompararPrivilegiosUser(userIndex, tUser.ArrayIndex) > 0 Then

138                 If minutos > 0 Then
140                     UserList(tUser.ArrayIndex).flags.Silenciado = 1
142                     UserList(tUser.ArrayIndex).flags.MinutosRestantes = minutos
144                     UserList(tUser.ArrayIndex).flags.SegundosPasados = 0

146                     Call SavePenaDatabase(username, .name & ": silencio por " & Time & " minutos. " & Date & " " & Time)
148                     Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha silenciado a " & UserList(tUser.ArrayIndex).Name & " por " & minutos & " minutos.", e_FontTypeNames.FONTTYPE_GM))
150                     Call WriteConsoleMsg(tUser.ArrayIndex, "Has sido silenciado por los administradores, no podrás hablar con otros usuarios. Utilice /GM para pedir ayuda.", e_FontTypeNames.FONTTYPE_GM)
152                     Call LogGM(.name, "Silenciar a " & UserList(tUser.ArrayIndex).name & " por " & minutos & " minutos.")

                    Else
                    
154                     UserList(tUser.ArrayIndex).flags.Silenciado = 1

156                     Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " ha desilenciado a " & UserList(tUser.ArrayIndex).Name & ".", e_FontTypeNames.FONTTYPE_GM))
158                     Call WriteConsoleMsg(tUser.ArrayIndex, "Has sido desilenciado.", e_FontTypeNames.FONTTYPE_GM)
160                     Call LogGM(.name, "Desilenciar a " & UserList(tUser.ArrayIndex).name & ".")

                    End If
                    
                Else
                
162                 Call WriteConsoleMsg(UserIndex, "No puedes silenciar a un administrador de mayor o igual rango.", e_FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End With
        
        Exit Sub
        
ErrHandler:
164     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSilence", Erl)
166

End Sub

Public Sub HandleSOSShowList(ByVal UserIndex As Integer)
        
        On Error GoTo HandleSOSShowList_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)

100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub

104         Call WriteShowSOSForm(UserIndex)

        End With
        
        Exit Sub

HandleSOSShowList_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSOSShowList", Erl)
108
        
End Sub


Public Sub HandleSOSRemove(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim username As String
102             username = Reader.ReadString8()
        
104         If Not .flags.Privilegios And e_PlayerType.user Then Call Ayuda.Quitar(username)

        End With
        
        Exit Sub
        
ErrHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSOSRemove", Erl)
108

End Sub

Public Sub HandleGoToChar(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim username As String
            Dim tUser    As t_UserReference
            Dim X        As Byte
            Dim y        As Byte
        
102         username = Reader.ReadString8()

104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             If LenB(username) <> 0 Then
108                 tUser = NameIndex(username)
                    
110                 If Not IsValidUserRef(tUser) Then
112                     Call WriteConsoleMsg(UserIndex, "El jugador no está online.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub

                    End If

                Else
114                 tUser = .flags.TargetUser

116                 If Not IsValidUserRef(tUser) Then Exit Sub

                End If
      
118             If CompararPrivilegiosUser(tUser.ArrayIndex, userIndex) > 0 Then
120                 Call WriteConsoleMsg(UserIndex, "Se le ha avisado a " & UserList(tUser.ArrayIndex).Name & " que quieres ir a su posición.", e_FontTypeNames.FONTTYPE_INFO)
122                 Call WriteConsoleMsg(tUser.ArrayIndex, .Name & " quiere transportarse a tu ubicación. Escribe /sum " & .Name & " para traerlo.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

124             x = UserList(tUser.ArrayIndex).pos.x
126             y = UserList(tUser.ArrayIndex).pos.y + 1
128             Call FindLegalPos(userIndex, UserList(tUser.ArrayIndex).pos.map, x, y)
130             Call WarpUserChar(userIndex, UserList(tUser.ArrayIndex).pos.map, x, y, True)
                    
132             If .flags.AdminInvisible = 0 Then
134                 Call WriteConsoleMsg(tUser.ArrayIndex, .name & " se ha trasportado hacia donde te encuentras.", e_FontTypeNames.FONTTYPE_INFO)
                End If
                
136             Call WriteConsoleMsg(userIndex, "Te has transportado hacia " & UserList(tUser.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
138             Call LogGM(.name, "/IRA " & username & " Mapa:" & UserList(tUser.ArrayIndex).pos.map & " X:" & UserList(tUser.ArrayIndex).pos.x & " Y:" & UserList(tUser.ArrayIndex).pos.y)
            Else
140             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo. solo puedes ir a Usuarios que piden SOS.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGoToChar", Erl)
144

End Sub

Public Sub HandleInvisible(ByVal UserIndex As Integer)
        On Error GoTo HandleInvisible_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)
102         If .flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero) Then Exit Sub
104         Call DoAdminInvisible(UserIndex)
        End With
        Exit Sub

HandleInvisible_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleInvisible", Erl)
108
        
End Sub

Public Sub HandleGMPanel(ByVal UserIndex As Integer)
        
        On Error GoTo HandleGMPanel_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

102         If .flags.Privilegios And e_PlayerType.user Then Exit Sub
        
104         Call WriteShowGMPanelForm(UserIndex)

        End With
        
        Exit Sub

HandleGMPanel_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGMPanel", Erl)
108
        
End Sub

Public Sub HandleRequestUserList(ByVal UserIndex As Integer)
        On Error GoTo HandleRequestUserList_Err

        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 01/09/07
        'Last modified by: Lucas Tavolaro Ortiz (Tavo)
        Dim i       As Long
        Dim names() As String
        Dim count   As Long
    
100     With UserList(UserIndex)
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         ReDim names(1 To LastUser) As String
108         count = 1
        
110         For i = 1 To LastUser

112             If (LenB(UserList(i).name) <> 0) Then
                
114                 names(count) = UserList(i).name
116                 count = count + 1
 
                End If

118         Next i
        
120         If count > 1 Then Call WriteUserNameList(UserIndex, names(), count - 1)

        End With
        
        Exit Sub

HandleRequestUserList_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestUserList", Erl)
124
        
End Sub

Public Sub HandleWorking(ByVal UserIndex As Integer)
        
        On Error GoTo HandleWorking_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        Dim i     As Long
        Dim Users As String
    
100     With UserList(UserIndex)

102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » /TRABAJANDO es un comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         For i = 1 To LastUser

108             If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
110                 Users = Users & ", " & UserList(i).name
                End If

114         Next i
        
116         If LenB(Users) <> 0 Then
118             Users = Right$(Users, Len(Users) - 2)
120             Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, e_FontTypeNames.FONTTYPE_INFO)
            Else
122             Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleWorking_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWorking", Erl)
126
        
End Sub

Public Sub HandleHiding(ByVal UserIndex As Integer)
        
        On Error GoTo HandleHiding_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        Dim i     As Long
        Dim Users As String
    
100     With UserList(UserIndex)

        
        
        
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         For i = 1 To LastUser

108             If (LenB(UserList(i).name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
110                 Users = Users & UserList(i).name & ", "

                End If

112         Next i
        
114         If LenB(Users) <> 0 Then
116             Users = Left$(Users, Len(Users) - 2)
118             Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, e_FontTypeNames.FONTTYPE_INFO)
            Else
120             Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End With
        
        Exit Sub

HandleHiding_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
124
        
End Sub

Public Sub HandleJail(ByVal UserIndex As Integer)

        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

        
        
        
            Dim username As String
            Dim Reason   As String
            Dim jailTime As Byte
            Dim count    As Byte
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         Reason = Reader.ReadString8()
106         jailTime = Reader.ReadInt8()
        
108         If InStr(1, username, "+") Then
110             username = Replace(username, "+", " ")

            End If
        
            '/carcel nick@motivo@<tiempo>
112         If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then

114             If LenB(username) = 0 Or LenB(Reason) = 0 Then
116                 Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", e_FontTypeNames.FONTTYPE_INFO)
                Else
118                 tUser = NameIndex(username)
                
120                 If Not IsValidUserRef(tUser) Then
122                     Call WriteConsoleMsg(UserIndex, "El usuario no está online.", e_FontTypeNames.FONTTYPE_INFO)
                    Else

124                     If EsGM(tUser.ArrayIndex) Then
126                         Call WriteConsoleMsg(UserIndex, "No podés encarcelar a administradores.", e_FontTypeNames.FONTTYPE_INFO)
                    
128                     ElseIf jailTime > 240 Then
130                         Call WriteConsoleMsg(UserIndex, "No podés encarcelar por más de 4 horas.", e_FontTypeNames.FONTTYPE_INFO)
                        Else

132                         If (InStrB(username, "\") <> 0) Then
134                             username = Replace(username, "\", "")

                            End If

136                         If (InStrB(username, "/") <> 0) Then
138                             username = Replace(username, "/", "")

                            End If
                        
140                         If PersonajeExiste(username) Then
144                                 Call SavePenaDatabase(username, .name & ": CARCEL " & jailTime & "m, MOTIVO: " & Reason & " " & Date & " " & Time)
                            End If
                        
152                         Call Encarcelar(tUser.ArrayIndex, jailTime, .name)
154                         Call LogGM(.Name, " encarceló a " & username)
                        End If
                    End If
                End If
            Else
156             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleHiding", Erl)
160

End Sub

Public Sub HandleKillNPC(ByVal UserIndex As Integer)
        
        On Error GoTo HandleKillNPC_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

102         If .flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero) Then
104             Call WriteConsoleMsg(UserIndex, "Solo los Administradores y Dioses pueden usar este comando.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
            End If
            
        
112         If IsValidNpcRef(.flags.TargetNPC) Then
                Dim tNPC As Integer
114             tNPC = .flags.TargetNPC.ArrayIndex
116             Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & NpcList(tNPC).name, e_FontTypeNames.FONTTYPE_INFO)
                Dim auxNPC As t_Npc
118             auxNPC = NpcList(tNPC)
120             Call QuitarNPC(tNPC, eGMCommand)
122             Call ReSpawnNpc(auxNPC)
            Else
124             Call WriteConsoleMsg(UserIndex, "Debes hacer click sobre el NPC antes", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        
        Exit Sub

HandleKillNPC_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillNPC", Erl)

128
        
End Sub

Public Sub HandleWarnUser(ByVal UserIndex As Integer)
    
        On Error GoTo ErrHandler

100     With UserList(UserIndex)

            Dim username As String
            Dim Reason   As String
102         username = Reader.ReadString8()
104         Reason = Reader.ReadString8()
        
            ' Tenes que ser Admin, Dios o Semi-Dios
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            ' Me fijo que esten todos los parametros.
110         If Len(username) = 0 Or Len(Trim$(Reason)) = 0 Then
112             Call WriteConsoleMsg(UserIndex, "Formato inválido. /advertencia nick@motivo", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim tUser As t_UserReference
114         tUser = NameIndex(username)
        
            ' No advertir a GM's
116         If EsGM(tUser.ArrayIndex) Then
118             Call WriteConsoleMsg(UserIndex, "No podes advertir a Game Masters.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
120         If (InStrB(username, "\") <> 0) Then
122             username = Replace(username, "\", "")
            End If

124         If (InStrB(username, "/") <> 0) Then
126             username = Replace(username, "/", "")

            End If
                    
128         If PersonajeExiste(username) Then
132             Call SaveWarnDatabase(username, "ADVERTENCIA: " & Reason & " " & Date & " " & Time, .name)

            
                ' Para el GM
140             Call WriteConsoleMsg(UserIndex, "Has advertido a " & username, e_FontTypeNames.FONTTYPE_CENTINELA)
142             Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " ha advertido a " & username & " por " & Reason, e_FontTypeNames.FONTTYPE_GM))
144             Call LogGM(.name, " advirtio a " & username & " por " & Reason)

                ' Si esta online...
146             If IsValidUserRef(tUser) Then
                    ' Actualizo el valor en la memoria.
148                 UserList(tUser.ArrayIndex).Stats.Advertencias = UserList(tUser.ArrayIndex).Stats.Advertencias + 1
                
                    ' Para el usuario advertido
150                 Call WriteConsoleMsg(tUser.ArrayIndex, "Has sido advertido por " & Reason, e_FontTypeNames.FONTTYPE_CENTINELA)
152                 Call WriteConsoleMsg(tUser.ArrayIndex, "Tenés " & UserList(tUser.ArrayIndex).Stats.Advertencias & " advertencias actualmente.", e_FontTypeNames.FONTTYPE_CENTINELA)
                    ' Cuando acumulas cierta cantidad de advertencias...
154                 Select Case UserList(tUser.ArrayIndex).Stats.Advertencias
                        Case 3
156                         Call Encarcelar(tUser.ArrayIndex, 30, "Servidor")
158                     Case 5
                            ' TODO: Banear PJ alv.
                    End Select
                End If
            End If
        
        End With
        Exit Sub
    
ErrHandler:

160     Call TraceError(Err.Number, Err.Description, "Protocol.HandleWarnUser", Erl)
End Sub

Public Sub HandleEditChar(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username      As String
            Dim tUser         As t_UserReference
            Dim opcion        As Byte
            Dim Arg1          As String
            Dim Arg2          As String
            Dim valido        As Boolean
            Dim LoopC         As Byte
            Dim commandString As String
            Dim n             As Byte
            Dim tmpLong       As Long
102         username = Replace(Reader.ReadString8(), "+", " ")
        
104         If UCase$(username) = "YO" Then
106             Call SetUserRef(tUser, userIndex)
            Else
108             tUser = NameIndex(username)
            End If
        
110         opcion = Reader.ReadInt8()
112         Arg1 = Reader.ReadString8()
114         Arg2 = Reader.ReadString8()

            ' Si no es GM, no hacemos nada.
116         If Not EsGM(UserIndex) Then Exit Sub

' From now on we disabled character edition in production environment
' This should have never been enabled in the first place
' If we require similar functionality for adhoc quests we will have to
' create new in-game tools focusing in TRANSPARENCY to prevent future abuse
#If DEBUGGING = 0 Then
            Exit Sub
#End If
            ' Si NO sos Dios o Admin,
118         If (.flags.Privilegios And e_PlayerType.Admin) = 0 Then
                ' Si te editas a vos mismo esta bien ;)
120             If userIndex <> tUser.ArrayIndex Then Exit Sub
            End If
            If Not IsValidUserRef(tUser) Then
122             Call WriteConsoleMsg(userIndex, "Usuario offline: " & username, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
124         Select Case opcion

                Case e_EditOptions.eo_Gold
126
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

130                 UserList(tUser.ArrayIndex).Stats.GLD = val(Arg1)
132                 Call WriteUpdateGold(tUser.ArrayIndex)
134             Case e_EditOptions.eo_Experience
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
140                 If UserList(tUser.ArrayIndex).Stats.ELV < STAT_MAXELV Then
142                     UserList(tUser.ArrayIndex).Stats.Exp = UserList(tUser.ArrayIndex).Stats.Exp + val(Arg1)
144                     Call CheckUserLevel(tUser.ArrayIndex)
146                     Call WriteUpdateExp(tUser.ArrayIndex)
                    Else
148                     Call WriteConsoleMsg(UserIndex, "El usuario es nivel máximo.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
150             Case e_EditOptions.eo_Body

152                 If Not IsValidUserRef(tUser) Then
156                     Call SaveUserBodyDatabase(username, val(Arg1))
160                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & username, e_FontTypeNames.FONTTYPE_INFO)
                    Else
162                     Call ChangeUserChar(tUser.ArrayIndex, val(Arg1), UserList(tUser.ArrayIndex).char.head, UserList(tUser.ArrayIndex).char.Heading, UserList(tUser.ArrayIndex).char.WeaponAnim, UserList(tUser.ArrayIndex).char.ShieldAnim, UserList(tUser.ArrayIndex).char.CascoAnim, UserList(UserIndex).char.CartAnim)
                    End If
                   
164             Case e_EditOptions.eo_Arma
172                 Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).char.body, UserList(tUser.ArrayIndex).char.head, UserList(tUser.ArrayIndex).char.Heading, val(Arg1), UserList(tUser.ArrayIndex).char.ShieldAnim, UserList(tUser.ArrayIndex).char.CascoAnim, UserList(UserIndex).char.CartAnim)
174             Case e_EditOptions.eo_Escudo
182                 Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).char.body, UserList(tUser.ArrayIndex).char.head, UserList(tUser.ArrayIndex).char.Heading, UserList(tUser.ArrayIndex).char.WeaponAnim, val(Arg1), UserList(tUser.ArrayIndex).char.CascoAnim, UserList(UserIndex).char.CartAnim)
184             Case e_EditOptions.eo_CASCO
192                     Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).char.body, UserList(tUser.ArrayIndex).char.head, UserList(tUser.ArrayIndex).char.Heading, UserList(tUser.ArrayIndex).char.WeaponAnim, UserList(tUser.ArrayIndex).char.ShieldAnim, val(Arg1), UserList(UserIndex).char.CartAnim)
194             Case e_EditOptions.eo_Particula
196                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
198                 If Not .flags.Privilegios = Consejero Then
206                     Call SendData(SendTarget.ToPCAliveArea, userIndex, PrepareMessageParticleFX(.Char.charindex, val(Arg1), 9999, False))
208                     .Char.ParticulaFx = val(Arg1)
210                     .Char.loops = 9999
                    End If
212             Case e_EditOptions.eo_Head
214                 If tUser.ArrayIndex <= 0 Then
218                     Call SaveUserHeadDatabase(username, val(Arg1))
222                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & username, e_FontTypeNames.FONTTYPE_INFO)
                    Else
224                     Call ChangeUserChar(tUser.ArrayIndex, UserList(tUser.ArrayIndex).char.body, val(Arg1), UserList(tUser.ArrayIndex).char.Heading, UserList(tUser.ArrayIndex).char.WeaponAnim, UserList(tUser.ArrayIndex).char.ShieldAnim, UserList(tUser.ArrayIndex).char.CascoAnim, UserList(UserIndex).char.CartAnim)
                    End If
226             Case e_EditOptions.eo_CriminalsKilled
228                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
234                 If val(Arg1) > MAXUSERMATADOS Then
236                     UserList(tUser.ArrayIndex).Faccion.CriminalesMatados = MAXUSERMATADOS
                    Else
238                     UserList(tUser.ArrayIndex).Faccion.CriminalesMatados = val(Arg1)
                    End If
240             Case e_EditOptions.eo_CiticensKilled
242                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
248                 If val(Arg1) > MAXUSERMATADOS Then
250                     UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados = MAXUSERMATADOS
                    Else
252                     UserList(tUser.ArrayIndex).Faccion.ciudadanosMatados = val(Arg1)
                    End If
254             Case e_EditOptions.eo_Level
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then Exit Sub
260                 If val(Arg1) > STAT_MAXELV Then
262                     Arg1 = CStr(STAT_MAXELV)
264                     Call WriteConsoleMsg(UserIndex, "No podés tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)
                    End If
266                     UserList(tUser.ArrayIndex).Stats.ELV = val(Arg1)
268                 Call WriteUpdateUserStats(UserIndex)
                
270             Case e_EditOptions.eo_Class
276                 For LoopC = 1 To NUMCLASES
278                     If Tilde(ListaClases(LoopC)) = Tilde(Arg1) Then Exit For
280                 Next LoopC
                        
282                 If LoopC > NUMCLASES Then
284                     Call WriteConsoleMsg(userIndex, "Clase desconocida. Intente nuevamente.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
286                     UserList(tUser.ArrayIndex).clase = LoopC
                    End If
                
288             Case e_EditOptions.eo_Skills
                    
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then Exit Sub
                    
290                 For LoopC = 1 To NUMSKILLS
292                     If Tilde(Replace$(SkillsNames(LoopC), " ", "+")) = Tilde(Arg1) Then Exit For
294                 Next LoopC
                    
296                 If LoopC > NUMSKILLS Then
298                     Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", e_FontTypeNames.FONTTYPE_INFO)
                    Else
300                     If Not IsValidUserRef(tUser) Then
304                         Call SaveUserSkillDatabase(username, LoopC, val(Arg2))
308                         Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & username, e_FontTypeNames.FONTTYPE_INFO)
                        Else
310                         UserList(tUser.ArrayIndex).Stats.UserSkills(LoopC) = val(Arg2)
                        End If
                    End If
                
312             Case e_EditOptions.eo_SkillPointsLeft
314                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then Exit Sub
316                 If Not IsValidUserRef(tUser) Then
320                     Call SaveUserSkillsLibres(username, val(Arg1))
324                     Call WriteConsoleMsg(UserIndex, "Usuario Offline Alterado: " & username, e_FontTypeNames.FONTTYPE_INFO)
                    Else
326                     UserList(tUser.ArrayIndex).Stats.SkillPts = val(Arg1)
                    End If
                
328             Case e_EditOptions.eo_Sex

330                 If Not IsValidUserRef(tUser) Then
332                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & username, e_FontTypeNames.FONTTYPE_INFO)
                    Else
334                     Arg1 = UCase$(Arg1)
336                     If (Arg1 = "MUJER") Then
338                         UserList(tUser.ArrayIndex).genero = e_Genero.Mujer
340                     ElseIf (Arg1 = "HOMBRE") Then
342                         UserList(tUser.ArrayIndex).genero = e_Genero.Hombre
                        End If
                    End If
                
344             Case e_EditOptions.eo_Raza
346                 If Not IsValidUserRef(tUser) Then
348                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & username, e_FontTypeNames.FONTTYPE_INFO)
                    Else
350                     Arg1 = UCase$(Arg1)
352                     If (Arg1 = "HUMANO") Then
354                         UserList(tUser.ArrayIndex).raza = e_Raza.Humano
356                     ElseIf (Arg1 = "ELFO") Then
358                         UserList(tUser.ArrayIndex).raza = e_Raza.Elfo
360                     ElseIf (Arg1 = "DROW") Then
362                         UserList(tUser.ArrayIndex).raza = e_Raza.Drow
364                     ElseIf (Arg1 = "ENANO") Then
366                         UserList(tUser.ArrayIndex).raza = e_Raza.Enano
368                     ElseIf (Arg1 = "GNOMO") Then
370                         UserList(tUser.ArrayIndex).raza = e_Raza.Gnomo
372                     ElseIf (Arg1 = "ORCO") Then
374                         UserList(tUser.ArrayIndex).raza = e_Raza.Orco
                        End If
                    End If
376             Case e_EditOptions.eo_Vida
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
378                 If Not IsValidUserRef(tUser) Then
380                     Call WriteConsoleMsg(UserIndex, "Usuario offline: " & username, e_FontTypeNames.FONTTYPE_INFO)
                    Else
382                     tmpLong = val(Arg1)
384                     If tmpLong > 0 Then
386                         UserList(tUser.ArrayIndex).Stats.MaxHp = Min(tmpLong, STAT_MAXHP)
388                         UserList(tUser.ArrayIndex).Stats.MinHp = UserList(tUser.ArrayIndex).Stats.MaxHp
390                         Call WriteUpdateUserStats(tUser.ArrayIndex)
                        End If
                    End If
                    
392             Case e_EditOptions.eo_Mana
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
398                     tmpLong = val(Arg1)
400                     If tmpLong > 0 Then
402                         UserList(tUser.ArrayIndex).Stats.MaxMAN = Min(tmpLong, STAT_MAXMP)
404                         UserList(tUser.ArrayIndex).Stats.MinMAN = UserList(tUser.ArrayIndex).Stats.MaxMAN
406                         Call WriteUpdateUserStats(tUser.ArrayIndex)
                        End If
                    
408             Case e_EditOptions.eo_Energia
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
414                     tmpLong = val(Arg1)
416                 If tmpLong > 0 Then
418                     UserList(tUser.ArrayIndex).Stats.MaxSta = Min(tmpLong, STAT_MAXSTA)
420                     UserList(tUser.ArrayIndex).Stats.MinSta = UserList(tUser.ArrayIndex).Stats.MaxSta
422                     Call WriteUpdateUserStats(tUser.ArrayIndex)
                    End If
                        
424             Case e_EditOptions.eo_MinHP
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
430                     tmpLong = val(Arg1)
432                 If tmpLong >= 0 Then
434                     UserList(tUser.ArrayIndex).Stats.MinHp = Min(tmpLong, STAT_MAXHP)
436                     Call WriteUpdateHP(tUser.ArrayIndex)
                    End If
                    
438             Case e_EditOptions.eo_MinMP
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
444                     tmpLong = val(Arg1)
446                 If tmpLong >= 0 Then
448                     UserList(tUser.ArrayIndex).Stats.MinMAN = Min(tmpLong, STAT_MAXMP)
450                     Call WriteUpdateMana(tUser.ArrayIndex)
                    End If
                    
452             Case e_EditOptions.eo_Hit
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
458                     tmpLong = val(Arg1)
460                 If tmpLong >= 0 Then
462                     UserList(tUser.ArrayIndex).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)
464                     UserList(tUser.ArrayIndex).Stats.MinHIT = UserList(tUser.ArrayIndex).Stats.MaxHit
                    End If

466             Case e_EditOptions.eo_MinHit
468                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
472                 tmpLong = val(Arg1)
474                 If tmpLong >= 0 Then
476                     UserList(tUser.ArrayIndex).Stats.MinHIT = Min(tmpLong, STAT_MAXHIT)
                    End If
                    
478             Case e_EditOptions.eo_MaxHit
                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
484                 tmpLong = val(Arg1)
486                 If tmpLong >= 0 Then
488                     UserList(tUser.ArrayIndex).Stats.MaxHit = Min(tmpLong, STAT_MAXHIT)
                    End If
                    
490             Case e_EditOptions.eo_Desc
                    If DescripcionValida(Arg1) Then
498                     UserList(tUser.ArrayIndex).Desc = Arg1
                    Else
500                     Call WriteConsoleMsg(UserIndex, "Caracteres inválidos en la descripción.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                    
502             Case e_EditOptions.eo_Intervalo
504                 If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub
510                     Arg1 = UCase$(Arg1)
512                     tmpLong = val(Arg2)
514                     If tmpLong >= 0 Then
516                         Select Case Arg1
                            Case "USAR"
518                             UserList(tUser.ArrayIndex).Intervals.UsarClic = tmpLong
520                             UserList(tUser.ArrayIndex).Intervals.UsarU = tmpLong
                                    
522                         Case "USAR_U", "USAR+U", "USAR-U"
524                             UserList(tUser.ArrayIndex).Intervals.UsarU = tmpLong
                                    
526                         Case "USAR_CLIC", "USAR+CLIC", "USAR-CLIC", "USAR_CLICK", "USAR+CLICK", "USAR-CLICK"
528                             UserList(tUser.ArrayIndex).Intervals.UsarClic = tmpLong
                                    
530                         Case "ARCO", "PROYECTILES"
532                             UserList(tUser.ArrayIndex).Intervals.Arco = tmpLong
                                
534                         Case "GOLPE", "GOLPES", "GOLPEAR"
536                             UserList(tUser.ArrayIndex).Intervals.Golpe = tmpLong
                                    
538                         Case "MAGIA", "HECHIZO", "HECHIZOS", "LANZAR"
540                             UserList(tUser.ArrayIndex).Intervals.Magia = tmpLong

542                         Case "COMBO"
544                             UserList(tUser.ArrayIndex).Intervals.GolpeMagia = tmpLong
546                             UserList(tUser.ArrayIndex).Intervals.MagiaGolpe = tmpLong

548                         Case "GOLPE-MAGIA", "GOLPE-HECHIZO"
550                             UserList(tUser.ArrayIndex).Intervals.GolpeMagia = tmpLong

552                         Case "MAGIA-GOLPE", "HECHIZO-GOLPE"
554                             UserList(tUser.ArrayIndex).Intervals.MagiaGolpe = tmpLong
                                
556                         Case "GOLPE-USAR"
558                             UserList(tUser.ArrayIndex).Intervals.GolpeUsar = tmpLong
                                    
560                         Case "TRABAJAR", "WORK", "TRABAJO"
562                             UserList(tUser.ArrayIndex).Intervals.TrabajarConstruir = tmpLong
564                             UserList(tUser.ArrayIndex).Intervals.TrabajarExtraer = tmpLong
                                
566                         Case "TRABAJAR_EXTRAER", "EXTRAER", "TRABAJO_EXTRAER"
568                             UserList(tUser.ArrayIndex).Intervals.TrabajarExtraer = tmpLong
                                
570                         Case "TRABAJAR_CONSTRUIR", "CONSTRUIR", "TRABAJO_CONSTRUIR"
572                             UserList(tUser.ArrayIndex).Intervals.TrabajarConstruir = tmpLong
                                
574                         Case Else
                                Exit Sub
                        End Select
576                     Call WriteIntervals(tUser.ArrayIndex)
                    End If
                    
578             Case e_EditOptions.eo_Hogar
584                 Arg1 = UCase$(Arg1)
586                 Select Case Arg1
                        Case "NIX"
588                         UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cNix
590                     Case "ULLA", "ULLATHORPE"
592                         UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cUllathorpe
594                     Case "BANDER", "BANDERBILL"
596                         UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cBanderbill
598                     Case "LINDOS"
600                           UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cLindos
602                     Case "ARGHAL"
604                         UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cArghal
606                     Case "ARKHEIN"
608                         UserList(tUser.ArrayIndex).Hogar = e_Ciudad.cArkhein
                    End Select

610             Case Else
612                 Call WriteConsoleMsg(UserIndex, "Comando no permitido.", e_FontTypeNames.FONTTYPE_INFO)
            End Select
            'Log it!
614         commandString = "/MOD "
        
616         Select Case opcion
                Case e_EditOptions.eo_Gold
618                 commandString = commandString & "ORO "
            
620             Case e_EditOptions.eo_Experience
622                 commandString = commandString & "EXP "
            
624             Case e_EditOptions.eo_Body
626                 commandString = vbNullString
            
628             Case e_EditOptions.eo_Head
630                 commandString = vbNullString
            
632             Case e_EditOptions.eo_CriminalsKilled
634                 commandString = commandString & "CRI "
            
636             Case e_EditOptions.eo_CiticensKilled
638                 commandString = commandString & "CIU "
            
640             Case e_EditOptions.eo_Level
642                 commandString = commandString & "LEVEL "
            
644             Case e_EditOptions.eo_Class
646                 commandString = commandString & "CLASE "
            
648             Case e_EditOptions.eo_Skills
650                 commandString = commandString & "SKILLS "
            
652             Case e_EditOptions.eo_SkillPointsLeft
654                 commandString = commandString & "SKILLSLIBRES "
                
656             Case e_EditOptions.eo_Sex
658                 commandString = commandString & "SEX "
                
660             Case e_EditOptions.eo_Raza
662                 commandString = commandString & "RAZA "

664             Case e_EditOptions.eo_Vida
666                 commandString = commandString & "VIDA "
                    
668             Case e_EditOptions.eo_Mana
670                 commandString = commandString & "MANA "
                    
672             Case e_EditOptions.eo_Energia
674                 commandString = commandString & "ENERGIA "
                    
676             Case e_EditOptions.eo_MinHP
678                 commandString = commandString & "MINHP "
                    
680             Case e_EditOptions.eo_MinMP
682                 commandString = commandString & "MINMP "
                    
684             Case e_EditOptions.eo_Hit
686                 commandString = commandString & "HIT "
                    
688             Case e_EditOptions.eo_MinHit
690                 commandString = commandString & "MINHIT "
                    
692             Case e_EditOptions.eo_MaxHit
694                 commandString = commandString & "MAXHIT "
                    
696             Case e_EditOptions.eo_Desc
698                 commandString = commandString & "DESC "
                    
700             Case e_EditOptions.eo_Intervalo
702                 commandString = commandString & "INTERVALO "
                    
704             Case e_EditOptions.eo_Hogar
706                 commandString = commandString & "HOGAR "
                
                Case e_EditOptions.eo_CASCO
                    commandString = vbNullString
                   
                Case e_EditOptions.eo_Arma
                    commandString = vbNullString
                    
                Case e_EditOptions.eo_Escudo
                    commandString = vbNullString

708             Case Else
710                 commandString = commandString & "UNKOWN "

            End Select
            
            If commandString <> vbNullString Then
714             Call LogGM(.name, commandString & Arg1 & " " & Arg2 & " " & username)
            End If
            
        End With

        Exit Sub

ErrHandler:
716     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEditChar", Erl)
End Sub

Public Sub HandleRequestCharInfo(ByVal UserIndex As Integer)

        'Author: Fredy Horacio Treboux (liquid)
        'Last Modification: 01/08/07
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim targetName  As String
            Dim tUser As t_UserReference
        
102         targetName = Replace$(Reader.ReadString8(), "+", " ")
104         tUser = NameIndex(targetName)
        
106         If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
                'is the player offline?
108             If IsValidUserRef(tUser) Then
                    'don't allow to retrieve administrator's info
116                 If UserList(tUser.ArrayIndex).flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
118                     Call SendUserStatsTxt(userIndex, tUser.ArrayIndex)
                    End If
                End If
            Else
120             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharInfo", Erl)
End Sub

' Handles the "RequestCharStats" message.
Public Sub HandleRequestCharStats(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference

102         username = Reader.ReadString8()
104         If (Not .flags.Privilegios And e_PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) <> 0 Then
106             Call LogGM(.name, "/STAT " & username)
108             tUser = NameIndex(username)
110             If IsValidUserRef(tUser) Then
116                 Call SendUserMiniStatsTxt(userIndex, tUser.ArrayIndex)
                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharStats", Erl)
End Sub

' Handles the "RequestCharGold" message.
Public Sub HandleRequestCharGold(ByVal UserIndex As Integer)

        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
        
            Dim username As String
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         tUser = NameIndex(username)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.name, "/BAL " & username)
            
110             If IsValidUserRef(tUser) Then
116                 Call WriteConsoleMsg(userIndex, "El usuario " & username & " tiene " & UserList(tUser.ArrayIndex).Stats.Banco & " en el banco", e_FontTypeNames.FONTTYPE_TALK)
                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub

ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharGold", Erl)
End Sub

' Handles the "RequestCharInventory" message.
Public Sub HandleRequestCharInventory(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
        
            Dim username As String
            Dim tUser    As t_UserReference
102         username = Reader.ReadString8()
104         tUser = NameIndex(username)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.name, "/INV " & username)
110             If IsValidUserRef(tUser) Then
116                 Call SendUserInvTxt(userIndex, tUser.ArrayIndex)
                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If

        End With
        Exit Sub

ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharInventory", Erl)
End Sub

Public Sub HandleRequestCharBank(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
102         username = Reader.ReadString8()
104         tUser = NameIndex(username)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.name, "/BOV " & username)
110             If Not IsValidUserRef(tUser) Then
112                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_TALK)
                Else
116                 Call SendUserBovedaTxt(userIndex, tUser.ArrayIndex)
                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)

            End If
        End With
        Exit Sub

ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharBank", Erl)
End Sub

Public Sub HandleRequestCharSkills(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
            Dim LoopC    As Long
            Dim Message  As String
        
102         username = Reader.ReadString8()
104         tUser = NameIndex(username)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call LogGM(.name, "/STATS " & username)
            
110             If Not IsValidUserRef(tUser) Then
112                 If (InStrB(username, "\") <> 0) Then
114                     username = Replace(username, "\", "")
                    End If

116                 If (InStrB(username, "/") <> 0) Then
118                     username = Replace(username, "/", "")
                    End If
                Else
128                 Call SendUserSkillsTxt(userIndex, tUser.ArrayIndex)
                End If
            Else
130             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
132     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRequestCharSkills", Erl)
End Sub

Public Sub HandleReviveChar(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
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
            
112             If Not IsValidUserRef(tUser) Then
114                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)
                Else
116                 With UserList(tUser.ArrayIndex)
118                     If .flags.Muerto = 1 Then
                            If UserList(UserIndex).flags.Privilegios And e_PlayerType.SemiDios Then
                                If MapInfo(.Pos.Map).Seguro = 0 Or Zona(.ZonaId).Segura = 0 Or EsMapaEvento(.Pos.Map) = False Then
                                     Call WriteConsoleMsg(UserIndex, "Servidor » No puedes revivir en una zona insegura.", e_FontTypeNames.FONTTYPE_INFO)
                                     Exit Sub
                                End If
                            End If
120                         .flags.Muerto = 0
122                         Call RevivirUsuario(tUser.ArrayIndex)
124                         Call WriteConsoleMsg(tUser.ArrayIndex, UserList(userIndex).name & " te ha resucitado.", e_FontTypeNames.FONTTYPE_INFO)
                        Else
126                         Call WriteConsoleMsg(tUser.ArrayIndex, UserList(userIndex).name & " te ha curado.", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    
128                     .Stats.MinHp = .Stats.MaxHp
                        .Stats.MinMAN = .Stats.MaxMAN
                        .Stats.MinSta = .Stats.MaxSta
                    End With
                
130                 Call WriteUpdateHP(tUser.ArrayIndex)
131                 Call WriteUpdateMana(tUser.ArrayIndex)
133                 Call WriteUpdateSta(tUser.ArrayIndex)
132                 Call ActualizarVelocidadDeUsuario(tUser.ArrayIndex)
134                 Call LogGM(.name, "Resucito a " & username)
                End If
            Else
136             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
End Sub

Public Sub HandleOnlineGM(ByVal UserIndex As Integer)
        On Error GoTo HandleOnlineGM_Err
        'Author: Fredy Horacio Treboux (liquid)
        Dim i    As Long
        Dim list As String
        Dim priv As e_PlayerType
    
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         priv = e_PlayerType.Consejero Or e_PlayerType.SemiDios
108         If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then priv = priv Or e_PlayerType.Dios Or e_PlayerType.Admin
      
110         For i = 1 To LastUser
112             If UserList(i).flags.UserLogged Then
114                 If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).name & ", "
                End If
116         Next i
        
118         If LenB(list) <> 0 Then
120             list = Left$(list, Len(list) - 2)
122             Call WriteConsoleMsg(UserIndex, list & ".", e_FontTypeNames.FONTTYPE_INFO)
            Else
124             Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub

HandleOnlineGM_Err:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineGM", Erl)
End Sub

Public Sub HandleOnlineMap(ByVal UserIndex As Integer)
        On Error GoTo HandleOnlineMap_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            Dim LoopC As Long
            Dim list  As String
            Dim priv  As e_PlayerType
        
106         priv = e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios
108         If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then priv = priv + (e_PlayerType.Dios Or e_PlayerType.Admin)
        
110         For LoopC = 1 To LastUser
112             If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).Pos.map = .Pos.map Then
114                 If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).name & ", "
                End If
116         Next LoopC
        
118         If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
120         Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, e_FontTypeNames.FONTTYPE_INFO)
        End With
        Exit Sub

HandleOnlineMap_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleOnlineMap", Erl)
End Sub

Public Sub HandleForgive(ByVal UserIndex As Integer)
        On Error GoTo HandleForgive_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            'Se asegura que el target es un npc
102         If Not IsValidNpcRef(.flags.TargetNPC) Then
104             Call WriteConsoleMsg(UserIndex, "Primero tenés que seleccionar al sacerdote.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            'Validate NPC and make sure player is not dead
106         If (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.Revividor And (NpcList(.flags.TargetNPC.ArrayIndex).npcType <> e_NPCType.ResucitadorNewbie Or Not EsNewbie(UserIndex))) Or .flags.Muerto = 1 Then Exit Sub
        
            Dim priest As t_Npc
108         priest = NpcList(.flags.TargetNPC.ArrayIndex)

            'Make sure it's close enough
110         If Distancia(.Pos, priest.Pos) > 3 Then
                'Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
112             Call WriteConsoleMsg(UserIndex, "El sacerdote no puede escuchar tus pecados debido a que estás demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
114         If .Faccion.Status = e_Facciones.Ciudadano Or .Faccion.Status = e_Facciones.Armada Or .Faccion.Status = e_Facciones.consejo Then
                'Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
116             Call WriteChatOverHead(UserIndex, "Tu alma ya esta libre de pecados hijo mio.", priest.Char.charindex, vbWhite)
                Exit Sub
            End If
        
118         If .Faccion.Status = e_Facciones.Caos Or .Faccion.Status = e_Facciones.consejo Then
120             Call WriteChatOverHead(UserIndex, "¡¡Dios no te perdonará mientras seas fiel al Demonio!!", priest.Char.charindex, vbWhite)
                Exit Sub
            End If

122         If .GuildIndex <> 0 Then
124             If modGuilds.Alineacion(.GuildIndex) = 1 Then
126                 Call WriteChatOverHead(UserIndex, "Te encuentras en un clan criminal... debes retirarte para que pueda perdonarte.", priest.Char.charindex, vbWhite)
                    Exit Sub
                End If
            End If

128         If .Faccion.ciudadanosMatados > 0 Then
                Dim Donacion As Long
130             Donacion = .Faccion.ciudadanosMatados * OroMult * CostoPerdonPorCiudadano
132             Call WriteChatOverHead(UserIndex, "Has matado a ciudadanos inocentes, Dios no puede perdonarte lo que has hecho. " & "Pero si haces una generosa donación de, digamos, " & PonerPuntos(Donacion) & " monedas de oro, tal vez cambie de opinión...", priest.Char.charindex, vbWhite)
                Exit Sub
            Else
                Call WriteChatOverHead(UserIndex, "Para volver a ser un ciudadano deberás donar 10.000 monedas de oro.", priest.Char.charindex, vbWhite)
            End If
                        
            Dim permitePerdon As Boolean
            permitePerdon = False
            
            If .GuildIndex > 0 And (GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CAOTICA Or GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CRIMINAL) Then
                permitePerdon = False
            Else
                permitePerdon = True
            End If
            
            If Not permitePerdon Then
                Call WriteChatOverHead(UserIndex, "No podrás ser perdonado perteneciendo a un clan de alineación Criminal o de Alineación Oscura.", priest.Char.charindex, vbYellow)
                Exit Sub
            End If

134         Call WriteChatOverHead(UserIndex, "Con estas palabras, te libero de todo tipo de pecados. ¡Que Dios te acompañe hijo mío!", priest.Char.charindex, vbYellow)
136         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, "80", 100, False))
138         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave("100", UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
140         Call VolverCiudadano(UserIndex)
        End With
        Exit Sub

HandleForgive_Err:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForgive", Erl)
End Sub

Public Sub HandleKick(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
            Dim rank     As Integer
        
102         rank = e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero
104         username = Reader.ReadString8()
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             tUser = NameIndex(username)
110             If Not IsValidUserRef(tUser) Then
112                 Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", e_FontTypeNames.FONTTYPE_INFO)
                Else
114                 If (UserList(tUser.ArrayIndex).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
116                     Call WriteConsoleMsg(UserIndex, "No podes echar a alguien con jerarquia mayor a la tuya.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
118                     Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " echo a " & username & ".", e_FontTypeNames.FONTTYPE_INFO))
120                     Call CloseSocket(tUser.ArrayIndex)
122                     Call LogGM(.name, "Echo a " & username)
                    End If
                End If
            Else
124             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKick", Erl)

End Sub

Public Sub HandleExecute(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)

            Dim username As String
            Dim tUser    As t_UserReference
102         username = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             tUser = NameIndex(username)
108             If IsValidUserRef(tUser) Then
                    Call CustomScenarios.UserDie(UserIndex)
110                 Call UserMod.UserDie(tUser.ArrayIndex)
112                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.name & " ha ejecutado a " & UserList(tUser.ArrayIndex).name, e_FontTypeNames.FONTTYPE_EJECUCION))
114                 Call LogGM(.name, " ejecuto a " & username)
                Else
116                 Call WriteConsoleMsg(UserIndex, "No está online", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
118             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.HandleExecute", Erl)
End Sub

Public Sub HandleBanChar(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim Reason   As String
102         username = Reader.ReadString8()
104         Reason = Reader.ReadString8()
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
108             Call BanPJ(UserIndex, username, Reason)
            Else
110             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBanChar", Erl)
End Sub

Public Sub HandleUnbanChar(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
102             username = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
106             If Not PersonajeExiste(username) Then
108                 Call WriteConsoleMsg(UserIndex, "El personaje no existe.", e_FontTypeNames.FONTTYPE_INFO)
                Else
110                 If BANCheck(username) Then
112                     Call SavePenaDatabase(username, .name & ": UNBAN. " & Date & " " & Time)
114                     Call UnBanDatabase(username)

116                     Call LogGM(.name, "/UNBAN a " & username)
118                     Call WriteConsoleMsg(UserIndex, username & " desbaneado.", e_FontTypeNames.FONTTYPE_INFO)
                    Else
120                     Call WriteConsoleMsg(UserIndex, username & " no esta baneado. Imposible desbanear.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            Else
122             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUnbanChar", Erl)
End Sub

Public Sub HandleNPCFollow(ByVal UserIndex As Integer)
        On Error GoTo HandleNPCFollow_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If IsValidNpcRef(.flags.TargetNPC) Then
108             Call DoFollow(.flags.TargetNPC.ArrayIndex, .Name)
110             NpcList(.flags.TargetNPC.ArrayIndex).flags.Inmovilizado = 0
112             NpcList(.flags.TargetNPC.ArrayIndex).flags.Paralizado = 0
114             NpcList(.flags.TargetNPC.ArrayIndex).Contadores.Paralisis = 0
            End If
        End With
        Exit Sub

HandleNPCFollow_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNPCFollow", Erl)
End Sub

Public Sub HandleSummonChar(ByVal UserIndex As Integer)
    'Author: Nicolas Matias Gonzalez (NIGO)
    On Error GoTo ErrHandler
100 With UserList(UserIndex)
        Dim username As String
        Dim tUser    As t_UserReference
        
102     username = Reader.ReadString8()
104     If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
106         If LenB(username) <> 0 Then
108             tUser = NameIndex(username)
110             If Not IsValidUserRef(tUser) Then
112                 Call WriteConsoleMsg(UserIndex, "El jugador no está online.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
114         ElseIf IsValidUserRef(.flags.targetUser) Then
116             tUser = .flags.TargetUser
                ' Mover NPCs
118         ElseIf IsValidNpcRef(.flags.TargetNPC) Then
120             If NpcList(.flags.TargetNPC.ArrayIndex).Pos.map = .Pos.map Then
122                 Call WarpNpcChar(.flags.TargetNPC.ArrayIndex, .Pos.map, .Pos.X, .Pos.y + 1, True)
124                 Call WriteConsoleMsg(UserIndex, "Has desplazado a la criatura.", e_FontTypeNames.FONTTYPE_INFO)
                Else
126                 Call WriteConsoleMsg(UserIndex, "Sólo puedes mover NPCs dentro del mismo mapa.", e_FontTypeNames.FONTTYPE_INFO)
                End If
                Exit Sub
            Else
                Exit Sub
            End If

128         If CompararPrivilegiosUser(tUser.ArrayIndex, userIndex) > 0 Then
130             Call WriteConsoleMsg(UserIndex, "Se le ha avisado a " & UserList(tUser.ArrayIndex).Name & " que quieres traerlo a tu posición.", e_FontTypeNames.FONTTYPE_INFO)
132             Call WriteConsoleMsg(tUser.ArrayIndex, .Name & " quiere transportarte a su ubicación. Escribe /ira " & .Name & " para ir.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
            Dim NotConsejero As Boolean
134         NotConsejero = (.flags.Privilegios And e_PlayerType.Consejero) = 0
                
            ' Consejeros sólo pueden traer en el mismo mapa
136         If NotConsejero Or .pos.map = UserList(tUser.ArrayIndex).pos.map Then
                
                 If .flags.Privilegios And (e_PlayerType.SemiDios) Then
                    If Not EsMapaEvento(.Pos.map) Then 'HarThaos ReyarB hay que sacar eso de mapa seguro y poner mapas de eventos [MapasEventos]
                        Call WriteConsoleMsg(UserIndex, "Solamente puedes traer usuarios a mapas de evento.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    If UserList(tUser.ArrayIndex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes transportar a un muerto. ellos tienen /hogar", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                
                ' Si el admin está invisible no mostramos el nombre
138             If NotConsejero And .flags.AdminInvisible = 1 Then
140                 Call WriteConsoleMsg(tUser.ArrayIndex, "Te han trasportado.", e_FontTypeNames.FONTTYPE_INFO)
                Else
142                 Call WriteConsoleMsg(tUser.ArrayIndex, .name & " te ha trasportado.", e_FontTypeNames.FONTTYPE_INFO)
                End If
                   
                'Si lo sumonean a un mapa interdimensional desde uno no interdimensional me guardo la posición de donde viene.
144             If EsMapaInterdimensional(.pos.map) And Not EsMapaInterdimensional(UserList(tUser.ArrayIndex).pos.map) Then
146                 UserList(tUser.ArrayIndex).flags.ReturnPos = UserList(tUser.ArrayIndex).pos
                End If

148             Call WarpToLegalPos(tUser.ArrayIndex, .pos.map, .pos.x, .pos.y + 1, True, True)
                Call WriteConsoleMsg(UserIndex, "Has traído a " & UserList(tUser.ArrayIndex).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
152             Call LogGM(.name, "/SUM " & username & " Map:" & .Pos.map & " X:" & .Pos.X & " Y:" & .Pos.y)
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " ha trasladado al usuario " & username & " al Mapa " & .pos.Map, e_FontTypeNames.FONTTYPE_INFO))
            End If
        Else
154         Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
        End If
    End With
    Exit Sub
ErrHandler:
156 Call TraceError(Err.Number, Err.Description, "Protocol.HandleSummonChar", Erl)
End Sub

Public Sub HandleSpawnListRequest(ByVal UserIndex As Integer)
        On Error GoTo HandleSpawnListRequest_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If .flags.Privilegios And e_PlayerType.user Then
                Exit Sub
104         ElseIf .flags.Privilegios And e_PlayerType.Consejero Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            
108         ElseIf .flags.Privilegios And (e_PlayerType.SemiDios) Then
110             Call WriteConsoleMsg(UserIndex, "Servidor » La cantidad de NPCs disponible para tu rango está limitada.", e_FontTypeNames.FONTTYPE_INFO)
            End If
112         Call WriteSpawnList(UserIndex, UserList(UserIndex).flags.Privilegios And e_PlayerType.Admin Or e_PlayerType.Dios)
        End With
        Exit Sub
HandleSpawnListRequest_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpawnListRequest", Erl)
End Sub

Public Sub HandleSpawnCreature(ByVal UserIndex As Integer)
        On Error GoTo HandleSpawnCreature_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim npc As Integer
102             npc = Reader.ReadInt16()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             If npc > 0 And npc <= UBound(Declaraciones.SpawnList()) Then
                    If Declaraciones.SpawnList(npc).NpcName <> "Nada" And (Declaraciones.SpawnList(npc).PuedeInvocar Or (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin)) <> 0) Then
108                     Call SpawnNpc(Declaraciones.SpawnList(npc).npcIndex, .Pos, True, False)
                    End If
                End If
110             Call LogGM(.name, "Sumoneo " & Declaraciones.SpawnList(npc).NpcName)
            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub

HandleSpawnCreature_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSpawnCreature", Erl)
End Sub

Public Sub HandleResetNPCInventory(ByVal UserIndex As Integer)
        On Error GoTo HandleResetNPCInventory_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
108         Call ResetNpcInv(.flags.TargetNPC.ArrayIndex)
110         Call LogGM(.Name, "/RESETINV " & NpcList(.flags.TargetNPC.ArrayIndex).Name)
        End With
        Exit Sub
HandleResetNPCInventory_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleResetNPCInventory", Erl)
End Sub

Public Sub HandleCleanWorld(ByVal UserIndex As Integer)
        On Error GoTo HandleCleanWorld_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
108         Call WriteConsoleMsg(UserIndex, "Se han limpiado los items del suelo.", e_FontTypeNames.FONTTYPE_INFO)
        End With
        Exit Sub

HandleCleanWorld_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCleanWorld", Erl)
End Sub

Public Sub HandleServerMessage(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim Message As String
102             Message = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) Then
106             If LenB(Message) <> 0 Then
108                 Call LogGM(.name, "Mensaje Broadcast:" & Message)
110                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & "> " & Message, e_FontTypeNames.FONTTYPE_SERVER))
                End If
            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub

ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleServerMessage", Erl)
End Sub

Public Sub HandleNickToIP(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        'Pablo (ToxicWaste): Agrego para uqe el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
            Dim priv     As e_PlayerType
        
102         username = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
        
106             tUser = NameIndex(username)
108             Call LogGM(.name, "NICK2IP Solicito la IP de " & username)

110             If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
112                 priv = e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin
                Else
114                 priv = e_PlayerType.user
                End If
            
116             If IsValidUserRef(tUser) Then
118                 If UserList(tUser.ArrayIndex).flags.Privilegios And priv Then
120                     Call WriteConsoleMsg(UserIndex, "El ip de " & username & " es " & UserList(tUser.ArrayIndex).ConnectionDetails.IP, e_FontTypeNames.FONTTYPE_INFO)
                        Dim IP    As String
                        Dim lista As String
                        Dim LoopC As Long
122                     IP = UserList(tUser.ArrayIndex).ConnectionDetails.IP
124                     For LoopC = 1 To LastUser
126                         If UserList(LoopC).ConnectionDetails.IP = IP Then
128                             If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
130                                 If UserList(LoopC).flags.Privilegios And priv Then
132                                     lista = lista & UserList(LoopC).name & ", "
                                    End If
                                End If
                            End If
134                     Next LoopC
136                     If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                
138                     Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
140                 Call WriteConsoleMsg(UserIndex, "No hay ningun personaje con ese nick", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
142             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNickToIP", Erl)
End Sub

Public Sub HandleIPToNick(ByVal UserIndex As Integer)
        On Error GoTo HandleIPToNick_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim IP    As String
            Dim LoopC As Long
            Dim lista As String
            Dim priv  As e_PlayerType
        
102         IP = Reader.ReadInt8() & "."
104         IP = IP & Reader.ReadInt8() & "."
106         IP = IP & Reader.ReadInt8() & "."
108         IP = IP & Reader.ReadInt8()
        
110         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
114         Call LogGM(.name, "IP2NICK Solicito los Nicks de IP " & IP)
        
116         If .flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin) Then
118             priv = e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios Or e_PlayerType.Admin
            Else
120             priv = e_PlayerType.user
            End If

122         For LoopC = 1 To LastUser
124             If UserList(LoopC).ConnectionDetails.IP = IP Then
126                 If LenB(UserList(LoopC).name) <> 0 And UserList(LoopC).flags.UserLogged Then
128                     If UserList(LoopC).flags.Privilegios And priv Then
130                         lista = lista & UserList(LoopC).name & ", "
                        End If
                    End If
                End If
132         Next LoopC
134         If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
136         Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, e_FontTypeNames.FONTTYPE_INFO)

        End With
        Exit Sub
HandleIPToNick_Err:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleIPToNick", Erl)
    
End Sub

Public Sub HandleTeleportCreate(ByVal UserIndex As Integer)
        On Error GoTo HandleTeleportCreate_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim Mapa As Integer
            Dim X    As Byte
            Dim y    As Byte
            Dim Radio As Byte
            Dim Motivo As String
102         Mapa = Reader.ReadInt16()
104         X = Reader.ReadInt8()
106         y = Reader.ReadInt8()
            Radio = Reader.ReadInt8()
            Motivo = Reader.ReadString8()
        
108         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
110             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
112         Call LogGM(.name, "/CT " & Mapa & "," & X & "," & y & "," & Motivo)
114         If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, y) Then Exit Sub
116         If MapData(.Pos.map, .Pos.X, .Pos.y - 1).ObjInfo.objIndex > 0 Then Exit Sub
118         If MapData(.Pos.map, .Pos.X, .Pos.y - 1).TileExit.map > 0 Then Exit Sub
        
120         If MapData(Mapa, X, y).ObjInfo.objIndex > 0 Then
122             Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
124         If MapData(Mapa, X, y).TileExit.map > 0 Then
126             Call WriteConsoleMsg(UserIndex, "No podés crear un teleport que apunte a la entrada de otro.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

            Dim Objeto As t_Obj
128         Objeto.amount = 1
            Select Case Radio
                Case 1
                    Objeto.objIndex = 370
                Case 2
                    Objeto.objIndex = 371
                Case Else
130                 Objeto.objIndex = 378
            End Select
132         Call MakeObj(Objeto, .Pos.map, .Pos.X, .Pos.y - 1)
        
134         With MapData(.Pos.map, .Pos.X, .Pos.y - 1)
136             .TileExit.map = Mapa
138             .TileExit.X = X
140             .TileExit.y = y
            End With
        End With
        Exit Sub
HandleTeleportCreate_Err:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTeleportCreate", Erl)
End Sub

Public Sub HandleTeleportDestroy(ByVal UserIndex As Integer)
        On Error GoTo HandleTeleportDestroy_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim Mapa As Integer
            Dim X    As Byte
            Dim y    As Byte
            '/dt
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
                
106         Mapa = .flags.TargetMap
108         X = .flags.TargetX
110         y = .flags.TargetY
        
112         If Not InMapBounds(Mapa, X, y) Then Exit Sub
        
114         With MapData(Mapa, X, y)
                'Si no tengo objeto y no tengo traslado
116             If .ObjInfo.objIndex = 0 And .TileExit.map = 0 Then Exit Sub
                'Si no tengo objeto pero tengo traslado
118             If .ObjInfo.objIndex = 0 And .TileExit.map > 0 Then
120                 Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & X & "," & y)
122                 .TileExit.map = 0
124                 .TileExit.X = 0
126                 .TileExit.y = 0
                    'si tengo objeto y traslado
128             ElseIf .ObjInfo.objIndex > 0 And ObjData(.ObjInfo.objIndex).OBJType = e_OBJType.otTeleport Then
130                 Call LogGM(UserList(UserIndex).name, "/DT: " & Mapa & "," & X & "," & y)
132                 Call EraseObj(.ObjInfo.amount, Mapa, X, y)
134                 If MapData(.TileExit.map, .TileExit.X, .TileExit.y).ObjInfo.objIndex = 651 Then
136                     Call EraseObj(1, .TileExit.map, .TileExit.X, .TileExit.y)
                    End If
138                 .TileExit.map = 0
140                 .TileExit.X = 0
142                 .TileExit.y = 0
                End If
            End With
        End With
        Exit Sub

HandleTeleportDestroy_Err:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTeleportDestroy", Erl)
End Sub

Public Sub HandleRainToggle(ByVal UserIndex As Integer)
        On Error GoTo HandleRainToggle_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         Call LogGM(.name, "/LLUVIA")
108         Lloviendo = Not Lloviendo
110         Nebando = Not Nebando
112         Call SendData(SendTarget.ToAll, 0, PrepareMessageRainToggle())
114         Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
116         If Lloviendo Then
118             Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(404, NO_3D_SOUND, NO_3D_SOUND)) ' Explota un trueno
120             Call SendData(SendTarget.ToAll, 0, PrepareMessageFlashScreen(&HF5D3F3, 250)) 'Rayo
122             Call ApagarFogatas
            End If
        End With
        Exit Sub
HandleRainToggle_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRainToggle", Erl)
End Sub

Public Sub HandleSetCharDescription(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim tUser As t_UserReference
            Dim Desc  As String
102         Desc = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
106             tUser = .flags.TargetUser
108             If IsValidUserRef(tUser) Then
110                 UserList(tUser.ArrayIndex).DescRM = Desc
                Else
112                 Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes!", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetCharDescription", Erl)
End Sub

Public Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)
        On Error GoTo HanldeForceMIDIToMap_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim midiID As Byte
            Dim Mapa   As Integer
        
102         midiID = Reader.ReadInt8
104         Mapa = Reader.ReadInt16
            'Solo dioses, admins y RMS
106         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
                'Si el mapa no fue enviado tomo el actual
108             If Not InMapBounds(Mapa, 50, 50) Then
110                 Mapa = .Pos.map
                End If
        
112             If midiID = 0 Then
                    'Ponemos el default del mapa
114                 Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(MapInfo(.Pos.map).music_numberLow))
                Else
                    'Ponemos el pedido por el GM
116                 Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMidi(midiID))
                End If
            End If
        End With
        Exit Sub
HanldeForceMIDIToMap_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HanldeForceMIDIToMap", Erl)
End Sub

Public Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)
        On Error GoTo HandleForceWAVEToMap_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim waveID As Byte
            Dim Mapa   As Integer
            Dim X      As Byte
            Dim y      As Byte
102         waveID = Reader.ReadInt8()
104         Mapa = Reader.ReadInt16()
106         X = Reader.ReadInt8()
108         y = Reader.ReadInt8()
            'Solo dioses, admins y RMS
110         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.RoleMaster)) Then
                'Si el mapa no fue enviado tomo el actual
112             If Not InMapBounds(Mapa, X, y) Then
114                 Mapa = .Pos.map
116                 X = .Pos.X
118                 y = .Pos.y
                End If
                'Ponemos el pedido por el GM
120             Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, y))
            End If
        End With
        Exit Sub
HandleForceWAVEToMap_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceWAVEToMap", Erl)
End Sub

Public Sub HandleTalkAsNPC(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim Message As String
102             Message = Reader.ReadString8()
            'Solo dioses, admins y RMS
104         If (.flags.Privilegios And (e_PlayerType.Dios Or e_PlayerType.Admin Or e_PlayerType.SemiDios)) Then
                'Asegurarse haya un NPC seleccionado
106             If IsValidNpcRef(.flags.TargetNPC) Then
108                 Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC.ArrayIndex, PrepareMessageChatOverHead(Message, NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex, vbWhite))
                Else
110                 Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTalkAsNPC", Erl)
End Sub

Public Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)
        On Error GoTo HandleDestroyAllItemsInArea_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim X As Long
            Dim y As Long
        
106         For y = .Pos.y - MinYBorder + 1 To .Pos.y + MinYBorder - 1
108             For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
110                 If X > 0 And y > 0 And X < 101 And y < 101 Then
112                     If MapData(.Pos.map, X, y).ObjInfo.objIndex > 0 Then
114                         If ItemNoEsDeMapa(MapData(.Pos.map, X, y).ObjInfo.objIndex) Then
116                             Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, X, y)
                            End If
                        End If
                    End If
118             Next X
120         Next y
122         Call LogGM(UserList(UserIndex).name, "/MASSDEST")
        End With
        Exit Sub
HandleDestroyAllItemsInArea_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDestroyAllItemsInArea", Erl)
End Sub

Public Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)
        On Error GoTo HandleItemsInTheFloor_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim tObj  As Integer
            Dim lista As String
            Dim X     As Long
            Dim y     As Long
106         For X = 5 To 95
108             For y = 5 To 95
110                 tObj = MapData(.Pos.map, X, y).ObjInfo.objIndex
112                 If tObj > 0 Then
114                     If ObjData(tObj).OBJType <> e_OBJType.otArboles Then
116                         Call WriteConsoleMsg(UserIndex, "(" & X & "," & y & ") " & ObjData(tObj).name, e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    End If
118             Next y
120         Next X
        End With
        Exit Sub
HandleItemsInTheFloor_Err:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleItemsInTheFloor", Erl)
End Sub

Public Sub HandleMakeDumb(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             tUser = NameIndex(username)
                'para deteccion de aoice
108             If Not IsValidUserRef(tUser) Then
110                 Call WriteConsoleMsg(UserIndex, "Offline", e_FontTypeNames.FONTTYPE_INFO)
                Else
112                 Call WriteDumb(tUser.ArrayIndex)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMakeDumb", Erl)
End Sub

Public Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)
        'Last Modification: 12/30/06
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             tUser = NameIndex(username)
                'para deteccion de aoice
108             If Not IsValidUserRef(tUser) Then
110                 Call WriteConsoleMsg(UserIndex, "Offline", e_FontTypeNames.FONTTYPE_INFO)
                Else
112                 Call WriteDumbNoMore(tUser.ArrayIndex)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMakeDumbNoMore", Erl)
End Sub

Public Sub HandleSetTrigger(ByVal UserIndex As Integer)
        On Error GoTo HandleSetTrigger_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim tTrigger As Byte
            Dim tLog     As String
        
102         tTrigger = Reader.ReadInt8()
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
106         If tTrigger >= 0 Then
108             MapData(.Pos.map, .Pos.X, .Pos.y).trigger = tTrigger
110             tLog = "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & "," & .Pos.y
112             Call LogGM(.name, tLog)
114             Call WriteConsoleMsg(UserIndex, tLog, e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
HandleSetTrigger_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetTrigger", Erl)
End Sub

Public Sub HandleAskTrigger(ByVal UserIndex As Integer)
        On Error GoTo HandleAskTrigger_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
        Dim tTrigger As Byte
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         tTrigger = MapData(.Pos.map, .Pos.X, .Pos.y).trigger
106         Call LogGM(.name, "Miro el trigger en " & .Pos.map & "," & .Pos.X & "," & .Pos.y & ". Era " & tTrigger)
108         Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.map & " " & .Pos.X & ", " & .Pos.y, e_FontTypeNames.FONTTYPE_INFO)
        End With
        Exit Sub
HandleAskTrigger_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAskTrigger", Erl)
End Sub



Public Sub HandleCreateItem(ByVal UserIndex As Integer)
        On Error GoTo HandleCreateItem_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim tObj    As Integer
            Dim Cuantos As Integer
        
102         tObj = Reader.ReadInt16()
104         Cuantos = Reader.ReadInt16()
    
            ' Si es usuario, lo sacamos cagando.
106         If Not EsGM(UserIndex) Or (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios)) Then
                Call WriteConsoleMsg(UserIndex, "Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            ' Si es Dios, dejamos crear un item siempre y cuando pueda estar en el inventario.
108         If (.flags.Privilegios And e_PlayerType.Dios) <> 0 And ObjData(tObj).Agarrable = 1 Then Exit Sub
            ' Si hace mas de 10000, lo sacamos cagando.
110         If Cuantos > MAX_INVENTORY_OBJS Then
112             Call WriteConsoleMsg(UserIndex, "Solo podés crear hasta " & CStr(MAX_INVENTORY_OBJS) & " unidades", e_FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
            ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
114         If tObj < 1 Or tObj > NumObjDatas Then Exit Sub
        
            ' El nombre del objeto es nulo?
116         If LenB(ObjData(tObj).name) = 0 Then Exit Sub
        
            Dim Objeto As t_Obj
118         Objeto.amount = Cuantos
120         Objeto.objIndex = tObj
            ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demAs t_Objs. que no deberian estar en el inventario)
            '   0 = SI
            '   1 = NO
122         If ObjData(tObj).Agarrable = 0 Then
                ' Trato de meterlo en el inventario.
124             If MeterItemEnInventario(UserIndex, Objeto) Then
126                 Call WriteConsoleMsg(UserIndex, "Has creado " & Objeto.amount & " unidades de " & ObjData(tObj).name & ".", e_FontTypeNames.FONTTYPE_INFO)
                Else
128                 Call WriteConsoleMsg(UserIndex, "No tenes espacio en tu inventario para crear el item.", e_FontTypeNames.FONTTYPE_INFO)
                    ' Si no hay espacio y es Admin, lo tiro al piso.
130                 If (.flags.Privilegios And e_PlayerType.Admin) <> 0 Then
132                     Call TirarItemAlPiso(.Pos, Objeto)
134                     Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", e_FontTypeNames.FONTTYPE_GUILD)
                    End If
                End If
            Else
                ' Crear el item NO AGARRARBLE y tirarlo al piso.
                ' Si no hay espacio y es Admin, lo tiro al piso.
136             If (.flags.Privilegios And e_PlayerType.Admin) <> 0 Then
138                 Call TirarItemAlPiso(.Pos, Objeto)
140                 Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", e_FontTypeNames.FONTTYPE_GUILD)
                End If
            End If

142         Call LogGM(.name, "/CI: " & tObj & " Cantidad : " & Cuantos)

        End With
        Exit Sub
HandleCreateItem_Err:
144     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateItem", Erl)
End Sub

Public Sub HandleDestroyItems(ByVal UserIndex As Integer)
        On Error GoTo HandleDestroyItems_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         If MapData(.Pos.map, .Pos.X, .Pos.y).ObjInfo.objIndex = 0 Then Exit Sub
108         Call LogGM(.name, "/DEST")
110         Call EraseObj(MAX_INVENTORY_OBJS, .Pos.map, .Pos.X, .Pos.y)
        End With
        Exit Sub
HandleDestroyItems_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDestroyItems", Erl)
End Sub

Public Sub HandleForceMIDIAll(ByVal UserIndex As Integer)
        On Error GoTo HandleForceMIDIAll_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim midiID As Byte
102             midiID = Reader.ReadInt8()
        
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
108         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " broadcast musica: " & midiID, e_FontTypeNames.FONTTYPE_SERVER))
110         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayMidi(midiID))
        End With
        Exit Sub
HandleForceMIDIAll_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceMIDIAll", Erl)
End Sub

Public Sub HandleForceWAVEAll(ByVal UserIndex As Integer)
        On Error GoTo HandleForceWAVEAll_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            Dim waveID As Byte
102             waveID = Reader.ReadInt8()
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
108         Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))
        End With
        Exit Sub
HandleForceWAVEAll_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleForceWAVEAll", Erl)
End Sub

Public Sub HandleRemovePunishment(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username   As String
            Dim punishment As Byte
            Dim NewText    As String
        
102         username = Reader.ReadString8()
104         punishment = Reader.ReadInt8
106         NewText = Reader.ReadString8()
        
108         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
110             If LenB(username) = 0 Then
112                 Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", e_FontTypeNames.FONTTYPE_INFO)
                Else
114                 If (InStrB(username, "\") <> 0) Then
116                     username = Replace(username, "\", "")
                    End If

118                 If (InStrB(username, "/") <> 0) Then
120                     username = Replace(username, "/", "")
                    End If
                
122                 If PersonajeExiste(username) Then
124                     Call LogGM(.Name, "Borro la pena " & punishment & " de " & username & " y la cambió por: " & NewText)
128                     Call CambiarPenaDatabase(username, punishment, .name & ": <" & NewText & "> " & Date & " " & Time)
132                     Call WriteConsoleMsg(UserIndex, "Pena Modificada.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
            End If
        End With
        Exit Sub
ErrHandler:
134     Call TraceError(Err.Number, Err.Description, "Protocol.HandleRemovePunishment", Erl)
End Sub

Public Sub HandleTile_BlockedToggle(ByVal UserIndex As Integer)
        On Error GoTo HandleTile_BlockedToggle_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

106         Call LogGM(.name, "/BLOQ")
108         If MapData(.Pos.map, .Pos.X, .Pos.y).Blocked = 0 Then
110             MapData(.Pos.map, .Pos.X, .Pos.y).Blocked = e_Block.ALL_SIDES Or e_Block.GM
            Else
112             MapData(.Pos.map, .Pos.X, .Pos.y).Blocked = 0
            End If
        
114         Call Bloquear(True, .Pos.map, .Pos.X, .Pos.y, IIf(MapData(.Pos.map, .Pos.X, .Pos.y).Blocked > 0, e_Block.ALL_SIDES, 0))
        End With
        Exit Sub
HandleTile_BlockedToggle_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTile_BlockedToggle", Erl)
End Sub

Public Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)
        On Error GoTo HandleKillNPCNoRespawn_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
100     With UserList(UserIndex)
            If Not EsGM(UserIndex) Then Exit Sub
102         If .flags.Privilegios And e_PlayerType.Consejero Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         If Not IsValidNpcRef(.flags.TargetNPC) Then Exit Sub
108         Call QuitarNPC(.flags.TargetNPC.ArrayIndex, eGMCommand)
110         Call LogGM(.Name, "/MATA " & NpcList(.flags.TargetNPC.ArrayIndex).Name)
        End With
        Exit Sub
HandleKillNPCNoRespawn_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillNPCNoRespawn", Erl)
End Sub

Public Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)
        On Error GoTo HandleKillAllNearbyNPCs_Err
        'Author: Nicolas Matias Gonzalez (NIGO)
        'ReyarB
100     With UserList(UserIndex)
            If Not EsGM(UserIndex) Then Exit Sub
102         If (.flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
            Dim X As Long
            Dim y As Long
108         For y = .Pos.y - MinYBorder + 1 To .Pos.y + MinYBorder - 1
110             For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1
112                 If X > 0 And y > 0 And X < 101 And y < 101 Then
114                     If MapData(.Pos.map, X, y).npcIndex > 0 Then
116                         Call QuitarNPC(MapData(.Pos.map, X, y).NpcIndex, eGMCommand)
                        End If
                    End If
118             Next X
120         Next y
122         Call LogGM(.name, "/MASSKILL")
        End With
        Exit Sub
HandleKillAllNearbyNPCs_Err:
124     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKillAllNearbyNPCs", Erl)
End Sub

Public Sub HandleLastIP(ByVal UserIndex As Integer)
        'Author: Martín Trionfetti (HarThaoS) - Fernando Quinteros (Lord Fers)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)

            Dim username   As String
            Dim LoopC      As Byte
        
104         username = Reader.ReadString8()
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
                'Handle special chars
108             If (InStrB(username, "\") <> 0) Then
110                 username = Replace(username, "\", "")
                End If

112             If (InStrB(username, "\") <> 0) Then
114                 username = Replace(username, "/", "")
                End If

116             If (InStrB(username, "+") <> 0) Then
118                 username = Replace(username, "+", " ")
                End If
                
                Dim tUser As t_UserReference
                tUser = NameIndex(username)
                Dim RS As ADODB.Recordset
                Dim ipStr As String
                
120             If IsValidUserRef(tUser) Then
122                 Call LogGM(.name, "/LASTIP " & username)
                    Set RS = Query("SELECT last_ip FROM account WHERE id = ?", UserList(tUser.ArrayIndex).AccountID)
            
                    'Revisamos si recibio un resultado
                    If RS Is Nothing Then Exit Sub
                    If RS.BOF Or RS.EOF Then Exit Sub
                    
                    'Obtenemos la variable
                    ipStr = RS.Fields(0).value
                Else
                    Dim account_id As String
                    Set RS = Query("SELECT u.account_id FROM user u WHERE LOWER(u.name) = LOWER(?)", username)
                    
                    'Revisamos si recibio un resultado
                    If RS Is Nothing Then Exit Sub
                    If RS.BOF Or RS.EOF Then Exit Sub
                    
                    'Obtenemos la variable
                    account_id = RS.Fields(0).value
                    Set RS = Query("SELECT last_ip FROM account WHERE id = ?", account_id)
            
                    'Revisamos si recibio un resultado
                    If RS Is Nothing Then Exit Sub
                    If RS.BOF Or RS.EOF Then Exit Sub
                    
                    'Obtenemos la variable
                    ipStr = RS.Fields(0).value
                End If
                Dim countIps As Long
                countIps = UBound(Split(ipStr, ";"))
                
                If countIps <= 0 Then Exit Sub
                
                ReDim ip_list(0 To (countIps - 1)) As String
                ip_list = Split(ipStr, ";")
                Call WriteConsoleMsg(UserIndex, "Las últimas ips para el personaje son: ", e_FontTypeNames.FONTTYPE_INFO)
                For LoopC = 0 To (countIps - 1)
                    Call WriteConsoleMsg(UserIndex, ip_list(LoopC), e_FontTypeNames.FONTTYPE_INFO)
                Next LoopC
            Else
146             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
148     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLastIP", Erl)
End Sub

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)
        On Error GoTo HandleChangeMOTD_Err
        'Author: Juan Martín sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.RoleMaster Or e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then Exit Sub

            Dim auxiliaryString As String
            Dim LoopC           As Long
            
104         For LoopC = LBound(MOTD()) To UBound(MOTD())
106             auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
108         Next LoopC
        
110         If Len(auxiliaryString) >= 2 Then
112             If Right$(auxiliaryString, 2) = vbCrLf Then
114                 auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)
                End If
            End If
116         Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)
        End With
        Exit Sub
HandleChangeMOTD_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMOTD", Erl)
End Sub

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)
        'Author: Lucas Tavolaro Ortiz (Tavo)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim newMOTD           As String
            Dim auxiliaryString() As String
            Dim LoopC             As Long
        
102         newMOTD = Reader.ReadString8()
104         auxiliaryString = Split(newMOTD, vbCrLf)
        
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
108             Call LogGM(.name, "Ha fijado un nuevo MOTD")
            
110             MaxLines = UBound(auxiliaryString()) + 1
                If MaxLines > 0 Then
112             ReDim MOTD(1 To MaxLines)
114                 Call WriteVar(DatPath & "Motd.ini", "INIT", "NumLines", CStr(MaxLines))
116                 For LoopC = 1 To MaxLines
118                     Call WriteVar(DatPath & "Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                    
120                     MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
122                 Next LoopC
                
124                 Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    ReDim MOTD(0 To 0)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetMOTD", Erl)
End Sub

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)
    On Error GoTo HandleCreateNPC_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)
            Dim npcIndex As Integer
102         npcIndex = Reader.ReadInt16()

            If Not EsGM(UserIndex) Then Exit Sub
        
104         If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
112         npcIndex = SpawnNpc(npcIndex, .Pos, True, False)
114         If npcIndex <> 0 Then
116             Call LogGM(.name, "Sumoneo a " & NpcList(npcIndex).name & " en mapa " & .Pos.map)
            End If
        End With
        Exit Sub
HandleCreateNPC_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNPC", Erl)
End Sub

Public Sub HandleCreateNPCWithRespawn(ByVal UserIndex As Integer)
        On Error GoTo HandleCreateNPCWithRespawn_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)
            Dim npcIndex As Integer
        
102         npcIndex = Reader.ReadInt16()
            If Not EsGM(UserIndex) Then Exit Sub
104         If .flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo. Si el motivo es probar algo ya saben ir a Test", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         npcIndex = SpawnNpc(npcIndex, .Pos, True, True)
110         If npcIndex <> 0 Then
112             Call LogGM(.name, "Sumoneo con respawn " & NpcList(npcIndex).name & " en mapa " & .Pos.map)
            End If
        End With
        Exit Sub

HandleCreateNPCWithRespawn_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateNPCWithRespawn", Erl)
End Sub

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)
        On Error GoTo HandleImperialArmour_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

            Dim Index    As Byte
            Dim objIndex As Integer
        
102         Index = Reader.ReadInt8()
104         objIndex = Reader.ReadInt16()
        
106         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
108         Select Case Index
                Case 1
                    ' ArmaduraImperial1 = objindex
110             Case 2
                    ' ArmaduraImperial2 = objindex
112             Case 3
                    ' ArmaduraImperial3 = objindex
114             Case 4
                    ' TunicaMagoImperial = objindex
            End Select
        End With
        Exit Sub
HandleImperialArmour_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleImperialArmour", Erl)
End Sub

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)
        
        On Error GoTo HandleChaosArmour_Err

        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)

            Dim Index    As Byte
            Dim objIndex As Integer
        
102         Index = Reader.ReadInt8()
104         objIndex = Reader.ReadInt16()
        
106         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
108         Select Case Index
                Case 1
                    '   ArmaduraCaos1 = objindex
110             Case 2
                    '   ArmaduraCaos2 = objindex
112             Case 3
                    '   ArmaduraCaos3 = objindex
114             Case 4
                    '  TunicaMagoCaos = objindex
            End Select
        End With
        Exit Sub
HandleChaosArmour_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChaosArmour", Erl)
End Sub

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.name, "/CONDEN " & username)
108             tUser = NameIndex(username)
110             If IsValidUserRef(tUser) Then Call VolverCriminal(tUser.ArrayIndex)
            End If
        End With
        Exit Sub
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleTurnCriminal", Erl)
End Sub

Public Sub HandleAlterName(ByVal UserIndex As Integer)
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        'Reads the userName and newUser Packets
        Dim username     As String
        Dim NewName      As String
        Dim tUser        As t_UserReference
        Dim GuildIndex   As Integer

        username = UCase$(Reader.ReadString8())
        NewName = Reader.ReadString8()

        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then Exit Sub

        If LenB(username) = 0 Or LenB(NewName) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        tUser = NameIndex(username)
        If IsValidUserRef(tUser) Then
            If UserList(tUser.ArrayIndex).GuildIndex > 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje " & username & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            If Not PersonajeExiste(username) Then
                Call WriteConsoleMsg(UserIndex, "El personaje " & username & " es inexistente.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            GuildIndex = GetUserGuildIndexDatabase(username)
            If GuildIndex > 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje " & username & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        If PersonajeExiste(NewName) Then
            Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Call ChangeNameDatabase(username, NewName)
        Call WriteConsoleMsg(UserIndex, "Transferencia exitosa", e_FontTypeNames.FONTTYPE_INFO)
        Call SavePenaDatabase(username, .name & ": nombre cambiado de """ & username & """ a """ & NewName & """. " & Date & " " & Time)
        Call SendData(SendTarget.ToGM, 0, PrepareMessageConsoleMsg("Administración » " & .Name & " cambió el nombre del usuario """ & username & """ por """ & NewName & """.", e_FontTypeNames.FONTTYPE_GM))
        Call LogGM(.name, "Ha cambiado de nombre al usuario """ & username & """. Ahora se llama """ & NewName & """.")
        
        If IsValidUserRef(tUser) Then
            UserList(tUser.ArrayIndex).name = NewName
            Call RefreshCharStatus(tUser.ArrayIndex)
        End If
    End With
    Exit Sub

ErrHandler:
150     Call TraceError(Err.Number, Err.Description, "Protocol.HandleAlterName", Erl)
End Sub

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)
        On Error GoTo HandleDoBackUp_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
104         Call LogGM(.name, .name & " ha hecho un backup")
106         Call ES.DoBackUp 'Sino lo confunde con la id del paquete
        End With
        Exit Sub
HandleDoBackUp_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDoBackUp", Erl)
End Sub

Public Sub HandleChangeMapInfoPK(ByVal UserIndex As Integer)
        
        On Error GoTo HandleChangeMapInfoPK_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
            Dim isMapPk As Boolean
102         isMapPk = Reader.ReadBool()
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
108         Call LogGM(.name, .name & " ha cambiado la informacion sobre si es seguro el mapa.")
110         MapInfo(.Pos.map).Seguro = IIf(isMapPk, 1, 0)
112         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Seguro: " & MapInfo(.Pos.map).Seguro, e_FontTypeNames.FONTTYPE_INFO)
        End With
        Exit Sub
HandleChangeMapInfoPK_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoPK", Erl)
End Sub

Public Sub HandleChangeMapInfoBackup(ByVal UserIndex As Integer)
        On Error GoTo HandleChangeMapInfoBackup_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)
            Dim doTheBackUp As Boolean
102         doTheBackUp = Reader.ReadBool()
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
106         Call LogGM(.Name, .Name & " ha cambiado la información sobre el BackUp")
            'Change the boolean to byte in a fast way
108         If doTheBackUp Then
110             MapInfo(.Pos.map).backup_mode = 1
            Else
112             MapInfo(.Pos.map).backup_mode = 0
            End If
        
            'Change the boolean to string in a fast way
114         Call WriteVar(MapPath & "mapa" & .Pos.map & ".dat", "Mapa" & .Pos.map, "backup", MapInfo(.Pos.map).backup_mode)
116         Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Backup: " & MapInfo(.Pos.map).backup_mode, e_FontTypeNames.FONTTYPE_INFO)
        End With
        Exit Sub

HandleChangeMapInfoBackup_Err:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoBackup", Erl)
End Sub

Public Sub HandleChangeMapInfoRestricted(ByVal UserIndex As Integer)
        'Author: Pablo (ToxicWaste)
        'Restringido -> Options: "NEWBIE", "SINMAGIA", "SININVI", "NOPKS", "NOCIUD".
        On Error GoTo ErrHandler
        Dim tStr As String
    
100     With UserList(UserIndex)
102         tStr = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) <> 0 Then
106             Select Case UCase$(tStr)
                    Case "NEWBIE"
108                     MapInfo(.Pos.map).Newbie = Not MapInfo(.Pos.map).Newbie
110                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & ": Newbie = " & MapInfo(.Pos.map).Newbie, e_FontTypeNames.FONTTYPE_INFO)
112                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": Newbie = " & MapInfo(.Pos.Map).Newbie)
                        
114                 Case "SINMAGIA"
116                     MapInfo(.Pos.map).SinMagia = Not MapInfo(.Pos.map).SinMagia
118                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & ": SinMagia = " & MapInfo(.Pos.map).SinMagia, e_FontTypeNames.FONTTYPE_INFO)
120                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": SinMagia = " & MapInfo(.Pos.Map).SinMagia)
                        
122                 Case "NOPKS"
124                     MapInfo(.Pos.map).NoPKs = Not MapInfo(.Pos.map).NoPKs
126                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & ": NoPKs = " & MapInfo(.Pos.map).NoPKs, e_FontTypeNames.FONTTYPE_INFO)
128                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": NoPKs = " & MapInfo(.Pos.Map).NoPKs)
                        
130                 Case "NOCIUD"
132                     MapInfo(.Pos.map).NoCiudadanos = Not MapInfo(.Pos.map).NoCiudadanos
134                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & ": NoCiudadanos = " & MapInfo(.Pos.map).NoCiudadanos, e_FontTypeNames.FONTTYPE_INFO)
136                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": NoCiudadanos = " & MapInfo(.Pos.Map).NoCiudadanos)
                        
138                 Case "SININVI"
140                     MapInfo(.Pos.map).SinInviOcul = Not MapInfo(.Pos.map).SinInviOcul
142                     Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & ": SinInvi = " & MapInfo(.Pos.map).SinInviOcul, e_FontTypeNames.FONTTYPE_INFO)
144                     Call LogGM(.Name, .Name & " ha cambiado la restricción del mapa " & .Pos.Map & ": SinInvi = " & MapInfo(.Pos.Map).SinInviOcul)
                
146                 Case Else
148                     Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'SINMAGIA', 'SININVI', 'NOPKS', 'NOCIUD'", e_FontTypeNames.FONTTYPE_INFO)

                End Select
            End If
        End With
        Exit Sub
ErrHandler:
150     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoRestricted", Erl)
End Sub

Public Sub HandleChangeMapInfoNoMagic(ByVal UserIndex As Integer)
        On Error GoTo HandleChangeMapInfoNoMagic_Err
        'Author: Pablo (ToxicWaste)
        'MagiaSinEfecto -> Options: "1" , "0".
        Dim nomagic As Boolean
100     With UserList(UserIndex)
102         nomagic = Reader.ReadBool
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar la Magia el mapa.")
            End If
        End With
        Exit Sub
HandleChangeMapInfoNoMagic_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoMagic", Erl)
End Sub

Public Sub HandleChangeMapInfoNoInvi(ByVal UserIndex As Integer)
        On Error GoTo HandleChangeMapInfoNoInvi_Err
        'Author: Pablo (ToxicWaste)
        'InviSinEfecto -> Options: "1", "0"
        Dim noinvi As Boolean
    
100     With UserList(UserIndex)
102         noinvi = Reader.ReadBool()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Invisibilidad el mapa.")
            End If
        End With
        Exit Sub
HandleChangeMapInfoNoInvi_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoInvi", Erl)
End Sub

Public Sub HandleChangeMapInfoNoResu(ByVal UserIndex As Integer)
        On Error GoTo HandleChangeMapInfoNoResu_Err
        Dim noresu As Boolean
    
100     With UserList(UserIndex)
102         noresu = Reader.ReadBool()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             Call LogGM(.name, .name & " ha cambiado la informacion sobre si esta permitido usar Resucitar el mapa.")
            End If
        End With
        Exit Sub
HandleChangeMapInfoNoResu_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapInfoNoResu", Erl)
End Sub

Public Sub HandleChangeMapInfoLand(ByVal UserIndex As Integer)
        'Author: Pablo (ToxicWaste)
        'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
        On Error GoTo ErrHandler
        Dim tStr As String
100     With UserList(UserIndex)
102         tStr = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
108                 Call LogGM(.name, .name & " ha cambiado la informacion del Terreno del mapa.")
110                 MapInfo(UserList(UserIndex).Pos.map).terrain = tStr
112                 Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Terreno", tStr)
114                 Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Terreno: " & MapInfo(.Pos.map).terrain, e_FontTypeNames.FONTTYPE_INFO)
                Else
116                 Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", e_FontTypeNames.FONTTYPE_INFO)
118                 Call WriteConsoleMsg(UserIndex, "Igualmente, el ínico ítil es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el Mapa", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleChangeMapInfoZone(ByVal UserIndex As Integer)
        'Author: Pablo (ToxicWaste)
        'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
        On Error GoTo ErrHandler
        Dim tStr As String
    
100     With UserList(UserIndex)
102         tStr = Reader.ReadString8()
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
106             If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
108                 Call LogGM(.name, .name & " ha cambiado la informacion de la Zona del mapa.")
110                 MapInfo(UserList(UserIndex).Pos.map).zone = tStr
112                 Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.map & ".dat", "Mapa" & UserList(UserIndex).Pos.map, "Zona", tStr)
114                 Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.map & " Zona: " & MapInfo(.Pos.map).zone, e_FontTypeNames.FONTTYPE_INFO)
                Else
116                 Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", e_FontTypeNames.FONTTYPE_INFO)
118                 Call WriteConsoleMsg(UserIndex, "Igualmente, el ínico ítil es 'DUNGEON' ya que al ingresarlo, NO se sentirí el efecto de la lluvia en este mapa.", e_FontTypeNames.FONTTYPE_INFO)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleChangeMapSetting(ByVal UserIndex As Integer)
On Error GoTo ChangeMapSetting_Err
        Dim SettingType As Byte
        SettingType = Reader.ReadInt8()
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.RoleMaster)) Then
104             Select Case SettingType
                    Case e_MapSetting.e_DropItems
108                     MapInfo(UserList(UserIndex).pos.map).DropItems = Reader.ReadInt8()
110                     Call LogGM(.name, .name & " ha cambiado la configuracion el dropeo de items en el mapa" & UserList(UserIndex).pos.map & " a " & MapInfo(UserList(UserIndex).pos.map).DropItems)
112                     Call WriteConsoleMsg(UserIndex, "Mapa actualizado correctamente", e_FontTypeNames.FONTTYPE_INFO)
                    Case e_MapSetting.e_SafeFight
116                     MapInfo(UserList(UserIndex).pos.map).SafeFightMap = Reader.ReadInt8()
118                     Call LogGM(.name, .name & " ha cambiado la configuracion el pelea segura del mapa" & UserList(UserIndex).pos.map & " a " & MapInfo(UserList(UserIndex).pos.map).DropItems)
120                     Call WriteConsoleMsg(UserIndex, "Mapa actualizado correctamente", e_FontTypeNames.FONTTYPE_INFO)
                    Case e_MapSetting.e_FriendlyFire
122                     MapInfo(UserList(UserIndex).pos.map).FriendlyFire = Reader.ReadInt8()
124                     Call LogGM(.name, .name & " ha cambiado la configuracion el friendly fire del mapa" & UserList(UserIndex).pos.map & " a " & MapInfo(UserList(UserIndex).pos.map).DropItems)
126                     Call WriteConsoleMsg(UserIndex, "Mapa actualizado correctamente", e_FontTypeNames.FONTTYPE_INFO)
                    Case Else
128                     Call WriteConsoleMsg(UserIndex, "Opcion no disponible", e_FontTypeNames.FONTTYPE_INFO)
                End Select
            End If
        End With
        Exit Sub
ChangeMapSetting_Err:
130     Call TraceError(Err.Number, Err.Description, "Protocol.HandleChangeMapSetting", Erl)
End Sub

Public Sub HandleSaveChars(ByVal UserIndex As Integer)
        On Error GoTo HandleSaveChars_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         Call LogGM(.name, .name & " ha guardado todos los chars")
108         Call GuardarUsuarios
        End With
        Exit Sub
HandleSaveChars_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSaveChars", Erl)
End Sub

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)
        On Error GoTo HandleCleanSOS_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
104         Call LogGM(.name, .name & " ha borrado los SOS")
106         Call Ayuda.Reset
        End With
        Exit Sub

HandleCleanSOS_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCleanSOS", Erl)
End Sub

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)
        On Error GoTo HandleShowServerForm_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
        
104         Call LogGM(.name, .name & " ha solicitado mostrar el formulario del servidor.")
106         Call frmMain.mnuMostrar_Click
        End With
        Exit Sub
HandleShowServerForm_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleShowServerForm", Erl)
End Sub

Public Sub HandleNight(ByVal UserIndex As Integer)
        On Error GoTo HandleNight_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         HoraMundo = GetTickCount()
108         Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
        End With
        Exit Sub
HandleNight_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNight", Erl)
End Sub

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)
        On Error GoTo HandleKickAllChars_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
104         Call LogGM(.name, .name & " ha echado a todos los personajes.")
106         Call EcharPjsNoPrivilegiados
        End With
        Exit Sub
HandleKickAllChars_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleKickAllChars", Erl)
End Sub

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)
        On Error GoTo HandleReloadNPCs_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Reload the Server`s NPC
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
104         Call LogGM(.name, .name & " ha recargado los NPCs.")
106         Call CargaNpcsDat
108         Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado exitosamente.", e_FontTypeNames.FONTTYPE_SERVER)
        End With
        Exit Sub
HandleReloadNPCs_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadNPCs", Erl)
End Sub

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)
        On Error GoTo HandleReloadServerIni_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
        'Reload the Server`s INI
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
104         Call LogGM(.name, .name & " ha recargado los INITs.")
106         Call LoadSini
            Call LoadMD5
            Call LoadPrivateKey
        End With
        Exit Sub
HandleReloadServerIni_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadServerIni", Erl)
End Sub

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)
        On Error GoTo HandleReloadSpells_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
104         Call LogGM(.name, .name & " ha recargado los hechizos.")
106         Call CargarHechizos
        End With
        Exit Sub
HandleReloadSpells_Err:
108     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadSpells", Erl)
End Sub

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)
        On Error GoTo HandleReloadObjects_Err
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.RoleMaster)) Then Exit Sub
104         Call LogGM(.name, .name & " ha recargado a los objetos.")
106         Call LoadOBJData
108         Call LoadPesca
110         Call LoadRecursosEspeciales
112         Call WriteConsoleMsg(UserIndex, "Obj.dat recargado exitosamente.", e_FontTypeNames.FONTTYPE_SERVER)
        End With
        Exit Sub
HandleReloadObjects_Err:
114     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReloadObjects", Erl)
End Sub

Public Sub HandleIgnored(ByVal UserIndex As Integer)
        'Author: Lucas Tavolaro Ortiz (Tavo)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) Then
104             .flags.AdminPerseguible = Not .flags.AdminPerseguible
            End If
        End With
        Exit Sub
HandleIgnored_Err:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleIgnored", Erl)
End Sub

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)
        'Author: Pablo (ToxicWaste)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            'Reads the UserName and Slot Packets
            Dim username As String
            Dim Slot     As Byte
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8() 'Que UserName?
104         Slot = Reader.ReadInt8() 'Que Slot?
106         tUser = NameIndex(username)
108         If Not EsGM(UserIndex) Then Exit Sub
110         Call LogGM(.name, .name & " Checkeo el slot " & Slot & " de " & username)
           
112         If IsValidUserRef(tUser) Then
114             If Slot > 0 And Slot <= UserList(UserIndex).CurrentInventorySlots Then
116                 If UserList(tUser.ArrayIndex).Invent.Object(Slot).ObjIndex > 0 Then
118                     Call WriteConsoleMsg(userIndex, " Objeto " & Slot & ") " & ObjData(UserList(tUser.ArrayIndex).Invent.Object(Slot).ObjIndex).name & " Cantidad:" & UserList(tUser.ArrayIndex).Invent.Object(Slot).amount, e_FontTypeNames.FONTTYPE_INFO)
                    Else
120                     Call WriteConsoleMsg(UserIndex, "No hay Objeto en slot seleccionado", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
122                 Call WriteConsoleMsg(UserIndex, "Slot Invílido.", e_FontTypeNames.FONTTYPE_TALK)
                End If
            Else
124             Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_TALK)
            End If
        End With
        Exit Sub
ErrHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCheckSlot", Erl)
End Sub

Public Sub HandleSetSpeed(ByVal UserIndex As Integer)
On Error GoTo HandleGlobalOnOff_Err
        Dim Speed As Single
        Speed = Reader.ReadReal32()
        'Author: Pablo Mercavides
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) = 0 Then Exit Sub
            UserList(UserIndex).Char.speeding = Speed
            Call WriteVelocidadToggle(Speed)
        End With
        Exit Sub
HandleGlobalOnOff_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGlobalOnOff", Erl)
End Sub

Public Sub HandleGlobalMessage(ByVal UserIndex As Integer)
        Dim TActual     As Long
        Dim ElapsedTime As Long

100     TActual = GetTickCount()
102     ElapsedTime = TActual - UserList(UserIndex).Counters.MensajeGlobal
        On Error GoTo ErrHandler
104     With UserList(UserIndex)
            Dim chat As String
106         chat = Reader.ReadString8()
108         If .flags.Silenciado = 1 Then
110             Call WriteLocaleMsg(UserIndex, "110", e_FontTypeNames.FONTTYPE_VENENO, .flags.MinutosRestantes)
112         ElseIf ElapsedTime < IntervaloMensajeGlobal Then
114             Call WriteConsoleMsg(UserIndex, "No puedes escribir mensajes globales tan rápido.", e_FontTypeNames.FONTTYPE_WARNING)
            Else
116             UserList(UserIndex).Counters.MensajeGlobal = TActual

118             If EstadoGlobal Then
120                 If LenB(chat) <> 0 Then
                        ' WyroX: Foto-denuncias - Push message
                        Dim i As Integer
124                     For i = 1 To UBound(.flags.ChatHistory) - 1
126                         .flags.ChatHistory(i) = .flags.ChatHistory(i + 1)
                        Next
128                     .flags.ChatHistory(UBound(.flags.ChatHistory)) = chat
130                     Call modSendData.SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("[" & .name & "] " & chat, e_FontTypeNames.FONTTYPE_GLOBAL))
                        'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
                        'Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbBlue & "í< " & rData & " >í" & CStr(UserList(UserIndex).Char.CharIndex))
                    End If
                Else
132                 Call WriteConsoleMsg(UserIndex, "El global se encuentra Desactivado.", e_FontTypeNames.FONTTYPE_GLOBAL)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
134     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleGlobalOnOff(ByVal UserIndex As Integer)
        On Error GoTo HandleGlobalOnOff_Err
        'Author: Pablo Mercavides
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then Exit Sub
104         Call LogGM(.name, " activo al Chat Global a las " & Now)
106         If EstadoGlobal = False Then
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Chat general habilitado. Escribe" & Chr(34) & "/CONSOLA" & Chr(34) & " o " & Chr(34) & ";" & Chr(34) & " y su mensaje para utilizarlo.", e_FontTypeNames.FONTTYPE_SERVER))
110             EstadoGlobal = True
            Else
112             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Chat General deshabilitado.", e_FontTypeNames.FONTTYPE_SERVER))
114             EstadoGlobal = False
            End If
        End With
        Exit Sub
HandleGlobalOnOff_Err:
116     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGlobalOnOff", Erl)
End Sub

Public Sub HandleDay(ByVal UserIndex As Integer)
        On Error GoTo HandleDay_Err
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         HoraMundo = GetTickCount() - DuracionDia \ 2
108         Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
        End With
        Exit Sub
HandleDay_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDay", Erl)
End Sub

Public Sub HandleSetTime(ByVal UserIndex As Integer)
        On Error GoTo HandleSetTime_Err
100     With UserList(UserIndex)
            Dim HoraDia As Long
102         HoraDia = Reader.ReadInt32
104         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.SemiDios)) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
108         HoraMundo = GetTickCount() - HoraDia
110         Call SendData(SendTarget.ToAll, 0, PrepareMessageHora())
        End With
        Exit Sub
HandleSetTime_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleSetTime", Erl)
End Sub

Public Sub HandleGiveItem(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim username As String
            Dim objIndex As Integer
            Dim Cantidad As Integer
            Dim Motivo   As String
            Dim tUser    As t_UserReference
        
102         username = Reader.ReadString8()
104         objIndex = Reader.ReadInt16()
106         Cantidad = Reader.ReadInt16()
108         Motivo = Reader.ReadString8()
        
110         If (.flags.Privilegios And e_PlayerType.Admin) Then
112             If ObjData(objIndex).Agarrable = 1 Then Exit Sub
114             If Cantidad > MAX_INVENTORY_OBJS Then Cantidad = MAX_INVENTORY_OBJS
                ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
116             If objIndex < 1 Or objIndex > NumObjDatas Then Exit Sub
                ' El nombre del objeto es nulo?
118             If LenB(ObjData(objIndex).name) = 0 Then Exit Sub
                ' Está online?
120             tUser = NameIndex(username)
122             If Not IsValidUserRef(tUser) Then
124                 Call WriteConsoleMsg(UserIndex, "El usuario " & username & " no está conectado.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

                Dim Objeto As t_Obj
126             Objeto.amount = Cantidad
128             Objeto.objIndex = objIndex
                ' Trato de meterlo en el inventario.
130             If MeterItemEnInventario(tUser.ArrayIndex, Objeto) Then
132                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(.name & " ha otorgado a " & UserList(tUser.ArrayIndex).name & " " & Cantidad & " " & ObjData(ObjIndex).name & ": " & Motivo, e_FontTypeNames.FONTTYPE_ROSA))
                Else
134                 Call WriteConsoleMsg(UserIndex, "El usuario no tiene espacio en el inventario.", e_FontTypeNames.FONTTYPE_INFO)
                End If
                ' Lo registro en los logs.
136             Call LogGM(.name, "/DAR " & username & " - Item: " & ObjData(objIndex).name & "(" & objIndex & ") Cantidad : " & Cantidad)
138             Call LogPremios(.name, username, objIndex, Cantidad, Motivo)
            Else
140             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo, debes pedir a un Admin que lo de.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
142     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGiveItem", Erl)
End Sub

Public Sub HandleQuestionGM(ByVal UserIndex As Integer)
        Dim TActual     As Long
        Dim ElapsedTime As Long
        
96      TActual = GetTickCount()
98      ElapsedTime = TActual - UserList(UserIndex).Counters.LastGmMessage
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim Consulta       As String
            Dim TipoDeConsulta As String

102         Consulta = Reader.ReadString8()
104         TipoDeConsulta = Reader.ReadString8()
            Dim PacketCounter As Long
            PacketCounter = Reader.ReadInt32
            Dim Packet_ID As Long
            Packet_ID = PacketNames.QuestionGM
            If Not verifyTimeStamp(PacketCounter, .PacketCounters(Packet_ID), .PacketTimers(Packet_ID), .MacroIterations(Packet_ID), UserIndex, "QuestionGM", PacketTimerThreshold(Packet_ID), MacroIterations(Packet_ID)) Then Exit Sub

            .Counters.CounterGmMessages = .Counters.CounterGmMessages + 1
            If .Counters.CounterGmMessages >= 20 Then
                Dim bannedip As String
                bannedip = UserList(UserIndex).ConnectionDetails.IP
                'Find every player with that ip and ban him!
                Dim i As Long
136             For i = 1 To LastUser
138                 If UserList(i).ConnectionDetails.ConnIDValida Then
140                     If UserList(i).ConnectionDetails.IP = bannedip Then
142                         Call WriteCerrarleCliente(i)
144                         Call CloseSocket(i)
                        End If
                    End If
146             Next i
            End If
            
            If ElapsedTime < IntervaloConsultaGM Then
115             Call WriteConsoleMsg(UserIndex, "Solo puedes enviar una consulta cada 5 minutos.", e_FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
            End If
            UserList(UserIndex).Counters.LastGmMessage = TActual
112         Call Ayuda.Push(.name, Consulta, TipoDeConsulta)
114         Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Se ha recibido un nuevo mensaje de soporte de " & UserList(UserIndex).name & ".", e_FontTypeNames.FONTTYPE_SERVER))
            .Counters.CounterGmMessages = 0
116         Call WriteConsoleMsg(UserIndex, "Tu mensaje fue recibido por el equipo de soporte.", e_FontTypeNames.FONTTYPE_INFOIAO)
118         Call LogConsulta(.name & " (" & TipoDeConsulta & ") " & Consulta)
        End With
        Exit Sub
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleCuentaRegresiva(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim Seconds As Byte
102         Seconds = Reader.ReadInt8()
104         If Not .flags.Privilegios And e_PlayerType.user Then
106             CuentaRegresivaTimer = Seconds
108             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("¡Empezando cuenta regresiva desde: " & Seconds & " segundos...!", e_FontTypeNames.FONTTYPE_GUILD))
            End If
        End With
        Exit Sub
ErrHandler:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCuentaRegresiva", Erl)
End Sub

Public Sub HandlePossUser(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
102         username = Reader.ReadString8()
            If Not InStrB(username, "(") Then Exit Sub
            'Modifico la forma en que se usa el destrabar, ahora solamente lo puedo destrabar si está online.
104         If (.flags.Privilegios And (e_PlayerType.user)) = 0 Then
                Dim tUser As t_UserReference
                username = Split(username, "(")(0)
                tUser = NameIndex(username)
106             If IsValidUserRef(tUser) Then
                    If LegalPosDestrabar(UserList(tUser.ArrayIndex).pos.map, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y, .flags.Navegando = 1, .flags.Navegando = 0) Then Exit Sub
                    Dim nPos As t_WorldPos
                    Call ClosestLegalPos(UserList(tUser.ArrayIndex).pos, nPos, False, True)
150                 Call FindLegalPos(tUser.ArrayIndex, UserList(tUser.ArrayIndex).pos.map, CByte(UserList(tUser.ArrayIndex).pos.x), CByte(UserList(tUser.ArrayIndex).pos.y))
152                 Call WarpUserChar(tUser.ArrayIndex, nPos.map, nPos.x, nPos.y, True)
116                 Call WriteConsoleMsg(UserIndex, "Servidor » Acción realizada con exito! La nueva posicion de " & username & " es: " & UserList(tUser.ArrayIndex).Pos.Map & "-" & UserList(tUser.ArrayIndex).Pos.X & "-" & UserList(tUser.ArrayIndex).Pos.Y & ".", e_FontTypeNames.FONTTYPE_INFO)
                    'ver porque si el usuario esta online lo dice igual
                Else
118                 Call WriteConsoleMsg(UserIndex, "Servidor » El usuario debe estar deslogueado para dicha solicitud!", e_FontTypeNames.FONTTYPE_INFO)
                    'avisar al usuario que un gm lo destrabara pero debe desloguear
                End If
            Else
120             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandlePossUser", Erl)
End Sub

Public Sub HandleNieveToggle(ByVal UserIndex As Integer)
        On Error GoTo HandleNieveToggle_Err
        'Author: Pablo Mercavides
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         Call LogGM(.name, "/NIEVE")
108         Nebando = Not Nebando
110         Call SendData(SendTarget.ToAll, 0, PrepareMessageNevarToggle())
        End With
        Exit Sub
HandleNieveToggle_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNieveToggle", Erl)
 
End Sub

Public Sub HandleNieblaToggle(ByVal UserIndex As Integer)
        On Error GoTo HandleNieblaToggle_Err
        'Author: Pablo Mercavides
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
106         Call LogGM(.name, "/NIEBLA")
108         Call ResetMeteo
        End With
        Exit Sub
HandleNieblaToggle_Err:
110     Call TraceError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
End Sub

Public Sub HandleGenio(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        On Error GoTo HandleGenio_Err
100     With UserList(UserIndex)
            'Si no es GM, no pasara nada :P
102         If (.flags.Privilegios And e_PlayerType.user) Then Exit Sub
            Dim i As Byte
104         For i = 1 To NUMSKILLS
106             .Stats.UserSkills(i) = 100
108         Next i
110         Call WriteConsoleMsg(UserIndex, "Tus skills fueron editados.", e_FontTypeNames.FONTTYPE_INFOIAO)
        End With
        Exit Sub
HandleGenio_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleGenio", Erl)
End Sub

Public Sub HandleBanCuenta(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
            Dim Reason   As String
102         username = Reader.ReadString8()
104         Reason = Reader.ReadString8()
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
108             Call BanearCuenta(UserIndex, username, Reason)
            Else
110             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBanCuenta", Erl)
End Sub

Public Sub HandleUnBanCuenta(ByVal UserIndex As Integer)
        ' /unbancuenta namepj
        ' /unbancuenta email
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim UserNameOEmail As String
102         UserNameOEmail = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
116             If DesbanearCuenta(UserIndex, UserNameOEmail) Then
118                 Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » " & .Name & " ha desbaneado la cuenta de " & UserNameOEmail & ".", e_FontTypeNames.FONTTYPE_SERVER))
                Else
                    Call WriteConsoleMsg(UserIndex, "No se ha podido desbanear la cuenta.", e_FontTypeNames.FONTTYPE_INFO)
                End If
            Else
120             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
122     Call TraceError(Err.Number, Err.Description, "Protocol.HandleUnBanCuenta", Erl)
End Sub

Public Sub HandleCerrarCliente(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)

        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
102         username = Reader.ReadString8()
            ' Solo administradores pueden cerrar clientes ajenos
104         If (.flags.Privilegios And e_PlayerType.Admin) Then
106             tUser = NameIndex(username)
108             If Not IsValidUserRef(tUser) Then
110                 Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", e_FontTypeNames.FONTTYPE_INFO)
                Else
112                 Call SendData(SendTarget.ToAdminsYDioses, 0, PrepareMessageConsoleMsg(.name & " cerro el cliente de " & username & ".", e_FontTypeNames.FONTTYPE_INFO))
114                 Call WriteCerrarleCliente(tUser.ArrayIndex)
116                 Call LogGM(.name, "Cerro el cliene de:" & username)
                End If
            End If
        End With
        Exit Sub
ErrHandler:
118     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCerrarCliente", Erl)
End Sub

Public Sub HandleEventoInfo(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        On Error GoTo HandleEventoInfo_Err

100     With UserList(UserIndex)
102         If EventoActivo Then
104             Call WriteConsoleMsg(UserIndex, PublicidadEvento & ". Tiempo restante: " & TiempoRestanteEvento & " minuto(s).", e_FontTypeNames.FONTTYPE_New_Eventos)
            Else
106             Call WriteConsoleMsg(UserIndex, "Eventos> Actualmente no hay ningun evento en curso.", e_FontTypeNames.FONTTYPE_New_Eventos)
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
134             Call WriteConsoleMsg(UserIndex, "Eventos> El proximo evento " & DescribirEvento(HoraProximo) & " iniciara a las " & HoraProximo & ":00 horas.", e_FontTypeNames.FONTTYPE_New_Eventos)
            Else
136             Call WriteConsoleMsg(UserIndex, "Eventos> No hay eventos proximos.", e_FontTypeNames.FONTTYPE_New_Eventos)
            End If
        End With
        Exit Sub
HandleEventoInfo_Err:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEventoInfo", Erl)
End Sub

Public Sub HandleCrearEvento(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
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
118                     Call WriteConsoleMsg(UserIndex, "Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.", e_FontTypeNames.FONTTYPE_New_Eventos)
                    Else
120                     Call ForzarEvento(Tipo, Duracion, multiplicacion, UserList(UserIndex).name)
                    End If
                Else
122                 Call WriteConsoleMsg(UserIndex, "Ya hay un evento en curso. Finalicelo con /FINEVENTO primero.", e_FontTypeNames.FONTTYPE_New_Eventos)
                End If
            Else
124             Call WriteConsoleMsg(UserIndex, "Servidor » Solo Administradores pueder crear estos eventos.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleBanTemporal(ByVal UserIndex As Integer)
        'Author: Nicolas Matias Gonzalez (NIGO)

        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
            Dim Reason   As String
            Dim dias     As Byte
        
102         username = Reader.ReadString8()
104         Reason = Reader.ReadString8()
106         dias = Reader.ReadInt8()
        
108         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
110             Call Admin.BanTemporal(username, dias, Reason, UserList(UserIndex).name)
            Else
112             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
114     Call TraceError(Err.Number, Err.Description, "Protocol.?", Erl)
End Sub

Public Sub HandleCancelarExit(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        On Error GoTo HandleCancelarExit_Err
100     Call CancelExit(UserIndex)
        Exit Sub

HandleCancelarExit_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCancelarExit", Erl)
End Sub

Public Sub HandleCrearTorneo(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim NivelMinimo As Byte
            Dim NivelMaximo As Byte
        
            Dim cupos       As Byte
            Dim costo       As Long
        
            Dim mago        As Byte
            Dim clerico     As Byte
            Dim guerrero    As Byte
            Dim asesino     As Byte
            Dim bardo       As Byte
            Dim druido      As Byte
            Dim Paladin     As Byte
            Dim cazador     As Byte
            Dim Trabajador  As Byte
            Dim Pirata      As Byte
            Dim Ladron      As Byte
            Dim Bandido     As Byte

            Dim Mapa        As Integer
            Dim X           As Byte
            Dim y           As Byte

            Dim nombre      As String
            Dim reglas      As String

102         NivelMinimo = Reader.ReadInt8
104         NivelMaximo = Reader.ReadInt8
        
106         cupos = Reader.ReadInt8
108         costo = Reader.ReadInt32
        
110         mago = Reader.ReadInt8
112         clerico = Reader.ReadInt8
114         guerrero = Reader.ReadInt8
116         asesino = Reader.ReadInt8
118         bardo = Reader.ReadInt8
120         druido = Reader.ReadInt8
122         Paladin = Reader.ReadInt8
124         cazador = Reader.ReadInt8
126         Trabajador = Reader.ReadInt8
128         Pirata = Reader.ReadInt8
130         Ladron = Reader.ReadInt8
132         Bandido = Reader.ReadInt8

134         Mapa = Reader.ReadInt16
136         X = Reader.ReadInt8
138         y = Reader.ReadInt8
        
140         nombre = Reader.ReadString8
142         reglas = Reader.ReadString8
  
144         If EsGM(UserIndex) And ((.flags.Privilegios And e_PlayerType.Consejero) = 0) Then
146             Torneo.NivelMinimo = NivelMinimo
148             Torneo.NivelMaximo = NivelMaximo
            
150             Torneo.cupos = cupos
152             Torneo.costo = costo
            
154             Torneo.mago = mago
156             Torneo.clerico = clerico
158             Torneo.guerrero = guerrero
160             Torneo.asesino = asesino
162             Torneo.bardo = bardo
164             Torneo.druido = druido
166             Torneo.Paladin = Paladin
168             Torneo.cazador = cazador
170             Torneo.Trabajador = Trabajador
172             Torneo.Pirata = Pirata
174             Torneo.Ladron = Ladron
176             Torneo.Bandido = Bandido
        
178             Torneo.Mapa = Mapa
180             Torneo.X = X
182             Torneo.y = y
            
184             Torneo.nombre = nombre
186             Torneo.reglas = reglas

188             Call IniciarTorneo
            End If
        End With
        Exit Sub

ErrHandler:
190     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCrearTorneo", Erl)
End Sub

Public Sub HandleComenzarTorneo(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
102         If EsGM(UserIndex) Then
104             Call ComenzarTorneoOk
            End If
        End With
        Exit Sub
ErrHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
End Sub

Public Sub HandleCancelarTorneo(ByVal UserIndex As Integer)

        On Error GoTo ErrHandler
100     With UserList(UserIndex)
102         If EsGM(UserIndex) Then
104             Call ResetearTorneo
            End If
        End With
        Exit Sub

ErrHandler:
106     Call TraceError(Err.Number, Err.Description, "Protocol.HandleComenzarTorneo", Erl)
End Sub

Public Sub HandleBusquedaTesoro(ByVal UserIndex As Integer)

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
116                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & TesoroNumMapa & "-" & TesoroX & "-" & TesoroY, e_FontTypeNames.FONTTYPE_INFO)
                            Else
118                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
120                 Case 1
122                     If Not BusquedaRegaloActiva And BusquedaTesoroActiva = False And BusquedaNpcActiva = False Then
124                         Call PerderRegalo
                        Else
126                         If BusquedaRegaloActiva Then
128                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Ningún valiente fue capaz de encontrar el item misterioso, recuerda que se encuentra en " & get_map_name(RegaloNumMapa) & "(" & RegaloNumMapa & "). ¡Ten cuidado!", e_FontTypeNames.FONTTYPE_TALK))
130                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa. El tesoro se encuentra en: " & RegaloNumMapa & "-" & RegaloX & "-" & RegaloY, e_FontTypeNames.FONTTYPE_INFO)
                            Else
132                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
134                 Case 2
136                     If Not BusquedaNpcActiva And BusquedaTesoroActiva = False And BusquedaRegaloActiva = False Then
                            Dim Pos As t_WorldPos
138                         Pos.map = TesoroNPCMapa(RandomNumber(1, UBound(TesoroNPCMapa)))
140                         Pos.y = 50
142                         Pos.X = 50
144                         npc_index_evento = SpawnNpc(TesoroNPC(RandomNumber(1, UBound(TesoroNPC))), Pos, True, False, True)
146                         BusquedaNpcActiva = True
                        Else
148                         If BusquedaNpcActiva Then
150                             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> Todavía nadie logró matar el NPC que se encuentra en el mapa " & NpcList(npc_index_evento).Pos.Map & ".", e_FontTypeNames.FONTTYPE_TALK))
152                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda de npc activo. El tesoro se encuentra en: " & NpcList(npc_index_evento).Pos.map & "-" & NpcList(npc_index_evento).Pos.X & "-" & NpcList(npc_index_evento).Pos.y, e_FontTypeNames.FONTTYPE_INFO)
                            Else
154                             Call WriteConsoleMsg(UserIndex, "Ya hay una busqueda del tesoro activa.", e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                End Select
            Else
156             Call WriteConsoleMsg(UserIndex, "Servidor » No estas habilitado para hacer Eventos.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
158     Call TraceError(Err.Number, Err.Description, "Protocol.HandleBusquedaTesoro", Erl)
End Sub

Public Sub HandleMarcaDeGM(ByVal UserIndex As Integer)
        'Author: Pablo Mercavides
        On Error GoTo HandleMarcaDeGM_Err
100     Call WriteWorkRequestTarget(UserIndex, e_Skill.MarcaDeGM)
        Exit Sub

HandleMarcaDeGM_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleMarcaDeGM", Erl)
End Sub

Public Sub HandleFinEvento(ByVal UserIndex As Integer)
        On Error GoTo HandleDenounce_Err
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
100     With UserList(UserIndex)
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Consejero Or e_PlayerType.RoleMaster)) Then
104             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
106         If EventoActivo Then
108             Call FinalizarEvento
            Else
110             Call WriteConsoleMsg(UserIndex, "No hay ningun evento activo.", e_FontTypeNames.FONTTYPE_New_Eventos)
            End If
        End With
        Exit Sub
HandleDenounce_Err:
112     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDenounce", Erl)
End Sub

Public Sub HandleCreateEvent(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler

100     With UserList(UserIndex)
            Dim name As String
102         name = Reader.ReadString8()
104         If LenB(name) = 0 Then Exit Sub
106         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
108             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
110         Select Case UCase$(name)
                Case "INVASION BANDER"
112                 Call IniciarEvento(TipoEvento.Invasion, 1)
114                 Call LogGM(.Name, "Forzó el evento Invasión en Banderbille.")
116             Case "INVASION CARCEL"
118                 Call IniciarEvento(TipoEvento.Invasion, 2)
120                 Call LogGM(.Name, "Forzó el evento Invasión en Carcel.")
122             Case Else
124                 Call WriteConsoleMsg(UserIndex, "No existe el evento """ & name & """.", e_FontTypeNames.FONTTYPE_INFO)
            End Select
        End With
        Exit Sub
ErrHandler:
126     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCreateEvent", Erl)
End Sub

Public Sub HandleStartEvent(ByVal UserIndex As Integer)
On Error GoTo ErrHandler
    Dim eventType As Byte
    eventType = Reader.ReadInt8()
    Dim LobbySettings As t_NewScenearioSettings
    LobbySettings.ScenearioType = Reader.ReadInt8
    LobbySettings.MinLevel = Reader.ReadInt8
    LobbySettings.MaxLevel = Reader.ReadInt8
    LobbySettings.MinPlayers = Reader.ReadInt8
    LobbySettings.MaxPlayers = Reader.ReadInt8
    LobbySettings.TeamSize = Reader.ReadInt8
    LobbySettings.TeamType = Reader.ReadInt8
    LobbySettings.RoundNumber = Reader.ReadInt8
    LobbySettings.InscriptionFee = Reader.ReadInt32
    LobbySettings.Description = Reader.ReadString8
    LobbySettings.Password = Reader.ReadString8
    If eventType = 0 Then
        CurrentActiveEventType = LobbySettings.ScenearioType
        Select Case LobbySettings.ScenearioType
            Case e_EventType.CaptureTheFlag
                Call HandleIniciarCaptura(UserIndex, LobbySettings)
            Case Else
                Call HandleStartGenericLobby(UserIndex, LobbySettings)
        End Select
    Else
        With UserList(UserIndex)
            If IsValidNpcRef(.flags.TargetNPC) Then
                If NpcList(.flags.TargetNPC.ArrayIndex).npcType = e_NPCType.EventMaster And .flags.Muerto = 0 Then
                    Call CreatePublicEvent(UserIndex, LobbySettings)
                End If
            End If
        End With
    End If
    Exit Sub
ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleStartEvent", Erl)
End Sub

Public Sub HandleCancelarEvento(ByVal UserIndex As Integer)
    On Error GoTo ErrHandler
    If (UserList(userIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) = 0 Then
        Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If CurrentActiveEventType = CaptureTheFlag Then
        If InstanciaCaptura Is Nothing Then
            Call WriteConsoleMsg(UserIndex, "Eventos » No hay ninguna instancia en curso para ese evento.", e_FontTypeNames.FONTTYPE_INFO)
        Else
            Call InstanciaCaptura.finalizarCaptura
        End If
    Else
        If GlobalLobbyIndex >= 0 Then
            Call CancelLobby(LobbyList(GlobalLobbyIndex))
            If LobbyList(GlobalLobbyIndex).Scenario Is Nothing Then Call ReleaseLobby(GlobalLobbyIndex)
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» El evento ha sido cancelado.", e_FontTypeNames.FONTTYPE_GUILD))
        Else
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("No se encontro ningun evento activo.", e_FontTypeNames.FONTTYPE_GUILD))
        End If
    End If
    Exit Sub
ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleCancelarCaptura", Erl)
End Sub

Public Sub HandleSeguirMouse(ByVal UserIndex As Integer)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)
            Dim username As String
            Dim tUser    As t_UserReference
            Dim LoopC    As Byte
            
            Dim tempArea As Long
102         username = Reader.ReadString8()
        
104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
106             If UCase$(username) <> "YO" Then
108                 tUser = NameIndex(username)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes seguirte a vos mismo", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
109             Call LogGM(.name, "/SM " & username)
                If username = "" And .flags.SigueUsuario.ArrayIndex > 0 Then
                    'Me devuelvo inventario y stats
                    UserList(UserIndex).Invent = UserList(UserIndex).Invent_bk
                    UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
                    'UserList(UserIndex).Char.charindex = UserList(UserIndex).Char.charindex_bk
                    Call WriteConsoleMsg(userIndex, "Dejas de seguir a " & UserList(.flags.SigueUsuario.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteCancelarSeguimiento(UserIndex)
                    Call WriteNotificarClienteSeguido(.flags.SigueUsuario.ArrayIndex, 0)
                    Call SetUserRef(UserList(.flags.SigueUsuario.ArrayIndex).flags.GMMeSigue, 0)
                    Call WriteUserCharIndexInServer(UserIndex)
                    Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y, False)
                    Call SetUserRef(UserList(userIndex).flags.SigueUsuario, 0)
                ElseIf tUser.ArrayIndex <> .flags.SigueUsuario.ArrayIndex And .flags.SigueUsuario.ArrayIndex > 0 And IsValidUserRef(tUser) Then
                    If IsValidUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue) And UserList(tUser.ArrayIndex).flags.GMMeSigue.ArrayIndex <> userIndex Then
                        Call WriteConsoleMsg(UserIndex, "El usuario está siendo seguido por " & UserList(UserList(tUser.ArrayIndex).flags.GMMeSigue.ArrayIndex).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Me devuelvo inventario y stats
                    UserList(UserIndex).Invent = UserList(UserIndex).Invent_bk
                    UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
                    'UserList(UserIndex).Char.charindex = UserList(UserIndex).Char.charindex_bk
                    Call WriteConsoleMsg(userIndex, "Dejas de seguir a " & UserList(.flags.SigueUsuario.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
                    Call WriteCancelarSeguimiento(UserIndex)
                    Call WriteNotificarClienteSeguido(.flags.SigueUsuario.ArrayIndex, 0)
                    Call SetUserRef(UserList(.flags.SigueUsuario.ArrayIndex).flags.GMMeSigue, 0)
                    Call WriteUserCharIndexInServer(UserIndex)
                    Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y, False)
                    Call SetUserRef(UserList(userIndex).flags.SigueUsuario, 0)
                    
                    Call DoAdminInvisible(UserIndex, 1) 'HACER LO DE ADMIN INVISIBLE
                    'Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, UserList(tUser).Pos.X, UserList(tUser).Pos.y, False)
                    
                    'Me backupeo el inventario y los hechizos
                    UserList(UserIndex).Invent_bk = UserList(UserIndex).Invent
                    UserList(UserIndex).Stats_bk = UserList(UserIndex).Stats
                    
                    'Me pego el inventario y los hechizos del usuario seguido
                    UserList(userIndex).Invent = UserList(tUser.ArrayIndex).Invent
                    UserList(userIndex).Stats = UserList(tUser.ArrayIndex).Stats
                    UserList(UserIndex).flags.SigueUsuario = tUser
                    'Actualizo flag en cliente para que empiece a enviar paquetes
                    Call WriteNotificarClienteSeguido(tUser.ArrayIndex, 1)
                    Call SetUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue, userIndex)
                    Call WriteConsoleMsg(userIndex, "Comienzas a seguir a " & UserList(tUser.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
                    tempArea = UserList(UserIndex).AreasInfo.AreaID
                    Call WarpUserChar(userIndex, UserList(tUser.ArrayIndex).pos.map, 15, 15)
                    Call WriteSendFollowingCharindex(userIndex, UserList(tUser.ArrayIndex).Char.charindex)
                    If tempArea <> UserList(tUser.ArrayIndex).AreasInfo.AreaID Then
                        Call MakeUserChar(False, userIndex, tUser.ArrayIndex, UserList(tUser.ArrayIndex).pos.map, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y)
                    End If
                    
                    Call WritePosUpdateCharIndex(userIndex, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y, UserList(tUser.ArrayIndex).Char.charindex)
                    Call WriteUserCharIndexInServer(tUser.ArrayIndex)
                    UserList(userIndex).AreasInfo = UserList(tUser.ArrayIndex).AreasInfo
                    Call ModAreas.CheckUpdateNeededUser(UserIndex, 5, 0)
112             ElseIf Not IsValidUserRef(tUser) Then
114                 Call WriteConsoleMsg(UserIndex, "Usuario offline.", e_FontTypeNames.FONTTYPE_INFO)
                Else
                    'Si empiezo a seguir a alguien
                    If Not IsValidUserRef(UserList(userIndex).flags.SigueUsuario) Then
                        If IsValidUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue) And UserList(tUser.ArrayIndex).flags.GMMeSigue.ArrayIndex <> userIndex Then
                            Call WriteConsoleMsg(UserIndex, "El usuario está siendo seguido por " & UserList(UserList(tUser.ArrayIndex).flags.GMMeSigue.ArrayIndex).Name & ".", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoAdminInvisible(UserIndex, 1) 'HACER LO DE ADMIN INVISIBLE
                        'Call WarpUserChar(UserIndex, UserList(tUser).Pos.map, UserList(tUser).Pos.X, UserList(tUser).Pos.y, False)
                        
                        'Me backupeo el inventario y los hechizos
                        UserList(UserIndex).Invent_bk = UserList(UserIndex).Invent
                        UserList(UserIndex).Stats_bk = UserList(UserIndex).Stats
                        'Me pego el inventario y los hechizos del usuario seguido
                        UserList(userIndex).Invent = UserList(tUser.ArrayIndex).Invent
                        UserList(userIndex).Stats = UserList(tUser.ArrayIndex).Stats
                        UserList(UserIndex).flags.SigueUsuario = tUser
                        'Actualizo flag en cliente para que empiece a enviar paquetes
                        Call WriteNotificarClienteSeguido(tUser.ArrayIndex, 1)
                        Call SetUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue, userIndex)
                        Call WriteConsoleMsg(userIndex, "Comienzas a seguir a " & UserList(tUser.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
                        tempArea = UserList(UserIndex).AreasInfo.AreaID
                        Call WarpUserChar(userIndex, UserList(tUser.ArrayIndex).pos.map, 15, 15)
                        Call WriteSendFollowingCharindex(userIndex, UserList(tUser.ArrayIndex).Char.charindex)
                        If tempArea <> UserList(tUser.ArrayIndex).AreasInfo.AreaID Then
                            Call MakeUserChar(False, userIndex, tUser.ArrayIndex, UserList(tUser.ArrayIndex).pos.map, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y)
                        End If
                        
                        Call WritePosUpdateCharIndex(userIndex, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y, UserList(tUser.ArrayIndex).Char.charindex)
                        Call WriteUserCharIndexInServer(tUser.ArrayIndex)
                        UserList(userIndex).AreasInfo = UserList(tUser.ArrayIndex).AreasInfo
                        Call ModAreas.CheckUpdateNeededUser(UserIndex, 5, 0)
                        Call WriteAreaChanged(UserIndex, UserList(tUser.ArrayIndex).pos.x, UserList(tUser.ArrayIndex).pos.y)
                    Else
                        If UserList(userIndex).flags.SigueUsuario.ArrayIndex <> tUser.ArrayIndex Then
                            Call WriteConsoleMsg(userIndex, "Ya te encuentras siguiendo a un usuario, para dejar de seguirlo escribe /SM " & UserList(UserList(userIndex).flags.SigueUsuario.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        'Me devuelvo inventario y stats
                        UserList(UserIndex).Invent = UserList(UserIndex).Invent_bk
                        UserList(UserIndex).Stats = UserList(UserIndex).Stats_bk
                        Call SetUserRef(UserList(userIndex).flags.SigueUsuario, 0)
                        Call WriteConsoleMsg(userIndex, "Dejas de seguir a " & UserList(tUser.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
                        Call WriteCancelarSeguimiento(UserIndex)
                        Call WriteNotificarClienteSeguido(tUser.ArrayIndex, 0)
                        Call SetUserRef(UserList(tUser.ArrayIndex).flags.GMMeSigue, 0)
                        Call WriteUserCharIndexInServer(UserIndex)
                        Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y, False)
                    End If
                End If
            Else
136             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        
        Call UpdateUserInv(True, UserIndex, 1)
        Call UpdateUserHechizos(True, UserIndex, 0)
900     Call WriteUpdateUserStats(UserIndex)
        Exit Sub
ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleReviveChar", Erl)
End Sub

Public Sub HandleEventoFaccionario(ByVal UserIndex As Integer)
    On Error GoTo HandleEventoFaccionario_Err:
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) Then
        EnEventoFaccionario = Not EnEventoFaccionario
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor » Se ha " & IIf(EnEventoFaccionario, "iniciado", "finalizado") & " el evento faccionario.", e_FontTypeNames.FONTTYPE_SERVER))
    End If
    Exit Sub
HandleEventoFaccionario_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleEventoFaccionario", Erl)
End Sub

Public Sub HandleDebugRequest(ByVal UserIndex As Integer)
On Error GoTo HandleDebugRequest_Err:
    Dim debugType As Byte
    Dim Args() As String
    debugType = Reader.ReadInt8()
    If debugType = 0 Then
        ReDim Args(0) As String
    ElseIf debugType = 1 Then
        ReDim Args(1) As String
        Args(0) = Reader.ReadString8()
    ElseIf debugType = 2 Then
        ReDim Args(0) As String
    Else
        Exit Sub
    End If
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios Or e_PlayerType.Consejero)) Then
        Call WriteDebugLogResponse(UserIndex, debugType, Args, UBound(Args))
    End If
    
    Exit Sub
HandleDebugRequest_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleDebugRequest", Erl)
End Sub

Public Sub HandleLobbyCommand(ByVal UserIndex As Integer)
On Error GoTo HandleLobbyCommand_err
    Dim Command As Byte
    Dim hasPermission As Integer
    Dim Params As String
    Command = Reader.ReadInt8()
    Params = Reader.ReadString8()
100 With UserList(UserIndex)
102     If .flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios) Then
104         If Not HandleRemoteLobbyCommand(Command, Params, UserIndex, GlobalLobbyIndex) Then
106             Call WriteConsoleMsg(UserIndex, "Servidor » No se pudo procesar el comando.", e_FontTypeNames.FONTTYPE_INFO)
108         End If
110     Else
112         Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
114     End If
116 End With
    Exit Sub
HandleLobbyCommand_err:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleLobbyCommand", Erl)

End Sub

Public Sub HandleFeatureToggle(ByVal UserIndex As Integer)
On Error GoTo HandleFeatureToggle_Err:
    Dim value As Byte
    Dim name As String
    Dim nameSize As Integer
    value = max(Min(1, Reader.ReadInt8), 0)
    name = Reader.ReadString8
    nameSize = Len(nameSize)
    If nameSize = 0 Or nameSize > 100 Then
        Exit Sub
    End If
    If (UserList(UserIndex).flags.Privilegios And (e_PlayerType.Admin)) Then
        Call SetFeatureToggle(name, value > 0)
        Call WriteConsoleMsg(UserIndex, "variable configurada correctamente.", e_FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "no tienes permisos para realizar esta accion.", e_FontTypeNames.FONTTYPE_INFO)
    End If
    Exit Sub
HandleFeatureToggle_Err:
102     Call TraceError(Err.Number, Err.Description, "Protocol.HandleFeatureToggle", Erl)
End Sub

'HarThaoS: Iniciar captura de bandera
Public Sub HandleIniciarCaptura(ByVal UserIndex As Integer, EventSettings As t_NewScenearioSettings)
        On Error GoTo ErrHandler
100     With UserList(UserIndex)

104         If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) Then
                If Not InstanciaCaptura Is Nothing Then
                    Call WriteConsoleMsg(UserIndex, "Ya hay un evento de captura de bandera en curso.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                Else
                    'El precio no puede ser negativo
                    If EventSettings.InscriptionFee < 0 Then
                        Call WriteConsoleMsg(UserIndex, "El valor de la entrada al evento no podrá ser menor que 0.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    'Me fijo si que la cantidad de participantes sea par
                    If EventSettings.MaxPlayers Mod 2 <> 0 Then
                        Call WriteConsoleMsg(UserIndex, "La cantidad de participantes debe ser un número par.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Permito un máximo de 48 participantes
                    If EventSettings.MaxPlayers > 48 Then 'Leer de una variable de configuración
                        Call WriteConsoleMsg(UserIndex, "La cantidad de participantes no podrá ser mayor que 48.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Me fijo si hay más participantes conectados que el cupo para jugar
                    If EventSettings.MaxPlayers > NumUsers Then
                        Call WriteConsoleMsg(UserIndex, "Hay pocos jugadores en el servidor, intenta con una cantidad menor de participantes.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If EventSettings.MinLevel < 1 Or EventSettings.MinLevel > 47 Then
                        Call WriteConsoleMsg(UserIndex, "El nivel para el evento debe ser entre 1 y 47.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If EventSettings.MinLevel > EventSettings.MaxLevel Then
                        Call WriteConsoleMsg(UserIndex, "El nivel minimo debe ser menor al maximo.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    Set InstanciaCaptura = New clsCaptura
                    Call InstanciaCaptura.inicializar(EventSettings.MaxPlayers, EventSettings.RoundNumber, EventSettings.MinLevel, EventSettings.MaxLevel, EventSettings.InscriptionFee)
                End If
            Else
136             Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub
ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleIniciarCaptura", Erl)
End Sub

Private Sub HandleStartGenericLobby(ByVal UserIndex As Integer, ByRef LobbySettings As t_NewScenearioSettings)
On Error GoTo ErrHandler

    If IsEventActive Then
        Call WriteConsoleMsg(UserIndex, "Ya hay un evento activo, debes cancelarlo primero.", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    GlobalLobbyIndex = GetAvailableLobby()
    If GlobalLobbyIndex < 0 Then
        Call WriteConsoleMsg(UserIndex, "No se pudo encontrar una sala activa.", e_FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    With UserList(UserIndex)
        If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios Or e_PlayerType.SemiDios)) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor » Comando deshabilitado para tu cargo.", e_FontTypeNames.FONTTYPE_INFO)
        Else
136         'Me fijo si hay más participantes conectados que el cupo para jugar
            If Not ValidateLobbySettings(UserIndex, LobbySettings) Then
                Exit Sub
            End If
            Call InitializeLobby(LobbyList(GlobalLobbyIndex))
            LobbyList(GlobalLobbyIndex).IsGlobal = True
            Call ModLobby.SetMinLevel(LobbyList(GlobalLobbyIndex), LobbySettings.MinLevel)
            Call ModLobby.SetMaxLevel(LobbyList(GlobalLobbyIndex), LobbySettings.MaxLevel)
            Call ModLobby.SetMaxPlayers(LobbyList(GlobalLobbyIndex), LobbySettings.MaxPlayers)
            Call CustomScenarios.PrepareNewEvent(LobbySettings.ScenearioType, GlobalLobbyIndex)
            Call WriteConsoleMsg(UserIndex, "Se creo el lobby, recorda que tenes que abrirlo para que se pueda anotar gente.", e_FontTypeNames.FONTTYPE_INFO)
            Call LogGM(.name, "Inicio un Lobby")
        End If
    End With
    Exit Sub
ErrHandler:
138     Call TraceError(Err.Number, Err.Description, "Protocol.HandleStartGenericLobby", Erl)
End Sub
