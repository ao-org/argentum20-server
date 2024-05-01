VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private participantes As Collection

Private cantidad_rondas As Byte
Private nivel_minimo As Byte
Private nivel_maximo As Byte
Private max_participantes As Byte
Private tiempo_espera As Integer
Private tiempo_ronda As Integer
Private precio As Long
Private ronda_actual As Byte
Private reintentos As Byte

Private Enum estadoCaptura
    Inscripcion = 1
    IniciaRonda = 2
    EnJuego = 3
    Finalizado = 4
End Enum

Private Estado As estadoCaptura

Private contadorBandera(1 To 2) As Integer

Private rondasGanadas(1 To 2) As Integer


Private colorEquipo(1 To 2) As String
Private mapaEquipo(1 To 2) As Integer
Private xEquipo(1 To 2) As Byte
Private yEquipo(1 To 2) As Byte
Private objEquipo(1 To 2) As Integer
Private Sub Class_Initialize()
    Set participantes = New Collection
    
    colorEquipo(1) = "azul"
    colorEquipo(2) = "rojo"
    
    mapaEquipo(1) = MAP_TEAM_1
    mapaEquipo(2) = MAP_TEAM_2
    
    xEquipo(1) = X_BANDERA_1
    xEquipo(2) = X_BANDERA_2
    yEquipo(1) = Y_BANDERA_1
    yEquipo(2) = Y_BANDERA_2
    
    objEquipo(1) = OBJ_CAPTURA_BANDERA_1
    objEquipo(2) = OBJ_CAPTURA_BANDERA_2
End Sub

Public Function GetPlayer(ByVal UserIndex As Integer) As clsCapturaParticipante
    Set GetPlayer = Participantes.Item(CStr(UserIndex))
End Function

Public Sub eliminarParticipante(ByRef player As clsCapturaParticipante)
    If player Is Nothing Then
        Exit Sub
    End If
    If player.IsValid Then
        Call tirarBandera(player.PlayerIndex)
    
        UserList(player.PlayerIndex).flags.jugando_captura = 0
        UserList(player.PlayerIndex).flags.CurrentTeam = 0
        
        Call WarpUserChar(player.PlayerIndex, player.LastPosMap, player.LastPosX, player.LastPosY)
    End If
    Call Participantes.Remove(CStr(player.PlayerIndex))
End Sub

Public Sub PasarSegundo()

    Dim participante As clsCapturaParticipante
    If Estado = Inscripcion Then
        tiempo_espera = tiempo_espera - 1
        If tiempo_espera Mod 30 = 0 And tiempo_espera > 5 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» En " & tiempo_espera & " segundos se cierran las inscripciones para el evento captura de bandera, escribe /PARTICIPAR para anotarte.", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_espera <= 5 And tiempo_espera > 0 Then
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Cerramos inscripciones en " & tiempo_espera & ".", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_espera = 0 Then
            If participantes.Count < max_participantes / 2 Then
                reintentos = reintentos + 1
                If reintentos >= 5 Then
                    Call finalizarCaptura(-2)
                Else
                    tiempo_espera = 60
                    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Se extendieron las inscripciones por falta de participantes.", e_FontTypeNames.FONTTYPE_GUILD))
                End If
            Else
                Call iniciarCaptura
            End If
        End If
    ElseIf Estado = IniciaRonda Then
        
        tiempo_ronda = tiempo_ronda - 1
        If tiempo_ronda Mod 30 = 0 And tiempo_ronda > 5 Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» En " & tiempo_ronda & " segundos comienza la ronda " & ronda_actual & ".", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_ronda <= 5 And tiempo_ronda > 0 Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» La ronda comienza en " & tiempo_ronda & ".", e_FontTypeNames.FONTTYPE_GUILD))
        ElseIf tiempo_ronda = 0 Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» ¡¡¡Yaaa!!!", e_FontTypeNames.FONTTYPE_GUILD))
            
            'Cambias triggers de los mapas para dejar pasar a los usuarios.

            
            For Each participante In participantes
                If participante.IsValid Then
                    Call enviarUsuarioBase(participante.PlayerIndex)
                End If
            Next participante
        
            Estado = EnJuego
        End If
    ElseIf Estado = EnJuego Then
        Dim Team As Byte
        Dim estaEnBandera(1 To 2) As Boolean
        For Each participante In participantes
            If participante.IsValid Then
                With UserList(participante.PlayerIndex)
                team = .flags.CurrentTeam
                If .pos.map = mapaEquipo(Team) Then 'Si estoy en mi mapa
                    If Abs(.pos.X - xEquipo(Team)) <= 8 And Abs(.pos.y - yEquipo(Team)) <= 5 Then ' .Pos.x Distance(.Pos.X, .Pos.Y, xEquipo(Team), yEquipo(Team)) <= 8 Then 'Y estoy 5 pasos del centro
                        If TieneObjetos(objEquipo(IIf(Team = 1, 2, 1)), 1, participante.PlayerIndex) Then 'Y tengo la bandera enemgia
                            contadorBandera(Team) = contadorBandera(Team) + 1
                            estaEnBandera(Team) = True
                            If contadorBandera(Team) >= CAPTURA_TIEMPO_BANDERA Then
                                Call finalizarRonda(Team)
                            Else
                                Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» ¡Atención! " & .Name & " del equipo " & colorEquipo(team) & " tiene la bandera enemiga en su base, si aguanta " & CAPTURA_TIEMPO_BANDERA - contadorBandera(team) & " segundos más ganará la ronda.", e_FontTypeNames.FONTTYPE_GUILD))
                            End If
                            
                        End If
                    End If
                ElseIf .flags.jugando_captura_timer > 0 Then
                    .flags.jugando_captura_timer = .flags.jugando_captura_timer - 1
                    If .flags.jugando_captura_timer = 0 And participante.IsValid Then
                        Call enviarUsuarioBase(participante.PlayerIndex)
                    End If
                End If
                End With
            End If
        Next participante
        If Not estaEnBandera(1) Then
            contadorBandera(1) = 0
        End If
        
        If Not estaEnBandera(2) Then
            contadorBandera(2) = 0
        End If
    End If
End Sub
Private Function sortParticipantes() As Collection
    Dim i As Integer
    Dim e As Integer
    Dim UI As Integer
    Dim added As Boolean
    Set sortParticipantes = New Collection
    
    For i = 1 To participantes.Count
        added = False
        Dim participante As clsCapturaParticipante
        Set participante = Participantes.Item(i)
        If participante.IsValid Then
            For e = 1 To sortParticipantes.count
                UI = participante.PlayerIndex
                If UserList(UI).Stats.ELV > UserList(sortParticipantes.Item(e).PlayerIndex).Stats.ELV Then
                    Call sortParticipantes.Add(participante, CStr(UI), e)
                    added = True
                    Exit For
                End If
            Next e
            If Not added Then
                Call sortParticipantes.Add(participante, CStr(participante.PlayerIndex))
            End If
        End If
    Next i
End Function

Private Sub iniciarCaptura()

    tiempo_espera = 0
    
    
    'Ordenar equipos
    Dim Team As Byte
    Dim counter As Byte
    Dim participante As clsCapturaParticipante
    
    If participantes.Count Mod 2 <> 0 Then
        Dim player As clsCapturaParticipante
        Set player = Participantes(Participantes.Count)
        Call eliminarParticipante(player)
        If player.IsValid Then
            Call WriteConsoleMsg(player.PlayerIndex, "Has sido eliminado del evento por ser el último en ingresar y los equipos eran impares.", e_FontTypeNames.FONTTYPE_INFOBOLD)
        End If
    End If
    
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» El evento captura la bandera ha iniciado.", e_FontTypeNames.FONTTYPE_GUILD))
    
    Dim sortedParticipantes As Collection
    Set sortedParticipantes = sortParticipantes()
    
    Team = 1
    counter = 1
    'Armamos los equipos segun level para que sea lo mas parejo.
    For Each participante In sortedParticipantes
        If participante.IsValid Then
            participante.Team = Team
            UserList(participante.PlayerIndex).flags.CurrentTeam = team
            UserList(participante.PlayerIndex).flags.tiene_bandera = 0
             Call SendData(SendTarget.ToAll, participante.PlayerIndex, PrepareMessageUpdateFlag(UserList(participante.PlayerIndex).flags.tiene_bandera, UserList(participante.PlayerIndex).Char.charindex))
            Call WriteConsoleMsg(participante.PlayerIndex, "Te ha tocado participar en el equipo " & colorEquipo(participante.team) & " deberás proteger a tus compañeros de equipo y capturar la bandera del equipo contrario para ganar.", e_FontTypeNames.FONTTYPE_GUILD)
            counter = counter + 1
            If counter = 2 Then
                Team = IIf(Team = 1, 2, 1)
                counter = 0
            End If
        End If
    Next participante

    Call iniciarRonda
End Sub

Public Sub muereUsuario(ByVal UserIndex As Integer)

    With UserList(UserIndex)
    
        If .flags.CurrentTeam <= 0 Then
            Exit Sub
        End If
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» " & .Name & " del equipo " & colorEquipo(.flags.CurrentTeam) & " ha muerto en batalla.", e_FontTypeNames.FONTTYPE_GUILD))
        
        Call tirarBandera(UserIndex)
        
        .flags.jugando_captura_muertes = .flags.jugando_captura_muertes + 1
        .flags.jugando_captura_timer = .flags.jugando_captura_muertes * CAPTURA_TIEMPO_POR_MUERTE + CAPTURA_TIEMPO_MUERTE
        
        
        Call WriteConsoleMsg(UserIndex, "Volverás a tu base en " & .flags.jugando_captura_timer & " segundos.", e_FontTypeNames.FONTTYPE_GUILD)
        
        
        Call enviarUsuarioSala(UserIndex)
    
        Call ResucitarOCurar(UserIndex)
       
    End With
End Sub

Public Function tomaBandera(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Boolean
    With UserList(UserIndex)
        Dim banderaTeam As Byte
        
        If objIndex = OBJ_CAPTURA_BANDERA_1 Then
            banderaTeam = 1
        ElseIf objIndex = OBJ_CAPTURA_BANDERA_2 Then
            banderaTeam = 2
        Else
            tomaBandera = True
            Exit Function
        End If
        
        
        If banderaTeam = .flags.CurrentTeam Then
             If .Pos.X = xEquipo(banderaTeam) And .Pos.Y = yEquipo(banderaTeam) Then
                tomaBandera = False
                Exit Function
            End If
        End If
        'Si toma la bandera del equipo contrario
        If UserList(UserIndex).flags.CurrentTeam <> banderaTeam Then
            UserList(UserIndex).flags.tiene_bandera = banderaTeam
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(UserList(UserIndex).flags.tiene_bandera, UserList(UserIndex).Char.charindex))
        Else
            UserList(UserIndex).flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(UserList(UserIndex).flags.tiene_bandera, UserList(UserIndex).Char.charindex))
        End If
        tomaBandera = True
    End With
End Function


Public Sub quitarBandera(ByVal UserIndex As Integer, ByVal objIndex As Integer)
    With UserList(UserIndex)
        Dim banderaTeam As Byte
        
        If objIndex = OBJ_CAPTURA_BANDERA_1 Then
            banderaTeam = 1
        ElseIf objIndex = OBJ_CAPTURA_BANDERA_2 Then
            banderaTeam = 2
        Else
            Exit Sub
        End If
                
        If banderaTeam <> .flags.CurrentTeam Then
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» ¡Atención! " & .Name & " del equipo " & colorEquipo(.flags.CurrentTeam) & " ha tomado la bandera del equipo " & colorEquipo(banderaTeam) & ".", e_FontTypeNames.FONTTYPE_GUILD))
            UserList(UserIndex).flags.tiene_bandera = banderaTeam
        Else
            UserList(UserIndex).flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, UserList(UserIndex).Char.charindex))
            Call QuitarObjetos(objIndex, 1, UserIndex)
            Call restaurarBandera(banderaTeam)
            
            Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» ¡Atención! La bandera del equipo " & colorEquipo(banderaTeam) & " ha sido recuperada.", e_FontTypeNames.FONTTYPE_GUILD))
        End If
    End With
End Sub

Public Sub tiraBandera(ByVal UserIndex As Integer, ByVal objIndex As Integer)
    'El usuario tira la bandera
    With UserList(UserIndex)
        Dim Team As Byte
        
        If objIndex = OBJ_CAPTURA_BANDERA_1 Then
            Team = 1
        ElseIf objIndex = OBJ_CAPTURA_BANDERA_2 Then
            Team = 2
        Else
            Exit Sub
        End If
        
        contadorBandera(Team) = 0
        
        UserList(UserIndex).flags.tiene_bandera = 0
        
        Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(UserList(UserIndex).flags.tiene_bandera, UserList(UserIndex).Char.charindex))
        Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» ¡Atención! la bandera del equipo " & colorEquipo(team) & " se encuentra en (" & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ")", e_FontTypeNames.FONTTYPE_GUILD))
        
    End With
End Sub

Public Sub tirarBandera(ByVal UserIndex As Integer)
    'Verificamos si tiene la bandera
    Dim Team As Byte
    Dim objBandera As Integer
    With UserList(UserIndex)
        team = .flags.CurrentTeam
        If Team = 0 Then
            Exit Sub
        ElseIf Team = 1 Then
            objBandera = OBJ_CAPTURA_BANDERA_2
        ElseIf Team = 2 Then
            objBandera = OBJ_CAPTURA_BANDERA_1
        End If
    
        
        If TieneObjetos(objBandera, 1, UserIndex) Then
            UserList(UserIndex).flags.tiene_bandera = 0
            Call SendData(SendTarget.ToAll, UserIndex, PrepareMessageUpdateFlag(0, UserList(UserIndex).Char.charindex))
            Call QuitarObjetos(objBandera, 1, UserIndex)
            
            If Estado = EnJuego Then
                Dim MiObj As t_Obj
        
                MiObj.amount = 1
                MiObj.objIndex = objBandera
            
                Call TirarItemAlPiso(.Pos, MiObj, False)
                Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» ¡Atención! la bandera del equipo " & colorEquipo(team) & " se encuentra en (" & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ")", e_FontTypeNames.FONTTYPE_GUILD))
            End If
            
        End If
    End With
End Sub

Private Sub enviarUsuarioBase(ByVal UserIndex As Integer)
    Dim pos_map As Integer
    Dim pos_X As Byte
    Dim pos_Y As Byte
    If UserList(UserIndex).flags.CurrentTeam = 1 Then
        pos_map = MAP_TEAM_1
        pos_X = X_TEAM_1
        pos_Y = Y_TEAM_1
    ElseIf UserList(UserIndex).flags.CurrentTeam = 2 Then
        pos_map = MAP_TEAM_2
        pos_X = X_TEAM_2
        pos_Y = Y_TEAM_2
    Else
        Exit Sub
    End If
    Call FindLegalPos(UserIndex, pos_map, pos_X, pos_Y)
    Call WarpUserChar(UserIndex, pos_map, pos_X, pos_Y, True)

End Sub

Private Sub enviarUsuarioSala(ByVal UserIndex As Integer)
    Dim pos_X As Byte, pos_Y As Byte
    pos_X = RandomNumber(MIN_SALA_ESPERA_X, MAX_SALA_ESPERA_X)
    pos_Y = RandomNumber(MIN_SALA_ESPERA_Y, MAX_SALA_ESPERA_Y)
    
    'No.
    Call FindLegalPos(UserIndex, MAP_SALA_ESPERA, pos_X, pos_Y)
    Call WarpUserChar(UserIndex, MAP_SALA_ESPERA, pos_X, pos_Y, True)
End Sub

Private Sub iniciarRonda()
    Estado = IniciaRonda
    ronda_actual = ronda_actual + 1
    tiempo_ronda = CAPTURA_TIEMPO_INICIO_RONDA
    
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» La ronda " & ronda_actual & " dará comienzo en " & tiempo_ronda & " segundos.", e_FontTypeNames.FONTTYPE_GUILD))

    
    Dim participante As clsCapturaParticipante
    For Each participante In participantes
        If participante.IsValid Then
            Call SendData(SendTarget.ToAll, participante.PlayerIndex, PrepareMessageUpdateFlag(0, UserList(participante.PlayerIndex).Char.charindex))
            UserList(participante.PlayerIndex).flags.jugando_captura_muertes = 0
            UserList(participante.PlayerIndex).flags.jugando_captura_timer = 0
            Call tirarBandera(participante.PlayerIndex)
            Call enviarUsuarioSala(participante.PlayerIndex)
        End If
    Next participante
    
    'Ponemos las banderas de nuevo
    Call restaurarBandera(1)
    Call restaurarBandera(2)
    

End Sub

Private Sub restaurarBandera(ByVal Team As Byte)
    Dim MiObj As t_Obj
    Dim MiPos As t_WorldPos
    MiObj.amount = 1
    MiObj.objIndex = objEquipo(Team)
    MiPos.map = mapaEquipo(Team)
    MiPos.X = xEquipo(Team)
    MiPos.Y = yEquipo(Team)
    
    If MapData(MiPos.map, MiPos.X, MiPos.Y).ObjInfo.amount <= 0 Then
        Call TirarItemAlPiso(MiPos, MiObj, False)
    End If
    
End Sub


Public Sub finalizarCaptura(Optional Team As Integer = -1)

    Dim participante As clsCapturaParticipante
    Estado = Finalizado
    If Team = -2 Then
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Se ha cancelado el evento de captura la bandera por falta de participantes.", e_FontTypeNames.FONTTYPE_GUILD))
    ElseIf Team = -1 Then 'Se aborto el evento
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» El evento de captura la bandera ha sido cancelado. Se te ha devuelto el precio de inscripción.", e_FontTypeNames.FONTTYPE_GUILD))
         For Each participante In participantes
            If participante.IsValid Then
                UserList(participante.PlayerIndex).Stats.GLD = UserList(participante.PlayerIndex).Stats.GLD + precio
                Call WriteUpdateGold(participante.PlayerIndex)
            End If
        Next participante
    ElseIf Team = 0 Then 'Hubo empate
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» ¡El evento de captura la bandera ha terminado en empate!.", e_FontTypeNames.FONTTYPE_GUILD))
    Else
    
        'Hay premios???
    
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» ¡El evento de captura la bandera ha terminado! Ganador el equipo " & colorEquipo(team) & ".", e_FontTypeNames.FONTTYPE_GUILD))
    End If

    
    For Each participante In participantes
        Call eliminarParticipante(participante)
    Next participante

    'Destruyo la instancia
    Set InstanciaCaptura = Nothing
End Sub

Public Sub finalizarRonda(ByVal Team As Byte)

    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessageConsoleMsg("Eventos» ¡El equipo " & colorEquipo(team) & " ha ganado la ronda " & ronda_actual & "!", e_FontTypeNames.FONTTYPE_GUILD))
    'El sonido de fundar clan para hacerlo mas magestuoso, mirar si hay un sonido mejor
    Call SendData(SendTarget.ToJugadoresCaptura, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
    
    contadorBandera(1) = 0
    contadorBandera(2) = 0
    
    
    rondasGanadas(Team) = rondasGanadas(Team) + 1
    If ronda_actual >= cantidad_rondas Then
        'Se termina la joda
        
        If rondasGanadas(1) > rondasGanadas(2) Then
            Team = 1
        ElseIf rondasGanadas(2) > rondasGanadas(1) Then
            Team = 2
        Else
            Team = 0
        End If
        
        Call finalizarCaptura(CInt(Team))
    Else
        Call iniciarRonda
    End If
End Sub

Public Function inicializar(ByVal MaxParticipantes As Long, ByVal CantidadRondas As Byte, ByVal NivelMinimo As Byte, ByVal NivelMaximo, ByVal precio_ As Long) As Boolean
    
    'Agregar random en comando para que no lo macreen rr-magax
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos» Se ha iniciado un Captura de la bandera de " & MaxParticipantes & " participantes, el precio de inscripción es de " & PonerPuntos(precio_) & ", el nivel mínimo para ingresar es " & NivelMinimo & " y el maximo es " & NivelMaximo & ", para participar envía /PARTICIPAR Las inscripciones se cerraran en 3 minutos.", e_FontTypeNames.FONTTYPE_GUILD))
        
    max_participantes = MaxParticipantes
    cantidad_rondas = CantidadRondas
    nivel_minimo = NivelMinimo
    nivel_maximo = NivelMaximo
    tiempo_espera = CAPTURA_TIEMPO_ESPERA
    precio = precio_
    Estado = Inscripcion
End Function

Public Sub inscribirse(ByVal UserIndex As Integer)
    'Validaciones al inscribirse
    
    With UserList(UserIndex)
        
        If Estado <> Inscripcion Then
            Call WriteConsoleMsg(UserIndex, "Ya se ha cerrado la inscripción para el evento.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Stats.GLD < precio Then
            Call WriteConsoleMsg(UserIndex, "No tienes oro suficiente para ingresar al evento.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.jugando_captura = 1 Then
            Call WriteConsoleMsg(UserIndex, "Ya estás anotado en el evento.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Nivel
        If .Stats.ELV < nivel_minimo Then
            Call WriteConsoleMsg(UserIndex, "Para ingresar al evento deberás ser al menos nivel " & nivel_minimo & ".", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .Stats.ELV > nivel_maximo Then
            Call WriteConsoleMsg(UserIndex, "Para ingresar al evento tu nivel no puede ser mayor a " & nivel_maximo & ".", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Zona segura
        If MapInfo(.Pos.Map).Seguro = 0 Or Zona(.ZonaId).Segura = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solamente podrás ingresar al evento en zona segura.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'cupos completos
        If participantes.Count >= max_participantes Then
            Call WriteConsoleMsg(UserIndex, "Los cupos ya están completos.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'que no esté en carcel
        If .Counters.Pena > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes anotarte al evento mientras estás encarcelado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'que no esté muerto
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes anotarte a un evento estando muerto.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'que no esté comerciando
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No puedes ingresar al evento mientras estás comerciando.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Montado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes ingresar al evento estando montado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes ingresar al evento mientras estás navegando.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Dim participante As clsCapturaParticipante
        Set participante = New clsCapturaParticipante
                
        Call participante.SetPlayer(UserIndex)
        participante.LastPosMap = UserList(UserIndex).Pos.map
        participante.LastPosX = UserList(UserIndex).Pos.X
        participante.LastPosY = UserList(UserIndex).Pos.Y
        
        
        Call participantes.Add(participante, CStr(UserIndex))
        
        UserList(UserIndex).flags.jugando_captura = 1
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - precio
        
        Call WriteUpdateGold(UserIndex)
        Call enviarUsuarioSala(UserIndex)
        
        If participantes.Count >= max_participantes Then
            Call iniciarCaptura
        End If
        
    End With
End Sub

