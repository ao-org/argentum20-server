Attribute VB_Name = "Trabajo"
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

Public Const GOLD_OBJ_INDEX As Long = 12
Public Const FISHING_NET_FX As Long = 12
Public Const NET_INMO_DURATION = 10

Function ExpectObjectTypeAt(ByVal objectType As Integer, ByVal Map As Integer, ByVal MapX As Byte, ByVal MapY As Byte) As Boolean
    Dim objIndex As Integer
    objIndex = MapData(Map, MapX, MapY).ObjInfo.objIndex
    If objIndex = 0 Then
        ExpectObjectTypeAt = False
        Exit Function
    End If
    ExpectObjectTypeAt = ObjData(objIndex).OBJType = objectType
End Function

Function IsUserAtPos(ByVal map As Integer, ByVal X As Byte, ByVal y As Byte) As Boolean
    IsUserAtPos = MapData(map, X, y).UserIndex > 0
End Function

Function IsNpcAtPos(ByVal map As Integer, ByVal X As Byte, ByVal y As Byte)
    IsNpcAtPos = MapData(map, X, y).npcIndex > 0
End Function

Sub HandleFishingNet(ByVal UserIndex As Integer)
On Error GoTo HandleFishingNet_Err:
    With UserList(UserIndex)
100 If (MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).Blocked And FLAG_AGUA) <> 0 _
        Or MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).trigger = e_Trigger.PESCAINVALIDA Then
102     If Abs(.pos.x - .Trabajo.Target_X) + Abs(.pos.y - .Trabajo.Target_Y) > 8 Then
104         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
106      Call WriteWorkRequestTarget(UserIndex, 0)
            Exit Sub
        End If
108     If UserList(UserIndex).Stats.UserSkills(e_Skill.Pescar) < 80 Then
110         Call WriteConsoleMsg(UserIndex, "Para utilizar la red de pesca debes tener 80 skills en recoleccion.", e_FontTypeNames.FONTTYPE_INFO)
112         Call WriteWorkRequestTarget(UserIndex, 0)
            Exit Sub
        End If
114     If MapInfo(UserList(UserIndex).pos.Map).Seguro = 1 Then
116         Call WriteConsoleMsg(UserIndex, "Esta prohibida la pesca masiva en las ciudades.", e_FontTypeNames.FONTTYPE_INFO)
118         Call WriteWorkRequestTarget(UserIndex, 0)
            Exit Sub
        End If
120         If UserList(UserIndex).flags.Navegando = 0 Then
122             Call WriteConsoleMsg(UserIndex, "Necesitas estar sobre tu barca para utilizar la red de pesca.", e_FontTypeNames.FONTTYPE_INFO)
124             Call WriteWorkRequestTarget(UserIndex, 0)
                Exit Sub
        End If
        Call DoPescar(UserIndex, True)
    Else
132     Call WriteConsoleMsg(UserIndex, "Zona de pesca no Autorizada. Busca otro lugar para hacerlo.", e_FontTypeNames.FONTTYPE_INFO)
142     Call WriteWorkRequestTarget(UserIndex, 0)
    End If
    End With
    Exit Sub
HandleFishingNet_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.HandleFishingNet", Erl)
End Sub

Public Sub Trabajar(ByVal UserIndex As Integer, ByVal Skill As e_Skill)
        Dim DummyInt As Integer

        With UserList(UserIndex)
                                      
               Select Case Skill
                  
                Case e_Skill.Pescar
288                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
290                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> e_OBJType.otHerramientas Then Exit Sub
                    
292                 If MapInfo(UserList(UserIndex).pos.map).Seguro = 0 Then
                        If Not ExpectObjectTypeAt(e_OBJType.otFishingPool, .pos.map, .Trabajo.Target_X, .Trabajo.Target_Y) Then
                            Call WriteConsoleMsg(UserIndex, "Necesitas un area de pesca en en el mar.", e_FontTypeNames.FONTTYPE_INFO)
                            Call WriteWorkRequestTarget(UserIndex, 0)
                            Call WriteMacroTrabajoToggle(UserIndex, False)
                            Exit Sub
                        End If
                    End If

294                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case e_ToolsSubtype.eFishingRod

296                         If (MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).Blocked And FLAG_AGUA) <> 0 And Not MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = e_Trigger.PESCAINVALIDA Then
298                             If (MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X + 1, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X, .Pos.Y + 1).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X - 1, .Pos.Y).Blocked And FLAG_AGUA) <> 0 Or (MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).Blocked And FLAG_AGUA) <> 0 Then
300                                 .flags.PescandoEspecial = False
                                    Call DoPescar(UserIndex, False)
                                Else
304                                 Call WriteConsoleMsg(UserIndex, "Acércate a la costa para pescar.", e_FontTypeNames.FONTTYPE_INFO)
306                                 Call WriteMacroTrabajoToggle(UserIndex, False)
                                End If
                            Else
308                             Call WriteConsoleMsg(UserIndex, "Zona de pesca no Autorizada. Busca otro lugar para hacerlo.", e_FontTypeNames.FONTTYPE_INFO)
310                             Call WriteMacroTrabajoToggle(UserIndex, False)
                            End If
312                     Case e_ToolsSubtype.eFishingNet
                            Call HandleFishingNet(UserIndex)
                    End Select
                    
                Case e_Skill.Carpinteria
                    'Veo cual es la cantidad máxima que puede construir de una
                    Dim cantidad_maxima As Long
                    
                    If UserList(UserIndex).clase = e_Class.Trabajador Then
                        cantidad_maxima = UserList(UserIndex).Stats.UserSkills(e_Skill.Carpinteria) / 10
                        If cantidad_maxima = 0 Then cantidad_maxima = 1
                    Else
                        'Si no hace de a 1
                        cantidad_maxima = 1
                    End If
                    
                    Call CarpinteroConstruirItem(UserIndex, UserList(UserIndex).Trabajo.Item, UserList(UserIndex).Trabajo.cantidad, cantidad_maxima)

                Case e_Skill.Mineria
            
454                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                    
456                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> e_OBJType.otHerramientas Then Exit Sub
                    
                    'Check interval
458                 If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

460                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case 8  ' Herramientas de Mineria - Piquete
                
                            'Target whatever is in the tile
462                         Call LookatTile(UserIndex, .Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y)
                            
464                         DummyInt = MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex
                            
466                         If DummyInt > 0 Then

                                'Check distance
468                             If Abs(.Pos.X - .Trabajo.Target_X) + Abs(.Pos.Y - .Trabajo.Target_Y) > 2 Then
470                                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estís demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
472                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If

                                '¡Hay un yacimiento donde clickeo?
474                             If ObjData(DummyInt).OBJType = e_OBJType.otYacimiento Then

                                    ' Si el Yacimiento requiere herramienta `Dorada` y la herramienta no lo es, o vice versa.
                                    ' Se usa para el yacimiento de Oro.
476                                 If ObjData(DummyInt).Dorada <> ObjData(.invent.HerramientaEqpObjIndex).Dorada Or ObjData(DummyInt).Blodium <> ObjData(.invent.HerramientaEqpObjIndex).Blodium Then
478
                                        If ObjData(DummyInt).Blodium <> ObjData(.invent.HerramientaEqpObjIndex).Blodium Then
                                            Call WriteConsoleMsg(UserIndex, "El pico minero especial solo puede extraer minerales del yacimiento de Blodium.", e_FontTypeNames.FONTTYPE_INFO)
                                            Call WriteWorkRequestTarget(UserIndex, 0)
                                            Exit Sub
                                        Else
                                            Call WriteConsoleMsg(UserIndex, "El pico dorado solo puede extraer minerales del yacimiento de Oro.", e_FontTypeNames.FONTTYPE_INFO)
480                                         Call WriteWorkRequestTarget(UserIndex, 0)
                                            Exit Sub
                                        End If
                                    End If

482                                 If MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount <= 0 Then
484                                     Call WriteConsoleMsg(UserIndex, "Este yacimiento no tiene mas minerales para entregar.", e_FontTypeNames.FONTTYPE_INFO)
486                                     Call WriteWorkRequestTarget(UserIndex, 0)
488                                     Call WriteMacroTrabajoToggle(UserIndex, False)
                                        Exit Sub

                                    End If

490                                 Call DoMineria(UserIndex, .Trabajo.Target_X, .Trabajo.Target_Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)

                                Else
492                                 Call WriteConsoleMsg(UserIndex, "Ahí no hay ningún yacimiento.", e_FontTypeNames.FONTTYPE_INFO)
494                                 Call WriteWorkRequestTarget(UserIndex, 0)

                                End If

                            Else
496                             Call WriteConsoleMsg(UserIndex, "Ahí no hay ningun yacimiento.", e_FontTypeNames.FONTTYPE_INFO)
498                             Call WriteWorkRequestTarget(UserIndex, 0)

                            End If

                    End Select

                Case e_Skill.Talar

350                 If .Invent.HerramientaEqpObjIndex = 0 Then Exit Sub

352                 If ObjData(.Invent.HerramientaEqpObjIndex).OBJType <> e_OBJType.otHerramientas Then Exit Sub
        
                    'Check interval
354                 If Not IntervaloPermiteTrabajarExtraer(UserIndex) Then Exit Sub

356                 Select Case ObjData(.Invent.HerramientaEqpObjIndex).Subtipo
                
                        Case 6      ' Herramientas de Carpinteria - Hacha

358                         DummyInt = MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex
                            
360                         If DummyInt > 0 Then
362                             If Abs(.Pos.X - .Trabajo.Target_X) + Abs(.Pos.Y - .Trabajo.Target_Y) > 1 Then
364                                 Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                                    'Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", e_FontTypeNames.FONTTYPE_INFO)
366                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
                                
368                             If .Pos.X = .Trabajo.Target_X And .Pos.Y = .Trabajo.Target_Y Then
370                                 Call WriteConsoleMsg(UserIndex, "No podés talar desde allí.", e_FontTypeNames.FONTTYPE_INFO)
372                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If

374                             If ObjData(DummyInt).Elfico <> ObjData(.Invent.HerramientaEqpObjIndex).Elfico Then
376                                 Call WriteConsoleMsg(UserIndex, "Sólo puedes talar árboles elficos con un hacha élfica.", e_FontTypeNames.FONTTYPE_INFO)
378                                 Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
379                             If ObjData(DummyInt).Pino <> ObjData(.invent.HerramientaEqpObjIndex).Pino Then
                                    Call WriteConsoleMsg(UserIndex, "Sólo puedes talar árboles de pino nudoso con un hacha de pino.", e_FontTypeNames.FONTTYPE_INFO)
                                    Call WriteWorkRequestTarget(UserIndex, 0)
                                    Exit Sub

                                End If
380                             If MapData(.Pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount <= 0 Then
382                                 Call WriteConsoleMsg(UserIndex, "El árbol ya no te puede entregar mas leña.", e_FontTypeNames.FONTTYPE_INFO)
384                                 Call WriteWorkRequestTarget(UserIndex, 0)
386                                 Call WriteMacroTrabajoToggle(UserIndex, False)
                                    Exit Sub

                                End If

                                '¡Hay un arbol donde clickeo?
388                             If ObjData(DummyInt).OBJType = e_OBJType.otArboles Then
390                                 Call DoTalar(UserIndex, .Trabajo.Target_X, .Trabajo.Target_Y, ObjData(.Invent.HerramientaEqpObjIndex).Dorada = 1)

                                End If

                            Else
392                             Call WriteConsoleMsg(UserIndex, "No hay ningún árbol ahí.", e_FontTypeNames.FONTTYPE_INFO)
394                             Call WriteWorkRequestTarget(UserIndex, 0)

396                             If UserList(UserIndex).Counters.Trabajando > 1 Then
398                                 Call WriteMacroTrabajoToggle(UserIndex, False)

                                End If

                            End If
                
                    End Select
                    
574             Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
            
                    'Check interval
576                 If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
                                
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
590                                 Call WriteConsoleMsg(UserIndex, "No tienes más minerales", e_FontTypeNames.FONTTYPE_INFO)
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
                    
600                         Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", e_FontTypeNames.FONTTYPE_INFO)
602                         Call WriteWorkRequestTarget(UserIndex, 0)

604                         If UserList(UserIndex).Counters.Trabajando > 1 Then
606                             Call WriteMacroTrabajoToggle(UserIndex, False)

                            End If

                        End If

                    Else
                
608                     Call WriteConsoleMsg(UserIndex, "Ahí no hay ninguna fragua.", e_FontTypeNames.FONTTYPE_INFO)
610                     Call WriteWorkRequestTarget(UserIndex, 0)

612                     If UserList(UserIndex).Counters.Trabajando > 1 Then
614                         Call WriteMacroTrabajoToggle(UserIndex, False)

                        End If

                    End If
            End Select
        End With
End Sub

Public Sub DoPermanecerOculto(ByVal UserIndex As Integer)
        '********************************************************
        'Autor: Nacho (Integer)
        'Last Modif: 28/01/2007
        'Chequea si ya debe mostrarse
        'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
        '********************************************************
        
        On Error GoTo DoPermanecerOculto_Err
    
100     With UserList(UserIndex)
            Dim velocidadOcultarse As Integer
            velocidadOcultarse = 1
            'HarThaoS: Si tiene armadura de cazador, dependiendo skills vemos cuanto tiempo se oculta
            If .clase = e_Class.Hunter Then
                If TieneArmaduraCazador(UserIndex) Then
                    Select Case .Stats.UserSkills(e_Skill.Ocultarse)
                        Case Is = 100
                            Exit Sub
                        Case Is < 100
                            velocidadOcultarse = RandomNumber(0, 1)
                    End Select
                End If
            End If
            
104         .Counters.TiempoOculto = .Counters.TiempoOculto - velocidadOcultarse
106         If .Counters.TiempoOculto <= 0 Then

108             .Counters.TiempoOculto = 0
110             .flags.Oculto = 0

112             If .flags.Navegando = 1 Then
            
114                 If .clase = e_Class.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
116                     Call EquiparBarco(UserIndex)
124                     Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", e_FontTypeNames.FONTTYPE_INFO)
126                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart)
                        Call RefreshCharStatus(UserIndex)
                    End If

                Else

128                 If .flags.invisible = 0 And .flags.AdminInvisible = 0 Then
130                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
132                     Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", e_FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            
            End If
    
        End With

        Exit Sub

DoPermanecerOculto_Err:
134     Call TraceError(Err.Number, Err.Description, "Trabajo.DoPermanecerOculto", Erl)

136
        
End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)

        'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
        'Modifique la fórmula y ahora anda bien.
        On Error GoTo ErrHandler

        Dim Suerte As Double
        Dim res    As Integer
        Dim Skill  As Integer
    
100     With UserList(UserIndex)

102         If .flags.Navegando = 1 And .clase <> e_Class.Pirat Then
104             Call WriteLocaleMsg(UserIndex, "56", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If GlobalFrameTime - .Counters.LastAttackTime < HideAfterHitTime Then
                Exit Sub
            End If
          
    
106         Skill = .Stats.UserSkills(e_Skill.Ocultarse)
108         Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
110         res = RandomNumber(1, 100)

112         If res <= Suerte Then

114             .flags.Oculto = 1
116             Suerte = (-0.000001 * (100 - Skill) ^ 3)
118             Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
120             Suerte = Suerte + (-0.0088 * (100 - Skill))
122             Suerte = Suerte + (0.9571)
124             Suerte = Suerte * IntervaloOculto
        
                Select Case .clase
                    Case e_Class.Bandit, e_Class.Thief
128                     .Counters.TiempoOculto = Int(Suerte / 2)
                    
                    Case e_Class.Hunter
130                     .Counters.TiempoOculto = Int(Suerte / 2)
                    
                    Case Else
129                     .Counters.TiempoOculto = Int(Suerte / 3)
                End Select
        
138             If .flags.Navegando = 1 Then
140                 If .clase = e_Class.Pirat Then
142                     .Char.Body = iFragataFantasmal
144                     .flags.Oculto = 1
146                     .Counters.TiempoOculto = IntervaloOculto
                         
148                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NoCart)
150                     Call WriteConsoleMsg(UserIndex, "¡Te has camuflado como barco fantasma!", e_FontTypeNames.FONTTYPE_INFO)
                        Call RefreshCharStatus(UserIndex)
                    End If
                Else
                    UserList(UserIndex).Counters.timeFx = 3
152                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                    'Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", e_FontTypeNames.FONTTYPE_INFO)
154                 Call WriteLocaleMsg(UserIndex, "55", e_FontTypeNames.FONTTYPE_INFO)
                End If


156             Call SubirSkill(UserIndex, Ocultarse)
            Else

158             If Not .flags.UltimoMensaje = 4 Then
                    'Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", e_FontTypeNames.FONTTYPE_INFO)
160                 Call WriteLocaleMsg(UserIndex, "57", e_FontTypeNames.FONTTYPE_INFO)
162                 .flags.UltimoMensaje = 4
                End If

            End If

164         .Counters.Ocultando = .Counters.Ocultando + 1
    
        End With

        Exit Sub

ErrHandler:
166     Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal UserIndex As Integer, _
                    ByRef Barco As t_ObjData, _
                    ByVal Slot As Integer)
        
        On Error GoTo DoNavega_Err

100     With UserList(UserIndex)

102         If .Invent.BarcoObjIndex <> .Invent.Object(Slot).ObjIndex Then

104             If Not EsGM(UserIndex) Then
            
106                 Select Case Barco.Subtipo
        
                        Case 2  'Galera
        
108                         If .clase <> e_Class.Trabajador And .clase <> e_Class.Pirat Then
110                             Call WriteConsoleMsg(UserIndex, "¡Solo Piratas y trabajadores pueden usar galera!", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        
112                     Case 3  'Galeón
                    
114                         If .clase <> e_Class.Pirat Then
116                             Call WriteConsoleMsg(UserIndex, "Solo los Piratas pueden usar Galeón!!", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                        
                    End Select
                    
                End If
            
                Dim SkillNecesario As Byte
                
118             SkillNecesario = IIf(.clase = e_Class.Trabajador Or .clase = e_Class.Pirat, Barco.MinSkill \ 2, Barco.MinSkill)
                
                ' Tiene el skill necesario?
120             If .Stats.UserSkills(e_Skill.Navegacion) < SkillNecesario Then
122                 Call WriteConsoleMsg(UserIndex, "Necesitas al menos " & SkillNecesario & " puntos en navegación para poder usar este " & IIf(Barco.Subtipo = 0, "traje", "barco") & ".", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
124             If .Invent.BarcoObjIndex = 0 Then
126                 Call WriteNavigateToggle(UserIndex, True)
128                 .flags.Navegando = 1
                    Call TargetUpdateTerrain(.EffectOverTime)
                End If
    
130             .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
132             .Invent.BarcoSlot = Slot
    
134             If .flags.Montado > 0 Then
136                 Call DoMontar(UserIndex, ObjData(.Invent.MonturaObjIndex), .Invent.MonturaSlot)
                End If
                
138             If .flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
140                 Call WriteConsoleMsg(UserIndex, "Pierdes el efecto del mimetismo.", e_FontTypeNames.FONTTYPE_INFO)
142                 .Counters.Mimetismo = 0
144                 .flags.Mimetizado = e_EstadoMimetismo.Desactivado
                    Call RefreshCharStatus(UserIndex)
                End If
                
    
146             Call EquiparBarco(UserIndex)
            
            Else
148             Call WriteNadarToggle(UserIndex, False)
                .flags.Navegando = 0
150             Call WriteNavigateToggle(UserIndex, False)
    
152
                Call TargetUpdateTerrain(.EffectOverTime)
154             .Invent.BarcoObjIndex = 0
156             .Invent.BarcoSlot = 0
    
158             If .flags.Muerto = 0 Then
160                 .Char.Head = .OrigChar.Head
        
162                 If .Invent.ArmourEqpObjIndex > 0 Then
164                     .Char.Body = ObtenerRopaje(UserIndex, ObjData(.Invent.ArmourEqpObjIndex))
                    Else
                        Call SetNakedBody(UserList(userIndex))
                    End If
        
168                 If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
170                 If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
174                 If .Invent.HerramientaEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.HerramientaEqpObjIndex).WeaponAnim
176                 If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
177                 If .invent.MagicoObjIndex > 0 Then
                        If ObjData(.invent.MagicoObjIndex).Ropaje > 0 Then .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje
                    End If
                Else
178                 .Char.Body = iCuerpoMuerto
180                 .Char.Head = 0
182                 Call ClearClothes(.Char)

                End If
                Call ActualizarVelocidadDeUsuario(UserIndex)

            End If
        
            ' Volver visible
190         If .flags.Oculto = 1 And .flags.AdminInvisible = 0 And .flags.invisible = 0 Then
192             .flags.Oculto = 0
194             .Counters.TiempoOculto = 0

                'Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", e_FontTypeNames.FONTTYPE_INFO)
196             Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
198             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
            End If

200         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
202         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(e_FXSound.BARCA_SOUND, .Pos.X, .Pos.y))
    
        End With
        
        Exit Sub

DoNavega_Err:
204     Call TraceError(Err.Number, Err.Description, "Trabajo.DoNavega", Erl)

206
        
End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)
        
        On Error GoTo FundirMineral_Err

100     If UserList(UserIndex).clase <> e_Class.Trabajador Then
102         Call WriteConsoleMsg(UserIndex, "Tu clase no tiene el conocimiento suficiente para trabajar este mineral.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
104     If UserList(userindex).flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios) Then
            Exit Sub
        End If

106     If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then

            Dim SkillRequerido As Integer
108         SkillRequerido = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill
   
110         If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = e_OBJType.otMinerales And _
                UserList(UserIndex).Stats.UserSkills(e_Skill.Mineria) >= SkillRequerido Then
            
112             Call DoLingotes(UserIndex)
        
114         ElseIf SkillRequerido > 100 Then
116             Call WriteConsoleMsg(UserIndex, "Los mortales no pueden fundir este mineral.", e_FontTypeNames.FONTTYPE_INFO)
                
            Else
118             Call WriteConsoleMsg(UserIndex, "No tenés conocimientos de minería suficientes para trabajar este mineral. Necesitas " & SkillRequerido & " puntos en minería.", e_FontTypeNames.FONTTYPE_INFO)

            End If

        End If

        
        Exit Sub

FundirMineral_Err:
120     Call TraceError(Err.Number, Err.Description, "Trabajo.FundirMineral", Erl)
122
        
End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
        On Error GoTo TieneObjetos_Err
        

        Dim i     As Long

        Dim Total As Long
        
        If (ItemIndex = GOLD_OBJ_INDEX) Then
            Total = UserList(UserIndex).Stats.GLD
        End If
        
100     For i = 1 To UserList(UserIndex).CurrentInventorySlots

102         If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
104             Total = Total + UserList(UserIndex).Invent.Object(i).amount

            End If

106     Next i

108     If cant <= Total Then
110         TieneObjetos = True
            Exit Function

        End If
        
        
        Exit Function

TieneObjetos_Err:
112     Call TraceError(Err.Number, Err.Description, "Trabajo.TieneObjetos", Erl)
114
        
End Function

Function QuitarObjetos(ByVal ItemIndex As Integer, ByVal cant As Integer, ByVal UserIndex As Integer) As Boolean
        On Error GoTo QuitarObjetos_Err
        
100     With UserList(UserIndex)
            Dim i As Long
            
102         For i = 1 To .CurrentInventorySlots
    
104             If .Invent.Object(i).ObjIndex = ItemIndex Then
    
106                 .Invent.Object(i).amount = .Invent.Object(i).amount - cant
    
108                 If .Invent.Object(i).amount <= 0 Then
110                     If .Invent.Object(i).Equipped Then
112                         Call Desequipar(UserIndex, i)
                        End If
    
114                     cant = Abs(.Invent.Object(i).amount)
116                     .Invent.Object(i).amount = 0
118                     .Invent.Object(i).ObjIndex = 0
                    Else
120                     cant = 0
    
                    End If

122                 Call UpdateUserInv(False, UserIndex, i)
            
124                 If cant = 0 Then
126                     QuitarObjetos = True
                        Exit Function
                    End If
    
                End If
    
128         Next i
           
            If (ItemIndex = GOLD_OBJ_INDEX And cant > 0) Then
                .Stats.GLD = .Stats.GLD - cant
                
                If (.Stats.GLD < 0) Then
                    cant = Abs(.Stats.GLD)
                    
                    .Stats.GLD = 0
                End If
                
                Call WriteUpdateGold(UserIndex)
                
                If (cant = 0) Then
                    QuitarObjetos = True
                    Exit Function
                End If
            End If
                
        End With
        
        Exit Function

QuitarObjetos_Err:
130     Call TraceError(Err.Number, Err.Description, "Trabajo.QuitarObjetos", Erl)
132
        
End Function

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then Call QuitarObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex)
102     If ObjData(ItemIndex).LingP > 0 Then Call QuitarObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex)
104     If ObjData(ItemIndex).LingO > 0 Then Call QuitarObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex)
106     If ObjData(ItemIndex).Coal > 0 Then Call QuitarObjetos(e_Minerales.Coal, ObjData(ItemIndex).Coal, UserIndex)
        Exit Sub
HerreroQuitarMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroQuitarMateriales", Erl)
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Integer)
        
        On Error GoTo CarpinteroQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Madera > 0 Then Call QuitarObjetos(Wood, Cantidad, UserIndex)
102     If ObjData(ItemIndex).MaderaElfica > 0 Then Call QuitarObjetos(ElvenWood, Cantidad, UserIndex)
104     If ObjData(ItemIndex).MaderaPino > 0 Then Call QuitarObjetos(PinoWood, Cantidad, UserIndex)
        Exit Sub
CarpinteroQuitarMateriales_Err:
106     Call TraceError(Err.Number, Err.Description, "Trabajo.CarpinteroQuitarMateriales", Erl)
End Sub

Sub AlquimistaQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaQuitarMateriales_Err
        

100     If ObjData(ItemIndex).Raices > 0 Then Call QuitarObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex)
        If ObjData(ItemIndex).Botella > 0 Then Call QuitarObjetos(Botella, ObjData(ItemIndex).Botella, UserIndex)
        If ObjData(ItemIndex).Cuchara > 0 Then Call QuitarObjetos(Cuchara, ObjData(ItemIndex).Cuchara, UserIndex)
        If ObjData(ItemIndex).Mortero > 0 Then Call QuitarObjetos(Mortero, ObjData(ItemIndex).Mortero, UserIndex)
        If ObjData(ItemIndex).FrascoAlq > 0 Then Call QuitarObjetos(FrascoAlq, ObjData(ItemIndex).FrascoAlq, UserIndex)
        If ObjData(ItemIndex).FrascoElixir > 0 Then Call QuitarObjetos(FrascoElixir, ObjData(ItemIndex).FrascoElixir, UserIndex)
        If ObjData(ItemIndex).Dosificador > 0 Then Call QuitarObjetos(Dosificador, ObjData(ItemIndex).Dosificador, UserIndex)
        If ObjData(ItemIndex).Orquidea > 0 Then Call QuitarObjetos(Orquidea, ObjData(ItemIndex).Orquidea, UserIndex)
        If ObjData(ItemIndex).Carmesi > 0 Then Call QuitarObjetos(Carmesi, ObjData(ItemIndex).Carmesi, UserIndex)
        If ObjData(ItemIndex).HongoDeLuz > 0 Then Call QuitarObjetos(HongoDeLuz, ObjData(ItemIndex).HongoDeLuz, UserIndex)
        If ObjData(ItemIndex).Esporas > 0 Then Call QuitarObjetos(Esporas, ObjData(ItemIndex).Esporas, UserIndex)
        If ObjData(ItemIndex).Tuna > 0 Then Call QuitarObjetos(Tuna, ObjData(ItemIndex).Tuna, UserIndex)
        If ObjData(ItemIndex).Cala > 0 Then Call QuitarObjetos(Cala, ObjData(ItemIndex).Cala, UserIndex)
        If ObjData(ItemIndex).ColaDeZorro > 0 Then Call QuitarObjetos(ColaDeZorro, ObjData(ItemIndex).ColaDeZorro, UserIndex)
        If ObjData(ItemIndex).FlorOceano > 0 Then Call QuitarObjetos(FlorOceano, ObjData(ItemIndex).FlorOceano, UserIndex)
        If ObjData(ItemIndex).FlorRoja > 0 Then Call QuitarObjetos(FlorRoja, ObjData(ItemIndex).FlorRoja, UserIndex)
        If ObjData(ItemIndex).Hierva > 0 Then Call QuitarObjetos(Hierva, ObjData(ItemIndex).Hierva, UserIndex)
        If ObjData(ItemIndex).HojasDeRin > 0 Then Call QuitarObjetos(HojasDeRin, ObjData(ItemIndex).HojasDeRin, UserIndex)
        If ObjData(ItemIndex).HojasRojas > 0 Then Call QuitarObjetos(HojasRojas, ObjData(ItemIndex).HojasRojas, UserIndex)
        If ObjData(ItemIndex).SemillasPros > 0 Then Call QuitarObjetos(SemillasPros, ObjData(ItemIndex).SemillasPros, UserIndex)
        If ObjData(ItemIndex).Pimiento > 0 Then Call QuitarObjetos(Pimiento, ObjData(ItemIndex).Pimiento, UserIndex)
        
        Exit Sub

AlquimistaQuitarMateriales_Err:
102     Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaQuitarMateriales", Erl)
104
        
End Sub

Sub SastreQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreQuitarMateriales_Err
        

100     If ObjData(ItemIndex).PielLobo > 0 Then Call QuitarObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex)
102     If ObjData(ItemIndex).PielOsoPardo > 0 Then Call QuitarObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex)
104     If ObjData(ItemIndex).PielOsoPolaR > 0 Then Call QuitarObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex)
106     If ObjData(ItemIndex).PielLoboNegro > 0 Then Call QuitarObjetos(PielLoboNegro, ObjData(ItemIndex).PielLoboNegro, UserIndex)
107     If ObjData(ItemIndex).PielTigre > 0 Then Call QuitarObjetos(PielTigre, ObjData(ItemIndex).PielTigre, UserIndex)
108     If ObjData(ItemIndex).PielTigreBengala > 0 Then Call QuitarObjetos(PielTigreBengala, ObjData(ItemIndex).PielTigreBengala, UserIndex)

        
        Exit Sub

SastreQuitarMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.SastreQuitarMateriales", Erl)
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Long) As Boolean
        
        On Error GoTo CarpinteroTieneMateriales_Err
    
100     If ObjData(ItemIndex).Madera > 0 Then
102         If Not TieneObjetos(Wood, ObjData(ItemIndex).Madera * Cantidad, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes madera.", e_FontTypeNames.FONTTYPE_INFO)
106             CarpinteroTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
        
110     If ObjData(ItemIndex).MaderaElfica > 0 Then
112         If Not TieneObjetos(ElvenWood, ObjData(ItemIndex).MaderaElfica * Cantidad, UserIndex) Then
114             Call WriteConsoleMsg(UserIndex, "No tenes suficiente madera elfica.", e_FontTypeNames.FONTTYPE_INFO)
116             CarpinteroTieneMateriales = False
118             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If

        End If
        
120     If ObjData(ItemIndex).MaderaPino > 0 Then
122         If Not TieneObjetos(PinoWood, ObjData(ItemIndex).MaderaPino * Cantidad, UserIndex) Then
124             Call WriteConsoleMsg(UserIndex, "No tenes suficiente madera pino nudoso.", e_FontTypeNames.FONTTYPE_INFO)
126             CarpinteroTieneMateriales = False
128             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If

        End If
    
        CarpinteroTieneMateriales = True
        
        Exit Function

CarpinteroTieneMateriales_Err:
        Call TraceError(Err.Number, Err.Description + " UI:" + UserIndex + " Item: " + ItemIndex, "Trabajo.CarpinteroTieneMateriales", Erl)
        
End Function

Function AlquimistaTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo AlquimistaTieneMateriales_Err
        
    
100     If ObjData(ItemIndex).Raices > 0 Then
102         If Not TieneObjetos(Raices, ObjData(ItemIndex).Raices, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes raices.", e_FontTypeNames.FONTTYPE_INFO)
106             AlquimistaTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
        
        If ObjData(ItemIndex).Botella > 0 Then
             If Not TieneObjetos(Botella, ObjData(ItemIndex).Botella, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Botellas.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Cuchara > 0 Then
             If Not TieneObjetos(Cuchara, ObjData(ItemIndex).Cuchara, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Cucharas.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Mortero > 0 Then
             If Not TieneObjetos(Mortero, ObjData(ItemIndex).Mortero, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Mortero.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
                
        If ObjData(ItemIndex).FrascoAlq > 0 Then
             If Not TieneObjetos(FrascoAlq, ObjData(ItemIndex).FrascoAlq, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Frascos de alquimistas.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        If ObjData(ItemIndex).FrascoElixir > 0 Then
             If Not TieneObjetos(FrascoElixir, ObjData(ItemIndex).FrascoElixir, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Frascos de elixir superior.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        If ObjData(ItemIndex).Dosificador > 0 Then
             If Not TieneObjetos(Dosificador, ObjData(ItemIndex).Dosificador, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Dosificadores.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Orquidea > 0 Then
             If Not TieneObjetos(Orquidea, ObjData(ItemIndex).Orquidea, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Orquídeas Silvestres.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        If ObjData(ItemIndex).Carmesi > 0 Then
             If Not TieneObjetos(Carmesi, ObjData(ItemIndex).Carmesi, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Raíces Carmesí.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).HongoDeLuz > 0 Then
             If Not TieneObjetos(HongoDeLuz, ObjData(ItemIndex).HongoDeLuz, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Hongos de Luz.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Esporas > 0 Then
             If Not TieneObjetos(Esporas, ObjData(ItemIndex).Esporas, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Esporas Silvestres.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Tuna > 0 Then
             If Not TieneObjetos(Tuna, ObjData(ItemIndex).Tuna, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Tunas Silvestres.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Cala > 0 Then
             If Not TieneObjetos(Cala, ObjData(ItemIndex).Cala, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Calas Venenosas.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).ColaDeZorro > 0 Then
             If Not TieneObjetos(ColaDeZorro, ObjData(ItemIndex).ColaDeZorro, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Colas de Zorro.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).FlorOceano > 0 Then
             If Not TieneObjetos(FlorOceano, ObjData(ItemIndex).FlorOceano, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Flores del óceano.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).FlorRoja > 0 Then
             If Not TieneObjetos(FlorRoja, ObjData(ItemIndex).FlorRoja, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Flores Roja.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Hierva > 0 Then
             If Not TieneObjetos(Hierva, ObjData(ItemIndex).Hierva, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Hiervas de Sangre.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).HojasDeRin > 0 Then
             If Not TieneObjetos(HojasDeRin, ObjData(ItemIndex).HojasDeRin, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Hojas de Rin.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).HojasRojas > 0 Then
             If Not TieneObjetos(HojasRojas, ObjData(ItemIndex).HojasRojas, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Hojas Rojas.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).SemillasPros > 0 Then
             If Not TieneObjetos(SemillasPros, ObjData(ItemIndex).SemillasPros, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Semillas Prosperas.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
        
        If ObjData(ItemIndex).Pimiento > 0 Then
             If Not TieneObjetos(Pimiento, ObjData(ItemIndex).Pimiento, UserIndex) Then
                 Call WriteConsoleMsg(UserIndex, "No tenes suficientes Pimientos Muerte.", e_FontTypeNames.FONTTYPE_INFO)
                 AlquimistaTieneMateriales = False
                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Function
    
                End If

        End If
    
110     AlquimistaTieneMateriales = True

        
        Exit Function

AlquimistaTieneMateriales_Err:
112     Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaTieneMateriales", Erl)
114
        
End Function

Function SastreTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo SastreTieneMateriales_Err
    
100     If ObjData(ItemIndex).PielLobo > 0 Then
102         If Not TieneObjetos(PieldeLobo, ObjData(ItemIndex).PielLobo, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de lobo.", e_FontTypeNames.FONTTYPE_INFO)
106             SastreTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If
        End If
    
110     If ObjData(ItemIndex).PielOsoPardo > 0 Then
112         If Not TieneObjetos(PieldeOsoPardo, ObjData(ItemIndex).PielOsoPardo, UserIndex) Then
114             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso pardo.", e_FontTypeNames.FONTTYPE_INFO)
116             SastreTieneMateriales = False
118             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If
        End If
    
120     If ObjData(ItemIndex).PielOsoPolaR > 0 Then
122         If Not TieneObjetos(PieldeOsoPolar, ObjData(ItemIndex).PielOsoPolaR, UserIndex) Then
124             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de oso polar.", e_FontTypeNames.FONTTYPE_INFO)
126             SastreTieneMateriales = False
128             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If
        End If
        
130     If ObjData(ItemIndex).PielLoboNegro > 0 Then
132         If Not TieneObjetos(PielLoboNegro, ObjData(ItemIndex).PielLoboNegro, UserIndex) Then
134             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de lobo negro.", e_FontTypeNames.FONTTYPE_INFO)
136             SastreTieneMateriales = False
138             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If
        End If


141     If ObjData(ItemIndex).PielTigre > 0 Then
142         If Not TieneObjetos(PielTigre, ObjData(ItemIndex).PielTigre, UserIndex) Then
143             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de tigre.", e_FontTypeNames.FONTTYPE_INFO)
144             SastreTieneMateriales = False
145             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If
        End If
        
146             If ObjData(ItemIndex).PielTigreBengala > 0 Then
147         If Not TieneObjetos(PielTigreBengala, ObjData(ItemIndex).PielTigreBengala, UserIndex) Then
148             Call WriteConsoleMsg(UserIndex, "No tenes suficientes pieles de tigre de bengala.", e_FontTypeNames.FONTTYPE_INFO)
149             SastreTieneMateriales = False
150             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If
        End If
154     SastreTieneMateriales = True

        
        Exit Function

SastreTieneMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.SastreTieneMateriales", Erl)
End Function

Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo HerreroTieneMateriales_Err
        

100     If ObjData(ItemIndex).LingH > 0 Then
102         If Not TieneObjetos(LingoteHierro, ObjData(ItemIndex).LingH, UserIndex) Then
104             Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de hierro.", e_FontTypeNames.FONTTYPE_INFO)
106             HerreroTieneMateriales = False
108             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If

110     If ObjData(ItemIndex).LingP > 0 Then
112         If Not TieneObjetos(LingotePlata, ObjData(ItemIndex).LingP, UserIndex) Then
114             Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de plata.", e_FontTypeNames.FONTTYPE_INFO)
116             HerreroTieneMateriales = False
118             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If

120     If ObjData(ItemIndex).LingO > 0 Then
122         If Not TieneObjetos(LingoteOro, ObjData(ItemIndex).LingO, UserIndex) Then
124             Call WriteConsoleMsg(UserIndex, "No tenes suficientes lingotes de oro.", e_FontTypeNames.FONTTYPE_INFO)
126             HerreroTieneMateriales = False
128             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function

            End If

        End If
        
130     If ObjData(ItemIndex).Coal > 0 Then
132         If Not TieneObjetos(e_Minerales.Coal, ObjData(ItemIndex).Coal, UserIndex) Then
134             Call WriteConsoleMsg(UserIndex, "No tenes suficientes carbon.", e_FontTypeNames.FONTTYPE_INFO)
136             HerreroTieneMateriales = False
138             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Function
            End If
        End If

140     HerreroTieneMateriales = True

        
        Exit Function

HerreroTieneMateriales_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroTieneMateriales", Erl)
End Function

Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruir_Err
        
100     PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex) And UserList(UserIndex).Stats.UserSkills(e_Skill.Herreria) >= ObjData(ItemIndex).SkHerreria

        
        Exit Function

PuedeConstruir_Err:
102     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruir", Erl)
104
        
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirHerreria_Err
        

        Dim i As Long

100     For i = 1 To UBound(ArmasHerrero)

102         If ArmasHerrero(i) = ItemIndex Then
104             PuedeConstruirHerreria = True
                Exit Function

            End If

106     Next i

108     For i = 1 To UBound(ArmadurasHerrero)

110         If ArmadurasHerrero(i) = ItemIndex Then
112             PuedeConstruirHerreria = True
                Exit Function

            End If

114     Next i

116     PuedeConstruirHerreria = False

        
        Exit Function

PuedeConstruirHerreria_Err:
118     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirHerreria", Erl)
120
        
End Function

Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo HerreroConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub
        
        If Not HayLugarEnInventario(UserIndex, ItemIndex, 1) Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente espacio en el inventario", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
102     If UserList(UserIndex).flags.Privilegios And (e_PlayerType.Consejero) Then
            Exit Sub
        End If
        
104     If PuedeConstruir(UserIndex, ItemIndex) And PuedeConstruirHerreria(ItemIndex) Then
106         Call HerreroQuitarMateriales(UserIndex, ItemIndex)
108         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
110         Call WriteUpdateSta(UserIndex)
            ' AGREGAR FX
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 253, 25, False, ObjData(ItemIndex).GrhIndex))
112         If ObjData(ItemIndex).OBJType = e_OBJType.otWeapon Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el arma!", e_FontTypeNames.FONTTYPE_INFO)
114             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
116         ElseIf ObjData(ItemIndex).OBJType = e_OBJType.otEscudo Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el escudo!", e_FontTypeNames.FONTTYPE_INFO)
118             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
120         ElseIf ObjData(ItemIndex).OBJType = e_OBJType.otCasco Then
                ' Call WriteConsoleMsg(UserIndex, "Has construido el casco!", e_FontTypeNames.FONTTYPE_INFO)
122             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)
124         ElseIf ObjData(ItemIndex).OBJType = e_OBJType.otArmadura Then
                'Call WriteConsoleMsg(UserIndex, "Has construido la armadura!", e_FontTypeNames.FONTTYPE_INFO)
126             Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)

            End If

            Dim MiObj As t_Obj

128         MiObj.amount = 1
130         MiObj.ObjIndex = ItemIndex

132         Call MeterItemEnInventario(UserIndex, MiObj)
    
136         Call SubirSkill(UserIndex, e_Skill.Herreria)
138         Call UpdateUserInv(True, UserIndex, 0)
140         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))

142         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

HerreroConstruirItem_Err:
144     Call TraceError(Err.Number, Err.Description, "Trabajo.HerreroConstruirItem", Erl)
146
        
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirCarpintero_Err
        

        Dim i As Long

100     For i = 1 To UBound(ObjCarpintero)

102         If ObjCarpintero(i) = ItemIndex Then
104             PuedeConstruirCarpintero = True
                Exit Function

            End If

106     Next i

108     PuedeConstruirCarpintero = False

        
        Exit Function

PuedeConstruirCarpintero_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirCarpintero", Erl)
112
        
End Function

Public Function PuedeConstruirAlquimista(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirAlquimista_Err
        

        Dim i As Long
              

100     For i = 1 To UBound(ObjAlquimista)

102         If ObjAlquimista(i) = ItemIndex Then
104             PuedeConstruirAlquimista = True
                Exit Function

            End If

106     Next i

108     PuedeConstruirAlquimista = False

        
        Exit Function

PuedeConstruirAlquimista_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirAlquimista", Erl)
112
        
End Function

Public Function PuedeConstruirSastre(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo PuedeConstruirSastre_Err
        

        Dim i As Long

100     For i = 1 To UBound(ObjSastre)

102         If ObjSastre(i) = ItemIndex Then
104             PuedeConstruirSastre = True
                Exit Function

            End If

106     Next i

108     PuedeConstruirSastre = False

        
        Exit Function

PuedeConstruirSastre_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeConstruirSastre", Erl)
112
        
End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal cantidad As Long, ByVal cantidad_maxima As Integer)
        
        On Error GoTo CarpinteroConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub

102     If UserList(userindex).flags.Privilegios And (e_PlayerType.Consejero Or e_PlayerType.SemiDios Or e_PlayerType.Dios) Then
            Exit Sub
        End If
        
104     If ItemIndex = 0 Then Exit Sub

        'Si no tiene equipado el serrucho
106     If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
            ' Antes de usar la herramienta deberias equipartela.
108         Call WriteLocaleMsg(UserIndex, "376", e_FontTypeNames.FONTTYPE_INFO)
110         Call WriteMacroTrabajoToggle(UserIndex, False)
            Exit Sub
        End If
        
        Dim cantidad_a_construir As Long
        Dim madera_requerida As Long
        cantidad_a_construir = IIf(UserList(UserIndex).Trabajo.cantidad >= cantidad_maxima, cantidad_maxima, UserList(UserIndex).Trabajo.cantidad)
        
        If cantidad_a_construir <= 0 Then
121             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
        End If
        
112     If CarpinteroTieneMateriales(UserIndex, ItemIndex, cantidad_a_construir) _
                And UserList(UserIndex).Stats.UserSkills(e_Skill.Carpinteria) >= ObjData(ItemIndex).SkCarpinteria _
                And PuedeConstruirCarpintero(ItemIndex) _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).OBJType = e_OBJType.otHerramientas _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo = 5 Then
    
114         If UserList(UserIndex).Stats.MinSta > 2 Then
116             Call QuitarSta(UserIndex, 2)
        
            Else
118             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para trabajar.", e_FontTypeNames.FONTTYPE_INFO)
120             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If
            
            If ObjData(ItemIndex).Madera > 0 Then
                madera_requerida = ObjData(ItemIndex).Madera
            ElseIf ObjData(ItemIndex).MaderaElfica > 0 Then
                madera_requerida = ObjData(ItemIndex).MaderaElfica
            ElseIf ObjData(ItemIndex).MaderaPino > 0 Then
                madera_requerida = ObjData(ItemIndex).MaderaPino
            End If
            
    
122         Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, madera_requerida * cantidad_a_construir)

            UserList(UserIndex).Trabajo.cantidad = UserList(UserIndex).Trabajo.cantidad - cantidad_a_construir
            
            
            'Call WriteConsoleMsg(UserIndex, "Has construido un objeto!", e_FontTypeNames.FONTTYPE_INFO)
            'Call WriteOroOverHead(UserIndex, 1, UserList(UserIndex).Char.CharIndex)
124         Call WriteTextCharDrop(UserIndex, "+" & cantidad_a_construir, UserList(UserIndex).Char.charindex, vbWhite)
    
            Dim MiObj As t_Obj

126         MiObj.amount = cantidad_a_construir
128         MiObj.ObjIndex = ItemIndex
             ' AGREGAR FX
130         Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.CharIndex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
132         If Not MeterItemEnInventario(UserIndex, MiObj) Then
134             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If
    
   
136         Call SubirSkill(UserIndex, e_Skill.Carpinteria)
            'Call UpdateUserInv(True, UserIndex, 0)
138         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))

140         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If

        
        Exit Sub

CarpinteroConstruirItem_Err:
142     Call TraceError(Err.Number, Err.Description, "Trabajo.CarpinteroConstruirItem", Erl)

        
End Sub

Public Sub AlquimistaConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo AlquimistaConstruirItem_Err
        

        Rem Debug.Print UserList(UserIndex).Invent.HerramientaEqpObjIndex

100     If Not UserList(UserIndex).Stats.MinSta > 0 Then
102         Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

104     If AlquimistaTieneMateriales(UserIndex, ItemIndex) _
                And UserList(UserIndex).Stats.UserSkills(e_Skill.Alquimia) >= ObjData(ItemIndex).SkPociones _
                And PuedeConstruirAlquimista(ItemIndex) _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).OBJType = e_OBJType.otHerramientas _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo = 4 Then
            
            'si tiene el hechizoAlquimista salgo
            Dim hIndex As Integer

            hIndex = ObjData(ItemIndex).Hechizo  'cambiar por la variable del Obj en Hechizo

            If TieneHechizo(hIndex, UserIndex) Then

106             UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 1
108             Call WriteUpdateSta(UserIndex)
                
                ' AGREGAR FX
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, 253, 25, False, ObjData(ItemIndex).GrhIndex))
110             Call AlquimistaQuitarMateriales(UserIndex, ItemIndex)
                'Call WriteConsoleMsg(UserIndex, "Has construido el objeto.", e_FontTypeNames.FONTTYPE_INFO)
112             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(1152, UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        
                Dim MiObj As t_Obj
    
114             MiObj.amount = 1
116             MiObj.ObjIndex = ItemIndex
    
118             If Not MeterItemEnInventario(UserIndex, MiObj) Then
120                 Call TirarItemAlPiso(UserList(UserIndex).pos, MiObj)
    
                End If
        
    
122             Call SubirSkill(UserIndex, e_Skill.Alquimia)
124             Call UpdateUserInv(True, UserIndex, 0)
                'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
126             UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
            Else
                Call WriteConsoleMsg(UserIndex, "Lamentablemente no aprendiste la receta para crear esta poción.", e_FontTypeNames.FONTTYPE_INFOBOLD)
            End If
        End If

        
        Exit Sub

AlquimistaConstruirItem_Err:
128     Call TraceError(Err.Number, Err.Description, "Trabajo.AlquimistaConstruirItem", Erl)
130
        
End Sub

Public Sub SastreConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
        
        On Error GoTo SastreConstruirItem_Err
        
100     If Not IntervaloPermiteTrabajarConstruir(UserIndex) Then Exit Sub

102     If Not UserList(UserIndex).Stats.MinSta > 0 Then
104         Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        If ItemIndex = 0 Then Exit Sub
        If UserList(UserIndex).invent.HerramientaEqpObjIndex = 0 Then
            Exit Sub
        End If
106     If SastreTieneMateriales(UserIndex, ItemIndex) _
                And UserList(UserIndex).Stats.UserSkills(e_Skill.Sastreria) >= ObjData(ItemIndex).SkSastreria _
                And PuedeConstruirSastre(ItemIndex) _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).OBJType = e_OBJType.otHerramientas _
                And ObjData(UserList(UserIndex).Invent.HerramientaEqpObjIndex).Subtipo = 9 Then
        
108         UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - 2
        
110         Call WriteUpdateSta(UserIndex)
    
112         Call SastreQuitarMateriales(UserIndex, ItemIndex)
    
114         Call WriteTextCharDrop(UserIndex, "+1", UserList(UserIndex).Char.CharIndex, vbWhite)

116         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(63, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
    
            Dim MiObj As t_Obj

118         MiObj.amount = 1
120         MiObj.ObjIndex = ItemIndex

122         If Not MeterItemEnInventario(UserIndex, MiObj) Then
124             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
            End If
    
126         Call SubirSkill(UserIndex, e_Skill.Sastreria)
128         Call UpdateUserInv(True, UserIndex, 0)

130         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

        End If
    
        
        Exit Sub

SastreConstruirItem_Err:
132     Call TraceError(Err.Number, Err.Description, "Trabajo.SastreConstruirItem", Erl)
134
        
End Sub

Private Function MineralesParaLingote(ByVal Lingote As e_Minerales, ByVal cant As Byte) As Integer
        
        On Error GoTo MineralesParaLingote_Err
        

100     Select Case Lingote

            Case e_Minerales.HierroCrudo
102             MineralesParaLingote = 13 * cant

104         Case e_Minerales.PlataCruda
106             MineralesParaLingote = 25 * cant

108         Case e_Minerales.OroCrudo
110             MineralesParaLingote = 50 * cant

112         Case Else
114             MineralesParaLingote = 10000

        End Select

        
        Exit Function

MineralesParaLingote_Err:
116     Call TraceError(Err.Number, Err.Description, "Trabajo.MineralesParaLingote", Erl)
118
        
End Function

Public Sub DoLingotes(ByVal UserIndex As Integer)
            On Error GoTo DoLingotes_Err

            Dim Slot As Integer
            Dim obji As Integer
            Dim cant As Byte
            Dim necesarios As Integer

100         If UserList(UserIndex).Stats.MinSta > 2 Then
102             Call QuitarSta(UserIndex, 2)

            Else
104             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", e_FontTypeNames.FONTTYPE_INFO)
106             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If

108         Slot = UserList(UserIndex).flags.TargetObjInvSlot
110         obji = UserList(UserIndex).Invent.Object(Slot).ObjIndex

112         cant = RandomNumber(10, 20)
114         necesarios = MineralesParaLingote(obji, cant)

116         If UserList(UserIndex).Invent.Object(Slot).amount < MineralesParaLingote(obji, cant) Or ObjData(obji).OBJType <> e_OBJType.otMinerales Then
118             Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", e_FontTypeNames.FONTTYPE_INFO)
120             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If

122         UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount - MineralesParaLingote(obji, cant)

124         If UserList(UserIndex).Invent.Object(Slot).amount < 1 Then
126             UserList(UserIndex).Invent.Object(Slot).amount = 0
128             UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0

            End If

            Dim nPos  As t_WorldPos

            Dim MiObj As t_Obj

130         MiObj.amount = cant
132         MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObjInvIndex).LingoteIndex

134         If Not MeterItemEnInventario(UserIndex, MiObj) Then
136             Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)

            End If

138         Call UpdateUserInv(False, UserIndex, Slot)
140         Call WriteTextCharDrop(UserIndex, "+" & cant, UserList(UserIndex).Char.CharIndex, vbWhite)
142         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(41, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.y))
144         Call SubirSkill(UserIndex, e_Skill.Mineria)

146         UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

148         If UserList(UserIndex).Counters.Trabajando = 1 And Not UserList(UserIndex).flags.UsandoMacro Then
150             Call WriteMacroTrabajoToggle(UserIndex, True)

            End If

            Exit Sub

DoLingotes_Err:
152         Call TraceError(Err.Number, Err.Description, "Trabajo.DoLingotes", Erl)
154

End Sub

Function ModAlquimia(ByVal clase As e_Class) As Integer
        On Error GoTo ModAlquimia_Err

100     Select Case clase

            Case e_Class.Druid
102             ModAlquimia = 1

104         Case e_Class.Trabajador
106             ModAlquimia = 1

108         Case Else
110             ModAlquimia = 3

        End Select

        Exit Function

ModAlquimia_Err:
112     Call TraceError(Err.Number, Err.Description, "Trabajo.ModAlquimia", Erl)
114

End Function

Function ModSastre(ByVal clase As e_Class) As Integer
        
        On Error GoTo ModSastre_Err
        

100     Select Case clase

            Case e_Class.Trabajador
102             ModSastre = 1

104         Case Else
106             ModSastre = 3

        End Select

        
        Exit Function

ModSastre_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ModSastre", Erl)
110
        
End Function

Function ModCarpinteria(ByVal clase As e_Class) As Integer
        
        On Error GoTo ModCarpinteria_Err
        

100     Select Case clase

            Case e_Class.Trabajador
102             ModCarpinteria = 1

104         Case Else
106             ModCarpinteria = 3

        End Select

        
        Exit Function

ModCarpinteria_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ModCarpinteria", Erl)
110
        
End Function

Function ModHerreria(ByVal clase As e_Class) As Single
        
        On Error GoTo ModHerreriA_Err
        

100     Select Case clase

            Case e_Class.Trabajador
102             ModHerreria = 1

104         Case Else
106             ModHerreria = 3

        End Select

        
        Exit Function

ModHerreriA_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ModHerreriA", Erl)
110
        
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer, Optional ByVal invisible As Byte = 2)
        
        On Error GoTo DoAdminInvisible_Err
        
    
100     With UserList(UserIndex)

        If invisible = 2 Then
           .flags.AdminInvisible = IIf(.flags.AdminInvisible = 1, 0, 1)
        Else
            .flags.AdminInvisible = invisible
        End If
        
102     If .flags.AdminInvisible = 1 Then
            
106         .flags.invisible = 1
108         .flags.Oculto = 1
                    
110         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))
            
112         Call SendData(SendTarget.ToPCAreaButGMs, UserIndex, PrepareMessageCharacterRemove(2, .Char.charindex, True))
        
        Else
    
116         .flags.invisible = 0
118         .flags.Oculto = 0
120         .Counters.TiempoOculto = 0
        
122         Call MakeUserChar(True, 0, UserIndex, .Pos.map, .Pos.X, .Pos.y, 1)
124         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))
        
        End If
            
        End With

        Exit Sub

DoAdminInvisible_Err:
126     Call TraceError(Err.Number, Err.Description, "Trabajo.DoAdminInvisible", Erl)

128
        
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)
        
        On Error GoTo TratarDeHacerFogata_Err
        

        Dim Suerte    As Byte

        Dim exito     As Byte

        Dim obj       As t_Obj

        Dim posMadera As t_WorldPos

100     If Not LegalPos(Map, X, Y) Then Exit Sub

102     With posMadera
104         .Map = Map
106         .X = X
108         .Y = Y

        End With

110     If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
112         Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre Leña para hacer ramitas", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

114     If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
116         Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
            ' Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

118     If UserList(UserIndex).flags.Muerto = 1 Then
120         Call WriteConsoleMsg(UserIndex, "No podés hacer fogatas estando muerto.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

122     If MapData(Map, X, Y).ObjInfo.amount < 3 Then
124         Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

126     If UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) < 6 Then
128         Suerte = 3
130     ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) <= 34 Then
132         Suerte = 2
134     ElseIf UserList(UserIndex).Stats.UserSkills(e_Skill.Supervivencia) >= 35 Then
136         Suerte = 1

        End If

138     exito = RandomNumber(1, Suerte)

140     If exito = 1 Then
142         obj.ObjIndex = FOGATA_APAG
144         obj.amount = MapData(Map, X, Y).ObjInfo.amount \ 3
    
146         Call WriteConsoleMsg(UserIndex, "Has hecho " & obj.amount & " ramitas.", e_FontTypeNames.FONTTYPE_INFO)
    
148         Call MakeObj(obj, Map, X, Y)
    
            'Seteamos la fogata como el nuevo TargetObj del user
150         UserList(UserIndex).flags.TargetObj = FOGATA_APAG

        End If

152     Call SubirSkill(UserIndex, Supervivencia)

        
        Exit Sub

TratarDeHacerFogata_Err:
154     Call TraceError(Err.Number, Err.Description, "Trabajo.TratarDeHacerFogata", Erl)
156
        
End Sub

Public Sub DoPescar(ByVal UserIndex As Integer, Optional ByVal RedDePesca As Boolean = False)

On Error GoTo ErrHandler
    Dim bonificacionPescaLvl(1 To 47) As Single
    Dim bonificacionCaña As Double
    Dim bonificacionZona As Double
    Dim bonificacionLvl As Double
    Dim bonificacionClase As Double
    Dim bonificacionTotal As Double
    Dim RestaStamina As Integer
    Dim Reward As Double
    Dim esEspecial As Boolean
    Dim i As Integer
    
    
        With UserList(UserIndex)
        
            RestaStamina = IIf(RedDePesca, 12, RandomNumber(2, 3))
104         If .flags.Privilegios And (e_PlayerType.Consejero) Then
                Exit Sub
            End If
            
106         If .Stats.MinSta > RestaStamina Then
108             Call QuitarSta(UserIndex, RestaStamina)
            Else
110             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
112             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
            End If
            
            If MapInfo(.Pos.Map).Seguro = 1 Or Zona(.ZonaId).Segura = 1 Then
120             Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageArmaMov(.Char.charindex, 0))
            Else
                Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageArmaMov(.Char.charindex, 0))
            End If

            bonificacionPescaLvl(1) = 0
            bonificacionPescaLvl(2) = 0.009
            bonificacionPescaLvl(3) = 0.015
            bonificacionPescaLvl(4) = 0.019
            bonificacionPescaLvl(5) = 0.025
            bonificacionPescaLvl(6) = 0.03
            bonificacionPescaLvl(7) = 0.035
            bonificacionPescaLvl(8) = 0.04
            bonificacionPescaLvl(9) = 0.045
            bonificacionPescaLvl(10) = 0.05
            bonificacionPescaLvl(11) = 0.06
            bonificacionPescaLvl(12) = 0.07
            bonificacionPescaLvl(13) = 0.08
            bonificacionPescaLvl(14) = 0.09
            bonificacionPescaLvl(15) = 0.1
            bonificacionPescaLvl(16) = 0.11
            bonificacionPescaLvl(17) = 0.13
            bonificacionPescaLvl(18) = 0.14
            bonificacionPescaLvl(19) = 0.16
            bonificacionPescaLvl(20) = 0.18
            bonificacionPescaLvl(21) = 0.2
            bonificacionPescaLvl(22) = 0.22
            bonificacionPescaLvl(23) = 0.24
            bonificacionPescaLvl(24) = 0.27
            bonificacionPescaLvl(25) = 0.3
            bonificacionPescaLvl(26) = 0.32
            bonificacionPescaLvl(27) = 0.35
            bonificacionPescaLvl(28) = 0.37
            bonificacionPescaLvl(29) = 0.4
            bonificacionPescaLvl(30) = 0.43
            bonificacionPescaLvl(31) = 0.47
            bonificacionPescaLvl(32) = 0.51
            bonificacionPescaLvl(33) = 0.55
            bonificacionPescaLvl(34) = 0.58
            bonificacionPescaLvl(35) = 0.62
            bonificacionPescaLvl(36) = 0.7
            bonificacionPescaLvl(37) = 0.77
            bonificacionPescaLvl(38) = 0.84
            bonificacionPescaLvl(39) = 0.92
            bonificacionPescaLvl(40) = 1#
            bonificacionPescaLvl(41) = 1.1
            bonificacionPescaLvl(42) = 1.15
            bonificacionPescaLvl(43) = 1.3
            bonificacionPescaLvl(44) = 1.5
            bonificacionPescaLvl(45) = 1.8
            bonificacionPescaLvl(46) = 2#
            bonificacionPescaLvl(47) = 2.5
        
                If MapInfo(.Pos.Map).Seguro Or RedDePesca Or Zona(.ZonaId).Segura Then
                    Select Case ObjData(.Invent.HerramientaEqpObjIndex).Power
                        Case 1 'Caña comun
                            bonificacionCaña = 1
                        Case 2 'Caña reforzada
                            bonificacionCaña = 1.5
                        Case 3 'Caña especial
                            bonificacionCaña = 1.9
                        Case 4 'Caña de plata
                            bonificacionCaña = 2.2
                        Case 5 'Red de pesca
                            bonificacionCaña = 6.5
                        Case 6 'Red lisa
                            bonificacionCaña = 9
                    End Select
                Else
                    Select Case ObjData(.Invent.HerramientaEqpObjIndex).Power
                        Case 1 'Caña comun
                            bonificacionCaña = 1.3
                        Case 2 'Caña reforzada
                            bonificacionCaña = 1.65
                        Case 3 'Caña especial
                            bonificacionCaña = 3
                        Case 4 'Caña de plata
                            bonificacionCaña = 6
                    End Select
                End If
                
                
                bonificacionLvl = 1 + bonificacionPescaLvl(.Stats.ELV) 'Segun el nivel se le bonifica extra
                bonificacionClase = IIf(.clase = Trabajador, 1, RandomNumber(1, 3) / 3) 'Si no es pescador va a pescar menos al azar.
                
                bonificacionTotal = bonificacionCaña * bonificacionLvl * bonificacionClase * RecoleccionMult
                
                'Calculo el botin esperado por iteracción. 'La base del calculo son 8000 por hora + 20% de chances de no pescar + un +/- 10%
                Reward = (IntervaloTrabajarExtraer / 3600000) * 8000 * bonificacionTotal * 1.2 * (1 + (RandomNumber(0, 20) - 10) / 100)
                
                
                'Calculo la suerte de pescar o no pescar y aplico eso sobre el reward para promediar.
                Dim Suerte As Integer
                Dim Pesco As Boolean
                If .Stats.UserSkills(e_Skill.Pescar) < 20 Then
                    Suerte = 20
                ElseIf .Stats.UserSkills(e_Skill.Pescar) < 40 Then
                    Suerte = 35
                ElseIf .Stats.UserSkills(e_Skill.Pescar) < 70 Then
                    Suerte = 55
                ElseIf .Stats.UserSkills(e_Skill.Pescar) < 100 Then
                    Suerte = 68
                Else
                    Suerte = 80
                End If
                
                Pesco = RandomNumber(1, 100) <= Suerte '80% de posibilidad de pescar
            
                If Pesco Then
            
                    
                    
                    Dim nPos  As t_WorldPos
                    Dim MiObj As t_Obj
                    Dim objValue As Integer
                    
124                 MiObj.objIndex = ObtenerPezRandom(ObjData(.Invent.HerramientaEqpObjIndex).Power)
125                 objValue = max(ObjData(MiObj.objIndex).Valor / 3, 1)
126                 MiObj.amount = Round(Reward / objValue)
                    If MiObj.amount <= 0 Then
                        MiObj.amount = 1
                    End If
                    Dim StopWorking As Boolean
                    StopWorking = False
                    If MapInfo(UserList(UserIndex).pos.map).Seguro = 0 Then
134                     If MiObj.amount > MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount Then
136                         MiObj.amount = MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount
                            Call CreateFishingPool(.pos.Map)
                            Call EraseObj(MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount, .pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y)
137                         Call WriteConsoleMsg(UserIndex, "No hay mas peces aqui.", e_FontTypeNames.FONTTYPE_INFO)
                            StopWorking = True
                        End If
                    End If
                    If FISHING_POOL_ID = MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.ObjIndex Then
138                     MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount = MapData(.pos.Map, .Trabajo.Target_X, .Trabajo.Target_Y).ObjInfo.amount - MiObj.amount
                    End If
                    If Not RedDePesca Then
                        esEspecial = False
                        For i = 1 To UBound(PecesEspeciales)
                            If PecesEspeciales(i).objIndex = MiObj.objIndex Then
                                esEspecial = True
                            End If
                        Next i
                    End If
                    
                    If Not esEspecial Then
                        Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, ObjData(MiObj.objIndex).GrhIndex))
                    ElseIf Not RedDePesca Then
                        .flags.PescandoEspecial = True
156                     Call WriteMacroTrabajoToggle(UserIndex, False)
                        .Stats.NumObj_PezEspecial = MiObj.objIndex
                        Call WritePelearConPezEspecial(UserIndex)
                        Exit Sub
                    End If
                    
158                 If MiObj.objIndex = 0 Then Exit Sub
            
160                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
162                     Call TirarItemAlPiso(.pos, MiObj)
                    End If
    
164                 Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.charindex, vbWhite)
166                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .pos.x, .pos.y))
                   
                    ' Al pescar también podés sacar cosas raras (se setean desde RecursosEspeciales.dat)
                    ' Por cada drop posible
                    Dim res As Long
168                 For i = 1 To UBound(EspecialesPesca)
                        ' Tiramos al azar entre 1 y la probabilidad
170                     res = RandomNumber(1, IIf(RedDePesca, EspecialesPesca(i).Data * 2, EspecialesPesca(i).Data)) ' Red de pesca chance x2 (revisar)
                
                        ' Si tiene suerte y le pega
172                     If res = 1 Then
174                         MiObj.objIndex = EspecialesPesca(i).objIndex
176                         MiObj.amount = 1 ' Solo un item por vez
                    
178                         If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.pos, MiObj)
                        
                            ' Le mandamos un mensaje
180                         Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(EspecialesPesca(i).ObjIndex).name & "!", e_FontTypeNames.FONTTYPE_INFO)
                        End If
    
                    Next
    
                Else
                     Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.charindex, 253, 25, False, GRH_FALLO_PESCA))
                End If
    
182             Call SubirSkill(UserIndex, e_Skill.Pescar)
                If StopWorking Then
184             Call WriteWorkRequestTarget(UserIndex, 0)
186                 Call WriteMacroTrabajoToggle(UserIndex, False)
                    Exit Sub
                End If
    
188             .Counters.Trabajando = .Counters.Trabajando + 1
                .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)
    
                'Ladder 06/07/14 Activamos el macro de trabajo
190             If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
                    Call WriteMacroTrabajoToggle(UserIndex, True)
                End If
        
        
    End With
    Exit Sub

ErrHandler:
192     Call LogError("Error en DoPescar. Error " & Err.Number & " - " & Err.Description & " Line number: " & Erl)

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadronIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadronIndex As Integer, ByVal VictimaIndex As Integer)
        '*************************************************
        'Author: Unknown
        'Last modified: 05/04/2010
        'Last Modification By: ZaMa
        '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadronIndex) when the thief stoles gold. (MarKoxX)
        '27/11/2009: ZaMa - Optimizacion de codigo.
        '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
        '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
        '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
        '23/04/2010: ZaMa - No se puede robar mas sin energia.
        '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
        '*************************************************

        On Error GoTo ErrHandler

        Dim OtroUserIndex As Integer
        
100     If UserList(LadronIndex).flags.Privilegios And (e_PlayerType.Consejero) Then Exit Sub

102     If MapInfo(UserList(VictimaIndex).Pos.Map).Seguro = 1 Then Exit Sub
        
        If Not UserMod.CanMove(UserList(VictimaIndex).flags, UserList(VictimaIndex).Counters) Then
            Call WriteConsoleMsg(LadronIndex, "No podes robarle a objetivos inmovilizados.", e_FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
    
104     If UserList(VictimaIndex).flags.EnConsulta Then
106         Call WriteConsoleMsg(LadronIndex, "¡No puedes robar a usuarios en consulta!", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim Penable As Boolean
        
108     With UserList(LadronIndex)

            If esCiudadano(LadronIndex) Then
                If (.flags.Seguro) Then
114                     Call WriteConsoleMsg(LadronIndex, "Debes quitarte el seguro para robarle a un ciudadano o a un miembro del Ejército Real", e_FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            ElseIf esArmada(LadronIndex) Then ' Armada robando a armada or ciudadano?
122              If (esCiudadano(VictimaIndex) Or esArmada(VictimaIndex)) Then
124                 Call WriteConsoleMsg(LadronIndex, "Los miembros del Ejército Real no tienen permitido robarle a ciudadanos o a otros miembros del Ejército Real", e_FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            ElseIf esCaos(LadronIndex) Then ' Caos robando a caos?
                If (esCaos(VictimaIndex)) Then
                    Call WriteConsoleMsg(LadronIndex, "No puedes robar a otros miembros de la Legión Oscura.", e_FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
            
            'Me fijo si el ladrón tiene clan
            If .GuildIndex > 0 Then
                'Si tiene clan me fijo si su clan es de alineación ciudadana
                If esCiudadano(LadronIndex) And GuildAlignmentIndex(.GuildIndex) = e_ALINEACION_GUILD.ALINEACION_CIUDADANA Then
                    If PersonajeEsLeader(.Id) Then
                        Call WriteConsoleMsg(LadronIndex, "No puedes robar siendo lider de un clan ciudadano.", e_FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
            End If

126         If TriggerZonaPelea(LadronIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
            ' Tiene energia?
128         If .Stats.MinSta < 15 Then
        
130             If .genero = e_Genero.Hombre Then
132                 Call WriteConsoleMsg(LadronIndex, "Estás muy cansado para robar.", e_FontTypeNames.FONTTYPE_INFO)
                
                Else
134                 Call WriteConsoleMsg(LadronIndex, "Estás muy cansada para robar.", e_FontTypeNames.FONTTYPE_INFO)

                End If
            
                Exit Sub

            End If
        
136         If .GuildIndex > 0 Then
                
138             If .flags.SeguroClan And NivelDeClan(.GuildIndex) >= 3 Then
            
140                 If .GuildIndex = UserList(VictimaIndex).GuildIndex Then
142                     Call WriteConsoleMsg(LadronIndex, "No podes robarle a un miembro de tu clan.", e_FontTypeNames.FONTTYPE_INFOIAO)
                        Exit Sub

                    End If

                End If

            End If

            ' Quito energia
144         Call QuitarSta(LadronIndex, 15)

146         If UserList(VictimaIndex).flags.Privilegios And e_PlayerType.user Then

                Dim Probabilidad As Byte

                Dim res          As Integer

                Dim RobarSkill   As Byte
            
148             RobarSkill = .Stats.UserSkills(e_Skill.Robar)

150             If (RobarSkill > 0 And RobarSkill < 10) Then
152                 Probabilidad = 1
154             ElseIf (RobarSkill >= 10 And RobarSkill <= 20) Then
156                 Probabilidad = 5
158             ElseIf (RobarSkill >= 20 And RobarSkill <= 30) Then
160                 Probabilidad = 10
162             ElseIf (RobarSkill >= 30 And RobarSkill <= 40) Then
164                 Probabilidad = 15
166             ElseIf (RobarSkill >= 40 And RobarSkill <= 50) Then
168                 Probabilidad = 25
170             ElseIf (RobarSkill >= 50 And RobarSkill <= 60) Then
172                 Probabilidad = 35
174             ElseIf (RobarSkill >= 60 And RobarSkill <= 70) Then
176                 Probabilidad = 40
178             ElseIf (RobarSkill >= 70 And RobarSkill <= 80) Then
180                 Probabilidad = 55
182             ElseIf (RobarSkill >= 80 And RobarSkill <= 90) Then
184                 Probabilidad = 70
186             ElseIf (RobarSkill >= 90 And RobarSkill < 100) Then
188                 Probabilidad = 80
190             ElseIf (RobarSkill = 100) Then
192                 Probabilidad = 90
                End If

194             If (RandomNumber(1, 100) < Probabilidad) Then 'Exito robo
                
196                 If UserList(VictimaIndex).flags.Comerciando Then
198                     OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu.ArrayIndex
                        
200                     If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
202                         Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, ¡te están robando!", e_FontTypeNames.FONTTYPE_TALK)
204                         Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado, al otro usuario le robaron.", e_FontTypeNames.FONTTYPE_TALK)
                        
206                         Call LimpiarComercioSeguro(VictimaIndex)

                        End If

                    End If
               
208                 If (RandomNumber(1, 50) < 25) And (.clase = e_Class.Thief) Then '50% de robar items
                    
210                     If TieneObjetosRobables(VictimaIndex) Then
212                         Call RobarObjeto(LadronIndex, VictimaIndex)
                        Else
214                         Call WriteConsoleMsg(LadronIndex, UserList(VictimaIndex).Name & " no tiene objetos.", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else '50% de robar oro

216                     If UserList(VictimaIndex).Stats.GLD > 0 Then

                            Dim n     As Long

                            Dim Extra As Single
                            
                            ' Multiplicador extra por niveles
                            
                            Select Case .Stats.ELV
                                Case Is < 13
                                    Extra = 1
                                Case Is < 25
                                    Extra = 1.1
                                Case Is < 35
                                    Extra = 1.2
                                Case Is >= 35 And .Stats.ELV <= 40
                                    Extra = 1.3
                                Case Is >= 41 And .Stats.ELV < 45
                                    Extra = 1.4
                                Case Is >= 45 And .Stats.ELV <= 46
                                    Extra = 1.5
                                Case Is = 47
                                    Extra = 5
                            End Select
                            
238                         If .clase = e_Class.Thief Then
                                'Si no tiene puestos los guantes de hurto roba un 50% menos.
240                             If .invent.WeaponEqpObjIndex > 0 Then
242                                 If ObjData(.invent.WeaponEqpObjIndex).Subtipo = 5 Then
244                                     n = RandomNumber(.Stats.ELV * 50 * Extra, .Stats.ELV * 100 * Extra) * OroMult
                                    Else
246                                     n = RandomNumber(.Stats.ELV * 25 * Extra, .Stats.ELV * 50 * Extra) * OroMult
                                    End If
                                Else
248                                 n = RandomNumber(.Stats.ELV * 25 * Extra, .Stats.ELV * 50 * Extra) * OroMult
                                End If
                            Else
250                             n = RandomNumber(1, 100) * OroMult
                            End If

252                         If n > UserList(VictimaIndex).Stats.GLD Then n = UserList(VictimaIndex).Stats.GLD
                            Dim prevGold As Long: prevGold = UserList(VictimaIndex).Stats.GLD
254                         UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - n
                            Dim ProtectedGold As Long
                            ProtectedGold = OroPorNivelBilletera * UserList(VictimaIndex).Stats.ELV
                            If prevGold >= ProtectedGold And UserList(VictimaIndex).Stats.GLD < ProtectedGold Then
                                n = prevGold - ProtectedGold
                                UserList(VictimaIndex).Stats.GLD = ProtectedGold
                            End If
256                         .Stats.GLD = .Stats.GLD + n

258                         If .Stats.GLD > MAXORO Then .Stats.GLD = MAXORO
                        
260                         Call WriteConsoleMsg(LadronIndex, "Le has robado " & PonerPuntos(n) & " monedas de oro a " & UserList(VictimaIndex).Name, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)
262                         Call WriteConsoleMsg(VictimaIndex, UserList(LadronIndex).Name & " te ha robado " & PonerPuntos(n) & " monedas de oro.", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)
264                         Call WriteUpdateGold(LadronIndex) 'Le actualizamos la billetera al ladron
                        
266                         Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                        Else
268                         Call WriteConsoleMsg(LadronIndex, UserList(VictimaIndex).Name & " no tiene oro.", e_FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If
                
270                 Call SubirSkill(LadronIndex, e_Skill.Robar)
            
                Else
272                 Call WriteConsoleMsg(LadronIndex, "¡No has logrado robar nada!", e_FontTypeNames.FONTTYPE_INFO)
274                 Call WriteConsoleMsg(VictimaIndex, "¡" & .name & " ha intentado robarte!", e_FontTypeNames.FONTTYPE_INFO)
                
276                 Call SubirSkill(LadronIndex, e_Skill.Robar)

                End If
                    
278             If Status(LadronIndex) = Ciudadano Then Call VolverCriminal(LadronIndex)

            End If

        End With

        Exit Sub

ErrHandler:
282     Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.Description)

End Sub

Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
        ' Agregué los barcos
        ' Agrego poción negra
        ' Esta funcion determina qué objetos son robables.
        
        On Error GoTo ObjEsRobable_Err
        

        Dim OI As Integer

100     OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

102     ObjEsRobable = ObjData(OI).OBJType <> e_OBJType.otLlaves And _
                       ObjData(OI).OBJType <> e_OBJType.otBarcos And _
                       ObjData(OI).OBJType <> e_OBJType.otMonturas And _
                       ObjData(OI).OBJType <> e_OBJType.otRunas And _
                       ObjData(OI).ObjDonador = 0 And _
                       ObjData(OI).Instransferible = 0 And _
                       ObjData(OI).Real = 0 And _
                       ObjData(OI).Caos = 0 And _
                       UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0

        
        Exit Function

ObjEsRobable_Err:
104     Call TraceError(Err.Number, Err.Description, "Trabajo.ObjEsRobable", Erl)
106
        
End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Private Sub RobarObjeto(ByVal LadronIndex As Integer, ByVal VictimaIndex As Integer)
        '***************************************************
        'Author: Unknown
        'Last Modification: 02/04/2010
        '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
        '***************************************************
        
        On Error GoTo RobarObjeto_Err
    
        

        Dim flag As Boolean
        Dim i    As Integer

100     flag = False

102     With UserList(VictimaIndex)

104         If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final del inventario?
106             i = 1

108             Do While Not flag And i <= .CurrentInventorySlots

                    'Hay objeto en este slot?
110                 If .Invent.Object(i).ObjIndex > 0 Then
                
112                     If ObjEsRobable(VictimaIndex, i) Then
                    
114                         If RandomNumber(1, 10) < 4 Then flag = True
                        
                        End If

                    End If

116                 If Not flag Then i = i + 1
                Loop
            Else
118             i = .CurrentInventorySlots

120             Do While Not flag And i > 0

                    'Hay objeto en este slot?
122                 If .Invent.Object(i).ObjIndex > 0 Then
                
124                     If ObjEsRobable(VictimaIndex, i) Then
                    
126                         If RandomNumber(1, 10) < 4 Then flag = True
                        
                        End If

                    End If

128                 If Not flag Then i = i - 1
                Loop

            End If
    
130         If flag Then

                Dim MiObj     As t_Obj
                Dim Num       As Integer
                Dim ObjAmount As Integer

132             ObjAmount = .Invent.Object(i).amount

                'Cantidad al azar entre el 3 y el 6% del total, con minimo 1.
134             Num = MaximoInt(1, RandomNumber(ObjAmount * 0.03, ObjAmount * 0.06))

136             MiObj.amount = Num
138             MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
140             .Invent.Object(i).amount = ObjAmount - Num
                    
142             If .Invent.Object(i).amount <= 0 Then
144                 Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

                End If
                
146             Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
148             If Not MeterItemEnInventario(LadronIndex, MiObj) Then
150                 Call TirarItemAlPiso(UserList(LadronIndex).Pos, MiObj)
                
                End If
        
152             If UserList(LadronIndex).clase = e_Class.Thief Then
154                 Call WriteConsoleMsg(LadronIndex, "Has robado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)
156                 Call WriteConsoleMsg(VictimaIndex, UserList(LadronIndex).Name & " te ha robado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)
                Else
158                 Call WriteConsoleMsg(LadronIndex, "Has hurtado " & MiObj.amount & " " & ObjData(MiObj.ObjIndex).Name, e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)
                
                End If

            Else
160             Call WriteConsoleMsg(LadronIndex, "No has logrado robar ningun objeto.", e_FontTypeNames.FONTTYPE_INFO)

            End If

            'If exiting, cancel de quien es robado
162         Call CancelExit(VictimaIndex)

        End With

        
        Exit Sub

RobarObjeto_Err:
164     Call TraceError(Err.Number, Err.Description, "Trabajo.RobarObjeto", Erl)

        
End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarSta_Err
        
100     UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad

102     If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
104     If UserList(UserIndex).Stats.MinSta = 0 Then Exit Sub
106     Call WriteUpdateSta(UserIndex)

        Exit Sub

QuitarSta_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.QuitarSta", Erl)
110
        
End Sub

Public Sub DoRaices(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer
    
100     With UserList(UserIndex)
    
102         If .flags.Privilegios And (e_PlayerType.Consejero) Then
                Exit Sub
            End If
            
104         If .Stats.MinSta > 2 Then
106             Call QuitarSta(UserIndex, 2)
        
            Else
            
108             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Estás muy cansado para obtener raices.", e_FontTypeNames.FONTTYPE_INFO)
110             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
    
            End If
    
            Dim Skill As Integer
112             Skill = .Stats.UserSkills(e_Skill.Alquimia)
        
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
116         res = RandomNumber(1, Suerte)
    
'118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))
    
            Rem Ladder 06/08/14 Subo un poco la probabilidad de sacar raices... porque era muy lento
120         If res < 7 Then
    
                Dim nPos  As t_WorldPos
                Dim MiObj As t_Obj
                
                Call ActualizarRecurso(.pos.Map, x, y)
                
                'If .clase = e_Class.Druid Then
                'MiObj.Amount = RandomNumber(6, 8)
                ' Else
122             MiObj.amount = RandomNumber(5, 7)
                ' End If

128             MiObj.amount = Round(MiObj.amount * 2.5 * RecoleccionMult)
130             MiObj.ObjIndex = Raices
        
132             MapData(.Pos.Map, X, Y).ObjInfo.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount - MiObj.amount
    
134             If MapData(.Pos.Map, X, Y).ObjInfo.amount < 0 Then
136                 MapData(.Pos.Map, X, Y).ObjInfo.amount = 0
                End If
        
140             If Not MeterItemEnInventario(UserIndex, MiObj) Then
            
142                 Call TirarItemAlPiso(.Pos, MiObj)
            
                End If
        
                'Call WriteConsoleMsg(UserIndex, "¡Has conseguido algunas raices!", e_FontTypeNames.FONTTYPE_INFO)
144             Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.CharIndex, vbWhite)
146             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .pos.x, .pos.y))
            Else
148             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .pos.x, .pos.y))
    
            End If
    
150         Call SubirSkill(UserIndex, e_Skill.Alquimia)
    
152         .Counters.Trabajando = .Counters.Trabajando + 1
            .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)
    
154         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
156             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
158     Call LogError("Error en DoRaices")

End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)
        On Error GoTo ErrHandler

        Dim Suerte As Integer
        Dim res    As Integer

100     With UserList(UserIndex)

102          If .flags.Privilegios And (e_PlayerType.Consejero) Then
                Exit Sub
             End If


                'EsfuerzoTalarLeñador = 1
104         If .Stats.MinSta > 5 Then
106             Call QuitarSta(UserIndex, 5)

            Else
108             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para talar.", e_FontTypeNames.FONTTYPE_INFO)
110             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub

            End If

            Dim Skill As Integer

112         Skill = .Stats.UserSkills(e_Skill.Talar)
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

            'HarThaoS: Le agrego más dificultad al talar en zona segura.  37% probabilidad de fallo en segura vs 16% en insegura
116         res = RandomNumber(1, IIf(MapInfo(UserList(userindex).Pos.map).Seguro = 1, Suerte + 4, Suerte))

            'ReyarB: aumento chances solamente si es el arbol de pino nudoso.
            If ObjData(MapData(.pos.map, x, y).ObjInfo.objIndex).Pino = 1 Then
                res = 1
                Suerte = 100
            End If

'118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))

120         If res < 6 Then

                Dim nPos  As t_WorldPos

                Dim MiObj As t_Obj

122             Call ActualizarRecurso(.Pos.Map, X, Y)
124             MapData(.Pos.Map, X, Y).ObjInfo.Data = GetTickCount() ' Ultimo uso
    
125             If .clase = Trabajador Then
126                 MiObj.amount = GetExtractResourceForLevel(.Stats.ELV)
                Else
127                 MiObj.amount = RandomNumber(1, 2)
                End If
128             MiObj.amount = MiObj.amount * RecoleccionMult

129             If ObjData(MapData(.pos.map, x, y).ObjInfo.objIndex).Elfico = 1 Then
130                 MiObj.objIndex = ElvenWood
                        
                ElseIf ObjData(MapData(.pos.map, x, y).ObjInfo.objIndex).Pino = 1 Then
                    MiObj.objIndex = PinoWood
                
                Else
132                 MiObj.objIndex = Wood
                End If
                
134             If MiObj.amount > MapData(.Pos.Map, X, Y).ObjInfo.amount Then
136                 MiObj.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount
                End If
            
138             MapData(.Pos.Map, X, Y).ObjInfo.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount - MiObj.amount
                 ' AGREGAR FX
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
140             If Not MeterItemEnInventario(UserIndex, MiObj) Then
142                 Call TirarItemAlPiso(.Pos, MiObj)
                End If
    
144             Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.CharIndex, vbWhite)

                If MapInfo(.Pos.Map).Seguro = 1 Or Zona(.ZonaId).Segura = 1 Then
146                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
                Else
145                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.y))
                End If

                ' Al talar también podés dropear cosas raras (se setean desde RecursosEspeciales.dat)
                Dim i As Integer

                ' Por cada drop posible
148             For i = 1 To UBound(EspecialesTala)
                    ' Tiramos al azar entre 1 y la probabilidad
150                 res = RandomNumber(1, EspecialesTala(i).Data)

                    ' Si tiene suerte y le pega
152                 If res = 1 Then
154                     MiObj.ObjIndex = EspecialesTala(i).ObjIndex
156                     MiObj.amount = 1 ' Solo un item por vez

                        ' Tiro siempre el item al piso, me parece más rolero, como que cae del árbol :P
158                     Call TirarItemAlPiso(.Pos, MiObj)
                    End If

160             Next i

            Else
162             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(64, .Pos.X, .Pos.y))

            End If
        
164         Call SubirSkill(UserIndex, e_Skill.Talar)
        
166         .Counters.Trabajando = .Counters.Trabajando + 1
            .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)
    
168         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
170             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With

        Exit Sub

ErrHandler:
172     Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte, Optional ByVal ObjetoDorado As Boolean = False)

        On Error GoTo ErrHandler

        Dim Suerte      As Integer
        Dim res         As Integer
        Dim Metal       As Integer
        Dim Yacimiento  As t_ObjData

100     With UserList(UserIndex)
    
102         If .flags.Privilegios And (e_PlayerType.Consejero) Then
                Exit Sub
            End If
    
            'Por Ladder 06/07/2014 Cuando la estamina llega a 0 , el macro se desactiva
104         If .Stats.MinSta > 5 Then
106             Call QuitarSta(UserIndex, 5)
            Else
108             Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                'Call WriteConsoleMsg(UserIndex, "Estás muy cansado para excavar.", e_FontTypeNames.FONTTYPE_INFO)
110             Call WriteMacroTrabajoToggle(UserIndex, False)
                Exit Sub
    
            End If
    
            'Por Ladder 06/07/2014
    
            Dim Skill As Integer
    
112         Skill = .Stats.UserSkills(e_Skill.Mineria)
114         Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

            'HarThaoS: Le agrego más dificultad al talar en zona segura.  37% probabilidad de fallo en segura vs 16% en insegura
116         res = RandomNumber(1, IIf(MapInfo(UserList(userindex).Pos.map).Seguro = 1, Suerte + 2, Suerte))
        
'118         Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageArmaMov(.Char.CharIndex))


            'ReyarB: aumento chances solamente si es mineria de blodium.
            If ObjData(MapData(.pos.map, x, y).ObjInfo.objIndex).MineralIndex = 3787 Then
                res = 1
                Suerte = 100
            End If
        
120         If res <= 5 Then
    
                Dim MiObj As t_Obj
                Dim nPos  As t_WorldPos
            
122             Call ActualizarRecurso(.Pos.Map, X, Y)
124             MapData(.Pos.Map, X, Y).ObjInfo.Data = GetTickCount() ' Ultimo uso

126             Yacimiento = ObjData(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex)
            
128             MiObj.ObjIndex = Yacimiento.MineralIndex
129             If .clase = Trabajador Then
130                 MiObj.amount = GetExtractResourceForLevel(.Stats.ELV)
                Else
131                 MiObj.amount = RandomNumber(1, 2)
                End If
132             MiObj.amount = MiObj.amount * RecoleccionMult
            
133             If MiObj.amount > MapData(.pos.map, X, y).ObjInfo.amount Then
134                 MiObj.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount
                End If
            
136             MapData(.Pos.Map, X, Y).ObjInfo.amount = MapData(.Pos.Map, X, Y).ObjInfo.amount - MiObj.amount
        
138             If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
                 ' AGREGAR FX
                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 253, 25, False, ObjData(MiObj.ObjIndex).GrhIndex))
                
139             Call WriteTextCharDrop(UserIndex, "+" & MiObj.amount, .Char.CharIndex, vbWhite)

140             Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", e_FontTypeNames.FONTTYPE_INFO)
                If MapInfo(.Pos.Map).Seguro = 1 Or Zona(.ZonaId).Segura = 1 Then
141                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessagePlayWave(15, .Pos.X, .Pos.Y))
                Else
142                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(15, .Pos.X, .Pos.y))
                End If
            
                ' Al minar también puede dropear una gema
                Dim i As Integer
    
                ' Por cada drop posible
144             For i = 1 To Yacimiento.CantItem
                    ' Tiramos al azar entre 1 y la probabilidad
146                 res = RandomNumber(1, Yacimiento.Item(i).amount)
                
                    ' Si tiene suerte y le pega
148                 If res = 1 Then
                        ' Se lo metemos al inventario (o lo tiramos al piso)
150                     MiObj.ObjIndex = Yacimiento.Item(i).ObjIndex
152                     MiObj.amount = 1 ' Solo una gema por vez
                    
154                     If Not MeterItemEnInventario(UserIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)

                        ' Le mandamos un mensaje
156                     Call WriteConsoleMsg(UserIndex, "¡Has conseguido " & ObjData(Yacimiento.Item(i).ObjIndex).name & "!", e_FontTypeNames.FONTTYPE_INFO)
                    End If
    
                Next
            
            Else
158             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(2185, .Pos.X, .Pos.y))

            End If

160         Call SubirSkill(UserIndex, e_Skill.Mineria)

162         .Counters.Trabajando = .Counters.Trabajando + 1
            .Counters.LastTrabajo = Int(IntervaloTrabajarExtraer / 1000)

164         If .Counters.Trabajando = 1 And Not .flags.UsandoMacro Then
166             Call WriteMacroTrabajoToggle(UserIndex, True)
            End If
    
        End With
    
        Exit Sub

ErrHandler:
168     Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)
        
        On Error GoTo DoMeditar_Err

        Dim Mana As Long
        
100     With UserList(UserIndex)

102         .Counters.TimerMeditar = .Counters.TimerMeditar + 1
            .Counters.TiempoInicioMeditar = .Counters.TiempoInicioMeditar + 1
104         If .Counters.TimerMeditar >= IntervaloMeditar And .Counters.TiempoInicioMeditar > 20 Then

106             Mana = Porcentaje(.Stats.MaxMAN, Porcentaje(PorcentajeRecuperoMana, 50 + .Stats.UserSkills(e_Skill.Meditar) * 0.5))

108             If Mana <= 0 Then Mana = 1

110             If .Stats.MinMAN + Mana >= .Stats.MaxMAN Then

112                 .Stats.MinMAN = .Stats.MaxMAN
114                 .flags.Meditando = False
116                 .Char.FX = 0

118                 Call WriteUpdateMana(UserIndex)
120                 Call SubirSkill(UserIndex, Meditar)

122                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
                
                Else
                    
124                 .Stats.MinMAN = .Stats.MinMAN + Mana
126                 Call WriteUpdateMana(UserIndex)
128                 Call SubirSkill(UserIndex, Meditar)

                End If

130             .Counters.TimerMeditar = 0
            End If

        End With

        
        Exit Sub

DoMeditar_Err:
132     Call TraceError(Err.Number, Err.Description, "Trabajo.DoMeditar", Erl)
134
        
End Sub

Public Sub DoMontar(ByVal UserIndex As Integer, ByRef Montura As t_ObjData, ByVal Slot As Integer)
        On Error GoTo DoMontar_Err

100     With UserList(UserIndex)
102         If PuedeUsarObjeto(UserIndex, .Invent.Object(Slot).ObjIndex, True) > 0 Then
                Exit Sub
            End If

104         If .flags.Montado = 0 And .Counters.EnCombate > 0 Then
106             Call WriteConsoleMsg(UserIndex, "Estás en combate, debes aguardar " & .Counters.EnCombate & " segundo(s) para montar...", e_FontTypeNames.FONTTYPE_INFOBOLD)
                Exit Sub
            End If

108         If .flags.EnReto Then
110             Call WriteConsoleMsg(UserIndex, "No podés montar en un reto.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

112         If (.flags.Oculto = 1 Or .flags.invisible = 1) And .flags.AdminInvisible = 0 Then
114             Call WriteConsoleMsg(UserIndex, "No podés montar estando oculto o invisible.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            If .flags.Mimetizado <> e_EstadoMimetismo.Desactivado Then
                Call WriteConsoleMsg(UserIndex, "Pierdes el efecto del mimetismo.", e_FontTypeNames.FONTTYPE_INFO)
                .Counters.Mimetismo = 0
                .flags.Mimetizado = e_EstadoMimetismo.Desactivado
                Call RefreshCharStatus(UserIndex)
            End If

116         If .flags.Montado = 0 And (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger > 10) Then
118             Call WriteConsoleMsg(UserIndex, "No podés montar aquí.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

120         If .flags.Meditando Then
122             .flags.Meditando = False
124             .Char.FX = 0
126             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
            End If

128         If .flags.Montado = 1 And .Invent.MonturaObjIndex > 0 Then
130             If ObjData(.Invent.MonturaObjIndex).ResistenciaMagica > 0 Then
132                 Call UpdateUserInv(False, UserIndex, .Invent.MonturaSlot)
                End If

            End If

134         .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
136         .Invent.MonturaSlot = Slot

138         If .flags.Montado = 0 Then
140             .Char.Body = Montura.Ropaje
142             .Char.Head = .OrigChar.Head
144             .Char.ShieldAnim = NingunEscudo
146             .Char.WeaponAnim = NingunArma
148             .Char.CascoAnim = .Char.CascoAnim
149             .Char.CartAnim = NoCart
150             .flags.Montado = 1
                Call TargetUpdateTerrain(.EffectOverTime)
            Else
152             .flags.Montado = 0
154             .Char.Head = .OrigChar.Head
                Call TargetUpdateTerrain(.EffectOverTime)
156             If .Invent.ArmourEqpObjIndex > 0 Then
158                 .Char.Body = ObtenerRopaje(UserIndex, ObjData(.Invent.ArmourEqpObjIndex))

                Else
                    Call SetNakedBody(UserList(userIndex))

                End If

162             If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
164             If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
166             If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
167             If .invent.MagicoObjIndex > 0 Then
                    If ObjData(.invent.MagicoObjIndex).Ropaje > 0 Then .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje
                End If

            End If

168         Call ActualizarVelocidadDeUsuario(UserIndex)
170         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)

172         Call UpdateUserInv(False, UserIndex, Slot)
174         Call WriteEquiteToggle(UserIndex)
        End With

        Exit Sub

DoMontar_Err:
176     Call TraceError(Err.Number, Err.Description, "Trabajo.DoMontar", Erl)
178

End Sub

Public Sub ActualizarRecurso(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo ActualizarRecurso_Err
        

        Dim ObjIndex As Integer

100     ObjIndex = MapData(Map, X, Y).ObjInfo.ObjIndex

        Dim TiempoActual As Long

102     TiempoActual = GetTickCount()

        ' Data = Ultimo uso
104     If (TiempoActual - MapData(Map, X, Y).ObjInfo.Data) * 0.001 > ObjData(ObjIndex).TiempoRegenerar Then
106         MapData(Map, X, Y).ObjInfo.amount = ObjData(ObjIndex).VidaUtil
108         MapData(Map, X, Y).ObjInfo.Data = &H7FFFFFFF   ' Ultimo uso = Max Long

        End If

        
        Exit Sub

ActualizarRecurso_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.ActualizarRecurso", Erl)
112
        
End Sub

Public Function ObtenerPezRandom(ByVal PoderCania As Integer) As Long
        
        On Error GoTo ObtenerPezRandom_Err
        

        Dim i As Long, SumaPesos As Long, ValorGenerado As Long
    
100     If PoderCania > UBound(PesoPeces) Then PoderCania = UBound(PesoPeces)
102     SumaPesos = PesoPeces(PoderCania)

104     ValorGenerado = RandomNumber(0, SumaPesos - 1)

106     ObtenerPezRandom = Peces(BinarySearchPeces(ValorGenerado)).ObjIndex

        Exit Function

ObtenerPezRandom_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.ObtenerPezRandom", Erl)
110
        
End Function

Function ModDomar(ByVal clase As e_Class) As Integer
        
        On Error GoTo ModDomar_Err
    
        

        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************
100     Select Case clase

            Case e_Class.Druid
102             ModDomar = 6

104         Case e_Class.Hunter
106             ModDomar = 6

108         Case e_Class.Cleric
110             ModDomar = 7

112         Case Else
114             ModDomar = 10

        End Select

        
        Exit Function

ModDomar_Err:
116     Call TraceError(Err.Number, Err.Description, "Trabajo.ModDomar", Erl)

        
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
        
        On Error GoTo FreeMascotaIndex_Err
    
        

        '***************************************************
        'Author: Unknown
        'Last Modification: 02/03/09
        '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
        '***************************************************
        Dim j As Integer

100     For j = 1 To MAXMASCOTAS

102         If UserList(UserIndex).MascotasType(j) = 0 Then
104             FreeMascotaIndex = j
                Exit Function

            End If

106     Next j
        
        FreeMascotaIndex = -1
        
        Exit Function

FreeMascotaIndex_Err:
108     Call TraceError(Err.Number, Err.Description, "Trabajo.FreeMascotaIndex", Erl)

        
End Function

Private Function HayEspacioMascotas(ByVal UserIndex As Integer) As Boolean
            HayEspacioMascotas = (FreeMascotaIndex(UserIndex) > 0)
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
        '***************************************************
        'Author: Nacho (Integer)
        'Last Modification: 01/05/2010
        '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
        '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
        '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
        '***************************************************

        On Error GoTo ErrHandler

        Dim puntosDomar      As Integer

        Dim CanStay          As Boolean

        Dim petType          As Integer

        Dim NroPets          As Integer
    
100     If IsValidUserRef(NpcList(npcIndex).MaestroUser) And NpcList(npcIndex).MaestroUser.ArrayIndex = userIndex Then
102         Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

104     With UserList(UserIndex)


106         If .flags.Privilegios And e_PlayerType.Consejero Then Exit Sub
            
108         If .NroMascotas < MAXMASCOTAS And HayEspacioMascotas(UserIndex) Then

110             If IsValidNpcRef(NpcList(NpcIndex).MaestroNPC) > 0 Or IsValidUserRef(NpcList(NpcIndex).MaestroUser) Then
112                 Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If

114             puntosDomar = CInt(.Stats.UserAtributos(e_Atributos.Carisma)) * CInt(.Stats.UserSkills(e_Skill.Domar))

116             If .clase = e_Class.Druid Then
118                 puntosDomar = puntosDomar / 6 'original es 6
                Else
120                 puntosDomar = puntosDomar / 11
                End If

122             If NpcList(NpcIndex).flags.Domable <= puntosDomar And RandomNumber(1, 5) = 1 Then
                    Dim Index As Integer
124                 .NroMascotas = .NroMascotas + 1
126                 Index = FreeMascotaIndex(UserIndex)
128                 Call SetNpcRef(.MascotasIndex(Index), NpcIndex)
130                 .MascotasType(Index) = NpcList(NpcIndex).Numero
132                 Call SetUserRef(NpcList(npcIndex).MaestroUser, userIndex)
                    .flags.ModificoMascotas = True
134                 Call FollowAmo(NpcIndex)
136                 Call ReSpawnNpc(NpcList(NpcIndex))

138                 Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", e_FontTypeNames.FONTTYPE_INFO)

                    ' Es zona segura?
140                 If MapInfo(.pos.Map).NoMascotas = 1 Then
142                     petType = NpcList(NpcIndex).Numero
144                     NroPets = .NroMascotas
146                     Call QuitarNPC(NpcIndex, eNewPet)
148                     .MascotasType(Index) = petType
150                     .NroMascotas = NroPets
152                     Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
154                 If Not .flags.UltimoMensaje = 5 Then
156                     Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", e_FontTypeNames.FONTTYPE_INFO)
158                     .flags.UltimoMensaje = 5
                    End If
                End If
160             Call SubirSkill(UserIndex, e_Skill.Domar)
            Else
162             Call WriteConsoleMsg(UserIndex, "No puedes controlar mas criaturas.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        End With
        Exit Sub

ErrHandler:
164     Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, _
                                   ByVal NpcIndex As Integer) As Boolean
        
        On Error GoTo PuedeDomarMascota_Err
    
        

        '***************************************************
        'Author: ZaMa
        'This function checks how many NPCs of the same type have
        'been tamed by the user.
        'Returns True if that amount is less than two.
        '***************************************************
        Dim i           As Long

        Dim numMascotas As Long
    
100     For i = 1 To MAXMASCOTAS

102         If UserList(UserIndex).MascotasType(i) = NpcList(NpcIndex).Numero Then
104             numMascotas = numMascotas + 1

            End If

106     Next i
    
108     If numMascotas <= 1 Then PuedeDomarMascota = True
    
        
        Exit Function

PuedeDomarMascota_Err:
110     Call TraceError(Err.Number, Err.Description, "Trabajo.PuedeDomarMascota", Erl)

        
End Function

Public Function EntregarPezEspecial(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .flags.PescandoEspecial Then
            Dim obj As t_Obj
            obj.amount = 1
            obj.ObjIndex = .Stats.NumObj_PezEspecial
            If Not MeterItemEnInventario(UserIndex, obj) Then
                Call TirarItemAlPiso(.Pos, obj)
            End If
            
            Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageParticleFX(.Char.CharIndex, 253, 25, False, ObjData(obj.ObjIndex).GrhIndex))
            Call WriteConsoleMsg(UserIndex, "Felicitaciones has pescado un pez de gran porte ( " & ObjData(obj.ObjIndex).Name & " )", e_FontTypeNames.FONTTYPE_INFOBOLD)
            .Stats.NumObj_PezEspecial = 0
            .flags.PescandoEspecial = False
        End If
    End With
End Function

Public Sub FishOrThrowNet(ByVal UserIndex As Integer)
On Error GoTo FishOrThrowNet_Err:
100 With UserList(UserIndex)
102     If ObjData(.invent.HerramientaEqpObjIndex).OBJType <> e_OBJType.otHerramientas Then Exit Sub
104     If ObjData(.invent.HerramientaEqpObjIndex).Subtipo = e_ToolsSubtype.eFishingNet Then
106         If MapInfo(.Pos.Map).Seguro = 1 Or Zona(.ZonaId).Segura = 1 Or _
                Not ExpectObjectTypeAt(e_OBJType.otFishingPool, .pos.map, .Trabajo.Target_X, .Trabajo.Target_Y) Then
108             If IsValidUserRef(.flags.TargetUser) Or IsValidNpcRef(.flags.TargetNPC) Then
110                 ThrowNetToTarget (UserIndex)
112                 Call WriteWorkRequestTarget(UserIndex, 0)
                    Exit Sub
                End If
            End If
        End If
114     Call Trabajar(UserIndex, e_Skill.Pescar)
    End With
    Exit Sub
FishOrThrowNet_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.FishOrThrowNet", Erl)
End Sub

Sub ThrowNetToTarget(ByVal UserIndex As Integer)
On Error GoTo ThrowNetToTarget_Err:
100 With UserList(UserIndex)
102     If .invent.HerramientaEqpObjIndex = 0 Then Exit Sub
104     If ObjData(.invent.HerramientaEqpObjIndex).OBJType <> e_OBJType.otHerramientas Then Exit Sub
106     If ObjData(.invent.HerramientaEqpObjIndex).Subtipo <> e_ToolsSubtype.eFishingNet Then Exit Sub

        'If it's outside range log it and exit
108     If Abs(.pos.X - .Trabajo.Target_X) > RANGO_VISION_X Or Abs(.pos.y - .Trabajo.Target_Y) > RANGO_VISION_Y Then
110         Call LogSecurity("Ataque fuera de rango de " & .name & "(" & .pos.Map & "/" & .pos.x & "/" & .pos.y & ") ip: " & .ConnectionDetails.IP & " a la posicion (" & .pos.Map & "/" & .Trabajo.Target_X & "/" & .Trabajo.Target_Y & ")")
            Exit Sub
        End If
        
        'Check bow's interval
112     If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        'Check attack-spell interval
114     If Not IntervaloPermiteGolpeMagia(UserIndex, False) Then Exit Sub
        'Check Magic interval
116     If Not IntervaloPermiteLanzarSpell(UserIndex) Then Exit Sub
        'check item cd
        
        
118     Dim ThrowNet As Boolean
120     ThrowNet = False
122     If IsValidUserRef(UserList(UserIndex).flags.TargetUser) Then
124         Dim tU As Integer
126         tU = UserList(UserIndex).flags.TargetUser.ArrayIndex
128         If UserIndex = tU Then
130             Call WriteLocaleMsg(UserIndex, MsgCantAttackYourself, e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            If IsSet(UserList(tU).flags.StatusMask, eCCInmunity) Then
                Call WriteLocaleMsg(UserIndex, MsgCCInunity, e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            If Not UserMod.CanMove(UserList(tU).flags, UserList(tU).Counters) Then
136             Call WriteConsoleMsg(UserIndex, "No podes inmovilizar un objetivo que no puede moverse.", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
140         If Not PuedeAtacar(UserIndex, tU) Then Exit Sub
            
142         Call UsuarioAtacadoPorUsuario(UserIndex, tU)
            
144         UserList(tU).Counters.Inmovilizado = NET_INMO_DURATION
146          If UserList(tU).flags.Inmovilizado = 0 Then
148             UserList(tU).flags.Inmovilizado = 1
150             Call SendData(SendTarget.ToPCAliveArea, tU, PrepareMessageCreateFX(UserList(tU).Char.charindex, FISHING_NET_FX, 0, UserList(tU).pos.X, UserList(tU).pos.y))
152             Call WriteInmovilizaOK(tU)
154             Call WritePosUpdate(tU)
156             ThrowNet = True
            End If
158         Call SetUserRef(UserList(UserIndex).flags.TargetUser, 0)
160     ElseIf IsValidNpcRef(UserList(UserIndex).flags.TargetNPC) Then

162         Dim npcIndex As Integer
164         NpcIndex = UserList(UserIndex).flags.TargetNPC.ArrayIndex
166         If NpcList(npcIndex).flags.AfectaParalisis = 0 Then

                Dim UserAttackInteractionResult As t_AttackInteractionResult
                UserAttackInteractionResult = UserCanAttackNpc(UserIndex, NpcIndex)
                Call SendAttackInteractionMessage(UserIndex, UserAttackInteractionResult.Result)
                If UserAttackInteractionResult.CanAttack Then
                    If UserAttackInteractionResult.TurnPK Then Call VolverCriminal(UserIndex)
                Else
                    Exit Sub
                End If
170             Call NPCAtacado(npcIndex, UserIndex)
172             NpcList(npcIndex).flags.Inmovilizado = 1
174             NpcList(npcIndex).Contadores.Inmovilizado = (NET_INMO_DURATION * 6.5) * 6
176             NpcList(npcIndex).flags.Paralizado = 0
178             NpcList(npcIndex).Contadores.Paralisis = 0
180             Call AnimacionIdle(npcIndex, True)
182             ThrowNet = True
184             Call SendData(SendTarget.ToNPCAliveArea, npcIndex, PrepareMessageFxPiso(FISHING_NET_FX, NpcList(npcIndex).pos.X, NpcList(npcIndex).pos.y))
186             Call ClearNpcRef(UserList(UserIndex).flags.TargetNPC)
            Else
188             Call WriteLocaleMsg(UserIndex, MSgNpcInmuneToEffect, e_FontTypeNames.FONTTYPE_INFOIAO)
            End If
        End If
        If ThrowNet Then
190         Call UpdateCd(UserIndex, ObjData(.invent.HerramientaEqpObjIndex).cdType)
192         Call QuitarUserInvItem(UserIndex, .invent.HerramientaEqpSlot, 1)
194         Call UpdateUserInv(True, UserIndex, .invent.HerramientaEqpSlot)
196         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(UserList(UserIndex).pos.X, UserList(UserIndex).pos.y, .Trabajo.Target_X, .Trabajo.Target_Y, 3))
        End If
    End With
    Exit Sub
ThrowNetToTarget_Err:
    Call TraceError(Err.Number, Err.Description, "Trabajo.ThrowNetToTarget", Erl)
End Sub

Public Function GetExtractResourceForLevel(ByVal level As Integer) As Integer
    Dim upper As Long
    Dim lower As Long
    lower = Int(CDbl(level + 0.000001) / 3.6)
    upper = Int(CDbl(level + 0.000001) / 2)
    GetExtractResourceForLevel = 5 + RandomNumber(lower, upper)
End Function
