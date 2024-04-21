Attribute VB_Name = "General"
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

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Global LeerNPCs As New clsIniManager


Sub SetNakedBody(ByRef user As t_User)
        Const man_human_naked_body As Integer = 3000
        Const man_drow_naked_body As Integer = 3001
        Const man_elf_naked_body As Integer = 3002
        Const man_gnome_naked_body As Integer = 3003
        Const man_dwarf_naked_body As Integer = 3004
        Const man_orc_naked_body As Integer = 3005
        
        Const woman_human_naked_body As Integer = 3006
        Const woman_drow_naked_body As Integer = 3007
        Const woman_elf_naked_body As Integer = 3008
        Const woman_gnome_naked_body As Integer = 3009
        Const woman_dwarf_naked_body As Integer = 3010
        Const woman_orc_naked_body As Integer = 3011
        
        user.flags.Desnudo = 1
        Select Case user.genero
            Case e_Genero.Hombre
               Select Case user.raza
                 Case e_Raza.Humano
                     user.Char.body = man_human_naked_body
                 Case e_Raza.Drow
                     user.Char.body = man_drow_naked_body
                 Case e_Raza.Elfo
                     user.Char.body = man_elf_naked_body
                 Case e_Raza.Gnomo
                     user.Char.body = man_gnome_naked_body
                 Case e_Raza.Enano
                     user.Char.body = man_dwarf_naked_body
                 Case e_Raza.Orco
                     user.Char.body = man_orc_naked_body
                 Case Else
                     user.Char.body = man_human_naked_body
                End Select

            Case e_Genero.Mujer
                Select Case user.raza
                    Case e_Raza.Humano
                        user.Char.body = woman_human_naked_body
                    Case e_Raza.Drow
                        user.Char.body = woman_drow_naked_body
                    Case e_Raza.Elfo
                        user.Char.body = woman_elf_naked_body
                    Case e_Raza.Gnomo
                        user.Char.body = woman_gnome_naked_body
                    Case e_Raza.Enano
                        user.Char.body = woman_dwarf_naked_body
                    Case e_Raza.Orco
                        user.Char.body = woman_orc_naked_body
                    Case Else
                        user.Char.body = woman_human_naked_body
                End Select
        End Select
End Sub

Sub Bloquear(ByVal toMap As Boolean, ByVal sndIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal b As Byte)
        'b ahora es boolean,
        'b=true bloquea el tile en (x,y)
        'b=false desbloquea el tile en (x,y)
        'toMap = true -> Envia los datos a todo el mapa
        'toMap = false -> Envia los datos al user
        'Unifique los tres parametros (sndIndex,sndMap y map) en sndIndex... pero de todas formas, el mapa jamas se indica.. eso esta bien asi?
        'Puede llegar a ser, que se quiera mandar el mapa, habria que agregar un nuevo parametro y modificar.. lo quite porque no se usaba ni aca ni en el cliente :s
        ' WyroX: Uso bloqueo parcial
        On Error GoTo Bloquear_Err
        ' Envío sólo los flags de bloq
100     b = b And e_Block.ALL_SIDES
102     If toMap Then
104         Call SendData(SendTarget.toMap, sndIndex, PrepareMessage_BlockPosition(X, Y, b))
        Else
106         Call Write_BlockPosition(sndIndex, X, Y, b)
        End If
        Exit Sub
Bloquear_Err:
108     Call TraceError(Err.Number, Err.Description, "General.Bloquear", Erl)
End Sub

Sub BlockAndInform(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal NewState As Integer)
100    If NewState Then
110        MapData(Map, X, Y).Blocked = e_Block.ALL_SIDES Or e_Block.GM
       Else
120        MapData(Map, X, Y).Blocked = 0
       End If
        
114    Call Bloquear(True, Map, X, Y, MapData(Map, X, Y).Blocked)
End Sub

Sub MostrarBloqueosPuerta(ByVal toMap As Boolean, _
                          ByVal sndIndex As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer)
        
        On Error GoTo MostrarBloqueosPuerta_Err
        
        Dim Map       As Integer
        Dim ModPuerta As Integer
        
100     If toMap Then
102         Map = sndIndex
        Else
104         Map = UserList(sndIndex).Pos.Map
        End If
        
106     ModPuerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo

108     Select Case ModPuerta
        
            Case 0
                ' Bloqueos superiores
110             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
112             Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)

                ' Bloqueos inferiores
114             Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)
116             Call Bloquear(toMap, sndIndex, X - 1, Y + 1, MapData(Map, X - 1, Y + 1).Blocked)

118         Case 1
                ' para palancas o teclas sin modicar bloqueos en X,Y
                
120         Case 2
                ' Bloqueos superiores
122             Call Bloquear(toMap, sndIndex, X, Y - 1, MapData(Map, X, Y - 1).Blocked)
124             Call Bloquear(toMap, sndIndex, X - 1, Y - 1, MapData(Map, X - 1, Y - 1).Blocked)
126             Call Bloquear(toMap, sndIndex, X + 1, Y - 1, MapData(Map, X + 1, Y - 1).Blocked)
                ' Bloqueos inferiores
128             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
130             Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
132             Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                
134         Case 3
                ' Bloqueos superiores
136             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
138             Call Bloquear(toMap, sndIndex, X - 1, Y, MapData(Map, X - 1, Y).Blocked)
140             Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
                ' Bloqueos inferiores
142             Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)
144             Call Bloquear(toMap, sndIndex, X - 1, Y + 1, MapData(Map, X - 1, Y + 1).Blocked)
146             Call Bloquear(toMap, sndIndex, X + 1, Y + 1, MapData(Map, X + 1, Y + 1).Blocked)

148         Case 4
                ' Bloqueos superiores
150             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
                ' Bloqueos inferiores
152             Call Bloquear(toMap, sndIndex, X, Y + 1, MapData(Map, X, Y + 1).Blocked)

154         Case 5 'Ver WyroX
                ' Bloqueos vertical ver ReyarB
156             Call Bloquear(toMap, sndIndex, X + 1, Y, MapData(Map, X + 1, Y).Blocked)
158             Call Bloquear(toMap, sndIndex, X + 1, Y - 1, MapData(Map, X + 1, Y - 1).Blocked)

                ' Bloqueos horizontal
160             Call Bloquear(toMap, sndIndex, X, Y - 2, MapData(Map, X, Y - 2).Blocked)
162             Call Bloquear(toMap, sndIndex, X - 1, Y - 2, MapData(Map, X - 1, Y - 2).Blocked)


164         Case 6 ' Ver WyroX
                ' Bloqueos superiores ver ReyarB
166             Call Bloquear(toMap, sndIndex, X, Y, MapData(Map, X, Y).Blocked)
168             Call Bloquear(toMap, sndIndex, X, Y - 1, MapData(Map, X, Y - 1).Blocked)

                ' Bloqueos inferiores
170             Call Bloquear(toMap, sndIndex, X, Y - 2, MapData(Map, X, Y - 2).Blocked)
172             Call Bloquear(toMap, sndIndex, X + 1, Y - 2, MapData(Map, X + 1, Y - 2).Blocked)

        End Select

        Exit Sub

MostrarBloqueosPuerta_Err:
174     Call TraceError(Err.Number, Err.Description, "General.MostrarBloqueosPuerta", Erl)
        
End Sub

Sub BloquearPuerta(ByVal Map As Integer, _
                   ByVal X As Integer, _
                   ByVal Y As Integer, _
                   ByVal Bloquear As Boolean)
        
        On Error GoTo BloquearPuerta_Err
        Dim ModPuerta As Integer
        
        'ver reyarb
100     ModPuerta = ObjData(MapData(Map, X, Y).ObjInfo.ObjIndex).Subtipo

102     Select Case ModPuerta
        
            Case 0 'puerta 2 tiles

                ' Bloqueos superiores
104             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.NORTH, MapData(Map, X, Y).Blocked And Not e_Block.NORTH)
106             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.NORTH, MapData(Map, X - 1, Y).Blocked And Not e_Block.NORTH)

                ' Cambio bloqueos inferiores
108             MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X, Y + 1).Blocked And Not e_Block.SOUTH)
110             MapData(Map, X - 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X - 1, Y + 1).Blocked And Not e_Block.SOUTH)

112         Case 1
                ' para palancas o teclas sin modicar bloqueos en X,Y

114         Case 2 ' puerta 3 tiles 1 arriba
                ' Bloqueos superiores
116             MapData(Map, X, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X, Y - 1).Blocked Or e_Block.NORTH, MapData(Map, X, Y - 1).Blocked And Not e_Block.NORTH)
118             MapData(Map, X - 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 1).Blocked Or e_Block.NORTH, MapData(Map, X - 1, Y - 1).Blocked And Not e_Block.NORTH)
120             MapData(Map, X + 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 1).Blocked Or e_Block.NORTH, MapData(Map, X + 1, Y - 1).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
122             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.SOUTH, MapData(Map, X, Y).Blocked And Not e_Block.SOUTH)
124             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.SOUTH, MapData(Map, X - 1, Y).Blocked And Not e_Block.SOUTH)
126             MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or e_Block.SOUTH, MapData(Map, X + 1, Y).Blocked And Not e_Block.SOUTH)
                
128         Case 3 ' puerta 3 tiles
                ' Bloqueos superiores
130             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.NORTH, MapData(Map, X, Y).Blocked And Not e_Block.NORTH)
132             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.NORTH, MapData(Map, X - 1, Y).Blocked And Not e_Block.NORTH)
134             MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or e_Block.NORTH, MapData(Map, X + 1, Y).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
136             MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X, Y + 1).Blocked And Not e_Block.SOUTH)
138             MapData(Map, X - 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X - 1, Y + 1).Blocked And Not e_Block.SOUTH)
140             MapData(Map, X + 1, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X + 1, Y + 1).Blocked And Not e_Block.SOUTH)
        
142         Case 4 'puerta 1 tiles
                ' Bloqueos superiores
144             MapData(Map, X, Y).Blocked = IIf(Bloquear, MapData(Map, X, Y).Blocked Or e_Block.NORTH, MapData(Map, X, Y).Blocked And Not e_Block.NORTH)
                ' Cambio bloqueos inferiores
146             MapData(Map, X, Y + 1).Blocked = IIf(Bloquear, MapData(Map, X, Y + 1).Blocked Or e_Block.SOUTH, MapData(Map, X, Y + 1).Blocked And Not e_Block.SOUTH)
                
148         Case 5 'Ver WyroX
                ' Bloqueos  vertical ver ReyarB
150             MapData(Map, X + 1, Y).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y).Blocked Or e_Block.ALL_SIDES, MapData(Map, X + 1, Y).Blocked And Not e_Block.ALL_SIDES)
152             MapData(Map, X + 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, X + 1, Y - 1).Blocked And Not e_Block.ALL_SIDES)
                
                ' Cambio horizontal
154             MapData(Map, X, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X, Y - 2).Blocked And Not e_Block.ALL_SIDES)
156             MapData(Map, X - 1, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X - 1, Y - 2).Blocked And Not e_Block.ALL_SIDES)


158         Case 6 ' Ver Wyrox
                ' Bloqueos vertical ver ReyarB
160             MapData(Map, X - 1, Y).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y).Blocked Or e_Block.ALL_SIDES, MapData(Map, X - 1, Y).Blocked And Not e_Block.ALL_SIDES)
162             MapData(Map, X - 1, Y - 1).Blocked = IIf(Bloquear, MapData(Map, X - 1, Y - 1).Blocked Or e_Block.ALL_SIDES, MapData(Map, X - 1, Y - 1).Blocked And Not e_Block.ALL_SIDES)
                
                ' Cambio bloqueos Puerta abierta
164             MapData(Map, X, Y - 2).Blocked = IIf(Bloquear, MapData(Map, X, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X, Y - 2).Blocked And Not e_Block.ALL_SIDES)
166             MapData(Map, X + 1, Y + 2).Blocked = IIf(Bloquear, MapData(Map, X + 1, Y - 2).Blocked Or e_Block.ALL_SIDES, MapData(Map, X + 1, Y - 2).Blocked And Not e_Block.ALL_SIDES)

                
        End Select

        ' Mostramos a todos
168     Call MostrarBloqueosPuerta(True, Map, X, Y)
        
        Exit Sub

BloquearPuerta_Err:
170     Call TraceError(Err.Number, Err.Description, "General.BloquearPuerta", Erl)
        
End Sub

Function HayCosta(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayCosta_Err
        

        'Ladder 10 - 2 - 2010
        'Chequea si hay costa en los tiles proximos al usuario
100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If ((MapData(Map, X, Y).Graphic(1) >= 22552 And MapData(Map, X, Y).Graphic(1) <= 22599) Or (MapData(Map, X, Y).Graphic(1) >= 7283 And MapData(Map, X, Y).Graphic(1) <= 7378) Or (MapData(Map, X, Y).Graphic(1) >= 13387 And MapData(Map, X, Y).Graphic(1) <= 13482)) And MapData(Map, X, Y).Graphic(2) = 0 Then
104             HayCosta = True
            Else
106             HayCosta = False

            End If

        Else
108         HayCosta = False

        End If

        
        Exit Function

HayCosta_Err:
110     Call TraceError(Err.Number, Err.Description, "General.HayCosta", Erl)

        
End Function

Function HayAgua(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayAgua_Err
        

100     With MapData(Map, X, Y)
102         If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
104             HayAgua = (.Graphic(1) >= 1505 And .Graphic(1) <= 1520) Or _
                    (.Graphic(1) >= 124 And .Graphic(1) <= 139) Or _
                    (.Graphic(1) >= 24223 And .Graphic(1) <= 24238) Or _
                    (.Graphic(1) >= 24303 And .Graphic(1) <= 24318) Or _
                    (.Graphic(1) >= 468 And .Graphic(1) <= 483) Or _
                    (.Graphic(1) >= 44668 And .Graphic(1) <= 44683) Or _
                    (.Graphic(1) >= 24143 And .Graphic(1) <= 24158) Or _
                    (.Graphic(1) >= 12628 And .Graphic(1) <= 12643) Or _
                    (.Graphic(1) >= 2948 And .Graphic(1) <= 2963)
            Else
106             HayAgua = False
    
            End If
        End With

        
        Exit Function

HayAgua_Err:
108     Call TraceError(Err.Number, Err.Description, "General.HayAgua", Erl)

        
End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean
        
        On Error GoTo EsArbol_Err
        
100     EsArbol = GrhIndex = 11905 Or GrhIndex = 644 Or GrhIndex = 1880 Or GrhIndex = 11906 Or GrhIndex = 12160 Or GrhIndex = 6597 Or GrhIndex = 2548 Or GrhIndex = 2549 Or _
           GrhIndex = 15110 Or GrhIndex = 15109 Or GrhIndex = 15108 Or GrhIndex = 11904 Or _
           GrhIndex = 7220 Or GrhIndex = 50990 Or GrhIndex = 55626 Or GrhIndex = 55627 Or GrhIndex = 55630 Or GrhIndex = 55632 Or GrhIndex = 55633 Or _
           GrhIndex = 55635 Or GrhIndex = 55638 Or GrhIndex = 12584 Or GrhIndex = 50985 Or GrhIndex = 15510 Or GrhIndex = 14775 Or GrhIndex = 14687 Or _
           GrhIndex = 11903 Or GrhIndex = 735 Or GrhIndex = 15698 Or GrhIndex = 14504 Or GrhIndex = 15697 Or _
           GrhIndex = 6598 Or GrhIndex = 1121 Or GrhIndex = 1878 Or GrhIndex = 9513 Or GrhIndex = 9514 Or _
           GrhIndex = 9515 Or GrhIndex = 9518 Or GrhIndex = 9519 Or GrhIndex = 9520 Or GrhIndex = 9529
           
        Exit Function


EsArbol_Err:
102     Call TraceError(Err.Number, Err.Description, "General.EsArbol", Erl)

        
End Function

Private Function HayLava(ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
        
        On Error GoTo HayLava_Err
        

        '***************************************************
        'Autor: Nacho (Integer)
        'Last Modification: 03/12/07
        '***************************************************
100     If Map > 0 And Map < NumMaps + 1 And X > 0 And X < 101 And Y > 0 And Y < 101 Then
102         If MapData(Map, X, Y).Graphic(1) >= 5837 And MapData(Map, X, Y).Graphic(1) <= 5852 Then
104             HayLava = True
            Else
106             HayLava = False

            End If

        Else
108         HayLava = False

        End If

        
        Exit Function

HayLava_Err:
110     Call TraceError(Err.Number, Err.Description, "General.HayLava", Erl)

        
End Function

Sub ApagarFogatas()

        'Ladder /ApagarFogatas
        On Error GoTo ErrHandler

        Dim obj As t_Obj
100         obj.ObjIndex = FOGATA_APAG
102         obj.amount = 1

        Dim MapaActual As Long
        Dim Y          As Long
        Dim X          As Long

104     For MapaActual = 1 To NumMaps
106         For Y = YMinMapSize To YMaxMapSize
108             For X = XMinMapSize To XMaxMapSize

110                 If MapInfo(MapaActual).lluvia Then
                
112                     If MapData(MapaActual, X, Y).ObjInfo.ObjIndex = FOGATA Then
                    
114                         Call EraseObj(MAX_INVENTORY_OBJS, MapaActual, X, Y)
116                         Call MakeObj(obj, MapaActual, X, Y)

                        End If

                    End If

118             Next X
120         Next Y
122     Next MapaActual

        Exit Sub
    
ErrHandler:
124     Call LogError("Error producido al apagar las fogatas de " & X & "-" & Y & " del mapa: " & MapaActual & "    -" & Err.Description)

End Sub


Private Sub InicializarConstantes()
        
        On Error GoTo InicializarConstantes_Err
    
        
    
100     LastBackup = Format$(Now, "Short Time")
102     minutos = Format$(Now, "Short Time")
    
104     IniPath = App.Path & "\"

106     ListaRazas(e_Raza.Humano) = "Humano"
108     ListaRazas(e_Raza.Elfo) = "Elfo"
110     ListaRazas(e_Raza.Drow) = "Elfo Oscuro"
112     ListaRazas(e_Raza.Gnomo) = "Gnomo"
114     ListaRazas(e_Raza.Enano) = "Enano"
        ListaRazas(e_Raza.Orco) = "Orco"
    
116     ListaClases(e_Class.Mage) = "Mago"
118     ListaClases(e_Class.Cleric) = "Clérigo"
120     ListaClases(e_Class.Warrior) = "Guerrero"
122     ListaClases(e_Class.Assasin) = "Asesino"
124     ListaClases(e_Class.Bard) = "Bardo"
126     ListaClases(e_Class.Druid) = "Druida"
128     ListaClases(e_Class.Paladin) = "Paladín"
130     ListaClases(e_Class.Hunter) = "Cazador"
132     ListaClases(e_Class.Trabajador) = "Trabajador"
134     ListaClases(e_Class.Pirat) = "Pirata"
136     ListaClases(e_Class.Thief) = "Ladrón"
138     ListaClases(e_Class.Bandit) = "Bandido"
    
140     SkillsNames(e_Skill.Magia) = "Magia"
142     SkillsNames(e_Skill.Robar) = "Robar"
144     SkillsNames(e_Skill.Tacticas) = "Destreza en combate"
146     SkillsNames(e_Skill.Armas) = "Combate con armas"
148     SkillsNames(e_Skill.Meditar) = "Meditar"
150     SkillsNames(e_Skill.Apuñalar) = "Apuñalar"
152     SkillsNames(e_Skill.Ocultarse) = "Ocultarse"
154     SkillsNames(e_Skill.Supervivencia) = "Supervivencia"
156     SkillsNames(e_Skill.Comerciar) = "Comercio"
158     SkillsNames(e_Skill.Defensa) = "Defensa con escudo"
160     SkillsNames(e_Skill.liderazgo) = "Liderazgo"
162     SkillsNames(e_Skill.Proyectiles) = "Armas a distancia"
164     SkillsNames(e_Skill.Wrestling) = "Combate sin armas"
166     SkillsNames(e_Skill.Navegacion) = "Navegación"
168     SkillsNames(e_Skill.equitacion) = "Equitación"
170     SkillsNames(e_Skill.Resistencia) = "Resistencia mágica"
172     SkillsNames(e_Skill.Talar) = "Tala"
174     SkillsNames(e_Skill.Pescar) = "Pesca"
176     SkillsNames(e_Skill.Mineria) = "Minería"
178     SkillsNames(e_Skill.Herreria) = "Herrería"
180     SkillsNames(e_Skill.Carpinteria) = "Carpintería"
182     SkillsNames(e_Skill.Alquimia) = "Alquimia"
184     SkillsNames(e_Skill.Sastreria) = "Sastrería"
186     SkillsNames(e_Skill.Domar) = "Domar"
   
188     ListaAtributos(e_Atributos.Fuerza) = "Fuerza"
190     ListaAtributos(e_Atributos.Agilidad) = "Agilidad"
192     ListaAtributos(e_Atributos.Inteligencia) = "Inteligencia"
194     ListaAtributos(e_Atributos.Constitucion) = "Constitución"
196     ListaAtributos(e_Atributos.Carisma) = "Carisma"
    
200     IniPath = App.Path & "\"
    
        'Bordes del mapa
202     MinXBorder = XMinMapSize + (XWindow \ 2)
204     MaxXBorder = XMaxMapSize - (XWindow \ 2)
206     MinYBorder = YMinMapSize + (YWindow \ 2)
208     MaxYBorder = YMaxMapSize - (YWindow \ 2)

210     RaceHeightOffset(Humano) = -35
212     RaceHeightOffset(Elfo) = -35
214     RaceHeightOffset(Drow) = -35
216     RaceHeightOffset(Gnomo) = -27
218     RaceHeightOffset(Enano) = -27
300     RaceHeightOffset(Orco) = -35

        WeaponTypeNames(eSword) = "Sword"
        WeaponTypeNames(eDagger) = "Dagger"
        WeaponTypeNames(eBow) = "Bow"
        WeaponTypeNames(eStaff) = "Staff"
        WeaponTypeNames(eMace) = "Mace"
        WeaponTypeNames(eThrowableAxe) = "ThrowableAxe"
        WeaponTypeNames(eAxe) = "Axe"
        WeaponTypeNames(eKnuckle) = "Knuckle"
        WeaponTypeNames(e_WeaponType.eFist) = "Fist"
        WeaponTypeNames(e_WeaponType.eSpear) = "Spear"
        WeaponTypeNames(e_WeaponType.eGunPowder) = "GunPowder"
        
        Exit Sub

InicializarConstantes_Err:
302     Call TraceError(Err.Number, Err.Description, "General.InicializarConstantes", Erl)
End Sub

Sub Main()
        On Error GoTo Handler
        Call InitializeCircularLogBuffer
        Call LogThis(0, "Starting the server " & Now, vbLogEventTypeInformation)

        Call load_stats
        ' Me fijo si ya hay un proceso llamado server.exe abierto
100     If GetProcessCount(App.EXEName & ".exe") > 1 Then
            ' Si lo hay, pregunto si lo queremos cerrar.
102         If MsgBox("Se ha encontrado mas de 1 instancia abierta de esta aplicación, ¿Desea continuar?", vbYesNo) = vbNo Then
104             End
            End If
        End If

        Dim f As Date
    
110     Call ChDir(App.Path)
112     Call ChDrive(App.Path)

114     Call InicializarConstantes
    
116     frmCargando.Show
    
118     frmMain.Caption = frmMain.Caption & " V." & App.Major & "." & App.Minor & "." & App.Revision
    
120     frmCargando.Label1(2).Caption = "Iniciando Arrays..."

        Call InitializeNpcIndexHeap
        Call InitializeLobbyList
126     Call loadAdministrativeUsers

        '¿?¿?¿?¿?¿?¿?¿?¿ CARGAMOS DATOS DESDE ARCHIVOS ¿??¿?¿?¿?¿?¿?¿?¿
128     frmCargando.Label1(2).Caption = "Cargando Server.ini"
    
130     MaxUsers = 0
132     Call LoadSini
137     Call LoadMD5
135     Call LoadPacketRatePolicy
#If PYMMO = 1 Then
133     Call LoadPrivateKey
#End If
138     Call LoadConfiguraciones
140     Call LoadIntervalos
142     Call CargarForbidenWords
144     Call CargaApuestas
146     Call CargarSpawnList
148     Call LoadMotd
        Call initBase64Chars
        
152     frmCargando.Label1(2).Caption = "Conectando base de datos y limpiando usuarios logueados"
        
        If Not FileExist(App.Path & "/" & DatabaseFileName) Then
            Call FileSystem.FileCopy(App.Path & "/Empty_db.db", App.Path & "/" & DatabaseFileName)
        End If
        ' ************************* Base de Datos ********************
        'Conecto base de datos
154     Call Database_Connect
        
        Call Database_Connect_Async
        
        ' Construimos las querys grandes
156     Call Contruir_Querys
158     Call LoadDBMigrations
        ' ******************* FIN - Base de Datos ********************
160     Call LoadGuildsDB
        '*************************************************
164     frmCargando.Label1(2).Caption = "Cargando NPCs.Dat"
166     Call CargaNpcsDat
        '*************************************************
    
168     frmCargando.Label1(2).Caption = "Cargando Obj.Dat"

170     Call LoadOBJData
        
172     frmCargando.Label1(2).Caption = "Cargando Hechizos.Dat"
174     Call CargarHechizos

        frmCargando.Label1(2).Caption = "Cargando EffectsOverTime.Dat"
175     Call LoadEffectOverTime
        
176     frmCargando.Label1(2).Caption = "Cargando Objetos de Herrería"
178     Call LoadArmasHerreria
180     Call LoadArmadurasHerreria
    
182     frmCargando.Label1(2).Caption = "Cargando Objetos de Carpintería"
184     Call LoadObjCarpintero
    
186     frmCargando.Label1(2).Caption = "Cargando Objetos de Alquimista"
188     Call LoadObjAlquimista
    
190     frmCargando.Label1(2).Caption = "Cargando Objetos de Sastre"
192     Call LoadObjSastre
    
194     frmCargando.Label1(2).Caption = "Cargando Pesca"
196     Call LoadPesca
    
198     frmCargando.Label1(2).Caption = "Cargando Recursos Especiales"
200     Call LoadRecursosEspeciales

202     frmCargando.Label1(2).Caption = "Cargando Rangos de Faccion"
204     Call LoadRangosFaccion

206     frmCargando.Label1(2).Caption = "Cargando Recompensas de Faccion"
208     Call LoadRecompensasFaccion
    
210     frmCargando.Label1(2).Caption = "Cargando Balance.dat"
212     Call LoadBalance    '4/01/08 Pablo ToxicWaste
    
214     frmCargando.Label1(2).Caption = "Cargando Ciudades.dat"
216     Call CargarCiudades
    
218     If BootDelBackUp Then
220         frmCargando.Label1(2).Caption = "Cargando WorldBackup"
222         Call CargarBackUp
        Else
224         frmCargando.Label1(2).Caption = "Cargando Mapas"
226         Call LoadMapData
        End If
        
        frmCargando.Label1(2).Caption = "Cargando Zonas y Areas"
        DoEvents
        Call LoadZonas
        
        frmCargando.Label1(2).Caption = "Cargando donadores"
        Call CargarDonadores
        
228     Call InitPathFinding

230     frmCargando.Label1(2).Caption = "Cargando informacion de eventos"
232     Call CargarInfoRetos
234     Call CargarInfoEventos
    
242     frmCargando.Label1(2).Caption = "Cargando Baneos Temporales"
244     Call LoadBans
    
246     frmCargando.Label1(2).Caption = "Cargando Quests"
252     Call LoadQuests

254     EstadoGlobal = False
    
        Call ResetLastLogoutAndIsLogged

        'Comentado porque hay worldsave en ese mapa!
    
        Dim LoopC As Integer
    
        'Resetea las conexiones de los usuarios
258     For LoopC = 1 To MaxUsers
262         UserList(LoopC).ConnectionDetails.ConnIDValida = False
268     Next LoopC
    
        '¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
    
270     With frmMain
272         .Minuto.Enabled = True
274         .TimerGuardarUsuarios.Enabled = True
276         .TimerGuardarUsuarios.Interval = IntervaloTimerGuardarUsuarios
278         .tPiqueteC.Enabled = True
280         .GameTimer.Enabled = True
282         .Segundo.Enabled = True
284         .KillLog.Enabled = True
286         .TIMER_AI.Enabled = True
            .T_UsersOnline.Enabled = True
            .t_Extraer.Enabled = True
            .t_Extraer.Interval = IntervaloTrabajarExtraer
            .tControlHechizos.Enabled = True
            .tControlHechizos.Interval = 60000
            If IsFeatureEnabled("ShipTravelEnabled") Then
                .TimerBarco.Enabled = True
                MapInfo(BarcoNavegando.map).ForceUpdate = True
            End If
            
            
        End With
    
290     Subasta.SubastaHabilitada = True
292     Subasta.HaySubastaActiva = False
294     Call ResetMeteo
    
        ' ----------------------------------------------------
        '           Configuracion de los sockets
        ' ----------------------------------------------------
296     Call InitializePacketList

302     Call modNetwork.Listen(MaxUsers, ListenIp, CStr(Puerto))

312     If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."
        ' ----------------------------------------------------
        '           Configuracion de los sockets
        ' ----------------------------------------------------
    
314     Call GetHoraActual
    
316     HoraMundo = GetTickCount() - DuracionDia \ 2

318     frmCargando.Visible = False
320     Unload frmCargando
    
        'Ocultar
330     Call frmMain.InitMain(HideMe)

        Call InitializeAntiCheat
    
    
332     tInicioServer = GetTickCount()
        #If UNIT_TEST = 1 Then
                    Call UnitTesting.init
                    Debug.Print "AO20 Unit Testing"
                    Dim suite_passed_ok As Boolean
                    suite_passed_ok = UnitTesting.test_suite()
                    If (suite_passed_ok) Then
                        Debug.Print "suite_passed_ok!!!"
                    Else
                        Debug.Print "suite failed!!!"
                    End If
                    Debug.Assert (suite_passed_ok)
                    
                    Debug.Print "Running proto suite, trying to connect to 127.0.0.1:7667"
                    Call UnitClient.Init
                    Call UnitClient.Connect("127.0.0.1", "7667")
        #End If
        
            
        While (True)
            GlobalFrameTime = GetTickCount()
            Dim PerformanceTimer As Long
            Call PerformanceTestStart(PerformanceTimer)
#If PYMMO = 1 Then
            Call modNetwork.close_not_logged_sockets_if_timeout
#End If
            Call PerformTimeLimitCheck(PerformanceTimer, "General modNetwork.close_not_logged_sockets_if_timeout")
            Call modNetwork.Tick(GetElapsed())
            Call PerformTimeLimitCheck(PerformanceTimer, "General modNetwork.Tick")
            Call UpdateEffectOverTime
            Call PerformTimeLimitCheck(PerformanceTimer, "General Update Effects over time")
            DoEvents
            Call PerformTimeLimitCheck(PerformanceTimer, "Do events")
            Call AntiCheatUpdate
            Call PerformTimeLimitCheck(PerformanceTimer, "Update anti cheat")
            ' Unlock main loop for maximum throughput but it can hog weak CPUs.
            #If UNLOCK_CPU = 0 Then
                Call Sleep(1)
            #End If
            
            #If UNIT_TEST = 1 Then
                Call UnitClient.Poll
            #End If
        Wend
        Call LogThis(0, "Closing the server " & Now, vbLogEventTypeInformation)
        Exit Sub
Handler:
334     Call TraceError(Err.Number, Err.Description, "General.Main", Erl)
End Sub

Function FileExist(ByVal File As String, Optional FileType As VbFileAttribute = vbNormal) As Boolean
        '*****************************************************************
        'Se fija si existe el archivo
        '*****************************************************************
        
        On Error GoTo FileExist_Err
        
100     FileExist = LenB(dir$(File, FileType)) <> 0

        
        Exit Function

FileExist_Err:
102     Call TraceError(Err.Number, Err.Description, "General.FileExist", Erl)

        
End Function

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
        
        On Error GoTo ReadField_Err
        

        '*****************************************************************
        'Gets a field from a string
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modify Date: 11/15/2004
        'Gets a field from a delimited string
        '*****************************************************************
        Dim i          As Long

        Dim LastPos    As Long

        Dim currentPos As Long

        Dim delimiter  As String * 1
    
100     delimiter = Chr$(SepASCII)
    
102     For i = 1 To Pos
104         LastPos = currentPos
106         currentPos = InStr(LastPos + 1, Text, delimiter, vbBinaryCompare)
108     Next i
    
110     If currentPos = 0 Then
112         ReadField = mid$(Text, LastPos + 1, Len(Text) - LastPos)
        Else
114         ReadField = mid$(Text, LastPos + 1, currentPos - LastPos - 1)

        End If

        
        Exit Function

ReadField_Err:
116     Call TraceError(Err.Number, Err.Description, "General.ReadField", Erl)

        
End Function

Function MapaValido(ByVal Map As Integer) As Boolean
        
        On Error GoTo MapaValido_Err
        
100     MapaValido = Map >= 1 And Map <= NumMaps

        
        Exit Function

MapaValido_Err:
102     Call TraceError(Err.Number, Err.Description, "General.MapaValido", Erl)

        
End Function

Sub MostrarNumUsers()

        On Error GoTo MostrarNumUsers_Err
        
        
100         Call SendData(SendTarget.ToAll, 0, PrepareMessageOnlineUser(NumUsers))

102         frmMain.CantUsuarios.Caption = "Numero de usuarios jugando: " & NumUsers
        

        Exit Sub

MostrarNumUsers_Err:
106     Call TraceError(Err.Number, Err.Description, "General.MostrarNumUsers", Erl)

        
End Sub




Sub Restart()
        
        On Error GoTo Restart_Err
        
        'Se asegura de que los sockets estan cerrados e ignora cualquier err

100     If frmMain.Visible Then frmMain.txStatus.Caption = "Reiniciando."

        Dim LoopC As Long

102     Call modNetwork.Disconnect

104     For LoopC = 1 To MaxUsers
106         Call CloseSocket(LoopC)
        Next

        'Initialize statistics!!
        'Call Statistics.Initialize

116     ReDim UserList(1 To MaxUsers) As t_User

118     For LoopC = 1 To MaxUsers
122         UserList(LoopC).ConnectionDetails.ConnIDValida = False
128     Next LoopC
        Call InitializeUserIndexHeap(MaxUsers)
130     LastUser = 0
132     NumUsers = 0

134     Call FreeNPCs
136     Call FreeCharIndexes

138     Call LoadSini
        Call LoadMD5
133     Call LoadPrivateKey
140     Call LoadIntervalos
142     Call LoadOBJData
144     Call LoadPesca
146     Call LoadRecursosEspeciales

148     Call LoadMapData

150     Call CargarHechizos
        
        Call modNetwork.Listen(MaxUsers, ListenIp, CStr(Puerto))
        
152     If frmMain.Visible Then frmMain.txStatus.Caption = "Escuchando conexiones entrantes ..."


        'Ocultar
162     Call frmMain.InitMain(HideMe)
    
        Exit Sub

Restart_Err:
164     Call TraceError(Err.Number, Err.Description, "General.Restart", Erl)

        
End Sub

Public Function Intemperie(ByVal UserIndex As Integer) As Boolean
        
        On Error GoTo Intemperie_Err
        
    
100     If MapInfo(UserList(UserIndex).Pos.Map).zone <> "DUNGEON" Then
102         If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger <> 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger < 10 Then Intemperie = True
        Else
104         Intemperie = False

        End If
    
        
        Exit Function

Intemperie_Err:
106     Call TraceError(Err.Number, Err.Description, "General.Intemperie", Erl)

        
End Function

Public Sub TiempoInvocacion(ByVal UserIndex As Integer)
        
        On Error GoTo TiempoInvocacion_Err
    
        
        Dim i As Integer
100     For i = 1 To MAXMASCOTAS
102         If UserList(UserIndex).MascotasIndex(i).ArrayIndex > 0 Then
                If Not IsValidNpcRef(UserList(UserIndex).MascotasIndex(i)) Then
                    Call ClearNpcRef(UserList(UserIndex).MascotasIndex(i))
                Else
104                If NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia > 0 Then
106                    NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = _
                       NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia - 1
108                    If NpcList(UserList(UserIndex).MascotasIndex(i).ArrayIndex).Contadores.TiempoExistencia = 0 Then Call MuereNpc(UserList(UserIndex).MascotasIndex(i).ArrayIndex, 0)
                   End If
                End If
            End If
110     Next i
        
        Exit Sub

TiempoInvocacion_Err:
112     Call TraceError(Err.Number, Err.Description, "General.TiempoInvocacion", Erl)

        
End Sub

Public Sub EfectoFrio(ByVal UserIndex As Integer)
        On Error GoTo EfectoFrio_Err
100     If Not Intemperie(UserIndex) Then Exit Sub
102     With UserList(UserIndex)
104         If .Invent.ArmourEqpObjIndex > 0 Then
                ' WyroX: Ropa invernal
106             If ObjData(.Invent.ArmourEqpObjIndex).Invernal Then Exit Sub
            End If
108         If .Counters.Frio < IntervaloFrio Then
110             .Counters.Frio = .Counters.Frio + 1
            Else
112             If MapInfo(.Pos.Map).terrain = Nieve Then
114                 Call WriteConsoleMsg(UserIndex, "¡Estás muriendo de frío, abrígate o morirás!", e_FontTypeNames.FONTTYPE_INFO)
                    ' WyroX: Sin ropa perdés vida más rápido que con una ropa no-invernal
                    Dim MinDamage As Integer, MaxDamage As Integer
116                 If .flags.Desnudo = 0 Then
118                     MinDamage = 17
120                     MaxDamage = 23
                    Else
122                     MinDamage = 27
124                     MaxDamage = 33
                    End If

                    ' WyroX: Agrego aleatoriedad
                    Dim Damage As Integer
126                 Damage = Porcentaje(.Stats.MaxHp, RandomNumber(MinDamage, MaxDamage))
128                 If UserMod.ModifyHealth(UserIndex, -Damage, 0) Then
132                     Call WriteConsoleMsg(UserIndex, "¡Has muerto de frío!", e_FontTypeNames.FONTTYPE_INFO)
134                     Call UserMod.UserDie(UserIndex)
                    End If
                End If
138             .Counters.Frio = 0
            End If
        End With
        Exit Sub
EfectoFrio_Err:
140     Call TraceError(Err.Number, Err.Description, "General.EfectoFrio", Erl)
End Sub

Public Sub EfectoStamina(ByVal UserIndex As Integer)

    Dim HambreOSed As Boolean
    Dim bEnviarStats_HP As Boolean
    Dim bEnviarStats_STA As Boolean
    
100 With UserList(UserIndex)
102     HambreOSed = .Stats.MinHam = 0 Or .Stats.MinAGU = 0
    
104     If Not HambreOSed Then 'Si no tiene hambre ni sed
106         If .Stats.MinHp < .Stats.MaxHp Then
108             Call Sanar(UserIndex, bEnviarStats_HP, IIf(.flags.Descansar, SanaIntervaloDescansar, SanaIntervaloSinDescansar))
            End If
        End If
                                
110     If .flags.Desnudo = 0 And Not HambreOSed Then
112         If Not Lloviendo Or Not Intemperie(UserIndex) Then
114             Call RecStamina(UserIndex, bEnviarStats_STA, IIf(.flags.Descansar, StaminaIntervaloDescansar, StaminaIntervaloSinDescansar))
            End If
        Else
116         If Lloviendo And Intemperie(UserIndex) Then
118             Call PierdeEnergia(UserIndex, bEnviarStats_STA, IntervaloPerderStamina * 0.5)
            Else
120             Call PierdeEnergia(UserIndex, bEnviarStats_STA, IIf(.flags.Descansar, IntervaloPerderStamina * 2, IntervaloPerderStamina))
            End If
        End If
        
122     If .flags.Descansar Then
            'termina de descansar automaticamente
124         If .Stats.MaxHp = .Stats.MinHp And .Stats.MaxSta = .Stats.MinSta Then
126             Call WriteRestOK(UserIndex)
128             Call WriteConsoleMsg(UserIndex, "Has terminado de descansar.", e_FontTypeNames.FONTTYPE_INFO)
130             .flags.Descansar = False
            End If
        
        End If
        
132     If bEnviarStats_STA Then
134         Call WriteUpdateSta(UserIndex)
        End If
        
136     If bEnviarStats_HP Then
138         Call WriteUpdateHP(UserIndex)
        End If
    End With
End Sub

Public Sub EfectoLava(ByVal UserIndex As Integer)
        On Error GoTo EfectoLava_Err
100     With UserList(UserIndex)
102         If .Counters.Lava < IntervaloFrio Then 'Usamos el mismo intervalo que el del frio
104             .Counters.Lava = .Counters.Lava + 1
            Else
106             If HayLava(.Pos.Map, .Pos.X, .Pos.Y) Then
108                 Call WriteConsoleMsg(UserIndex, "¡Quítate de la lava, te estás quemando!", e_FontTypeNames.FONTTYPE_INFO)
110                 If UserMod.ModifyHealth(UserIndex, -Porcentaje(.Stats.MaxHp, 5)) Then
114                     Call WriteConsoleMsg(UserIndex, "¡Has muerto quemado!", e_FontTypeNames.FONTTYPE_INFO)
                        Call CustomScenarios.UserDie(UserIndex)
116                     Call UserMod.UserDie(UserIndex)
                    End If
                End If
120             .Counters.Lava = 0
            End If
        End With
        Exit Sub
EfectoLava_Err:
122     Call TraceError(Err.Number, Err.Description, "General.EfectoLava", Erl)
End Sub

''
' Maneja el tiempo y el efecto del mimetismo
'
' @param UserIndex  El index del usuario a ser afectado por el mimetismo
'

Public Sub EfectoMimetismo(ByVal UserIndex As Integer)
    '******************************************************
    'Author: Unknown
    'Last Update: 04/11/2008 (NicoNZ)
    '
    '******************************************************
        
        On Error GoTo EfectoMimetismo_Err
    
        
        Dim Barco As t_ObjData
    
100     With UserList(UserIndex)
102         If .Counters.Mimetismo < IntervaloInvisible Then
104             .Counters.Mimetismo = .Counters.Mimetismo + 1

            Else
                'restore old char
106             Call WriteConsoleMsg(UserIndex, "Recuperas tu apariencia normal.", e_FontTypeNames.FONTTYPE_INFO)
            
108             If .flags.Navegando Then
110                 Call EquiparBarco(UserIndex)
                Else
112                 .Char.Body = .CharMimetizado.Body
114                 .Char.Head = .CharMimetizado.Head
116                 .Char.CascoAnim = .CharMimetizado.CascoAnim
118                 .Char.ShieldAnim = .CharMimetizado.ShieldAnim
120                 .Char.WeaponAnim = .CharMimetizado.WeaponAnim
121                 .char.CartAnim = .CharMimetizado.CartAnim
                End If
                
122             .Counters.Mimetismo = 0
124             .flags.Mimetizado = e_EstadoMimetismo.Desactivado
            
126             With .Char
128                 Call ChangeUserChar(UserIndex, .body, .head, .Heading, .WeaponAnim, .ShieldAnim, .CascoAnim, .CartAnim)
130                 Call RefreshCharStatus(UserIndex)
                End With
                
            End If
            
        End With
        
        Exit Sub

EfectoMimetismo_Err:
132     Call TraceError(Err.Number, Err.Description, "General.EfectoMimetismo", Erl)

        
End Sub

Public Sub EfectoInvisibilidad(ByVal UserIndex As Integer)
        
    On Error GoTo EfectoInvisibilidad_Err
        
    With UserList(UserIndex)
        If .Counters.Invisibilidad > 0 Then
            .Counters.Invisibilidad = UserList(UserIndex).Counters.Invisibilidad - 1
            If .Counters.DisabledInvisibility > 0 Then
                .Counters.DisabledInvisibility = .Counters.DisabledInvisibility - 1
                If .Counters.DisabledInvisibility = 0 And .Counters.Invisibilidad > 0 Then
                    .flags.invisible = 1
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, .pos.x, .pos.y))
                End If
            End If
        Else
            .Counters.Invisibilidad = 0
            .flags.invisible = 0
            .Counters.DisabledInvisibility = 0
            If .flags.Oculto = 0 Then
                ' Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", e_FontTypeNames.FONTTYPE_INFO)
                Call WriteLocaleMsg(UserIndex, "307", e_FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False, .pos.x, .pos.y))
                Call WriteContadores(UserIndex)
            End If
        End If
    End With
        
    Exit Sub

EfectoInvisibilidad_Err:
    Call TraceError(Err.Number, Err.Description, "General.EfectoInvisibilidad", Erl)

        
End Sub

Public Sub EfectoParalisisNpc(ByVal NpcIndex As Integer)
        On Error GoTo EfectoParalisisNpc_Err
        
100     If NpcList(NpcIndex).Contadores.Paralisis > 0 Then
102         NpcList(NpcIndex).Contadores.Paralisis = NpcList(NpcIndex).Contadores.Paralisis - 1
        Else
104         NpcList(NpcIndex).flags.Paralizado = 0

        End If
        
        Exit Sub

EfectoParalisisNpc_Err:
106     Call TraceError(Err.Number, Err.Description, "General.EfectoParalisisNpc", Erl)

        
End Sub

Public Sub EfectoInmovilizadoNpc(ByVal NpcIndex As Integer)
        On Error GoTo EfectoInmovilizadoNpc_Err

100     If NpcList(NpcIndex).Contadores.Inmovilizado > 0 Then
102         NpcList(NpcIndex).Contadores.Inmovilizado = NpcList(NpcIndex).Contadores.Inmovilizado - 1
        Else
104         NpcList(NpcIndex).flags.Inmovilizado = 0

        End If

        Exit Sub

EfectoInmovilizadoNpc_Err:
106     Call TraceError(Err.Number, Err.Description, "General.EfectoInmovilizadoNpc", Erl)

        
End Sub


Public Sub EfectoCeguera(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoCeguera_Err
        

100     If UserList(UserIndex).Counters.Ceguera > 0 Then
102         UserList(UserIndex).Counters.Ceguera = UserList(UserIndex).Counters.Ceguera - 1
        Else

104         If UserList(UserIndex).flags.Ceguera = 1 Then
106             UserList(UserIndex).flags.Ceguera = 0
108             Call WriteBlindNoMore(UserIndex)

            End If

        End If

        
        Exit Sub

EfectoCeguera_Err:
110     Call TraceError(Err.Number, Err.Description, "General.EfectoCeguera", Erl)

        
End Sub

Public Sub EfectoEstupidez(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoEstupidez_Err
        

100     If UserList(UserIndex).Counters.Estupidez > 0 Then
102         UserList(UserIndex).Counters.Estupidez = UserList(UserIndex).Counters.Estupidez - 1

        Else

104         If UserList(UserIndex).flags.Estupidez = 1 Then
106             UserList(UserIndex).flags.Estupidez = 0
108             Call WriteDumbNoMore(UserIndex)

            End If

        End If

        
        Exit Sub

EfectoEstupidez_Err:
110     Call TraceError(Err.Number, Err.Description, "General.EfectoEstupidez", Erl)

        
End Sub

Public Sub EfectoParalisisUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoParalisisUser_Err
        
        With UserList(UserIndex)
100         If .Counters.Paralisis > 0 Then
102             .Counters.Paralisis = .Counters.Paralisis - 1
            Else
104             .flags.Paralizado = 0
    
                If .clase = e_Class.Warrior Or .clase = e_Class.Hunter Or .clase = e_Class.Thief Or .clase = e_Class.Pirat Then
                    .Counters.TiempoDeInmunidadParalisisNoMagicas = 3
                End If
                'UserList(UserIndex).Flags.AdministrativeParalisis = 0
106             Call WriteParalizeOK(UserIndex)
    
            End If
        End With

        
        Exit Sub

EfectoParalisisUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoParalisisUser", Erl)

        
End Sub

Public Sub EfectoVelocidadUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoVelocidadUser_Err
        

100     If UserList(UserIndex).Counters.velocidad > 0 Then
102         UserList(UserIndex).Counters.velocidad = UserList(UserIndex).Counters.velocidad - 1
        Else
104         UserList(UserIndex).flags.VelocidadHechizada = 0
106         Call ActualizarVelocidadDeUsuario(UserIndex)
        End If

        Exit Sub

EfectoVelocidadUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoVelocidadUser", Erl)

        
End Sub

Public Sub EfectoMaldicionUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoMaldicionUser_Err
        

100     If UserList(UserIndex).Counters.Maldicion > 0 Then
102         UserList(UserIndex).Counters.Maldicion = UserList(UserIndex).Counters.Maldicion - 1
    
        Else
104         UserList(UserIndex).flags.Maldicion = 0
106         Call WriteConsoleMsg(UserIndex, "¡La magia perdió su efecto! Ya puedes atacar.", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
        End If

        
        Exit Sub

EfectoMaldicionUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoMaldicionUser", Erl)

        
End Sub

Public Sub EfectoInmoUser(ByVal UserIndex As Integer)
        
        On Error GoTo EfectoInmoUser_Err
        
        With UserList(UserIndex)
100         If .Counters.Inmovilizado > 0 Then
102             .Counters.Inmovilizado = .Counters.Inmovilizado - 1
            Else
104             .flags.Inmovilizado = 0

                If .clase = e_Class.Warrior Or .clase = e_Class.Hunter Or .clase = e_Class.Thief Or .clase = e_Class.Pirat Then
                    .Counters.TiempoDeInmunidadParalisisNoMagicas = 3
                End If
106             Call WriteInmovilizaOK(UserIndex)
    
            End If
        End With

        
        Exit Sub

EfectoInmoUser_Err:
108     Call TraceError(Err.Number, Err.Description, "General.EfectoInmoUser", Erl)

        
End Sub

Public Sub RecStamina(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
            On Error GoTo RecStamina_Err

            Dim trigger As Byte
            Dim Suerte As Integer

100         With UserList(UserIndex)
102             trigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger

104             If trigger = 1 And trigger = 2 And trigger = 4 Then Exit Sub

106             If .Stats.MinSta < .Stats.MaxSta Then

108                 If .Counters.STACounter < Intervalo Then
110                     .Counters.STACounter = .Counters.STACounter + 1
                        Exit Sub

                    End If

112                 .Counters.STACounter = 0

114                 If .Counters.Trabajando > 0 Or IsSet(.flags.StatusMask, ePreventEnergyRestore) Then Exit Sub  'Trabajando no sube energía. (ToxicWaste)

116                 EnviarStats = True

118                 Select Case .Stats.UserSkills(e_Skill.Supervivencia)
                        Case 0 To 10
120                         Suerte = 5
122                     Case 11 To 20
124                         Suerte = 7
126                     Case 21 To 30
128                         Suerte = 9
130                     Case 31 To 40
132                         Suerte = 11
134                     Case 41 To 50
136                         Suerte = 13
138                     Case 51 To 60
140                         Suerte = 15
142                     Case 61 To 70
144                         Suerte = 17
146                     Case 71 To 80
148                         Suerte = 19
150                     Case 81 To 90
152                         Suerte = 21
154                     Case 91 To 99
156                         Suerte = 23
158                     Case 100
160                         Suerte = 25
                    End Select
                    
                    Dim NuevaStamina As Long
                        If .clase = e_Class.Trabajador Then
                            NuevaStamina = .Stats.MinSta + RandomNumber(1, CInt(Porcentaje(.Stats.MaxSta, Suerte)))
                        Else
                            NuevaStamina = .Stats.MinSta + RandomNumber(1, CInt(Porcentaje(.Stats.MaxSta, Suerte)) / 1.6)
                        End If

                    
                    ' Jopi: Prevenimos overflow al acotar la stamina que se puede recuperar en cualquier caso.
                    ' Cuando te editabas la energia con el GM causaba este error.
166                 If NuevaStamina < 32000 Then
168                     .Stats.MinSta = NuevaStamina
                    Else
170                     .Stats.MinSta = 32000
                    End If

172                 If .Stats.MinSta > .Stats.MaxSta Then
174                     .Stats.MinSta = .Stats.MaxSta
                    End If

                End If
            End With

            Exit Sub

RecStamina_Err:
176         Call TraceError(Err.Number, Err.Description, "General.RecStamina", Erl)


End Sub

Public Sub PierdeEnergia(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)

        On Error GoTo RecStamina_Err

100     With UserList(UserIndex)

102         If .Stats.MinSta > 0 Then
    
104             If .Counters.STACounter < Intervalo Then
106                 .Counters.STACounter = .Counters.STACounter + 1
                Else
            
108                 .Counters.STACounter = 0
    
110                 EnviarStats = True
            
                    Dim Cantidad As Integer
    
112                 Cantidad = RandomNumber(1, Porcentaje(.Stats.MaxSta, (MAXSKILLPOINTS * 1.5 - .Stats.UserSkills(e_Skill.Supervivencia)) * 0.25))
114                 .Stats.MinSta = .Stats.MinSta - Cantidad
    
116                 If .Stats.MinSta < 0 Then
118                     .Stats.MinSta = 0
                    End If
    
                End If
    
            End If

        End With
        
        Exit Sub

RecStamina_Err:
120     Call TraceError(Err.Number, Err.Description, "General.PierdeEnergia", Erl)

        
End Sub

Public Sub EfectoVeneno(ByVal UserIndex As Integer)
    On Error GoTo EfectoVeneno_Err
        Dim damage As Long
100     If UserList(UserIndex).Counters.Veneno < IntervaloVeneno Then
102         UserList(UserIndex).Counters.Veneno = UserList(UserIndex).Counters.Veneno + 1
        Else
104         Call CancelExit(UserIndex)
106         With UserList(UserIndex)
              'Call WriteConsoleMsg(UserIndex, "Estás envenenado, si no te curas morirás.", e_FontTypeNames.FONTTYPE_VENENO)
108           Call WriteLocaleMsg(UserIndex, "47", e_FontTypeNames.FONTTYPE_VENENO)
              UserList(userindex).Counters.timeFx = 3
110           Call SendData(SendTarget.ToPCAliveArea, userindex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Envenena, 30, False, , UserList(userindex).Pos.X, UserList(userindex).Pos.y))
112           .Counters.Veneno = 0
              ' El veneno saca un porcentaje de vida random.
114           damage = RandomNumber(3, 5)
116           damage = (1 + damage * .Stats.MaxHp \ 100) ' Redondea para arriba
120           If .ChatCombate = 1 Then
                  ' "El veneno te ha causado ¬1 puntos de daño."
122               Call WriteLocaleMsg(UserIndex, "390", e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(damage))
              End If
1224           If UserMod.ModifyHealth(UserIndex, -Damage) Then
                  Call CustomScenarios.UserDie(UserIndex)
126               Call UserMod.UserDie(UserIndex)
              End If
            End With
        End If
        Exit Sub
EfectoVeneno_Err:
130     Call TraceError(Err.Number, Err.Description, "General.EfectoVeneno", Erl)
End Sub


' El incineramiento tiene una logica particular, que es hacer daño sostenido en el tiempo.
Public Sub EfectoIncineramiento(ByVal UserIndex As Integer)
    On Error GoTo EfectoIncineramiento_Err
            Dim damage As Integer
100         With UserList(UserIndex)
                ' 4 Mini intervalitos, dentro del intervalo total de incineracion
102             If .Counters.Incineracion Mod (IntervaloIncineracion \ 4) = 0 Then
                    ' "Te estás incinerando, si no te curas morirás.
104                 Call WriteLocaleMsg(UserIndex, "392", e_FontTypeNames.FONTTYPE_FIGHT)
                    UserList(userindex).Counters.timeFx = 3
106                 Damage = RandomNumber(20, 30)
108                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 73, 0, .pos.x, .pos.y))
112                 If .ChatCombate = 1 Then
114                     Call WriteLocaleMsg(UserIndex, "391", e_FontTypeNames.FONTTYPE_FIGHT, PonerPuntos(damage))
                    End If
110                 If UserMod.ModifyHealth(UserIndex, -Damage) Then
                        Call CustomScenarios.UserDie(UserIndex)
                        Call UserMod.UserDie(UserIndex)
                    End If
                End If
122             .Counters.Incineracion = .Counters.Incineracion + 1
124             If .Counters.Incineracion > IntervaloIncineracion Then
                    ' Se termino la incineracion
126                 .flags.Incinerado = 0
128                 .Counters.Incineracion = 0
                    Exit Sub
                End If
            End With
            Exit Sub
EfectoIncineramiento_Err:
130         Call TraceError(Err.Number, Err.Description, "General.EfectoIncineramiento", Erl)
End Sub

Public Sub DuracionPociones(ByVal UserIndex As Integer)
        
        On Error GoTo DuracionPociones_Err
        

        'Controla la duracion de las pociones
100     If UserList(UserIndex).flags.DuracionEfecto > 0 Then
102         UserList(UserIndex).flags.DuracionEfecto = UserList(UserIndex).flags.DuracionEfecto - 1

104         If UserList(UserIndex).flags.DuracionEfecto = 0 Then
106             UserList(UserIndex).flags.TomoPocion = False
108             UserList(UserIndex).flags.TipoPocion = 0

                'volvemos los atributos al estado normal
                Dim LoopX As Integer

110             For LoopX = 1 To NUMATRIBUTOS
112                 UserList(UserIndex).Stats.UserAtributos(LoopX) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopX)
                Next
114             Call WriteFYA(UserIndex)

            End If

        End If

        
        Exit Sub

DuracionPociones_Err:
116     Call TraceError(Err.Number, Err.Description, "General.DuracionPociones", Erl)

        
End Sub

Public Function HambreYSed(ByVal UserIndex As Integer) As Boolean
         
        On Error GoTo HambreYSed_Err
        

100     If (UserList(UserIndex).flags.Privilegios And e_PlayerType.user) = 0 Then Exit Function

        'Sed
102     If UserList(UserIndex).Stats.MinAGU > 0 Then
104         If UserList(UserIndex).Counters.AGUACounter < IntervaloSed Then
106             UserList(UserIndex).Counters.AGUACounter = UserList(UserIndex).Counters.AGUACounter + 1
            Else
108             UserList(UserIndex).Counters.AGUACounter = 0
110             UserList(UserIndex).Stats.MinAGU = UserList(UserIndex).Stats.MinAGU - 10
        
112             If UserList(UserIndex).Stats.MinAGU <= 0 Then
114                 UserList(UserIndex).Stats.MinAGU = 0

                End If
        
118             HambreYSed = True

            End If

        End If

        'hambre
120     If UserList(UserIndex).Stats.MinHam > 0 Then
122         If UserList(UserIndex).Counters.COMCounter < IntervaloHambre Then
124             UserList(UserIndex).Counters.COMCounter = UserList(UserIndex).Counters.COMCounter + 1
            Else
126             UserList(UserIndex).Counters.COMCounter = 0
128             UserList(UserIndex).Stats.MinHam = UserList(UserIndex).Stats.MinHam - 10

130             If UserList(UserIndex).Stats.MinHam <= 0 Then
132                 UserList(UserIndex).Stats.MinHam = 0

                End If

136             HambreYSed = True

            End If

        End If

        
        Exit Function

HambreYSed_Err:
138     Call TraceError(Err.Number, Err.Description, "General.HambreYSed", Erl)

        
End Function

Public Sub Sanar(ByVal UserIndex As Integer, ByRef EnviarStats As Boolean, ByVal Intervalo As Integer)
        
        On Error GoTo Sanar_Err
        
        ' Desnudo no regenera vida
100     If UserList(UserIndex).flags.Desnudo = 1 Then Exit Sub
        
102     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 2 And MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 4 Then Exit Sub

        Dim mashit As Integer

        'con el paso del tiempo va sanando....pero muy lentamente ;-)
104         If UserList(UserIndex).flags.RegeneracionHP = 1 Then
106             Intervalo = 400

            End If
    
108         If UserList(UserIndex).Counters.HPCounter < Intervalo Then
110             UserList(UserIndex).Counters.HPCounter = UserList(UserIndex).Counters.HPCounter + 1
            Else
112             mashit = RandomNumber(2, Porcentaje(UserList(UserIndex).Stats.MaxSta, 5)) * UserMod.GetSelfHealingBonus(UserList(UserIndex))
        
114             UserList(UserIndex).Counters.HPCounter = 0
116             Call UserMod.ModifyHealth(UserIndex, mashit)
120             Call WriteConsoleMsg(UserIndex, "Has sanado.", e_FontTypeNames.FONTTYPE_INFO)
122             EnviarStats = True
            End If
        Exit Sub
Sanar_Err:
124     Call TraceError(Err.Number, Err.Description, "General.Sanar", Erl)
End Sub

Public Sub CargaNpcsDat(Optional ByVal ActualizarNPCsExistentes As Boolean = False)
        
            On Error GoTo CargaNpcsDat_Err
        
            ' Leemos el NPCs.dat y lo almacenamos en la memoria.
100         Set LeerNPCs = New clsIniManager
102         Call LeerNPCs.Initialize(DatPath & "NPCs.dat")
        
            ' Cargamos la lista de NPC's hostiles disponibles para spawnear.
104         Call CargarSpawnList
    
            ' Actualizamos la informacion de los NPC's ya spawneados.
106         If ActualizarNPCsExistentes Then
    
                Dim i As Long
108             For i = 1 To NumNPCs
    
110                 If NpcList(i).flags.NPCActive Then
112                     Call OpenNPC(CInt(i), False, True)
                    End If
    
114                 DoEvents
    
116             Next i
    
            End If
        
            Exit Sub

CargaNpcsDat_Err:
118         Call TraceError(Err.Number, Err.Description, "General.CargaNpcsDat", Erl)

        
End Sub

Sub PasarSegundo()

        On Error GoTo ErrHandler

        Dim i    As Long

        Dim h    As Byte

        Dim Mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        If TiempoPesca > 0 Then TiempoPesca = TiempoPesca + 1
        
100     If CuentaRegresivaTimer > 0 Then
102         If CuentaRegresivaTimer > 1 Then
104             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(CuentaRegresivaTimer - 1 & " segundos...!", e_FontTypeNames.FONTTYPE_GUILD))
            Else
106             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Ya!!!", e_FontTypeNames.FONTTYPE_FIGHT))

            End If

108         CuentaRegresivaTimer = CuentaRegresivaTimer - 1

        End If
        
        
        If Not InstanciaCaptura Is Nothing Then
            Call InstanciaCaptura.PasarSegundo
        End If
        
        Segundos = Segundos + 1
        
   
110     For i = 1 To LastUser

112         With UserList(i)

114             If .flags.UserLogged Then
                    
                    If .Counters.timeChat > 0 Then
                        .Counters.timeChat = .Counters.timeChat - 1
                    End If
                    
                    If .Counters.LastTrabajo > 0 Then
                        .Counters.LastTrabajo = .Counters.LastTrabajo - 1
                    End If
                    
                    If .Counters.timeFx > 0 Then
                        .Counters.timeFx = .Counters.timeFx - 1
                    End If
                    
                    If .Counters.timeGuildChat > 0 Then
                        .Counters.timeGuildChat = .Counters.timeGuildChat - 1
                    End If
                                      
116                 If .flags.Silenciado = 1 Then
118                     .flags.SegundosPasados = .flags.SegundosPasados + 1
        
120                     If .flags.SegundosPasados = 60 Then
122                         .flags.MinutosRestantes = .flags.MinutosRestantes - 1
124                         .flags.SegundosPasados = 0
        
                        End If
                    
126                     If .flags.MinutosRestantes = 0 Then
128                         .flags.SegundosPasados = 0
130                         .flags.Silenciado = 0
132                         .flags.MinutosRestantes = 0
134                         Call WriteConsoleMsg(i, "Has sido liberado del silencio.", e_FontTypeNames.FONTTYPE_SERVER)
        
                        End If
        
                    End If

136                 If .flags.Muerto = 0 Then
138                     Call DuracionPociones(i)
142                     If .flags.invisible = 1 Or .Counters.DisabledInvisibility > 0 Then Call EfectoInvisibilidad(i)
144                     If .flags.Paralizado = 1 Then Call EfectoParalisisUser(i)
146                     If .flags.Inmovilizado = 1 Then Call EfectoInmoUser(i)
148                     If .flags.Ceguera = 1 Then Call EfectoCeguera(i)
150                     If .flags.Estupidez = 1 Then Call EfectoEstupidez(i)
152                     If .flags.Maldicion = 1 Then Call EfectoMaldicionUser(i)
154                     If .flags.VelocidadHechizada > 0 Then Call EfectoVelocidadUser(i)
    
156                     If HambreYSed(i) Then
158                         Call WriteUpdateHungerAndThirst(i)
                        End If
                    
                    Else
160                     If .flags.Traveling <> 0 Then Call TravelingEffect(i)
                    End If
        
162                 If .Counters.TimerBarra > 0 Then
164                     .Counters.TimerBarra = .Counters.TimerBarra - 1
                        
166                     If .Counters.TimerBarra = 0 Then
        
168                         Select Case .Accion.TipoAccion
                                Case e_AccionBarra.Hogar
170                                 Call HomeArrival(i)
                            End Select
                            
182                         .Accion.Particula = 0
184                         .Accion.TipoAccion = e_AccionBarra.CancelarAccion
186                         .Accion.HechizoPendiente = 0
188                         .Accion.RunaObj = 0
190                         .Accion.ObjSlot = 0
192                         .Accion.AccionPendiente = False
                            
                        End If
                    End If
        
194                 If .flags.UltimoMensaje > 0 Then
196                     .Counters.RepetirMensaje = .Counters.RepetirMensaje + 1
198                     If .Counters.RepetirMensaje >= 3 Then
200                         .flags.UltimoMensaje = 0
202                         .Counters.RepetirMensaje = 0
                        End If
                    End If
                    
204                 If .Counters.CuentaRegresiva >= 0 Then
206                     If .Counters.CuentaRegresiva > 0 Then
208                         Call WriteConsoleMsg(i, ">>>  " & .Counters.CuentaRegresiva & "  <<<", e_FontTypeNames.FONTTYPE_New_Gris)
                        Else
210                         Call WriteConsoleMsg(i, ">>> YA! <<<", e_FontTypeNames.FONTTYPE_FIGHT)
212                         Call WriteStopped(i, False)
                        End If
                        
214                     .Counters.CuentaRegresiva = .Counters.CuentaRegresiva - 1
                    End If
    
216                 If .flags.Portal > 1 Then
218                     .flags.Portal = .flags.Portal - 1
                
220                     If .flags.Portal = 1 Then
222                         Mapa = .flags.PortalM
224                         X = .flags.PortalX
226                         Y = .flags.PortalY
228                         Call SendData(SendTarget.toMap, .flags.PortalM, PrepareMessageParticleFXToFloor(X, Y, e_ParticulasIndex.TpVerde, 0))
230                         Call SendData(SendTarget.toMap, .flags.PortalM, PrepareMessageLightFXToFloor(X, Y, 0, 105))
        
232                         If MapData(Mapa, X, Y).TileExit.Map > 0 Then
234                             MapData(Mapa, X, Y).TileExit.Map = 0
236                             MapData(Mapa, X, Y).TileExit.X = 0
238                             MapData(Mapa, X, Y).TileExit.Y = 0
        
                            End If
        
240                         MapData(Mapa, X, Y).Particula = 0
242                         MapData(Mapa, X, Y).TimeParticula = 0
244                         MapData(Mapa, X, Y).Particula = 0
246                         MapData(Mapa, X, Y).TimeParticula = 0
248                         .flags.Portal = 0
250                         .flags.PortalM = 0
252                         .flags.PortalY = 0
254                         .flags.PortalX = 0
256                         .flags.PortalMDestino = 0
258                         .flags.PortalYDestino = 0
260                         .flags.PortalXDestino = 0
        
                        End If
        
                    End If
                
262                 If .Counters.EnCombate > 0 Then
264                     .Counters.EnCombate = .Counters.EnCombate - 1
                    End If
                
                    If .Counters.TiempoDeInmunidadParalisisNoMagicas > 0 Then
                        .Counters.TiempoDeInmunidadParalisisNoMagicas = .Counters.TiempoDeInmunidadParalisisNoMagicas - 1
                    End If
                

                
266                 If .Counters.TiempoDeInmunidad > 0 Then
268                     .Counters.TiempoDeInmunidad = .Counters.TiempoDeInmunidad - 1
270                     If .Counters.TiempoDeInmunidad = 0 Then
272                         .flags.Inmunidad = 0
                        End If
                    End If
                
274                 If .flags.Subastando Then
276                     .Counters.TiempoParaSubastar = .Counters.TiempoParaSubastar - 1
        
278                     If .Counters.TiempoParaSubastar = 0 Then
280                         Call CancelarSubasta
                        End If
                    End If
        
                    'Cerrar usuario
282                 If .Counters.Saliendo Then
                        '  If .flags.Muerto = 1 Then .Counters.Salir = 0
284                     .Counters.Salir = .Counters.Salir - 1
                        ' Call WriteConsoleMsg(i, "Se saldrá del juego en " & .Counters.Salir & " segundos...", e_FontTypeNames.FONTTYPE_INFO)
286                     Call WriteLocaleMsg(i, "203", e_FontTypeNames.FONTTYPE_INFO, .Counters.Salir)
        
288                     If .Counters.Salir <= 0 Then
290                         Call WriteConsoleMsg(i, "Gracias por jugar Argentum 20.", e_FontTypeNames.FONTTYPE_INFO)
292                         Call WriteDisconnect(i)
                            
294                         Call CloseSocket(i)
        
                        End If
        
                    End If

                End If ' If UserLogged

            End With
302     Next i

        ' **********************************
        ' **********  Invasiones  **********
        ' **********************************
304     For i = 1 To UBound(Invasiones)
306         With Invasiones(i)

                ' Si la invasión está activa
308             If .Activa Then
310                 .TimerSpawn = .TimerSpawn + 1

                    ' Comprobamos si hay que spawnear NPCs
312                 If .TimerSpawn >= .IntervaloSpawn Then
314                     Call InvasionSpawnNPC(i)
316                     .TimerSpawn = 0
                    End If
                    
                    ' ------------------------------------
                    
318                 .TimerMostrarInfo = .TimerMostrarInfo + 1
                    
                    ' Comprobamos si hay que mostrar la info
320                 If .TimerMostrarInfo >= 5 Then
322                     Call EnviarInfoInvasion(i)
324                     .TimerMostrarInfo = 0
                    End If
                End If
            
            End With
        Next
        ' **********************************

        Exit Sub

ErrHandler:
326     Call TraceError(Err.Number, Err.Description, "General.PasarSegundo", Erl)

End Sub


Sub ForzarActualizar()
    
        On Error Resume Next
    
        Dim i As Long

100     For i = 1 To LastUser

102         If UserList(i).ConnectionDetails.ConnIDValida Then
        
104             Call WriteForceUpdate(i)
    
            End If
    
106     Next i
    
End Sub
 
Sub GuardarUsuarios()
        
        On Error GoTo GuardarUsuarios_Err
        
100     haciendoBK = True
    
102     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())
104     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Grabando Personajes", e_FontTypeNames.FONTTYPE_SERVER))
    
        Dim i As Long
        
106     For i = 1 To LastUser

108         If UserList(i).flags.UserLogged Then
110             Call modNetwork.Poll
            End If

112     Next i

114     For i = 1 To LastUser

116         If UserList(i).flags.UserLogged Then

118              Call SaveUser(i)

            End If

120     Next i
    
122     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor » Personajes Grabados", e_FontTypeNames.FONTTYPE_SERVER))
124     Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

126     haciendoBK = False

        
        Exit Sub

GuardarUsuarios_Err:
128     Call TraceError(Err.Number, Err.Description, "General.GuardarUsuarios", Erl)

        
End Sub

Public Sub FreeNPCs()
        
        On Error GoTo FreeNPCs_Err
        

        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all NPC Indexes
        '***************************************************
        Dim LoopC As Long
    
        ' Free all NPC indexes
100     For LoopC = 1 To MaxNPCs
            Call ReleaseNpc(LoopC, e_DeleteSource.eReleaseAll)
104     Next LoopC

        
        Exit Sub

FreeNPCs_Err:
106     Call TraceError(Err.Number, Err.Description, "General.FreeNPCs", Erl)

        
End Sub

Public Sub FreeCharIndexes()
        '***************************************************
        'Autor: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/06
        'Releases all char indexes
        '***************************************************
        ' Free all char indexes (set them all to 0)
        
        On Error GoTo FreeCharIndexes_Err
        
100     Call ZeroMemory(CharList(1), MAXCHARS * Len(CharList(1)))

        
        Exit Sub

FreeCharIndexes_Err:
102     Call TraceError(Err.Number, Err.Description, "General.FreeCharIndexes", Erl)

        
End Sub

Function RandomString(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String
        
        On Error GoTo RandomString_Err
        

100     Randomize Time

        Dim rgch As String

102     rgch = "abcdefghijklmnopqrstuvwxyz"
    
104     If OnlyUpper Then
106         rgch = UCase(rgch)
        Else
108         rgch = rgch & UCase(rgch)

        End If
    
110     rgch = rgch & "0123456789"  ' & "#@!~$()-_"

        Dim i As Long

112     For i = 1 To cb
114         RandomString = RandomString & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Next

        
        Exit Function

RandomString_Err:
116     Call TraceError(Err.Number, Err.Description, "General.RandomString", Erl)

        
End Function


Function RandomName(cb As Integer, Optional ByVal OnlyUpper As Boolean = False) As String
        
        On Error GoTo RandomString_Err
        

100     Randomize Time

        Dim rgch As String

102     rgch = "abcdefghijklmnopqrstuvwxyz"
    
104     If OnlyUpper Then
106         rgch = UCase$(rgch)
        Else
108         rgch = rgch & UCase$(rgch)
        End If
    
        Dim i As Long

112     For i = 1 To cb
114         RandomName = RandomName & mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
        Next

        
        Exit Function

RandomString_Err:
116     Call TraceError(Err.Number, Err.Description, "General.RandomString", Erl)

        
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean

        On Error GoTo errHnd

        Dim lPos As Long

        Dim lX   As Long

        Dim iAsc As Integer
    
        '1er test: Busca un simbolo @
100     lPos = InStr(sString, "@")

102     If (lPos <> 0) Then

            '2do test: Busca un simbolo . después de @ + 1
104         If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then Exit Function
        
            '3er test: Recorre todos los caracteres y los valída
106         For lX = 0 To Len(sString) - 1

108             If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
110                 iAsc = Asc(mid$(sString, (lX + 1), 1))

112                 If Not CMSValidateChar_(iAsc) Then Exit Function

                End If

114         Next lX
        
            'Finale
116         CheckMailString = True

        End If

errHnd:

End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
        
        On Error GoTo CMSValidateChar__Err
        
100     CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or (iAsc >= 65 And iAsc <= 90) Or (iAsc >= 97 And iAsc <= 122) Or (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)

        
        Exit Function

CMSValidateChar__Err:
102     Call TraceError(Err.Number, Err.Description, "General.CMSValidateChar_", Erl)

        
End Function

Public Function Tilde(ByRef Data As String) As String
    
        On Error GoTo Tilde_Err
    

100     Tilde = UCase$(Data)
 
102     Tilde = Replace$(Tilde, "Á", "A")
104     Tilde = Replace$(Tilde, "É", "E")
106     Tilde = Replace$(Tilde, "Í", "I")
108     Tilde = Replace$(Tilde, "Ó", "O")
110     Tilde = Replace$(Tilde, "Ú", "U")
        
    
        Exit Function

Tilde_Err:
112     Call TraceError(Err.Number, Err.Description, "Mod_General.Tilde", Erl)

    
End Function

Public Sub CerrarServidor()
        'Save stats!!!
102     Call frmMain.QuitarIconoSystray
    
        ' Limpieza del socket del servidor.
104     Call modNetwork.Disconnect
    
        Dim LoopC As Long
106     For LoopC = 1 To MaxUsers
108         If UserList(LoopC).ConnectionDetails.ConnIDValida Then
110             Call CloseSocket(LoopC)
            End If
        Next
        Call UnloadAntiCheat
112     If Database_Enabled Then Database_Close
 
124     End
   
End Sub


Public Function PonerPuntos(ByVal Numero As Long) As String
    
        On Error GoTo PonerPuntos_Err
    

        Dim i     As Integer

        Dim Cifra As String
 
100     Cifra = str$(Numero)
102     Cifra = Right$(Cifra, Len(Cifra) - 1)

104     For i = 0 To 4

106         If Len(Cifra) - 3 * i >= 3 Then
108             If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
110                 PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos

                End If

            Else

112             If Len(Cifra) - 3 * i > 0 Then
114                 PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos

                End If

                Exit For

            End If

        Next
 
116     PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
    
        Exit Function

PonerPuntos_Err:
118     Call TraceError(Err.Number, Err.Description, "ModLadder.PonerPuntos", Erl)

    
End Function

' Autor: WyroX
Function CalcularPromedioVida(ByVal UserIndex As Integer) As Double

100     With UserList(UserIndex)
102         If .Stats.ELV = 1 Then
                ' Siempre estamos promedio al lvl 1
104             CalcularPromedioVida = ModClase(.clase).Vida - (21 - .Stats.UserAtributos(e_Atributos.Constitucion)) * 0.5
            Else
106             CalcularPromedioVida = (.Stats.MaxHp - .Stats.UserAtributos(e_Atributos.Constitucion)) / (.Stats.ELV - 1)
            End If
        End With

End Function

' Adaptado desde https://stackoverflow.com/questions/29325069/how-to-generate-random-numbers-biased-towards-one-value-in-a-range/29325222#29325222
' By WyroX
Function RandomIntBiased(ByVal Min As Double, ByVal max As Double, ByVal Bias As Double, ByVal Influence As Double) As Double

        On Error GoTo handle

        Dim RandomRango As Double, Mix As Double
    
100     RandomRango = Rnd * (max - Min) + Min
102     Mix = Rnd * Influence
    
104     RandomIntBiased = RandomRango * (1 - Mix) + Bias * Mix
    
        Exit Function
    
handle:
106     Call TraceError(Err.Number, Err.Description, "General.RandomIntBiased")
108     RandomIntBiased = Bias

End Function

'Very efficient function for testing whether this code is running in the IDE or compiled
'https://www.vbforums.com/showthread.php?231468-VB-Detect-if-you-are-running-in-the-IDE&p=5413357&viewfull=1#post5413357
Public Function RunningInVB(Optional ByRef b As Boolean = True) As Boolean
100     If b Then Debug.Assert Not RunningInVB(RunningInVB) Else b = True
End Function

' WyroX: Mensaje a todo el mundo
Public Sub MensajeGlobal(texto As String, Fuente As e_FontTypeNames)
100     Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(texto, Fuente))
End Sub

' WyroX: Devuelve si X e Y están dentro del Rectangle
Public Function InsideRectangle(R As t_Rectangle, ByVal X As Integer, ByVal Y As Integer) As Boolean
100     If X < R.X1 Then Exit Function
102     If X > R.X2 Then Exit Function
104     If Y < R.Y1 Then Exit Function
106     If Y > R.Y2 Then Exit Function
108     InsideRectangle = True
End Function

' Based on: https://stackoverflow.com/questions/1378604/end-process-from-task-manager-using-vb-6-code (ultima respuesta)
Public Function GetProcessCount(ByVal processName As String) As Byte
    
        Dim oService As Object
        Dim servicename As String
        Dim processCount As Byte
    
100     Dim oWMI As Object: Set oWMI = GetObject("winmgmts:")
102     Dim oServices As Object: Set oServices = oWMI.InstancesOf("win32_process")

104     For Each oService In oServices

106         servicename = CStr(oService.name)

108         If StrComp(servicename, processName, vbTextCompare) = 0 Then
            
                ' Para matar un proceso adentro de este loop usar.
                ' oService.Terminate
            
110             processCount = processCount + 1
            
            End If

        Next
    
112     GetProcessCount = processCount

End Function

Public Function EsMapaInterdimensional(ByVal Map As Integer) As Boolean
        Dim i As Integer
100     For i = 1 To UBound(MapasInterdimensionales)
102         If Map = MapasInterdimensionales(i) Then
104             EsMapaInterdimensional = True
                Exit Function
            End If
        Next
End Function

Public Function IsValidIPAddress(ByVal IP As String) As Boolean

        On Error GoTo Handler

        Dim varAddress As Variant, n As Long, lCount As Long
100     varAddress = Split(IP, ".", 4, vbTextCompare)

102     If IsArray(varAddress) Then

104         For n = LBound(varAddress) To UBound(varAddress)
106             lCount = lCount + 1
108             varAddress(n) = CByte(varAddress(n))
            Next
        
110         IsValidIPAddress = (lCount = 4)

        End If

Handler:

End Function

Function Ceil(X As Variant) As Variant
        
        On Error GoTo Ceil_Err
        
100     Ceil = IIf(Fix(X) = X, X, Fix(X) + 1)
        
        Exit Function

Ceil_Err:
102     Call TraceError(Err.Number, Err.Description & "Ceil_Err", Erl)


        
End Function

Function Clamp(X As Variant, a As Variant, b As Variant) As Variant
        
        On Error GoTo Clamp_Err
        
100     Clamp = IIf(X < a, a, IIf(X > b, b, X))
        
        Exit Function

Clamp_Err:
102     Call TraceError(Err.Number, Err.Description & "Clamp_Err", Erl)


        
End Function

Private Function GetElapsed() As Single
    Static sTime1     As Currency
    Static sTime2     As Currency
    Static sFrequency As Currency
    
    'Get the timer frequency
    If sFrequency = 0 Then
        Call QueryPerformanceFrequency(sFrequency)
    End If
    
    'Get current time
    Call QueryPerformanceCounter(sTime1)

     'Calculate elapsed time
    GetElapsed = ((sTime1 - sTime2) / sFrequency * 1000)
    
    'Get next end time
    Call QueryPerformanceCounter(sTime2)
End Function

Public Function RunScriptInFile(ByVal FilePath As String) As Boolean
    Dim Script As String
    Script = FileText(FilePath)
    Script = Replace(Replace(Script, Chr(10), ""), Chr(13), "")
    Dim RS As Recordset
    If Script <> vbNullString Then
        Set RS = Query(Script)
        If RS Is Nothing Then
            RunScriptInFile = False
            Exit Function
        End If
    End If
    RunScriptInFile = True
End Function
'Reads the files inside the ScriptsDB folder, it can be a create table, alter, etc.
'we are calling this files dbmigrations, this function check this
'folder and the db, and run all the files that are not registered in the db migration table
'the file should store the name in the format of YYYYMMDD-XX-description text.sql
'where the XX is the number of migrations generated the same day
Public Sub LoadDBMigrations()
   
        On Error GoTo LoadDBMigrations_Err
    'Consulto a la DB a ver si existe la tabla migrations
    
    Dim RS As Recordset
    Set RS = Query("select * from migrations")
    Dim LastScript As String: LastScript = ""
    
    If RS Is Nothing Then
        Call Query("CREATE TABLE ""migrations"" (    ""id"" INTEGER NOT NULL,    ""date"" VARCHAR(11) NOT NULL,    ""description"" VARCHAR(50) NULL,    Primary key(""id""));")
    Else
        Set RS = Query("select date from migrations order by id desc LIMIT 1;")
        If RS.RecordCount > 0 Then LastScript = RS!Date
    End If
    
    Dim sFilename  As String
    sFilename = dir(App.Path & "/ScriptsDB/")
        
    Do While sFilename <> ""
        If Len(sFilename) > 11 Then
            Dim date_ As String
            date_ = Left(sFilename, 11)
            If LastScript < date_ Then
                'Leemos el archivo
                Dim script As String
                Dim Description As String
                Description = mid(sFilename, 13, Len(sFilename) - 16)
                If RunScriptInFile(App.Path & "/ScriptsDB/" & sFilename) Then
                    Call Query("insert into migrations (date, description) values (?,?);", Date_, Description)
                Else
                    Call Err.raise(5, , "invalid - " & Description)
                End If
            End If
        End If
        sFilename = dir()
    Loop
        
    Exit Sub

LoadDBMigrations_Err:
122     Call TraceError(err.Number, err.Description, "modGuilds.LoadDBMigrations", Erl)
        Call MsgBox(DBError & vbNewLine & "Script:" & err.Description, vbCritical, "ERROR MIGRATIONS")
        
End Sub
Function FileText(filename$) As String
    Dim handle As Integer
    handle = FreeFile
    Open filename$ For Input As #handle
    FileText = Input$(LOF(handle), handle)
    Close #handle
End Function

Public Function IsArrayInitialized(ByRef arr) As Boolean
  Dim rv As Long
  On Error Resume Next
  rv = UBound(arr)
  IsArrayInitialized = (Err.Number = 0) And rv >= 0
End Function
