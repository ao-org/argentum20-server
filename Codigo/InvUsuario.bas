Attribute VB_Name = "InvUsuario"
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

Public Function TieneObjEnInv(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional ObjIndex2 As Integer = 0) As Boolean
        On Error GoTo TieneObjEnInv_Err
        
        'Devuelve el slot del inventario donde se encuentra el obj
        'Creaado por Ladder 25/09/2014
        Dim i As Byte

     For i = 1 To 36

         If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex Then
             TieneObjEnInv = True
                Exit Function

            End If

         If ObjIndex2 > 0 Then
             If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex2 Then
                 TieneObjEnInv = True
                    Exit Function

                End If

            End If

     Next i

     TieneObjEnInv = False

        
        Exit Function

TieneObjEnInv_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.TieneObjEnInv", Erl)

        
End Function


Public Function CantidadObjEnInv(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Integer
        On Error GoTo CantidadObjEnInv_Err
        'Devuelve el amount si tiene el ObjIndex en el inventario, sino devuelve 0
        'Creaado por Ladder 25/09/2014
        Dim i As Byte

     For i = 1 To 36

         If UserList(UserIndex).Invent.Object(i).ObjIndex = ObjIndex Then
             CantidadObjEnInv = UserList(UserIndex).Invent.Object(i).amount
                Exit Function
            End If


     Next i

     CantidadObjEnInv = 0

        
        Exit Function

CantidadObjEnInv_Err:
     Call TraceError(Err.Number, Err.Description, "ModLadder.CantidadObjEnInv", Erl)

        
End Function


Public Function TieneObjetosRobables(ByVal UserIndex As Integer) As Boolean
        On Error GoTo TieneObjetosRobables_Err
      

        Dim i        As Integer
        Dim ObjIndex As Integer
        
        If UserList(UserIndex).CurrentInventorySlots > 0 Then
100         For i = 1 To UserList(UserIndex).CurrentInventorySlots
102             ObjIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
    
104             If ObjIndex > 0 Then
106                 If (ObjData(ObjIndex).OBJType <> e_OBJType.otLlaves And ObjData(ObjIndex).OBJType <> e_OBJType.otBarcos And ObjData(ObjIndex).OBJType <> e_OBJType.otMonturas And ObjData(ObjIndex).OBJType <> e_OBJType.OtDonador And ObjData(ObjIndex).OBJType <> e_OBJType.otRunas) Then
108                     TieneObjetosRobables = True
                        Exit Function
    
                    End If
        
                End If
    
110         Next i
        End If

        
        Exit Function

TieneObjetosRobables_Err:
112     Call TraceError(Err.Number, Err.Description, "InvUsuario.TieneObjetosRobables", Erl)

        
End Function

Function ClasePuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean

        On Error GoTo manejador

        Dim flag As Boolean

100     If Slot <> 0 Then
102         If UserList(UserIndex).Invent.Object(Slot).Equipped Then
104             ClasePuedeUsarItem = True
                Exit Function

            End If

        End If

106     If EsGM(UserIndex) Then
108         ClasePuedeUsarItem = True
            Exit Function

        End If

        Dim i As Integer

110     For i = 1 To NUMCLASES

112         If ObjData(ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
114             ClasePuedeUsarItem = False
                Exit Function

            End If

116     Next i

118     ClasePuedeUsarItem = True

        Exit Function

manejador:
120     LogError ("Error en ClasePuedeUsarItem")

End Function

Function RazaPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer, Optional Slot As Byte) As Boolean
        On Error GoTo RazaPuedeUsarItem_Err

        Dim Objeto As t_ObjData, i As Long
        
100     Objeto = ObjData(ObjIndex)
        
102     If EsGM(UserIndex) Then
104         RazaPuedeUsarItem = True
            Exit Function
        End If

106     For i = 1 To NUMRAZAS
108         If Objeto.RazaProhibida(i) = UserList(UserIndex).raza Then
110             RazaPuedeUsarItem = False
                Exit Function
            End If

112     Next i
        
        ' Si el objeto no define una raza en particular
114     If Objeto.RazaDrow + Objeto.RazaElfa + Objeto.RazaEnana + Objeto.RazaGnoma + Objeto.RazaHumana + Objeto.RazaOrca = 0 Then
116         RazaPuedeUsarItem = True
        
        Else ' El objeto esta definido para alguna raza en especial
118         Select Case UserList(UserIndex).raza
                Case e_Raza.Humano
120                 RazaPuedeUsarItem = Objeto.RazaHumana > 0

122             Case e_Raza.Elfo
124                 RazaPuedeUsarItem = Objeto.RazaElfa > 0
                
126             Case e_Raza.Drow
128                 RazaPuedeUsarItem = Objeto.RazaDrow > 0
    
130             Case e_Raza.Orco
132                 RazaPuedeUsarItem = Objeto.RazaOrca > 0
                    
134             Case e_Raza.Gnomo
136                 RazaPuedeUsarItem = Objeto.RazaGnoma > 0
                    
138             Case e_Raza.Enano
140                 RazaPuedeUsarItem = Objeto.RazaEnana > 0

            End Select
        End If
        
        If RazaPuedeUsarItem And Objeto.OBJType = e_OBJType.otArmadura Then
            RazaPuedeUsarItem = ObtenerRopaje(UserIndex, Objeto) <> 0
        End If
        
        Exit Function

RazaPuedeUsarItem_Err:
142     LogError ("Error en RazaPuedeUsarItem")

End Function

Sub QuitarNewbieObj(ByVal UserIndex As Integer)
        
        On Error GoTo QuitarNewbieObj_Err
        

        Dim j As Integer
        
        If UserList(UserIndex).CurrentInventorySlots > 0 Then
100         For j = 1 To UserList(UserIndex).CurrentInventorySlots
    
102             If UserList(UserIndex).Invent.Object(j).ObjIndex > 0 Then
                 
104                 If ObjData(UserList(UserIndex).Invent.Object(j).ObjIndex).Newbie = 1 Then
106                     Call QuitarUserInvItem(UserIndex, j, MAX_INVENTORY_OBJS)
108                     Call UpdateUserInv(False, UserIndex, j)
    
                    End If
            
                End If
    
110         Next j
        End If
    
        'Si el usuario dejó de ser Newbie, y estaba en el Newbie Dungeon

112     If MapInfo(UserList(UserIndex).pos.map).Newbie Then
                   
            'Mandamos a la isla de renacimiento
            Call WarpUserChar(UserIndex, Renacimiento.map, Renacimiento.X, Renacimiento.y, True)
            Call WriteConsoleMsg(UserIndex, "Has dejado de ser Newbie, Te orientaremos que hacer ahora.", e_FontTypeNames.FONTTYPE_INFO)
            
    
        End If

        
        Exit Sub

QuitarNewbieObj_Err:
144     Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarNewbieObj", Erl)

        
End Sub

Sub LimpiarInventario(ByVal UserIndex As Integer)
        
        On Error GoTo LimpiarInventario_Err
        

        Dim j As Integer

        If UserList(UserIndex).CurrentInventorySlots > 0 Then
100         For j = 1 To UserList(UserIndex).CurrentInventorySlots
102             UserList(UserIndex).Invent.Object(j).ObjIndex = 0
104             UserList(UserIndex).Invent.Object(j).amount = 0
106             UserList(UserIndex).Invent.Object(j).Equipped = 0
            Next
        End If

108     UserList(UserIndex).Invent.NroItems = 0

110     UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
112     UserList(UserIndex).Invent.ArmourEqpSlot = 0

114     UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
116     UserList(UserIndex).Invent.WeaponEqpSlot = 0

118     UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0
120     UserList(UserIndex).Invent.HerramientaEqpSlot = 0

122     UserList(UserIndex).Invent.CascoEqpObjIndex = 0
124     UserList(UserIndex).Invent.CascoEqpSlot = 0

126     UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
128     UserList(UserIndex).Invent.EscudoEqpSlot = 0

130     UserList(UserIndex).invent.DañoMagicoEqpObjIndex = 0
132     UserList(UserIndex).invent.DañoMagicoEqpSlot = 0

134     UserList(UserIndex).Invent.ResistenciaEqpObjIndex = 0
136     UserList(UserIndex).Invent.ResistenciaEqpSlot = 0

142     UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
144     UserList(UserIndex).Invent.MunicionEqpSlot = 0

146     UserList(UserIndex).Invent.BarcoObjIndex = 0
148     UserList(UserIndex).Invent.BarcoSlot = 0

150     UserList(UserIndex).Invent.MonturaObjIndex = 0
152     UserList(UserIndex).Invent.MonturaSlot = 0

154     UserList(UserIndex).Invent.MagicoObjIndex = 0
156     UserList(UserIndex).Invent.MagicoSlot = 0

        
        Exit Sub

LimpiarInventario_Err:
158     Call TraceError(Err.Number, Err.Description, "InvUsuario.LimpiarInventario", Erl)

        
End Sub

Sub TirarOro(ByVal Cantidad As Long, ByVal UserIndex As Integer)

        '***************************************************
        'Autor: Unknown (orginal version)
        'Last Modification: 23/01/2007
        '23/01/2007 -> Pablo (ToxicWaste): Billetera invertida y explotar oro en el agua.
        '***************************************************
        On Error GoTo ErrHandler
        Dim OriginalAmount As Long
        OriginalAmount = Cantidad
100     With UserList(UserIndex)
        
            ' GM's (excepto Dioses y Admins) no pueden tirar oro
102         If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin Or e_PlayerType.Dios)) = 0 Then
104             Call LogGM(.Name, " trató de tirar " & PonerPuntos(Cantidad) & " de oro en " & .Pos.Map & "-" & .Pos.X & "-" & .Pos.Y)
                Exit Sub
            End If
         
            ' Si el usuario tiene ORO, entonces lo tiramos
106         If (Cantidad > 0) And (Cantidad <= .Stats.GLD) Then

                Dim i     As Byte
                Dim MiObj As t_Obj

                'info debug
                Dim loops As Long
                
116             Do While (Cantidad > 0)
            
118                 If Cantidad > MAX_INVENTORY_OBJS And .Stats.GLD > MAX_INVENTORY_OBJS Then
120                     MiObj.amount = MAX_INVENTORY_OBJS
122                     Cantidad = Cantidad - MiObj.amount
                    Else
124                     MiObj.amount = Cantidad
126                     Cantidad = Cantidad - MiObj.amount

                    End If

128                 MiObj.ObjIndex = iORO

                    Dim AuxPos As t_WorldPos

130                 If .clase = e_Class.Pirat Then
132                     AuxPos = TirarItemAlPiso(.Pos, MiObj, False)
                    Else
134                     AuxPos = TirarItemAlPiso(.Pos, MiObj, True)
                    End If
            
136                 If AuxPos.X <> 0 And AuxPos.Y <> 0 Then
138                     .Stats.GLD = .Stats.GLD - MiObj.amount

                    End If
            
                    'info debug
140                 loops = loops + 1

142                 If loops > 100000 Then 'si entra aca y se cuelga mal el server revisen al tipo porque tiene much oro (NachoP) seguramente es dupero
144                     Call LogError("Se ha superado el limite de iteraciones(100000) permitido en el Sub TirarOro() - posible Nacho P")
                        Exit Sub

                    End If
            
                Loop
                
                ' Si es GM, registramos lo q hizo
146             If EsGM(UserIndex) Then
148                 If MiObj.ObjIndex = iORO Then
150                     Call LogGM(.name, "Tiro: " & PonerPuntos(OriginalAmount) & " monedas de oro.")
                    Else
152                     Call LogGM(.name, "Tiro cantidad:" & PonerPuntos(OriginalAmount) & " Objeto:" & ObjData(MiObj.ObjIndex).name)
                    End If
                End If
    
160             Call WriteUpdateGold(UserIndex)

            End If
        
        End With

        Exit Sub

ErrHandler:
162 Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarOro", Erl())
    
End Sub

Public Sub QuitarUserInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Cantidad As Integer)
        
        On Error GoTo QuitarUserInvItem_Err
        

100     If Slot < 1 Or Slot > UserList(UserIndex).CurrentInventorySlots Then Exit Sub
    
102     With UserList(UserIndex).Invent.Object(Slot)

104         If .amount <= Cantidad And .Equipped = 1 Then
106             Call Desequipar(UserIndex, Slot)
            End If
        
            'Quita un objeto
108         .amount = .amount - Cantidad

            '¿Quedan mas?
110         If .amount <= 0 Then
112             UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
114             .ObjIndex = 0
116             .amount = 0
            End If
            
            UserList(UserIndex).flags.ModificoInventario = True
            
        End With
        
        If IsValidUserRef(UserList(UserIndex).flags.GMMeSigue) And UserIndex <> UserList(UserIndex).flags.GMMeSigue.ArrayIndex Then
            Call QuitarUserInvItem(UserList(UserIndex).flags.GMMeSigue.ArrayIndex, Slot, Cantidad)
        End If

        
        Exit Sub

QuitarUserInvItem_Err:
118     Call TraceError(Err.Number, Err.Description, "InvUsuario.QuitarUserInvItem", Erl)

        
End Sub

Public Sub UpdateUserInv(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
        
        On Error GoTo UpdateUserInv_Err
        

        Dim NullObj As t_UserOBJ

        Dim LoopC   As Byte

        'Actualiza un solo slot
100     If Not UpdateAll And Slot > 0 Then
    
            'Actualiza el inventario
102         If UserList(UserIndex).Invent.Object(Slot).ObjIndex > 0 Then
104             Call ChangeUserInv(UserIndex, Slot, UserList(UserIndex).Invent.Object(Slot))
            Else
106             Call ChangeUserInv(UserIndex, Slot, NullObj)
            End If
                        
            UserList(UserIndex).flags.ModificoInventario = True
        Else

            'Actualiza todos los slots
            If UserList(UserIndex).CurrentInventorySlots > 0 Then
108             For LoopC = 1 To UserList(UserIndex).CurrentInventorySlots
                    'Actualiza el inventario
110                 If UserList(UserIndex).Invent.Object(LoopC).ObjIndex > 0 Then
112                     Call ChangeUserInv(UserIndex, LoopC, UserList(UserIndex).Invent.Object(LoopC))
                    Else
114                     Call ChangeUserInv(UserIndex, LoopC, NullObj)
                    End If
116             Next LoopC
            End If

        End If
        

        
        Exit Sub

UpdateUserInv_Err:
118     Call TraceError(Err.Number, Err.Description, "InvUsuario.UpdateUserInv", Erl)

        
End Sub

Sub DropObj(ByVal UserIndex As Integer, _
            ByVal Slot As Byte, _
            ByVal num As Integer, _
            ByVal Map As Integer, _
            ByVal X As Integer, _
            ByVal Y As Integer)
        
        On Error GoTo DropObj_Err
        Dim obj As t_Obj

100     If num > 0 Then
102         With UserList(UserIndex)
104             If num > .Invent.Object(Slot).amount Then
106                 num = .Invent.Object(Slot).amount
                End If
108             obj.ObjIndex = .Invent.Object(Slot).ObjIndex
110             obj.amount = num
                If Not CustomScenarios.UserCanDropItem(UserIndex, Slot, Map, x, y) Then
                    Exit Sub
                End If
                
112             If ObjData(obj.ObjIndex).Destruye = 0 Then
                    Dim Suma As Long
                    Suma = num + MapData(.Pos.Map, X, Y).ObjInfo.amount
                    'Check objeto en el suelo
114                 If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = 0 Or (MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex And Suma <= MAX_INVENTORY_OBJS) Then
116                     If Suma > MAX_INVENTORY_OBJS Then
118                         num = MAX_INVENTORY_OBJS - MapData(.Pos.Map, X, Y).ObjInfo.amount
                        End If
                        ' Si sos Admin, Dios o Usuario, crea el objeto en el piso.
120                     If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
                            ' Tiramos el item al piso
122                         Call MakeObj(obj, Map, X, Y)
                        End If
                        Call CustomScenarios.UserDropItem(UserIndex, Slot, Map, x, y)
124                     Call QuitarUserInvItem(UserIndex, Slot, num)
126                     Call UpdateUserInv(False, UserIndex, Slot)

                        If .flags.jugando_captura = 1 Then
                            If Not InstanciaCaptura Is Nothing Then
                                Call InstanciaCaptura.tiraBandera(UserIndex, obj.objIndex)
                            End If
                        End If
                        
128                     If Not .flags.Privilegios And e_PlayerType.user Then
                            If (.flags.Privilegios And (e_PlayerType.Admin Or e_PlayerType.Dios)) <> 0 Then
130                             Call LogGM(.Name, "Tiro cantidad:" & num & " Objeto:" & ObjData(obj.ObjIndex).Name)
                            End If
                        End If
                    Else
132                     Call WriteLocaleMsg(UserIndex, "262", e_FontTypeNames.FONTTYPE_INFO)
                    End If
                Else
134                 Call QuitarUserInvItem(UserIndex, Slot, num)
136                 Call UpdateUserInv(False, UserIndex, Slot)
                End If
            End With
        End If
        Exit Sub
DropObj_Err:
138     Call TraceError(Err.Number, Err.Description, "InvUsuario.DropObj", Erl)
End Sub

Sub EraseObj(ByVal num As Integer, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer)
        
        On Error GoTo EraseObj_Err
        

        Dim Rango As Byte

100     MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount - num

102     If MapData(Map, X, Y).ObjInfo.amount <= 0 Then

            
108         MapData(Map, X, Y).ObjInfo.ObjIndex = 0
110         MapData(Map, X, Y).ObjInfo.amount = 0
    
    
112         Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectDelete(X, Y))

        End If

        
        Exit Sub

EraseObj_Err:
114     Call TraceError(Err.Number, Err.Description, "InvUsuario.EraseObj", Erl)

        
End Sub

Sub MakeObj(ByRef obj As t_Obj, ByVal Map As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal Limpiar As Boolean = True)
        
        On Error GoTo MakeObj_Err

        Dim Color As Long

        Dim Rango As Byte

100     If obj.ObjIndex > 0 And obj.ObjIndex <= UBound(ObjData) Then
    
102         If MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex Then
104             MapData(Map, X, Y).ObjInfo.amount = MapData(Map, X, Y).ObjInfo.amount + obj.amount
            Else
110             MapData(Map, X, Y).ObjInfo.ObjIndex = obj.ObjIndex

112             If ObjData(obj.ObjIndex).VidaUtil <> 0 Then
114                 MapData(Map, X, Y).ObjInfo.amount = ObjData(obj.ObjIndex).VidaUtil
                Else
116                 MapData(Map, X, Y).ObjInfo.amount = obj.amount

                End If
                
            End If
            
118         Call modSendData.SendToAreaByPos(Map, X, Y, PrepareMessageObjectCreate(obj.ObjIndex, MapData(Map, X, Y).ObjInfo.amount, X, Y))
    
        End If
        
        Exit Sub

MakeObj_Err:
120     Call TraceError(Err.Number, Err.Description, "InvUsuario.MakeObj", Erl)
End Sub

Function GetSlotForItemInInventory(ByVal UserIndex As Integer, ByRef MyObject As t_Obj) As Integer
On Error GoTo GetSlotForItemInInventory_Err
    GetSlotForItemInInventory = -1
100 Dim i As Integer
    
102 For i = 1 To UserList(UserIndex).CurrentInventorySlots
104    If UserList(UserIndex).invent.Object(i).objIndex = 0 And GetSlotForItemInInventory = -1 Then
106        GetSlotForItemInInventory = i 'we found a valid place but keep looking in case we can stack
108    ElseIf UserList(UserIndex).invent.Object(i).objIndex = MyObject.objIndex And _
              UserList(UserIndex).invent.Object(i).amount + MyObject.amount <= MAX_INVENTORY_OBJS Then
110        GetSlotForItemInInventory = i 'we can stack the item, let use this slot
112        Exit Function
       End If
    Next i
    Exit Function
GetSlotForItemInInventory_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.GetSlotForItemInInventory", Erl)
End Function

Function GetSlotInInventory(ByVal UserIndex As Integer, ByVal objIndex As Integer) As Integer
    On Error GoTo GetSlotInInventory_Err
    GetSlotInInventory = -1
    Dim i As Integer
    
    For i = 1 To UserList(UserIndex).CurrentInventorySlots
        If UserList(UserIndex).invent.Object(i).objIndex = objIndex Then
            GetSlotInInventory = i
            Exit Function
        End If
    Next i
    Exit Function
GetSlotInInventory_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.GetSlotInInventory", Erl)
End Function

Function MeterItemEnInventario(ByVal UserIndex As Integer, ByRef MiObj As t_Obj) As Boolean

        On Error GoTo MeterItemEnInventario_Err

        Dim X    As Integer

        Dim Y    As Integer

        Dim Slot As Integer
        If MiObj.ObjIndex = 12 Then
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + MiObj.amount
            MeterItemEnInventario = True
            Call WriteUpdateGold(UserIndex)
            Exit Function
        End If
        '¿el user ya tiene un objeto del mismo tipo? ?????
100     Slot = GetSlotForItemInInventory(UserIndex, MiObj)
        
        If Slot <= 0 Then
118        Call WriteLocaleMsg(UserIndex, MsgInventoryIsFull, e_FontTypeNames.FONTTYPE_FIGHT)
120        MeterItemEnInventario = False
           Exit Function
        End If
        If UserList(UserIndex).invent.Object(Slot).objIndex = 0 Then
            UserList(UserIndex).invent.NroItems = UserList(UserIndex).invent.NroItems + 1
        End If
        'Mete el objeto
124     If UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
126         UserList(UserIndex).Invent.Object(Slot).ObjIndex = MiObj.ObjIndex
128         UserList(UserIndex).Invent.Object(Slot).amount = UserList(UserIndex).Invent.Object(Slot).amount + MiObj.amount
            
        
        Else
130         UserList(UserIndex).Invent.Object(Slot).amount = MAX_INVENTORY_OBJS
        End If
        
132     Call UpdateUserInv(False, UserIndex, Slot)
        
134     MeterItemEnInventario = True
        UserList(UserIndex).flags.ModificoInventario = True

        Exit Function
MeterItemEnInventario_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.MeterItemEnInventario", Erl)
End Function

Function HayLugarEnInventario(ByVal UserIndex As Integer, ByVal TargetItemIndex As Integer, ByVal ItemCount) As Boolean
On Error GoTo HayLugarEnInventario_err
        Dim X    As Integer
        Dim Y    As Integer
        Dim Slot As Byte
100     Slot = 1

102     Do Until UserList(UserIndex).invent.Object(Slot).ObjIndex = 0 Or _
            (UserList(UserIndex).invent.Object(Slot).ObjIndex = TargetItemIndex And UserList(UserIndex).invent.Object(Slot).amount + ItemCount < 10000)
104         Slot = Slot + 1
106         If Slot > UserList(UserIndex).CurrentInventorySlots Then
108             HayLugarEnInventario = False
                Exit Function
            End If
        Loop
110     HayLugarEnInventario = True
        Exit Function
HayLugarEnInventario_err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.HayLugarEnInventario", Erl)
End Function

Sub PickObj(ByVal UserIndex As Integer)
        
        On Error GoTo PickObj_Err
        
        Dim X    As Integer
        Dim Y    As Integer
        Dim Slot As Byte
        Dim obj   As t_ObjData
        Dim MiObj As t_Obj

        '¿Hay algun obj?
100     If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex > 0 Then

            '¿Esta permitido agarrar este obj?
102         If ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex).Agarrable <> 1 Then

104             If UserList(UserIndex).flags.Montado = 1 Then
106                 Call WriteConsoleMsg(UserIndex, "Debes descender de tu montura para agarrar objetos del suelo.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
                If Not UserCanPickUpItem(UserIndex) Then
                    Exit Sub
                End If
108             X = UserList(UserIndex).Pos.X
110             Y = UserList(UserIndex).Pos.Y

                If UserList(UserIndex).flags.jugando_captura = 1 Then
                    If Not InstanciaCaptura Is Nothing Then
                        If Not InstanciaCaptura.tomaBandera(UserIndex, MapData(UserList(UserIndex).Pos.map, X, Y).ObjInfo.objIndex) Then
                            Exit Sub
                        End If
                    End If
                End If
                
                
112             obj = ObjData(MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).ObjInfo.ObjIndex)
114             MiObj.amount = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.amount
116             MiObj.ObjIndex = MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.ObjIndex
        
118             If Not MeterItemEnInventario(UserIndex, MiObj) Then
                    'Call WriteConsoleMsg(UserIndex, "No puedo cargar mas objetos.", e_FontTypeNames.FONTTYPE_INFO)
                Else
            
                    'Quitamos el objeto
120                 Call EraseObj(MapData(UserList(UserIndex).Pos.Map, X, Y).ObjInfo.amount, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)

122                 If Not UserList(UserIndex).flags.Privilegios And e_PlayerType.user Then Call LogGM(UserList(UserIndex).Name, "Agarro:" & MiObj.amount & " Objeto:" & ObjData(MiObj.ObjIndex).Name)
                    
                    Call UserDidPickupItem(UserIndex, MiObj.ObjIndex)
                    If UserList(UserIndex).flags.jugando_captura = 1 Then
                    If Not InstanciaCaptura Is Nothing Then
                            Call InstanciaCaptura.quitarBandera(UserIndex, MiObj.objIndex)
                    End If
                    End If
    
124                 If BusquedaTesoroActiva Then
126                     If UserList(UserIndex).Pos.Map = TesoroNumMapa And UserList(UserIndex).Pos.X = TesoroX And UserList(UserIndex).Pos.Y = TesoroY Then
    
128                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).Name & " encontro el tesoro ¡Felicitaciones!", e_FontTypeNames.FONTTYPE_TALK))
130                         BusquedaTesoroActiva = False

                        End If

                    End If
                
132                 If BusquedaRegaloActiva Then
134                     If UserList(UserIndex).Pos.Map = RegaloNumMapa And UserList(UserIndex).Pos.X = RegaloX And UserList(UserIndex).Pos.Y = RegaloY Then
136                         Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Eventos> " & UserList(UserIndex).Name & " fue el valiente que encontro el gran item magico ¡Felicitaciones!", e_FontTypeNames.FONTTYPE_TALK))
138                         BusquedaRegaloActiva = False

                        End If

                    End If
                
                End If

            End If

        Else

144         If Not UserList(UserIndex).flags.UltimoMensaje = 261 Then
146             Call WriteLocaleMsg(UserIndex, "261", e_FontTypeNames.FONTTYPE_INFO)
148             UserList(UserIndex).flags.UltimoMensaje = 261
            End If
        End If

        
        Exit Sub

PickObj_Err:
150     Call TraceError(Err.Number, Err.Description, "InvUsuario.PickObj", Erl)

        
End Sub

Sub Desequipar(ByVal UserIndex As Integer, ByVal Slot As Byte)
        
        On Error GoTo Desequipar_Err
    
        'Desequipa el item slot del inventario
        Dim obj As t_ObjData

100     If (Slot < LBound(UserList(UserIndex).Invent.Object)) Or (Slot > UBound(UserList(UserIndex).Invent.Object)) Then
            Exit Sub
102     ElseIf UserList(UserIndex).Invent.Object(Slot).ObjIndex = 0 Then
            Exit Sub
        End If

104     obj = ObjData(UserList(UserIndex).Invent.Object(Slot).ObjIndex)

106     Select Case obj.OBJType

            Case e_OBJType.otWeapon
108             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
110             UserList(UserIndex).Invent.WeaponEqpObjIndex = 0
112             UserList(UserIndex).Invent.WeaponEqpSlot = 0
114             UserList(UserIndex).Char.Arma_Aura = ""
116             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 1))
        
118             UserList(UserIndex).Char.WeaponAnim = NingunArma
            
120             If UserList(UserIndex).flags.Montado = 0 Then
122                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                End If
                
124             If obj.MagicDamageBonus > 0 Then
126                 Call WriteUpdateDM(UserIndex)
                End If
    
128         Case e_OBJType.otFlechas
130             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
132             UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
134             UserList(UserIndex).Invent.MunicionEqpSlot = 0
    
                ' Case e_OBJType.otAnillos
                '    UserList(UserIndex).Invent.Object(slot).Equipped = 0
                '    UserList(UserIndex).Invent.AnilloEqpObjIndex = 0
                ' UserList(UserIndex).Invent.AnilloEqpSlot = 0
            
136         Case e_OBJType.otHerramientas
137             If UserList(UserIndex).flags.PescandoEspecial = False Then
138                 UserList(UserIndex).invent.Object(Slot).Equipped = 0
140                 UserList(UserIndex).invent.HerramientaEqpObjIndex = 0
142                 UserList(UserIndex).invent.HerramientaEqpSlot = 0
    
144                 If UserList(UserIndex).flags.UsandoMacro = True Then
146                     Call WriteMacroTrabajoToggle(UserIndex, False)
                    End If
            
148                 UserList(UserIndex).Char.WeaponAnim = NingunArma
                
150                 If UserList(UserIndex).flags.Montado = 0 Then
152                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                    End If
                End If
       
154         Case e_OBJType.otMagicos
    
156             Select Case obj.EfectoMagico

                    Case e_MagicItemEffect.eModifyAttributes
                        If obj.QueAtributo <> 0 Then
162                         UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
164                         UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributosBackUP(obj.QueAtributo) - obj.CuantoAumento
                            ' UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) = UserList(UserIndex).Stats.UserAtributos(obj.QueAtributo) - obj.CuantoAumento
                            
166                         Call WriteFYA(UserIndex)
                        End If

168                 Case e_MagicItemEffect.eModifySkills
                        If obj.Que_Skill <> 0 Then
170                         UserList(UserIndex).Stats.UserSkills(obj.Que_Skill) = UserList(UserIndex).Stats.UserSkills(obj.Que_Skill) - obj.CuantoAumento
                        End If
                        
172                 Case e_MagicItemEffect.eRegenerateHealth
174                     UserList(UserIndex).flags.RegeneracionHP = 0

176                 Case e_MagicItemEffect.eRegenerateMana
178                     UserList(UserIndex).flags.RegeneracionMana = 0

180                 Case e_MagicItemEffect.eIncreaseDamageToNpc
182                     UserList(UserIndex).Stats.MaxHit = UserList(UserIndex).Stats.MaxHit - obj.CuantoAumento
184                     UserList(UserIndex).Stats.MinHIT = UserList(UserIndex).Stats.MinHIT - obj.CuantoAumento

188                 Case e_MagicItemEffect.eInmunityToNpcMagic 'Orbe ignea
190                     UserList(UserIndex).flags.NoMagiaEfecto = 0

192                 Case e_MagicItemEffect.eIncinerate
194                     UserList(UserIndex).flags.incinera = 0

196                 Case e_MagicItemEffect.eParalize
198                     UserList(UserIndex).flags.Paraliza = 0

200                 Case e_MagicItemEffect.eProtectedResources
202                     If UserList(UserIndex).flags.Muerto = 0 Then
                            UserList(UserIndex).Char.CartAnim = NoCart
203                         Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                        End If
                        
206                 Case e_MagicItemEffect.eProtectedInventory
208                     UserList(UserIndex).flags.PendienteDelSacrificio = 0
                 
210                 Case e_MagicItemEffect.ePreventMagicWords
212                     UserList(UserIndex).flags.NoPalabrasMagicas = 0

214                 Case e_MagicItemEffect.ePreventInvisibleDetection
216                     UserList(UserIndex).flags.NoDetectable = 0

218                 Case e_MagicItemEffect.eIncreaseLearningSkills
220                     UserList(UserIndex).flags.PendienteDelExperto = 0

222                 Case e_MagicItemEffect.ePoison
224                     UserList(UserIndex).flags.Envenena = 0

226                 Case e_MagicItemEffect.eRingOfShadows
228                     UserList(UserIndex).flags.AnilloOcultismo = 0

                    Case e_MagicItemEffect.eTalkToDead
                        Call UnsetMask(UserList(UserIndex).flags.StatusMask, e_StatusMask.eTalkToDead)
                        Call WriteConsoleMsg(UserIndex, "Dejas el mundo de los muertos, ya no podrás comunicarte con ellos.", e_FontTypeNames.FONTTYPE_WARNING)
                        Call SendData(SendTarget.ToPCDeadAreaButIndex, UserIndex, PrepareMessageCharacterRemove(4, UserList(UserIndex).Char.charindex, False, True))
                End Select
        
230             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 5))
232             UserList(UserIndex).Char.Otra_Aura = 0
234             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
236             UserList(UserIndex).Invent.MagicoObjIndex = 0
238             UserList(UserIndex).Invent.MagicoSlot = 0
        
256         Case e_OBJType.otArmadura
258             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
260             UserList(UserIndex).Invent.ArmourEqpObjIndex = 0
262             UserList(UserIndex).Invent.ArmourEqpSlot = 0
        
264             If UserList(UserIndex).flags.Navegando = 0 Then
266                 If UserList(UserIndex).flags.Montado = 0 Then
                        Call SetNakedBody(UserList(userIndex))
270                     Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                    End If
                End If
        
272             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 2))
        
274             UserList(UserIndex).Char.Body_Aura = 0

276             If obj.ResistenciaMagica > 0 Then
278                 Call WriteUpdateRM(UserIndex)
                End If
    
280         Case e_OBJType.otCasco
282             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
284             UserList(UserIndex).Invent.CascoEqpObjIndex = 0
286             UserList(UserIndex).Invent.CascoEqpSlot = 0
288             UserList(UserIndex).Char.Head_Aura = 0
290             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 4))

292             UserList(UserIndex).Char.CascoAnim = NingunCasco
294             Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
    
296             If obj.ResistenciaMagica > 0 Then
298                 Call WriteUpdateRM(UserIndex)
                End If
    
300         Case e_OBJType.otEscudo
302             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
304             UserList(UserIndex).Invent.EscudoEqpObjIndex = 0
306             UserList(UserIndex).Invent.EscudoEqpSlot = 0
308             UserList(UserIndex).Char.Escudo_Aura = 0
310             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 3))
        
312             UserList(UserIndex).Char.ShieldAnim = NingunEscudo

314             If UserList(UserIndex).flags.Montado = 0 Then
316                 Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.body, UserList(UserIndex).Char.head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                End If
                
318             If obj.ResistenciaMagica > 0 Then
320                 Call WriteUpdateRM(UserIndex)
                End If
                
322         Case e_OBJType.otDañoMagico
324             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
326             UserList(UserIndex).invent.DañoMagicoEqpObjIndex = 0
328             UserList(UserIndex).invent.DañoMagicoEqpSlot = 0
330             UserList(UserIndex).Char.DM_Aura = 0
332             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 6))
334             Call WriteUpdateDM(UserIndex)
                
336         Case e_OBJType.otResistencia
338             UserList(UserIndex).Invent.Object(Slot).Equipped = 0
340             UserList(UserIndex).Invent.ResistenciaEqpObjIndex = 0
342             UserList(UserIndex).Invent.ResistenciaEqpSlot = 0
344             UserList(UserIndex).Char.RM_Aura = 0
346             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(UserList(UserIndex).Char.charindex, 0, True, 7))
348             Call WriteUpdateRM(UserIndex)
        
        End Select
        
350     Call UpdateUserInv(False, UserIndex, Slot)

        
        Exit Sub

Desequipar_Err:
352     Call TraceError(Err.Number, Err.Description, "InvUsuario.Desequipar", Erl)

        
End Sub

Function SexoPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then
102         SexoPuedeUsarItem = True
            Exit Function

        End If

104     If ObjData(ObjIndex).Mujer = 1 Then
106         SexoPuedeUsarItem = UserList(UserIndex).genero <> e_Genero.Hombre
108     ElseIf ObjData(ObjIndex).Hombre = 1 Then
110         SexoPuedeUsarItem = UserList(UserIndex).genero <> e_Genero.Mujer
        Else
112         SexoPuedeUsarItem = True
        End If

        Exit Function
ErrHandler:
114     Call LogError("SexoPuedeUsarItem")

End Function

Function FaccionPuedeUsarItem(ByVal UserIndex As Integer, ByVal ObjIndex As Integer) As Boolean
        
        On Error GoTo FaccionPuedeUsarItem_Err
        
100     If EsGM(UserIndex) Then
102         FaccionPuedeUsarItem = True
            Exit Function
        End If
        
104     If ObjIndex < 1 Then Exit Function

106     If ObjData(ObjIndex).Real = 1 Then
108         If Status(UserIndex) = e_Facciones.Armada Or Status(UserIndex) = e_Facciones.consejo Then
110             FaccionPuedeUsarItem = esArmada(UserIndex)
            Else
112             FaccionPuedeUsarItem = False
            End If

114     ElseIf ObjData(ObjIndex).Caos = 1 Then

116         If Status(UserIndex) = e_Facciones.Caos Or Status(UserIndex) = e_Facciones.concilio Then
118             FaccionPuedeUsarItem = esCaos(UserIndex)
            Else
120             FaccionPuedeUsarItem = False
            End If
        Else
122         FaccionPuedeUsarItem = True
        End If
        
        Exit Function

FaccionPuedeUsarItem_Err:
124     Call TraceError(Err.Number, Err.Description, "InvUsuario.FaccionPuedeUsarItem", Erl)

        
End Function

Function JerarquiaPuedeUsarItem(ByVal userindex As Integer, ByVal ObjIndex As Integer) As Boolean
       
    With UserList(userindex)
        If .Faccion.RecompensasCaos >= ObjData(ObjIndex).Jerarquia Then
            JerarquiaPuedeUsarItem = True
            Exit Function
        End If

        If .Faccion.RecompensasReal >= ObjData(ObjIndex).Jerarquia Then
            JerarquiaPuedeUsarItem = True
            Exit Function
        End If
    End With
        
End Function
'Equipa barco y hace el cambio de ropaje correspondiente
Sub EquiparBarco(ByVal UserIndex As Integer)
        On Error GoTo EquiparBarco_Err

        Dim Barco As t_ObjData

100     With UserList(UserIndex)
            If .Invent.BarcoObjIndex <= 0 Or .Invent.BarcoObjIndex > UBound(ObjData) Then Exit Sub
102         Barco = ObjData(.Invent.BarcoObjIndex)

104         If .flags.Muerto = 1 Then
106             If Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw Then
                    ' No tenemos la cabeza copada que va con iRopaBuceoMuerto,
                    ' asique asignamos el casper directamente caminando sobre el agua.
108                 .Char.Body = iCuerpoMuerto 'iRopaBuceoMuerto
110                 .Char.Head = iCabezaMuerto
                ElseIf Barco.Ropaje = iTrajeAltoNw Then
          
                ElseIf Barco.Ropaje = iTrajeBajoNw Then
          
                Else
112                 .Char.Body = iFragataFantasmal
114                 .Char.Head = 0
                End If
      
            Else ' Esta vivo

116             If Barco.Ropaje = iTraje Then
118                 .Char.Body = iTraje
120                 .Char.Head = .OrigChar.Head

122                 If .Invent.CascoEqpObjIndex > 0 Then
124                     .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                    End If
                ElseIf Barco.Ropaje = iTrajeAltoNw Then
                    .Char.Body = iTrajeAltoNw
                    .Char.Head = .OrigChar.Head

                    If .Invent.CascoEqpObjIndex > 0 Then
                        .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                    End If
                ElseIf Barco.Ropaje = iTrajeBajoNw Then
                    .Char.Body = iTrajeBajoNw
                    .Char.Head = .OrigChar.Head

                    If .Invent.CascoEqpObjIndex > 0 Then
                        .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                    End If
                Else
126                 .Char.Head = 0
128                 .Char.CascoAnim = NingunCasco
                End If

130             If .Faccion.status = e_Facciones.Armada Or .Faccion.status = e_Facciones.consejo Then
132                 If Barco.Ropaje = iBarca Then .Char.Body = iBarcaArmada
134                 If Barco.Ropaje = iGalera Then .Char.Body = iGaleraArmada
136                 If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonArmada

138             ElseIf .Faccion.status = e_Facciones.Caos Or .Faccion.status = e_Facciones.concilio Then

140                 If Barco.Ropaje = iBarca Then .Char.Body = iBarcaCaos
142                 If Barco.Ropaje = iGalera Then .Char.Body = iGaleraCaos
144                 If Barco.Ropaje = iGaleon Then .Char.Body = iGaleonCaos
          
                Else

146                 If Barco.Ropaje = iBarca Then .Char.Body = IIf(.Faccion.Status = 0, iBarcaCrimi, iBarcaCiuda)
148                 If Barco.Ropaje = iGalera Then .Char.Body = IIf(.Faccion.Status = 0, iGaleraCrimi, iGaleraCiuda)
150                 If Barco.Ropaje = iGaleon Then .Char.Body = IIf(.Faccion.Status = 0, iGaleonCrimi, iGaleonCiuda)
                End If
            End If

152         .Char.ShieldAnim = NingunEscudo
154         .Char.WeaponAnim = NingunArma
    
            Call WriteNavigateToggle(UserIndex, .flags.Navegando)
156         Call WriteNadarToggle(UserIndex, (Barco.Ropaje = iTraje Or Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw), (Barco.Ropaje = iTrajeAltoNw Or Barco.Ropaje = iTrajeBajoNw))
            Call ActualizarVelocidadDeUsuario(UserIndex)
        End With
  
        Exit Sub

EquiparBarco_Err:
158     Call TraceError(Err.Number, Err.Description, "InvUsuario.EquiparBarco", Erl)

End Sub

'Equipa un item del inventario
Sub EquiparInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte)
        On Error GoTo ErrHandler

        Dim obj       As t_ObjData
        Dim ObjIndex  As Integer
        Dim errordesc As String

100     ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
102     obj = ObjData(ObjIndex)
        
104     If PuedeUsarObjeto(UserIndex, ObjIndex, True) > 0 Then
            Exit Sub
        End If

106     With UserList(UserIndex)

108          If .flags.Muerto = 1 Then
                 'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
110              Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                 Exit Sub

             End If

112         Select Case obj.OBJType
                Case e_OBJType.otWeapon
114                 errordesc = "Arma"
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eWeapon) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Si esta equipado lo quita
116                 If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
118                     Call Desequipar(UserIndex, Slot)

                        'Animacion por defecto
120                     .Char.WeaponAnim = NingunArma

122                     If .flags.Montado = 0 Then
124                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                        End If

                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
126                 If .Invent.WeaponEqpObjIndex > 0 Then
128                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
            
130                 If .Invent.HerramientaEqpObjIndex > 0 Then
132                     Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
                    End If

138                 .Invent.Object(Slot).Equipped = 1
140                 .Invent.WeaponEqpObjIndex = .Invent.Object(Slot).ObjIndex
142                 .Invent.WeaponEqpSlot = Slot
            
144                 If obj.Proyectil = 1 And obj.Municion > 0 Then 'Si es un arco, desequipa el escudo.

146                     If .Invent.EscudoEqpObjIndex = 1700 Or _
                           .Invent.EscudoEqpObjIndex = 1730 Or _
                           .Invent.EscudoEqpObjIndex = 1724 Or _
                           .Invent.EscudoEqpObjIndex = 1717 Or _
                           .Invent.EscudoEqpObjIndex = 1699 Then
                           ' Estos escudos SI pueden ser usados con arco.
                        Else

148                         If .Invent.EscudoEqpObjIndex > 0 Then
150                             Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
152                             Call WriteConsoleMsg(UserIndex, "No podes tirar flechas si tenés un escudo equipado. Tu escudo fue desequipado.", e_FontTypeNames.FONTTYPE_INFOIAO)

                            End If

                        End If

                    End If

154                 If obj.DosManos = 1 Then
156                     If .Invent.EscudoEqpObjIndex > 0 Then
158                         Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
160                         Call WriteConsoleMsg(UserIndex, "No puedes usar armas dos manos si tienes un escudo equipado. Tu escudo fue desequipado.", e_FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    End If

                    'Sonido
162                 If obj.SndAura = 0 Then
164                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_SACARARMA, .Pos.X, .Pos.y))
                    Else
166                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.y))
                    End If

168                 If Len(obj.CreaGRH) <> 0 Then
170                     .Char.Arma_Aura = obj.CreaGRH
172                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Arma_Aura, False, 1))
                    End If

174                 If obj.MagicDamageBonus > 0 Then
176                     Call WriteUpdateDM(UserIndex)
                    End If
                
178                 If .flags.Montado = 0 Then
                
180                     If .flags.Navegando = 0 Then
182                         .Char.WeaponAnim = obj.WeaponAnim
184                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                        End If

                    End If
      
186             Case e_OBJType.otHerramientas
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eTool) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Si esta equipado lo quita
188                 If .Invent.Object(Slot).Equipped Then

                        'Quitamos del inv el item
190                     Call Desequipar(UserIndex, Slot)
                        Exit Sub

                    End If

                    'Quitamos el elemento anterior
192                 If .Invent.HerramientaEqpObjIndex > 0 Then
194                     Call Desequipar(UserIndex, .Invent.HerramientaEqpSlot)
                    End If
             
196                 If .Invent.WeaponEqpObjIndex > 0 Then
198                     Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
                    End If
             
200                 .Invent.Object(Slot).Equipped = 1
202                 .Invent.HerramientaEqpObjIndex = ObjIndex
204                 .Invent.HerramientaEqpSlot = Slot
             
206                 If .flags.Montado = 0 Then
                
208                     If .flags.Navegando = 0 Then
210                         .Char.WeaponAnim = obj.WeaponAnim
212                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)

                        End If

                    End If
       
214             Case e_OBJType.otMagicos
216                 errordesc = "Magico"
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eMagicItem) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Si esta equipado lo quita
218                 If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
220                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If

                    'Quitamos el elemento anterior
222                 If .Invent.MagicoObjIndex > 0 Then
224                     Call Desequipar(UserIndex, .Invent.MagicoSlot)
                    End If
        
226                 .Invent.Object(Slot).Equipped = 1
228                 .Invent.MagicoObjIndex = .Invent.Object(Slot).ObjIndex
230                 .Invent.MagicoSlot = Slot
                
232                 Select Case obj.EfectoMagico
                        Case e_MagicItemEffect.eModifyAttributes 'Modif la fuerza, agilidad, carisma, etc
238                         .Stats.UserAtributosBackUP(obj.QueAtributo) = .Stats.UserAtributosBackUP(obj.QueAtributo) + obj.CuantoAumento
240                         .Stats.UserAtributos(obj.QueAtributo) = MinimoInt(.Stats.UserAtributos(obj.QueAtributo) + obj.CuantoAumento, .Stats.UserAtributosBackUP(obj.QueAtributo) * 2)
                
242                         Call WriteFYA(UserIndex)

244                     Case e_MagicItemEffect.eModifySkills
            
246                         .Stats.UserSkills(obj.Que_Skill) = .Stats.UserSkills(obj.Que_Skill) + obj.CuantoAumento

248                     Case e_MagicItemEffect.eRegenerateHealth
250                         .flags.RegeneracionHP = 1

252                     Case e_MagicItemEffect.eRegenerateMana
254                         .flags.RegeneracionMana = 1

256                     Case e_MagicItemEffect.eIncreaseDamageToNpc
258                         .Stats.MaxHit = .Stats.MaxHit + obj.CuantoAumento
260                         .Stats.MinHIT = .Stats.MinHIT + obj.CuantoAumento

262                     Case e_MagicItemEffect.eInmunityToNpcMagic
264                         .flags.NoMagiaEfecto = 1

266                     Case e_MagicItemEffect.eIncinerate
268                         .flags.incinera = 1

270                     Case e_MagicItemEffect.eParalize
272                         .flags.Paraliza = 1

274                     Case e_MagicItemEffect.eProtectedResources
                            If .flags.Navegando = 0 And .flags.Montado = 0 Then
                                 .Char.CartAnim = obj.Ropaje
                                 Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                            End If
                        
280                     Case e_MagicItemEffect.eProtectedInventory
282                         .flags.PendienteDelSacrificio = 1

284                     Case e_MagicItemEffect.ePreventMagicWords
286                         .flags.NoPalabrasMagicas = 1

288                     Case e_MagicItemEffect.ePreventInvisibleDetection
290                         .flags.NoDetectable = 1
                   
292                     Case e_MagicItemEffect.eIncreaseLearningSkills
294                         .flags.PendienteDelExperto = 1

296                     Case e_MagicItemEffect.ePoison
298                         .flags.Envenena = 1

300                     Case e_MagicItemEffect.eRingOfShadows
302                         .flags.AnilloOcultismo = 1

                        Case e_MagicItemEffect.eTalkToDead
                            Call SetMask(.flags.StatusMask, e_StatusMask.eTalkToDead)
                            Call WriteConsoleMsg(UserIndex, "Entras al mundo de los muertos, ahora podrás comunicarte con ellos.", e_FontTypeNames.FONTTYPE_WARNING)
                            Call CheckUpdateNeededUser(UserIndex, USER_NUEVO, True, 1)
                    End Select
            
                    'Sonido
304                 If obj.SndAura <> 0 Then
306                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.SndAura, .Pos.X, .Pos.y))
                    End If
            
308                 If Len(obj.CreaGRH) <> 0 Then
310                     .Char.Otra_Aura = obj.CreaGRH
312                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Otra_Aura, False, 5))
                    End If
    
354             Case e_OBJType.otFlechas
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eAmunition) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Si esta equipado lo quita
356                 If .Invent.Object(Slot).Equipped Then
                        'Quitamos del inv el item
358                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If
                
                    'Quitamos el elemento anterior
360                 If .Invent.MunicionEqpObjIndex > 0 Then
362                     Call Desequipar(UserIndex, .Invent.MunicionEqpSlot)
                    End If
        
364                 .Invent.Object(Slot).Equipped = 1
366                 .Invent.MunicionEqpObjIndex = .Invent.Object(Slot).ObjIndex
368                 .Invent.MunicionEqpSlot = Slot

370             Case e_OBJType.otArmadura
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eArmor) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                
                    Dim Ropaje As Integer
                    Ropaje = ObtenerRopaje(UserIndex, Obj)

372                 If Ropaje = 0 Then
374                     Call WriteConsoleMsg(UserIndex, "Hay un error con este objeto. Infórmale a un administrador.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    'Si esta equipado lo quita
376                 If .Invent.Object(Slot).Equipped Then
378                     Call Desequipar(UserIndex, Slot)

380                     If .flags.Navegando = 0 And .flags.Montado = 0 Then
                            Call SetNakedBody(UserList(userIndex))
384                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                        Else
386                         .flags.Desnudo = 1
                        End If

                        Exit Sub

                    End If

                    'Quita el anterior
388                 If .Invent.ArmourEqpObjIndex > 0 Then
390                     errordesc = "Armadura 2"
392                     Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)
394                     errordesc = "Armadura 3"

                    End If
  
                    'Lo equipa
396                 If Len(obj.CreaGRH) <> 0 Then
398                     .Char.Body_Aura = obj.CreaGRH
400                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Body_Aura, False, 2))

                    End If
            
402                 .Invent.Object(Slot).Equipped = 1
404                 .Invent.ArmourEqpObjIndex = .Invent.Object(Slot).ObjIndex
406                 .Invent.ArmourEqpSlot = Slot

408                 If .flags.Montado = 0 And .flags.Navegando = 0 Then
410                     .Char.Body = Ropaje

412                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                    End If
                    
414                 .flags.Desnudo = 0

416                 If obj.ResistenciaMagica > 0 Then
418                     Call WriteUpdateRM(UserIndex)
                    End If
    
420             Case e_OBJType.otCasco
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eHelm) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Si esta equipado lo quita
422                 If .Invent.Object(Slot).Equipped Then
424                     Call Desequipar(UserIndex, Slot)
                
426                     .Char.CascoAnim = NingunCasco
428                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
                        Exit Sub

                    End If
    
                    'Quita el anterior
430                 If .Invent.CascoEqpObjIndex > 0 Then
432                     Call Desequipar(UserIndex, .Invent.CascoEqpSlot)
                    End If

434                 errordesc = "Casco"

                    'Lo equipa
436                 If Len(obj.CreaGRH) <> 0 Then
438                     .Char.Head_Aura = obj.CreaGRH
440                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Head_Aura, False, 4))
                    End If
            
442                 .Invent.Object(Slot).Equipped = 1
444                 .Invent.CascoEqpObjIndex = .Invent.Object(Slot).ObjIndex
446                 .Invent.CascoEqpSlot = Slot
            
448                 If .flags.Navegando = 0 Then
450                     .Char.CascoAnim = obj.CascoAnim
452                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
                    End If
                
454                 If obj.ResistenciaMagica > 0 Then
456                     Call WriteUpdateRM(UserIndex)
                    End If

458             Case e_OBJType.otEscudo
                    If IsSet(.flags.DisabledSlot, e_InventorySlotMask.eShiled) Then
                        Call WriteLocaleMsg(UserIndex, MsgCantEquipYet, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    'Si esta equipado lo quita
460                 If .Invent.Object(Slot).Equipped Then
462                     Call Desequipar(UserIndex, Slot)
                 
464                     .Char.ShieldAnim = NingunEscudo

466                     If .flags.Montado = 0 And .flags.Navegando = 0 Then
468                         Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)

                        End If

                        Exit Sub

                    End If
     
                    'Quita el anterior
470                 If .Invent.EscudoEqpObjIndex > 0 Then
472                     Call Desequipar(UserIndex, .Invent.EscudoEqpSlot)
                    End If
     
                    'Lo equipa
474                 If .Invent.Object(Slot).ObjIndex = 1700 Or _
                       .Invent.Object(Slot).ObjIndex = 1730 Or _
                       .Invent.Object(Slot).ObjIndex = 1724 Or _
                       .Invent.Object(Slot).ObjIndex = 1717 Or _
                       .Invent.Object(Slot).ObjIndex = 1699 Then
             
                    Else

476                     If .Invent.WeaponEqpObjIndex > 0 Then
478                         If ObjData(.invent.WeaponEqpObjIndex).Proyectil = 1 And ObjData(.invent.WeaponEqpObjIndex).Municion > 0 Then
480                             Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
482                             Call WriteConsoleMsg(UserIndex, "No podes sostener el escudo si tenes que tirar flechas. Tu arco fue desequipado.", e_FontTypeNames.FONTTYPE_INFOIAO)
                            End If
                        End If

                    End If

484                 If .Invent.WeaponEqpObjIndex > 0 Then
486                     If ObjData(.Invent.WeaponEqpObjIndex).DosManos = 1 Then
488                         Call Desequipar(UserIndex, .Invent.WeaponEqpSlot)
490                         Call WriteConsoleMsg(UserIndex, "No puedes equipar un escudo si tienes un arma dos manos equipada. Tu arma fue desequipada.", e_FontTypeNames.FONTTYPE_INFOIAO)
                        End If
                    End If

492                 errordesc = "Escudo"

494                 If Len(obj.CreaGRH) <> 0 Then
496                     .Char.Escudo_Aura = obj.CreaGRH
498                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.Escudo_Aura, False, 3))
                    End If

500                 .Invent.Object(Slot).Equipped = 1
502                 .Invent.EscudoEqpObjIndex = .Invent.Object(Slot).ObjIndex
504                 .Invent.EscudoEqpSlot = Slot

506                 If .flags.Navegando = 0 And .flags.Montado = 0 Then
508                     .Char.ShieldAnim = obj.ShieldAnim
510                     Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
                    End If

512                 If obj.ResistenciaMagica > 0 Then
514                     Call WriteUpdateRM(UserIndex)
                    End If

516             Case e_OBJType.otDañoMagico, e_OBJType.otResistencia
                    'Si esta equipado lo quita
518                 If .Invent.Object(Slot).Equipped Then
520                     Call Desequipar(UserIndex, Slot)
                        Exit Sub
                    End If
     
                    'Quita el anterior
522                 If .invent.DañoMagicoEqpSlot > 0 Then
524                     Call Desequipar(UserIndex, .invent.DañoMagicoEqpSlot)
                    End If

546                 If .Invent.ResistenciaEqpSlot > 0 Then
548                     Call Desequipar(UserIndex, .Invent.ResistenciaEqpSlot)
                    End If
                

526                 .Invent.Object(Slot).Equipped = 1

                    If ObjData(.Invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otResistencia Then
                        .Invent.ResistenciaEqpObjIndex = .Invent.Object(Slot).ObjIndex
530                     .Invent.ResistenciaEqpSlot = Slot
                        Call WriteUpdateRM(userindex)
                    ElseIf ObjData(.invent.Object(Slot).ObjIndex).OBJType = e_OBJType.otDañoMagico Then
528                     .invent.DañoMagicoEqpObjIndex = .invent.Object(Slot).ObjIndex
                        .invent.DañoMagicoEqpSlot = Slot
538                     Call WriteUpdateDM(userindex)
                        
                    End If
532                 If Len(obj.CreaGRH) <> 0 Then
534                     .Char.DM_Aura = obj.CreaGRH
536                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageAuraToChar(.Char.charindex, .Char.DM_Aura, False, 6))
                    End If
                Case e_OBJType.OtDonador
                    If obj.Subtipo = 4 Then
                       Call EquipAura(Slot, .Invent, UserIndex)
                    End If

            End Select
            
        End With

        'Actualiza
564     Call UpdateUserInv(False, UserIndex, Slot)

        Exit Sub
    
ErrHandler:
566     Debug.Print errordesc
568     Call LogError("EquiparInvItem Slot:" & Slot & " - Error: " & Err.Number & " - Error Description : " & Err.Description & "- " & errordesc)

End Sub

Public Sub EquipAura(ByVal Slot As Integer, ByRef inventory As t_Inventario, ByVal UserIndex As Integer)
    
    If inventory.Object(Slot).Equipped Then
        inventory.Object(Slot).Equipped = False
        Exit Sub
    End If
    If Slot < 1 Or Slot > UBound(inventory.Object) Then Exit Sub

    Dim Index As Integer
    Dim obj As t_ObjData
    For Index = 1 To UBound(inventory.Object)
        If Index <> Slot And inventory.Object(Index).Equipped Then
            If inventory.Object(Index).objIndex > 0 Then
                If inventory.Object(Index).objIndex > 0 Then
                    obj = ObjData(inventory.Object(Index).objIndex)
                    If obj.OBJType = OtDonador And obj.Subtipo = 4 Then
                       inventory.Object(Index).Equipped = 0
                       Call UpdateUserInv(False, UserIndex, Index)
                    End If
                End If
             End If
        End If
    Next Index
    inventory.Object(Slot).Equipped = 1
End Sub

Public Function CheckClaseTipo(ByVal UserIndex As Integer, ItemIndex As Integer) As Boolean

        On Error GoTo ErrHandler

100     If EsGM(UserIndex) Then

102         CheckClaseTipo = True
            Exit Function

        End If

104     Select Case ObjData(ItemIndex).ClaseTipo

            Case 0
106             CheckClaseTipo = True
                Exit Function

108         Case 2

110             If UserList(UserIndex).clase = e_Class.Mage Then CheckClaseTipo = True
112             If UserList(UserIndex).clase = e_Class.Druid Then CheckClaseTipo = True
                Exit Function

114         Case 1

116             If UserList(UserIndex).clase = e_Class.Warrior Then CheckClaseTipo = True
118             If UserList(UserIndex).clase = e_Class.Assasin Then CheckClaseTipo = True
120             If UserList(UserIndex).clase = e_Class.Bard Then CheckClaseTipo = True
122             If UserList(UserIndex).clase = e_Class.Cleric Then CheckClaseTipo = True
124             If UserList(UserIndex).clase = e_Class.Paladin Then CheckClaseTipo = True
126             If UserList(UserIndex).clase = e_Class.Trabajador Then CheckClaseTipo = True
128             If UserList(UserIndex).clase = e_Class.Hunter Then CheckClaseTipo = True
                Exit Function

        End Select

        Exit Function
ErrHandler:
130     Call LogError("Error CheckClaseTipo ItemIndex:" & ItemIndex)

End Function

Sub UseInvItem(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal ByClick As Byte)

        On Error GoTo hErr

        '*************************************************
        'Author: Unknown
        'Last modified: 24/01/2007
        'Handels the usage of items from inventory box.
        '24/01/2007 Pablo (ToxicWaste) - Agrego el Cuerno de la Armada y la Legión.
        '24/01/2007 Pablo (ToxicWaste) - Utilización nueva de Barco en lvl 20 por clase Pirata y Pescador.
        '*************************************************

        Dim obj      As t_ObjData

        Dim ObjIndex As Integer

        Dim TargObj  As t_ObjData

        Dim MiObj    As t_Obj
        
100     With UserList(UserIndex)

102         If .Invent.Object(Slot).amount = 0 Then Exit Sub
            If Not CanUseItem(.flags, .Counters) Then
                Call WriteLocaleMsg(UserIndex, 395, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If PuedeUsarObjeto(UserIndex, .invent.Object(Slot).objIndex, True) > 0 Then
                Exit Sub
            End If
104         obj = ObjData(.Invent.Object(Slot).ObjIndex)
            Dim TimeSinceLastUse As Long: TimeSinceLastUse = GetTickCount() - .CdTimes(obj.cdType)
            If TimeSinceLastUse < obj.Cooldown Then Exit Sub
            
            If IsSet(obj.ObjFlags, e_ObjFlags.e_UseOnSafeAreaOnly) Then
                If MapInfo(.Pos.Map).Seguro = 0 Or Zona(.ZonaId).Segura = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Solo podes usar este objeto en mapas seguros.", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
    
106         If obj.OBJType = e_OBJType.otWeapon Then
108             If obj.Proyectil = 1 Then
    
                    'valido para evitar el flood pero no bloqueo. El bloqueo se hace en WLC con proyectiles.
110                 If ByClick <> 0 Then
                        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
                    Else
                        If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
                    End If
                Else
                    'dagas
112                 If ByClick <> 0 Then
                        If Not IntervaloPermiteUsar(UserIndex) Then Exit Sub
                    Else
                        If Not IntervaloPermiteUsarClick(UserIndex) Then Exit Sub
                    End If
                End If
    
            Else
                If ByClick <> 0 Then
                    If Not IntervaloPermiteUsarClick(userindex) Then Exit Sub
                Else
                    If Not IntervaloPermiteUsar(userindex) Then Exit Sub
                End If
            End If
    
118         If .flags.Meditando Then
120             .flags.Meditando = False
122             .Char.FX = 0
124             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))
            End If
    
126         If obj.Newbie = 1 And Not EsNewbie(UserIndex) And Not EsGM(UserIndex) Then
128             Call WriteConsoleMsg(UserIndex, "Solo los newbies pueden usar estos objetos.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
    
130         If .Stats.ELV < obj.MinELV Then
132             Call WriteConsoleMsg(UserIndex, "Necesitas ser nivel " & obj.MinELV & " para usar este item.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
    
134         ObjIndex = .Invent.Object(Slot).ObjIndex
136         .flags.TargetObjInvIndex = ObjIndex
138         .flags.TargetObjInvSlot = Slot
    
140         Select Case obj.OBJType
    
                Case e_OBJType.otUseOnce
    
142                 If .flags.Muerto = 1 Then
144                     Call WriteLocaleMsg(UserIndex, 77, e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
                    'Usa el item
146                 .Stats.MinHam = .Stats.MinHam + obj.MinHam
    
148                 If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
152                 Call WriteUpdateHungerAndThirst(UserIndex)

                    'Sonido
154                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(e_SoundIndex.SOUND_COMIDA, .Pos.X, .Pos.y))

                    'Quitamos del inv el item
156                 Call QuitarUserInvItem(UserIndex, Slot, 1)
            
158                 Call UpdateUserInv(False, UserIndex, Slot)
                    
                    UserList(UserIndex).flags.ModificoInventario = True
                    
160             Case e_OBJType.otGuita
    
162                 If .flags.Muerto = 1 Then
164                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If

166                 .Stats.GLD = .Stats.GLD + .Invent.Object(Slot).amount
168                 .Invent.Object(Slot).amount = 0
170                 .Invent.Object(Slot).ObjIndex = 0
172                 .Invent.NroItems = .Invent.NroItems - 1
                    .flags.ModificoInventario = True
174                 Call UpdateUserInv(False, UserIndex, Slot)
176                 Call WriteUpdateGold(UserIndex)


            
178             Case e_OBJType.otWeapon
    
180                 If .flags.Muerto = 1 Then
182                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
184                 If Not .Stats.MinSta > 0 Then
186                     Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
188                 If ObjData(ObjIndex).Proyectil = 1 Then
                        If IsSet(.flags.StatusMask, e_StatusMask.eTransformed) Then
                            Call WriteLocaleMsg(UserIndex, MsgCantUseBowTransformed, e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
190                     Call WriteWorkRequestTarget(UserIndex, Proyectiles)
                    Else
192                     If .flags.TargetObj = Wood Then
194                         If .Invent.Object(Slot).ObjIndex = DAGA Then
196                             Call TratarDeHacerFogata(.flags.TargetObjMap, .flags.TargetObjX, .flags.TargetObjY, UserIndex)
                            End If
                        End If
    
                    End If
            
198                 If .Invent.Object(Slot).Equipped = 0 Then
                        Exit Sub
                    End If
            
200             Case e_OBJType.otHerramientas
    
202                 If .flags.Muerto = 1 Then
204                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
            
206                 If Not .Stats.MinSta > 0 Then
208                     Call WriteLocaleMsg(UserIndex, "93", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
                    'Solo si es herramienta ;) (en realidad si no es ni proyectil ni daga)
210                 If .Invent.Object(Slot).Equipped = 0 Then
                        'Call WriteConsoleMsg(UserIndex, "Antes de usar la herramienta deberias equipartela.", e_FontTypeNames.FONTTYPE_INFO)
212                     Call WriteLocaleMsg(UserIndex, "376", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
214                 Select Case obj.Subtipo
                    
                        Case 1, 2  ' Herramientas del Pescador - Caña y Red
216                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Pescar)
                    
218                     Case 3     ' Herramientas de Alquimia - Tijeras
220                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Alquimia)
                    
222                     Case 4     ' Herramientas de Alquimia - Olla
224                         Call EnivarObjConstruiblesAlquimia(UserIndex)
226                         Call WriteShowAlquimiaForm(UserIndex)
                    
228                     Case 5     ' Herramientas de Carpinteria - Serrucho
230                         Call EnivarObjConstruibles(UserIndex)
232                         Call WriteShowCarpenterForm(UserIndex)
                    
234                     Case 6     ' Herramientas de Tala - Hacha
236                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Talar)
    
238                     Case 7     ' Herramientas de Herrero - Martillo
240                         Call WriteConsoleMsg(UserIndex, "Debes hacer click derecho sobre el yunque.", e_FontTypeNames.FONTTYPE_INFOIAO)
    
242                     Case 8     ' Herramientas de Mineria - Piquete
244                         Call WriteWorkRequestTarget(UserIndex, e_Skill.Mineria)
                    
246                     Case 9     ' Herramientas de Sastreria - Costurero
248                         Call EnivarObjConstruiblesSastre(UserIndex)
250                         Call WriteShowSastreForm(UserIndex)
    
                    End Select
        
252             Case e_OBJType.otPociones
    
254                 If .flags.Muerto = 1 Then
256                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
                    
                    If Not IntervaloPermiteGolpeUsar(UserIndex, False) Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Debes esperar unos momentos para tomar otra poción!!", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
            
258                 .flags.TomoPocion = True
260                 .flags.TipoPocion = obj.TipoPocion
                    
                    Dim CabezaFinal  As Integer
    
                    Dim CabezaActual As Integer
    
262                 Select Case .flags.TipoPocion
            
                        Case 1 'Modif la agilidad
264                         .flags.DuracionEfecto = obj.DuracionEfecto
            
                            'Usa el item
266                         .Stats.UserAtributos(e_Atributos.Agilidad) = MinimoInt(.Stats.UserAtributos(e_Atributos.Agilidad) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Agilidad) * 2)
                    
268                         Call WriteFYA(UserIndex)
                    
                            'Quitamos del inv el item
270                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
272                         If obj.Snd1 <> 0 Then
274                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                            Else
276                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                            End If
            
278                     Case 2 'Modif la fuerza
280                         .flags.DuracionEfecto = obj.DuracionEfecto
            
                            'Usa el item
282                         .Stats.UserAtributos(e_Atributos.Fuerza) = MinimoInt(.Stats.UserAtributos(e_Atributos.Fuerza) + RandomNumber(obj.MinModificador, obj.MaxModificador), .Stats.UserAtributosBackUP(e_Atributos.Fuerza) * 2)
                    
                            'Quitamos del inv el item
284                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
286                         If obj.Snd1 <> 0 Then
288                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                            Else
290                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                            End If
    
292                         Call WriteFYA(UserIndex)
    
294                     Case 3 'Pocion roja, restaura HP
                    
                            'Usa el item
                            Dim HealingAmount As Long
                            HealingAmount = RandomNumber(obj.MinModificador, obj.MaxModificador) * UserMod.GetSelfHealingBonus(UserList(UserIndex))
296                         Call UserMod.ModifyHealth(UserIndex, HealingAmount)
                            'Quitamos del inv el item
300                         Call QuitarUserInvItem(UserIndex, Slot, 1)
302                         If obj.Snd1 <> 0 Then
304                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                            Else
306                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
                            End If
                
308                     Case 4 'Pocion azul, restaura MANA
                
                            Dim porcentajeRec As Byte
310                         porcentajeRec = obj.Porcentaje
                    
                            'Usa el item
312                          .Stats.MinMAN = IIf(.Stats.MinMAN > 20000, 20000, .Stats.MinMAN + Porcentaje(.Stats.MaxMAN, porcentajeRec))
314                         If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN

                    
                            'Quitamos del inv el item
316                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
318                         If obj.Snd1 <> 0 Then
320                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                        
                            Else
322                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                            End If
                    
324                     Case 5 ' Pocion violeta
    
326                         If .flags.Envenenado > 0 Then
328                             .flags.Envenenado = 0
330                             Call WriteConsoleMsg(UserIndex, "Te has curado del envenenamiento.", e_FontTypeNames.FONTTYPE_INFO)
                                'Quitamos del inv el item
332                             Call QuitarUserInvItem(UserIndex, Slot, 1)
    
334                             If obj.Snd1 <> 0 Then
336                                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                        
                                Else
338                                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                                End If
    
                            Else
340                             Call WriteConsoleMsg(UserIndex, "¡No te encuentras envenenado!", e_FontTypeNames.FONTTYPE_INFO)
    
                            End If
                    
342                     Case 6  ' Remueve Parálisis
    
344                         If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
346                             If .flags.Paralizado = 1 Then
348                                 .flags.Paralizado = 0
350                                 Call WriteParalizeOK(UserIndex)
    
                                End If
                            
352                             If .flags.Inmovilizado = 1 Then
354                                 .Counters.Inmovilizado = 0
356                                 .flags.Inmovilizado = 0
358                                 Call WriteInmovilizaOK(UserIndex)
    
                                End If
                            
                            
                            
360                             Call QuitarUserInvItem(UserIndex, Slot, 1)
    
362                             If obj.Snd1 <> 0 Then
364                                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                        
                                Else
366                                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(255, .Pos.X, .Pos.y))
    
                                End If
    
368                             Call WriteConsoleMsg(UserIndex, "Te has removido la paralizis.", e_FontTypeNames.FONTTYPE_INFOIAO)
                            Else
370                             Call WriteConsoleMsg(UserIndex, "No estas paralizado.", e_FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
                    
372                     Case 7  ' Pocion Naranja
374                         .Stats.MinSta = .Stats.MinSta + RandomNumber(obj.MinModificador, obj.MaxModificador)
    
376                         If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
                        
                            'Quitamos del inv el item
378                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
380                         If obj.Snd1 <> 0 Then
382                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                                
                            Else
384                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                            End If
    
386                     Case 8  ' Pocion cambio cara
    
388                         Select Case .genero
    
                                Case e_Genero.Hombre
    
390                                 Select Case .raza
    
                                        Case e_Raza.Humano
392                                         CabezaFinal = RandomNumber(1, 40)
    
394                                     Case e_Raza.Elfo
396                                         CabezaFinal = RandomNumber(101, 132)
    
398                                     Case e_Raza.Drow
400                                         CabezaFinal = RandomNumber(201, 229)
    
402                                     Case e_Raza.Enano
404                                         CabezaFinal = RandomNumber(301, 329)
    
406                                     Case e_Raza.Gnomo
408                                         CabezaFinal = RandomNumber(401, 429)
    
410                                     Case e_Raza.Orco
412                                         CabezaFinal = RandomNumber(501, 529)
    
                                    End Select
    
414                             Case e_Genero.Mujer
    
416                                 Select Case .raza
    
                                        Case e_Raza.Humano
418                                         CabezaFinal = RandomNumber(50, 80)
    
420                                     Case e_Raza.Elfo
422                                         CabezaFinal = RandomNumber(150, 179)
    
424                                     Case e_Raza.Drow
426                                         CabezaFinal = RandomNumber(250, 279)
    
428                                     Case e_Raza.Gnomo
430                                         CabezaFinal = RandomNumber(350, 379)
    
432                                     Case e_Raza.Enano
434                                         CabezaFinal = RandomNumber(450, 479)
    
436                                     Case e_Raza.Orco
438                                         CabezaFinal = RandomNumber(550, 579)
    
                                    End Select
    
                            End Select
                
440                         .Char.Head = CabezaFinal
442                         .OrigChar.Head = CabezaFinal
444                         Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
                            'Quitamos del inv el item
                            
                            UserList(UserIndex).Counters.timeFx = 3
446                         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
    
448                         If CabezaActual <> CabezaFinal Then
450                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
452                             Call WriteConsoleMsg(UserIndex, "¡Rayos! Te tocó la misma cabeza, item no consumido. Tienes otra oportunidad.", e_FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
454                         Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                        
456                     Case 9  ' Pocion sexo
        
458                         Select Case .genero
    
                                Case e_Genero.Hombre
460                                 .genero = e_Genero.Mujer
                        
462                             Case e_Genero.Mujer
464                                 .genero = e_Genero.Hombre
                        
                            End Select
                
466                         Select Case .genero
    
                                Case e_Genero.Hombre
    
468                                 Select Case .raza
    
                                        Case e_Raza.Humano
470                                         CabezaFinal = RandomNumber(1, 40)
    
472                                     Case e_Raza.Elfo
474                                         CabezaFinal = RandomNumber(101, 132)
    
476                                     Case e_Raza.Drow
478                                         CabezaFinal = RandomNumber(201, 229)
    
480                                     Case e_Raza.Enano
482                                         CabezaFinal = RandomNumber(301, 329)
    
484                                     Case e_Raza.Gnomo
486                                         CabezaFinal = RandomNumber(401, 429)
    
488                                     Case e_Raza.Orco
490                                         CabezaFinal = RandomNumber(501, 529)
    
                                    End Select
    
492                             Case e_Genero.Mujer
    
494                                 Select Case .raza
    
                                        Case e_Raza.Humano
496                                         CabezaFinal = RandomNumber(50, 80)
    
498                                     Case e_Raza.Elfo
500                                         CabezaFinal = RandomNumber(150, 179)
    
502                                     Case e_Raza.Drow
504                                         CabezaFinal = RandomNumber(250, 279)
    
506                                     Case e_Raza.Gnomo
508                                         CabezaFinal = RandomNumber(350, 379)
    
510                                     Case e_Raza.Enano
512                                         CabezaFinal = RandomNumber(450, 479)
    
514                                     Case e_Raza.Orco
516                                         CabezaFinal = RandomNumber(550, 579)
    
                                    End Select
    
                            End Select
                
518                         .Char.Head = CabezaFinal
520                         .OrigChar.Head = CabezaFinal
522                         Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
                            'Quitamos del inv el item
                            UserList(UserIndex).Counters.timeFx = 3
524                         Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
526                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
528                         If obj.Snd1 <> 0 Then
530                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                            Else
532                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                            End If
                    
534                     Case 10  ' Invisibilidad
                
536                         If .flags.invisible = 0 And .Counters.DisabledInvisibility = 0 Then
                                If IsSet(.flags.StatusMask, eTaunting) Then
                                    Call WriteConsoleMsg(UserIndex, "No tiene efecto.", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                                    Exit Sub
                                End If
538                             .flags.invisible = 1
540                             .Counters.Invisibilidad = obj.DuracionEfecto
542                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True, .Pos.X, .Pos.Y))
544                             Call WriteContadores(UserIndex)
546                             Call QuitarUserInvItem(UserIndex, Slot, 1)
    
548                             If obj.Snd1 <> 0 Then
550                                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                                
                                Else
552                                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave("123", .Pos.X, .Pos.y))
    
                                End If
    
554                             Call WriteConsoleMsg(UserIndex, "Te has escondido entre las sombras...", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                            
                            Else
556                             Call WriteConsoleMsg(UserIndex, "Ya estas invisible.", e_FontTypeNames.FONTTYPE_New_Amarillo_Oscuro)
                                Exit Sub
                            End If
                            
                        ' Poción que limpia todo
626                     Case 13
                    
628                         Call QuitarUserInvItem(UserIndex, Slot, 1)
630                         .flags.Envenenado = 0
632                         .flags.Incinerado = 0
                        
634                         If .flags.Inmovilizado = 1 Then
636                             .Counters.Inmovilizado = 0
638                             .flags.Inmovilizado = 0
640                             Call WriteInmovilizaOK(UserIndex)
                            
    
                            End If
                        
642                         If .flags.Paralizado = 1 Then
644                             .flags.Paralizado = 0
646                             Call WriteParalizeOK(UserIndex)
                            
    
                            End If
                        
648                         If .flags.Ceguera = 1 Then
650                             .flags.Ceguera = 0
652                             Call WriteBlindNoMore(UserIndex)
                            
    
                            End If
                        
654                         If .flags.Maldicion = 1 Then
656                             .flags.Maldicion = 0
658                             .Counters.Maldicion = 0
    
                            End If
                        
660                         .Stats.MinSta = .Stats.MaxSta
662                         .Stats.MinAGU = .Stats.MaxAGU
664                         .Stats.MinMAN = .Stats.MaxMAN
666                         .Stats.MinHp = .Stats.MaxHp
668                         .Stats.MinHam = .Stats.MaxHam
                        
                        
674                         Call WriteUpdateHungerAndThirst(UserIndex)
676                         Call WriteConsoleMsg(UserIndex, "Donador> Te sentis sano y lleno.", e_FontTypeNames.FONTTYPE_WARNING)
    
678                         If obj.Snd1 <> 0 Then
680                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                            
                            Else
682                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                            End If
    
                        ' Poción runa
684                     Case 14
                                       
686                         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
688                             Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                        
                            Dim Map     As Integer
    
                            Dim X       As Byte
    
                            Dim Y       As Byte
    
                            Dim DeDonde As t_WorldPos
    
690                         Call QuitarUserInvItem(UserIndex, Slot, 1)
                
692                         Select Case .Hogar
    
                                Case e_Ciudad.cUllathorpe
694                                 DeDonde = Ullathorpe
                                
696                             Case e_Ciudad.cNix
698                                 DeDonde = Nix
                    
700                             Case e_Ciudad.cBanderbill
702                                 DeDonde = Banderbill
                            
704                             Case e_Ciudad.cLindos
706                                 DeDonde = Lindos
                                
708                             Case e_Ciudad.cArghal
710                                 DeDonde = Arghal
                                
712                             Case e_Ciudad.cArkhein
714                                 DeDonde = Arkhein
                                
716                             Case Else
718                                 DeDonde = Ullathorpe
    
                            End Select
                        
720                         Map = DeDonde.Map
722                         X = DeDonde.X
724                         Y = DeDonde.Y
                        
726                         Call FindLegalPos(UserIndex, Map, X, Y)
728                         Call WarpUserChar(UserIndex, Map, X, Y, True)
730                         Call WriteConsoleMsg(UserIndex, "Ya estas a salvo...", e_FontTypeNames.FONTTYPE_WARNING)
    
732                         If obj.Snd1 <> 0 Then
734                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                            
                            Else
736                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                            End If
        
774                     Case 16 ' Divorcio
    
776                         If .flags.Casado = 1 Then
    
                                Dim tUser As t_UserReference
    
                                '.flags.Pareja
778                             tUser = NameIndex(GetUserSpouse(.flags.SpouseId))

                            
782                             If Not IsValidUserRef(tUser) Then
790                                 Call WriteConsoleMsg(UserIndex, "Tu pareja deberás estar conectada para divorciarse.", e_FontTypeNames.FONTTYPE_INFOIAO)
                                Else
780                                 Call QuitarUserInvItem(UserIndex, Slot, 1)
794                                 UserList(tUser.ArrayIndex).flags.Casado = 0
796                                 UserList(tUser.ArrayIndex).flags.SpouseId = 0
798                                 .flags.Casado = 0
800                                 .flags.SpouseId = 0
802                                 Call WriteConsoleMsg(UserIndex, "Te has divorciado.", e_FontTypeNames.FONTTYPE_INFOIAO)
804                                 Call WriteConsoleMsg(tUser.ArrayIndex, .name & " se ha divorciado de ti.", e_FontTypeNames.FONTTYPE_INFOIAO)

                                    If obj.Snd1 <> 0 Then
808                                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                                    Else
810                                     Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
                                    End If
                                
                                End If
    
806
                        
                            Else
812                             Call WriteConsoleMsg(UserIndex, "No estas casado.", e_FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
814                     Case 17 'Cara legendaria
    
816                         Select Case .genero
    
                                Case e_Genero.Hombre
    
818                                 Select Case .raza
    
                                        Case e_Raza.Humano
820                                         CabezaFinal = RandomNumber(684, 686)
    
822                                     Case e_Raza.Elfo
824                                         CabezaFinal = RandomNumber(690, 692)
    
826                                     Case e_Raza.Drow
828                                         CabezaFinal = RandomNumber(696, 698)
    
830                                     Case e_Raza.Enano
832                                         CabezaFinal = RandomNumber(702, 704)
    
834                                     Case e_Raza.Gnomo
836                                         CabezaFinal = RandomNumber(708, 710)
    
838                                     Case e_Raza.Orco
840                                         CabezaFinal = RandomNumber(714, 716)
    
                                    End Select
    
842                             Case e_Genero.Mujer
    
844                                 Select Case .raza
    
                                        Case e_Raza.Humano
846                                         CabezaFinal = RandomNumber(687, 689)
    
848                                     Case e_Raza.Elfo
850                                         CabezaFinal = RandomNumber(693, 695)
    
852                                     Case e_Raza.Drow
854                                         CabezaFinal = RandomNumber(699, 701)
    
856                                     Case e_Raza.Gnomo
858                                         CabezaFinal = RandomNumber(705, 707)
    
860                                     Case e_Raza.Enano
862                                         CabezaFinal = RandomNumber(711, 713)
    
864                                     Case e_Raza.Orco
866                                         CabezaFinal = RandomNumber(717, 719)
    
                                    End Select
    
                            End Select
    
868                         CabezaActual = .OrigChar.Head
                            
870                         .Char.Head = CabezaFinal
872                         .OrigChar.Head = CabezaFinal
874                         Call ChangeUserChar(UserIndex, .Char.body, CabezaFinal, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.CartAnim)
    
                            'Quitamos del inv el item
876                         If CabezaActual <> CabezaFinal Then
                                UserList(UserIndex).Counters.timeFx = 3
878                             Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, 102, 0, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
880                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
882                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
884                             Call WriteConsoleMsg(UserIndex, "¡Rayos! No pude asignarte una cabeza nueva, item no consumido. ¡Proba de nuevo!", e_FontTypeNames.FONTTYPE_INFOIAO)
    
                            End If
    
886                     Case 18  ' tan solo crea una particula por determinado tiempo
    
                            Dim Particula           As Integer
    
                            Dim Tiempo              As Long
    
                            Dim ParticulaPermanente As Byte
    
                            Dim sobrechar           As Byte
    
888                         If obj.CreaParticula <> "" Then
890                             Particula = val(ReadField(1, obj.CreaParticula, Asc(":")))
892                             Tiempo = val(ReadField(2, obj.CreaParticula, Asc(":")))
894                             ParticulaPermanente = val(ReadField(3, obj.CreaParticula, Asc(":")))
896                             sobrechar = val(ReadField(4, obj.CreaParticula, Asc(":")))
                                
898                             If ParticulaPermanente = 1 Then
900                                 .Char.ParticulaFx = Particula
902                                 .Char.loops = Tiempo
    
                                End If
                                
904                             If sobrechar = 1 Then
906                                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageParticleFXToFloor(.Pos.X, .Pos.y, Particula, Tiempo))
                                Else
                                    UserList(UserIndex).Counters.timeFx = 3
908                                 Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, Particula, Tiempo, False, , UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                                End If
    
                            End If
                            
910                         If obj.CreaFX <> 0 Then
912                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageFxPiso(obj.CreaFX, .Pos.X, .Pos.y))
                            End If
                            
914                         If obj.Snd1 <> 0 Then
916                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
    
                            End If
                            
918                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
920                     Case 19 ' Reseteo de skill
    
                            Dim S As Byte
                    
922                         If .Stats.UserSkills(e_Skill.liderazgo) >= 80 Then
924                             Call WriteConsoleMsg(UserIndex, "Has fundado un clan, no podes resetar tus skills. ", e_FontTypeNames.FONTTYPE_INFOIAO)
                                Exit Sub
    
                            End If
                        
926                         For S = 1 To NUMSKILLS
928                             .Stats.UserSkills(S) = 0
930                         Next S
                        
                            Dim SkillLibres As Integer
                        
932                         SkillLibres = 5
934                         SkillLibres = SkillLibres + (5 * .Stats.ELV)
                         
936                         .Stats.SkillPts = SkillLibres
938                         Call WriteLevelUp(UserIndex, .Stats.SkillPts)
                        
940                         Call WriteConsoleMsg(UserIndex, "Tus skills han sido reseteados.", e_FontTypeNames.FONTTYPE_INFOIAO)
942                         Call QuitarUserInvItem(UserIndex, Slot, 1)
    
                        ' Mochila
944                     Case 20
                    
946                         If .Stats.InventLevel < INVENTORY_EXTRA_ROWS Then
948                             .Stats.InventLevel = .Stats.InventLevel + 1
950                             .CurrentInventorySlots = getMaxInventorySlots(UserIndex)
952                             Call WriteInventoryUnlockSlots(UserIndex)
954                             Call WriteConsoleMsg(UserIndex, "Has aumentado el espacio de tu inventario!", e_FontTypeNames.FONTTYPE_INFO)
956                             Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Else
958                             Call WriteConsoleMsg(UserIndex, "Ya has desbloqueado todos los casilleros disponibles.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
    
                            End If
                            
                        ' Poción negra (suicidio)
960                     Case 21
                            'Quitamos del inv el item
962                         Call QuitarUserInvItem(UserIndex, Slot, 1)
                            
964                         If obj.Snd1 <> 0 Then
966                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                            Else
968                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
                            End If

970                         Call WriteConsoleMsg(UserIndex, "Te has suicidado.", e_FontTypeNames.FONTTYPE_EJECUCION)
                            Call CustomScenarios.UserDie(UserIndex)
972                         Call UserMod.UserDie(UserIndex)
                        'Poción de reset (resetea el personaje)
                        Case 22
                            If GetTickCount - .Counters.LastResetTick > 3000 Then
                                Call writeAnswerReset(UserIndex)
                                .Counters.LastResetTick = GetTickCount
                            Else
                                Call WriteConsoleMsg(UserIndex, "Debes esperar unos momentos para tomar esta poción.", e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        Case 23
                             If obj.ApplyEffectId > 0 Then
                                Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
                            End If
                            Call UpdateCd(UserIndex, ObjData(objIndex).cdType)
                            'Quitamos del inv el item
                            Call QuitarUserInvItem(UserIndex, Slot, 1)
                            Call UpdateUserInv(False, UserIndex, Slot)
                            Exit Sub
                    End Select
                    If obj.ApplyEffectId > 0 Then
                        Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
                    End If
974                 Call WriteUpdateUserStats(UserIndex)
976                 Call UpdateUserInv(False, UserIndex, Slot)
    
978             Case e_OBJType.otBebidas
    
980                 If .flags.Muerto = 1 Then
982                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
    
                    End If
    
984                 .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
    
986                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
990                 Call WriteUpdateHungerAndThirst(UserIndex)
            
                    'Quitamos del inv el item
992                 Call QuitarUserInvItem(UserIndex, Slot, 1)
                    If obj.ApplyEffectId > 0 Then
                        Call AddOrResetEffect(UserIndex, obj.ApplyEffectId)
                    End If
            
994                 If obj.Snd1 <> 0 Then
996                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                
                    Else
998                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.y))
    
                      End If
            
1000                 Call UpdateUserInv(False, UserIndex, Slot)
            
1002             Case e_OBJType.OtCofre
    
1004                 If .flags.Muerto = 1 Then
1006                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                            'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
    
                        End If
    
                        'Quitamos del inv el item
1008                 Call QuitarUserInvItem(UserIndex, Slot, 1)
1010                 Call UpdateUserInv(False, UserIndex, Slot)
            
1012                 Call WriteConsoleMsg(UserIndex, "Has abierto un " & obj.Name & " y obtuviste...", e_FontTypeNames.FONTTYPE_New_DONADOR)
            
1014                 If obj.Snd1 <> 0 Then
1016                     Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
                          End If
            
1018                 If obj.CreaFX <> 0 Then
                        UserList(UserIndex).Counters.timeFx = 3
1020                    Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, obj.CreaFX, 0, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y))
                     End If
            
                          Dim i As Byte
    
1022                Select Case obj.Subtipo

                        Case 1
    
1024                             For i = 1 To obj.CantItem
    
1026                                If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
                                    
1028                                     If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
1030                                         Call TirarItemAlPiso(.Pos, obj.Item(i))
                                         End If
                                    
                                     End If
                                
1032                                Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).Name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
    
1034                             Next i
                
                        Case 2
            
1036                             For i = 1 To obj.CantEntrega
        
                                      Dim indexobj As Byte
1038                                    indexobj = RandomNumber(1, obj.CantItem)
                    
                                      Dim Index As t_Obj
        
1040                                 Index.ObjIndex = obj.Item(indexobj).ObjIndex
1042                                 Index.amount = obj.Item(indexobj).amount
        
1044                                 If Not MeterItemEnInventario(UserIndex, Index) Then
    
1046                                    If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
1048                                         Call TirarItemAlPiso(.Pos, Index)
                                         End If
                                    
                                      End If
    
1050                                 Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(Index.ObjIndex).Name & " (" & Index.amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
1052                             Next i
    
                        Case 3
                        
                            For i = 1 To obj.CantItem
                            
                                If RandomNumber(1, obj.Item(i).Data) = 1 Then
                            
                                    If Not MeterItemEnInventario(UserIndex, obj.Item(i)) Then
                                    
                                        If (.flags.Privilegios And (e_PlayerType.user Or e_PlayerType.Dios Or e_PlayerType.Admin)) Then
                                            Call TirarItemAlPiso(.Pos, obj.Item(i))
                                        End If
                                    
                                    End If
                                    
                                    Call SendData(SendTarget.ToIndex, UserIndex, PrepareMessageConsoleMsg(ObjData(obj.Item(i).ObjIndex).Name & " (" & obj.Item(i).amount & ")", e_FontTypeNames.FONTTYPE_INFOBOLD))
                                    
                                End If
                            
                            Next i
    
                    End Select
        
1054             Case e_OBJType.otLlaves
                    If UserList(UserIndex).flags.Muerto = 1 Then
                        Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
                    TargObj = ObjData(UserList(UserIndex).flags.TargetObj)
                    '¿El objeto clickeado es una puerta?
                    If TargObj.OBJType = e_OBJType.otPuertas Then
                        If TargObj.clave < 1000 Then
                            Call WriteConsoleMsg(UserIndex, "Las llaves en el inventario están desactivadas. Sólo se permiten en el llavero.", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        '¿Esta cerrada?
                        If TargObj.Cerrada = 1 Then
                              '¿Cerrada con llave?
                              If TargObj.Llave > 0 Then
                                 If TargObj.clave = obj.clave Then
                                    MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                                    = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerrada
                                    UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                                    Call WriteConsoleMsg(UserIndex, "Has abierto la puerta.", e_FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                 Else
                                    Call WriteConsoleMsg(UserIndex, "La llave no sirve.", e_FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                 End If
                              Else
                                 If TargObj.clave = obj.clave Then
                                    MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex _
                                    = ObjData(MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex).IndexCerradaLlave
                                    Call WriteConsoleMsg(UserIndex, "Has cerrado con llave la puerta.", e_FontTypeNames.FONTTYPE_INFO)
                                    UserList(UserIndex).flags.TargetObj = MapData(UserList(UserIndex).flags.TargetObjMap, UserList(UserIndex).flags.TargetObjX, UserList(UserIndex).flags.TargetObjY).ObjInfo.ObjIndex
                                    Exit Sub
                                 Else
                                    Call WriteConsoleMsg(UserIndex, "La llave no sirve.", e_FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                 End If
                              End If
                        Else
                              Call WriteConsoleMsg(UserIndex, "No esta cerrada.", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
                        End If
                    End If
                    
1058             Case e_OBJType.otBotellaVacia
    
1060                If .flags.Muerto = 1 Then
1062                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    If Not InMapBounds(.flags.TargetMap, .flags.TargetX, .flags.TargetY) Then
                        Exit Sub
                    End If
                    
1064                 If (MapData(.Pos.Map, .flags.TargetX, .flags.TargetY).Blocked And FLAG_AGUA) = 0 Then
1066                     Call WriteConsoleMsg(UserIndex, "No hay agua allí.", e_FontTypeNames.FONTTYPE_INFO)
                         Exit Sub
                    End If
                    
                    If Distance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, .flags.TargetX, .flags.TargetY) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "Debes acercarte más al agua.", e_FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
    
1068                 MiObj.amount = 1
1070                 MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexAbierta

1072                 Call QuitarUserInvItem(UserIndex, Slot, 1)
    
1074                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
1076                     Call TirarItemAlPiso(.Pos, MiObj)
                          End If
            
1078                 Call UpdateUserInv(False, UserIndex, Slot)
        
1080             Case e_OBJType.otBotellaLlena
    
1082                 If .flags.Muerto = 1 Then
1084                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                              ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
    
1086                 .Stats.MinAGU = .Stats.MinAGU + obj.MinSed
    
1088                 If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
1092                 Call WriteUpdateHungerAndThirst(UserIndex)
1094                 MiObj.amount = 1
1096                 MiObj.ObjIndex = ObjData(.Invent.Object(Slot).ObjIndex).IndexCerrada
1098                 Call QuitarUserInvItem(UserIndex, Slot, 1)
    
1100                 If Not MeterItemEnInventario(UserIndex, MiObj) Then
1102                     Call TirarItemAlPiso(.Pos, MiObj)
    
                          End If
            
1104                 Call UpdateUserInv(False, UserIndex, Slot)
        
1106             Case e_OBJType.otPergaminos
    
1108                 If .flags.Muerto = 1 Then
1110                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                              ' Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
            
                          'Call LogError(.Name & " intento aprender el hechizo " & ObjData(.Invent.Object(slot).ObjIndex).HechizoIndex)
            
1112                 If ClasePuedeUsarItem(UserIndex, .invent.Object(Slot).objIndex, Slot) And RazaPuedeUsarItem(UserIndex, .invent.Object(Slot).objIndex, Slot) Then
    
                              'If .Stats.MaxMAN > 0 Then
1114                     If .Stats.MinHam > 0 And .Stats.MinAGU > 0 Then
1116                         Call AgregarHechizo(UserIndex, Slot)
1118                         Call UpdateUserInv(False, UserIndex, Slot)
                                  ' Call LogError(.Name & " lo aprendio.")
                        Else
1120                         Call WriteConsoleMsg(UserIndex, "Estas demasiado hambriento y sediento.", e_FontTypeNames.FONTTYPE_INFO)
    
                        End If
    
                              ' Else
                              '    Call WriteConsoleMsg(UserIndex, "No tienes conocimientos de las Artes Arcanas.", e_FontTypeNames.FONTTYPE_WARNING)
                              'End If
                          Else
                 
1122                     Call WriteConsoleMsg(UserIndex, "Por mas que lo intentas, no podés comprender el manuescrito.", e_FontTypeNames.FONTTYPE_INFO)
       
                          End If
            
1124             Case e_OBJType.otMinerales
    
1126                 If .flags.Muerto = 1 Then
1128                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                              'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
    
1130                 Call WriteWorkRequestTarget(UserIndex, FundirMetal)
           
1132             Case e_OBJType.otInstrumentos
    
1134                 If .flags.Muerto = 1 Then
1136                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                              'Call WriteConsoleMsg(UserIndex, "¡¡Estas muerto!! Solo podes usar items cuando estas vivo. ", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
    
                          End If
            
1138                 If obj.Real Then '¿Es el Cuerno Real?
1140                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1142                         If MapInfo(.Pos.Map).Seguro = 1 Or Zona(.ZonaId).Segura = 1 Then
1144                             Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", e_FontTypeNames.FONTTYPE_INFO)
                                      Exit Sub
    
                                  End If
    
1146                         Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                  Exit Sub
                              Else
1148                         Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Armada Real pueden usar este cuerno.", e_FontTypeNames.FONTTYPE_INFO)
                                  Exit Sub
    
                              End If
    
1150                 ElseIf obj.Caos Then '¿Es el Cuerno Legión?
    
1152                     If FaccionPuedeUsarItem(UserIndex, ObjIndex) Then
1154                         If MapInfo(.Pos.Map).Seguro = 1 Or Zona(.ZonaId).Segura = 1 Then
1156                             Call WriteConsoleMsg(UserIndex, "No hay Peligro aquí. Es Zona Segura ", e_FontTypeNames.FONTTYPE_INFO)
                                      Exit Sub
    
                                  End If
    
1158                         Call SendData(SendTarget.toMap, .Pos.Map, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.Y))
                                  Exit Sub
                              Else
1160                         Call WriteConsoleMsg(UserIndex, "Solo Miembros de la Legión Oscura pueden usar este cuerno.", e_FontTypeNames.FONTTYPE_INFO)
                                  Exit Sub
    
                              End If
    
                          End If
    
                          'Si llega aca es porque es o Laud o Tambor o Flauta
1162                 Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessagePlayWave(obj.Snd1, .Pos.X, .Pos.y))
           
1164             Case e_OBJType.otBarcos
                
                        ' Piratas y trabajadores navegan al nivel 23
                     If .invent.Object(Slot).objIndex <> iObjTrajeAltoNw And .invent.Object(Slot).objIndex <> iObjTrajeBajoNw And .invent.Object(Slot).objIndex <> iObjTraje Then
1166                     If .clase = e_Class.Trabajador Or .clase = e_Class.Pirat Then
1168                         If .Stats.ELV < 23 Then
1170                             Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 23 o superior.", e_FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub
                                End If
                        ' Nivel mínimo 25 para navegar, si no sos pirata ni trabajador
1172                    ElseIf .Stats.ELV < 25 Then
1174                        Call WriteConsoleMsg(UserIndex, "Para recorrer los mares debes ser nivel 25 o superior.", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    ElseIf .invent.Object(Slot).objIndex = iObjTrajeAltoNw Or .invent.Object(Slot).objIndex = iObjTrajeBajoNw Then
                        If (.flags.Navegando = 0 Or (.invent.BarcoObjIndex <> iObjTrajeAltoNw And .invent.BarcoObjIndex <> iObjTrajeBajoNw)) And MapData(.pos.Map, .pos.X + 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.X - 1, .pos.y).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.X, .pos.y + 1).trigger <> e_Trigger.DETALLEAGUA And MapData(.pos.Map, .pos.X, .pos.y - 1).trigger <> e_Trigger.DETALLEAGUA Then
                            Call WriteConsoleMsg(UserIndex, "Este traje es para aguas contaminadas.", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    ElseIf .invent.Object(Slot).objIndex = iObjTraje Then
                          If (.flags.Navegando = 0 Or .invent.BarcoObjIndex <> iObjTraje) And MapData(.pos.Map, .pos.X + 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.X - 1, .pos.y).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.X, .pos.y + 1).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.X, .pos.y - 1).trigger <> e_Trigger.NADOCOMBINADO And MapData(.pos.Map, .pos.X + 1, .pos.y).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.X - 1, .pos.y).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.X, .pos.y + 1).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.X, .pos.y - 1).trigger <> e_Trigger.VALIDONADO And MapData(.pos.Map, .pos.X + 1, .pos.y).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.X - 1, .pos.y).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.X, .pos.y + 1).trigger <> e_Trigger.NADOBAJOTECHO And MapData(.pos.Map, .pos.X, .pos.y - 1).trigger <> e_Trigger.NADOBAJOTECHO Then
                            Call WriteConsoleMsg(UserIndex, "Este traje es para zonas poco profundas.", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    End If
                    

1176                If .flags.Navegando = 0 Then
1178                    If LegalWalk(.Pos.Map, .Pos.X - 1, .Pos.Y, e_Heading.WEST, True, False) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y - 1, e_Heading.NORTH, True, False) Or LegalWalk(.Pos.Map, .Pos.X + 1, .Pos.Y, e_Heading.EAST, True, False) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y + 1, e_Heading.SOUTH, True, False) Then
1180                        Call DoNavega(UserIndex, obj, Slot)
                        Else
1182                        Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte al agua para usar el barco o traje de baño!", e_FontTypeNames.FONTTYPE_INFO)
                        End If
                    
                    Else
1184                    If .invent.BarcoObjIndex <> .invent.Object(Slot).ObjIndex Then
1186                        Call DoNavega(UserIndex, obj, Slot)
                        Else
1188                        If LegalWalk(.Pos.Map, .Pos.X - 1, .Pos.Y, e_Heading.WEST, False, True) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y - 1, e_Heading.NORTH, False, True) Or LegalWalk(.Pos.Map, .Pos.X + 1, .Pos.Y, e_Heading.EAST, False, True) Or LegalWalk(.Pos.Map, .Pos.X, .Pos.Y + 1, e_Heading.SOUTH, False, True) Then
1190                            Call DoNavega(UserIndex, obj, Slot)
                            Else
1192                            Call WriteConsoleMsg(UserIndex, "¡Debes aproximarte a la costa para dejar la barca!", e_FontTypeNames.FONTTYPE_INFO)
                            End If
                        End If
                    End If
            
1194             Case e_OBJType.otMonturas
                          'Verifica todo lo que requiere la montura
        
1196                If .flags.Muerto = 1 Then
1198                    Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                           'Call WriteConsoleMsg(UserIndex, "¡Estas muerto! Los fantasmas no pueden montar.", e_FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
                
1200                If .flags.Navegando = 1 Then
1202                    Call WriteConsoleMsg(UserIndex, "Debes dejar de navegar para poder cabalgar.", e_FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
    
1204                If MapInfo(.Pos.Map).zone = "DUNGEON" Then
1206                    Call WriteConsoleMsg(UserIndex, "No podes cabalgar dentro de un dungeon.", e_FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1208                Call DoMontar(UserIndex, obj, Slot)
                
                 Case e_OBJType.OtDonador
                    Select Case obj.Subtipo
                        Case 1
1214                        If .Counters.Pena <> 0 Then
1216                            Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
1218                        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
1220                            Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", e_FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If

1222                         Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1224                         Call WriteConsoleMsg(UserIndex, "Has viajado por el mundo.", e_FontTypeNames.FONTTYPE_WARNING)
1226                         Call QuitarUserInvItem(UserIndex, Slot, 1)
1228                         Call UpdateUserInv(False, UserIndex, Slot)

1230                     Case 2
                            Exit Sub
1252                     Case 3
                            Exit Sub
                    End Select
1262             Case e_OBJType.otpasajes
    
1264                 If .flags.Muerto = 1 Then
1266                     Call WriteLocaleMsg(UserIndex, "77", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
                          End If
            
1268                 If .flags.TargetNpcTipo <> Pirata Then
1270                     Call WriteConsoleMsg(UserIndex, "Primero debes hacer click sobre el pirata.", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
                          End If
            
1272                 If Distancia(NpcList(.flags.TargetNPC.ArrayIndex).Pos, .Pos) > 3 Then
1274                     Call WriteLocaleMsg(UserIndex, "8", e_FontTypeNames.FONTTYPE_INFO)
                              Exit Sub
                          End If
            
1276                 If .Pos.Map <> obj.DesdeMap Then
1278                     Call WriteChatOverHead(UserIndex, "El pasaje no lo compraste aquí! Largate!", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)
                              Exit Sub
                          End If
            
1280                 If Not MapaValido(obj.HastaMap) Then
1282                     Call WriteChatOverHead(UserIndex, "El pasaje lleva hacia un mapa que ya no esta disponible! Disculpa las molestias.", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)
                              Exit Sub
                          End If
    
1284                 If obj.NecesitaNave > 0 Then
1286                     If .Stats.UserSkills(e_Skill.Navegacion) < 80 Then
1288                         Call WriteChatOverHead(UserIndex, "Debido a la peligrosidad del viaje, no puedo llevarte, ya que al menos necesitas saber manejar una barca.", str$(NpcList(.flags.TargetNPC.ArrayIndex).Char.charindex), vbWhite)
                                  Exit Sub
                              End If
                          End If
                
1290                 Call WarpUserChar(UserIndex, obj.HastaMap, obj.HastaX, obj.HastaY, True)
1292                 Call WriteConsoleMsg(UserIndex, "Has viajado por varios días, te sientes exhausto!", e_FontTypeNames.FONTTYPE_WARNING)
1294                 .Stats.MinAGU = 0
1296                 .Stats.MinHam = 0
1302                 Call WriteUpdateHungerAndThirst(UserIndex)
1304                 Call QuitarUserInvItem(UserIndex, Slot, 1)
1306                 Call UpdateUserInv(False, UserIndex, Slot)
            
1308             Case e_OBJType.otRunas
        
1310                If .Counters.Pena <> 0 Then
1312                    Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", e_FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1314                If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
1316                    Call WriteConsoleMsg(UserIndex, "No podes usar la runa estando en la carcel.", e_FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
                        
1318                If MapInfo(.Pos.Map).Seguro = 0 Or Zona(.ZonaId).Segura = 0 And .flags.Muerto = 0 Then
1320                    Call WriteConsoleMsg(UserIndex, "Solo podes usar tu runa en zonas seguras.", e_FontTypeNames.FONTTYPE_INFO)
                           Exit Sub
    
                       End If
            
1322                If .Accion.AccionPendiente Then
                           Exit Sub
    
                       End If
            
1324                 Select Case ObjData(ObjIndex).TipoRuna
            
                              Case 1, 2
    
1326                         If Not EsGM(UserIndex) Then
1328                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, 400, False))
1330                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 350, e_AccionBarra.Runa))
                                  Else
1332                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageParticleFX(.Char.charindex, e_ParticulasIndex.Runa, 50, False))
1334                             Call SendData(SendTarget.toPCAliveArea, UserIndex, PrepareMessageBarFx(.Char.charindex, 100, e_AccionBarra.Runa))
    
                                  End If
    
1336                         .Accion.Particula = e_ParticulasIndex.Runa
1338                         .Accion.AccionPendiente = True
1340                         .Accion.TipoAccion = e_AccionBarra.Runa
1342                         .Accion.RunaObj = ObjIndex
1344                         .Accion.ObjSlot = Slot
        
                          End Select
            
1346             Case e_OBJType.otmapa
1348                 Call WriteShowFrmMapa(UserIndex)
                 Case e_OBJType.OtQuest
1349                 If obj.QuestId > 0 Then Call WriteObjQuestSend(UserIndex, obj.QuestId, Slot)
                 Case e_OBJType.otMagicos
                        Select Case ObjData(objIndex).Subtipo
                            Case e_MagicItemSubType.TargetUsable
                                Call WriteWorkRequestTarget(UserIndex, e_Skill.TargetableItem)
                        End Select
                        Select Case ObjData(ObjIndex).EfectoMagico
                            Case e_MagicItemEffect.eProtectedResources
                                If ObjData(ObjIndex).ApplyEffectId <= 0 Then
                                    Exit Sub
                                End If
                                Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
                                Call AddOrResetEffect(UserIndex, ObjData(ObjIndex).ApplyEffectId)
                        End Select
                 Case e_OBJType.otUsableOntarget
                    .flags.UsingItemSlot = .flags.TargetObjInvSlot
                    Call WriteWorkRequestTarget(UserIndex, e_Skill.TargetableItem)
                End Select
             End With

             Exit Sub

hErr:
1350    LogError "Error en useinvitem Usuario: " & UserList(UserIndex).Name & " item:" & obj.Name & " index: " & UserList(UserIndex).Invent.Object(Slot).ObjIndex

End Sub

Sub EnivarArmasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmasConstruibles_Err
        

100     Call WriteBlacksmithWeapons(UserIndex)

        
        Exit Sub

EnivarArmasConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmasConstruibles", Erl)

        
End Sub
 
Sub EnivarObjConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruibles_Err
        

100     Call WriteCarpenterObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruibles", Erl)

        
End Sub

Sub EnivarObjConstruiblesAlquimia(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesAlquimia_Err
        

100     Call WriteAlquimistaObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesAlquimia_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesAlquimia", Erl)

        
End Sub

Sub EnivarObjConstruiblesSastre(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarObjConstruiblesSastre_Err
        

100     Call WriteSastreObjects(UserIndex)

        
        Exit Sub

EnivarObjConstruiblesSastre_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarObjConstruiblesSastre", Erl)

        
End Sub

Sub EnivarArmadurasConstruibles(ByVal UserIndex As Integer)
        
        On Error GoTo EnivarArmadurasConstruibles_Err
        

100     Call WriteBlacksmithArmors(UserIndex)

        
        Exit Sub

EnivarArmadurasConstruibles_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.EnivarArmadurasConstruibles", Erl)

        
End Sub

Public Function ItemSeCae(ByVal Index As Integer) As Boolean
        
        On Error GoTo ItemSeCae_Err
        

100     ItemSeCae = (ObjData(Index).Real <> 1 Or ObjData(Index).NoSeCae = 0) And (ObjData(Index).Caos <> 1 Or ObjData(Index).NoSeCae = 0) And ObjData(Index).OBJType <> e_OBJType.otLlaves And ObjData(Index).OBJType <> e_OBJType.otBarcos And ObjData(Index).OBJType <> e_OBJType.otMonturas And ObjData(Index).NoSeCae = 0 And Not ObjData(Index).Intirable = 1 And Not ObjData(Index).Destruye = 1 And Not ObjData(Index).Instransferible = 1

        
        Exit Function

ItemSeCae_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemSeCae", Erl)

        
End Function

Public Function PirataCaeItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

        On Error GoTo PirataCaeItem_Err

100     With UserList(UserIndex)

102         If .clase = e_Class.Pirat And .Stats.ELV >= 37 And .flags.Navegando = 1 Then

                ' Si no está navegando, se caen los items
104             If .Invent.BarcoObjIndex > 0 Then

                    ' Con galeón cada item tiene una probabilidad de caerse del 67%
106                 If ObjData(.Invent.BarcoObjIndex).Ropaje = iGaleon Then

108                     If RandomNumber(1, 100) <= 33 Then
                            Exit Function
                        End If

                    End If

                End If

            End If

        End With

110     PirataCaeItem = True

        Exit Function

PirataCaeItem_Err:
112     Call TraceError(Err.Number, Err.Description, "InvUsuario.PirataCaeItem", Erl)

End Function

Sub TirarTodosLosItems(ByVal UserIndex As Integer)
        
        On Error GoTo TirarTodosLosItems_Err

        Dim i         As Byte
        Dim NuevaPos  As t_WorldPos
        Dim MiObj     As t_Obj
        Dim ItemIndex As Integer
       
100     With UserList(UserIndex)

            
            If ((.Pos.map = 58 Or .Pos.map = 59 Or .Pos.map = 60 Or .Pos.map = 61) And EnEventoFaccionario) Then Exit Sub
            ' Tambien se cae el oro de la billetera
            Dim GoldToDrop As Long
            GoldToDrop = .Stats.GLD - (OroPorNivelBilletera * .Stats.ELV)
102         If GoldToDrop > 0 And Not EsGM(UserIndex) Then
104             Call TirarOro(GoldToDrop, UserIndex)
            End If
            
106         For i = 1 To .CurrentInventorySlots
    
108             ItemIndex = .Invent.Object(i).ObjIndex

110             If ItemIndex > 0 Then

112                 If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
114                     NuevaPos.X = 0
116                     NuevaPos.Y = 0
                    
118                     MiObj.amount = DropAmmount(.invent, i)
120                     MiObj.ObjIndex = ItemIndex
                        
                        If .flags.Navegando Then
128                         Call Tilelibre(.Pos, NuevaPos, MiObj, True, True)
                        Else
129                         Call Tilelibre(.Pos, NuevaPos, MiObj, .flags.Navegando = True, (Not .flags.Navegando) = True)
                            Call ClosestLegalPos(.Pos, NuevaPos, .flags.Navegando, Not .flags.Navegando)
                        End If
130                     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
132                         Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                        
                        ' WyroX: Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                        Else
134                         Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
136                         Call UpdateUserInv(False, UserIndex, i)
                        End If
                
                    End If

                End If
    
138         Next i
    
        End With
 
        Exit Sub

TirarTodosLosItems_Err:
140     Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarTodosLosItems", Erl)
End Sub

Function DropAmmount(ByRef invent As t_Inventario, ByVal objectIndex As Integer) As Integer
100 DropAmmount = invent.Object(objectIndex).amount
102 If invent.MagicoObjIndex > 0 Then
        With ObjData(invent.MagicoObjIndex)
104     If .EfectoMagico = 12 Then
            Dim unprotected As Single
            unprotected = 1
106         If invent.Object(objectIndex).ObjIndex = ORO_MINA Then 'ore types
108             unprotected = CSng(1) - (CSng(.LingO) / 100)
110         ElseIf invent.Object(objectIndex).ObjIndex = PLATA_MINA Then
112             unprotected = CSng(1) - (CSng(.LingP) / 100)
114         ElseIf invent.Object(objectIndex).ObjIndex = HIERRO_MINA Then
116             unprotected = CSng(1) - (CSng(.LingH) / 100)
118         ElseIf invent.Object(objectIndex).ObjIndex = Wood Then ' wood types
120             unprotected = CSng(1) - (CSng(.Madera) / 100)
122         ElseIf invent.Object(objectIndex).ObjIndex = ElvenWood Then
124             unprotected = CSng(1) - (CSng(.MaderaElfica) / 100)
129         ElseIf invent.Object(objectIndex).objIndex = PinoWood Then
130             unprotected = CSng(1) - (CSng(.MaderaPino) / 100)
131         ElseIf invent.Object(objectIndex).objIndex = BLODIUM_MINA Then
132             unprotected = CSng(1) - (CSng(.Blodium) / 100)
            ElseIf invent.Object(objectIndex).ObjIndex > 0 Then 'fish types
                If ObjData(invent.Object(objectIndex).ObjIndex).OBJType = otUseOnce And _
                   ObjData(invent.Object(objectIndex).ObjIndex).Subtipo = e_UseOnceSubType.eFish Then
                   unprotected = CSng(1) - (CSng(.MaxItems) / 100)
                End If
            End If
            DropAmmount = Int(DropAmmount * unprotected)
        End If
        End With
    End If
End Function

Function ItemNewbie(ByVal ItemIndex As Integer) As Boolean
        
        On Error GoTo ItemNewbie_Err
        
100     ItemNewbie = ObjData(ItemIndex).Newbie = 1

        Exit Function

ItemNewbie_Err:
102     Call TraceError(Err.Number, Err.Description, "InvUsuario.ItemNewbie", Erl)
End Function

Public Function IsItemInCooldown(ByRef User As t_User, ByRef obj As t_UserOBJ) As Boolean
    Dim elapsedTime As Long
    ElapsedTime = GetTickCount() - User.CdTimes(ObjData(obj.objIndex).cdType)
    IsItemInCooldown = ElapsedTime < ObjData(obj.objIndex).Cooldown
End Function

Public Sub UserTargetableItem(ByVal UserIndex As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
On Error GoTo UserTargetableItem_Err
    With UserList(UserIndex)
        If IsItemInCooldown(UserList(UserIndex), .invent.Object(.flags.UsingItemSlot)) Then
            Exit Sub
        End If
        If .flags.UsingItemSlot = 0 Then Exit Sub
        Dim objIndex As Integer
        ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
        With ObjData(ObjIndex)
            If .MinHp > UserList(UserIndex).Stats.MinHp Then
                Call WriteLocaleMsg(UserIndex, MsgRequiresMoreHealth, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If .MinSta > UserList(UserIndex).Stats.MinSta Then
                Call WriteLocaleMsg(UserIndex, MsgTiredToPerformAction, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            Select Case .Subtipo
                Case e_UssableOnTarget.eRessurectionItem
                    Call ResurrectWithItem(UserIndex)
                Case e_UssableOnTarget.eTrap
                    Call PlaceTrap(UserIndex, TileX, TileY)
                Case e_UssableOnTarget.eArpon
                    Call UseArpon(UserIndex)
                Case e_UssableOnTarget.eHandCannon
                    Call UseHandCannon(UserIndex, TileX, TileY)
            End Select
        End With
        .flags.UsingItemSlot = 0
    End With
    Exit Sub
UserTargetableItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.UserTargetableItem", Erl)
End Sub

Public Sub ResurrectWithItem(ByVal UserIndex As Integer)
On Error GoTo ResurrectWithItem_Err
    With UserList(UserIndex)
100     Dim CanHelpResult As e_InteractionResult
102     If Not IsValidUserRef(.flags.TargetUser) Then
104         Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If .flags.TargetUser.ArrayIndex = UserIndex Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
114     Dim TargetUser As Integer
116     TargetUser = .flags.TargetUser.ArrayIndex
        If UserList(TargetUser).flags.Muerto = 0 Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
106     CanHelpResult = UserMod.CanHelpUser(UserIndex, targetUser)
        If UserList(TargetUser).flags.SeguroResu Then
            Call WriteConsoleMsg(UserIndex, "El usuario tiene el seguro de resurrección activado.", e_FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(TargetUser, UserList(UserIndex).name & " está intentando revivirte. Desactiva el seguro de resurrección para permitirle hacerlo.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If CanHelpResult <> eInteractionOk Then
            Call SendHelpInteractionMessage(UserIndex, CanHelpResult)
        End If
        
118     Dim costoVidaResu As Long
120     costoVidaResu = UserList(TargetUser).Stats.ELV * 1.5 + .Stats.MinHp * 0.5
    
122     Call UserMod.ModifyHealth(UserIndex, -costoVidaResu, 1)
124     Call ModifyStamina(UserIndex, -UserList(UserIndex).Stats.MinSta, False, 0)
        Dim objIndex As Integer
126     ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
128     Call UpdateCd(UserIndex, ObjData(objIndex).cdType)
192     Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.UsingItemSlot)
196     Call ResurrectUser(TargetUser)
        If IsFeatureEnabled("remove-inv-on-attack") Then
            Call RemoveUserInvisibility(UserIndex)
        End If
    End With
    Exit Sub
ResurrectWithItem_Err:
    Call TraceError(Err.Number, Err.Description, "InvUsuario.ResurrectWithItem", Erl)
End Sub

Public Sub RemoveItemFromInventory(ByVal UserIndex As Integer, ByVal Slot As Integer)
    Call QuitarUserInvItem(UserIndex, Slot, 1)
    Call UpdateUserInv(True, UserIndex, Slot)
End Sub

Public Sub PlaceTrap(ByVal UserIndex As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    With UserList(UserIndex)
        If Distance(TileX, TileY, .pos.x, .pos.y) > 3 Then
            Call WriteLocaleMsg(UserIndex, MsgToFar, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If Not CanAddTrapAt(.pos.map, TileX, TileY) Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTile, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim i As Integer
        Dim OlderTrapTime As Long
        Dim OlderTrapIndex As Integer
        OlderTrapTime = 0
        Dim TrapCount As Integer
        Dim Trap As clsTrap
        For i = 0 To .EffectOverTime.EffectCount - 1
            If .EffectOverTime.EffectList(i).TypeId = e_EffectOverTimeType.eTrap Then
                TrapCount = TrapCount + 1
                Set Trap = .EffectOverTime.EffectList(i)
                If Trap.ElapsedTime > OlderTrapTime Then
                    OlderTrapIndex = i
                    OlderTrapTime = Trap.ElapsedTime
                End If
            End If
        Next i
        If TrapCount >= 3 Then
            Set Trap = .EffectOverTime.EffectList(OlderTrapIndex)
            Call Trap.Disable
        End If
        Dim objIndex As Integer
        ObjIndex = UserList(UserIndex).invent.Object(UserList(UserIndex).flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(objIndex).cdType)
        Call EffectsOverTime.CreateTrap(UserIndex, eUser, .pos.map, TileX, TileY, ObjData(objIndex).EfectoMagico)
        Call RemoveItemFromInventory(UserIndex, UserList(UserIndex).flags.UsingItemSlot)
    End With
End Sub

Public Sub UseArpon(ByVal UserIndex As Integer)
    With UserList(UserIndex)
100     Dim CanAttackResult As e_AttackInteractionResult
        Dim TargetRef As t_AnyReference
        If IsValidUserRef(.flags.targetUser) Then
            Call CastUserToAnyRef(.flags.targetUser, TargetRef)
        Else
            Call CastNpcToAnyRef(.flags.TargetNPC, TargetRef)
        End If
102     If Not IsValidRef(TargetRef) Then
104         Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If TargetRef.RefType = eUser Then
            If UserList(TargetRef.ArrayIndex).flags.Muerto <> 0 Then
                Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            If TargetRef.RefType = eUser And TargetRef.ArrayIndex = UserIndex Then
                Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        End If
        CanAttackResult = UserCanAttack(UserIndex, UserList(UserIndex).VersionId, TargetRef)
        If CanAttackResult <> e_AttackInteractionResult.eCanAttack Then
            Call WriteLocaleMsg(UserIndex, MsgInvalidTarget, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim ObjIndex As Integer
        ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
        Dim Damage As Integer
        Damage = GetUserDamageWithItem(UserIndex, ObjIndex, 0)
        If TargetRef.RefType = eUser Then
            UserList(TargetRef.ArrayIndex).Counters.timeFx = 3
            Call RemoveUserInvisibility(UserIndex)
            Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessageCreateFX(UserList(TargetRef.ArrayIndex).Char.charindex, FXSANGRE, 0, UserList(TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
            Call SendData(SendTarget.ToPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO, UserList(TargetRef.ArrayIndex).pos.x, UserList(TargetRef.ArrayIndex).pos.y))
        Else
            If NpcList(TargetRef.ArrayIndex).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(NpcList(TargetRef.ArrayIndex).flags.Snd2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))
            Else
                Call SendData(SendTarget.ToNPCAliveArea, TargetRef.ArrayIndex, PrepareMessagePlayWave(SND_IMPACTO2, NpcList(TargetRef.ArrayIndex).pos.x, NpcList(TargetRef.ArrayIndex).pos.y))
            End If
        End If
        If DoDamageToTarget(UserIndex, TargetRef, Damage, e_phisical, ObjIndex) = eStillAlive Then
            If TargetRef.RefType = eUser Then
                If Not IsSet(UserList(TargetRef.ArrayIndex).flags.StatusMask, eCCInmunity) Then
                    Call CreateEffect(UserIndex, eUser, TargetRef.ArrayIndex, TargetRef.RefType, ObjData(ObjIndex).ApplyEffectId)
                End If
            End If
        End If
        
        If .flags.Oculto = 0 Then
            Dim TargetPos As t_WorldPos
            TargetPos = GetPosition(TargetRef)
            Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(.pos.x, .pos.y, TargetPos.x, TargetPos.y, ObjData(ObjIndex).ProjectileType))
        End If
    End With
End Sub

Public Sub UseHandCannon(ByVal UserIndex As Integer, ByVal TileX As Integer, ByVal TileY As Integer)
    With UserList(UserIndex)
        If Distance(TileX, TileY, .pos.x, .pos.y) > 10 Then
            Call WriteLocaleMsg(UserIndex, MsgToFar, e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        Dim ObjIndex As Integer
        ObjIndex = .invent.Object(.flags.UsingItemSlot).ObjIndex
        Call UpdateCd(UserIndex, ObjData(ObjIndex).cdType)
        Dim Particula As Integer
        Dim Tiempo    As Long
        Particula = val(ReadField(1, ObjData(ObjIndex).CreaParticula, Asc(":")))
        Tiempo = val(ReadField(2, ObjData(ObjIndex).CreaParticula, Asc(":")))
        UserList(UserIndex).Counters.timeFx = 3
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessageParticleFX(UserList(UserIndex).Char.charindex, Particula, Tiempo, False, , UserList(UserIndex).pos.x, UserList(UserIndex).pos.y))
        Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareCreateProjectile(.pos.x, .pos.y, TileX, TileY, ObjData(ObjIndex).ProjectileType))
        Call CreateDelayedBlast(UserIndex, eUser, .pos.Map, TileX, TileY, ObjData(ObjIndex).ApplyEffectId, ObjIndex)
        If ObjData(ObjIndex).Snd1 <> 0 Then Call SendData(SendTarget.ToPCAliveArea, UserIndex, PrepareMessagePlayWave(ObjData(ObjIndex).Snd1, .pos.x, .pos.y))
        Call RemoveUserInvisibility(UserIndex)
    End With
End Sub

Public Sub AddOrResetEffect(ByVal UserIndex As Integer, ByVal EffectId As Integer)
    With UserList(UserIndex)
        Dim Effect As IBaseEffectOverTime
        Set Effect = EffectsOverTime.FindEffectOnTarget(UserIndex, .EffectOverTime, EffectId)
        If Effect Is Nothing Then
            Call CreateEffect(UserIndex, eUser, UserIndex, eUser, EffectId)
        Else
            If EffectOverTime(EffectId).Override Then
                Call Effect.Reset(UserIndex, eUser, EffectId)
            End If
        End If
    End With
End Sub

Public Sub UpdateCharWithEquipedItems(ByVal UserIndex As Integer)
    With UserList(UserIndex)
        If .flags.Navegando > 0 Then
            Call EquiparBarco(UserIndex)
            .Char.CascoAnim = 0
            .Char.CartAnim = 0
            .Char.ShieldAnim = 0
            .Char.WeaponAnim = 0
            'TODO place ship body
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
            Exit Sub
        End If
        If .flags.Muerto = 1 Then
            .Char.body = iCuerpoMuerto
204         .Char.head = 0
206         .Char.ShieldAnim = NingunEscudo
208         .Char.WeaponAnim = NingunArma
210         .Char.CascoAnim = NingunCasco
211         .Char.CartAnim = NoCart
            Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
        Exit Sub
        End If
        
        .Char.head = .OrigChar.head
        If .invent.WeaponEqpObjIndex > 0 Then
            .Char.WeaponAnim = ObjData(.invent.WeaponEqpObjIndex).WeaponAnim
        ElseIf .invent.HerramientaEqpObjIndex > 0 Then
            .Char.WeaponAnim = ObjData(.invent.HerramientaEqpObjIndex).WeaponAnim
        Else
            .Char.WeaponAnim = 0
        End If
        If .invent.ArmourEqpObjIndex > 0 Then
            .Char.body = ObtenerRopaje(UserIndex, ObjData(.invent.ArmourEqpObjIndex))
        Else
            Call SetNakedBody(UserList(UserIndex))
        End If
        If .invent.CascoEqpObjIndex > 0 Then
            .Char.CascoAnim = ObjData(.invent.CascoEqpObjIndex).CascoAnim
        Else
            .Char.CascoAnim = 0
        End If
        If .invent.MagicoObjIndex > 0 Then
            .Char.CartAnim = ObjData(.invent.MagicoObjIndex).Ropaje
        Else
            .Char.CartAnim = 0
        End If
        If .invent.EscudoEqpObjIndex > 0 Then
            .Char.ShieldAnim = ObjData(.invent.EscudoEqpObjIndex).ShieldAnim
        Else
            .Char.ShieldAnim = 0
        End If
        Call ChangeUserChar(UserIndex, .Char.body, .Char.head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, UserList(UserIndex).Char.CartAnim)
    End With
End Sub

Function RemoveGold(ByVal UserIndex As Integer, ByVal Amount As Long) As Boolean
    With UserList(UserIndex)
        If .Stats.GLD < Amount Then Exit Function
        .Stats.GLD = .Stats.GLD - Amount
        Call WriteUpdateGold(UserIndex)
        RemoveGold = True
    End With
End Function

Sub AddGold(ByVal UserIndex As Integer, ByVal Amount As Long)
    With UserList(UserIndex)
        .Stats.GLD = .Stats.GLD + Amount
        Call WriteUpdateGold(UserIndex)
    End With
End Sub

Function ObtenerRopaje(ByVal UserIndex As Integer, ByRef Obj As t_ObjData) As Integer
    Dim Race As e_Raza
    Race = UserList(UserIndex).raza
    
    Dim EsMujer As Boolean
    EsMujer = UserList(UserIndex).genero = e_Genero.Mujer

    Select Case Race
        Case e_Raza.Humano
            If EsMujer And Obj.RopajeHumana > 0 Then
                ObtenerRopaje = Obj.RopajeHumana
                Exit Function
            ElseIf Obj.RopajeHumano > 0 Then
                ObtenerRopaje = Obj.RopajeHumano
                Exit Function
            End If
        Case e_Raza.Elfo
            If EsMujer And Obj.RopajeElfa > 0 Then
                ObtenerRopaje = Obj.RopajeElfa
                Exit Function
            ElseIf Obj.RopajeElfo > 0 Then
                ObtenerRopaje = Obj.RopajeElfo
                Exit Function
            End If
        Case e_Raza.Drow
            If EsMujer And Obj.RopajeElfaOscura > 0 Then
                ObtenerRopaje = Obj.RopajeElfaOscura
                Exit Function
            ElseIf Obj.RopajeElfoOscuro > 0 Then
                ObtenerRopaje = Obj.RopajeElfoOscuro
                Exit Function
            End If
        Case e_Raza.Orco
            If EsMujer And Obj.RopajeOrca > 0 Then
                ObtenerRopaje = Obj.RopajeOrca
                Exit Function
            ElseIf Obj.RopajeOrco > 0 Then
                ObtenerRopaje = Obj.RopajeOrco
                Exit Function
            End If
        Case e_Raza.Enano
            If EsMujer And Obj.RopajeEnana > 0 Then
                ObtenerRopaje = Obj.RopajeEnana
                Exit Function
            ElseIf Obj.RopajeEnano > 0 Then
                ObtenerRopaje = Obj.RopajeEnano
                Exit Function
            End If
        Case e_Raza.Gnomo
            If EsMujer And Obj.RopajeGnoma > 0 Then
                ObtenerRopaje = Obj.RopajeGnoma
                Exit Function
            ElseIf Obj.RopajeGnomo > 0 Then
                ObtenerRopaje = Obj.RopajeGnomo
                Exit Function
            End If
    End Select
    
    ObtenerRopaje = Obj.Ropaje
End Function
