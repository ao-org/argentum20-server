Attribute VB_Name = "ModRetos"
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

Private Const APUESTA_MAXIMA = 100000000

Public Retos As t_Retos
Private ListaDeEspera As New Dictionary

Public Sub CargarInfoRetos()
    Dim File As clsIniManager
100 Set File = New clsIniManager

102     Call File.Initialize(DatPath & "Retos.dat")
    
104     With Retos

106         .TamañoMaximoEquipo = val(File.GetValue("Retos", "MaximoEquipo"))
108         .ApuestaMinima = val(File.GetValue("Retos", "ApuestaMinima"))
110         .ImpuestoApuesta = val(File.GetValue("Retos", "ImpuestoApuesta"))
112         .DuracionMaxima = val(File.GetValue("Retos", "DuracionMaxima"))
#If DEBUGGING Then
114         .TiempoConteo = 3
#Else
            .TiempoConteo = val(File.GetValue("Retos", "TiempoConteo"))
#End If
116         .TotalSalas = val(File.GetValue("Salas", "Cantidad"))
        
118         If .TotalSalas <= 0 Then Exit Sub
        
120         ReDim .Salas(1 To .TotalSalas)
        
122         .SalasLibres = .TotalSalas
        
124         .AnchoSala = val(File.GetValue("Salas", "Ancho"))
126         .AltoSala = val(File.GetValue("Salas", "Alto"))
        
            Dim Sala As Integer, SalaStr As String
128         For Sala = 1 To .TotalSalas
130             SalaStr = "Sala" & Sala
            
132             With .Salas(Sala)
134                 .PosIzquierda.Map = val(File.GetValue(SalaStr, "Mapa"))
136                 .PosIzquierda.X = val(File.GetValue(SalaStr, "X"))
138                 .PosIzquierda.Y = val(File.GetValue(SalaStr, "Y"))
                
140                 .PosDerecha.Map = .PosIzquierda.Map
142                 .PosDerecha.X = .PosIzquierda.X + Retos.AnchoSala - 1
144                 .PosDerecha.Y = .PosIzquierda.Y + Retos.AltoSala - 1
                End With
            Next
        
        End With
    
146     Set File = Nothing
End Sub

Public Sub CrearReto(ByVal UserIndex As Integer, JugadoresStr As String, ByVal Apuesta As Long, ByVal PocionesMaximas As Integer, Optional ByVal CaenItems As Boolean = False)
    
        On Error GoTo ErrHandler
    
100     With UserList(UserIndex)

102         If .flags.SolicitudReto.estado <> e_SolicitudRetoEstado.Libre Then
104             Call CancelarSolicitudReto(UserIndex, .Name & " ha cancelado la solicitud.")

106         ElseIf IsValidUserRef(.flags.AceptoReto) Then
108             Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .Name & " ha cancelado su admisión.")
            End If
        
110         Dim TamanoReal As Byte: TamanoReal = Retos.TamañoMaximoEquipo * 2 - 1
        
112         If LenB(JugadoresStr) <= 0 Then Exit Sub
    
114         Dim Jugadores() As String: Jugadores = Split(JugadoresStr, ";", TamanoReal)
        
116         If UBound(Jugadores) > TamanoReal - 1 Or UBound(Jugadores) Mod 2 = 1 Then Exit Sub
        
118         Dim MaxIndexEquipo As Integer: MaxIndexEquipo = UBound(Jugadores) \ 2
    
120         If Apuesta < Retos.ApuestaMinima Or Apuesta > APUESTA_MAXIMA Then
122             Call WriteConsoleMsg(UserIndex, "La apuesta mínima es de " & PonerPuntos(Retos.ApuestaMinima) & " monedas de oro.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

124         If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub

126         If .Stats.GLD < Apuesta Then
128             Call WriteConsoleMsg(UserIndex, "No tienes el oro suficiente.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
130         If PocionesMaximas >= 0 Then
132             If TieneObjetos(38, PocionesMaximas + 1, UserIndex) Then
134                 Call WriteConsoleMsg(UserIndex, "Tienes demasiadas pociones rojas (Cantidad máxima: " & PocionesMaximas & ").", e_FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
        
136         With .flags.SolicitudReto
138             .Apuesta = Apuesta
140             .PocionesMaximas = PocionesMaximas
142             .CaenItems = CaenItems
144             ReDim .Jugadores(0 To UBound(Jugadores))
            
                Dim i As Integer, tIndex As t_UserReference
                Dim Equipo1 As String, Equipo2 As String
            
146             Equipo1 = UserList(UserIndex).Name

148             For i = 0 To UBound(.Jugadores)
150                 With .Jugadores(i)
152                     If EsGmChar(Jugadores(i)) Then
154                         Call WriteConsoleMsg(UserIndex, "¡No puedes jugar retos con administradores!", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If

156                     tIndex = NameIndex(Jugadores(i))
                                                                                
158                     If Not IsValidUserRef(tIndex) Then
160                         Call WriteConsoleMsg(UserIndex, "El usuario " & Jugadores(i) & " no puede jugar un reto en este momento.", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If
                    
162                     If Not PuedeReto(tIndex.ArrayIndex) Then
164                         Call WriteConsoleMsg(UserIndex, "El usuario " & UserList(tIndex.ArrayIndex).name & " no puede jugar un reto en este momento.", e_FontTypeNames.FONTTYPE_INFO)
                            Exit Sub
                        End If

166                     .CurIndex = tIndex
168                     .nombre = UserList(.CurIndex.ArrayIndex).name
170                     .Aceptado = False
                    
172                     If i Mod 2 Then
174                         Equipo1 = Equipo1 & IIf((i + 1) \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                        Else
176                         If LenB(Equipo2) > 0 Then
178                             Equipo2 = Equipo2 & IIf(i \ 2 < MaxIndexEquipo, ", ", " y ") & .nombre
                            Else
180                             Equipo2 = .nombre
                            End If
                        End If
                    End With
                Next
            
                Dim Texto1 As String, Texto2 As String, Texto3 As String
182             Texto1 = UserList(UserIndex).Name & "(" & UserList(UserIndex).Stats.ELV & ") te invita a jugar el siguiente reto:"
184             Texto2 = Equipo1 & " vs " & Equipo2 & ". Apuesta: " & PonerPuntos(Apuesta) & " monedas de oro" & IIf(CaenItems, " y los items.", ".")
186             Texto3 = "Escribe /ACEPTAR " & UCase$(UserList(UserIndex).Name) & " para participar en el reto."
            
188             If PocionesMaximas >= 0 Then
190                 Texto2 = Texto2 & " Máximo " & PocionesMaximas & " pociones rojas."
                End If

192             For i = 0 To UBound(.Jugadores)
194                 With .Jugadores(i)
196                     Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto1, e_FontTypeNames.FONTTYPE_INFO)
198                     Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto2, e_FontTypeNames.FONTTYPE_New_Naranja)
200                     Call WriteConsoleMsg(.CurIndex.ArrayIndex, Texto3, e_FontTypeNames.FONTTYPE_INFO)
                    End With
                Next

202             .estado = e_SolicitudRetoEstado.Enviada
            End With

204         Call WriteConsoleMsg(UserIndex, "Has enviado una solicitud para el siguiente reto:", e_FontTypeNames.FONTTYPE_INFO)
206         Call WriteConsoleMsg(UserIndex, Texto2, e_FontTypeNames.FONTTYPE_New_Naranja)
208         Call WriteConsoleMsg(UserIndex, "Escribe /CANCELAR para anular la solicitud.", e_FontTypeNames.FONTTYPE_New_Gris)
    
        End With
    
        Exit Sub
    
ErrHandler:
210     Call TraceError(Err.Number, Err.Description, "ModRetos.CrearReto", Erl)

End Sub

Public Sub AceptarReto(ByVal UserIndex As Integer, OferenteName As String)

        On Error GoTo ErrHandler

100     If Not PuedeRetoConMensaje(UserIndex) Then Exit Sub
    
102     With UserList(UserIndex)
104         If .flags.SolicitudReto.estado <> e_SolicitudRetoEstado.Libre Then
106             Call CancelarSolicitudReto(UserIndex, .Name & " ha cancelado la solicitud.")
            
108         ElseIf IsValidUserRef(.flags.AceptoReto) Then
110             Call CancelarSolicitudReto(.flags.AceptoReto.ArrayIndex, .Name & " ha cancelado su admisión.")
            End If
        End With
    
112     If EsGmChar(OferenteName) Then
114         Call WriteConsoleMsg(UserIndex, "¡No puedes jugar retos con administradores!", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        Dim Oferente As t_UserReference
116     Oferente = NameIndex(OferenteName)
    
118     If Not IsValidUserRef(Oferente) Then
120         Call WriteConsoleMsg(UserIndex, OferenteName & " no está conectado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
    

122     With UserList(Oferente.ArrayIndex).flags.SolicitudReto

        Dim JugadorIndex As Integer
124     JugadorIndex = IndiceJugadorEnSolicitud(UserIndex, Oferente.ArrayIndex)

126     If JugadorIndex < 0 Then
128         Call WriteConsoleMsg(UserIndex, UserList(Oferente.ArrayIndex).Name & " no te ha invitado a ningún reto o ha sido cancelado.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

130     If UserList(UserIndex).Stats.GLD < .Apuesta Then
132         Call WriteConsoleMsg(UserIndex, "Necesitas al menos " & PonerPuntos(.Apuesta) & " monedas de oro para aceptar este reto.", e_FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
134     If .PocionesMaximas >= 0 Then
136         If TieneObjetos(38, .PocionesMaximas + 1, UserIndex) Then
138             Call WriteConsoleMsg(UserIndex, "Tienes demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").", e_FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
140     Call MensajeATodosSolicitud(Oferente.ArrayIndex, UserList(UserIndex).name & " ha aceptado el reto.", e_FontTypeNames.FONTTYPE_INFO)
        
142     .Jugadores(JugadorIndex).Aceptado = True
144     Call SetUserRef(.Jugadores(JugadorIndex).CurIndex, UserIndex)
146     UserList(UserIndex).flags.AceptoReto = Oferente
        
148     Call WriteConsoleMsg(UserIndex, "Has aceptado el reto de " & UserList(Oferente.ArrayIndex).name & ".", e_FontTypeNames.FONTTYPE_INFO)
        
            Dim FaltanAceptar As String

            Dim i As Integer
150         For i = 0 To UBound(.Jugadores)
152             If Not .Jugadores(i).Aceptado Then
154                 FaltanAceptar = FaltanAceptar & .Jugadores(i).nombre & " - "
                End If
            Next
        
156         If LenB(FaltanAceptar) > 0 Then
158             FaltanAceptar = Left$(FaltanAceptar, Len(FaltanAceptar) - 3)
160             Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Faltan aceptar: " & FaltanAceptar, e_FontTypeNames.FONTTYPE_New_Gris)
                Exit Sub
            End If
        
162         Call MensajeATodosSolicitud(Oferente.ArrayIndex, "Todos los jugadores han aceptado el reto. Buscando sala...", e_FontTypeNames.FONTTYPE_New_Gris)

164         Call BuscarSala(Oferente.ArrayIndex)

        End With
    
        Exit Sub
    
ErrHandler:
166     Call TraceError(Err.Number, Err.Description, "ModRetos.AceptarReto", Erl)
End Sub

Public Sub CancelarSolicitudReto(ByVal Oferente As Integer, Mensaje As String)
    
        On Error GoTo ErrHandler
    
100     With UserList(Oferente).flags.SolicitudReto
    
102         If .estado = e_SolicitudRetoEstado.EnCola Then
104             Call ListaDeEspera.Remove(Oferente)
            End If

106         .estado = e_SolicitudRetoEstado.Libre
        
            Dim i As Integer, tUser As t_UserReference

            ' Enviamos a los invitados
108         For i = 0 To UBound(.Jugadores)

110             tUser = NameIndex(.Jugadores(i).nombre)
            
112             If IsValidUserRef(tUser) Then
114                 Call WriteConsoleMsg(tUser.ArrayIndex, mensaje, e_FontTypeNames.FONTTYPE_WARNING)
116                 Call WriteConsoleMsg(tUser.ArrayIndex, "El reto ha sido cancelado.", e_FontTypeNames.FONTTYPE_WARNING)

118                 If .Jugadores(i).Aceptado Then
120                     Call SetUserRef(UserList(tUser.ArrayIndex).flags.AceptoReto, 0)
                    End If
                End If

            Next

            ' Y al oferente por separado
122         Call WriteConsoleMsg(Oferente, Mensaje, e_FontTypeNames.FONTTYPE_WARNING)
124         Call WriteConsoleMsg(Oferente, "El reto ha sido cancelado.", e_FontTypeNames.FONTTYPE_WARNING)

        End With
    
        Exit Sub
    
ErrHandler:
126     Call TraceError(Err.Number, Err.Description, "ModRetos.CancelarSolicitudReto", Erl)
    
End Sub

Private Sub BuscarSala(ByVal Oferente As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(Oferente).flags.SolicitudReto

102         If Retos.SalasLibres <= 0 Then
104             Call ListaDeEspera.Add(Oferente, 0)
106             Call MensajeATodosSolicitud(Oferente, "No hay salas disponibles. El reto comenzará cuando se desocupe una sala.", e_FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        
            Dim Sala As Integer, SalaAleatoria As Integer
108         SalaAleatoria = RandomNumber(1, Retos.SalasLibres)
        
110         For Sala = 1 To Retos.TotalSalas
112             If Not Retos.Salas(Sala).EnUso Then
114                 SalaAleatoria = SalaAleatoria - 1
116                 If SalaAleatoria = 0 Then Exit For
                End If
            Next
        
118         Call IniciarReto(Oferente, Sala)
    
        End With
    
        Exit Sub
    
ErrHandler:
120     Call TraceError(Err.Number, Err.Description, "ModRetos.BuscarSala", Erl)
End Sub

Private Sub IniciarReto(ByVal Oferente As Integer, ByVal Sala As Integer)

        On Error GoTo ErrHandler
    
100     With UserList(Oferente).flags.SolicitudReto
    
            ' Última comprobación de si todos pueden entrar/pagar
102         If Not TodosPuedenReto(Oferente) Then Exit Sub
        
            Dim Apuesta As Long, ApuestaStr As String
104         Apuesta = .Apuesta
106         ApuestaStr = PonerPuntos(Apuesta)

            ' Calculamos el tamaño del equipo
108         Retos.Salas(Sala).TamañoEquipoIzq = UBound(.Jugadores) \ 2 + 1
110         Retos.Salas(Sala).TamañoEquipoDer = Retos.Salas(Sala).TamañoEquipoIzq
            ' Reservamos espacio para los jugadores (incluyendo al oferente)
112         ReDim Retos.Salas(Sala).Jugadores(0 To UBound(.Jugadores) + 1)
        
            ' Tiramos una moneda (50-50) y decidimos si agregar al oferente al inicio o al final de la lista
            Dim Moneda As Byte
114         Moneda = RandomNumber(0, 1)
        
            Dim CurIndex As Integer
        
116         If Moneda = 0 Then
                ' Agregamos al oferente al inicio (su equipo juega a la izquierda)
118             Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
120             CurIndex = CurIndex + 1
            End If
        
            Dim i As Integer
        
            ' Agregamos los jugadores alternando 1 y 1 (en los índices pares está el equipo izquierdo y en los impares el derecho - el array empieza en cero)
122         For i = 0 To UBound(.Jugadores)
124             Retos.Salas(Sala).Jugadores(CurIndex) = .Jugadores(i).CurIndex
126             CurIndex = CurIndex + 1
                ' Reset flag
128             Call SetUserRef(UserList(.Jugadores(i).CurIndex.ArrayIndex).flags.AceptoReto, 0)
            Next
        
130         If Moneda = 1 Then
                ' Agregamos al oferente al final (su equipo juega a la derecha)
132             Call SetUserRef(Retos.Salas(Sala).Jugadores(CurIndex), Oferente)
            End If
        
            ' Reset estado de la solicitud, ya que no la necesitamos más
134         .estado = e_SolicitudRetoEstado.Libre
        End With

136     With Retos.Salas(Sala)
138         .EnUso = True
140         .Puntaje = 0
142         .Ronda = 1
144         .Apuesta = Apuesta
146         .TiempoRestante = Retos.DuracionMaxima
148         .CaenItems = UserList(Oferente).flags.SolicitudReto.CaenItems
            Dim tUser As t_UserReference

150         For i = 0 To UBound(.Jugadores)

152             tUser = .Jugadores(i)

                ' Le cobramos
154             UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD - Apuesta
156             Call WriteUpdateGold(tUser.ArrayIndex)
158             Call WriteConsoleMsg(tUser.ArrayIndex, "Otorgas " & ApuestaStr & " monedas de oro al pozo del reto.", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)
            
                ' Desmontamos
160             If UserList(tUser.ArrayIndex).flags.Montado <> 0 Then
162                 Call DoMontar(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).Invent.MonturaObjIndex), UserList(tUser.ArrayIndex).Invent.MonturaSlot)
                End If
            
                ' Dejamos de navegar
164             If UserList(tUser.ArrayIndex).flags.Nadando <> 0 Or UserList(tUser.ArrayIndex).flags.Navegando <> 0 Then
166                 Call DoNavega(tUser.ArrayIndex, ObjData(UserList(tUser.ArrayIndex).Invent.BarcoObjIndex), UserList(tUser.ArrayIndex).Invent.BarcoSlot)
                End If
            
                ' Asignamos flags
168             With UserList(tUser.ArrayIndex).flags
170                 .EnReto = True
172                 .EquipoReto = IIf(i Mod 2, e_EquipoReto.Derecha, e_EquipoReto.Izquierda)
174                 .SalaReto = Sala
                    ' Guardar posición
176                 .LastPos = UserList(tUser.ArrayIndex).pos
                End With
            
178             Call WriteConsoleMsg(tUser.ArrayIndex, "¡Ha comenzado el reto!", e_FontTypeNames.FONTTYPE_New_Rojo_Salmon)
180             Call WriteConsoleMsg(tUser.ArrayIndex, "Para admitir la derrota escribe /ABANDONAR.", e_FontTypeNames.FONTTYPE_New_Gris)

            Next

        End With
    
182     Retos.SalasLibres = Retos.SalasLibres - 1

184     Call IniciarRonda(Sala)

        Exit Sub
    
ErrHandler:
186     Call TraceError(Err.Number, Err.Description, "ModRetos.IniciarReto", Erl)
    
End Sub

Private Sub IniciarRonda(ByVal Sala As Integer)

100     With Retos.Salas(Sala)
    
            Dim i As Integer, tUser As t_UserReference
        
102         For i = 0 To UBound(.Jugadores)

104             tUser = .Jugadores(i)

106             If tUser.ArrayIndex <> 0 Then

108                 Call RevivirYLimpiar(tUser.ArrayIndex)

                    ' Usando el número de ronda y el índice, decidimos el lado al que corresponde
110                 If (.Ronda + i) Mod 2 = 1 Then
                        ' Lado izquierdo
112                     Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.map, .PosIzquierda.x, .PosIzquierda.y, True)
                    Else
                        ' Lado derecho
114                     Call WarpToLegalPos(tUser.ArrayIndex, .PosDerecha.map, .PosDerecha.x, .PosDerecha.y, True)
                    End If

                    ' Si usamos el conteo
116                 If Retos.TiempoConteo > 0 Then
                        ' Le ponemos el conteo
118                     UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = Retos.TiempoConteo
                        ' Lo stoppeamos
120                     Call WriteStopped(tUser.ArrayIndex, True)
                    End If
                
122                 Call WriteConsoleMsg(tUser.ArrayIndex, "Comienza la ronda Nº" & .Ronda, e_FontTypeNames.FONTTYPE_GUILD)

                End If
            Next
    
        End With
    
End Sub

Public Sub MuereEnReto(ByVal UserIndex As Integer)
        On Error GoTo ErrorHandler
        
        Dim Sala As Integer, Equipo As e_EquipoReto

100     With UserList(UserIndex)
102         Sala = .flags.SalaReto
104         Equipo = .flags.EquipoReto
        End With
    
106     With Retos.Salas(Sala)
    
            Dim CurIndex As Integer
        
            ' El equipo derecho está en índices pares
108         If Equipo = e_EquipoReto.Derecha Then CurIndex = 1
        
110         For CurIndex = CurIndex To UBound(.Jugadores) Step 2
112             If .Jugadores(CurIndex).ArrayIndex <> 0 Then
                    ' Si todavía hay alguno vivo del equipo
114                 If UserList(.Jugadores(CurIndex).ArrayIndex).flags.Muerto = 0 Then
                        Exit Sub
                    End If
                End If
            Next
        
            ' Están todos muertos, ganó el equipo contrario
116         Call ProcesarRondaGanada(Sala, EquipoContrario(Equipo))
    
        End With
        
        Exit Sub
ErrorHandler:
118     Call TraceError(Err.Number, Err.Description, "ModRetos.MuereEnReto", Erl)
End Sub

Private Sub ProcesarRondaGanada(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto)

100     With Retos.Salas(Sala)

            ' Sumamos puntaje o restamos según el equipo
102         If Equipo = e_EquipoReto.Derecha Then
104             .Puntaje = .Puntaje + 1
            Else
106             .Puntaje = .Puntaje - 1
            End If
        
            ' Si terminó la tercer ronda o bien algún equipo obtuvo 2 victorias seguidas
108         If .Ronda >= 3 Or Abs(.Puntaje) >= 2 Then
110             Call FinalizarReto(Sala)
                Exit Sub
            End If
        
            ' Aumentamos el número de ronda
112         .Ronda = .Ronda + 1
        
            ' Obtenemos el tamaño actual del equipo (por si alguno abandonó)
            Dim TamañoEquipo As Integer, TamañoEquipo2 As Integer
114         TamañoEquipo = ObtenerTamañoEquipo(Sala, Equipo)
            ' Menos cálculos en el bucle
116         TamañoEquipo2 = TamañoEquipo * 2
        
            ' Obtenemos los nombres del equipo ganador
            Dim i As Integer, nombres As String
118         For i = IIf(Equipo = e_EquipoReto.Izquierda, 0, 1) To TamañoEquipo2 - 1 Step 2

120             If .Jugadores(i).ArrayIndex <> 0 Then
122                 nombres = nombres & UserList(.Jugadores(i).ArrayIndex).name
                
124                 If i < TamañoEquipo2 - 2 Then
126                     nombres = nombres & IIf(i > TamañoEquipo2 - 5, " y ", ", ")
                    End If
                End If
            Next
        
            ' Informamos el ganador de esta ronda
128         For i = 0 To UBound(.Jugadores)
130             If .Jugadores(i).ArrayIndex <> 0 Then
132                 Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, "Esta ronda es para " & nombres & ".", e_FontTypeNames.FONTTYPE_GUILD)
134                 Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, "", 0) ' Dejamos un espacio vertical
                End If
            Next
        
            ' Iniciamos la próxima ronda
136         Call IniciarRonda(Sala)
    
        End With
        
        Exit Sub
ErrorHandler:
138     Call TraceError(Err.Number, Err.Description, "ModRetos.ProcesarRondaGanada", Erl)
End Sub

Public Sub FinalizarReto(ByVal Sala As Integer, Optional ByVal TiempoAgotado As Boolean)
        On Error GoTo ErrorHandler
        
100     With Retos.Salas(Sala)
    
            ' Calculamos el oro total del premio
            Dim OroTotal As Long, Oro As Long, OroStr As String
102         OroTotal = .Apuesta * (UBound(.Jugadores) + 1)
        
            ' Descontamos el impuesto
104         OroTotal = OroTotal * (1 - Retos.ImpuestoApuesta)
    
            ' Decidimos el resultado del reto según el puntaje:
            Dim i As Integer, tUser As t_UserReference, Equipo1 As String, Equipo2 As String
            Dim eloTotalIzquierda As Long, eloTotalDerecha As Long, winsIzquierda As Long, winsDerecha As Long
            Dim todosMayorA35 As Boolean
            todosMayorA35 = True

106         For i = 0 To UBound(.Jugadores)
108           tUser = .Jugadores(i)

110           If tUser.ArrayIndex <> 0 Then
                todosMayorA35 = todosMayorA35 And (UserList(tUser.ArrayIndex).Stats.ELV >= 35)

112             If i Mod 2 = 0 Then
114               eloTotalIzquierda = eloTotalIzquierda + UserList(tUser.ArrayIndex).Stats.ELO
                Else
116               eloTotalDerecha = eloTotalDerecha + UserList(tUser.ArrayIndex).Stats.ELO
                End If
              End If

118         Next i

            ' Empate
120         If .Puntaje = 0 Then
                ' Pagamos a todos los que no abandonaron
122             Oro = OroTotal \ (UBound(.Jugadores) + 1)
124             OroStr = PonerPuntos(Oro)

                ' No hubo ganadores, entonces el ELO no les da el bonus.
126             winsIzquierda = 0
128             winsDerecha = 0
    
130             For i = 0 To UBound(.Jugadores)
132                 tUser = .Jugadores(i)

134                 If IsValidUserRef(tUser) Then
136                     UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + Oro
138                     Call WriteUpdateGold(tUser.ArrayIndex)
140                     Call WriteLocaleMsg(tUser.ArrayIndex, "29", e_FontTypeNames.FONTTYPE_MP, OroStr) ' Has ganado X monedas de oro
142                     Call RevivirYLimpiar(tUser.ArrayIndex)
144                     Call DevolverPosAnterior(tUser.ArrayIndex)
                    
                        ' Reset flags
146                     UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
148                     UserList(tUser.ArrayIndex).flags.EnReto = False
                    
                        ' Nombres
150                     If i Mod 2 Then
152                         If LenB(Equipo2) > 0 Then
154                             Equipo2 = Equipo2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).Name
                            Else
156                             Equipo2 = UserList(tUser.ArrayIndex).name
                            End If
                        Else
158                         If LenB(Equipo1) > 0 Then
160                             Equipo1 = Equipo2 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).Name
                            Else
162                             Equipo1 = UserList(tUser.ArrayIndex).name
                            End If
                        End If
                    End If
                
                Next
            
                ' Anuncio global
164             Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos » " & Equipo1 & " vs " & Equipo2 & ". Ninguno pudo vencer a su rival.", e_FontTypeNames.FONTTYPE_INFO))
166             Call SalaLiberada(Sala)
            ' Hubo un ganador
            Else
                Dim Ganador As e_EquipoReto
            
168             If .Puntaje < 0 Then
170                 Ganador = e_EquipoReto.Izquierda
172                 winsIzquierda = .TamañoEquipoDer
174                 winsDerecha = -.TamañoEquipoIzq
                Else
176                 Ganador = e_EquipoReto.Derecha
178                 winsIzquierda = -.TamañoEquipoDer
180                 winsDerecha = .TamañoEquipoIzq
                End If

                ' Pagamos a los ganadores que no abandonaron
182             Oro = OroTotal \ ObtenerTamañoEquipo(Sala, Ganador)
184             OroStr = PonerPuntos(Oro)

186             For i = 0 To UBound(.Jugadores)
188                 tUser = .Jugadores(i)

190                 If IsValidUserRef(tUser) Then
192                     Call RevivirYLimpiar(tUser.ArrayIndex)
194                     If UserList(tUser.ArrayIndex).flags.EquipoReto = Ganador Then
196                         UserList(tUser.ArrayIndex).Stats.GLD = UserList(tUser.ArrayIndex).Stats.GLD + Oro
198                         Call WriteUpdateGold(tUser.ArrayIndex)
200                         Call WriteLocaleMsg(tUser.ArrayIndex, "29", e_FontTypeNames.FONTTYPE_MP, OroStr) ' Has ganado X monedas de oro


202                         If .CaenItems Then
206                                Call WarpToLegalPos(tUser.ArrayIndex, .PosIzquierda.map, .PosIzquierda.x, .PosIzquierda.y, True)
                            Else
210                             UserList(tUser.ArrayIndex).flags.EnReto = False
212                             Call DevolverPosAnterior(tUser.ArrayIndex)
                            End If
                        Else
214                         If .CaenItems Then
216                             Call TirarItemsEnPos(tUser.ArrayIndex, ((.PosDerecha.x - .PosIzquierda.x) \ 2) + .PosIzquierda.x, ((.PosDerecha.y - .PosIzquierda.y) \ 2) + .PosIzquierda.y)
                            End If
218                             UserList(tUser.ArrayIndex).flags.EnReto = False
220                             Call DevolverPosAnterior(tUser.ArrayIndex)
                        End If
                    
                    
                    
                        ' Reset flags
222                     UserList(tUser.ArrayIndex).Counters.CuentaRegresiva = -1
                    
224                     If TiempoAgotado Then
226                         Call WriteConsoleMsg(tUser.ArrayIndex, "Se ha agotado el tiempo del reto.", e_FontTypeNames.FONTTYPE_New_Gris)
                        End If

                        ' Nombres
228                     If i Mod 2 Then
                    
230                         If LenB(Equipo2) > 0 Then
232                             Equipo2 = Equipo2 & IIf((i + 1) \ 2 < .TamañoEquipoDer - 2, ", ", " y ") & UserList(tUser.ArrayIndex).Name
                            Else
234                             Equipo2 = UserList(tUser.ArrayIndex).name
                            End If
                        Else
236                         If LenB(Equipo1) > 0 Then
238                             Equipo1 = Equipo1 & IIf(i \ 2 < .TamañoEquipoIzq - 2, ", ", " y ") & UserList(tUser.ArrayIndex).Name
                            Else
240                             Equipo1 = UserList(tUser.ArrayIndex).name
                            End If
                        End If
                    
                    End If
                Next

                Dim equipoGanador As String, equipoPerdedor As String
242             equipoGanador = IIf(Ganador = e_EquipoReto.Izquierda, Equipo1, Equipo2)
244             equipoPerdedor = IIf(Ganador = e_EquipoReto.Izquierda, Equipo2, Equipo1)

                ' Anuncio global
246             If UBound(.Jugadores) > 1 Then
248                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos » El equipo " & equipoGanador & " venció al equipo " & equipoPerdedor & " y se quedo con el botín de: " & PonerPuntos(.Apuesta) & " monedas de oro. ", e_FontTypeNames.FONTTYPE_INFO))
        
                Else ' 1 vs 1
250                 Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos » " & equipoGanador & " venció a " & equipoPerdedor & " y se quedo con el botín de: " & PonerPuntos(.Apuesta) & " monedas de oro. ", e_FontTypeNames.FONTTYPE_INFO))

                End If
            
252             If .CaenItems Then
254                 Call IniciarDepositoItems(Sala)
                Else
256                 Call SalaLiberada(Sala)
                End If
            
            End If

            ' Actualizamos el ELO de cada jugador, inspirados en `Algoritmo de 400`
            ' https://en.wikipedia.org/wiki/Elo_rating_system
            Dim eloDiff As Long
258         For i = 0 To UBound(.Jugadores)
260           tUser = .Jugadores(i)

262           If IsValidUserRef(tUser) Then
                If todosMayorA35 Then
266               If i Mod 2 = 0 Then ' Jugadores en el equipo Izquierdo
268                 eloDiff = winsIzquierda * (eloTotalDerecha * 0.1)
                  Else
270                 eloDiff = winsDerecha * (eloTotalIzquierda * 0.1)
                  End If

                  If eloDiff > 0 Then
                    Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageConsoleMsg("Has ganado " & Abs(eloDiff) & " puntos de ELO!", e_FontTypeNames.FONTTYPE_ROSA))
                  Else
272                 If UserList(tUser.ArrayIndex).Stats.ELO < Abs(eloDiff) Then
274                   eloDiff = -UserList(tUser.ArrayIndex).Stats.ELO
                    End If

                    Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageConsoleMsg("Has perdido " & Abs(eloDiff) & " puntos de ELO!", e_FontTypeNames.FONTTYPE_ROSA))
                  End If

276               UserList(tUser.ArrayIndex).Stats.ELO = UserList(tUser.ArrayIndex).Stats.ELO + eloDiff
                Else ' Alguno es menor a level 35
                  Call SendData(SendTarget.ToIndex, tUser.ArrayIndex, PrepareMessageConsoleMsg("Al menos un participante del reto tiene nivel menor a 35, tu ELO permanece igual.", e_FontTypeNames.FONTTYPE_INFOIAO))
                End If
              End If

278         Next i
    
        End With
        
        Exit Sub
ErrorHandler:
280     Call TraceError(Err.Number, Err.Description, "ModRetos.FinalizarReto", Erl)
End Sub
Public Sub TirarItemsEnPos(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)
            
        On Error GoTo TirarItemsEnPos_Err

        Dim i         As Byte
        Dim NuevaPos  As t_WorldPos
        Dim MiObj     As t_Obj
        Dim ItemIndex As Integer
        Dim posItems As t_WorldPos
        
              
100     With UserList(UserIndex)
102         posItems.Map = .Pos.Map
104         posItems.X = X
106         posItems.Y = Y
            
108         For i = 1 To .CurrentInventorySlots
110             ItemIndex = .Invent.Object(i).ObjIndex
112             If ItemIndex > 0 Then
114                 If ItemSeCae(ItemIndex) And PirataCaeItem(UserIndex, i) And (Not EsNewbie(UserIndex) Or Not ItemNewbie(ItemIndex)) Then
116                     NuevaPos.X = 0
118                     NuevaPos.Y = 0
120                     MiObj.amount = .Invent.Object(i).amount
122                     MiObj.ObjIndex = ItemIndex
                        
124                     Call Tilelibre(posItems, NuevaPos, MiObj, True, True, False)
            
126                     If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
128                         Call DropObj(UserIndex, i, MiObj.amount, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)
                        
                        ' WyroX: Si no hay lugar, quemamos el item del inventario (nada de mochilas gratis)
                        Else
130                         Call QuitarUserInvItem(UserIndex, i, MiObj.amount)
132                         Call UpdateUserInv(False, UserIndex, i)
                        End If
                
                    End If

                End If
    
134         Next i
    
        End With
 
        Exit Sub

TirarItemsEnPos_Err:
136     Call TraceError(Err.Number, Err.Description, "InvUsuario.TirarItemsEnPos", Erl)


            
End Sub


Public Sub IniciarDepositoItems(ByVal Sala As Integer)
        Dim i As Byte
         Dim Ganador As e_EquipoReto
            
        
100     With Retos.Salas(Sala)
102         If .Puntaje < 0 Then
104             Ganador = e_EquipoReto.Izquierda
            Else
106             Ganador = e_EquipoReto.Derecha
            End If
        
108         For i = 0 To UBound(.Jugadores)
110             If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Ganador Then
112                 Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, "Tienes 1 minuto para levantar los items del piso.", e_FontTypeNames.FONTTYPE_INFO)
                End If
114         Next i
        
            Dim Pos As t_WorldPos
        
116         Pos.Map = .PosIzquierda.Map
118         Pos.X = ((.PosDerecha.X - .PosIzquierda.X) \ 2) + .PosIzquierda.X
120         Pos.Y = ((.PosDerecha.Y - .PosIzquierda.Y) \ 2) + .PosIzquierda.Y
            'Spawneo un banquero.
122         .IndexBanquero = SpawnNpc(3, Pos, True, False)
#If DEBUGGING Then
            .TiempoItems = 20
#Else
124         .TiempoItems = 60
#End If

        End With
    
    
End Sub

Public Sub TerminarTiempoAgarrarItems(ByVal Sala As Integer)
        
        
    
        Dim Ganador As e_EquipoReto
100     With Retos.Salas(Sala)
            'Mato al banquero
102         Call QuitarNPC(.IndexBanquero, eChallenge)
        
104         If .Puntaje < 0 Then
106             Ganador = e_EquipoReto.Izquierda
            Else
108             Ganador = e_EquipoReto.Derecha
            End If
        
            Dim i As Byte
110         For i = 0 To UBound(.Jugadores)
112             If IsValidUserRef(.Jugadores(i)) Then
114                 If UserList(.Jugadores(i).ArrayIndex).flags.EquipoReto = Ganador Then
116                     UserList(.Jugadores(i).ArrayIndex).flags.EnReto = False
118                     Call DevolverPosAnterior(.Jugadores(i).ArrayIndex)
                    End If
                End If
120         Next i
122         .TiempoItems = 0
        
            Dim X As Byte
            Dim Y As Byte
        
124         For X = .PosIzquierda.X To .PosDerecha.X
126             For Y = .PosIzquierda.Y To .PosDerecha.Y
128                 Call EraseObj(MAX_INVENTORY_OBJS, .PosIzquierda.Map, X, Y)
130             Next Y
132         Next X
        
        End With
    
    
134     Call SalaLiberada(Sala)
End Sub

Public Sub AbandonarReto(ByVal UserIndex As Integer, Optional ByVal Desconexion As Boolean)
    
        Dim Sala As Integer, Equipo As e_EquipoReto
100     With UserList(UserIndex)
102         Sala = .flags.SalaReto
104         Equipo = .flags.EquipoReto

106         .Counters.CuentaRegresiva = -1
108         .flags.EnReto = False
        End With
    
110     With Retos.Salas(Sala)
        
        
        
112         If .CaenItems And Abs(.Puntaje) >= 2 Then
114                 If .Puntaje < 0 Then
116                     .TamañoEquipoIzq = .TamañoEquipoIzq - 1
118                     If .TamañoEquipoIzq <= 0 Then
120                         TerminarTiempoAgarrarItems (Sala)
                        End If
                    Else
122                     .TamañoEquipoDer = .TamañoEquipoDer - 1
124                     If .TamañoEquipoDer <= 0 Then
126                         TerminarTiempoAgarrarItems (Sala)
                        End If
                    End If
                Exit Sub
            End If
        
128         If Not Desconexion Then
130             Call WriteConsoleMsg(UserIndex, "Has abandonado el reto.", e_FontTypeNames.FONTTYPE_INFO)
            End If
        
            ' Restamos un miembro al equipo y si llega a cero entonces procesamos la derrota
132         If Equipo = e_EquipoReto.Izquierda Then
134             If .TamañoEquipoIzq > 1 Then
136                 .TamañoEquipoIzq = .TamañoEquipoIzq - 1
                Else
138                 .Puntaje = 123 ' Forzamos puntaje positivo
140                 Call FinalizarReto(Sala)
                    Exit Sub
                End If

            Else
142             If .TamañoEquipoDer > 1 Then
144                 .TamañoEquipoDer = .TamañoEquipoDer - 1
                Else
146                 .Puntaje = -123 ' Forzamos puntaje negativo
148                 Call FinalizarReto(Sala)
                    Exit Sub
                End If
            End If
        
150         Call RevivirYLimpiar(UserIndex)
152         Call DevolverPosAnterior(UserIndex)
        
            Dim texto As String
154         If Desconexion Then
156             texto = UserList(UserIndex).Name & " es descalificado por desconectarse."
            Else
158             texto = UserList(UserIndex).Name & " ha abandonado el reto."
            End If
        
            Dim i As Integer
160         For i = 0 To UBound(.Jugadores)
162             If .Jugadores(i).ArrayIndex = UserIndex Then
164                 Call SetUserRef(.Jugadores(i), 0)
                Else
166                 Call WriteConsoleMsg(.Jugadores(i).ArrayIndex, texto, e_FontTypeNames.FONTTYPE_New_Gris)
                End If
            Next
    
        End With
    
End Sub

Private Sub SalaLiberada(ByVal Sala As Integer)

        On Error GoTo ErrHandler
    
100     Retos.Salas(Sala).EnUso = False
102     Retos.SalasLibres = Retos.SalasLibres + 1
104     If ListaDeEspera.Count > 0 Then
    
            Dim Oferente As Integer
106         Oferente = ListaDeEspera.Keys(0)
108         Call ListaDeEspera.Remove(Oferente)
            
110         Call IniciarReto(Oferente, Sala)

        End If
    
        Exit Sub
    
ErrHandler:
112     Call TraceError(Err.Number, Err.Description, "ModRetos.SalaLiberada", Erl)
    
End Sub

Public Function PuedeReto(ByVal UserIndex As Integer) As Boolean
    
100     With UserList(UserIndex)
        
102         If .flags.EnReto Then Exit Function
        
104         If .flags.EnConsulta Then Exit Function

106         If .Pos.Map = 0 Or .Pos.X = 0 Or .Pos.Y = 0 Then Exit Function
            
108         If MapInfo(.Pos.Map).Seguro = 0 Or Zona(.ZonaId).Segura = 0 Then Exit Function
        
110         If .flags.EnTorneo Then Exit Function
            
112         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then Exit Function
        
        End With
    
114     PuedeReto = True
    
End Function

Public Function PuedeRetoConMensaje(ByVal UserIndex As Integer) As Boolean

100     With UserList(UserIndex)
        
102         If .flags.EnReto Then
104             Call WriteConsoleMsg(UserIndex, "Ya te encuentras en un reto.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
106         If .flags.EnConsulta Then
108             Call WriteConsoleMsg(UserIndex, "No puedes acceder a un reto si estás en consulta.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
            If .flags.jugando_captura = 1 Then
109             Call WriteConsoleMsg(UserIndex, "No puedes jugar un reto estando en un evento.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
110         If Not esCiudad(.Pos.map) Then
112             Call WriteConsoleMsg(UserIndex, "No puedes participar de un reto en un mapa inseguro.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
114         If .flags.EnTorneo Then
116             Call WriteConsoleMsg(UserIndex, "No puedes ir a un reto si participas de un torneo.", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
118         If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = CARCEL Then
120             Call WriteConsoleMsg(UserIndex, "¡Estás encarcelado!", e_FontTypeNames.FONTTYPE_INFO)
                Exit Function
            End If
        
        End With

122     PuedeRetoConMensaje = True

End Function

Private Function IndiceJugadorEnSolicitud(ByVal UserIndex As Integer, ByVal Oferente As Integer) As Integer

100     With UserList(Oferente).flags.SolicitudReto
    
102         IndiceJugadorEnSolicitud = -1

104         If .estado <> e_SolicitudRetoEstado.Enviada Then Exit Function

            Dim i As Integer
106         For i = 0 To UBound(.Jugadores)
108             If .Jugadores(i).nombre = UserList(UserIndex).Name Then
110                 IndiceJugadorEnSolicitud = i
                    Exit Function
                End If
            Next
    
        End With

End Function

Private Sub MensajeATodosSolicitud(ByVal Oferente As Integer, Mensaje As String, ByVal Fuente As e_FontTypeNames)
    
100     With UserList(Oferente).flags.SolicitudReto

            Dim i As Integer
102         For i = 0 To UBound(.Jugadores)
104             If .Jugadores(i).Aceptado Then
106                 Call WriteConsoleMsg(.Jugadores(i).CurIndex.ArrayIndex, mensaje, Fuente)
                End If
            Next
        
108         Call WriteConsoleMsg(Oferente, Mensaje, Fuente)

        End With
    
End Sub

Private Function TodosPuedenReto(ByVal Oferente As Integer) As Boolean

        On Error GoTo ErrHandler
    
100     With UserList(Oferente).flags.SolicitudReto
    
102         If Not PuedeReto(Oferente) Then
104             Call CancelarSolicitudReto(Oferente, UserList(Oferente).Name & " no puede entrar al reto en este momento.")
                Exit Function
106         ElseIf UserList(Oferente).Stats.GLD < .Apuesta Then
108             Call CancelarSolicitudReto(Oferente, UserList(Oferente).Name & " no tiene las monedas de oro suficientes.")
                Exit Function

110         ElseIf .PocionesMaximas >= 0 Then
112             If TieneObjetos(38, .PocionesMaximas + 1, Oferente) Then
114                 Call CancelarSolicitudReto(Oferente, UserList(Oferente).Name & " tiene demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").")
                    Exit Function
                End If
            End If


            Dim i As Integer
        
116         For i = 0 To UBound(.Jugadores)
118             If Not PuedeReto(.Jugadores(i).CurIndex.ArrayIndex) Then
120                 Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no puede entrar al reto en este momento.")
                    Exit Function

122             ElseIf UserList(.Jugadores(i).CurIndex.ArrayIndex).Stats.GLD < .Apuesta Then
124                 Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).name & " no tiene las monedas de oro suficientes.")
                    Exit Function
                
126             ElseIf .PocionesMaximas >= 0 Then
128                 If TieneObjetos(38, .PocionesMaximas + 1, Oferente) Then
130                     Call CancelarSolicitudReto(Oferente, UserList(.Jugadores(i).CurIndex.ArrayIndex).Name & " tiene demasiadas pociones rojas (Cantidad máxima: " & .PocionesMaximas & ").")
                        Exit Function
                    End If
                End If
            Next
        
132         TodosPuedenReto = True
    
        End With
    
        Exit Function
    
ErrHandler:
134     Call TraceError(Err.Number, Err.Description, "ModRetos.TodosPuedenReto", Erl)
    
End Function

Private Function EquipoContrario(ByVal Equipo As e_EquipoReto) As e_EquipoReto
100     If Equipo = e_EquipoReto.Izquierda Then
102         EquipoContrario = e_EquipoReto.Derecha
        Else
104         EquipoContrario = e_EquipoReto.Izquierda
        End If
End Function

Private Function ObtenerTamañoEquipo(ByVal Sala As Integer, ByVal Equipo As e_EquipoReto) As Integer
100     If Equipo = e_EquipoReto.Izquierda Then
102         ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoIzq
        Else
104         ObtenerTamañoEquipo = Retos.Salas(Sala).TamañoEquipoDer
        End If
End Function

Private Sub RevivirYLimpiar(ByVal UserIndex As Integer)
    
100         Call WriteStopped(UserIndex, False)
    
        ' Si está vivo
102     If UserList(UserIndex).flags.Muerto = 0 Then
104         Call LimpiarEstadosAlterados(UserIndex)
        End If

        ' Si está muerto lo revivimos, sino lo curamos
106     Call RevivirUsuario(UserIndex)

End Sub
