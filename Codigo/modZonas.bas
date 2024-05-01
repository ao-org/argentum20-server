Attribute VB_Name = "modZonas"
Public Type t_ZonaInfo
    Zona_name As String
    
    Deleted As Byte
    Map As Integer
    X As Integer
    Y As Integer
    X2 As Integer
    Y2 As Integer
    
    Backup As Byte
    
    Lluvia As Byte
    Nieve As Byte
    Niebla As Byte
    
    Ambient As String
    Base_light As Long
    Terreno As String
    
    MinLevel As Integer
    MaxLevel As Integer
    Segura As Byte
    Newbie As Byte
    
    SoloClanes As Byte
    SoloFaccion As Byte
    SinMagia As Byte
    SinInvi As Byte
    SinMascotas As Byte
    SinResucitar As Byte
    Interdimensional As Byte
    
    Faccion As Byte
          
    Musica1 As Integer
    Musica2 As Integer
    Musica3 As Integer
        
    Salida As t_WorldPos
    
End Type

Public Type t_NpcSpawn_List
    NpcIndex As Integer
    Cantidad As Integer
End Type

Public Type t_NpcSpawn
    Deleted As Byte
    Map As Integer
    X As Integer
    Y As Integer
    X2 As Integer
    Y2 As Integer
    CantNPCs As Byte
    NPCs() As t_NpcSpawn_List
End Type

Public Zona() As t_ZonaInfo
Public NpcSpawn() As t_NpcSpawn


Public Sub LoadNpcSpawn()
    On Error GoTo ErrHandler

    Dim i As Integer

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(DatPath & "NpcSpawns.dat")
    
    Dim Cant As Integer
    Dim Key As String
    Cant = Lector.GetValue("INIT", "Cantidad")
    
    ReDim NpcSpawn(1 To Cant)
    
    For i = 1 To Cant
        Key = "NpcSpawn" & i
        With NpcSpawn(i)
            
            .Deleted = val(Lector.GetValue(Key, "Deleted"))
            .Map = val(Lector.GetValue(Key, "Map"))
            .X = val(Lector.GetValue(Key, "X"))
            .Y = val(Lector.GetValue(Key, "Y"))
            .X2 = val(Lector.GetValue(Key, "X2"))
            .Y2 = val(Lector.GetValue(Key, "Y2"))
            .CantNPCs = val(Lector.GetValue(Key, "CantNpcs"))
            If .CantNPCs > 0 Then
                ReDim .NPCs(1 To .CantNPCs)
                For X = 1 To .CantNPCs
                    .NPCs(X).NpcIndex = val(Lector.GetValue(Key, "NpcIndex" & X))
                    .NPCs(X).Cantidad = val(Lector.GetValue(Key, "NpcCantidad" & X))
                Next X
            End If
        End With
        
        
    Next i
    
    Exit Sub
    
ErrHandler:
302     Call TraceError(Err.Number, Err.Description, "modZonas.LoadNpcSpawn", Erl)
End Sub

Public Sub InitNpcSpawn()
    On Error GoTo ErrHandler

    Dim i As Integer
    Dim X As Integer
    Dim LoopC As Integer
    Dim Pos As t_WorldPos

    For i = 1 To UBound(NpcSpawn)
        
        With NpcSpawn(i)
            If .X > 0 And .Deleted = 0 Then
                For X = 1 To NpcSpawn(i).CantNPCs
                    For LoopC = 1 To .NPCs(X).Cantidad
                        Call CrearNPC(.NPCs(X).NpcIndex, i, Pos)
                    Next LoopC
                Next X
            End If
        End With
    Next i
    
    Exit Sub
    
ErrHandler:
302     Call TraceError(Err.Number, Err.Description, "modZonas.InitNpcSpawn", Erl)
End Sub

Public Sub LoadZonas()
    On Error GoTo ErrHandler

    Dim i As Integer

    Dim Lector As clsIniManager
    Set Lector = New clsIniManager
    Call Lector.Initialize(DatPath & "Zonas.dat")
    
    Dim Key As String
    Dim Cant As Integer
    Cant = Lector.GetValue("Config", "Cantidad")
    
    ReDim Zona(1 To Cant)
    
    For i = 1 To Cant
        Key = "Zona" & i
        With Zona(i)
            
            .Zona_name = Lector.GetValue(Key, "Nombre")
            .Map = val(Lector.GetValue(Key, "Mapa"))
            .Deleted = val(Lector.GetValue(Key, "Deleted"))
            .X = val(Lector.GetValue(Key, "X1"))
            .Y = val(Lector.GetValue(Key, "Y1"))
            .X2 = val(Lector.GetValue(Key, "X2"))
            .Y2 = val(Lector.GetValue(Key, "Y2"))
            
            .Backup = val(Lector.GetValue(Key, "Backup"))
            .Lluvia = val(Lector.GetValue(Key, "Lluvia"))
            .Nieve = val(Lector.GetValue(Key, "Nieve"))
            .Niebla = val(Lector.GetValue(Key, "Niebla"))
            .MinLevel = val(Lector.GetValue(Key, "MinLevel"))
            .MaxLevel = val(Lector.GetValue(Key, "MaxLevel"))
            .Segura = val(Lector.GetValue(Key, "Segura"))
            .Newbie = val(Lector.GetValue(Key, "Newbie"))
            .SinMagia = val(Lector.GetValue(Key, "SinMagia"))
            .SinInvi = val(Lector.GetValue(Key, "SinInvi"))
            .SinMascotas = val(Lector.GetValue(Key, "SinMascotas"))
            .SinResucitar = val(Lector.GetValue(Key, "SinResucitar"))
            .SoloClanes = val(Lector.GetValue(Key, "SoloClanes"))
            .SoloFaccion = val(Lector.GetValue(Key, "SoloFaccion"))
            .Interdimensional = val(Lector.GetValue(Key, "Interdimensional"))

            .Faccion = val(Lector.GetValue(Key, "Faccion"))
            .Terreno = Lector.GetValue(Key, "Terreno")
            .Ambient = Lector.GetValue(Key, "Ambient")
            .Base_light = val(Lector.GetValue(Key, "Base_light"))
            .Salida.Map = val(Lector.GetValue(Key, "SalidaMap"))
            .Salida.X = val(Lector.GetValue(Key, "SalidaX"))
            .Salida.Y = val(Lector.GetValue(Key, "SalidaY"))
            
            .Musica1 = val(Lector.GetValue(Key, "Musica1"))
            .Musica2 = val(Lector.GetValue(Key, "Musica2"))
            .Musica3 = val(Lector.GetValue(Key, "Musica3"))
        End With
        
        
    Next i
    
    Exit Sub
    
ErrHandler:
302     Call TraceError(Err.Number, Err.Description, "modZonas.LoadZonas", Erl)
End Sub
Public Sub InitZonas()
    Dim i As Integer
    Dim X As Integer
    Dim Y As Integer
    
    'Precargo todas las zonas en el mapa para que la busqueda sea mucho mas rapida.
    For i = 1 To UBound(Zona)
        With Zona(i)
            If .X > 0 And .Deleted = 0 And .Map <= NumMaps Then
                For Y = .Y To .Y2
                    For X = .X To .X2
                        MapData(.Map).Tile(X, Y).ZonaId = i
                    Next X
                Next Y
            End If
        End With
    Next i
    
    'Seteo los tiles que no entran en ninguna zona con la zona por default del mapa o segun la textura que haya.
    For i = 1 To NumMaps
        For Y = 1 To MapInfo(i).Height
            For X = 1 To MapInfo(i).Width
                If MapData(i).Tile(X, Y).ZonaId = 0 Then
                    MapData(i).Tile(X, Y).ZonaId = 1
                End If
            Next X
        Next Y
    Next i
    
    'Actualizamos la zona de los npcs creados con el mapa
    For i = 1 To LastNPC
        If NpcList(i).ZonaId = 0 Then
            NpcList(i).ZonaId = ZonaByPos(NpcList(i).Pos)
        End If
    Next i
End Sub
Public Sub RestoreZonaBackup()

    On Error GoTo Handler:
    
    Dim j As Integer, X As Integer, Y As Integer, Item As Integer, Cant As Long
    Dim fh As Integer
    
    For j = 1 To UBound(Zona)
        If Zona(j).Deleted = 0 And Zona(j).Map <= NumMaps And Zona(j).X > 0 And Zona(j).Backup = 1 Then
            fh = FreeFile
            If FileExist(App.Path & "\ZonaBackups\Zona" & j & ".bk") Then
                Open App.Path & "\ZonaBackups\Zona" & j & ".bk" For Binary As fh
                    Do While Not EOF(fh)
                        Get #fh, , X
                        Get #fh, , Y
                        Get #fh, , Item
                        Get #fh, , Cant
                        If Item > 0 Then
                            MapData(Zona(j).Map).Tile(X, Y).ObjInfo.ObjIndex = Item
                            MapData(Zona(j).Map).Tile(X, Y).ObjInfo.amount = Cant
                            If ObjData(Item).OBJType = otPuertas Then
                                Call BloquearPuerta(Zona(j).Map, X, Y, ObjData(Item).Cerrada = 1)
                            End If
                        End If
                    Loop
                Close #fh
            End If
        End If
    Next j
    
    Exit Sub
    
Handler:
128 Call TraceError(Err.Number, Err.Description, "Admin.WorldSave", Erl)

End Sub
Public Function ZonaByPos(Pos As t_WorldPos) As Integer
    ZonaByPos = MapData(Pos.Map).Tile(Pos.X, Pos.Y).ZonaId
End Function
