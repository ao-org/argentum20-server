Attribute VB_Name = "EffectsOverTime"
Option Explicit

Private LastUpdateTime As Long
Private UniqueIdCounter As Long
Const ACTIVE_EFFECTS_MIN_SIZE As Integer = 500
Private ActiveEffects As t_EffectOverTimeList

Const INITIAL_POOL_SIZE = 200
Private EffectPools() As t_EffectOverTimeList

Public Sub InitializePools()
On Error GoTo InitializePools_Err
    Dim i As Integer
    Dim j As Integer
100 ReDim EffectPools(1 To e_EffectOverTimeType.EffectTypeCount - 1) As t_EffectOverTimeList
102 For i = 1 To e_EffectOverTimeType.EffectTypeCount - 1
104     ReDim EffectPools(i).EffectList(INITIAL_POOL_SIZE) As IBaseEffectOverTime
106     For j = 0 To INITIAL_POOL_SIZE
108         Call AddEffect(EffectPools(i), InstantiateEOT(i))
110     Next j
    Next i
    Exit Sub
InitializePools_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.InitializePools", Erl)
End Sub

Public Sub UpdateEffectOverTime()
On Error GoTo Update_Err
    Dim CurrTime As Long
    Dim ElapsedTime As Long
100 CurrTime = GetTickCount()
102 If CurrTime < LastUpdateTime Then ' GetTickCount can overflow se we take care of that
104     ElapsedTime = 0
    Else
106     ElapsedTime = CurrTime - LastUpdateTime
    End If
108 LastUpdateTime = CurrTime
    
    
    Dim i As Integer
200 Do While i < ActiveEffects.EffectCount
202     If UpdateEffect(i, ElapsedTime) Then
204         i = i + 1
        End If
    Loop
    Exit Sub
Update_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.Update", Erl)
End Sub

Private Function UpdateEffect(ByVal Index As Integer, ByVal ElapsedTime As Long) As Boolean
On Error GoTo UpdateEffect_Err
    'this should never happend but it covers us for breaking all effects if something goes wrong
100 If ActiveEffects.EffectList(index) Is Nothing Then
102     UpdateEffect = True
        Exit Function
    End If
    Dim CurrentEffect As IBaseEffectOverTime
    Set CurrentEffect = ActiveEffects.EffectList(index)
104 CurrentEffect.Update (ElapsedTime)
106 If CurrentEffect.RemoveMe Then
108     If CurrentEffect.TargetIsValid Then
110         If CurrentEffect.TargetRefType = eUser Then
112             Call RemoveEffect(UserList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
114         ElseIf CurrentEffect.TargetRefType = eNpc Then
116             Call RemoveEffect(NpcList(CurrentEffect.TargetArrayIndex).EffectOverTime, CurrentEffect)
            End If
        End If
        Call RemoveEffectAtPos(ActiveEffects, index)
120     Call RecycleEffect(CurrentEffect)
134     UpdateEffect = False
    Else
138     UpdateEffect = True
    End If
    Exit Function
UpdateEffect_Err:
    Call TraceError(Err.Number, Err.Description, "EffectsOverTime.UpdateEffect", Erl)
    Set ActiveEffects.EffectList(index) = Nothing
    UpdateEffect = True
End Function

Private Function GetNextId() As Long
    UniqueIdCounter = (UniqueIdCounter + 1) And &H7FFFFFFF
    GetNextId = UniqueIdCounter
End Function

Public Sub CreateEffect(ByVal sourceIndex As Integer, ByVal sourceType As e_ReferenceType, _
                                  ByVal TargetIndex As Integer, ByVal TargetType As e_ReferenceType, _
                                  ByVal EffectIndex As Integer)
On Error GoTo CreateEffect_Err
    Dim EffectType As e_EffectOverTimeType
100 EffectType = EffectOverTime(EffectIndex).Type
    Select Case EffectType
        Case e_EffectOverTimeType.eHealthModifier
102         Dim Dot As UpdateHpOverTime
104         Set Dot = GetEOT(EffectType)
106         UniqueIdCounter = GetNextId()
108         Call Dot.Setup(sourceIndex, sourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
110         Call AddEffectToUpdate(Dot)
112         If TargetType = eUser Then
114             Call AddEffect(UserList(TargetIndex).EffectOverTime, Dot)
116         ElseIf TargetType = eNpc Then
118             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Dot)
            End If
        Case e_EffectOverTimeType.eApplyModifiers
130         Dim StatDot As StatModifier
132         Set StatDot = GetEOT(EffectType)
134         UniqueIdCounter = GetNextId()
136         Call StatDot.Setup(sourceIndex, sourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
138         Call AddEffectToUpdate(StatDot)
140         If TargetType = eUser Then
142             Call AddEffect(UserList(TargetIndex).EffectOverTime, StatDot)
144         ElseIf TargetType = eNpc Then
146             Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatDot)
            End If
        Case e_EffectOverTimeType.eProvoke
150         Dim Provoke As EffectProvoke
152         Set Provoke = GetEOT(EffectType)
154         UniqueIdCounter = GetNextId()
156         Call Provoke.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
158         Call AddEffectToUpdate(Provoke)
160         If TargetType = eUser Then
162             Call AddEffect(UserList(TargetIndex).EffectOverTime, Provoke)
164         ElseIf TargetType = eNpc Then
166             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Provoke)
            End If
        Case e_EffectOverTimeType.eProvoked
170         Dim StatProvoked As EffectProvoked
172         Set StatProvoked = GetEOT(EffectType)
174         UniqueIdCounter = GetNextId()
176         Call StatProvoked.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
178         Call AddEffectToUpdate(StatProvoked)
180         If TargetType = eUser Then
182             Call AddEffect(UserList(TargetIndex).EffectOverTime, StatProvoked)
184         ElseIf TargetType = eNpc Then
186             Call AddEffect(NpcList(TargetIndex).EffectOverTime, StatProvoked)
            End If
        Case e_EffectOverTimeType.eDrunk
190         Dim Drunk As DrunkEffect
192         Set Drunk = GetEOT(EffectType)
194         UniqueIdCounter = GetNextId()
196         Call Drunk.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
198         Call AddEffectToUpdate(Drunk)
200         If TargetType = eUser Then
202             Call AddEffect(UserList(TargetIndex).EffectOverTime, Drunk)
204         ElseIf TargetType = eNpc Then
206             Call AddEffect(NpcList(TargetIndex).EffectOverTime, Drunk)
            End If
        Case e_EffectOverTimeType.eTranslation
230      Dim TE As TranslationEffect
232         Set TE = GetEOT(EffectType)
236         UniqueIdCounter = GetNextId()
238         Call TE.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
240         Call AddEffectToUpdate(TE)
242         If TargetType = eUser Then
244             Call AddEffect(UserList(TargetIndex).EffectOverTime, TE)
246         ElseIf TargetType = eNpc Then
248             Call AddEffect(NpcList(TargetIndex).EffectOverTime, TE)
            End If
        Case e_EffectOverTimeType.eApplyEffectOnHit
390         Dim EOH As ApplyEffectOnHit
392         Set EOH = GetEOT(EffectType)
394         UniqueIdCounter = GetNextId()
396         Call EOH.Setup(SourceIndex, SourceType, EffectIndex, UniqueIdCounter)
398         Call AddEffectToUpdate(EOH)
400         If TargetType = eUser Then
402             Call AddEffect(UserList(TargetIndex).EffectOverTime, EOH)
404         ElseIf TargetType = eNpc Then
406             Call AddEffect(NpcList(TargetIndex).EffectOverTime, EOH)
            End If
        Case e_EffectOverTimeType.eManaModifier
420         Dim Mot As UpdateManaOverTime
422         Set Mot = GetEOT(EffectType)
426         UniqueIdCounter = GetNextId()
428         Call Mot.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
430         Call AddEffectToUpdate(Mot)
432         If TargetType = eUser Then
434             Call AddEffect(UserList(TargetIndex).EffectOverTime, Mot)
438             'npc doesn't have mana
            End If
        Case e_EffectOverTimeType.ePartyBonus
450         Dim PartyEffect As ApplyEffectToParty
452         Set PartyEffect = GetEOT(EffectType)
456         UniqueIdCounter = GetNextId()
458         Call PartyEffect.Setup(SourceIndex, SourceType, TargetIndex, TargetType, EffectIndex, UniqueIdCounter)
460         Call AddEffectToUpdate(PartyEffect)
462         If TargetType = eUser Then
464             Call AddEffect(UserList(TargetIndex).EffectOverTime, PartyEffect)
468             'npc doesn't have groups
            End If
        Case Else
            Debug.Assert False
    End Select
    Exit Sub
CreateEffect_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateEffect", Erl)
End Sub

Public Sub CreateTrap(ByVal SourceIndex As Integer, ByVal SourceType As e_ReferenceType, ByVal map As Integer, ByVal TileX As Integer, ByVal TileY As Integer, ByVal EffectTypeId As Integer)
On Error GoTo CreateTrap_Err
    Dim EffectType As e_EffectOverTimeType
100 EffectType = e_EffectOverTimeType.eTrap
    Dim Trap As clsTrap
104 Set Trap = GetEOT(EffectType)
106 UniqueIdCounter = GetNextId()
108 Call Trap.Setup(SourceIndex, SourceType, EffectTypeId, UniqueIdCounter, map, TileX, TileY)
110 Call AddEffectToUpdate(Trap)
112 If SourceType = eUser Then
114     Call AddEffect(UserList(SourceIndex).EffectOverTime, Trap)
116 ElseIf SourceType = eNpc Then
118     Call AddEffect(NpcList(SourceIndex).EffectOverTime, Trap)
    End If
    Exit Sub
CreateTrap_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.CreateTrap", Erl)
End Sub

Private Function InstantiateEOT(ByVal EffectType As e_EffectOverTimeType) As IBaseEffectOverTime
    Select Case EffectType
        Case e_EffectOverTimeType.eHealthModifier
            Set InstantiateEOT = New UpdateHpOverTime
        Case e_EffectOverTimeType.eApplyModifiers
            Set InstantiateEOT = New StatModifier
        Case e_EffectOverTimeType.eProvoke
            Set InstantiateEOT = New EffectProvoke
        Case e_EffectOverTimeType.eProvoked
            Set InstantiateEOT = New EffectProvoked
        Case e_EffectOverTimeType.eTrap
            Set InstantiateEOT = New clsTrap
        Case e_EffectOverTimeType.eDrunk
            Set InstantiateEOT = New DrunkEffect
        Case e_EffectOverTimeType.eTranslation
            Set InstantiateEOT = New TranslationEffect
        Case e_EffectOverTimeType.eApplyEffectOnHit
            Set InstantiateEOT = New ApplyEffectOnHit
        Case e_EffectOverTimeType.eManaModifier
            Set InstantiateEOT = New UpdateManaOverTime
        Case e_EffectOverTimeType.ePartyBonus
            Set InstantiateEOT = New ApplyEffectToParty
        Case Else
            Debug.Assert False
    End Select
End Function

Private Function GetEOT(ByVal EffectType As e_EffectOverTimeType) As IBaseEffectOverTime
On Error GoTo GetEOT_Err
100 Set GetEOT = Nothing
102 If EffectPools(EffectType).EffectCount = 0 Then
104     Set GetEOT = InstantiateEOT(EffectType)
        Exit Function
    End If
108 Set GetEOT = EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1)
120 Set EffectPools(EffectType).EffectList(EffectPools(EffectType).EffectCount - 1) = Nothing
126 EffectPools(EffectType).EffectCount = EffectPools(EffectType).EffectCount - 1
    Exit Function
GetEOT_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.GetEOT", Erl)
End Function

Private Sub RecycleEffect(ByRef Effect As IBaseEffectOverTime)
    Call AddEffect(EffectPools(Effect.TypeId), Effect)
End Sub

Public Sub AddEffectToUpdate(ByRef Effect As IBaseEffectOverTime)
On Error GoTo AddEffectToUpdate_Err
    Call AddEffect(ActiveEffects, Effect)
    Exit Sub
AddEffectToUpdate_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.AddEffectToUpdate", Erl)
End Sub

Public Sub AddEffect(ByRef EffectList As t_EffectOverTimeList, ByRef Effect As IBaseEffectOverTime)
On Error GoTo AddEffect_Err
100 If Not IsArrayInitialized(EffectList.EffectList) Then
104     ReDim EffectList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As IBaseEffectOverTime
    ElseIf EffectList.EffectCount >= UBound(EffectList.EffectList) Then
108     ReDim Preserve EffectList.EffectList(EffectList.EffectCount * 1.2) As IBaseEffectOverTime
    End If
116 Set EffectList.EffectList(EffectList.EffectCount) = Effect
120 EffectList.EffectCount = EffectList.EffectCount + 1
    Exit Sub
AddEffect_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.AddEffect", Erl)
End Sub

Public Sub RemoveEffect(ByRef EffectList As t_EffectOverTimeList, ByRef Effect As IBaseEffectOverTime)
On Error GoTo RemoveEffect_Err
    Dim i As Integer
100 For i = 0 To EffectList.EffectCount - 1
106     If EffectList.EffectList(i).UniqueId() = Effect.UniqueId() Then
            Call RemoveEffectAtPos(EffectList, i)
            Exit Sub
        End If
    Next i
    Exit Sub
RemoveEffect_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffect", Erl)
End Sub

Public Function FindEffectOnTarget(ByVal CasterIndex As Integer, ByRef EffectList As t_EffectOverTimeList, ByVal EffectId As Integer) As IBaseEffectOverTime
On Error GoTo FindEffectOnTarget_Err
100 Set FindEffectOnTarget = Nothing
102 Dim EffectLimit As e_EOTTargetLimit
104 EffectLimit = EffectOverTime(EffectId).Limit
106 Dim i As Integer
108 If EffectLimit = e_EOTTargetLimit.eAny Then
        Exit Function
    End If
120 For i = 0 To EffectList.EffectCount - 1
        If EffectLimit = eSingle Or EffectLimit = eSingleByCaster Then
126         If EffectList.EffectList(i).EotId = EffectId Then
130             If EffectLimit = eSingle Then
132                 Set FindEffectOnTarget = EffectList.EffectList(i)
                    Exit Function
                Else
140                 If EffectList.EffectList(i).CasterRefType = eUser Then
142                     If EffectList.EffectList(i).CasterUserId = UserList(CasterIndex).ID Then
144                         Set FindEffectOnTarget = EffectList.EffectList(i)
                            Exit Function
                        End If
150                 ElseIf EffectList.EffectList(i).CasterRefType = eNpc Then
152                     If EffectList.EffectList(i).CasterIsValid Then
154                         Set FindEffectOnTarget = EffectList.EffectList(i)
                            Exit Function
                        End If
                    End If
                End If
            End If
        ElseIf EffectLimit = eSingleByType Then
            If EffectList.EffectList(i).TypeId = EffectOverTime(EffectId).Type Then
                Set FindEffectOnTarget = EffectList.EffectList(i)
                Exit Function
            End If
        End If
    Next i
    Exit Function
FindEffectOnTarget_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.FindEffectOnTarget", Erl)
End Function

Public Sub ClearEffectList(ByRef EffectList As t_EffectOverTimeList, Optional ByVal Filter As e_EffectType = e_EffectType.eAny)
On Error GoTo ClearEffectList_Err
    Dim i As Integer
100 Do While i < EffectList.EffectCount
102     If Filter = e_EffectType.eAny Or Filter = EffectList.EffectList(i).EffectType Then
104         EffectList.EffectList(i).RemoveMe = True
            Call RemoveEffectAtPos(EffectList, i)
        Else
112         i = i + 1
        End If
    Loop
Exit Sub
ClearEffectList_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.ClearEffectList", Erl)
End Sub

Public Sub RemoveEffectAtPos(ByRef EffectList As t_EffectOverTimeList, ByVal position As Integer)
On Error GoTo RemoveEffectAtPos_Err
    Call EffectList.EffectList(position).OnRemove
106 Set EffectList.EffectList(position) = EffectList.EffectList(EffectList.EffectCount - 1)
108 Set EffectList.EffectList(EffectList.EffectCount - 1) = Nothing
110 EffectList.EffectCount = EffectList.EffectCount - 1
RemoveEffectAtPos_Err:
      Call TraceError(Err.Number, Err.Description, "EffectsOverTime.RemoveEffectAtPos", Erl)
End Sub


Public Sub TargetUseMagic(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal MagicId As Integer)
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
         Call EffectList.EffectList(i).TargetUseMagic(TargetUserId, SourceType, MagicId)
    Next i
End Sub

Public Sub TartgetWillAtack(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
         Call EffectList.EffectList(i).TartgetWillAtack(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TartgetDidHit(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
         Call EffectList.EffectList(i).TartgetDidHit(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TargetFailedAttack(ByRef EffectList As t_EffectOverTimeList, ByVal TargetUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
         Call EffectList.EffectList(i).TargetFailedAttack(TargetUserId, SourceType, AttackType)
    Next i
End Sub

Public Sub TargetWasDamaged(ByRef EffectList As t_EffectOverTimeList, ByVal SourceUserId As Integer, ByVal SourceType As e_ReferenceType, ByVal AttackType As e_DamageSourceType)
    Dim i As Integer
    For i = 0 To EffectList.EffectCount - 1
         Call EffectList.EffectList(i).TargetWasDamaged(SourceUserId, SourceType, AttackType)
    Next i
End Sub

Public Function ConvertToClientBuff(ByVal buffType As e_EffectType) As e_EffectType
    Select Case buffType
        Case e_EffectType.eInformativeBuff
            ConvertToClientBuff = eBuff
        Case e_EffectType.eInformativeDebuff
            ConvertToClientBuff = eDebuff
        Case Else
        ConvertToClientBuff = buffType
    End Select
End Function
