Imports Topomatic.Dwg.Entities

Friend NotInheritable Class BridgeTaggedEntityRecord
    Friend Sub New(idStructure As String)
        Me.IdStructure = If(idStructure, String.Empty).Trim()
    End Sub

    Friend ReadOnly Property IdStructure As String
    Friend Property Bridge As Bridges
    Friend ReadOnly Property Entities As New HashSet(Of DwgEntity)()
End Class

Friend NotInheritable Class BridgeTaggedEntityIndex
    Private Sub New()
    End Sub

    Friend Shared Function Build(entities As IEnumerable(Of DwgEntity)) As Dictionary(Of String, BridgeTaggedEntityRecord)
        Dim result As New Dictionary(Of String, BridgeTaggedEntityRecord)(StringComparer.Ordinal)
        If entities Is Nothing Then Return result

        For Each entity As DwgEntity In entities
            Dim data As StructureElement = Nothing
            If Not TryReadStructureData(entity, data) Then Continue For

            Dim canonicalId As String = BridgeDrawingGroupManager.GetGroupName(data.IdStructure)
            If canonicalId.Length = 0 Then Continue For

            Dim record As BridgeTaggedEntityRecord = Nothing
            If Not result.TryGetValue(canonicalId, record) Then
                record = New BridgeTaggedEntityRecord(data.IdStructure)
                result.Add(canonicalId, record)
            End If
            record.Entities.Add(entity)

            If record.Bridge Is Nothing AndAlso
               data.Name = StructureElement.typeObject.axisBridge Then
                record.Bridge = ReadBridge(data)
            End If
        Next

        Return result
    End Function

    Friend Shared Function TryReadStructureData(entity As DwgEntity,
                                                 ByRef data As StructureElement) As Boolean
        If entity Is Nothing Then Return False
        Try
            Return FuncXRecords.getXRecords(entity, data) AndAlso
                   data IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(data.IdStructure)
        Catch
            data = Nothing
            Return False
        End Try
    End Function

    Private Shared Function ReadBridge(data As StructureElement) As Bridges
        Try
            Return data.getBridge()
        Catch
            Return Nothing
        End Try
    End Function
End Class
