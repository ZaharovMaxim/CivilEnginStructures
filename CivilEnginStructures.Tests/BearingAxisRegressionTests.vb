Imports System.Collections
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg
Imports Topomatic.Dwg.Entities

Namespace Tests
    <TestFixture>
    Public Class BearingAxisRegressionTests
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub LastSpanKeepsPenultimateBearingLineAndJsonAtBeamStarts()
            Dim structureType As Type = MainType("CivilEnginStructures.StructureElement")
            Dim beamType As Type = MainType("CivilEnginStructures.BeamI")
            Dim pillarType As Type = MainType("CivilEnginStructures.Pillar")
            Dim bearingType As Type = MainType("CivilEnginStructures.AxisBeamsPillars")
            Dim beams As IDictionary = CreateIntegerDictionary(
                GetType(Dictionary(Of ,)).MakeGenericType(GetType(Integer), structureType))

            beams.Add(1, CreateSpan(structureType, beamType, 0.0R, 10.0R))
            beams.Add(2, CreateSpan(structureType, beamType, 11.0R, 20.0R))

            Dim previousBearing As Object = CreateBearingData(structureType, bearingType, 2, 1)
            Dim nextBearing As Object = CreateBearingData(structureType, bearingType, 2, 2)
            Dim supports As IDictionary = CreateIntegerDictionary(
                GetType(List(Of )).MakeGenericType(structureType))
            supports.Add(1, CreateStructureList(
                structureType,
                Nothing,
                CreatePillarData(structureType, pillarType, 1),
                Nothing))
            supports.Add(2, CreateStructureList(
                structureType,
                previousBearing,
                CreatePillarData(structureType, pillarType, 2),
                nextBearing))
            Dim terminalPillar As Object = CreatePillarData(structureType, pillarType, 3)
            supports.Add(3, CreateStructureList(structureType, Nothing, terminalPillar, Nothing))

            Dim method As MethodInfo = pillarType.GetMethod(
                "calculateAxisBeamsPillar", BindingFlags.Public Or BindingFlags.Static)
            Assert.That(method, [Is].Not.Null)
            Dim succeeded As Boolean = CBool(InvokeWithoutWrapper(
                method, Nothing, New Object() {beams, supports, Nothing}))

            Dim nextLine As DwgLine = DirectCast(ReadProperty(nextBearing, "DWGEntity"), DwgLine)
            Dim nextDto As Object = JsonConvert.DeserializeObject(
                CStr(ReadProperty(nextBearing, "KeyParameter")), bearingType)
            Dim terminalLine As DwgLine = DirectCast(ReadProperty(terminalPillar, "DWGEntity"), DwgLine)
            Dim terminalDto As Object = JsonConvert.DeserializeObject(
                CStr(ReadProperty(terminalPillar, "KeyParameter")), pillarType)

            Assert.Multiple(
                Sub()
                    Assert.That(succeeded, [Is].True)
                    AssertLineAtX(nextLine, 11.0R, "penultimate support next-bearing line")
                    AssertPointsAtX(nextDto, 11.0R, "penultimate support next-bearing JSON")
                    AssertLineAtX(terminalLine, 20.0R, "terminal pillar line")
                    AssertPointsAtX(terminalDto, 20.0R, "terminal pillar JSON")
                End Sub)
        End Sub

        <Test, Explicit("Requires an initialized Topomatic application; New Drawing() blocks outside the host.")>
        <Category("TopomaticIntegration")>
        Public Sub DrawWritesDistinctPreviousAndNextBearingMetadataToTheirOwnEntities()
            Dim structureType As Type = MainType("CivilEnginStructures.StructureElement")
            Dim bearingType As Type = MainType("CivilEnginStructures.AxisBeamsPillars")
            Dim previousBearing As Object = CreateBearingData(structureType, bearingType, 5, 4)
            Dim nextBearing As Object = CreateBearingData(structureType, bearingType, 5, 5)
            Dim previousLine As DwgLine = DirectCast(ReadProperty(previousBearing, "DWGEntity"), DwgLine)
            Dim nextLine As DwgLine = DirectCast(ReadProperty(nextBearing, "DWGEntity"), DwgLine)
            SetLine(previousLine, 10.0R)
            SetLine(nextLine, 11.0R)

            Dim drawing As New Drawing()
            drawing.ActiveSpace.Entities.Add(previousLine)
            drawing.ActiveSpace.Entities.Add(nextLine)
            Dim axes As IDictionary = CreateIntegerDictionary(GetType(List(Of )).MakeGenericType(structureType))
            axes.Add(5, CreateStructureList(structureType, previousBearing, Nothing, nextBearing))
            Dim bridgeObjects As IDictionary = CreateBridgeObjectsDictionary(structureType)
            Dim drawMethod As MethodInfo = bearingType.GetMethod(
                "drawAxisBeamsPillar", BindingFlags.Public Or BindingFlags.Static)
            Assert.That(drawMethod, [Is].Not.Null)

            Dim succeeded As Boolean = CBool(InvokeWithoutWrapper(
                drawMethod,
                Nothing,
                New Object() {drawing, axes, bridgeObjects, Nothing}))
            Dim repeated As Boolean = CBool(InvokeWithoutWrapper(
                drawMethod,
                Nothing,
                New Object() {drawing, axes, bridgeObjects, Nothing}))
            Dim previousStored As Object = ReadStructureMetadata(previousLine, structureType)
            Dim nextStored As Object = ReadStructureMetadata(nextLine, structureType)
            Dim previousDto As Object = JsonConvert.DeserializeObject(
                CStr(ReadProperty(previousStored, "KeyParameter")), bearingType)
            Dim nextDto As Object = JsonConvert.DeserializeObject(
                CStr(ReadProperty(nextStored, "KeyParameter")), bearingType)

            Assert.Multiple(
                Sub()
                    Assert.That(succeeded, [Is].True)
                    Assert.That(repeated, [Is].True)
                    Assert.That(CInt(ReadProperty(previousDto, "numberPillar")), [Is].EqualTo(5))
                    Assert.That(CInt(ReadProperty(previousDto, "numberProlet")), [Is].EqualTo(4))
                    Assert.That(CInt(ReadProperty(nextDto, "numberPillar")), [Is].EqualTo(5))
                    Assert.That(CInt(ReadProperty(nextDto, "numberProlet")), [Is].EqualTo(5))
                End Sub)
        End Sub

        Private Shared Function CreateSpan(structureType As Type,
                                           beamType As Type,
                                           startX As Double,
                                           endX As Double) As IDictionary
            Dim span As IDictionary = DirectCast(Activator.CreateInstance(
                GetType(Dictionary(Of ,)).MakeGenericType(GetType(Integer), structureType)), IDictionary)
            span.Add(1, CreateBeamData(structureType, beamType, 1, startX, endX, 0.0R))
            span.Add(2, CreateBeamData(structureType, beamType, 2, startX, endX, 10.0R))
            Return span
        End Function

        Private Shared Function CreateBeamData(structureType As Type,
                                               beamType As Type,
                                               rowNumber As Integer,
                                               startX As Double,
                                               endX As Double,
                                               y As Double) As Object
            Dim beam As Object = Activator.CreateInstance(beamType)
            WriteProperty(beam, "numberRow", rowNumber)
            SetPoints(beam, New Vector3D(startX, y, 0.0R), New Vector3D(endX, y, 0.0R))
            Dim line As New DwgLine With {
                .StartPoint = New Vector3D(startX, y, 0.0R),
                .EndPoint = New Vector3D(endX, y, 0.0R)
            }
            Return CreateStructureData(structureType, beam, line)
        End Function

        Private Shared Function CreatePillarData(structureType As Type,
                                                 pillarType As Type,
                                                 number As Integer) As Object
            Dim pillar As Object = Activator.CreateInstance(pillarType)
            WriteProperty(pillar, "Number", number)
            SetPoints(pillar, New Vector3D(), New Vector3D())
            Return CreateStructureData(structureType, pillar, New DwgLine())
        End Function

        Private Shared Function CreateBearingData(structureType As Type,
                                                  bearingType As Type,
                                                  pillarNumber As Integer,
                                                  spanNumber As Integer) As Object
            Dim bearing As Object = Activator.CreateInstance(bearingType)
            WriteProperty(bearing, "numberPillar", pillarNumber)
            WriteProperty(bearing, "numberProlet", spanNumber)
            SetPoints(bearing, New Vector3D(), New Vector3D())
            Return CreateStructureData(structureType, bearing, New DwgLine())
        End Function

        Private Shared Function CreateStructureData(structureType As Type,
                                                    dto As Object,
                                                    line As DwgLine) As Object
            Dim data As Object = Activator.CreateInstance(structureType)
            WriteProperty(data, "KeyParameter", JsonConvert.SerializeObject(dto))
            WriteProperty(data, "DWGEntity", line)
            Return data
        End Function

        Private Shared Function CreateIntegerDictionary(valueType As Type) As IDictionary
            Return DirectCast(Activator.CreateInstance(
                GetType(Dictionary(Of ,)).MakeGenericType(GetType(Integer), valueType)), IDictionary)
        End Function

        Private Shared Function CreateBridgeObjectsDictionary(structureType As Type) As IDictionary
            Dim objectKind As Type = structureType.GetNestedType("typeObject")
            Dim listType As Type = GetType(List(Of )).MakeGenericType(structureType)
            Return DirectCast(Activator.CreateInstance(
                GetType(Dictionary(Of ,)).MakeGenericType(objectKind, listType)), IDictionary)
        End Function

        Private Shared Function CreateStructureList(structureType As Type,
                                                    ParamArray items As Object()) As IList
            Dim result As IList = DirectCast(Activator.CreateInstance(
                GetType(List(Of )).MakeGenericType(structureType)), IList)
            For Each item As Object In items
                result.Add(item)
            Next
            Return result
        End Function

        Private Shared Sub SetPoints(dto As Object, startPoint As Vector3D, endPoint As Vector3D)
            Dim points As Object = dto.GetType().GetField(
                "_elementBridgePoint", BindingFlags.Public Or BindingFlags.Instance).GetValue(dto)
            WriteProperty(points, "StartAxisPoint", startPoint)
            WriteProperty(points, "EndAxisPoint", endPoint)
        End Sub

        Private Shared Function ReadPoint(dto As Object, propertyName As String) As Vector3D
            Dim points As Object = dto.GetType().GetField(
                "_elementBridgePoint", BindingFlags.Public Or BindingFlags.Instance).GetValue(dto)
            Return DirectCast(ReadProperty(points, propertyName), Vector3D)
        End Function

        Private Shared Sub AssertLineAtX(line As DwgLine, expectedX As Double, description As String)
            Assert.That(line.StartPoint.X, [Is].EqualTo(expectedX).Within(0.000000001R),
                        description & " start")
            Assert.That(line.EndPoint.X, [Is].EqualTo(expectedX).Within(0.000000001R),
                        description & " end")
        End Sub

        Private Shared Sub AssertPointsAtX(dto As Object, expectedX As Double, description As String)
            Assert.That(ReadPoint(dto, "StartAxisPoint").X,
                        [Is].EqualTo(expectedX).Within(0.000000001R), description & " start")
            Assert.That(ReadPoint(dto, "EndAxisPoint").X,
                        [Is].EqualTo(expectedX).Within(0.000000001R), description & " end")
        End Sub

        Private Shared Sub SetLine(line As DwgLine, x As Double)
            line.StartPoint = New Vector3D(x, 0.0R, 0.0R)
            line.EndPoint = New Vector3D(x, 10.0R, 0.0R)
        End Sub

        Private Shared Function ReadStructureMetadata(line As DwgLine, structureType As Type) As Object
            Dim xrecordsType As Type = MainType("CivilEnginStructures.FuncXRecords")
            Dim tableType As Type = structureType.GetNestedType("tableXRecords")
            Dim projectStructures As Object = [Enum].Parse(tableType, "PROJECT_STRUCTURES")
            Dim method As MethodInfo = xrecordsType.GetMethod(
                "getXRecords",
                BindingFlags.Public Or BindingFlags.Static,
                binder:=Nothing,
                types:=New Type() {GetType(DwgEntity), structureType.MakeByRefType(), tableType},
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null)
            Dim arguments As Object() = {line, Nothing, projectStructures}
            Dim found As Boolean = CBool(InvokeWithoutWrapper(method, Nothing, arguments))
            Assert.That(found, [Is].True, "Each bearing line must receive its own structure metadata.")
            Return arguments(1)
        End Function

        Private Shared Function ReadProperty(target As Object, propertyName As String) As Object
            Return target.GetType().GetProperty(propertyName).GetValue(target, Nothing)
        End Function

        Private Shared Sub WriteProperty(target As Object, propertyName As String, value As Object)
            target.GetType().GetProperty(propertyName).SetValue(target, value, Nothing)
        End Sub

        Private Shared Function InvokeWithoutWrapper(method As MethodInfo,
                                                     target As Object,
                                                     arguments As Object()) As Object
            Try
                Return method.Invoke(target, arguments)
            Catch exception As TargetInvocationException
                If exception.InnerException IsNot Nothing Then Throw exception.InnerException
                Throw
            End Try
        End Function

        Private Shared Function MainType(fullName As String) As Type
            Dim result As Type = MainAssembly().GetType(fullName, throwOnError:=False)
            Assert.That(result, [Is].Not.Null, fullName & " must exist in the production assembly.")
            Return result
        End Function

        Private Shared Function MainAssembly() As Assembly
            If _mainAssembly IsNot Nothing Then Return _mainAssembly

            Dim root As String = RepositoryRoot()
            Dim candidates As String() = {
                Path.Combine(root, "bin", "Debug", "CivilEnginStructures.dll"),
                Path.Combine(root, "bin", "Release", "CivilEnginStructures.dll")
            }
            Dim assemblyPath As String = candidates.
                Where(Function(candidate) File.Exists(candidate)).
                OrderByDescending(Function(candidate) File.GetLastWriteTimeUtc(candidate)).
                FirstOrDefault()
            Assert.That(assemblyPath, [Is].Not.Null,
                        "Build CivilEnginStructures.vbproj before running bearing-axis regression tests.")
            _mainAssembly = Assembly.LoadFrom(assemblyPath)
            Return _mainAssembly
        End Function

        Private Shared Function RepositoryRoot() As String
            Dim directory As New DirectoryInfo(TestContext.CurrentContext.TestDirectory)
            While directory IsNot Nothing
                If File.Exists(Path.Combine(directory.FullName, "CivilEnginStructures.vbproj")) Then
                    Return directory.FullName
                End If
                directory = directory.Parent
            End While
            Assert.Fail("Could not locate the repository root from the test directory.")
            Return Nothing
        End Function
    End Class
End Namespace
