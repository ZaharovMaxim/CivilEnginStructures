Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Dwg.Entities

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class BridgePropertyRoundTripTests
        Private Const Tolerance As Double = 0.000000001R
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub BeamOffsetsSurviveRussianJsonAndGlobalOffsetEditsWithoutAccumulation()
            Dim beamType As Type = MainType("CivilEnginStructures.BeamI")
            Dim leftBeam As Object = CreateBeam(beamType, -0.71R, 0.15R)
            Dim rightBeam As Object = CreateBeam(beamType, 1.21R, 0.15R)
            Dim previousCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
            Dim restoredLeft As Object
            Dim restoredRight As Object

            Try
                Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU")
                restoredLeft = JsonRoundTripBeam(leftBeam, beamType)
                restoredRight = JsonRoundTripBeam(rightBeam, beamType)
            Finally
                Thread.CurrentThread.CurrentCulture = previousCulture
            End Try

            ApplyBeamOffsetDelta(restoredLeft, 0.25R, 0.4R, 0.05R, 0.08R)
            ApplyBeamOffsetDelta(restoredRight, 0.25R, 0.4R, 0.05R, 0.08R)

            Assert.Multiple(
                Sub()
                    AssertBeamOffsets(restoredLeft, -0.56R, 0.18R, "left edited beam")
                    AssertBeamOffsets(restoredRight, 1.36R, 0.18R, "right edited beam")
                    Assert.That(GetRebuildRowOffsets(restoredLeft),
                                [Is].EqualTo(New Double() {-0.56R, 0.18R}))
                    Assert.That(GetRebuildRowOffsets(restoredLeft),
                                [Is].EqualTo(New Double() {-0.56R, 0.18R}),
                                "Reading rebuild offsets repeatedly must not add the global offsets.")
                    Assert.That(GetRebuildRowOffsets(restoredRight),
                                [Is].EqualTo(New Double() {1.36R, 0.18R}))
                    Assert.That(FormPlacementBeams.GetSavedRowValues(-0.56R, 0.18R, 0.4R, 0.08R),
                                [Is].EqualTo(New Double() {960.0R, 100.0R}))
                    Assert.That(FormPlacementBeams.GetSavedRowValues(1.36R, 0.18R, 0.4R, 0.08R),
                                [Is].EqualTo(New Double() {960.0R, 100.0R}))
                End Sub)

            ApplyBeamOffsetDelta(restoredLeft, 0.4R, 0.25R, 0.08R, 0.05R)
            ApplyBeamOffsetDelta(restoredRight, 0.4R, 0.25R, 0.08R, 0.05R)

            Assert.Multiple(
                Sub()
                    AssertBeamOffsets(restoredLeft, -0.71R, 0.15R, "left restored beam")
                    AssertBeamOffsets(restoredRight, 1.21R, 0.15R, "right restored beam")
                End Sub)
        End Sub

        <Test>
        Public Sub BridgeCountsAreReadOnlyInPropertyGridButRemainSettableAndSerializable()
            Dim bridgeType As Type = MainType("CivilEnginStructures.Bridges")
            Dim json As String = "{""ProletCount"":4,""LeftRowsCount"":2,""RightRowsCount"":5}"
            Dim bridge As Object = JsonConvert.DeserializeObject(json, bridgeType)
            Dim descriptors As PropertyDescriptorCollection = TypeDescriptor.GetProperties(bridge)

            Assert.Multiple(
                Sub()
                    AssertCountContract(bridgeType, descriptors, "ProletCount", 4, bridge)
                    AssertCountContract(bridgeType, descriptors, "LeftRowsCount", 2, bridge)
                    AssertCountContract(bridgeType, descriptors, "RightRowsCount", 5, bridge)
                End Sub)

            WriteProperty(bridge, "ProletCount", 6)
            WriteProperty(bridge, "LeftRowsCount", 3)
            WriteProperty(bridge, "RightRowsCount", 7)
            Dim persisted As JObject = JObject.Parse(JsonConvert.SerializeObject(bridge))

            Assert.Multiple(
                Sub()
                    AssertNumericJsonProperty(persisted, "ProletCount", 6)
                    AssertNumericJsonProperty(persisted, "LeftRowsCount", 3)
                    AssertNumericJsonProperty(persisted, "RightRowsCount", 7)
                End Sub)
        End Sub

        <TestCase("LeftStructureWidth", Double.NaN)>
        <TestCase("RightStructureWidth", Double.NaN)>
        <TestCase("startPlacementPosition", Double.NaN)>
        <TestCase("HorizontalOffset", Double.NaN)>
        <TestCase("TransverseOffset", Double.NaN)>
        <TestCase("VerticalOffset", Double.NaN)>
        <TestCase("HorizontalOffset", Double.PositiveInfinity)>
        <TestCase("VerticalOffset", Double.NegativeInfinity)>
        Public Sub NonFiniteBridgePlacementValueIsRejectedWithoutMutatingDto(propertyName As String,
                                                                             invalidValue As Double)
            Dim bridgeType As Type = MainType("CivilEnginStructures.Bridges")
            Dim bridge As Object = Activator.CreateInstance(bridgeType)
            WriteProperty(bridge, propertyName, invalidValue)
            Dim beforeJson As String = JsonConvert.SerializeObject(bridge)
            Dim method As MethodInfo = FindMainMethod(
                "CivilEnginStructures.Bridges",
                "EnsureFinitePlacementValues",
                New Type() {bridgeType})

            Dim exception As Exception = Assert.Catch(Of Exception)(
                Sub()
                    InvokeWithoutWrapper(method, Nothing, New Object() {bridge})
                End Sub)

            Assert.Multiple(
                Sub()
                    Assert.That(exception.GetType().Name, [Is].EqualTo("BuildStageException"))
                    Assert.That(JsonConvert.SerializeObject(bridge), [Is].EqualTo(beforeJson),
                                propertyName & " guard must not mutate the bridge DTO.")
                End Sub)
        End Sub

        <Test>
        Public Sub FiniteBridgePlacementValuesPassGuardWithoutMutatingDto()
            Dim bridgeType As Type = MainType("CivilEnginStructures.Bridges")
            Dim bridge As Object = Activator.CreateInstance(bridgeType)
            WriteProperty(bridge, "LeftStructureWidth", 8.5R)
            WriteProperty(bridge, "RightStructureWidth", 9.25R)
            WriteProperty(bridge, "startPlacementPosition", 0.0R)
            WriteProperty(bridge, "HorizontalOffset", 10.0R)
            WriteProperty(bridge, "TransverseOffset", -0.4R)
            WriteProperty(bridge, "VerticalOffset", 0.08R)
            Dim beforeJson As String = JsonConvert.SerializeObject(bridge)
            Dim method As MethodInfo = FindMainMethod(
                "CivilEnginStructures.Bridges",
                "EnsureFinitePlacementValues",
                New Type() {bridgeType})

            Assert.DoesNotThrow(
                Sub()
                    InvokeWithoutWrapper(method, Nothing, New Object() {bridge})
                End Sub)
            Assert.That(JsonConvert.SerializeObject(bridge), [Is].EqualTo(beforeJson))
        End Sub

        <TestCase(0.0R, 10.0R)>
        <TestCase(10.0R, 0.0R)>
        Public Sub RebuildPreflightRejectsEitherZeroBridgeWidthBeforeCheckingBeams(
            leftWidth As Double,
            rightWidth As Double)

            Dim bridgeType As Type = MainType("CivilEnginStructures.Bridges")
            Dim structureType As Type = MainType("CivilEnginStructures.StructureElement")
            Dim operationType As Type = MainType("CivilEnginStructures.BuildOperationContext")
            Dim panelType As Type = MainType("CivilEnginStructures.PanelProjectBridge")
            Dim innerDictionaryType As Type = GetType(System.Collections.Generic.Dictionary(Of ,)).
                MakeGenericType(GetType(Integer), structureType)
            Dim bridgeBeamsType As Type = GetType(System.Collections.Generic.Dictionary(Of ,)).
                MakeGenericType(GetType(Integer), innerDictionaryType)
            Dim bridge As Object = Activator.CreateInstance(bridgeType)
            WriteProperty(bridge, "LeftStructureWidth", leftWidth)
            WriteProperty(bridge, "RightStructureWidth", rightWidth)
            WriteProperty(bridge, "ProletCount", 1)
            WriteProperty(bridge, "LeftRowsCount", 1)
            WriteProperty(bridge, "RightRowsCount", 1)
            Dim emptyBeams As Object = Activator.CreateInstance(bridgeBeamsType)
            Dim operation As Object = Activator.CreateInstance(
                operationType,
                New Object() {"Проверка перестройки", "Проверьте параметры моста.", "regression-test"})
            Dim method As MethodInfo = panelType.GetMethod(
                "ValidateBridgeRebuildInputs",
                BindingFlags.Static Or BindingFlags.Public Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=New Type() {bridgeType, bridgeBeamsType, operationType},
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null,
                        "PanelProjectBridge.ValidateBridgeRebuildInputs preflight method")

            Dim exception As Exception = Assert.Catch(Of Exception)(
                Sub()
                    InvokeWithoutWrapper(method, Nothing, New Object() {bridge, emptyBeams, operation})
                End Sub)

            Assert.Multiple(
                Sub()
                    Assert.That(exception.GetType().Name, [Is].EqualTo("BuildStageException"))
                    Assert.That(exception.Message.ToLowerInvariant(), Does.Contain("ширин"),
                                "An invalid side width must be reported before missing beams.")
                End Sub)
        End Sub

        <TestCase(100.0R, 5.0R, 103.0R, 2.0R)>
        <TestCase(100.0R, 5.0R, 105.0R, 0.0R)>
        <TestCase(-100.0R, -5.0R, -103.0R, -2.0R)>
        <TestCase(100.1236R, -0.0002R, 100.123R, 0.0R)>
        Public Sub LongitudinalRebuildDeltaUsesSavedAbsoluteStationWithoutAccumulation(
            startStation As Double,
            horizontalOffset As Double,
            currentStation As Double,
            expectedDelta As Double)

            Assert.That(GetLongitudinalRebuildDelta(startStation, horizontalOffset, currentStation),
                        [Is].EqualTo(expectedDelta).Within(Tolerance))
        End Sub

        <Test>
        Public Sub ZeroSavedStartStationRemainsAuthoritativeAcrossRepeatedRebuilds()
            Dim currentStation As Double = 0.0R

            Dim firstDelta As Double = GetLongitudinalRebuildDelta(0.0R, 5.0R, currentStation)
            currentStation += firstDelta
            Dim firstRepeat As Double = GetLongitudinalRebuildDelta(0.0R, 5.0R, currentStation)
            Dim secondDelta As Double = GetLongitudinalRebuildDelta(0.0R, 10.0R, currentStation)
            currentStation += secondDelta
            Dim secondRepeat As Double = GetLongitudinalRebuildDelta(0.0R, 10.0R, currentStation)

            Assert.Multiple(
                Sub()
                    Assert.That(firstDelta, [Is].EqualTo(5.0R).Within(Tolerance),
                                "PK0 + 5 from current PK0")
                    Assert.That(firstRepeat, [Is].EqualTo(0.0R).Within(Tolerance),
                                "repeating PK0 + 5 at current PK5")
                    Assert.That(secondDelta, [Is].EqualTo(5.0R).Within(Tolerance),
                                "editing the offset to 10 from current PK5")
                    Assert.That(secondRepeat, [Is].EqualTo(0.0R).Within(Tolerance),
                                "repeating PK0 + 10 at current PK10")
                    Assert.That(currentStation, [Is].EqualTo(10.0R).Within(Tolerance))
                End Sub)
        End Sub

        <Test>
        Public Sub DetachedAxisTranslationPreservesShapeElevationAndIsStableAtCurrentTarget()
            Dim axis As New DwgLine With {
                .StartPoint = New Vector3D(100.0R, 8.0R, 7.25R),
                .EndPoint = New Vector3D(106.0R, 12.0R, 11.75R)
            }
            Dim originalDelta As New Vector3D(
                axis.EndPoint.X - axis.StartPoint.X,
                axis.EndPoint.Y - axis.StartPoint.Y,
                axis.EndPoint.Z - axis.StartPoint.Z)
            Dim originalLength As Double = axis.Length

            TranslateAxisBetweenPoints(axis, New Vector2D(103.0R, 10.0R), New Vector2D(105.0R, 10.0R))
            Dim firstStart As Vector3D = axis.StartPoint
            Dim firstEnd As Vector3D = axis.EndPoint
            TranslateAxisBetweenPoints(axis, New Vector2D(105.0R, 10.0R), New Vector2D(105.0R, 10.0R))

            Assert.Multiple(
                Sub()
                    Assert.That(firstStart.X, [Is].EqualTo(102.0R).Within(Tolerance))
                    Assert.That(firstStart.Y, [Is].EqualTo(8.0R).Within(Tolerance))
                    Assert.That(firstStart.Z, [Is].EqualTo(7.25R).Within(Tolerance))
                    Assert.That(firstEnd.X, [Is].EqualTo(108.0R).Within(Tolerance))
                    Assert.That(firstEnd.Y, [Is].EqualTo(12.0R).Within(Tolerance))
                    Assert.That(firstEnd.Z, [Is].EqualTo(11.75R).Within(Tolerance))
                    Assert.That(firstEnd.X - firstStart.X, [Is].EqualTo(originalDelta.X).Within(Tolerance))
                    Assert.That(firstEnd.Y - firstStart.Y, [Is].EqualTo(originalDelta.Y).Within(Tolerance))
                    Assert.That(firstEnd.Z - firstStart.Z, [Is].EqualTo(originalDelta.Z).Within(Tolerance))
                    Assert.That(axis.Length, [Is].EqualTo(originalLength).Within(Tolerance))
                    AssertVector(axis.StartPoint, firstStart, "repeated start")
                    AssertVector(axis.EndPoint, firstEnd, "repeated end")
                End Sub)
        End Sub

        <Test>
        Public Sub ElevationCalculationFailureIdentifiesBeamAndSurfaceWhileSuccessLeavesBeamUnchanged()
            Dim beamType As Type = MainType("CivilEnginStructures.BeamI")
            Dim beam As Object = CreateBeam(beamType, -0.71R, 0.15R)
            WriteProperty(beam, "numberProlet", 3)
            WriteProperty(beam, "numberRow", -2)
            Dim beforeJson As String = JsonConvert.SerializeObject(beam)
            Dim method As MethodInfo = FindMainMethod(
                "CivilEnginStructures.CalculationBeams",
                "EnsureElevationCalculated",
                New Type() {GetType(Boolean), beamType})

            Assert.DoesNotThrow(
                Sub()
                    InvokeWithoutWrapper(method, Nothing, New Object() {True, beam})
                End Sub)
            Assert.That(JsonConvert.SerializeObject(beam), [Is].EqualTo(beforeJson),
                        "A successful elevation calculation must not mutate the beam.")

            Dim exception As Exception = Assert.Catch(Of Exception)(
                Sub()
                    InvokeWithoutWrapper(method, Nothing, New Object() {False, beam})
                End Sub)
            Assert.Multiple(
                Sub()
                    Assert.That(exception.GetType().Name, [Is].EqualTo("BuildStageException"))
                    Assert.That(exception.Message, Does.Contain("3"), "span number")
                    Assert.That(exception.Message, Does.Contain("-2"), "row number")
                    Assert.That(exception.Message.ToLowerInvariant(), Does.Contain("поверх"),
                                "surface diagnostic hint")
                    Assert.That(JsonConvert.SerializeObject(beam), [Is].EqualTo(beforeJson),
                                "A failed guard must not conceal itself by mutating the beam.")
                End Sub)
        End Sub

        <Test>
        Public Sub PlacementFormRestoresSignedUnboundedOffsetsAndKeepsNineArgumentCompatibility()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                InvokeFormMethod(form, "ApplyModernAppearance", Type.EmptyTypes, New Object() {})

                Dim nineArgumentTypes As Type() = {
                    GetType(Integer), GetType(Double), GetType(Double), GetType(Double),
                    GetType(Integer), GetType(Integer), GetType(Integer), GetType(Double), GetType(Double)
                }
                Dim commonValues As Object() = {0, 0.0R, -0.25R, 0.0R, 1, 0, 0, 0.0R, 0.0R}
                InvokeFormMethod(form, "RestoreSavedPlacementOptions", nineArgumentTypes, commonValues)
                Assert.That(form.NUpD_VerticalOffset.Value, [Is].EqualTo(-250D))

                commonValues(2) = 0.5R
                InvokeFormMethod(form, "RestoreSavedPlacementOptions", nineArgumentTypes, commonValues)
                Assert.That(form.NUpD_VerticalOffset.Value, [Is].EqualTo(500D))

                Dim longitudinal As NumericUpDown = GetFormNumeric(form, "NUpD_LongitudinalOffset")
                InvokeFormMethod(form,
                                 "RestoreSavedLongitudinalOffset",
                                 New Type() {GetType(Double)},
                                 New Object() {5.0R})
                Assert.That(longitudinal.Value, [Is].EqualTo(5000D))

                InvokeFormMethod(form,
                                 "RestoreSavedLongitudinalOffset",
                                 New Type() {GetType(Double)},
                                 New Object() {-2.5R})
                Assert.That(longitudinal.Value, [Is].EqualTo(-2500D))
            End Using
        End Sub

        <Test>
        Public Sub ExplicitSavedZeroStartStationIsDisplayedAsSet()
            Using form As FormPlacementBeams = FormPlacementBeams.CreateDesignerOnly()
                InvokeFormMethod(
                    form,
                    "RestoreSavedPlacementOptions",
                    New Type() {
                        GetType(Integer), GetType(Double), GetType(Double), GetType(Double),
                        GetType(Integer), GetType(Integer), GetType(Integer), GetType(Double), GetType(Double)
                    },
                    New Object() {0, 0.0R, 0.0R, 0.0R, 1, 0, 0, 0.0R, 0.0R})

                Assert.That(form.CheckBox3.Checked, [Is].False,
                            "The existing nine-argument compatibility path cannot distinguish a draft zero.")
                InvokeFormMethod(
                    form,
                    "RestoreSavedStartStation",
                    New Type() {GetType(Double)},
                    New Object() {0.0R})

                Dim parsedStation As Double
                Assert.Multiple(
                    Sub()
                        Assert.That(form.CheckBox3.Checked, [Is].True)
                        Assert.That(form.MaskTB_PK.Enabled, [Is].True)
                        Assert.That(FuncFormatZn.TryParsePKText(form.MaskTB_PK.Text, parsedStation), [Is].True)
                        Assert.That(parsedStation, [Is].EqualTo(0.0R).Within(Tolerance))
                    End Sub)
            End Using
        End Sub

        Private Shared Function CreateBeam(beamType As Type,
                                           axisOffset As Double,
                                           offsetSurface As Double) As Object
            Dim beam As Object = Activator.CreateInstance(beamType)
            WriteProperty(beam, "axisOffset", axisOffset)
            WriteProperty(beam, "offsetSurface", offsetSurface)
            Return beam
        End Function

        Private Shared Function JsonRoundTripBeam(beam As Object, beamType As Type) As Object
            Dim json As String = JsonConvert.SerializeObject(beam)
            Dim parsed As JObject = JObject.Parse(json)
            AssertNumericJsonProperty(parsed, "axisOffset", ReadDouble(beam, "axisOffset"))
            AssertNumericJsonProperty(parsed, "offsetSurface", ReadDouble(beam, "offsetSurface"))
            Return JsonConvert.DeserializeObject(json, beamType)
        End Function

        Private Shared Sub AssertNumericJsonProperty(json As JObject,
                                                     propertyName As String,
                                                     expectedValue As Double)
            Dim token As JToken = json(propertyName)
            Assert.That(token, [Is].Not.Null, propertyName & " JSON property")
            Assert.That(token.Type,
                        [Is].AnyOf(JTokenType.Float, JTokenType.Integer),
                        propertyName & " must remain a JSON number under ru-RU culture.")
            Assert.That(CDbl(token), [Is].EqualTo(expectedValue).Within(Tolerance))
        End Sub

        Private Shared Sub AssertCountContract(bridgeType As Type,
                                               descriptors As PropertyDescriptorCollection,
                                               propertyName As String,
                                               expectedValue As Integer,
                                               bridge As Object)
            Dim descriptor As PropertyDescriptor = descriptors(propertyName)
            Dim [property] As PropertyInfo = bridgeType.GetProperty(propertyName)
            Assert.That(descriptor, [Is].Not.Null, propertyName & " descriptor")
            Assert.That(descriptor.IsBrowsable, [Is].True, propertyName & " must remain visible")
            Assert.That(descriptor.IsReadOnly, [Is].True, propertyName & " must be read-only in PropertyGrid")
            Assert.That([property], [Is].Not.Null)
            Assert.That([property].CanWrite, [Is].True, propertyName & " public setter is required by the form and JSON")
            Assert.That(CInt([property].GetValue(bridge, Nothing)), [Is].EqualTo(expectedValue))
        End Sub

        Private Shared Sub ApplyBeamOffsetDelta(beam As Object,
                                                oldTransverse As Double,
                                                newTransverse As Double,
                                                oldVertical As Double,
                                                newVertical As Double)
            Dim method As MethodInfo = FindMainMethod(
                "CivilEnginStructures.Bridges",
                "ApplyBeamOffsetDelta",
                New Type() {beam.GetType(), GetType(Double), GetType(Double), GetType(Double), GetType(Double)})
            InvokeWithoutWrapper(method,
                                 Nothing,
                                 New Object() {beam, oldTransverse, newTransverse, oldVertical, newVertical})
        End Sub

        Private Shared Function GetRebuildRowOffsets(beam As Object) As Double()
            Dim method As MethodInfo = FindMainMethod(
                "CivilEnginStructures.Bridges",
                "GetRebuildRowOffsets",
                New Type() {beam.GetType()})
            Return DirectCast(InvokeWithoutWrapper(method, Nothing, New Object() {beam}), Double())
        End Function

        Private Shared Function GetLongitudinalRebuildDelta(startStation As Double,
                                                            horizontalOffset As Double,
                                                            currentStation As Double) As Double
            Dim method As MethodInfo = FindMainMethod(
                "CivilEnginStructures.Bridges",
                "GetLongitudinalRebuildDelta",
                New Type() {GetType(Double), GetType(Double), GetType(Double)})
            Return CDbl(InvokeWithoutWrapper(
                method, Nothing, New Object() {startStation, horizontalOffset, currentStation}))
        End Function

        Private Shared Sub TranslateAxisBetweenPoints(axis As DwgLine,
                                                      sourceCenter As Vector2D,
                                                      targetCenter As Vector2D)
            Dim method As MethodInfo = FindMainMethod(
                "CivilEnginStructures.Pillar",
                "TranslateAxisBetweenPoints",
                New Type() {GetType(DwgLine), GetType(Vector2D), GetType(Vector2D)})
            InvokeWithoutWrapper(method, Nothing, New Object() {axis, sourceCenter, targetCenter})
        End Sub

        Private Shared Function FindMainMethod(typeName As String,
                                               methodName As String,
                                               parameterTypes As Type()) As MethodInfo
            Dim method As MethodInfo = MainType(typeName).GetMethod(
                methodName,
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Static,
                binder:=Nothing,
                types:=parameterTypes,
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null, typeName & "." & methodName & " contract method")
            Return method
        End Function

        Private Shared Sub InvokeFormMethod(form As FormPlacementBeams,
                                            methodName As String,
                                            parameterTypes As Type(),
                                            arguments As Object())
            Dim method As MethodInfo = GetType(FormPlacementBeams).GetMethod(
                methodName,
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=parameterTypes,
                modifiers:=Nothing)
            Assert.That(method, [Is].Not.Null, "FormPlacementBeams." & methodName & " contract method")
            InvokeWithoutWrapper(method, form, arguments)
        End Sub

        Private Shared Function GetFormNumeric(form As FormPlacementBeams,
                                               fieldName As String) As NumericUpDown
            Dim field As FieldInfo = GetType(FormPlacementBeams).GetField(
                fieldName,
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic)
            Assert.That(field, [Is].Not.Null, fieldName & " designer field")
            Dim control As NumericUpDown = TryCast(field.GetValue(form), NumericUpDown)
            Assert.That(control, [Is].Not.Null, fieldName & " must be a NumericUpDown")
            Assert.That(control.Name, [Is].EqualTo(fieldName))
            Return control
        End Function

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

        Private Shared Sub AssertBeamOffsets(beam As Object,
                                             expectedAxisOffset As Double,
                                             expectedSurfaceOffset As Double,
                                             description As String)
            Assert.That(ReadDouble(beam, "axisOffset"),
                        [Is].EqualTo(expectedAxisOffset).Within(Tolerance),
                        description & " axis offset")
            Assert.That(ReadDouble(beam, "offsetSurface"),
                        [Is].EqualTo(expectedSurfaceOffset).Within(Tolerance),
                        description & " surface offset")
        End Sub

        Private Shared Sub AssertVector(actual As Vector3D,
                                        expected As Vector3D,
                                        description As String)
            Assert.That(actual.X, [Is].EqualTo(expected.X).Within(Tolerance), description & " X")
            Assert.That(actual.Y, [Is].EqualTo(expected.Y).Within(Tolerance), description & " Y")
            Assert.That(actual.Z, [Is].EqualTo(expected.Z).Within(Tolerance), description & " Z")
        End Sub

        Private Shared Function ReadDouble(target As Object, propertyName As String) As Double
            Return CDbl(target.GetType().GetProperty(propertyName).GetValue(target, Nothing))
        End Function

        Private Shared Sub WriteProperty(target As Object, propertyName As String, value As Object)
            target.GetType().GetProperty(propertyName).SetValue(target, value, Nothing)
        End Sub

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
                        "Build CivilEnginStructures.vbproj before running bridge property round-trip tests.")
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
