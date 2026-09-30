Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports NUnit.Framework
Imports Topomatic.Cad.Foundation
Imports Topomatic.Cad.View
Imports Topomatic.Cad.View.Hints
Imports Topomatic.Stg

Namespace Tests
    <TestFixture>
    Public Class BridgeSelectionSetTests
        Private Shared _mainAssembly As Assembly

        <Test>
        Public Sub MixedNativeSelectionContainingBridgeContainerDoesNotSupportTransform()
            Dim layer As Object = CreateLayer()
            Dim inner As New RecordingSelectionSet(DirectCast(layer, CadViewLayer))
            inner.Items.Add(New Object())
            inner.Items.Add(CreatePresentationObject("bridge-a"))
            Dim wrapper As SelectionSet = CreateWrapper(layer, inner)

            Assert.Multiple(
                Sub()
                    Assert.That(wrapper.Count, [Is].EqualTo(2))
                    Assert.That(wrapper.SupportTransform, [Is].False)
                    Assert.That(Sub() wrapper.GetTransformData(), Throws.TypeOf(Of NotSupportedException)())
                End Sub)
        End Sub

        <Test>
        Public Sub NativeTransformOperationsDelegateToInnerSelectionWhenCustomSelectionIsEmpty()
            Dim layer As Object = CreateLayer()
            Dim inner As New RecordingSelectionSet(DirectCast(layer, CadViewLayer))
            Dim wrapper As SelectionSet = CreateWrapper(layer, inner)
            Dim sentinel As New Object()
            inner.TransformData = sentinel

            Dim actual As Object = wrapper.GetTransformData()
            wrapper.PaintTransformData(sentinel, Nothing)
            wrapper.Transform(sentinel, Matrix.Identity, True)

            Assert.Multiple(
                Sub()
                    Assert.That(actual, [Is].SameAs(sentinel))
                    Assert.That(inner.GetTransformDataCalls, [Is].EqualTo(1))
                    Assert.That(inner.PaintTransformDataCalls, [Is].EqualTo(1))
                    Assert.That(inner.PaintedTransformData, [Is].SameAs(sentinel))
                    Assert.That(inner.TransformCalls, [Is].EqualTo(1))
                    Assert.That(inner.TransformedData, [Is].SameAs(sentinel))
                    Assert.That(inner.TransformMatrix, [Is].EqualTo(Matrix.Identity))
                    Assert.That(inner.TransformCopy, [Is].True)
                End Sub)
        End Sub

        <Test>
        Public Sub NativeClipboardOperationsDelegateToInnerSelectionWhenCustomSelectionIsEmpty()
            Dim layer As Object = CreateLayer()
            Dim inner As New RecordingSelectionSet(DirectCast(layer, CadViewLayer))
            Dim wrapper As SelectionSet = CreateWrapper(layer, inner)
            Dim sentinel As New ClipboardSentinel()
            inner.ClipboardData = sentinel
            Dim basePoint As New Vector3D(10, -20, 30)

            Dim actual As Object = ClipboardMethod(wrapper, "GetClipboardData", Type.EmptyTypes).
                Invoke(wrapper, Nothing)
            ClipboardMethod(
                wrapper,
                "PaintClipboardData",
                New Type() {GetType(IStgSerializable), GetType(CadPen)}).
                Invoke(wrapper, New Object() {sentinel, Nothing})
            ClipboardMethod(
                wrapper,
                "AddClipboardData",
                New Type() {
                    GetType(IStgSerializable),
                    GetType(Matrix),
                    GetType(Vector3D),
                    GetType(Boolean)
                }).Invoke(wrapper, New Object() {sentinel, Matrix.Identity, basePoint, True})

            Assert.Multiple(
                Sub()
                    Assert.That(wrapper.SupportClipboard, [Is].True)
                    Assert.That(actual, [Is].SameAs(sentinel))
                    Assert.That(inner.GetClipboardDataCalls, [Is].EqualTo(1))
                    Assert.That(inner.PaintClipboardDataCalls, [Is].EqualTo(1))
                    Assert.That(inner.PaintedClipboardData, [Is].SameAs(sentinel))
                    Assert.That(inner.AddClipboardDataCalls, [Is].EqualTo(1))
                    Assert.That(inner.AddedClipboardData, [Is].SameAs(sentinel))
                    Assert.That(inner.AddClipboardTransform, [Is].EqualTo(Matrix.Identity))
                    Assert.That(inner.AddClipboardBasePoint, [Is].EqualTo(basePoint))
                    Assert.That(inner.AddClipboardSelect, [Is].True)
                End Sub)
        End Sub

        <Test>
        Public Sub ReconcileRemovesPresentationObjectMissingFromCurrentLayer()
            Dim layer As Object = CreateLayer()
            Dim inner As New RecordingSelectionSet(DirectCast(layer, CadViewLayer))
            Dim wrapper As SelectionSet = CreateWrapper(layer, inner)
            Dim presentationObject As Object = CreatePresentationObject("removed-bridge")
            wrapper.Select(presentationObject, True)
            Assert.That(wrapper.Count, [Is].EqualTo(1))

            Dim reconcile As MethodInfo = wrapper.GetType().GetMethod(
                "Reconcile", BindingFlags.NonPublic Or BindingFlags.Instance)
            Assert.That(reconcile, [Is].Not.Null)
            reconcile.Invoke(wrapper, Nothing)

            Assert.Multiple(
                Sub()
                    Assert.That(wrapper.Count, [Is].EqualTo(0))
                    Assert.That(wrapper.IsSelected(presentationObject), [Is].False)
                    Assert.That(inner.SelectCalls, [Is].EqualTo(0))
                End Sub)
        End Sub

        <Test>
        Public Sub CountMatchesResolvedEnumeratorForGroupLikeSelectionWithRepeatedBridgeLeaves()
            Dim layer As Object = CreateLayer()
            Dim inner As New RecordingSelectionSet(DirectCast(layer, CadViewLayer))
            Dim foreignGroupSelectionItem As New Object()
            Dim sharedBridgeContainer As Object = CreatePresentationObject("bridge-from-foreign-group")
            inner.Items.Add(foreignGroupSelectionItem)
            inner.Items.Add(sharedBridgeContainer)
            inner.Items.Add(sharedBridgeContainer)
            Dim wrapper As SelectionSet = CreateWrapper(layer, inner)

            Dim enumerated As List(Of Object) = wrapper.Cast(Of Object)().ToList()

            Assert.Multiple(
                Sub()
                    Assert.That(enumerated, Has.Count.EqualTo(2))
                    Assert.That(enumerated, Does.Contain(foreignGroupSelectionItem))
                    Assert.That(enumerated, Does.Contain(sharedBridgeContainer))
                    Assert.That(wrapper.Count, [Is].EqualTo(enumerated.Count),
                                "Count must report the resolved, reference-deduplicated selection exposed by GetEnumerator.")
                End Sub)
        End Sub

        Private Shared Function CreateLayer() As Object
            Dim drawingLayerContract As Type = GetType(Topomatic.Dwg.Layer.DrawingLayer)
            Assert.That(drawingLayerContract, [Is].Not.Null)
            Dim layerType As Type = MainType("CivilEnginStructures.BridgeDrawingLayer")
            Dim constructor As ConstructorInfo = layerType.GetConstructor(
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance,
                Nothing,
                New Type() {GetType(String)},
                Nothing)
            Assert.That(constructor, [Is].Not.Null)
            Return constructor.Invoke(New Object() {"Offline bridge selection tests"})
        End Function

        Private Shared Function CreateWrapper(layer As Object,
                                              inner As RecordingSelectionSet) As SelectionSet
            Dim wrapperType As Type = MainType("CivilEnginStructures.BridgeSelectionSet")
            Dim constructor As ConstructorInfo = wrapperType.GetConstructor(
                BindingFlags.Public Or BindingFlags.NonPublic Or BindingFlags.Instance,
                Nothing,
                New Type() {layer.GetType(), GetType(SelectionSet)},
                Nothing)
            Assert.That(constructor, [Is].Not.Null)
            Return DirectCast(constructor.Invoke(New Object() {layer, inner}), SelectionSet)
        End Function

        Private Shared Function CreatePresentationObject(idStructure As String) As Object
            Dim stateType As Type = MainType("CivilEnginStructures.BridgeContainerState")
            Dim state As Object = Activator.CreateInstance(stateType)
            Dim resolve As MethodInfo = stateType.GetMethod(
                "ResolveSelectionObject",
                BindingFlags.Public Or BindingFlags.Instance,
                Nothing,
                New Type() {GetType(Object), GetType(String)},
                Nothing)
            Assert.That(resolve, [Is].Not.Null)
            Return resolve.Invoke(state, New Object() {New Object(), idStructure})
        End Function

        Private Shared Function ClipboardMethod(wrapper As SelectionSet,
                                                methodName As String,
                                                parameterTypes As Type()) As MethodInfo
            Dim method As MethodInfo = wrapper.GetType().GetMethod(
                methodName,
                BindingFlags.NonPublic Or BindingFlags.Instance,
                Nothing,
                parameterTypes,
                Nothing)
            Assert.That(method, [Is].Not.Null, methodName & " must remain a protected SelectionSet override.")
            Return method
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
                        "Build CivilEnginStructures.vbproj before running bridge selection tests.")
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

        Private NotInheritable Class ClipboardSentinel
            Implements IStgSerializable

            Public Sub SaveToStg(node As StgNode) Implements IStgSerializable.SaveToStg
            End Sub

            Public Sub LoadFromStg(node As StgNode) Implements IStgSerializable.LoadFromStg
            End Sub
        End Class

        Private NotInheritable Class RecordingSelectionSet
            Inherits SelectionSet

            Public Sub New(layer As CadViewLayer)
                MyBase.New(layer)
            End Sub

            Public ReadOnly Items As New List(Of Object)()
            Public Property TransformData As Object
            Public Property ClipboardData As IStgSerializable
            Public Property GetTransformDataCalls As Integer
            Public Property PaintTransformDataCalls As Integer
            Public Property PaintedTransformData As Object
            Public Property TransformCalls As Integer
            Public Property TransformedData As Object
            Public Property TransformMatrix As Matrix
            Public Property TransformCopy As Boolean
            Public Property GetClipboardDataCalls As Integer
            Public Property PaintClipboardDataCalls As Integer
            Public Property PaintedClipboardData As IStgSerializable
            Public Property AddClipboardDataCalls As Integer
            Public Property AddedClipboardData As IStgSerializable
            Public Property AddClipboardTransform As Matrix
            Public Property AddClipboardBasePoint As Vector3D
            Public Property AddClipboardSelect As Boolean
            Public Property SelectCalls As Integer

            Public Overrides ReadOnly Property Count As Integer
                Get
                    Return Items.Count
                End Get
            End Property

            Public Overrides ReadOnly Property SupportClipboard As Boolean
                Get
                    Return True
                End Get
            End Property

            Public Overrides ReadOnly Property SupportCopyTransform As Boolean
                Get
                    Return True
                End Get
            End Property

            Public Overrides ReadOnly Property SupportTransform As Boolean
                Get
                    Return True
                End Get
            End Property

            Public Overrides Sub Clear()
                Items.Clear()
            End Sub

            Public Overrides Sub [Erase]()
                Items.Clear()
            End Sub

            Public Overrides Function GetEnumerator() As IEnumerator
                Return Items.GetEnumerator()
            End Function

            Public Overrides Function GetObjectsAtPoint(
                point As Vector3D,
                match As Predicate(Of Object),
                waitTimeOut As Integer) As IEnumerable(Of KeyValuePair(Of Double, Object))

                Return New KeyValuePair(Of Double, Object)() {}
            End Function

            Public Overrides Sub GetObjectsByFrame(mode As FrameSelectType,
                                                   rect As RectangleD,
                                                   match As Predicate(Of Object),
                                                   action As Action(Of Object))
            End Sub

            Public Overrides Sub GetObjectsByPolygon(mode As FrameSelectType,
                                                     pointsList As List(Of Vector2D),
                                                     match As Predicate(Of Object),
                                                     action As Action(Of Object))
            End Sub

            Public Overrides Function GetSelectable() As IEnumerable
                Return Items
            End Function

            Public Overrides Function IsOwned(obj As Object) As Boolean
                Return Items.Contains(obj)
            End Function

            Public Overrides Function IsSelected(obj As Object) As Boolean
                Return Items.Contains(obj)
            End Function

            Public Overrides Function IsEnable(obj As Object) As Boolean
                Return True
            End Function

            Public Overrides Sub [Select](item As Object, bFlag As Boolean)
                SelectCalls += 1
                If bFlag Then
                    If Not Items.Contains(item) Then Items.Add(item)
                Else
                    Items.Remove(item)
                End If
            End Sub

            Public Overrides Function GetTransformData() As Object
                GetTransformDataCalls += 1
                Return TransformData
            End Function

            Public Overrides Sub PaintTransformData(data As Object, pen As CadPen)
                PaintTransformDataCalls += 1
                PaintedTransformData = data
            End Sub

            Public Overrides Sub Transform(data As Object, transformMatrix As Matrix, copy As Boolean)
                TransformCalls += 1
                TransformedData = data
                Me.TransformMatrix = transformMatrix
                TransformCopy = copy
            End Sub

            Protected Overrides Function GetClipboardData() As IStgSerializable
                GetClipboardDataCalls += 1
                Return ClipboardData
            End Function

            Protected Overrides Sub PaintClipboardData(data As IStgSerializable, pen As CadPen)
                PaintClipboardDataCalls += 1
                PaintedClipboardData = data
            End Sub

            Protected Overrides Sub AddClipboardData(data As IStgSerializable,
                                                      transform As Matrix,
                                                      basePoint As Vector3D,
                                                      [select] As Boolean)
                AddClipboardDataCalls += 1
                AddedClipboardData = data
                AddClipboardTransform = transform
                AddClipboardBasePoint = basePoint
                AddClipboardSelect = [select]
            End Sub
        End Class
    End Class
End Namespace
