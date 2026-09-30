Imports System
Imports System.Drawing
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class PreviewImageZoomTests
        Private Const WmMouseWheel As Integer = &H20A

        <Test>
        Public Sub FilterIsInstalledOnlyWhileTheOwnerIsShownAndDoesNotOwnTheImage()
            Dim image As Bitmap = CreateMarkerImage()
            Dim owner As Form = CreateOwner()
            Dim picture As PictureBox = CreatePicture("front", New Rectangle(12, 42, 200, 200), image)
            owner.Controls.Add(picture)
            Dim zoom As IDisposable = CreateZoom(owner, picture)
            Try
                Dim clientPoint As New Point(60, 80)
                Dim beforeShow As Message = WheelMessage(picture.Handle, 120, picture.PointToScreen(clientPoint))
                Assert.That(Application.FilterMessage(beforeShow), [Is].False,
                            "The process-wide filter must not be installed by a form constructor that is never shown.")

                ShowOwner(owner)
                Dim whileShown As Message = WheelMessage(picture.Handle, 120, picture.PointToScreen(clientPoint))
                Assert.That(Application.FilterMessage(whileShown), [Is].True,
                            "A shown owner must route wheel input over its preview through the installed filter.")
                Assert.That(picture.Image, [Is].SameAs(image))

                Dim staleHandle As IntPtr = picture.Handle
                Dim stalePoint As Point = picture.PointToScreen(clientPoint)
                owner.Close()
                Application.DoEvents()
                Dim afterClose As Message = WheelMessage(staleHandle, 120, stalePoint)
                Assert.That(Application.FilterMessage(afterClose), [Is].False,
                            "Closing the owner must remove its process-wide message filter.")

                Assert.DoesNotThrow(Sub() image.SetPixel(0, 0, Color.Lime),
                                    "PreviewImageZoom must not dispose images supplied by the caller.")
            Finally
                zoom.Dispose()
                owner.Dispose()
                image.Dispose()
            End Try
        End Sub

        <Test>
        Public Sub WheelZoomsAroundTheCursorWithoutChangingImageOrKeyboardFocus()
            Dim image As Bitmap = CreateMarkerImage()
            Dim owner As Form = CreateOwner()
            Dim editor As New TextBox() With {.Name = "Editor", .Bounds = New Rectangle(12, 10, 120, 24)}
            Dim picture As PictureBox = CreatePicture("front", New Rectangle(12, 42, 200, 200), image)
            owner.Controls.Add(editor)
            owner.Controls.Add(picture)
            Dim zoom As IDisposable = CreateZoom(owner, picture)
            Try
                ShowOwner(owner, editor)
                Dim before As Rectangle = MarkerBounds(Render(picture))
                Dim cursor As New Point(CInt(Math.Round(before.Left + before.Width / 2.0R)),
                                        CInt(Math.Round(before.Top + before.Height / 2.0R)))
                Dim message As Message = WheelMessage(editor.Handle, 120, picture.PointToScreen(cursor))

                Assert.That(DirectCast(zoom, IMessageFilter).PreFilterMessage(message), [Is].True)
                Application.DoEvents()
                Dim after As Rectangle = MarkerBounds(Render(picture))

                Assert.Multiple(
                    Sub()
                        Assert.That(after.Width, [Is].GreaterThanOrEqualTo(before.Width + 3), "one wheel notch zoom")
                        Assert.That(Math.Abs(RectangleCenter(after).X - RectangleCenter(before).X), [Is].LessThanOrEqualTo(2.0R),
                                    "horizontal cursor anchor")
                        Assert.That(Math.Abs(RectangleCenter(after).Y - RectangleCenter(before).Y), [Is].LessThanOrEqualTo(2.0R),
                                    "vertical cursor anchor")
                        Assert.That(picture.Image, [Is].SameAs(image), "image identity")
                        Assert.That(editor.Focused, [Is].True, "wheel zoom must not steal keyboard focus")
                    End Sub)
            Finally
                zoom.Dispose()
                owner.Dispose()
                image.Dispose()
            End Try
        End Sub

        <Test>
        Public Sub FractionalWheelDeltasAndBothZoomLimitsAreObservableInRendering()
            Dim image As Bitmap = CreateMarkerImage()
            Dim owner As Form = CreateOwner()
            Dim picture As PictureBox = CreatePicture("front", New Rectangle(12, 12, 200, 200), image)
            owner.Controls.Add(picture)
            Dim zoom As IDisposable = CreateZoom(owner, picture)
            Try
                ShowOwner(owner, picture)
                Dim filter As IMessageFilter = DirectCast(zoom, IMessageFilter)
                Dim baseline As Rectangle = MarkerBounds(Render(picture))
                Dim cursor As New Point(CInt(RectangleCenter(baseline).X), CInt(RectangleCenter(baseline).Y))

                SendWheel(filter, picture.Handle, picture, cursor, 60)
                Dim halfNotch As Rectangle = MarkerBounds(Render(picture))
                SendWheel(filter, picture.Handle, picture, cursor, 60)
                Dim fullNotch As Rectangle = MarkerBounds(Render(picture))
                Assert.That(halfNotch.Width, [Is].GreaterThan(baseline.Width), "a half notch must not be rounded away")
                Assert.That(halfNotch.Width, [Is].LessThan(fullNotch.Width), "fractional delta must use an exponential zoom step")

                For index As Integer = 1 To 30
                    SendWheel(filter, picture.Handle, picture, cursor, 120)
                Next
                Dim maximum As Rectangle = MarkerBounds(Render(picture))
                SendWheel(filter, picture.Handle, picture, cursor, 120)
                Assert.That(MarkerBounds(Render(picture)), [Is].EqualTo(maximum), "zoom must clamp at 8x")

                For index As Integer = 1 To 60
                    SendWheel(filter, picture.Handle, picture, cursor, -120)
                Next
                Dim minimum As Rectangle = MarkerBounds(Render(picture))
                SendWheel(filter, picture.Handle, picture, cursor, -120)
                Assert.Multiple(
                    Sub()
                        Assert.That(MarkerBounds(Render(picture)), [Is].EqualTo(minimum), "zoom must clamp at 0.25x")
                        Assert.That(minimum.Width, [Is].LessThan(baseline.Width))
                    End Sub)
            Finally
                zoom.Dispose()
                owner.Dispose()
                image.Dispose()
            End Try
        End Sub

        <Test>
        Public Sub ViewsZoomIndependentlyAndAssigningANewImageResetsThatView()
            Dim firstImage As Bitmap = CreateMarkerImage()
            Dim secondImage As Bitmap = CreateMarkerImage()
            Dim thirdImage As Bitmap = CreateMarkerImage()
            Dim replacement As Bitmap = CreateMarkerImage()
            Dim owner As Form = CreateOwner(New Size(540, 220))
            Dim first As PictureBox = CreatePicture("front", New Rectangle(10, 10, 160, 160), firstImage)
            Dim second As PictureBox = CreatePicture("left", New Rectangle(180, 10, 160, 160), secondImage)
            Dim third As PictureBox = CreatePicture("right", New Rectangle(350, 10, 160, 160), thirdImage)
            owner.Controls.AddRange(New Control() {first, second, third})
            Dim zoom As IDisposable = CreateZoom(owner, first, second, third)
            Try
                ShowOwner(owner, first)
                Dim firstBaseline As Rectangle = MarkerBounds(Render(first))
                Dim secondBaseline As Rectangle = MarkerBounds(Render(second))
                Dim thirdBaseline As Rectangle = MarkerBounds(Render(third))
                Dim cursor As New Point(CInt(RectangleCenter(firstBaseline).X), CInt(RectangleCenter(firstBaseline).Y))
                SendWheel(DirectCast(zoom, IMessageFilter), first.Handle, first, cursor, 120)

                Assert.That(MarkerBounds(Render(first)).Width, [Is].GreaterThan(firstBaseline.Width))
                Assert.That(MarkerBounds(Render(second)), [Is].EqualTo(secondBaseline))
                Assert.That(MarkerBounds(Render(third)), [Is].EqualTo(thirdBaseline))
                Assert.That(second.Image, [Is].SameAs(secondImage))
                Assert.That(third.Image, [Is].SameAs(thirdImage))

                first.Image = replacement
                Application.DoEvents()
                Assert.Multiple(
                    Sub()
                        Assert.That(MarkerBounds(Render(first)), [Is].EqualTo(firstBaseline),
                                    "assigning a new image must reset only that view to aspect-fit")
                        Assert.That(first.Image, [Is].SameAs(replacement))
                        Assert.That(MarkerBounds(Render(second)), [Is].EqualTo(secondBaseline))
                        Assert.That(MarkerBounds(Render(third)), [Is].EqualTo(thirdBaseline))
                    End Sub)
            Finally
                zoom.Dispose()
                owner.Dispose()
                firstImage.Dispose()
                secondImage.Dispose()
                thirdImage.Dispose()
                replacement.Dispose()
            End Try
        End Sub

        <Test>
        Public Sub WheelIsIgnoredOutsideVisibleOwnedPreviewsAndForForeignWindows()
            Dim image As Bitmap = CreateMarkerImage()
            Dim owner As Form = CreateOwner()
            Dim picture As PictureBox = CreatePicture("front", New Rectangle(12, 12, 200, 200), image)
            owner.Controls.Add(picture)
            Dim zoom As IDisposable = CreateZoom(owner, picture)
            Dim foreign As Form = CreateOwner()
            Dim foreignEditor As New TextBox() With {.Bounds = New Rectangle(10, 10, 100, 24)}
            foreign.Controls.Add(foreignEditor)
            Try
                ShowOwner(owner, picture)
                Dim filter As IMessageFilter = DirectCast(zoom, IMessageFilter)
                Dim baseline As Rectangle = MarkerBounds(Render(picture))
                Dim inside As Point = picture.PointToScreen(New Point(80, 80))
                Dim outside As Point = picture.PointToScreen(New Point(picture.Width + 20, picture.Height + 20))

                Dim outsideMessage As Message = WheelMessage(picture.Handle, 120, outside)
                Assert.That(filter.PreFilterMessage(outsideMessage), [Is].False, "cursor outside previews")

                foreign.Show()
                foreign.Activate()
                foreignEditor.Focus()
                Application.DoEvents()
                Dim foreignMessage As Message = WheelMessage(foreignEditor.Handle, 120, inside)
                Assert.That(filter.PreFilterMessage(foreignMessage), [Is].False, "foreign HWND and inactive owner")

                foreign.Hide()
                owner.Activate()
                picture.Visible = False
                Application.DoEvents()
                Dim hiddenMessage As Message = WheelMessage(picture.Handle, 120, inside)
                Assert.That(filter.PreFilterMessage(hiddenMessage), [Is].False, "hidden preview")
                Assert.That(MarkerBounds(Render(picture)), [Is].EqualTo(baseline))
            Finally
                zoom.Dispose()
                foreign.Dispose()
                owner.Dispose()
                image.Dispose()
            End Try
        End Sub

        Private Shared Function CreateZoom(owner As Form, ParamArray pictures As PictureBox()) As IDisposable
            Dim helperType As Type = GetType(PreviewImageZoomTests).Assembly.GetType(
                "CivilEnginStructures.PreviewImageZoom", throwOnError:=False)
            If helperType Is Nothing Then
                Assert.Fail("PreviewImageZoom has not been implemented yet. The wheel-zoom characterization tests are expected to be red.")
            End If
            Dim constructor As ConstructorInfo = helperType.GetConstructor(
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=New Type() {GetType(Form), GetType(PictureBox())},
                modifiers:=Nothing)
            Assert.That(constructor, [Is].Not.Null,
                        "PreviewImageZoom must expose constructor(owner As Form, ParamArray pictures As PictureBox()).")
            Dim instance As Object = constructor.Invoke(New Object() {owner, pictures})
            Assert.That(instance, [Is].InstanceOf(Of IMessageFilter)())
            Assert.That(instance, [Is].InstanceOf(Of IDisposable)())
            Return DirectCast(instance, IDisposable)
        End Function

        Private Shared Function CreateOwner(Optional size As Size = Nothing) As Form
            Dim result As New Form() With {
                .ShowInTaskbar = False,
                .StartPosition = FormStartPosition.Manual,
                .Location = New Point(-30000, -30000),
                .ClientSize = If(size.IsEmpty, New Size(240, 250), size)
            }
            Return result
        End Function

        Private Shared Function CreatePicture(name As String, bounds As Rectangle, image As Image) As PictureBox
            Return New PictureBox() With {
                .Name = name,
                .Bounds = bounds,
                .BackColor = Color.FromArgb(12, 34, 56),
                .BorderStyle = BorderStyle.None,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .Image = image,
                .TabStop = False
            }
        End Function

        Private Shared Sub ShowOwner(owner As Form, Optional focusTarget As Control = Nothing)
            owner.Show()
            owner.Activate()
            If focusTarget IsNot Nothing Then focusTarget.Focus()
            Application.DoEvents()
        End Sub

        Private Shared Sub SendWheel(filter As IMessageFilter, hwnd As IntPtr, picture As PictureBox,
                                     clientPoint As Point, delta As Integer)
            Dim message As Message = WheelMessage(hwnd, delta, picture.PointToScreen(clientPoint))
            Assert.That(filter.PreFilterMessage(message), [Is].True, "wheel message should be consumed")
            Application.DoEvents()
        End Sub

        Private Shared Function WheelMessage(hwnd As IntPtr, delta As Integer, screenPoint As Point) As Message
            Dim wParamValue As Long = (CLng(delta) And &HFFFFL) << 16
            Dim lParamValue As Long = (CLng(screenPoint.X) And &HFFFFL) Or
                                      ((CLng(screenPoint.Y) And &HFFFFL) << 16)
            Return Message.Create(hwnd, WmMouseWheel, New IntPtr(wParamValue), New IntPtr(lParamValue))
        End Function

        Private Shared Function CreateMarkerImage() As Bitmap
            Dim result As New Bitmap(100, 100)
            Using graphics As Graphics = Graphics.FromImage(result)
                graphics.Clear(Color.Transparent)
                graphics.FillRectangle(Brushes.Red, 25, 35, 10, 10)
            End Using
            Return result
        End Function

        Private Shared Function Render(picture As PictureBox) As Bitmap
            Dim result As New Bitmap(picture.Width, picture.Height)
            picture.DrawToBitmap(result, picture.ClientRectangle)
            Return result
        End Function

        Private Shared Function MarkerBounds(rendered As Bitmap) As Rectangle
            Using rendered
                Dim left As Integer = rendered.Width
                Dim top As Integer = rendered.Height
                Dim right As Integer = -1
                Dim bottom As Integer = -1
                For y As Integer = 0 To rendered.Height - 1
                    For x As Integer = 0 To rendered.Width - 1
                        Dim pixel As Color = rendered.GetPixel(x, y)
                        If pixel.R >= 180 AndAlso pixel.G <= 90 AndAlso pixel.B <= 90 Then
                            left = Math.Min(left, x)
                            top = Math.Min(top, y)
                            right = Math.Max(right, x)
                            bottom = Math.Max(bottom, y)
                        End If
                    Next
                Next
                Assert.That(right, [Is].GreaterThanOrEqualTo(left), "The marker image must be visible in the rendered preview.")
                Return Rectangle.FromLTRB(left, top, right + 1, bottom + 1)
            End Using
        End Function

        Private Shared Function RectangleCenter(bounds As Rectangle) As PointF
            Return New PointF(bounds.Left + bounds.Width / 2.0F, bounds.Top + bounds.Height / 2.0F)
        End Function
    End Class
End Namespace
