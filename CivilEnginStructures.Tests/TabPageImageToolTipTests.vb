Imports System
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Windows.Forms
Imports NUnit.Framework

Namespace Tests
    <TestFixture, Apartment(ApartmentState.STA)>
    Public Class TabPageImageToolTipTests
        <TestCase(815, 266, 416, 141)>
        <TestCase(814, 267, 415, 142)>
        <TestCase(1, 1, 9, 9)>
        Public Sub PopupUsesHalfImageSizeRoundedUpAndFourPixelPadding(imageWidth As Integer,
                                                                      imageHeight As Integer,
                                                                      expectedToolTipWidth As Integer,
                                                                      expectedToolTipHeight As Integer)
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "popup.png")
            SaveImage(imagePath, imageWidth, imageHeight)

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))

                    Dim actualSize As Size = InvokePopup(imageToolTip, tabs, New Size(1, 1))

                    Assert.That(actualSize,
                                [Is].EqualTo(New Size(expectedToolTipWidth, expectedToolTipHeight)))
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub DrawScalesImageToHalfSizeInsideFourPixelPadding()
            Const imageWidth As Integer = 815
            Const imageHeight As Integer = 266
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "quadrants.png")
            SaveQuadrantImage(imagePath, imageWidth, imageHeight)

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs),
                      renderedToolTip As New Bitmap(416, 141)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))

                    InvokeDraw(imageToolTip, tabs, renderedToolTip)

                    Assert.That(renderedToolTip.GetPixel(30, 30).ToArgb(), [Is].EqualTo(Color.Red.ToArgb()),
                                "Top-left quadrant must be drawn inside the scaled destination rectangle.")
                    Assert.That(renderedToolTip.GetPixel(310, 30).ToArgb(), [Is].EqualTo(Color.Blue.ToArgb()),
                                "Top-right source quadrant must be visible after scaling to 408 pixels.")
                    Assert.That(renderedToolTip.GetPixel(30, 100).ToArgb(), [Is].EqualTo(Color.Green.ToArgb()),
                                "Bottom-left source quadrant must be visible after scaling to 133 pixels.")
                    Assert.That(renderedToolTip.GetPixel(310, 100).ToArgb(), [Is].EqualTo(Color.Yellow.ToArgb()),
                                "Bottom-right source quadrant must be visible after scaling.")
                    Assert.That(renderedToolTip.GetPixel(411, 136).ToArgb(), [Is].EqualTo(Color.Yellow.ToArgb()),
                                "The 408x133 destination rectangle must include its bottom-right pixel.")
                    Assert.That(renderedToolTip.GetPixel(412, 136).ToArgb(), [Is].EqualTo(Color.Black.ToArgb()),
                                "The image must stop before the right four-pixel padding.")
                    Assert.That(renderedToolTip.GetPixel(411, 137).ToArgb(), [Is].EqualTo(Color.Black.ToArgb()),
                                "The image must stop before the bottom four-pixel padding.")
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub ConstructorAcceptsControlUsedAsToolTipDisplayArea()
            Dim constructor As ConstructorInfo = GetType(TabPageImageToolTip).GetConstructor(
                BindingFlags.Instance Or BindingFlags.Public Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=New Type() {GetType(ToolTip), GetType(TabControl), GetType(Control)},
                modifiers:=Nothing)

            Assert.That(constructor, [Is].Not.Null,
                        "The helper constructor must accept ToolTip, TabControl and displayArea Control.")

            Using tabs As TestTabControl = CreateTabControl("Nozzle"),
                  displayArea As New PictureBox(),
                  toolTip As New ToolTip()
                Dim imageToolTip As IDisposable = DirectCast(
                    constructor.Invoke(New Object() {toolTip, tabs, displayArea}),
                    IDisposable)
                imageToolTip.Dispose()
            End Using
        End Sub

        <TestCase(1000, 500, 416, 141, 292, 179)>
        <TestCase(416, 141, 416, 141, 0, 0)>
        <TestCase(200, 100, 416, 141, 0, 0)>
        Public Sub CalculateToolTipLocationCentersWithinDisplayAreaAndClampsAtZero(
            areaWidth As Integer,
            areaHeight As Integer,
            toolTipWidth As Integer,
            toolTipHeight As Integer,
            expectedX As Integer,
            expectedY As Integer)
            Dim calculateLocation As MethodInfo = GetType(TabPageImageToolTip).GetMethod(
                "CalculateToolTipLocation",
                BindingFlags.Static Or BindingFlags.Public Or BindingFlags.NonPublic,
                binder:=Nothing,
                types:=New Type() {GetType(Size), GetType(Size)},
                modifiers:=Nothing)

            Assert.That(calculateLocation, [Is].Not.Null,
                        "A pure CalculateToolTipLocation(displayAreaSize, toolTipSize) helper is required.")

            Dim actualLocation As Point = DirectCast(
                calculateLocation.Invoke(Nothing,
                                         New Object() {New Size(areaWidth, areaHeight),
                                                       New Size(toolTipWidth, toolTipHeight)}),
                Point)

            Assert.That(actualLocation, [Is].EqualTo(New Point(expectedX, expectedY)))
        End Sub

        <Test>
        Public Sub RegisteredTabHeadersSelectTheirOwnImagesAndUnregisteredHeaderIsIgnored()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim firstImagePath As String = Path.Combine(tempDirectory, "first.png")
            Dim secondImagePath As String = Path.Combine(tempDirectory, "second.png")
            SaveImage(firstImagePath, 11, 7)
            SaveImage(secondImagePath, 13, 9)

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle", "Crossbar", "Posts"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    imageToolTip.Register(tabs.TabPages(0), firstImagePath)
                    imageToolTip.Register(tabs.TabPages(1), secondImagePath)

                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))
                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)
                    Assert.That(imageToolTip.CurrentImage.Size, [Is].EqualTo(New Size(11, 7)))

                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 1))
                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)
                    Assert.That(imageToolTip.CurrentImage.Size, [Is].EqualTo(New Size(13, 9)))

                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 2))
                    Assert.That(imageToolTip.CurrentImage, [Is].Null)
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub TabContentDoesNotSelectRegisteredTabImage()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "nozzle.png")
            SaveImage(imagePath, 17, 10)

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle", "Crossbar"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)

                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))
                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)

                    tabs.RaiseMouseMoveAt(ContentCenter(tabs))
                    Assert.That(imageToolTip.CurrentImage, [Is].Null)
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub ValidImageIsDetachedFromItsSourceFile()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "nozzle.png")
            SaveImage(imagePath, 19, 12)

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))

                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)
                    Assert.DoesNotThrow(Sub() File.Delete(imagePath))
                    Assert.That(imageToolTip.CurrentImage.Size, [Is].EqualTo(New Size(19, 12)))
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub MissingAndDamagedImagesDoNotThrowAndDoNotBecomeCurrentImage()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim missingImagePath As String = Path.Combine(tempDirectory, "missing.png")
            Dim damagedImagePath As String = Path.Combine(tempDirectory, "damaged.jpg")
            File.WriteAllBytes(damagedImagePath, New Byte() {0, 1, 2, 3, 4})

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle", "Crossbar"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    Assert.DoesNotThrow(Sub() imageToolTip.Register(tabs.TabPages(0), missingImagePath))
                    Assert.DoesNotThrow(Sub() imageToolTip.Register(tabs.TabPages(1), damagedImagePath))

                    Assert.DoesNotThrow(Sub() tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0)))
                    Assert.That(imageToolTip.CurrentImage, [Is].Null)

                    Assert.DoesNotThrow(Sub() tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 1)))
                    Assert.That(imageToolTip.CurrentImage, [Is].Null)
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub RegisterRetriesImageThatWasMissingDuringPreviousLoad()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "appears-later.png")

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))
                    Assert.That(imageToolTip.CurrentImage, [Is].Null)

                    SaveImage(imagePath, 23, 14)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))

                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)
                    Assert.That(imageToolTip.CurrentImage.Size, [Is].EqualTo(New Size(23, 14)))
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub RegisterReloadsImageWhenFileAtSamePathWasReplaced()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "replaceable.png")
            SaveImage(imagePath, 25, 15)

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))
                    Assert.That(imageToolTip.CurrentImage.Size, [Is].EqualTo(New Size(25, 15)))

                    File.Delete(imagePath)
                    SaveImage(imagePath, 31, 18)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))

                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)
                    Assert.That(imageToolTip.CurrentImage.Size, [Is].EqualTo(New Size(31, 18)))
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub MouseLeaveClearsCurrentImage()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "nozzle.png")
            SaveImage(imagePath, 21, 13)

            Try
                Using tabs As TestTabControl = CreateTabControl("Nozzle"),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    imageToolTip.Register(tabs.TabPages(0), imagePath)
                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, 0))
                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)

                    tabs.RaiseMouseLeave()

                    Assert.That(imageToolTip.CurrentImage, [Is].Null)
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        <Test>
        Public Sub RegisteredHeaderInLastMultilineRowSelectsItsImage()
            Dim tempDirectory As String = CreateTempDirectory()
            Dim imagePath As String = Path.Combine(tempDirectory, "last-row.png")
            SaveImage(imagePath, 27, 16)

            Try
                Using tabs As TestTabControl = CreateMultilineTabControl(),
                      toolTip As New ToolTip(),
                      imageToolTip As New TabPageImageToolTip(toolTip, tabs)
                    Dim lastRowTabIndex As Integer = FindLastRowTabIndex(tabs)
                    imageToolTip.Register(tabs.TabPages(lastRowTabIndex), imagePath)

                    tabs.RaiseMouseMoveAt(HeaderCenter(tabs, lastRowTabIndex))

                    Assert.That(imageToolTip.CurrentImage, [Is].Not.Null)
                    Assert.That(imageToolTip.CurrentImage.Size, [Is].EqualTo(New Size(27, 16)))
                End Using
            Finally
                Directory.Delete(tempDirectory, recursive:=True)
            End Try
        End Sub

        Private Shared Function CreateTabControl(ParamArray tabNames As String()) As TestTabControl
            Dim tabs As New TestTabControl With {
                .Size = New Size(480, 300)
            }

            For Each tabName As String In tabNames
                tabs.TabPages.Add(New TabPage(tabName))
            Next

            tabs.CreateControl()
            Return tabs
        End Function

        Private Shared Function CreateMultilineTabControl() As TestTabControl
            Dim tabs As New TestTabControl With {
                .Multiline = True,
                .Size = New Size(170, 300)
            }

            For index As Integer = 0 To 7
                tabs.TabPages.Add(New TabPage("Long tab " & index.ToString()))
            Next

            tabs.CreateControl()
            Return tabs
        End Function

        Private Shared Function FindLastRowTabIndex(tabs As TabControl) As Integer
            Dim minimumTop As Integer = Integer.MaxValue
            Dim maximumTop As Integer = Integer.MinValue
            Dim lastRowTabIndex As Integer = -1

            For index As Integer = 0 To tabs.TabPages.Count - 1
                Dim headerTop As Integer = tabs.GetTabRect(index).Top
                minimumTop = Math.Min(minimumTop, headerTop)
                If headerTop > maximumTop Then
                    maximumTop = headerTop
                    lastRowTabIndex = index
                End If
            Next

            Assert.That(maximumTop, [Is].GreaterThan(minimumTop),
                        "Test setup must create at least two tab header rows.")
            Return lastRowTabIndex
        End Function

        Private Shared Function HeaderCenter(tabs As TabControl, tabIndex As Integer) As Point
            Dim headerRectangle As Rectangle = tabs.GetTabRect(tabIndex)
            Return New Point(headerRectangle.Left + headerRectangle.Width \ 2,
                             headerRectangle.Top + headerRectangle.Height \ 2)
        End Function

        Private Shared Function ContentCenter(tabs As TabControl) As Point
            Dim contentRectangle As Rectangle = tabs.DisplayRectangle
            Return New Point(contentRectangle.Left + contentRectangle.Width \ 2,
                             contentRectangle.Top + contentRectangle.Height \ 2)
        End Function

        Private Shared Function CreateTempDirectory() As String
            Dim directoryPath As String = Path.Combine(TestContext.CurrentContext.WorkDirectory,
                                                       NameOf(TabPageImageToolTipTests),
                                                       Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(directoryPath)
            Return directoryPath
        End Function

        Private Shared Sub SaveImage(path As String, width As Integer, height As Integer)
            Using bitmap As New Bitmap(width, height)
                bitmap.Save(path, ImageFormat.Png)
            End Using
        End Sub

        Private Shared Sub SaveQuadrantImage(path As String, width As Integer, height As Integer)
            Dim rightStart As Integer = width \ 2
            Dim bottomStart As Integer = height \ 2

            Using bitmap As New Bitmap(width, height),
                  graphics As Graphics = Graphics.FromImage(bitmap)
                graphics.Clear(Color.Red)
                graphics.FillRectangle(Brushes.Blue,
                                       rightStart,
                                       0,
                                       width - rightStart,
                                       bottomStart)
                graphics.FillRectangle(Brushes.Green,
                                       0,
                                       bottomStart,
                                       rightStart,
                                       height - bottomStart)
                graphics.FillRectangle(Brushes.Yellow,
                                       rightStart,
                                       bottomStart,
                                       width - rightStart,
                                       height - bottomStart)
                bitmap.Save(path, ImageFormat.Png)
            End Using
        End Sub

        Private Shared Function InvokePopup(imageToolTip As TabPageImageToolTip,
                                            associatedControl As Control,
                                            initialSize As Size) As Size
            Dim eventArguments As New PopupEventArgs(associatedControl,
                                                     associatedControl,
                                                     False,
                                                     initialSize)
            InvokePrivateHandler(imageToolTip, "ToolTip_Popup", eventArguments)
            Return eventArguments.ToolTipSize
        End Function

        Private Shared Sub InvokeDraw(imageToolTip As TabPageImageToolTip,
                                      associatedControl As Control,
                                      target As Bitmap)
            Using graphics As Graphics = Graphics.FromImage(target)
                Dim eventArguments As New DrawToolTipEventArgs(graphics,
                                                              associatedControl,
                                                              associatedControl,
                                                              New Rectangle(Point.Empty, target.Size),
                                                              String.Empty,
                                                              Color.Black,
                                                              Color.White,
                                                              SystemFonts.DefaultFont)
                InvokePrivateHandler(imageToolTip, "ToolTip_Draw", eventArguments)
            End Using
        End Sub

        Private Shared Sub InvokePrivateHandler(imageToolTip As TabPageImageToolTip,
                                                handlerName As String,
                                                eventArguments As EventArgs)
            Dim handler As MethodInfo = GetType(TabPageImageToolTip).GetMethod(handlerName,
                                                                               BindingFlags.Instance Or
                                                                               BindingFlags.NonPublic)
            Assert.That(handler, [Is].Not.Null, handlerName & " handler must exist.")
            handler.Invoke(imageToolTip, New Object() {Nothing, eventArguments})
        End Sub

        Private NotInheritable Class TestTabControl
            Inherits TabControl

            Public Sub RaiseMouseMoveAt(location As Point)
                MyBase.OnMouseMove(New MouseEventArgs(MouseButtons.None, 0, location.X, location.Y, 0))
            End Sub

            Public Sub RaiseMouseLeave()
                MyBase.OnMouseLeave(EventArgs.Empty)
            End Sub
        End Class
    End Class
End Namespace
