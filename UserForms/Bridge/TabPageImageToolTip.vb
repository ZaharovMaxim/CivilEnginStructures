Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Friend NotInheritable Class TabPageImageToolTip
    Implements IDisposable

    Private Const ImagePadding As Integer = 4

    Private ReadOnly toolTip As ToolTip
    Private ReadOnly hoverTabControl As TabControl
    Private ReadOnly showOwner As Control
    Private ReadOnly imagePaths As New Dictionary(Of TabPage, String)()
    Private ReadOnly imageCache As New Dictionary(Of String, Image)(StringComparer.OrdinalIgnoreCase)
    Private hoveredTabPage As TabPage
    Private currentImageValue As Image
    Private placementBounds As Rectangle
    Private hasPlacementBounds As Boolean
    Private disposed As Boolean

    Public Sub New(toolTip As ToolTip, hoverTabControl As TabControl)
        Me.New(toolTip, hoverTabControl, hoverTabControl)
    End Sub

    Public Sub New(toolTip As ToolTip, hoverTabControl As TabControl, showOwner As Control)
        If toolTip Is Nothing Then Throw New ArgumentNullException(NameOf(toolTip))
        If hoverTabControl Is Nothing Then Throw New ArgumentNullException(NameOf(hoverTabControl))
        If showOwner Is Nothing Then Throw New ArgumentNullException(NameOf(showOwner))

        Me.toolTip = toolTip
        Me.hoverTabControl = hoverTabControl
        Me.showOwner = showOwner
        Me.toolTip.OwnerDraw = True

        AddHandler Me.hoverTabControl.MouseMove, AddressOf TabControl_MouseMove
        AddHandler Me.hoverTabControl.MouseLeave, AddressOf TabControl_MouseLeave
        AddHandler Me.toolTip.Popup, AddressOf ToolTip_Popup
        AddHandler Me.toolTip.Draw, AddressOf ToolTip_Draw
    End Sub

    Public Sub SetPlacementBounds(bounds As Rectangle)
        placementBounds = bounds
        hasPlacementBounds = True
    End Sub

    Public Sub Register(tabPage As TabPage, imagePath As String)
        If tabPage Is Nothing Then Throw New ArgumentNullException(NameOf(tabPage))

        Dim previousImagePath As String = Nothing
        If imagePaths.TryGetValue(tabPage, previousImagePath) Then
            InvalidateImage(previousImagePath)
        End If
        imagePaths(tabPage) = imagePath
        If Object.ReferenceEquals(hoveredTabPage, tabPage) Then
            HideCurrentImage()
        End If
    End Sub

    Friend ReadOnly Property CurrentImage As Image
        Get
            Return currentImageValue
        End Get
    End Property

    Private Sub TabControl_MouseMove(sender As Object, e As MouseEventArgs)
        Dim tabPage As TabPage = FindTabPageAt(e.Location)
        Dim imagePath As String = Nothing
        If tabPage Is Nothing OrElse Not imagePaths.TryGetValue(tabPage, imagePath) Then
            HideCurrentImage()
            Return
        End If

        If Object.ReferenceEquals(hoveredTabPage, tabPage) Then Return

        HideCurrentImage()
        hoveredTabPage = tabPage
        currentImageValue = GetImage(imagePath)
        If CurrentImage Is Nothing Then Return

        If showOwner.IsHandleCreated AndAlso showOwner.Visible Then
            Dim toolTipSize As Size = GetToolTipSize()
            Dim displayBounds As Rectangle = If(hasPlacementBounds,
                                                placementBounds,
                                                showOwner.ClientRectangle)
            Dim relativeLocation As Point = CalculateToolTipLocation(displayBounds.Size, toolTipSize)
            toolTip.Show(" ",
                         showOwner,
                         displayBounds.Left + relativeLocation.X,
                         displayBounds.Top + relativeLocation.Y)
        End If
    End Sub

    Private Sub TabControl_MouseLeave(sender As Object, e As EventArgs)
        HideCurrentImage()
    End Sub

    Private Function FindTabPageAt(location As Point) As TabPage
        For index As Integer = 0 To hoverTabControl.TabPages.Count - 1
            If hoverTabControl.GetTabRect(index).Contains(location) Then
                Return hoverTabControl.TabPages(index)
            End If
        Next

        Return Nothing
    End Function

    Private Function GetImage(imagePath As String) As Image
        If String.IsNullOrWhiteSpace(imagePath) Then Return Nothing

        Dim image As Image = Nothing
        If imageCache.TryGetValue(imagePath, image) Then Return image

        Try
            Using sourceImage As Image = Image.FromFile(imagePath)
                image = New Bitmap(sourceImage)
            End Using
        Catch ex As UnauthorizedAccessException
        Catch ex As IOException
        Catch ex As ArgumentException
        Catch ex As OutOfMemoryException
        End Try

        If image IsNot Nothing Then
            imageCache(imagePath) = image
        End If
        Return image
    End Function

    Private Sub InvalidateImage(imagePath As String)
        If String.IsNullOrWhiteSpace(imagePath) Then Return

        Dim image As Image = Nothing
        If Not imageCache.TryGetValue(imagePath, image) Then Return

        If Object.ReferenceEquals(currentImageValue, image) Then
            HideCurrentImage()
        End If
        imageCache.Remove(imagePath)
        image.Dispose()
    End Sub

    Private Sub ToolTip_Popup(sender As Object, e As PopupEventArgs)
        If CurrentImage Is Nothing Then Return

        e.ToolTipSize = GetToolTipSize()
    End Sub

    Private Sub ToolTip_Draw(sender As Object, e As DrawToolTipEventArgs)
        e.DrawBackground()
        e.DrawBorder()
        If CurrentImage IsNot Nothing Then
            Dim displaySize As Size = GetDisplayImageSize()
            e.Graphics.DrawImage(CurrentImage,
                                 New Rectangle(ImagePadding,
                                               ImagePadding,
                                               displaySize.Width,
                                               displaySize.Height))
        End If
    End Sub

    Private Function GetDisplayImageSize() As Size
        If CurrentImage Is Nothing Then Return Size.Empty

        Return New Size((CurrentImage.Width + 1) \ 2,
                        (CurrentImage.Height + 1) \ 2)
    End Function

    Private Function GetToolTipSize() As Size
        Dim displaySize As Size = GetDisplayImageSize()
        Return New Size(displaySize.Width + ImagePadding * 2,
                        displaySize.Height + ImagePadding * 2)
    End Function

    Friend Shared Function CalculateToolTipLocation(displayAreaSize As Size,
                                                     toolTipSize As Size) As Point
        Return New Point(Math.Max(0, (displayAreaSize.Width - toolTipSize.Width) \ 2),
                         Math.Max(0, (displayAreaSize.Height - toolTipSize.Height) \ 2))
    End Function

    Private Sub HideCurrentImage()
        If hoveredTabPage Is Nothing AndAlso CurrentImage Is Nothing Then Return

        Try
            toolTip.Hide(showOwner)
        Catch ex As ObjectDisposedException
        End Try
        hoveredTabPage = Nothing
        currentImageValue = Nothing
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If disposed Then Return
        disposed = True

        RemoveHandler hoverTabControl.MouseMove, AddressOf TabControl_MouseMove
        RemoveHandler hoverTabControl.MouseLeave, AddressOf TabControl_MouseLeave
        RemoveHandler toolTip.Popup, AddressOf ToolTip_Popup
        RemoveHandler toolTip.Draw, AddressOf ToolTip_Draw
        HideCurrentImage()

        Dim disposedImages As New HashSet(Of Image)()
        For Each image As Image In imageCache.Values
            If image IsNot Nothing AndAlso disposedImages.Add(image) Then
                image.Dispose()
            End If
        Next
        imageCache.Clear()
        imagePaths.Clear()
    End Sub
End Class
