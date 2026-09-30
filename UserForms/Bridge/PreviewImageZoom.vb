Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Friend NotInheritable Class PreviewImageZoom
    Implements IMessageFilter, IDisposable

    Private Const WmMouseWheel As Integer = &H20A
    Private Const MinimumZoom As Double = 0.25R
    Private Const MaximumZoom As Double = 8.0R

    Private NotInheritable Class ViewState
        Friend Zoom As Double = 1.0R
        Friend Offset As PointF = PointF.Empty
        Friend LastImage As Image
    End Class

    Private ReadOnly owner As Form
    Private ReadOnly views As New Dictionary(Of PictureBox, ViewState)()
    Private filterInstalled As Boolean
    Private disposed As Boolean

    Friend Sub New(owner As Form, ParamArray pictures As PictureBox())
        If owner Is Nothing Then Throw New ArgumentNullException(NameOf(owner))
        Me.owner = owner
        For Each picture As PictureBox In pictures
            If picture Is Nothing OrElse views.ContainsKey(picture) Then Continue For
            views.Add(picture, New ViewState() With {.LastImage = picture.Image})
            AddHandler picture.Paint, AddressOf PaintPreview
            AddHandler picture.SizeChanged, AddressOf PreviewSizeChanged
        Next
        AddHandler owner.Shown, AddressOf OwnerShown
        AddHandler owner.FormClosed, AddressOf OwnerClosed
        AddHandler owner.Disposed, AddressOf OwnerDisposed
        If owner.Visible Then InstallFilter()
    End Sub

    Public Function PreFilterMessage(ByRef message As Message) As Boolean Implements IMessageFilter.PreFilterMessage
        If disposed OrElse message.Msg <> WmMouseWheel OrElse
           owner.IsDisposed OrElse Not owner.Visible OrElse Not owner.Enabled Then Return False
        If Form.ActiveForm IsNot owner AndAlso Not owner.ContainsFocus Then Return False

        Dim source As Control = Control.FromChildHandle(message.HWnd)
        If source Is Nothing OrElse (source IsNot owner AndAlso source.FindForm() IsNot owner) Then Return False

        Dim delta As Integer = SignedWord(message.WParam.ToInt64(), 16)
        If delta = 0 Then Return False
        Dim screenPoint As New Point(SignedWord(message.LParam.ToInt64(), 0),
                                     SignedWord(message.LParam.ToInt64(), 16))

        For Each pair As KeyValuePair(Of PictureBox, ViewState) In views
            Dim picture As PictureBox = pair.Key
            If picture.IsDisposed OrElse Not picture.Visible OrElse Not picture.Enabled OrElse
               picture.Image Is Nothing Then Continue For
            Dim clientPoint As Point = picture.PointToClient(screenPoint)
            If Not picture.ClientRectangle.Contains(clientPoint) Then Continue For
            ZoomAt(picture, pair.Value, clientPoint, delta)
            Return True
        Next
        Return False
    End Function

    Private Sub ZoomAt(picture As PictureBox, state As ViewState, anchor As Point, delta As Integer)
        ResetForNewImage(picture, state)
        Dim oldZoom As Double = state.Zoom
        Dim requestedZoom As Double = oldZoom * Math.Pow(1.2R, delta / 120.0R)
        Dim newZoom As Double = Math.Max(MinimumZoom, Math.Min(MaximumZoom, requestedZoom))
        If newZoom <> oldZoom Then
            Dim factor As Double = newZoom / oldZoom
            Dim centerX As Double = picture.ClientSize.Width / 2.0R
            Dim centerY As Double = picture.ClientSize.Height / 2.0R
            state.Offset = New PointF(
                CSng(anchor.X - centerX - (anchor.X - centerX - state.Offset.X) * factor),
                CSng(anchor.Y - centerY - (anchor.Y - centerY - state.Offset.Y) * factor))
            state.Zoom = newZoom
            picture.Invalidate()
        End If
    End Sub

    Private Sub PaintPreview(sender As Object, e As PaintEventArgs)
        Dim picture As PictureBox = DirectCast(sender, PictureBox)
        Dim state As ViewState = views(picture)
        ResetForNewImage(picture, state)
        Dim graphicsState As GraphicsState = e.Graphics.Save()
        Try
            e.Graphics.Clear(picture.BackColor)
            If picture.Image Is Nothing Then Return
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
            e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality
            e.Graphics.DrawImage(picture.Image, ImageBounds(picture, picture.Image, state))
        Finally
            e.Graphics.Restore(graphicsState)
        End Try
    End Sub

    Private Shared Function ImageBounds(picture As PictureBox, image As Image, state As ViewState) As RectangleF
        If picture.ClientSize.Width <= 0 OrElse picture.ClientSize.Height <= 0 OrElse
           image.Width <= 0 OrElse image.Height <= 0 Then Return RectangleF.Empty
        Dim fitScale As Double = Math.Min(picture.ClientSize.Width / CDbl(image.Width),
                                          picture.ClientSize.Height / CDbl(image.Height))
        Dim width As Double = image.Width * fitScale * state.Zoom
        Dim height As Double = image.Height * fitScale * state.Zoom
        Return New RectangleF(CSng((picture.ClientSize.Width - width) / 2.0R + state.Offset.X),
                              CSng((picture.ClientSize.Height - height) / 2.0R + state.Offset.Y),
                              CSng(width), CSng(height))
    End Function

    Private Shared Sub ResetForNewImage(picture As PictureBox, state As ViewState)
        If Object.ReferenceEquals(state.LastImage, picture.Image) Then Return
        state.LastImage = picture.Image
        state.Zoom = 1.0R
        state.Offset = PointF.Empty
    End Sub

    Private Shared Function SignedWord(value As Long, shift As Integer) As Integer
        Dim word As Integer = CInt((value >> shift) And &HFFFFL)
        Return If(word >= &H8000, word - &H10000, word)
    End Function

    Private Sub PreviewSizeChanged(sender As Object, e As EventArgs)
        DirectCast(sender, PictureBox).Invalidate()
    End Sub

    Private Sub OwnerShown(sender As Object, e As EventArgs)
        InstallFilter()
    End Sub

    Private Sub InstallFilter()
        If disposed OrElse filterInstalled Then Return
        Application.AddMessageFilter(Me)
        filterInstalled = True
    End Sub

    Private Sub OwnerClosed(sender As Object, e As FormClosedEventArgs)
        Dispose()
    End Sub

    Private Sub OwnerDisposed(sender As Object, e As EventArgs)
        Dispose()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If disposed Then Return
        disposed = True
        If filterInstalled Then
            Application.RemoveMessageFilter(Me)
            filterInstalled = False
        End If
        RemoveHandler owner.Shown, AddressOf OwnerShown
        RemoveHandler owner.FormClosed, AddressOf OwnerClosed
        RemoveHandler owner.Disposed, AddressOf OwnerDisposed
        For Each picture As PictureBox In views.Keys
            RemoveHandler picture.Paint, AddressOf PaintPreview
            RemoveHandler picture.SizeChanged, AddressOf PreviewSizeChanged
        Next
        views.Clear()
    End Sub
End Class
