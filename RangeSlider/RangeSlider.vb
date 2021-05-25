Imports System.ComponentModel

Public Class RangeSlider

    Private Const MIN_THICKNESS As Single = 0.1

    Public Event MaximumChanged(sender As Object, e As EventArgs)
    Public Event MinimumChanged(sender As Object, e As EventArgs)
    Public Event SelectedRangeChanged(sender As Object, e As EventArgs)
    Public Event PropertyChanged(sender As Object, e As EventArgs)

    Private _maximum As Long = 10

    <DefaultValue(10)>
    Public Property Maximum As Long
        Get
            Return _maximum
        End Get
        Set(value As Long)
            _maximum = Math.Max(Minimum, value)
            RangeStart = Math.Min(RangeStart, Maximum)
            RangeEnd = Math.Min(RangeEnd, Maximum)
            Invalidate()

            RaiseEvent MaximumChanged(Me, New EventArgs)
        End Set
    End Property

    Private _minimum As Long = 0

    <DefaultValue(0)>
    Public Property Minimum As Long
        Get
            Return _minimum
        End Get
        Set(value As Long)
            _minimum = Math.Min(Maximum, value)
            RangeEnd = Math.Max(RangeEnd, Minimum)
            RangeStart = Math.Max(RangeStart, Minimum)
            Invalidate()

            RaiseEvent MinimumChanged(Me, New EventArgs)
        End Set
    End Property

    Public Sub SetLimits(minimum As Long, maximum As Long)
        If minimum > maximum Then Throw New ArgumentException

        If minimum < Me.Maximum Then
            Me.Minimum = minimum
            Me.Maximum = maximum
        Else
            Me.Maximum = maximum
            Me.Minimum = minimum
        End If
    End Sub

    Private _rangeStart As Long = 0

    <DefaultValue(0)>
    Public Property RangeStart As Long
        Get
            Return _rangeStart
        End Get
        Set(value As Long)
            _rangeStart = Math.Max(Math.Min(value, RangeEnd - MinRangeLength), Minimum)
            Invalidate()

            RaiseEvent SelectedRangeChanged(Me, New EventArgs)
        End Set
    End Property

    Private _rangeEnd As Long = 10

    <DefaultValue(10)>
    Public Property RangeEnd As Long
        Get
            Return _rangeEnd
        End Get
        Set(value As Long)
            _rangeEnd = Math.Min(Math.Max(value, RangeStart + MinRangeLength), Maximum)
            Invalidate()

            RaiseEvent SelectedRangeChanged(Me, New EventArgs)
        End Set
    End Property

    Private _SelectedRangeBarThickness As Single = 1

    Public Sub SetRange(rangeStart As Long, rangeEnd As Long)
        rangeStart = Math.Max(rangeStart, Minimum)
        rangeEnd = Math.Max(rangeEnd, rangeStart + MinRangeLength)
        rangeEnd = Math.Min(rangeEnd, Maximum)
        If rangeEnd < rangeStart + MinRangeLength Then
            rangeStart = rangeEnd - MinRangeLength
        End If
        If rangeStart < Minimum Then
            Exit Sub
        End If

        _rangeStart = rangeStart
        _rangeEnd = rangeEnd

        Invalidate()

        RaiseEvent SelectedRangeChanged(Me, New EventArgs)
    End Sub

    <DefaultValue(1)>
    Public Property SelectedRangeBarThickness As Single
        Get
            Return _SelectedRangeBarThickness
        End Get
        Set(value As Single)
            _SelectedRangeBarThickness = Math.Max(MIN_THICKNESS, value)

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private _UnselectedRangeBarThickness As Single = 1

    <DefaultValue(1)>
    Public Property UnselectedRangeBarThickness As Single
        Get
            Return _UnselectedRangeBarThickness
        End Get
        Set(value As Single)
            _UnselectedRangeBarThickness = Math.Max(MIN_THICKNESS, value)

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private ReadOnly Property BarHeight As Single
        Get
            Return Math.Max(SelectedRangeBarThickness, UnselectedRangeBarThickness)
        End Get
    End Property

    Private _SelectedRangeBarColor As Color = Color.DimGray

    Public Property SelectedRangeBarColor As Color
        Get
            Return _SelectedRangeBarColor
        End Get
        Set(value As Color)
            _SelectedRangeBarColor = value

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private _UnselectedRangeBarColor As Color

    Public Property UnselectedRangeBarColor As Color
        Get
            Return _UnselectedRangeBarColor
        End Get
        Set(value As Color)
            _UnselectedRangeBarColor = value

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private _BarOffsetTop As Integer = 10

    <DefaultValue(10)>
    Public Property BarOffsetTop As Integer
        Get
            Return _BarOffsetTop
        End Get
        Set(value As Integer)
            _BarOffsetTop = Math.Max(0, value)

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private _HorizontalPadding As Integer = 10

    <DefaultValue(10)>
    Public Property HorizontalPadding As Integer
        Get
            Return _HorizontalPadding
        End Get
        Set(value As Integer)
            _HorizontalPadding = Math.Max(0, value)

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private _thumbColor As Color = Color.Black

    Public Property ThumbColor As Color
        Get
            Return _thumbColor
        End Get
        Set(value As Color)
            _thumbColor = value

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private _minRangeLength As Long = 0

    <DefaultValue(0)>
    Public Property MinRangeLength As Long
        Get
            Return _minRangeLength
        End Get
        Set(value As Long)
            _minRangeLength = Math.Max(0, value)

            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Private _tickStep As Long = 1
    <DefaultValue(1)>
    Public Property TickStep As Long
        Get
            Return _tickStep
        End Get
        Set(value As Long)
            _tickStep = Math.Max(1, value)

            Invalidate()
            RaiseEvent PropertyChanged(Me, New EventArgs)
        End Set
    End Property

    Public Property FirstItemLegend As String
    Public Property LastItemLegend As String

    Private Property OtherLegends As IEnumerator(Of KeyValuePair(Of Long, String))

    Public Sub SetOtherLegends(legends As List(Of KeyValuePair(Of Long, String)))
        If legends Is Nothing Then
            OtherLegends = Nothing
            Exit Sub
        End If
        legends = New List(Of KeyValuePair(Of Long, String))(legends)
        legends.Sort(Function(x, y) x.Key.CompareTo(y.Key))
        OtherLegends = legends.GetEnumerator
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        Dim barCenter As Integer = CInt(BarOffsetTop + BarHeight / 2)

        e.Graphics.DrawLine(New Pen(UnselectedRangeBarColor, UnselectedRangeBarThickness), New Point(HorizontalPadding, barCenter),
New Point(e.ClipRectangle.Width - HorizontalPadding, barCenter))

        If Minimum >= Maximum Then Exit Sub

        Dim usableWidth As Integer = e.ClipRectangle.Width - 2 * HorizontalPadding
        Dim selLeft As Integer = CInt(usableWidth * (RangeStart - Minimum) / (Maximum - Minimum) + HorizontalPadding)
        Dim selRight As Integer = CInt(usableWidth * (RangeEnd - Minimum) / (Maximum - Minimum) + HorizontalPadding)

        e.Graphics.DrawLine(New Pen(SelectedRangeBarColor, SelectedRangeBarThickness), New Point(selLeft, barCenter),
New Point(selRight, barCenter))

        Dim height As Integer = CInt(BarHeight * 1.5)
        e.Graphics.FillEllipse(New SolidBrush(ThumbColor), selLeft - CInt(height / 2), barCenter - CInt(height / 2), height, height)

        e.Graphics.FillEllipse(New SolidBrush(ThumbColor), selRight - CInt(height / 2), barCenter - CInt(height / 2), height, height)

        TrackBarRenderer.DrawHorizontalTicks(e.Graphics,
           New Rectangle(HorizontalPadding, CInt(BarOffsetTop + BarHeight + 5), usableWidth, 10), CInt((Maximum - Minimum) / TickStep + 1), VisualStyles.EdgeStyle.Raised)

        Dim LegendsTop As Integer = CInt(BarOffsetTop + BarHeight + 20)

        Dim lastLegendPos As Integer = Me.Width + 1
        Dim lastPos As Integer = -1
        If Not String.IsNullOrEmpty(FirstItemLegend) Then
            lastPos = DrawLegend(e.Graphics, Minimum, FirstItemLegend, LegendsTop, usableWidth, lastPos, lastLegendPos)(0)
        End If

        If Not String.IsNullOrEmpty(LastItemLegend) Then
            Dim result As Integer() = DrawLegend(e.Graphics, Maximum, LastItemLegend, LegendsTop, usableWidth, lastPos, lastLegendPos)
            If result(1) > 0 Then
                lastLegendPos = result(0) - result(1)
            End If
        End If

        If OtherLegends Is Nothing Then Exit Sub

        OtherLegends.Reset()
        While OtherLegends.MoveNext
            lastPos = DrawLegend(e.Graphics, OtherLegends.Current.Key, OtherLegends.Current.Value, LegendsTop, usableWidth, lastPos, lastLegendPos)(0)
        End While
    End Sub

    Public Function DrawLegend(graphics As Graphics, index As Long, value As String, legendsTop As Integer, usableWidth As Integer, lastPos As Integer, lastLegendPos As Integer) As Integer()

        If index < Minimum Or index > Maximum Then Return {lastPos, 0}

        Dim textWidth As Integer = CInt(graphics.MeasureString(value, Font).Width)
        Dim currentPos As Integer = CInt(usableWidth * ((index - Minimum) / (Maximum - Minimum)) + HorizontalPadding - textWidth / 2)
        If currentPos < 0 Then currentPos = 0
        If currentPos > Me.Width - textWidth Then currentPos = Me.Width - textWidth

        If currentPos > lastPos And currentPos + textWidth < lastLegendPos Then
            graphics.DrawString(value, Font, New SolidBrush(ForeColor), currentPos, legendsTop)
            Return {currentPos + textWidth, textWidth}
        End If
        Return {lastPos, 0}
    End Function

    Private rangeStartScrolling As Boolean
    Private rangeEndScrolling As Boolean
    Private translating As Boolean
    Private suppressClick As Boolean
    Private translateStart As Long

    Private Function GetValueFromPoint(x As Integer) As Long
        Return CLng((x - HorizontalPadding) / (Width - 2 * HorizontalPadding) * (Maximum - Minimum) + Minimum)
    End Function

    Private Sub RangeSlider_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If Maximum <= Minimum Then Exit Sub
        If Width = 2 * HorizontalPadding Then Exit Sub

        If e.Button = MouseButtons.Left Then
            If e.Y >= BarOffsetTop And e.Y <= BarOffsetTop + BarHeight Then
                Dim ptn As Long = GetValueFromPoint(e.X)
                If ptn >= Minimum And ptn <= Maximum Then
                    If ptn > RangeStart + 20 And ptn < RangeEnd - 20 Then
                        translateStart = ptn
                        translating = True
                    End If
                    If RangeEnd > Minimum And (ptn <= RangeStart Or (ptn < RangeEnd And ptn - RangeStart < RangeEnd - ptn)) Then
                        rangeStartScrolling = True
                    Else
                        rangeEndScrolling = True
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub RangeSlider_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        Dim ptn As Long = GetValueFromPoint(e.X)
        Dim actualPos As Long = ptn
        If My.Computer.Keyboard.CtrlKeyDown Then actualPos -= (actualPos Mod 15)
        If translating Then
            Dim diff As Long = ptn - translateStart
            If My.Computer.Keyboard.CtrlKeyDown Then diff -= ((RangeStart + diff) Mod 15)
            If diff > 0 Then
                diff = Math.Min(diff, Maximum - RangeEnd)
            Else
                diff = Math.Max(diff, Minimum - RangeStart)
            End If

            SetRange(RangeStart + diff, RangeEnd + diff)
            translateStart += diff

            suppressClick = True
        ElseIf rangeStartScrolling Then
            RangeStart = actualPos
        ElseIf rangeEndScrolling Then
            RangeEnd = actualPos
        End If
    End Sub

    Private Sub RangeSlider_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        rangeStartScrolling = False
        rangeEndScrolling = False
        translating = False
    End Sub

    Private Sub RangeSlider_Click(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        If suppressClick Then
            suppressClick = False
            Exit Sub
        End If
        Dim ptn As Long = GetValueFromPoint(e.X)
        If My.Computer.Keyboard.CtrlKeyDown Then ptn -= (ptn Mod 15)
        If rangeStartScrolling Then
            RangeStart = ptn
        ElseIf rangeEndScrolling Then
            RangeEnd = ptn
        End If
    End Sub

    Private Sub RangeSlider_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Invalidate()
    End Sub
End Class
