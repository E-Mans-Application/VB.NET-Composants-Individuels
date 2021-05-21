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
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        Dim barCenter As Integer = BarOffsetTop + BarHeight / 2

        e.Graphics.DrawLine(New Pen(UnselectedRangeBarColor, UnselectedRangeBarThickness), New Point(HorizontalPadding, barCenter),
New Point(e.ClipRectangle.Width - HorizontalPadding, barCenter))

        If Minimum >= Maximum Then Exit Sub

        Dim usableWidth As Integer = e.ClipRectangle.Width - 2 * HorizontalPadding
        Dim selLeft As Integer = usableWidth * (RangeStart - Minimum) / (Maximum - Minimum) + HorizontalPadding
        Dim selRight As Integer = usableWidth * (RangeEnd - Minimum) / (Maximum - Minimum) + HorizontalPadding

        e.Graphics.DrawLine(New Pen(SelectedRangeBarColor, SelectedRangeBarThickness), New Point(selLeft, barCenter),
New Point(selRight, barCenter))

        Dim height As Integer = BarHeight * 1.5
        e.Graphics.FillEllipse(New SolidBrush(ThumbColor), selLeft - CInt(height / 2), barCenter - CInt(height / 2), height, height)

        e.Graphics.FillEllipse(New SolidBrush(ThumbColor), selRight - CInt(height / 2), barCenter - CInt(height / 2), height, height)

        TrackBarRenderer.DrawHorizontalTicks(e.Graphics,
           New Rectangle(HorizontalPadding, BarOffsetTop + BarHeight + 5, usableWidth, 10), (Maximum - Minimum + 1) / TickStep, VisualStyles.EdgeStyle.Raised)
    End Sub

    Private rangeStartScrolling As Boolean
    Private rangeEndScrolling As Boolean

    Private Function GetValueFromPoint(x As Integer) As Integer
        Return (x - HorizontalPadding) / (Width - 2 * HorizontalPadding) * (Maximum - Minimum) + Minimum
    End Function

    Private Sub RangeSlider_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If Maximum <= Minimum Then Exit Sub
        If Width = 2 * HorizontalPadding Then Exit Sub

        If e.Button = MouseButtons.Left Then
            If e.Y >= BarOffsetTop And e.Y <= BarOffsetTop + BarHeight Then
                Dim ptn As Integer = GetValueFromPoint(e.X)
                If ptn >= Minimum And ptn <= Maximum Then
                    If RangeEnd > Minimum And (ptn <= RangeStart Or (ptn < RangeEnd And ptn - RangeStart < RangeEnd - ptn)) Then
                        RangeStart = ptn
                        rangeStartScrolling = True
                    Else
                        RangeEnd = ptn
                        rangeEndScrolling = True
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub RangeSlider_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        Dim ptn As Integer = GetValueFromPoint(e.X)
        If rangeStartScrolling Then
            RangeStart = ptn
        ElseIf rangeEndScrolling Then
            RangeEnd = ptn
        End If
    End Sub

    Private Sub RangeSlider_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        rangeStartScrolling = False
        rangeEndScrolling = False
    End Sub
End Class
