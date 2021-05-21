''' <summary>
''' GroupBox pliable
''' </summary>
Public Class CollapsableGroupBox

    ''' <summary>
    ''' Evénement déclenché lorsque la propriété Expanded du contrôle est modifiée
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Public Event ExpandedChanged(ByVal sender As Object, ByVal e As ExpandedEventArgs)


    ''' <summary>
    ''' Obtient ou définit une valeur indiqué si le contrôle est déplié.
    ''' </summary>
    ''' <returns></returns>
    Public Property Expanded As Boolean
        Get
            Return CheckBox1.Checked
        End Get
        Set(value As Boolean)
            CheckBox1.Checked = value
        End Set
    End Property

    ''' <summary>
    ''' Obtient ou définit le titre du GroupBox
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Property Text As String
        Get
            Return GroupBox1.Text
        End Get
        Set(value As String)
            GroupBox1.Text = value
        End Set
    End Property

    ''' <summary>
    ''' Substitut à la propriété Text
    ''' </summary>
    ''' <returns></returns>
    Public Property Title As String
        Get
            Return Text
        End Get
        Set(value As String)
            Text = value
        End Set
    End Property

    Private _expandedHeight As Integer

    ''' <summary>
    ''' La hauteur du contrôle déplié
    ''' </summary>
    ''' <returns></returns>
    Public Property ExpandedHeight As Integer
        Get
            Return _expandedHeight
        End Get
        Set(value As Integer)
            _expandedHeight = value
            UpdateGroupBox()
        End Set
    End Property

    ''' <summary>
    ''' La hauteur du contrôle plié. Calculée automatiquement.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CollapsedHeight As Integer
        Get
            Return CInt(GroupBox1.CreateGraphics().MeasureString(GroupBox1.Text, GroupBox1.Font, Width).Height) + 10
        End Get
    End Property

    Private Sub CollapsableGroupBox_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If Expanded Then
            If WrapContent Then
                Height = VerticalLinearLayout1.Height
            End If
            ExpandedHeight = Height
        Else
            Height = CollapsedHeight
        End If
        GroupBox1.Height = Height
        GroupBox1.Width = Width

        CheckBox1.Left = Width - CheckBox1.Width
        CheckBox1.Top = 0

    End Sub

    Private Sub CollapsableGroupBox_FontChanged(sender As Object, e As EventArgs) Handles Me.FontChanged
        GroupBox1.Font = Font
    End Sub

    Private Sub CollapsableGroupBox_ControlAdded(sender As Object, e As ControlEventArgs) Handles Me.ControlAdded
        ' If Not e.Control Is GroupBox1 And Not e.Control Is Button1 Then
        If Expanded Then
            e.Control.BringToFront()
        Else
            GroupBox1.BringToFront()
        End If

        '      e.Control.Font = DefaultFont
        '      GroupBox1.Controls.Add(e.Control)
        '  End If
    End Sub

    '  Public Sub SetToolTip(ByVal toolTip As ToolTip, ByVal message As String)
    '      toolTip.SetToolTip(Me, message)
    '      toolTip.SetToolTip(GroupBox1, message)
    '  End Sub

    Private Sub UpdateGroupBox()
        If CheckBox1.Checked Then
            CheckBox1.Text = "-"
            Height = ExpandedHeight
            GroupBox1.SendToBack()

        Else
            CheckBox1.Text = "+"
            Height = CollapsedHeight
            GroupBox1.BringToFront()

            GroupBox1.BringToFront()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        UpdateGroupBox()
        RaiseEvent ExpandedChanged(Me, New ExpandedEventArgs(CheckBox1.Checked))
    End Sub

    ''' <summary>
    ''' Ajoute une ligne contenant les contrôles indiqués. Ces contrôles seront positionnés sur la même ligne. La taille de la ligne est déterminée par le contrôle le plus haut.
    ''' </summary>
    ''' <param name="controls"></param>
    Public Sub AddLayer(ParamArray controls As Control())
        AddLayer(VisualStyles.VerticalAlignment.Top, controls)
    End Sub

    ''' <summary>
    ''' Ajoute une ligne contenant les contrôles indiqués. Ces contrôles seront positionnés sur la même ligne selon l'alignement spécifié. La taille de la ligne est déterminée par le contrôle le plus haut.
    ''' </summary>
    ''' <param name="controls"></param>
    Public Sub AddLayer(ByVal align As VisualStyles.VerticalAlignment, ParamArray controls As Control())
        VerticalLinearLayout1.AddLayer(align, controls)
    End Sub

    Private Sub CollapsableGroupBox_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        UpdateGroupBox()
    End Sub

    ''' <summary>
    ''' Obtient ou définit une valeur indiquant si la hauteur de la fenêtre doit être automatiquement modifiée pour correspondre exactement à la taille cumulée de tous ses contrôles.
    ''' </summary>
    ''' <returns></returns>
    Public Property WrapContent As Boolean
        Get
            Return VerticalLinearLayout1.WrapContent
        End Get
        Set(value As Boolean)
            VerticalLinearLayout1.WrapContent = value
        End Set
    End Property
End Class

Public Class ExpandedEventArgs
    Inherits EventArgs


    ''' <summary>
    ''' Obtient une valeur indiquant si le contrôle vient d'être plié ou déplié.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Expanded As Boolean

    Public Sub New(ByVal expanded As Boolean)
        _Expanded = expanded
    End Sub

End Class
