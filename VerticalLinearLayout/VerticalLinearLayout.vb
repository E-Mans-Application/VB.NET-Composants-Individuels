''' <summary>
''' Contrôle conteneur avec une logique de positionnement vertical linéaire des contrôles enfants.
''' Les contrôles sont positionnés par couche, les uns en dessous des autres.
''' La position horizontale des contrôles peut être définie manuellement et n'est pas modifiée.
''' </summary>
Public Class VerticalLinearLayout

    Private Layers As List(Of Layer)

    ''' <summary>
    ''' Obtient une valeur indiquant la taille qu'aurait prise le formulaire si la propriété WrapContent était activée.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property WrapHeight As Integer

    ''' <summary>
    ''' Positionne les contrôles du parent, comme s'ils étaient dans le composant, plutôt que les contrôles du composant lui-même.
    ''' </summary>
    ''' <returns></returns>
    Public Property UseParentControls As Boolean = False

    ' Public Event LayerInitializationNeeded(ByVal sender As Object, ByVal e As EventArgs)

    ''' <summary>
    ''' Réinitialise la liste des lignes
    ''' </summary>
    Public Sub ResetLayers()
        If Layers Is Nothing Then Layers = New List(Of Layer)
        Layers.Clear()
    End Sub

    '  ''' <summary>
    ' ''' Appelé au chargement du formulaire pour définir les règles de positionnement des éléments
    '  ''' </summary>
    '  Private Sub InitializeLayers()
    '      ResetLayers()
    '     RaiseEvent LayerInitializationNeeded(Me, New EventArgs)
    '  End Sub

    ''' <summary>
    ''' Ajoute une ligne contenant les contrôles indiqués. Ces contrôles seront positionnés sur la même ligne. La taille de la ligne est déterminée par le contrôle le plus haut.
    ''' Au sein de la ligne, seule la position horizontale des contrôles peut être définie manuellement.
    ''' </summary>
    ''' <param name="controls"></param>
    Public Sub AddLayer(ParamArray controls As Control())
        AddLayer(VisualStyles.VerticalAlignment.Top, controls)
    End Sub

    ''' <summary>
    ''' Ajoute une ligne contenant les contrôles indiqués. Ces contrôles seront positionnés sur la même ligne selon l'alignement spécifié. La taille de la ligne est déterminée par le contrôle le plus haut.
    ''' Au sein de la ligne, seule la position horizontale des contrôles peut être définie manuellement.
    ''' </summary>
    ''' <param name="controls"></param>
    Public Sub AddLayer(ByVal align As VisualStyles.VerticalAlignment, ParamArray controls As Control())
        If Layers Is Nothing Then Layers = New List(Of Layer)
        Layers.Add(New Layer(align, controls))
        AddHandlers(controls)
        Reorganize(Nothing, Nothing)
    End Sub

    'Public Sub InsertLayer(ByVal i As Integer, ParamArray controls As Control())
    '    InsertLayer(i, VisualStyles.VerticalAlignment.Top, controls)
    'End Sub

    'Public Sub InsertLayer(ByVal i As Integer, ByVal align As VisualStyles.VerticalAlignment, ParamArray controls As Control())
    '    If Layers Is Nothing Then Layers = New List(Of Layer)
    '    Layers.Insert(i, New Layer(align, controls))
    '    AddHandlers(controls)
    '    Reorganize(Nothing, Nothing)
    'End Sub

    ''' <summary>
    ''' Enregistre les gestionnaires d'événements pour les contrôles, afin de mettre à jour automatiquement le positionnement des contrôles.
    ''' </summary>
    ''' <param name="controls"></param>
    Private Sub AddHandlers(ByVal controls As Control())
        For Each ctrl As Control In controls
            If Not UseParentControls And Not Me.Controls.Contains(ctrl) Then
                Me.Controls.Add(ctrl)
            End If
            AddHandlers(ctrl)
        Next
    End Sub

    ''' <summary>
    ''' Enregistre les gestionnaires d'événements pour le contrôle.
    ''' </summary>
    ''' <param name="ctrl"></param>
    Private Sub AddHandlers(ByVal ctrl As Control)
        '    UpdatePos(ctrl, New EventArgs)
        AddHandler ctrl.SizeChanged, AddressOf Reorganize
        AddHandler ctrl.LocationChanged, AddressOf Reorganize
        AddHandler ctrl.MarginChanged, AddressOf Reorganize
        AddHandler ctrl.PaddingChanged, AddressOf Reorganize
    End Sub

    ''' <summary>
    ''' Obtient ou définit une valeur indiquant si la hauteur de la fenêtre doit être automatiquement modifiée pour correspondre exactement à la taille cumulée de tous ses contrôles.
    ''' </summary>
    ''' <returns></returns>
    Public Property WrapContent As Boolean

    ''' <summary>
    ''' Force le repositionnement des contrôles enfants selon la logique de positionnement par lignes.
    ''' </summary>
    Public Sub Reorganize()
        Reorganize(Nothing, Nothing)
    End Sub

    Private Sub Reorganize(ByVal sender As Object, ByVal e As EventArgs)
        If Layers Is Nothing Then Exit Sub

        SuspendLayout()
        Dim y As Integer = 0
        For i As Integer = 0 To Layers.Count - 1

            Dim h As Integer = 0
            For j As Integer = 0 To Layers(i).Length - 1
                If Layers(i)(j) Is Nothing Then
                    Continue For
                End If
                h = Math.Max(h, Layers(i)(j).Height + Layers(i)(j).Margin.Top + Layers(i)(j).Margin.Bottom)
            Next
            For j As Integer = 0 To Layers(i).Length - 1
                If Layers(i)(j) Is Nothing Then
                    Continue For
                End If
                Select Case Layers(i).Align
                    Case VisualStyles.VerticalAlignment.Top
                        Layers(i)(j).Top = y + Layers(i)(j).Margin.Top
                    Case VisualStyles.VerticalAlignment.Center
                        Layers(i)(j).Top = CInt(y + (h - Layers(i)(j).Height + Layers(i)(j).Margin.Top - Layers(i)(j).Margin.Bottom) / 2)
                    Case VisualStyles.VerticalAlignment.Bottom
                        Layers(i)(j).Top = y + h - Layers(i)(j).Margin.Bottom - Layers(i)(j).Height
                End Select
            Next
            y += h
        Next

        If WrapContent Then
            ClientSize = New Size(ClientSize.Width, y)
        End If
        _WrapHeight = y

        ResumeLayout(False)
    End Sub

    ''' <summary>
    ''' Retire le contrôle sélectionné de la logique de positionnement du conteneur et supprime les gestionnaires d'événements associés.
    ''' </summary>
    ''' <param name="control"></param>
    Public Sub RemoveControlFromLayer(ByVal control As Control)

        RemoveHandler control.SizeChanged, AddressOf Reorganize
        RemoveHandler control.LocationChanged, AddressOf Reorganize
        RemoveHandler control.MarginChanged, AddressOf Reorganize
        RemoveHandler control.PaddingChanged, AddressOf Reorganize

        If Layers Is Nothing Then Exit Sub

        RemoveControl(control)

    End Sub

    Private Sub VerticalLinearLayout_ControlRemoved(sender As Object, e As ControlEventArgs) Handles Me.ControlRemoved
        RemoveControlFromLayer(e.Control)
    End Sub

    ''' <summary>
    ''' Retire le contrôle et met à jour la liste des contrôles de la ligne associée.
    ''' </summary>
    ''' <param name="control"></param>
    Private Sub RemoveControl(ByVal control As Control)

        Dim cLayers As New List(Of Layer)(Layers)

        For i As Integer = 0 To cLayers.Count - 1
            If cLayers(i).Contains(control) Then
                If cLayers(i).Length = 1 Then
                    Layers.Remove(cLayers(i))
                Else
                    Dim ctrls(cLayers(i).Length - 2) As Control
                    Dim j As Integer = 0
                    For Each ctrl As Control In cLayers(i)
                        If ctrl IsNot control Then
                            ctrls(j) = ctrl
                            j += 1
                        End If
                    Next
                    Layers(i).Controls = ctrls
                End If
            End If
        Next
    End Sub

    '  Private Sub VerticalLinearLayout_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
    '  If Layers Is Nothing Then
    '  InitializeLayers()
    '  End If
    '  End Sub

    Private Sub VerticalLinearLayout_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Layers Is Nothing Then
            Exit Sub
        End If
        If WrapContent Then
            Reorganize(Nothing, Nothing)
        End If
    End Sub

End Class
''' <summary>
''' Ligne d'affichage
''' </summary>
Public Class Layer
    Implements IEnumerable(Of Control)

    Public Property Controls As Control()

    Public Property Align As VisualStyles.VerticalAlignment

    Public Sub New(ParamArray controls As Control())
        Me.New(VisualStyles.VerticalAlignment.Top, controls)
    End Sub

    Public Sub New(ByVal align As VisualStyles.VerticalAlignment, ParamArray controls As Control())
        Me.Align = align
        Me.Controls = controls
    End Sub

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Controls.GetEnumerator
    End Function

    Public Function GetEnumerator() As IEnumerator(Of Control) Implements IEnumerable(Of Control).GetEnumerator
        Return Controls.AsEnumerable.GetEnumerator
    End Function

    Public ReadOnly Property Length As Integer
        Get
            Return Controls.Length
        End Get
    End Property

    Public Property Item(ByVal i As Integer) As Control
        Get
            Return Controls(i)
        End Get
        Set(value As Control)
            Controls(i) = value
        End Set
    End Property

    Public Shared Widening Operator CType(ByVal ctrl As Control) As Layer
        Return New Layer(VisualStyles.VerticalAlignment.Top, ctrl)
    End Operator


End Class
