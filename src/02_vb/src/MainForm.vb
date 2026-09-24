Option Explicit On
Option Strict Off

Imports System.Text
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Diagnostics

' Window logic: keep the screen, the limits and the model in step.
'
' The order on every change is always the same:
'   1. read the controls into the design         (Collect)
'   2. recompute the allowed ranges              (ApplyLimits, may clamp a value)
'   3. read the controls again and redraw        (Collect + Refresh)
' Clamping in step 2 raises ValueChanged again, so the whole cycle runs behind one
' re-entrancy guard.
Partial Class MainForm
    Inherits Form

    Private _app As Object = Nothing
    Private _masterDoc As Object = Nothing
    Private _mainDoc As Object = Nothing   ' main_assembly.CATProduct (table + robot)
    Private _masterPart As Object = Nothing
    Private _tableDoc As Object = Nothing
    ' The robot assembly is opened LAZILY - only when the tool has to be swapped.
    ' Somebody working on the table alone should not wait for it.
    Private _robotDoc As Object = Nothing
    Private _params As MasterParameters = Nothing

    Private ReadOnly _design As New TableDesign()
    Private _busy As Boolean = False
    Private _connected As Boolean = False

    ' True while the screen holds values the model has not been told yet. Only when it is
    ' False may a difference between the two readout columns be called a disagreement.
    Private _dirty As Boolean = True

    ' Optional parameters the master does not have yet - the app skips those on Apply.
    Private ReadOnly _missing As New List(Of String)()

    Public Sub New()
        InitializeComponent()
        WireInputs()
        SeedDefaults()
        ' The inputs work offline on purpose: the whole table can be configured and
        ' checked before CATIA is even running. Only writing needs the connection.
        SetInputsEnabled(True)
        btnReload.Enabled = False
        sketch.Design = _design
        RefreshEverything()
    End Sub

    ' --- wiring ---

    Private Function AllSliders() As ParameterSlider()
        Return New ParameterSlider() {sldLength, sldWidth, sldHeight, sldUpperPlate, sldLowerPlate,
                                      sldOffsetX, sldOffsetY, sldBaseOverhang,
                                      sldJ2J3, sldJ3J5,
                                      sldP1X, sldP1Y, sldP1Z,
                                      sldP3X, sldP3Y, sldP3Z,
                                      sldP6X, sldP6Y, sldP6Z,
                                      sldP2P3, sldP2P4,
                                      sldSimLeg, sldSimSpeed}
    End Function

    Private Function AllPickers() As OptionPicker()
        Return New OptionPicker() {optProfile, optSimplification, optShelfVariant,
                                   optRobotMount, optObjectType}
    End Function

    Private Sub WireInputs()
        For Each s As ParameterSlider In AllSliders()
            AddHandler s.ValueChanged, AddressOf OnInputChanged
            AddHandler s.Activated, AddressOf OnInputActivated
        Next
        For Each p As OptionPicker In AllPickers()
            AddHandler p.SelectionChanged, AddressOf OnInputChanged
            AddHandler p.Activated, AddressOf OnInputActivated
        Next
        AddHandler tabs.SelectedIndexChanged, AddressOf OnTabChanged
    End Sub

    ' Starting point before CATIA is asked: the values the master currently ships with.
    Private Sub SeedDefaults()
        _busy = True
        optProfile.SelectedValue = _design.ProfileName
        optSimplification.SelectedValue = If(_design.Simplified, "true", "false")
        optShelfVariant.SelectedValue = _design.LowerPlateVariant
        optRobotMount.SelectedValue = _design.RobotMountVariant

        ' Values FIRST, limits after. The other way round - which this was until
        ' 29.08.2026 - hands every slider a range while it is still brand new, and
        ' ParameterSlider.SetRange takes the MINIMUM as the current value when Minimum and
        ' Maximum are both still zero. So each slider snapped to the bottom of its range
        ' before anything sensible was ever written into it, and the worktop opened at 5 mm
        ' rather than the 15 the design carries. Cosmetic until someone presses Apply on it.
        sldBaseOverhang.Value = _design.BaseOverhang
        sldLength.Value = _design.Length
        sldWidth.Value = _design.Width
        sldHeight.Value = _design.Height
        sldUpperPlate.Value = _design.UpperPlateThickness
        sldLowerPlate.Value = _design.LowerPlateThickness
        sldOffsetX.Value = _design.UserOffsetX
        sldOffsetY.Value = _design.UserOffsetY
        sldJ2J3.Value = _design.LenJ2J3
        sldJ3J5.Value = _design.LenJ3J5
        sldP1X.Value = _design.P1X
        sldP1Y.Value = _design.P1Y
        sldP1Z.Value = _design.P1Z
        sldP3X.Value = _design.P3X
        sldP3Y.Value = _design.P3Y
        sldP3Z.Value = _design.P3Z
        sldP6X.Value = _design.P6X
        sldP6Y.Value = _design.P6Y
        sldP6Z.Value = _design.P6Z
        sldP2P3.Value = _design.P2P3
        sldP2P4.Value = _design.P2P4
        sldSimLeg.Value = _design.SimLeg
        sldSimSpeed.Value = _design.SimSpeed
        optObjectType.SelectedValue = _design.ObjectType

        ' Now the ranges, computed from values that exist.
        ApplyLimits()
        _busy = False
    End Sub

    Private Sub SetInputsEnabled(ByVal enabled As Boolean)
        For Each s As ParameterSlider In AllSliders()
            s.Enabled = enabled
        Next
        For Each p As OptionPicker In AllPickers()
            p.Enabled = enabled
        Next
        btnApply.Enabled = enabled
        btnReload.Enabled = enabled
    End Sub

    ' --- the change cycle ---

    Private Sub OnInputChanged(ByVal sender As Object, ByVal e As EventArgs)
        If _busy Then Return
        _busy = True
        Try
            Collect()
            ApplyLimits()   ' may pull a value back into range
            Collect()
            _dirty = True
            RefreshEverything()
        Finally
            _busy = False
        End Try
    End Sub

    Private Sub OnInputActivated(ByVal sender As Object, ByVal e As EventArgs)
        Dim key As String = ""
        If TypeOf sender Is ParameterSlider Then key = DirectCast(sender, ParameterSlider).SketchKey
        If TypeOf sender Is OptionPicker Then key = DirectCast(sender, OptionPicker).SketchKey
        If key = "" Then Return
        sketch.View = SketchPanel.ViewFor(key)
        sketch.Highlight = key
    End Sub

    Private Sub OnTabChanged(ByVal sender As Object, ByVal e As EventArgs)
        sketch.Highlight = ""
        ' Tabs are Table, Robot, Tool, Path. The Tool tab has no drawing of its own, so
        ' it borrows the arm view - which is where the gripper hangs, and therefore the
        ' picture somebody choosing an object actually wants to see.
        Select Case tabs.SelectedIndex
            Case 1 : sketch.View = SketchView.Plan
            Case 2 : sketch.View = SketchView.Arm
            Case 3 : sketch.View = SketchView.Path
            Case Else : sketch.View = SketchView.Front
        End Select
    End Sub

    ' Controls -> design.
    Private Sub Collect()
        _design.Length = sldLength.Value
        _design.Width = sldWidth.Value
        _design.Height = sldHeight.Value
        _design.UpperPlateThickness = sldUpperPlate.Value
        _design.LowerPlateThickness = sldLowerPlate.Value
        _design.UserOffsetX = sldOffsetX.Value
        _design.UserOffsetY = sldOffsetY.Value
        _design.BaseOverhang = sldBaseOverhang.Value
        ' MountPitchLong/Short and MountHoleDiameter are no longer UI controls - they keep
        ' the fixed raster of the assigned robot (RobotCatalog.Assigned), and Apply still
        ' writes them so the master always carries 205 / 160 / 11.

        If optProfile.SelectedValue <> "" Then _design.ProfileName = optProfile.SelectedValue
        If optShelfVariant.SelectedValue <> "" Then _design.LowerPlateVariant = optShelfVariant.SelectedValue
        If optRobotMount.SelectedValue <> "" Then _design.RobotMountVariant = optRobotMount.SelectedValue
        _design.Simplified = (optSimplification.SelectedValue = "true")

        _design.LenJ2J3 = sldJ2J3.Value
        _design.LenJ3J5 = sldJ3J5.Value
        _design.P1X = sldP1X.Value
        _design.P1Y = sldP1Y.Value
        _design.P1Z = sldP1Z.Value
        _design.P3X = sldP3X.Value
        _design.P3Y = sldP3Y.Value
        _design.P3Z = sldP3Z.Value
        _design.P6X = sldP6X.Value
        _design.P6Y = sldP6Y.Value
        _design.P6Z = sldP6Z.Value
        _design.P2P3 = sldP2P3.Value
        _design.P2P4 = sldP2P4.Value
        _design.SimLeg = sldSimLeg.Value
        _design.SimSpeed = sldSimSpeed.Value

        ' The object decides what the task's own 4a and c are, so changing it re-proposes
        ' both segments. Only on a CHANGE: re-applying the suggestion on every keystroke
        ' would make the two sliders impossible to move, and section 12 asks for them to
        ' be changeable.
        Dim wantType As String = optObjectType.SelectedValue
        If wantType <> "" AndAlso wantType <> _design.ObjectType Then
            _design.ObjectType = wantType
            _design.P2P4 = _design.SuggestedP2P4()
            _design.P2P3 = _design.SuggestedP2P3()
            sldP2P4.Value = _design.P2P4
            sldP2P3.Value = _design.P2P3
        End If
    End Sub

    ' Design -> the ranges every control is allowed to move in.
    Private Sub ApplyLimits()
        Push(sldLength, DesignLimits.LengthRange(_design))
        Push(sldWidth, DesignLimits.WidthRange(_design))
        Push(sldHeight, DesignLimits.HeightRange(_design))
        Push(sldUpperPlate, DesignLimits.UpperPlateRange(_design))
        Push(sldLowerPlate, DesignLimits.LowerPlateRange(_design))
        Push(sldBaseOverhang, DesignLimits.BaseOverhangRange())
        Push(sldOffsetX, DesignLimits.OffsetXRange(_design))
        Push(sldOffsetY, DesignLimits.OffsetYRange(_design))
        Push(sldJ2J3, DesignLimits.ArmJ2J3Range(_design))
        Push(sldJ3J5, DesignLimits.ArmJ3J5Range(_design))
        ' All three points share one set of envelopes: they are the same kind of thing -
        ' a tool-tip target in the robot's own frame - and the arm does not care which of
        ' them it is being asked to reach.
        Push(sldP1X, DesignLimits.PathForwardRange(_design))
        Push(sldP1Y, DesignLimits.PathSidewaysRange(_design))
        Push(sldP1Z, DesignLimits.PathHeightRange(_design))
        Push(sldP3X, DesignLimits.PathForwardRange(_design))
        Push(sldP3Y, DesignLimits.PathSidewaysRange(_design))
        Push(sldP3Z, DesignLimits.PathHeightRange(_design))
        Push(sldP6X, DesignLimits.PathForwardRange(_design))
        Push(sldP6Y, DesignLimits.PathSidewaysRange(_design))
        Push(sldP6Z, DesignLimits.PathHeightRange(_design))
        Push(sldP2P3, DesignLimits.SegmentRange(400))
        Push(sldP2P4, DesignLimits.SegmentRange(600))
        Push(sldSimLeg, DesignLimits.SimLegRange())
        Push(sldSimSpeed, DesignLimits.SimSpeedRange())

        ' An offset that the chosen base position ignores is greyed out instead of
        ' silently doing nothing - that ambiguity is what made the old window confusing.
        UpdateOffsetAvailability(sldOffsetX, _design.UsesOffsetX, "along the length")
        UpdateOffsetAvailability(sldOffsetY, _design.UsesOffsetY, "across the width")
    End Sub

    Private Sub UpdateOffsetAvailability(ByVal s As ParameterSlider, ByVal used As Boolean, ByVal axis As String)
        If used Then
            s.SetAvailable()
        Else
            s.SetUnavailable("Not used: this base position keeps the robot centred " & axis & ".")
        End If
    End Sub

    Private Sub Push(ByVal s As ParameterSlider, ByVal r As ValueRange)
        s.SetRange(r.Minimum, r.Maximum, r.Reason)
    End Sub

    Private Sub RefreshEverything()
        sketch.Design = _design
        sketch.Invalidate()
        ShowPreview()
        ShowChecks()
    End Sub

    ' --- derived values ---

    ' The robot readout: the assigned robot's own figures, plus where they came from - and,
    ' when the inputs have been moved off them, what they are now instead.
    Private Sub ShowRobotSpec()
        Dim r As RobotInfo = _design.Robot

        If r.HasRaster Then
            lblRobotRaster.Text = DesignLimits.Fmt(r.BoltPitchLong) & " x " &
                                  DesignLimits.Fmt(r.BoltPitchShort) & " mm"
            lblRobotHole.Text = "d " & DesignLimits.Fmt(r.HoleDiameter) & " mm" &
                                If(r.BoltThread = "", "", "   for " & r.BoltThread)
        Else
            lblRobotRaster.Text = "not recorded yet"
            lblRobotHole.Text = "not recorded yet"
        End If

        If r.Reach > 0 Then
            lblRobotReach.Text = DesignLimits.Fmt(r.Reach) & " mm   /   " &
                                 DesignLimits.Fmt(r.Payload) & " kg"
        Else
            lblRobotReach.Text = "-"
        End If

        If r.LenJ2J3 > 0 Then
            lblRobotArm.Text = DesignLimits.Fmt(r.LenJ2J3) & " / " & DesignLimits.Fmt(r.LenJ3J5) & " mm"
        Else
            lblRobotArm.Text = "-   (app default " & DesignLimits.Fmt(_design.Constants.NominalLenJ2J3) &
                               " / " & DesignLimits.Fmt(_design.Constants.NominalLenJ3J5) & ")"
        End If

        Dim note As String = r.Source
        If r.HasRaster AndAlso Not _design.MatchesCatalogue Then
            note = "EDITED - the fields below are " & DesignLimits.Fmt(_design.MountPitchLong) & " x " &
                   DesignLimits.Fmt(_design.MountPitchShort) & ", hole " &
                   DesignLimits.Fmt(_design.MountHoleDiameter) & " mm, not the robot's." &
                   vbCrLf & note
        End If
        lblRobotSource.Text = note
        lblRobotSource.ForeColor = If(r.Verified AndAlso _design.MatchesCatalogue, Theme.Muted, Theme.Bad)
    End Sub

    Private Sub ShowPreview()
        ShowRobotSpec()
        lblGussetSize.Text = DesignLimits.Fmt(_design.GussetSize) & " mm"
        lblGussetCornerSize.Text = DesignLimits.Fmt(_design.GussetSizeCorner) & " mm"
        lblLowerPlateLen.Text = DesignLimits.Fmt(_design.LowerPlateLength) & " mm"
        lblLowerPlateWid.Text = DesignLimits.Fmt(_design.LowerPlateWidth) & " mm"

        lblGussetModel.Text = FromModel(MasterParameters.P_GUSSET_SIZE)
        lblGussetCornerModel.Text = FromModel(MasterParameters.P_GUSSET_SIZE_CORNER)
        lblLowerLenModel.Text = FromModel(MasterParameters.P_LOW_PLATE_LEN)
        lblLowerWidModel.Text = FromModel(MasterParameters.P_LOW_PLATE_WID)

        lblAxisOffset.Text = Pair(_design.AxisOffsetX, _design.AxisOffsetY)
        lblMountPitch.Text = Pair(_design.MountPitchX, _design.MountPitchY)
        lblMountEdge.Text = Pair(_design.MountEdgeX, _design.MountEdgeY)
        lblAxisOffsetModel.Text = PairFromModel(MasterParameters.P_AXIS_OFF_X, MasterParameters.P_AXIS_OFF_Y)
        lblMountPitchModel.Text = PairFromModel(MasterParameters.P_PITCH_X, MasterParameters.P_PITCH_Y)
        lblMountEdgeModel.Text = PairFromModel(MasterParameters.P_MOUNT_EDGE_X, MasterParameters.P_MOUNT_EDGE_Y)

        ' A rule that drifted from rules/R_RobotMountOffset.txt shows up right here, as a
        ' mismatch between the two columns. Only meaningful once the model has been told
        ' the current values - before Apply the columns are supposed to differ.
        MarkAgreement(lblAxisOffsetModel, lblAxisOffset)
        MarkAgreement(lblMountPitchModel, lblMountPitch)
        MarkAgreement(lblMountEdgeModel, lblMountEdge)
    End Sub

    Private Sub MarkAgreement(ByVal modelLabel As Label, ByVal previewLabel As Label)
        ' "-" means the parameter is not in the master yet - nothing to disagree about.
        If _params Is Nothing OrElse _dirty OrElse modelLabel.Text = "-" Then
            modelLabel.ForeColor = Theme.Ink
            Return
        End If
        modelLabel.ForeColor = If(modelLabel.Text = previewLabel.Text, Theme.Ink, Theme.Bad)
    End Sub

    Private Shared Function Pair(ByVal a As Double, ByVal b As Double) As String
        Return DesignLimits.Fmt(a) & " / " & DesignLimits.Fmt(b) & " mm"
    End Function

    ' A rule older than rules/R_RobotMountOffset.txt does not know the new standing
    ' positions and quietly drops them into its else branch - the base then sits right but
    ' turned the wrong way. Compared only once the master already carries this variant
    ' name, because before an Apply the model still holds the previous one.
    Private Function RuleIsStale() As Boolean
        If _params Is Nothing OrElse _dirty Then Return False

        Dim modelVariant As String = _params.ReadText(MasterParameters.P_ROB_MNT).Trim()
        If String.Equals(modelVariant, "corner", StringComparison.OrdinalIgnoreCase) Then
            modelVariant = "cornerLongAxis"
        End If
        If Not String.Equals(modelVariant, _design.RobotMountVariant, StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        Dim px As Double = ModelNumber(MasterParameters.P_PITCH_X)
        Dim py As Double = ModelNumber(MasterParameters.P_PITCH_Y)
        If px <= 0 OrElse py <= 0 Then Return False   ' parameter missing - nothing to compare

        Return Math.Abs(px - _design.MountPitchX) > 0.01 OrElse Math.Abs(py - _design.MountPitchY) > 0.01
    End Function

    Private Function ModelNumber(ByVal paramName As String) As Double
        If _params Is Nothing Then Return 0
        Dim raw As String = _params.ReadNumber(paramName)
        If String.IsNullOrEmpty(raw) Then Return 0
        Dim d As Double
        If Double.TryParse(Num.Normalize(raw), Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture, d) Then Return d
        Return 0
    End Function

    Private Function PairFromModel(ByVal nameA As String, ByVal nameB As String) As String
        If _params Is Nothing Then Return "not connected"
        Dim a As String = _params.ReadNumber(nameA)
        Dim b As String = _params.ReadNumber(nameB)
        If String.IsNullOrEmpty(a) OrElse String.IsNullOrEmpty(b) Then Return "-"
        Return DesignLimits.Fmt(CDbl(Val(Num.Normalize(a)))) & " / " &
               DesignLimits.Fmt(CDbl(Val(Num.Normalize(b)))) & " mm"
    End Function

    ' Value a rule or formula produced in the master. "-" means the parameter is not
    ' there yet (for example the master macro has not been run).
    Private Function FromModel(ByVal paramName As String) As String
        If _params Is Nothing Then Return "not connected"
        Dim v As String = _params.ReadNumber(paramName)
        If String.IsNullOrEmpty(v) Then Return "-"
        Dim d As Double
        If Double.TryParse(Num.Normalize(v), Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture, d) Then
            Return DesignLimits.Fmt(d) & " mm"
        End If
        Return v
    End Function

    ' --- checks ---

    Private Sub ShowChecks()
        Dim checks As List(Of DesignCheck) = DesignLimits.Review(_design)

        ' A sketch still dimensioned from the plate centre cannot express a base past the
        ' middle: the dimension has no sign, so CATIA takes the absolute value and mirrors
        ' the base back. Refuse exactly that case rather than write a model that looks
        ' applied and is not.
        If _connected AndAlso Not _params.Exists(MasterParameters.P_MOUNT_EDGE_X) Then
            If _design.AxisOffsetX < 0 OrElse _design.AxisOffsetY < 0 Then
                checks.Add(New DesignCheck(CheckLevel.Problem,
                    "This puts the base past the middle, but the master has no mount_edge_x/y - " &
                    "the hole sketch would mirror it back. Run 01_master_parameters, then re-dimension the sketch."))
            Else
                checks.Add(New DesignCheck(CheckLevel.Info,
                    "No mount_edge_x/y in the master yet, so the base cannot cross the middle of the plate. " &
                    "Run 01_master_parameters and re-dimension the hole sketch to it."))
            End If
        End If

        ' The master and the assembly can name different grippers. Say so out loud - read
        ' back on its own just looks like the window got it wrong.
        '
        ' Warning, NOT Problem: a Problem disables Apply, and Apply is exactly what settles
        ' the disagreement - ToolSwitcher fits the gripper the master names. Blocking it
        ' would leave the user stuck with the very state the message complains about.
        If _fittedMismatch <> "" Then
            checks.Add(New DesignCheck(CheckLevel.Warning, _fittedMismatch))
        End If

        If _missing.Count > 0 Then
            checks.Add(New DesignCheck(CheckLevel.Info,
                "Skipped, not in the master yet: " & String.Join(", ", _missing.ToArray()) & "."))
        End If
        If Not _connected Then
            checks.Insert(0, New DesignCheck(CheckLevel.Info,
                "Not connected yet. The ranges below come from the default table."))
        ElseIf RuleIsStale() Then
            checks.Insert(0, New DesignCheck(CheckLevel.Problem,
                "The master's R_RobotMountOffset does not know this standing position - it produced " &
                Pair(ModelNumber(MasterParameters.P_PITCH_X), ModelNumber(MasterParameters.P_PITCH_Y)) &
                " instead of " & Pair(_design.MountPitchX, _design.MountPitchY) &
                ". Paste rules/R_RobotMountOffset.txt into the rule."))
        End If

        flowChecks.SuspendLayout()
        flowChecks.Controls.Clear()
        If checks.Count = 0 Then
            flowChecks.Controls.Add(CheckRow(New DesignCheck(CheckLevel.Info,
                "Nothing to flag - every value is inside its allowed range.")))
        Else
            For Each c As DesignCheck In checks
                flowChecks.Controls.Add(CheckRow(c))
            Next
        End If
        flowChecks.ResumeLayout()

        Dim blocking As Integer = 0
        For Each c As DesignCheck In checks
            If c.Level = CheckLevel.Problem Then blocking += 1
        Next

        btnApply.Enabled = _connected AndAlso blocking = 0
        If blocking > 0 Then
            lblApplyNote.ForeColor = Theme.Bad
            lblApplyNote.Text = "Fix the highlighted problem before applying - CATIA would fail the update."
            btnApply.BackColor = Theme.Border
        ElseIf Not _connected Then
            lblApplyNote.ForeColor = Theme.Muted
            lblApplyNote.Text = "Connect to CATIA first. The master part and the table assembly are opened for you."
            btnApply.BackColor = Theme.Border
        Else
            lblApplyNote.ForeColor = Theme.Muted
            lblApplyNote.Text = "Writes the parameters into master.CATPart, swaps the profile section and rebuilds the assembly."
            btnApply.BackColor = Theme.Accent
        End If
    End Sub

    Private Function CheckRow(ByVal c As DesignCheck) As Control
        Dim col As Color = Theme.Muted
        Dim mark As String = "i"
        Select Case c.Level
            Case CheckLevel.Problem
                col = Theme.Bad
                mark = "!"
            Case CheckLevel.Warning
                col = Theme.Warn
                mark = "!"
        End Select

        Dim row As New Label()
        row.AutoSize = False
        row.Width = flowChecks.ClientSize.Width - 24
        row.Height = 34
        row.Margin = New Padding(0, 0, 0, 4)
        row.Font = Theme.Small
        row.ForeColor = col
        row.Text = "  " & mark & "   " & c.Text
        row.TextAlign = ContentAlignment.MiddleLeft
        Theme.Tips.SetToolTip(row, c.Text)
        Return row
    End Function

    Private Sub SetStatus(ByVal text As String, ByVal ok As Boolean)
        lblStatus.Text = text
        lblStatus.ForeColor = If(ok, Theme.Ok, Theme.Bad)
    End Sub

    ' --- connecting ---

    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            OleMessageFilter.Register()
            _app = CatiaConnector.Connect()

            _masterDoc = CatiaConnector.GetOrOpenDocument(_app, AppConfig.MasterPartPath())
            _masterPart = _masterDoc.Part
            _params = New MasterParameters(_masterPart)

            ' Table first (the switchers act on it), then main_assembly on top so it
            ' reuses the table as an already-open component instead of a second copy.
            If AppConfig.OpenTableAssembly() Then
                Try
                    _tableDoc = CatiaConnector.GetOrOpenDocument(_app, AppConfig.TableAssemblyPath())
                Catch ex As Exception
                    _tableDoc = Nothing
                End Try
            End If

            If AppConfig.OpenMainAssembly() Then
                Try
                    _mainDoc = CatiaConnector.GetOrOpenDocument(_app, AppConfig.MainAssemblyPath())
                Catch ex As Exception
                    _mainDoc = Nothing
                End Try
            End If

            ' Reading the model needs the master active - parameter and GSD access both
            ' go through the active document.
            CatiaConnector.Activate(_masterDoc)

            _design.Constants.ReadFrom(_params)
            DetectMissingOptionals()
            LoadValuesFromModel()

            ' ...but what the user should be LOOKING at is the whole cell, not the
            ' skeleton part. main_assembly has been opened a few lines above and then
            ' immediately hidden behind the master, which is why connecting used to leave
            ' master.CATPart on screen with nothing recognisable in it.
            '
            ' Safe to leave the assembly active: btnApply_Click activates the master again
            ' at its start, before it writes anything.
            Dim shown As String = "master.CATPart"
            If _mainDoc IsNot Nothing Then
                CatiaConnector.Activate(_mainDoc)
                shown = "main_assembly.CATProduct"
            ElseIf _tableDoc IsNot Nothing Then
                CatiaConnector.Activate(_tableDoc)
                shown = "table_mainAssembly.CATProduct"
            End If

            ' Activate only picks the document INSIDE CATIA; this raises CATIA over this
            ' window, which is the difference between "it is open somewhere" and "you can
            ' see it".
            Try
                CatiaConnector.BringToFront()
            Catch
                ' Not being able to raise the window is cosmetic; the document is active.
            End Try

            _connected = True
            SetInputsEnabled(True)
            SetStatus("Connected. Showing " & shown, True)
            RefreshEverything()
        Catch ex As Exception
            SetStatus("Could not connect.", False)
            MessageBox.Show("Connecting failed:" & vbCrLf & vbCrLf & ex.Message,
                            "CATIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ' Optional parameters (sections 6 and 12) only exist once they are modelled. Locking
    ' the input is more honest than accepting a number the app will drop.
    Private Sub DetectMissingOptionals()
        _missing.Clear()
        CheckOptional(sldJ2J3, MasterParameters.P_LEN_J2J3)
        CheckOptional(sldJ3J5, MasterParameters.P_LEN_J3J5)
        CheckOptional(sldP1X, MasterParameters.P_P1X)
        CheckOptional(sldP1Y, MasterParameters.P_P1Y)
        CheckOptional(sldP1Z, MasterParameters.P_P1Z)
        CheckOptional(sldP3X, MasterParameters.P_P3X)
        CheckOptional(sldP3Y, MasterParameters.P_P3Y)
        CheckOptional(sldP3Z, MasterParameters.P_P3Z)
        CheckOptional(sldP6X, MasterParameters.P_P6X)
        CheckOptional(sldP6Y, MasterParameters.P_P6Y)
        CheckOptional(sldP6Z, MasterParameters.P_P6Z)
        CheckOptional(sldP2P3, MasterParameters.P_P2P3)
        CheckOptional(sldP2P4, MasterParameters.P_P2P4)
        CheckOptional(sldSimLeg, MasterParameters.P_SIM_LEG)
        CheckOptional(sldSimSpeed, MasterParameters.P_SIM_SPEED)
        ' object_type has no slider to grey out, so it is only recorded as missing.
        If Not _params.Exists(MasterParameters.P_OBJECT_TYPE) Then
            _missing.Add(MasterParameters.P_OBJECT_TYPE)
        End If
    End Sub

    Private Sub CheckOptional(ByVal s As ParameterSlider, ByVal paramName As String)
        If _params.Exists(paramName) Then
            s.SetAvailable()
        Else
            s.SetUnavailable("'" & paramName & "' is not in the master yet - this value is skipped.")
            _missing.Add(paramName)
        End If
    End Sub

    Private Sub btnReload_Click(sender As Object, e As EventArgs) Handles btnReload.Click
        If _params Is Nothing Then Return
        Try
            ' Reload is the button you press when the model changed under the application -
            ' a macro run in CATIA beside it - so it is also the one place where a cached
            ' "this parameter does not exist" could be stale. Drop the cache first.
            _params.ForgetCache()
            LoadValuesFromModel()
            SetStatus("Values reloaded from the model.", True)
        Catch ex As Exception
            MessageBox.Show("Reading the model failed:" & vbCrLf & vbCrLf & ex.Message,
                            "CATIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadValuesFromModel()
        _busy = True
        Try
            ' Options first: they decide the ranges of everything else.
            Dim profileName As String = _params.ReadText(MasterParameters.P_PROFILE)
            If profileName <> "" Then optProfile.SelectedValue = profileName.Trim()
            optProfile.SelectFirstIfEmpty()

            Dim shelfMount As String = _params.ReadText(MasterParameters.P_LOWER_M)
            If shelfMount <> "" Then optShelfVariant.SelectedValue = shelfMount.Trim()
            optShelfVariant.SelectFirstIfEmpty()

            ' "corner" is the retired name of cornerLongAxis. Mapping it here keeps an
            ' older master from silently landing on the first item in the list.
            Dim mount As String = _params.ReadText(MasterParameters.P_ROB_MNT).Trim()
            If String.Equals(mount, "corner", StringComparison.OrdinalIgnoreCase) Then mount = "cornerLongAxis"
            If mount <> "" Then optRobotMount.SelectedValue = mount
            optRobotMount.SelectFirstIfEmpty()

            optSimplification.SelectedValue = NormalizeBool(_params.ReadText(MasterParameters.P_SIMP))
            optSimplification.SelectFirstIfEmpty()

            ' NO ApplyLimits HERE. It used to sit on this line, and it computed every
            ' range from a _design that still held the PREVIOUS screen - the model has not
            ' been read yet at this point. UpperPlateRange and LowerPlateRange depend on
            ' the table height, so on a cold connect a 225 mm stale height collapsed them
            ' to (5,5) and (2,2), and the Pulls below then clamped the real 23.8 / 15.8 to
            ' 5 / 2. The second ApplyLimits widened the ranges again, but by then the
            ' numbers were gone - which is why pressing "Reload from model" afterwards
            ' appeared to fix a model that had been read correctly all along.
            '
            ' The invariant this file states at the top - collect, then limit, then
            ' collect - is about the SCREEN's own values. A value arriving from the model
            ' has to land first and be limited afterwards, which is what the Pulls below
            ' and the ApplyLimits at the end now do.
            Collect()

            Pull(sldLength, MasterParameters.P_LENGTH)
            Pull(sldWidth, MasterParameters.P_WIDTH)
            Pull(sldHeight, MasterParameters.P_HEIGHT)
            Pull(sldUpperPlate, MasterParameters.P_UP_T)
            Pull(sldLowerPlate, MasterParameters.P_LOW_T)
            Pull(sldOffsetX, MasterParameters.P_OFFX)
            Pull(sldOffsetY, MasterParameters.P_OFFY)
            Pull(sldJ2J3, MasterParameters.P_LEN_J2J3)
            Pull(sldJ3J5, MasterParameters.P_LEN_J3J5)
            Pull(sldP1X, MasterParameters.P_P1X)
            Pull(sldP1Y, MasterParameters.P_P1Y)
            Pull(sldP1Z, MasterParameters.P_P1Z)
            Pull(sldP3X, MasterParameters.P_P3X)
            Pull(sldP3Y, MasterParameters.P_P3Y)
            Pull(sldP3Z, MasterParameters.P_P3Z)
            Pull(sldP6X, MasterParameters.P_P6X)
            Pull(sldP6Y, MasterParameters.P_P6Y)
            Pull(sldP6Z, MasterParameters.P_P6Z)
            Pull(sldP2P3, MasterParameters.P_P2P3)
            Pull(sldP2P4, MasterParameters.P_P2P4)
            Pull(sldSimLeg, MasterParameters.P_SIM_LEG)
            Pull(sldSimSpeed, MasterParameters.P_SIM_SPEED)

            ' tool_len is set by the rule R_Tool and only ever READ here. It is paired with
            ' the object type it belongs to, because picking the other object in the window
            ' does not change what is fitted in CATIA - see TableDesign.ToolLen.
            Try
                Dim tl As String = _params.ReadNumber(MasterParameters.P_TOOL_LEN)
                Dim tlv As Double
                If Not String.IsNullOrEmpty(tl) AndAlso
                   Double.TryParse(Num.Normalize(tl), Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture, tlv) Then
                    _design.ToolLenFromModel = tlv
                    _design.ToolLenForType = _params.ReadText(MasterParameters.P_OBJECT_TYPE).Trim()
                End If
            Catch
                ' An older master has no tool_len; the catalogue value stands in.
            End Try

            ' object_type is a string, so it does not go through Pull. The model wins on
            ' connect, exactly as it does for the profile and the variants.
            Try
                Dim ot As String = _params.ReadText(MasterParameters.P_OBJECT_TYPE)
                If ot <> "" Then optObjectType.SelectedValue = ot.Trim()
                _fittedMismatch = FittedToolMismatch(ot.Trim())
            Catch
                ' An older master has no object_type; the picker keeps its default.
            End Try

            Collect()
            ApplyLimits()
            Collect()
            _dirty = False   ' screen and model now hold the same thing
        Finally
            _busy = False
        End Try
        RefreshEverything()
    End Sub

    ' Set on connect: the master says one gripper and the assembly holds another.
    ' Empty when they agree, or when the robot assembly is not open to be asked.
    Private _fittedMismatch As String = ""

    ' The master's object_type is not the whole truth. It is what R_Tool branches on, but
    ' the gripper that is physically in the assembly is the part number of the tool_1
    ' instance, and the two drift apart the moment one is changed without the other -
    ' a hand-made Replace Component, say. When that happened on 29.08.2026 the window
    ' read back "1" and showed the MPG while sub_tool_2 was mounted, and nothing on screen
    ' explained it.
    '
    ' Only asks a robot assembly that is ALREADY open: opening it here would make every
    ' connect wait several seconds for a document most sessions never touch.
    Private Function FittedToolMismatch(ByVal masterType As String) As String
        If masterType = "" Then Return ""

        Dim rd As Object = _robotDoc
        If rd Is Nothing Then
            rd = CatiaConnector.FindOpenDocument(_app, AppConfig.RobotAssemblyPath())
        End If
        If rd Is Nothing Then Return ""

        Dim fitted As String = ToolSwitcher.FittedObjectType(rd)
        If fitted = "" OrElse fitted = masterType Then Return ""

        Return "The master says object_type = " & masterType & " (" &
               GripperCatalog.ByType(masterType).Caption & "), but " &
               ToolSwitcher.INSTANCE_NAME & " in the robot assembly is sub_tool_" & fitted &
               " (" & GripperCatalog.ByType(fitted).Caption & ")." & vbCrLf &
               "The window follows the master. Apply will fit the gripper it names; " &
               "pick " & fitted & " first if the assembly is the one that is right."
    End Function

    ' Reads one number out of the model. A value the model holds but the UI no longer
    ' allows (an older, broken table) is clamped and reported as a change - but by the
    ' ApplyLimits that follows, against the range derived from the model's own values,
    ' NOT by whatever range the previous screen happened to leave behind.
    Private Sub Pull(ByVal s As ParameterSlider, ByVal paramName As String)
        Dim raw As String = _params.ReadNumber(paramName)
        If String.IsNullOrEmpty(raw) Then Return
        Dim d As Double
        If Not Double.TryParse(Num.Normalize(raw), Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture, d) Then Return
        s.AcceptModelValue(d)
    End Sub

    Private Shared Function NormalizeBool(ByVal s As String) As String
        If s Is Nothing Then Return "false"
        Dim t As String = s.Trim().ToLowerInvariant()
        If t = "true" OrElse t = "1" Then Return "true"
        Return "false"
    End Function

    ' --- applying ---

    ' The robot assembly, opened on first use, and QUIETLY: the one caller left is Apply,
    ' where a missing robot assembly is an ordinary state (a table-only project has none
    ' yet) and the tool swap simply reports that it could not run.
    '
    ' It used to take a `complain` flag, because the Simulate button wanted a dialog
    ' instead of silence. That button is gone - the simulation is played in DMU - and with
    ' it the only caller that ever passed True.
    Private Function RobotDoc() As Object
        If _robotDoc IsNot Nothing Then Return _robotDoc
        If _app Is Nothing Then Return Nothing

        Dim path As String = AppConfig.RobotAssemblyPath()
        If Not IO.File.Exists(path) Then Return Nothing

        Try
            _robotDoc = CatiaConnector.GetOrOpenDocument(_app, path)
        Catch ex As Exception
            _robotDoc = Nothing
        End Try
        Return _robotDoc
    End Function


    ' NO Simulate BUTTON (removed 30.08.2026, user's decision). It drove the DMU
    ' mechanism straight from here - channel B of TODO2.md phase 15 - by computing the
    ' command values in RobotKinematics and pushing them in. Channel A won: the Knowledge
    ' laws recompute the whole path from KINTime inside the model, and the run is played
    ' and recorded in DMU, where the video has to be made anyway.
    '
    ' Two drivers on one mechanism cannot both hold it. The button had to switch the laws
    ' off before it could move anything - a command a formula owns is read-only - and left
    ' them off, so the model came back from a "successful" simulation in a state the laws
    ' no longer described. RobotKinematics stays: the window still uses it to draw the path
    ' and to refuse an unreachable point before it is ever written to CATIA.

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If _params Is Nothing Then
            MessageBox.Show("Connect to CATIA first.", "CATIA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        Collect()

        ' Backstop: the sliders already make an out-of-range value impossible, so this
        ' only catches combinations no single range can express.
        Dim blocking As New StringBuilder()
        For Each c As DesignCheck In DesignLimits.Review(_design)
            If c.Level = CheckLevel.Problem Then blocking.AppendLine("  - " & c.Text)
        Next
        If blocking.Length > 0 Then
            MessageBox.Show("This table cannot be built:" & vbCrLf & vbCrLf & blocking.ToString(),
                            "Check the inputs", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        Me.Cursor = Cursors.WaitCursor

        ' Timed, in three parts, and the numbers go into the result dialog. "It feels
        ' faster" is not a measurement, and the three phases have completely different
        ' causes - parameter writes are COM name lookups, the swaps are component surgery,
        ' the update is CATIA solving. Without the split, a change that helps one and hurts
        ' another looks like no change at all.
        Dim swAll As Stopwatch = Stopwatch.StartNew()
        Dim msWrite As Long = 0
        Dim msSwap As Long = 0
        Dim msUpdate As Long = 0

        Try
            CatiaConnector.Activate(_masterDoc)

            ' 1) lengths
            _params.WriteLengthMm(MasterParameters.P_LENGTH, _design.Length)
            _params.WriteLengthMm(MasterParameters.P_WIDTH, _design.Width)
            _params.WriteLengthMm(MasterParameters.P_HEIGHT, _design.Height)
            _params.WriteLengthMm(MasterParameters.P_UP_T, _design.UpperPlateThickness)
            _params.WriteLengthMm(MasterParameters.P_LOW_T, _design.LowerPlateThickness)
            _params.WriteLengthMm(MasterParameters.P_OFFX, _design.UserOffsetX)
            _params.WriteLengthMm(MasterParameters.P_OFFY, _design.UserOffsetY)
            ' Robot bolt raster. R_RobotMountOffset turns these two into
            ' robot_mount_pitch_x/y according to the standing position, so the rule stays
            ' untouched - only its inputs moved from hand-set constants to the app.
            _params.WriteLengthMm(MasterParameters.P_MOUNT_LEN, _design.MountPitchLong)
            _params.WriteLengthMm(MasterParameters.P_MOUNT_WID, _design.MountPitchShort)

            ' 2) strings and the boolean
            _params.WriteString(MasterParameters.P_PROFILE, _design.ProfileName)
            _params.WriteString(MasterParameters.P_LOWER_M, _design.LowerPlateVariant)
            _params.WriteString(MasterParameters.P_ROB_MNT, _design.RobotMountVariant)
            _params.WriteBool(MasterParameters.P_SIMP, _design.Simplified)

            ' 3) optional ones, only where the parameter exists
            Dim skipped As New StringBuilder()
            ' A parameter that exists but is read-only (a formula or rule took it over) is
            ' collected here instead of aborting Apply - len_J2J3 / P1_* MUST be free inputs
            ' (t. 2.2, t. 12), so a lock is a real defect, but the rest of the write should
            ' still land and the user should see the whole list, not just the first lock.
            Dim locked As New StringBuilder()
            ' robot_mount_hole_d is new (17.08.2026) and an older master does not have it;
            ' the plate's hole sketch has to be re-dimensioned onto it by hand first.
            TryOptional(MasterParameters.P_MOUNT_HOLE_D, _design.MountHoleDiameter, skipped, locked)
            TryOptional(MasterParameters.P_LEN_J2J3, _design.LenJ2J3, skipped, locked)
            TryOptional(MasterParameters.P_LEN_J3J5, _design.LenJ3J5, skipped, locked)
            TryOptional(MasterParameters.P_P1X, _design.P1X, skipped, locked)
            TryOptional(MasterParameters.P_P1Y, _design.P1Y, skipped, locked)
            TryOptional(MasterParameters.P_P1Z, _design.P1Z, skipped, locked)
            TryOptional(MasterParameters.P_P3X, _design.P3X, skipped, locked)
            TryOptional(MasterParameters.P_P3Y, _design.P3Y, skipped, locked)
            TryOptional(MasterParameters.P_P3Z, _design.P3Z, skipped, locked)
            TryOptional(MasterParameters.P_P6X, _design.P6X, skipped, locked)
            TryOptional(MasterParameters.P_P6Y, _design.P6Y, skipped, locked)
            TryOptional(MasterParameters.P_P6Z, _design.P6Z, skipped, locked)
            TryOptional(MasterParameters.P_P2P3, _design.P2P3, skipped, locked)
            TryOptional(MasterParameters.P_P2P4, _design.P2P4, skipped, locked)
            ' P4P5 is not written: the model derives it from P2P3, because the task figure
            ' puts both stations at one height. Writing it would report as locked.
            '
            ' path_radius and station_height are not written either - the path is no longer
            ' laid out on a circle, so nothing in the master consumes them.
            TryOptional(MasterParameters.P_SIM_LEG, _design.SimLeg, skipped, locked)
            TryOptional(MasterParameters.P_SIM_SPEED, _design.SimSpeed, skipped, locked)
            ' adapter_t is NOT written any more (29.08.2026). Nothing in the master consumes
            ' it since tool_len stopped being a formula over it, so writing it would only
            ' suggest it still drives the plate.

            ' object_type is a STRING, not a length, so it needs its own guarded write.
            ' It is what R_Tool branches on; ToolSwitcher then exchanges the component,
            ' because a rule can set a number but cannot replace a part.
            Try
                If _params.Exists(MasterParameters.P_OBJECT_TYPE) Then
                    _params.WriteString(MasterParameters.P_OBJECT_TYPE, _design.ObjectType)
                Else
                    skipped.Append(" " & MasterParameters.P_OBJECT_TYPE)
                End If
            Catch
                skipped.Append(" " & MasterParameters.P_OBJECT_TYPE)
            End Try

            ' 4) GSD swap of the active profile section, simplified one included.
            '    The join feeds the parts' extrusion, which works on a SURFACE - every
            '    detailed section is an Extract off the imported one. The simplified
            '    section must therefore be a surface too (a Fill over profile_simple_sk),
            '    not the bare sketch; a sketch is a wire and fails the assembly update.
            '    The section offsets themselves belong to rule R_Profile - the app must
            '    not write them, otherwise a dead rule stays invisible.
            msWrite = swAll.ElapsedMilliseconds

            Dim sectionUsed As String = ProfileSwitcher.EnsureJoinHasSingleProfile(
                _masterPart, _design.ProfileName, _design.Simplified)

            ' 4b) swap the gusset kits when the section change moved either bracket size
            '     onto another one. R_Profile and F_gusset_size_corner recompute the
            '     numbers, but which COMPONENT sits in the assembly is not something
            '     Knowledge can change - same reason the profile section needs a switcher.
            '     TWO sizes go in: the 16 post-to-rail brackets follow gusset_size, the 8
            '     corner brackets gusset_size_corner, and on 45x60 / 45x90 those differ.
            '     Runs before the update so the assembly rebuilds once, with the right kits
            '     already in place.
            Dim gussetMsg As String = ""
            Try
                Dim gs As String = _params.ReadNumber(MasterParameters.P_GUSSET_SIZE)
                Dim gsc As String = _params.ReadNumber(MasterParameters.P_GUSSET_SIZE_CORNER)
                ' A master that predates F_gusset_size_corner still has to be usable: fall
                ' back to gusset_size, which is what the assembly held before the corner
                ' brackets existed. The switcher then reports "0 found" for that group.
                If gsc = "" Then gsc = gs
                If gs <> "" AndAlso _tableDoc IsNot Nothing Then
                    gussetMsg = GussetSwitcher.EnsureGussetSize(
                        _app, _tableDoc, CDbl(gs), CDbl(gsc), AppConfig.CadRoot())
                End If
            Catch exG As Exception
                ' Never fail Apply over the gussets: the parameters are already written and
                ' the swap can be redone by hand.
                gussetMsg = "Gussets: swap failed (" & exG.Message &
                            "). Replace the kit by hand if the size changed."
            End Try

            ' 4c) put the right DIN 7991 in each plate joint. R_Fasteners has just
            '     recomputed which length each one needs; every stock length is its own
            '     document, so making the assembly agree means exchanging components -
            '     the price of a Part Number that is the real designation.
            '     Runs before the update, like the gussets, so the assembly rebuilds once.
            Dim screwMsg As String = ""
            Try
                If _tableDoc IsNot Nothing Then
                    Dim lp As Double = ReadLen(MasterParameters.P_SCREW_LEN_POST)
                    Dim lu As Double = ReadLen(MasterParameters.P_SCREW_LEN_UPPER)
                    Dim ll As Double = ReadLen(MasterParameters.P_SCREW_LEN_LOWER)
                    If lp > 0 OrElse lu > 0 OrElse ll > 0 Then
                        screwMsg = ScrewSwitcher.EnsureScrewLengths(
                            _app, _tableDoc, lp, lu, ll, AppConfig.CadRoot())
                    End If
                End If
            Catch exS As Exception
                screwMsg = "Screws: swap failed (" & exS.Message &
                           "). Replace them by hand if a length changed."
            End Try

            ' 4c-bis) the same for the four DIN 912 that bolt the robot base through the
            '     worktop. A separate call because it is a separate DOCUMENT: the base and
            '     the plate meet only in main_assembly, so the table assembly - where the
            '     three plate joints live - does not contain these bolts at all.
            '     The washers and nuts under the plate are not touched: their size never
            '     follows the plate, so there is nothing to swap and no constraint worth
            '     risking. That is also why they, unlike the bolts, may stay patterned -
            '     robot_washer.* hangs off the plate, robot_nut.* off the washer.
            Dim mountMsg As String = ""
            Try
                If _mainDoc IsNot Nothing Then
                    Dim lm As Double = ReadLen(MasterParameters.P_SCREW_LEN_ROBOT_MOUNT)
                    If lm > 0 Then
                        mountMsg = ScrewSwitcher.EnsureRobotMountScrews(
                            _app, _mainDoc, lm, AppConfig.CadRoot())
                    End If
                End If
            Catch exM As Exception
                mountMsg = "Robot mount: swap failed (" & exM.Message &
                           "). Replace the four M10 by hand if the worktop thickness changed."
            End Try

            ' 4d) fit the tool the object type asks for. Same reason as the gussets and
            '     the screws: R_Tool has just set the numbers, but exchanging the
            '     adapter plate, gripper, fingers and object is a change of geometry, and
            '     Knowledge cannot replace a component. They travel together in one
            '     sub-assembly so this is a single swap rather than four.
            '     There is nothing else to swap. The two cubes that used to stand on the
            '     stations are gone (30.08.2026): DMU has no attach/detach, so they were
            '     three copies of one object faked with visibility, and they needed their
            '     own switcher because they sat OUTSIDE sub_tool_x. The animation now runs
            '     with the single cube that is inside the tool, which the tool swap already
            '     exchanges - see HOWTO_dmu_kinematics.md section 6.2.
            Dim toolMsg As String = ""
            Try
                Dim rd As Object = RobotDoc()
                If rd IsNot Nothing Then
                    toolMsg = ToolSwitcher.EnsureTool(
                        _app, rd, _design.ObjectType, AppConfig.CadRoot())
                    ' Persist changes and repaired mechanism in robot_mainAssembly to disk
                    Try
                        rd.Save()
                    Catch
                    End Try
                End If
            Catch exT As Exception
                toolMsg = "Tool: swap failed (" & exT.Message &
                          "). Replace sub_tool_* by hand if the object type changed."
            End Try

            ' Clean up imported mechanisms and conflicting Fix constraints on rb_base from main_assembly
            ' so CATIA does not lock constraints or hold dangling references to swapped tools.
            If _mainDoc IsNot Nothing Then
                Try
                    Dim cleanMsg As String = MechanismRepair.CleanupImportedMechanisms(_mainDoc)
                    If cleanMsg <> "" AndAlso toolMsg <> "" Then
                        toolMsg &= vbCrLf & cleanMsg
                    ElseIf cleanMsg <> "" Then
                        toolMsg = cleanMsg
                    End If
                Catch
                End Try
            End If

            msSwap = swAll.ElapsedMilliseconds - msWrite

            ' 5) update the master, then the assembly. Split on purpose: CATIA reports the
            '    same "The method Update failed" for both.
            Try
                _params.UpdatePart()
            Catch exUp As Exception
                Throw New Exception(
                    "The MASTER did not update (master.CATPart)." & vbCrLf & vbCrLf &
                    DiagnosticHint() & vbCrLf & vbCrLf &
                    "CATIA says: " & exUp.Message, exUp)
            End Try

            Try
                ' One pass: a part reached through both the robot assembly and main_assembly
                ' is the same DOCUMENT, and updating it twice is a COM round trip for
                ' nothing. This does NOT skip any of the three documents below - see the
                ' note on the robot assembly for why each one must be updated in its own
                ' right - it only stops the same part being updated again on the way.
                AssemblyUpdater.BeginPass()

                ' Order matters, and it is bottom-up: every document is updated before the
                ' one that nests it.
                '
                ' Table assembly first (holds the profiles/plates/brackets the switchers
                ' just touched), then the ROBOT assembly, then main_assembly on top - it
                ' nests both, so updating it re-solves the robot's mount and moves it.
                If _tableDoc IsNot Nothing Then
                    AssemblyUpdater.UpdateProductDocument(_tableDoc)
                End If

                ' The robot assembly was missing from this list until 29.08.2026, and it is
                ' the one ToolSwitcher writes to: a swap removes the tool instance and adds
                ' a new one, which stays unresolved until its OWN document updates. The
                ' symptom was exact - after Apply the gripper and the cube sat in the wrong
                ' place, and opening robot_mainAssembly and pressing Update by hand fixed
                ' it. Updating main_assembly does not do it: a parent update does not
                ' re-solve a child's constraints.
                If _robotDoc IsNot Nothing Then
                    AssemblyUpdater.UpdateProductDocument(_robotDoc)

                    ' Back to home and re-read the joint geometry. Without it a change of
                    ' segment length leaves the mechanism solving on the old joint axes
                    ' (P1 missed by 57 mm) - see RobotHome.
                    Dim homeMsg As String = RobotHome.Reset(_app, _robotDoc)
                    If homeMsg <> "" Then
                        toolMsg = If(toolMsg = "", homeMsg, toolMsg & vbCrLf & homeMsg)
                    End If
                    ' Saved again: the save after the tool swap (step 4d) comes BEFORE this
                    ' reset, so the file kept the stale link positions and a model reopened
                    ' without saving missed P1 by 40 mm until the next Apply.
                    Try
                        _robotDoc.Save()
                    Catch
                    End Try
                End If

                If _mainDoc IsNot Nothing Then
                    ' The robot must be RIGID in main_assembly, or a change of table height
                    ' moves rb_base alone and leaves the arm behind - see MechanismImport.
                    Dim rigidMsg As String = MechanismImport.EnsureRobotRigid(_app, _mainDoc, "robot_mainAssembly.1")
                    If rigidMsg <> "" Then
                        toolMsg = If(toolMsg = "", rigidMsg, toolMsg & vbCrLf & rigidMsg)
                    End If

                    AssemblyUpdater.UpdateProductDocument(_mainDoc)

                    ' A SECOND pass over main_assembly, and it is not superstition.
                    '
                    ' Reported 30.08.2026: after a change of table size, shelf mounting or
                    ' profile section, the profiles came up red in main_assembly and only a
                    ' manual Update in CATIA cleared them. Every other change - plate
                    ' thickness, robot dimensions, gripper, path points, simplified - went
                    ' through clean. What those three have in common is that they change the
                    ' SHAPE of the profiles or where they sit, which is the geometry the
                    ' assembly's constraints are built on: the first pass re-solves them
                    ' against parts that are still being rebuilt underneath it.
                    '
                    ' So this is literally the Update the user was pressing by hand, and with
                    ' it the last manual step in Apply is gone - which is the point of the
                    ' whole application.
                    '
                    ' Deliberately an EXTRA pass, not a rearranged one. Yesterday a pass was
                    ' REMOVED here as an optimisation, the rebuild came out wrong and it had
                    ' to be reverted. In this chain the number of passes is measured, not
                    ' shortened from first principles - and the per-phase timing in the
                    ' result dialog is what measures it.
                    AssemblyUpdater.UpdateProductDocument(_mainDoc)

                    ' The imported mechanisms were deleted before the update (they go
                    ' DOF=Undefined once the joint geometry moves); bring them back now that
                    ' the geometry is final - see MechanismImport.
                    Dim impMsg As String = MechanismImport.Reimport(_app, _mainDoc)
                    If impMsg <> "" Then
                        toolMsg = If(toolMsg = "", impMsg, toolMsg & vbCrLf & impMsg)
                    End If

                    CatiaConnector.Activate(_mainDoc)   ' bring the full view to the front
                ElseIf _tableDoc IsNot Nothing Then
                    CatiaConnector.Activate(_tableDoc)
                Else
                    AssemblyUpdater.UpdateActive(_app)
                End If
            Catch exAsm As Exception
                Throw New Exception(
                    "The master went through, the ASSEMBLY did not update." & vbCrLf & vbCrLf &
                    DiagnosticHint() & vbCrLf & vbCrLf &
                    "CATIA says: " & exAsm.Message, exAsm)
            End Try

            msUpdate = swAll.ElapsedMilliseconds - msWrite - msSwap

            Dim msg As String = "Parameters written, assembly rebuilt.  Section " & sectionUsed & "." &
                                vbCrLf & "  " & swAll.ElapsedMilliseconds & " ms  (parameters " &
                                msWrite & ", swaps " & msSwap & ", update " & msUpdate & ")"
            ' The switchers are chatty (every bracket, every screw, the constraint path).
            ' On a clean run that is noise - so only a switcher in trouble is surfaced.
            '
            ' It used to test for the word "fail", which almost nothing says. ToolSwitcher
            ' words its refusals "Tool: no instance named tool_1...", "Tool NOT swapped: ...
            ' does not exist", "Tool: nothing to swap" - not one of them matched, so every
            ' one of them was swallowed and the dialog looked exactly like a clean run.
            ' Found 29.08.2026, when a silent Apply could not be told from a working swap.
            For Each sw As String In New String() {gussetMsg, screwMsg, mountMsg, toolMsg}
                If IsSwitcherTrouble(sw) Then msg &= vbCrLf & vbCrLf & sw
            Next

            ' The tool line is shown even when it is good news. Which gripper is fitted is
            ' the one thing about a run that the model does not make obvious at a glance,
            ' and it is a single line.
            If toolMsg <> "" AndAlso Not IsSwitcherTrouble(toolMsg) Then
                msg &= vbCrLf & vbCrLf & toolMsg
            End If
            ' The switcher falls back to the full section instead of failing. Say so, or
            ' the simplified view silently looks exactly like the detailed one - which is
            ' precisely how this went unnoticed for a whole evening.
            If _design.Simplified AndAlso sectionUsed = _design.ProfileName Then
                msg &= vbCrLf & vbCrLf &
                       "Simplified was requested, but '" & ProfileSwitcher.SHARED_SIMPLE &
                       "' is not in the master, so the full section is still in use. It has " &
                       "to be a surface (Fill over the sketch), not the sketch itself."
            End If
            If skipped.Length > 0 Then
                msg &= vbCrLf & vbCrLf & "Skipped, not in the model:" & vbCrLf & skipped.ToString()
            End If
            If locked.Length > 0 Then
                msg &= vbCrLf & vbCrLf &
                       "NOT WRITTEN - these are read-only in the model (a formula or rule owns" & vbCrLf &
                       "them). They must be free inputs. Remove the formula in Tools -> Formula:" & vbCrLf &
                       locked.ToString()
            End If
            SetStatus("Update finished.", True)
            _dirty = False
            RefreshEverything()
            MessageBox.Show(msg, "CATIA", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            SetStatus("Apply failed.", False)
            MessageBox.Show("Applying failed:" & vbCrLf & vbCrLf & ex.Message,
                            "CATIA", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    ' Does a switcher's message report something the user has to see?
    '
    ' The switchers return free text, so this reads their vocabulary rather than a flag.
    ' Every phrase below is one a switcher actually produces when it declines or fails; a
    ' clean run says "Tool: sub_tool_1 -> sub_tool_2 (...)" or "Tool: already sub_tool_1"
    ' and matches none of them. Keep this list in step with the switchers' wording.
    Private Shared ReadOnly TroubleMarks As String() = {
        "fail", "not swapped", "no instance", "nothing to swap",
        "does not exist", "by hand", "could not"}

    Private Function IsSwitcherTrouble(ByVal msg As String) As Boolean
        If String.IsNullOrEmpty(msg) Then Return False
        For Each mark As String In TroubleMarks
            If msg.IndexOf(mark, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        Next
        Return False
    End Function

    ' A length the MODEL owns, read back in mm. 0 when the parameter is not there yet -
    ' the app has to stay usable against an older master.
    Private Function ReadLen(ByVal paramName As String) As Double
        Try
            Dim raw As String = _params.ReadNumber(paramName)
            If raw = "" Then Return 0
            Dim d As Double
            If Double.TryParse(Num.Normalize(raw), Globalization.NumberStyles.Float,
                               Globalization.CultureInfo.InvariantCulture, d) Then
                Return d
            End If
        Catch
        End Try
        Return 0
    End Function

    Private Sub TryOptional(ByVal paramName As String, ByVal value As Double,
                            ByVal skipped As StringBuilder, ByVal locked As StringBuilder)
        If Not _params.TryWriteLengthReport(paramName, DesignLimits.Fmt(value), locked) Then
            skipped.AppendLine("  - " & paramName)
        End If
    End Sub

    ' "The method Update failed" never says WHICH feature broke - that lives in CATIA's
    ' own Update Diagnosis window. Print the values that used to cause it and the way to
    ' the real diagnosis.
    Private Function DiagnosticHint() As String
        Dim sb As New StringBuilder()
        sb.AppendLine("CATIA has the real diagnosis, this app does not:")
        sb.AppendLine("  1. switch to CATIA (the values are already written)")
        sb.AppendLine("  2. Edit -> Update (or Ctrl+U)")
        sb.AppendLine("  3. Update Diagnosis opens and NAMES the feature that broke")
        sb.AppendLine()
        sb.AppendLine("Plate envelope: X +-" & DesignLimits.Fmt(_design.Length / 2) &
                      "   Y +-" & DesignLimits.Fmt(_design.Width / 2))
        sb.AppendLine("Robot base:     outer hole " & DesignLimits.Fmt(_design.AxisOffsetX) & " / " &
                      DesignLimits.Fmt(_design.AxisOffsetY) &
                      "   pitch " & DesignLimits.Fmt(_design.MountPitchX) & " / " &
                      DesignLimits.Fmt(_design.MountPitchY))
        Return sb.ToString()
    End Function

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        Try
            OleMessageFilter.Revoke()
        Catch
        End Try
        MyBase.OnFormClosed(e)
    End Sub

End Class
