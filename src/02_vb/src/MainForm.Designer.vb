Option Explicit On
Option Strict Off

Imports System.Windows.Forms
Imports System.Drawing

' Layout only. Everything that talks to CATIA lives in MainForm.vb.
'
' Naming rule for this window: the user reads plain English ("Distance from the front
' edge"), the model keeps its parameter name ("user_offset_x"). The parameter name is
' shown as a quiet hint next to the input so the report and the CATIA tree still line up.
Partial Class MainForm

    Private components As System.ComponentModel.IContainer = Nothing

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    ' --- header / footer ---
    Friend WithEvents btnConnect As Button
    Friend WithEvents btnApply As Button
    Friend WithEvents btnReload As Button
    ' No btnSimulate (removed 30.08.2026, user's decision). The simulation belongs to the
    ' model: the Knowledge laws drive the mechanism from KINTime, and it is played in DMU
    ' where it is also recorded. A second channel that pushed command values from here was
    ' a way of moving the arm WITHOUT the laws - two drivers for one mechanism, of which
    ' only one can be right, and the button had to switch the laws off to work at all.
    Friend WithEvents btnClose As Button
    Private lblStatus As Label
    Private lblApplyNote As Label

    ' --- table inputs ---
    Friend WithEvents sldLength As ParameterSlider
    Friend WithEvents sldWidth As ParameterSlider
    Friend WithEvents sldHeight As ParameterSlider
    Friend WithEvents sldUpperPlate As ParameterSlider
    Friend WithEvents sldLowerPlate As ParameterSlider
    Friend WithEvents optProfile As OptionPicker
    Friend WithEvents optSimplification As OptionPicker
    Friend WithEvents optShelfVariant As OptionPicker

    ' --- robot inputs ---
    Friend WithEvents optRobotMount As OptionPicker
    Friend WithEvents sldOffsetX As ParameterSlider
    Friend WithEvents sldOffsetY As ParameterSlider
    Friend WithEvents sldBaseOverhang As ParameterSlider
    Friend WithEvents sldJ2J3 As ParameterSlider
    Friend WithEvents sldJ3J5 As ParameterSlider

    ' --- path inputs ---
    Friend WithEvents sldP1X As ParameterSlider
    Friend WithEvents sldP1Y As ParameterSlider
    Friend WithEvents sldP1Z As ParameterSlider
    Friend WithEvents sldP3X As ParameterSlider
    Friend WithEvents sldP3Y As ParameterSlider
    Friend WithEvents sldP3Z As ParameterSlider
    Friend WithEvents sldP6X As ParameterSlider
    Friend WithEvents sldP6Y As ParameterSlider
    Friend WithEvents sldP6Z As ParameterSlider
    Friend WithEvents sldP2P3 As ParameterSlider
    Friend WithEvents sldP2P4 As ParameterSlider
    ' No sldP4P5: that descent is P2P3 by construction (both stations at one height).
    ' No sldPathRadius / sldStationHeight: the stations are no longer placed on a circle,
    ' they follow from P3 and the two segments.
    Friend WithEvents sldSimLeg As ParameterSlider
    Friend WithEvents sldSimSpeed As ParameterSlider

    ' --- tool inputs ---
    Friend WithEvents optObjectType As OptionPicker

    ' --- right column ---
    Private sketch As SketchPanel
    Private flowChecks As FlowLayoutPanel
    Private lblRobotRaster As Label
    Private lblRobotHole As Label
    Private lblRobotReach As Label
    Private lblRobotArm As Label
    Private lblRobotSource As Label
    Private lblGussetSize As Label
    Private lblGussetCornerSize As Label
    Private lblLowerPlateLen As Label
    Private lblLowerPlateWid As Label
    Private lblGussetModel As Label
    Private lblGussetCornerModel As Label
    Private lblLowerLenModel As Label
    Private lblLowerWidModel As Label
    Private lblAxisOffset As Label
    Private lblAxisOffsetModel As Label
    Private lblMountPitch As Label
    Private lblMountPitchModel As Label
    Private lblMountEdge As Label
    Private lblMountEdgeModel As Label

    Private flowTable As FlowLayoutPanel
    Private flowRobot As FlowLayoutPanel
    Private flowPath As FlowLayoutPanel
    Private flowTool As FlowLayoutPanel
    Friend WithEvents tabs As TabControl

    Private Const ColW As Integer = 470

    Private Sub InitializeComponent()
        Me.SuspendLayout()

        Me.AutoScaleDimensions = New SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = AutoScaleMode.Font
        Me.ClientSize = New Size(1240, 840)
        Me.MinimumSize = New Size(1120, 760)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Theme.Background
        Me.Font = Theme.Body
        Me.Text = "Robot table configurator - CATIA"

        ' ---------------- left column: the inputs ----------------
        Dim pnlLeft As New Panel()
        pnlLeft.Dock = DockStyle.Fill
        pnlLeft.Padding = New Padding(14, 10, 8, 10)
        pnlLeft.BackColor = Theme.Background

        tabs = New TabControl()
        tabs.Dock = DockStyle.Fill
        tabs.Font = Theme.Body
        tabs.Padding = New Point(16, 6)

        flowTable = NewFlow()
        flowRobot = NewFlow()
        flowPath = NewFlow()
        flowTool = NewFlow()
        tabs.TabPages.Add(NewTab("Table", flowTable))
        tabs.TabPages.Add(NewTab("Robot", flowRobot))
        tabs.TabPages.Add(NewTab("Tool", flowTool))
        ' No longer "(optional)": the path parameters exist in the master from
        ' 09_robot_parameters onwards, and the simulation is driven from them.
        tabs.TabPages.Add(NewTab("Path", flowPath))
        pnlLeft.Controls.Add(tabs)

        ' --- Table tab ---
        flowTable.Controls.Add(Theme.SectionHeader("Table size",
            "Outer dimensions of the frame. The limits follow the chosen profile.", ColW))

        sldLength = New ParameterSlider("Table length", "Front to back, along X   (model: length)", "mm", 0, "length")
        sldWidth = New ParameterSlider("Table width", "Left to right, along Y   (model: width)", "mm", 0, "width")
        sldHeight = New ParameterSlider("Table height", "Floor to the top of the worktop   (model: height)", "mm", 0, "height")
        flowTable.Controls.AddRange(New Control() {sldLength, sldWidth, sldHeight})

        flowTable.Controls.Add(Theme.SectionHeader("Plates",
            "Steel worktop on top, shelf plate on the lower rails.", ColW))
        sldUpperPlate = New ParameterSlider("Worktop thickness",
            "Steel plate the robot is bolted to   (model: upper_plate_thickness)", "mm", 1, "upperPlate")
        sldLowerPlate = New ParameterSlider("Shelf thickness",
            "Lower plate, the counterweight   (model: lower_plate_thickness)", "mm", 1, "lowerPlate")
        flowTable.Controls.AddRange(New Control() {sldUpperPlate, sldLowerPlate})

        flowTable.Controls.Add(Theme.SectionHeader("Aluminium profile",
            "Section of every post and rail. It drives the bracket size and the minimum table size.", ColW))
        optProfile = New OptionPicker("Profile section", "profile", ProfileChoices())
        optSimplification = New OptionPicker("Level of detail", "simplification", New ChoiceItem() {
            New ChoiceItem("Full section - slots and bores modelled", "false",
                           "Realistic profile geometry. Heavier to rebuild."),
            New ChoiceItem("Simplified - solid block of the same envelope", "true",
                           "Section 5 of the task. Outer envelope and all constraints stay identical.")})
        flowTable.Controls.AddRange(New Control() {optProfile, optSimplification})

        flowTable.Controls.Add(Theme.SectionHeader("Shelf mounting",
            "Which side of the lower rails carries the brackets.", ColW))
        optShelfVariant = New OptionPicker("Bracket position", "variant", New ChoiceItem() {
            New ChoiceItem("V1 - brackets under the rails", "v1",
                           "The lower level climbs by one bracket size; the shelf keeps the full span."),
            New ChoiceItem("V2 - brackets on top of the rails", "v2",
                           "Shelf is flush with the post ends but loses one bracket per side.")})
        flowTable.Controls.Add(optShelfVariant)

        ' --- Robot tab ---
        flowRobot.Controls.Add(Theme.SectionHeader("Robot",
            "The robot the task assigns, and the only one modelled. Its base bolt raster " &
            "(205 x 160, M10 clearance) is a fixed property of that robot - the app writes it " &
            "but does not offer it as a control; what stays adjustable is where and how the " &
            "base sits on the worktop, below.", ColW))
        flowRobot.Controls.Add(BuildRobotSpecPanel())

        flowRobot.Controls.Add(Theme.SectionHeader("Robot base position",
            "Where the four base bolts land on the worktop. Both distances are measured from the edge.", ColW))
        optRobotMount = New OptionPicker("Standing position", "robotMount", MountChoices())
        sldOffsetX = New ParameterSlider("Distance from the +X edge",
            "Right-hand edge in the top view to the nearest bolt; past half the length the base crosses the middle   (model: user_offset_x)",
            "mm", 0, "offsetX")
        sldOffsetY = New ParameterSlider("Distance from the +Y edge",
            "Top edge in the top view to the nearest bolt; past half the width the base crosses the middle   (model: user_offset_y)",
            "mm", 0, "offsetY")
        sldBaseOverhang = New ParameterSlider("Robot base overhang past the bolts",
            "Half of (base size - bolt pitch). Sets the minimum distances here; not written to the model",
            "mm", 0, "baseOverhang")
        flowRobot.Controls.AddRange(New Control() {optRobotMount, sldOffsetX, sldOffsetY, sldBaseOverhang})

        flowRobot.Controls.Add(Theme.SectionHeader("Arm segments",
            "Optional. Written only when the master already has these parameters.", ColW))
        sldJ2J3 = New ParameterSlider("Upper arm, J2 to J3",
            "Nominal RoboDK length up to +20 %   (model: len_J2J3)", "mm", 0, "lenJ2J3")
        sldJ3J5 = New ParameterSlider("Forearm, J3 to J5",
            "Nominal RoboDK length up to +20 %   (model: len_J3J5)", "mm", 0, "lenJ3J5")
        flowRobot.Controls.AddRange(New Control() {sldJ2J3, sldJ3J5})

        ' --- Path tab ---
        ' All coordinates here are in the ROBOT BASE FRAME: origin on the robot's
        ' mounting face, X forward, Z up. Not the table frame and not the floor - the
        ' inverse kinematics works in the robot's own frame, and converting twice is how
        ' a path ends up 450 mm below where it was meant to be.
        ' Two drawings, not one, and which one appears follows the input under the cursor:
        ' forward (X) and sideways (Y) bring up the TOP view, height (Z) the FRONT view -
        ' each slider carries its own sketch key for that. The two can also be swapped by hand - SketchPanel draws a Top/Front switch
        ' on both path drawings.
        flowPath.Controls.Add(Theme.SectionHeader("Start point P1",
            "Free by the task. Where the arm begins, before it goes for the object.", ColW))
        sldP1X = New ParameterSlider("P1 forward", "Along the robot's X axis, negative is behind it   (model: P1_x)", "mm", 0, "p1x")
        sldP1Y = New ParameterSlider("P1 sideways", "Across the robot's own X axis   (model: P1_y)", "mm", 0, "p1y")
        sldP1Z = New ParameterSlider("P1 height", "Above the robot's mounting face   (model: P1_z)", "mm", 0, "p1")
        flowPath.Controls.AddRange(New Control() {sldP1X, sldP1Y, sldP1Z})

        ' P3 is the one point that is also a place in the model: the cube is initialised
        ' with its CENTRE here, because that is the point the tool grips.
        flowPath.Controls.Add(Theme.SectionHeader("Pick station P3",
            "Where the object starts, and where the tool goes to take it. " &
            "P2 sits above it, and P4/P5 follow from the two segments below.", ColW))
        sldP3X = New ParameterSlider("P3 forward", "Along the robot's X axis, negative is behind it   (model: P3_x)", "mm", 0, "p3x")
        sldP3Y = New ParameterSlider("P3 sideways", "The place station is this plus P2 to P4   (model: P3_y)", "mm", 0, "p3y")
        sldP3Z = New ParameterSlider("P3 height", "Centre of the object above the mounting face   (model: P3_z)", "mm", 0, "p3")
        flowPath.Controls.AddRange(New Control() {sldP3X, sldP3Y, sldP3Z})

        flowPath.Controls.Add(Theme.SectionHeader("End point P6",
            "Free by the task. Where the arm parks after it has let go.", ColW))
        sldP6X = New ParameterSlider("P6 forward", "Along the robot's X axis, negative is behind it   (model: P6_x)", "mm", 0, "p6x")
        sldP6Y = New ParameterSlider("P6 sideways", "Across the robot's own X axis   (model: P6_y)", "mm", 0, "p6y")
        sldP6Z = New ParameterSlider("P6 height", "Above the robot's mounting face   (model: P6_z)", "mm", 0, "p6")
        flowPath.Controls.AddRange(New Control() {sldP6X, sldP6Y, sldP6Z})

        ' No P4 to P5 slider: the task figure puts both stations at one height, so that
        ' descent equals P2 to P3. An independent value would float one station above the
        ' other, which the figure does not show.
        flowPath.Controls.Add(Theme.SectionHeader("Path segments",
            "The moves the simulation drives. Task section 12 asks for these to be " &
            "changeable; they start at the task's own 4a and c for the object fitted.", ColW))
        sldP2P3 = New ParameterSlider("P2 to P3", "Down onto the object, task value c   (model: P2P3)", "mm", 0, "p2p3")
        sldP2P4 = New ParameterSlider("P2 to P4", "Across to the other station, task value 4a   (model: P2P4)", "mm", 0, "p2p4")
        flowPath.Controls.AddRange(New Control() {sldP2P3, sldP2P4})

        flowPath.Controls.Add(Theme.SectionHeader("Speed",
            "The tool goes straight between P2 and P5 and swings through free space on " &
            "P1 to P2 and P4 to P6 - so the two have their own speeds, as on a real robot.", ColW))
        sldSimLeg = New ParameterSlider("Joint move", "Seconds for each free-space swing   (model: sim_leg)", "s", 1, "simLeg")
        sldSimSpeed = New ParameterSlider("Straight move", "Tool speed on the straight legs   (model: sim_speed)", "mm/s", 0, "simSpeed")
        flowPath.Controls.AddRange(New Control() {sldSimLeg, sldSimSpeed})

        ' --- Tool tab ---
        flowTool.Controls.Add(Theme.SectionHeader("Object and gripper",
            "Task section 9: one choice reconfigures the adapter plate, the gripper, the " &
            "fingers and the object. The rule sets the numbers; this program exchanges " &
            "the sub-assembly, because Knowledge cannot replace a component.", ColW))
        optObjectType = New OptionPicker("Object", "objectType", New ChoiceItem() {
            New ChoiceItem("Object 1 - 40 x 50 x 60, SCHUNK MPG-plus 40", "1",
                           "Stroke 6 mm per jaw; fingers up to 50 mm. Fitted as sub_tool_1."),
            New ChoiceItem("Object 2 - 80 x 80 x 80, SCHUNK EGP 64-N-N-IOL", "2",
                           "Stroke 10 mm per jaw; fingers up to 80 mm. Fitted as sub_tool_2.")})
        flowTool.Controls.Add(optObjectType)

        ' NO "Adapter plate" SLIDER (removed 29.08.2026). It offered adapter_t as an input
        ' driving the plate thickness and, through it, tool_len. That direction is reversed:
        ' the plates are shaped freely and the tool length is measured off the finished
        ' assembly, FLANGE_AXS to TCP_AXS. A slider that changes nothing in the model is
        ' worse than no slider - see 05_robot/HOWTO_dmu_kinematics.md section 12.

        ' ---------------- right column: sketch, derived values, checks ----------------
        Dim pnlRight As New Panel()
        pnlRight.Dock = DockStyle.Right
        pnlRight.Width = 520
        pnlRight.Padding = New Padding(6, 10, 14, 10)
        pnlRight.BackColor = Theme.Background

        sketch = New SketchPanel()
        sketch.Dock = DockStyle.Fill

        Dim pnlDerived As Panel = BuildDerivedPanel()
        Dim pnlChecks As Panel = BuildChecksPanel()

        pnlRight.Controls.Add(sketch)
        pnlRight.Controls.Add(pnlDerived)
        pnlRight.Controls.Add(pnlChecks)

        ' ---------------- header ----------------
        ' Width is set before the children so the Right anchors start from the real edge.
        Dim pnlHeader As New Panel()
        pnlHeader.Size = New Size(Me.ClientSize.Width, 76)
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.BackColor = Theme.Card

        Dim lblTitle As New Label()
        lblTitle.AutoSize = True
        lblTitle.Font = Theme.Title
        lblTitle.ForeColor = Theme.Ink
        lblTitle.Location = New Point(18, 14)
        lblTitle.Text = "Robot table configurator"

        Dim lblSub As New Label()
        lblSub.AutoSize = True
        lblSub.Font = Theme.Small
        lblSub.ForeColor = Theme.Muted
        lblSub.Location = New Point(20, 42)
        lblSub.Text = "Sets the master parameters in CATIA and rebuilds the table assembly."

        btnConnect = Theme.PrimaryButton("Connect to CATIA", 190, 34)
        btnConnect.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnConnect.Location = New Point(pnlHeader.Width - 210, 12)

        lblStatus = New Label()
        lblStatus.AutoSize = False
        lblStatus.TextAlign = ContentAlignment.MiddleRight
        lblStatus.Font = Theme.Small
        lblStatus.ForeColor = Theme.Muted
        lblStatus.Size = New Size(480, 18)
        lblStatus.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        lblStatus.Location = New Point(pnlHeader.Width - 500, 50)
        lblStatus.Text = "Not connected - configure freely, connecting is only needed to apply."

        Dim rule As New Panel()
        rule.Dock = DockStyle.Bottom
        rule.Height = 1
        rule.BackColor = Theme.Border

        pnlHeader.Controls.AddRange(New Control() {lblTitle, lblSub, btnConnect, lblStatus, rule})

        ' ---------------- footer ----------------
        Dim pnlFooter As New Panel()
        pnlFooter.Size = New Size(Me.ClientSize.Width, 66)
        pnlFooter.Dock = DockStyle.Bottom
        pnlFooter.BackColor = Theme.Card

        btnApply = Theme.PrimaryButton("Apply and rebuild", 200, 38)
        btnApply.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnApply.Location = New Point(pnlFooter.Width - 220, 14)

        btnReload = Theme.QuietButton("Reload from model", 160, 38)
        btnReload.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnReload.Location = New Point(pnlFooter.Width - 390, 14)

        btnClose = Theme.QuietButton("Close", 92, 38)
        btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnClose.Location = New Point(pnlFooter.Width - 502, 14)

        lblApplyNote = New Label()
        lblApplyNote.AutoSize = False
        lblApplyNote.Font = Theme.Small
        lblApplyNote.ForeColor = Theme.Muted
        lblApplyNote.Location = New Point(18, 24)
        lblApplyNote.Size = New Size(560, 30)
        lblApplyNote.Text = ""

        Dim rule2 As New Panel()
        rule2.Dock = DockStyle.Top
        rule2.Height = 1
        rule2.BackColor = Theme.Border

        pnlFooter.Controls.AddRange(New Control() {lblApplyNote, btnApply, btnReload,
                                                   btnClose, rule2})

        ' Fill first, then the edges - a docked control added later is placed further out.
        Me.Controls.Add(pnlLeft)
        Me.Controls.Add(pnlRight)
        Me.Controls.Add(pnlFooter)
        Me.Controls.Add(pnlHeader)

        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    ' --- small builders ---

    Private Function NewFlow() As FlowLayoutPanel
        Dim f As New FlowLayoutPanel()
        f.Dock = DockStyle.Fill
        f.FlowDirection = FlowDirection.TopDown
        f.WrapContents = False
        f.AutoScroll = True
        f.Padding = New Padding(16, 8, 16, 16)
        f.BackColor = Theme.Card
        AddHandler f.Resize, AddressOf OnFlowResize
        Return f
    End Function

    Private Sub OnFlowResize(ByVal sender As Object, ByVal e As EventArgs)
        Theme.StretchChildren(DirectCast(sender, FlowLayoutPanel))
    End Sub

    Private Function NewTab(ByVal caption As String, ByVal body As Control) As TabPage
        Dim t As New TabPage(caption)
        t.BackColor = Theme.Card
        t.Padding = New Padding(0)
        t.Controls.Add(body)
        Return t
    End Function

    ' The small robot readout: what the assigned robot's own data says, next to where those
    ' numbers came from. The rows are filled in by ShowRobotSpec.
    Private Function BuildRobotSpecPanel() As Panel
        Dim card As New CardBox()
        card.Size = New Size(ColW, 148)
        card.Margin = New Padding(0, 2, 0, 10)
        card.Padding = New Padding(12, 8, 12, 8)

        Dim colName As Integer = 12, colVal As Integer = 210

        Dim head As New Label()
        head.AutoSize = True
        head.Font = Theme.BodyBold
        head.ForeColor = Theme.Ink
        head.Location = New Point(colName, 8)
        head.Text = RobotCatalog.Assigned.Caption

        lblRobotRaster = Value(colVal, 30)
        lblRobotHole = Value(colVal, 48)
        lblRobotReach = Value(colVal, 66)
        lblRobotArm = Value(colVal, 84)

        lblRobotSource = New Label()
        lblRobotSource.AutoSize = False
        lblRobotSource.Size = New Size(ColW - 24, 44)
        lblRobotSource.Location = New Point(colName, 104)
        lblRobotSource.Font = Theme.Small
        lblRobotSource.ForeColor = Theme.Muted

        card.Controls.AddRange(New Control() {head,
            Caption("Bolt raster  long / short", colName, 30), lblRobotRaster,
            Caption("Hole / bolt", colName, 48), lblRobotHole,
            Caption("Reach / payload", colName, 66), lblRobotReach,
            Caption("Arm J2-J3 / J3-J5", colName, 84), lblRobotArm,
            lblRobotSource})
        Return card
    End Function

    Private Function ProfileChoices() As ChoiceItem()
        Dim list As New List(Of ChoiceItem)()
        For Each p As ProfileInfo In ProfileCatalog.All
            list.Add(New ChoiceItem(p.Caption, p.Name, p.Hint))
        Next
        Return list.ToArray()
    End Function

    ' Position and orientation are separate questions, so every position offers both
    ' orientations. "Long side" is the 100 mm side of the bolt raster.
    Private Function MountChoices() As ChoiceItem()
        Return New ChoiceItem() {
            New ChoiceItem("Centred  -  long side along the length", "centerLongAxis",
                           "Raster centred on the plate, long side along X. Neither distance below is used."),
            New ChoiceItem("Centred  -  long side across the width", "centerShortAxis",
                           "Raster centred on the plate, long side along Y. Neither distance below is used."),
            New ChoiceItem("Free along the length, centred across", "edgeLongAxis",
                           "Only the distance from the +X edge places the base; it stays centred across the width."),
            New ChoiceItem("Free across the width, centred along", "edgeShortAxis",
                           "Only the distance from the +Y edge places the base; it stays centred along the length."),
            New ChoiceItem("Free anywhere  -  long side along the length", "cornerLongAxis",
                           "Both distances place the base. It reaches the whole plate, not just a corner."),
            New ChoiceItem("Free anywhere  -  long side across the width", "cornerShortAxis",
                           "Both distances place the base, with the long side turned across the width.")}
    End Function

    Private Function BuildDerivedPanel() As Panel
        Dim card As New CardBox()
        card.Size = New Size(500, 182)
        card.Dock = DockStyle.Bottom
        card.Padding = New Padding(14, 10, 14, 10)

        Dim head As New Label()
        head.AutoSize = True
        head.Font = Theme.BodyBold
        head.ForeColor = Theme.Ink
        head.Location = New Point(14, 8)
        head.Text = "Sizes the model works out for itself"

        Dim colName As Integer = 14, colPreview As Integer = 250, colModel As Integer = 380
        Dim hdr1 As Label = Caption("Preview", colPreview, 28)
        Dim hdr2 As Label = Caption("In CATIA", colModel, 28)

        lblGussetSize = Value(colPreview, 46)
        lblGussetModel = Value(colModel, 46)
        ' The second bracket size. It has its own row rather than sharing one, because the
        ' whole point is that on 45x60 and 45x90 the two numbers disagree.
        lblGussetCornerSize = Value(colPreview, 64)
        lblGussetCornerModel = Value(colModel, 64)
        lblLowerPlateLen = Value(colPreview, 82)
        lblLowerLenModel = Value(colModel, 82)
        lblLowerPlateWid = Value(colPreview, 100)
        lblLowerWidModel = Value(colModel, 100)
        ' The two rows that make a wrong robot base visible: what the app expects against
        ' what the rule in the master actually produced.
        lblAxisOffset = Value(colPreview, 122)
        lblAxisOffsetModel = Value(colModel, 122)
        lblMountPitch = Value(colPreview, 140)
        lblMountPitchModel = Value(colModel, 140)
        lblMountEdge = Value(colPreview, 158)
        lblMountEdgeModel = Value(colModel, 158)

        card.Controls.AddRange(New Control() {head, hdr1, hdr2,
            Caption("Bracket size  -  post to rail  (16x)", colName, 46), lblGussetSize, lblGussetModel,
            Caption("Bracket size  -  rail to rail  (8x)", colName, 64), lblGussetCornerSize, lblGussetCornerModel,
            Caption("Shelf plate length", colName, 82), lblLowerPlateLen, lblLowerLenModel,
            Caption("Shelf plate width", colName, 100), lblLowerPlateWid, lblLowerWidModel,
            Caption("First robot bolt, from centre  X / Y", colName, 122), lblAxisOffset, lblAxisOffsetModel,
            Caption("Robot bolt pitch  X / Y", colName, 140), lblMountPitch, lblMountPitchModel,
            Caption("First robot bolt, from edge  X / Y", colName, 158), lblMountEdge, lblMountEdgeModel})
        Return card
    End Function

    Private Function Caption(ByVal text As String, ByVal x As Integer, ByVal y As Integer) As Label
        Dim l As New Label()
        l.AutoSize = True
        l.Font = Theme.Small
        l.ForeColor = Theme.Muted
        l.Location = New Point(x, y)
        l.Text = text
        Return l
    End Function

    Private Function Value(ByVal x As Integer, ByVal y As Integer) As Label
        Dim l As New Label()
        l.AutoSize = True
        l.Font = Theme.Small
        l.ForeColor = Theme.Ink
        l.Location = New Point(x, y)
        l.Text = "-"
        Return l
    End Function

    Private Function BuildChecksPanel() As Panel
        Dim card As New CardBox()
        card.Size = New Size(500, 178)
        card.Dock = DockStyle.Bottom
        card.Padding = New Padding(14, 10, 14, 10)

        Dim head As New Label()
        head.AutoSize = True
        head.Font = Theme.BodyBold
        head.ForeColor = Theme.Ink
        head.Location = New Point(14, 8)
        head.Text = "Checks"

        flowChecks = New FlowLayoutPanel()
        flowChecks.Location = New Point(12, 30)
        flowChecks.Size = New Size(card.Width - 26, card.Height - 42)
        flowChecks.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        flowChecks.FlowDirection = FlowDirection.TopDown
        flowChecks.WrapContents = False
        flowChecks.AutoScroll = True
        flowChecks.BackColor = Theme.Card

        card.Controls.Add(head)
        card.Controls.Add(flowChecks)
        Return card
    End Function

End Class
