Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Path where commands are stored (JSON)
$commandsPath = Join-Path $env:USERPROFILE 'PowerShellCommandGuiCommands.json'

# Global list of commands
$script:Commands = @()

# Load existing commands or keep empty
if (Test-Path $commandsPath) {
    $loaded = Get-Content $commandsPath -Raw | ConvertFrom-Json
    if ($loaded -is [System.Array]) {
        $script:Commands = $loaded
    } elseif ($null -ne $loaded) {
        $script:Commands = @($loaded)
    }
}

function Save-Commands {
    $script:Commands | ConvertTo-Json -Depth 5 | Set-Content -Path $commandsPath -Encoding UTF8
}

function Refresh-CommandList {
    $listCommands.Items.Clear()
    foreach ($cmd in $script:Commands) {
        [void]$listCommands.Items.Add($cmd.Name)
    }
}

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell Command Launcher"
$form.Size = New-Object System.Drawing.Size(900, 800)
$form.StartPosition = "CenterScreen"
$form.WindowState = "Maximized"
$form.MaximizeBox = $true
$form.MinimizeBox = $true

# List of commands
$listCommands = New-Object System.Windows.Forms.ListBox
$listCommands.Location = New-Object System.Drawing.Point(10, 10)
$listCommands.Size = New-Object System.Drawing.Size(250, 630)
$listCommands.AllowDrop = $true
$listCommands.Anchor = "Top, Bottom, Left"

# Index of the item currently being dragged
$script:DragIndex = -1

$form.Controls.Add($listCommands)

# Command name
$lblName = New-Object System.Windows.Forms.Label
$lblName.Text = "Command name:"
$lblName.Location = New-Object System.Drawing.Point(270, 10)
$lblName.AutoSize = $true
$form.Controls.Add($lblName)

$txtName = New-Object System.Windows.Forms.TextBox
$txtName.Location = New-Object System.Drawing.Point(270, 30)
$txtName.Size = New-Object System.Drawing.Size(300, 20)
$txtName.Anchor = "Top, Left, Right"
$form.Controls.Add($txtName)

# Command description
$lblDescription = New-Object System.Windows.Forms.Label
$lblDescription.Text = "Description (what this command does):"
$lblDescription.Location = New-Object System.Drawing.Point(270, 60)
$lblDescription.AutoSize = $true
$form.Controls.Add($lblDescription)

$txtDescription = New-Object System.Windows.Forms.TextBox
$txtDescription.Location = New-Object System.Drawing.Point(270, 80)
$txtDescription.Size = New-Object System.Drawing.Size(600, 50)
$txtDescription.Multiline = $true
$txtDescription.ScrollBars = "Vertical"
$txtDescription.Anchor = "Top, Left, Right"
$form.Controls.Add($txtDescription)

# Command script
$lblScript = New-Object System.Windows.Forms.Label
$lblScript.Text = "Command template (use $ to mark editable variables):"
$lblScript.Location = New-Object System.Drawing.Point(270, 140)
$lblScript.AutoSize = $true
$form.Controls.Add($lblScript)

$txtScript = New-Object System.Windows.Forms.TextBox
$txtScript.Location = New-Object System.Drawing.Point(270, 160)
$txtScript.Size = New-Object System.Drawing.Size(600, 200)
$txtScript.Multiline = $true
$txtScript.ScrollBars = "Vertical"
$txtScript.Anchor = "Top, Left, Right"
$form.Controls.Add($txtScript)

# Panel where variable controls will be created dynamically
$pnlVariables = New-Object System.Windows.Forms.Panel
$pnlVariables.Location = New-Object System.Drawing.Point(270, 370)
$pnlVariables.Size = New-Object System.Drawing.Size(600, 160)
$pnlVariables.AutoScroll = $true
$pnlVariables.Anchor = "Top, Left, Right"
$form.Controls.Add($pnlVariables)

# Hashtable that will hold "variable name -> TextBox"
$script:VarTextBoxes = @{}

# Current running job (if any)
$script:CurrentJob = $null

# Timer to poll job output
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500  # ms

# When panel resizes, adjust textbox widths
$pnlVariables.Add_Resize({
    foreach ($tb in $script:VarTextBoxes.Values) {
        $tb.Width = [Math]::Max($pnlVariables.ClientSize.Width - 160, 50)
    }
})

# Buttons
$btnNew = New-Object System.Windows.Forms.Button
$btnNew.Text = "New"
$btnNew.Location = New-Object System.Drawing.Point(270, 540)
$btnNew.Size = New-Object System.Drawing.Size(80, 30)
$btnNew.Anchor = "Bottom, Left"
$form.Controls.Add($btnNew)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save"
$btnSave.Location = New-Object System.Drawing.Point(360, 540)
$btnSave.Size = New-Object System.Drawing.Size(80, 30)
$btnSave.Anchor = "Bottom, Left"
$form.Controls.Add($btnSave)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Delete"
$btnDelete.Location = New-Object System.Drawing.Point(450, 540)
$btnDelete.Size = New-Object System.Drawing.Size(80, 30)
$btnDelete.Anchor = "Bottom, Left"
$form.Controls.Add($btnDelete)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Run"
$btnRun.Location = New-Object System.Drawing.Point(540, 540)
$btnRun.Size = New-Object System.Drawing.Size(80, 30)
$btnRun.Anchor = "Bottom, Left"
$form.Controls.Add($btnRun)

$btnTerminate = New-Object System.Windows.Forms.Button
$btnTerminate.Text = "Terminate"
$btnTerminate.Location = New-Object System.Drawing.Point(630, 540)
$btnTerminate.Size = New-Object System.Drawing.Size(80, 30)
$btnTerminate.Enabled = $false
$btnTerminate.Anchor = "Bottom, Left"
$form.Controls.Add($btnTerminate)

# Output
$lblOutput = New-Object System.Windows.Forms.Label
$lblOutput.Text = "Output:"
$lblOutput.Location = New-Object System.Drawing.Point(270, 580)
$lblOutput.AutoSize = $true
$lblOutput.Anchor = "Bottom, Left"
$form.Controls.Add($lblOutput)

$txtOutput = New-Object System.Windows.Forms.TextBox
$txtOutput.Location = New-Object System.Drawing.Point(270, 600)
$txtOutput.Size = New-Object System.Drawing.Size(600, 120)
$txtOutput.Multiline = $true
$txtOutput.ScrollBars = "Vertical"
$txtOutput.ReadOnly = $true
$txtOutput.Anchor = "Bottom, Left, Right"
$form.Controls.Add($txtOutput)

# === detect all variables in the script and create textboxes dynamically ===
function Update-VariableControls {
    $scriptText = $txtScript.Text

    # Clear previous controls and mapping
    $pnlVariables.Controls.Clear()
    $script:VarTextBoxes.Clear()

    if ([string]::IsNullOrWhiteSpace($scriptText)) {
        return
    }

    # Find all PowerShell-style variable names in the script
    $matches = [regex]::Matches($scriptText, '\$(?<name>[A-Za-z_][A-Za-z0-9_]*)')

    # Unique variable names
    $varNames =
        $matches |
        ForEach-Object { $_.Groups['name'].Value } |
        Select-Object -Unique

    # Optionally ignore some built-in ones
    $ignore = @('env', 'HOME', 'PWD', 'PROFILE')
    $varNames = $varNames | Where-Object { $_ -notin $ignore }

    # Create one label + textbox per variable
    $top = 0
    foreach ($name in $varNames) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $name + ":"
        $lbl.Location = New-Object System.Drawing.Point(0, $top)
        $lbl.AutoSize = $true
        $pnlVariables.Controls.Add($lbl)

        $txt = New-Object System.Windows.Forms.TextBox
        $txt.Location = New-Object System.Drawing.Point(150, $top)
        $txt.Size = New-Object System.Drawing.Size(
            [Math]::Max($pnlVariables.ClientSize.Width - 160, 50), 20
        )
        $txt.Anchor = "Top, Left, Right"
        $pnlVariables.Controls.Add($txt)

        $script:VarTextBoxes[$name] = $txt

        $top += 25
    }
}

# Timer tick: read job output and detect completion
$timer.Add_Tick({
    if ($script:CurrentJob -ne $null) {

        if ($script:CurrentJob.HasMoreData) {
            $data = Receive-Job -Job $script:CurrentJob -Keep | Out-String
            if ($data) {
                $txtOutput.AppendText($data)
            }
        }

        if ($script:CurrentJob.State -ne 'Running' -and $script:CurrentJob.State -ne 'NotStarted') {

            if ($script:CurrentJob.HasMoreData) {
                $data = Receive-Job -Job $script:CurrentJob | Out-String
                if ($data) {
                    $txtOutput.AppendText($data)
                }
            }

            $btnRun.Enabled = $true
            $btnTerminate.Enabled = $false

            Remove-Job -Job $script:CurrentJob -Force -ErrorAction SilentlyContinue
            $script:CurrentJob = $null

            $timer.Stop()
        }
    }
})

# --- Drag-and-drop events for listCommands ---

$listCommands.Add_MouseDown({
    param($sender, $e)

    $script:DragIndex = $sender.IndexFromPoint($e.Location)
})

$listCommands.Add_MouseMove({
    param($sender, $e)

    if (($e.Button -band [System.Windows.Forms.MouseButtons]::Left) -and
        ($script:DragIndex -ge 0)) {

        [void]$sender.DoDragDrop(
            $sender.Items[$script:DragIndex],
            [System.Windows.Forms.DragDropEffects]::Move
        )
    }
})

$listCommands.Add_DragOver({
    param($sender, $e)

    $e.Effect = [System.Windows.Forms.DragDropEffects]::Move
})

$listCommands.Add_DragDrop({
    param($sender, $e)

    try {
        if ($script:DragIndex -lt 0 -or $script:DragIndex -ge $script:Commands.Count) {
            $script:DragIndex = -1
            return
        }

        $dropPoint = $sender.PointToClient(
            [System.Drawing.Point]::new($e.X, $e.Y)
        )

        $dropIndex = $sender.IndexFromPoint($dropPoint)

        if ($dropIndex -lt 0) {
            $dropIndex = $sender.Items.Count - 1
        }

        if ($dropIndex -eq $script:DragIndex) {
            $script:DragIndex = -1
            return
        }

        $list = New-Object 'System.Collections.Generic.List[object]'
        $list.AddRange($script:Commands)

        $moved = $list[$script:DragIndex]
        $list.RemoveAt($script:DragIndex)

        if ($dropIndex -gt $script:DragIndex) {
            $dropIndex--
        }

        $list.Insert($dropIndex, $moved)

        $script:Commands = $list.ToArray()
        Save-Commands
        Refresh-CommandList

        $listCommands.SelectedIndex = $dropIndex
    }
    finally {
        $script:DragIndex = -1
    }
})

# Events
$listCommands.Add_SelectedIndexChanged({
    $idx = $listCommands.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:Commands.Count) {
        $txtName.Text        = $script:Commands[$idx].Name
        $txtDescription.Text = $script:Commands[$idx].Description
        $txtScript.Text      = $script:Commands[$idx].Script
    } else {
        $txtName.Text        = ""
        $txtDescription.Text = ""
        $txtScript.Text      = ""
    }
    Update-VariableControls
})

$txtScript.Add_TextChanged({
    Update-VariableControls
})

$btnNew.Add_Click({
    $listCommands.ClearSelected()
    $txtName.Text        = ""
    $txtDescription.Text = ""
    $txtScript.Text      = ""
    Update-VariableControls
})

$btnSave.Add_Click({
    $name        = $txtName.Text.Trim()
    $description = $txtDescription.Text
    $scriptText  = $txtScript.Text

    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($scriptText)) {
        [System.Windows.Forms.MessageBox]::Show("Name and command cannot be empty.","Error","OK","Error") | Out-Null
        return
    }

    $idx = $listCommands.SelectedIndex

    if ($idx -ge 0 -and $idx -lt $script:Commands.Count) {
        $script:Commands[$idx].Name        = $name
        $script:Commands[$idx].Description = $description
        $script:Commands[$idx].Script      = $scriptText
    } else {
        $new = [PSCustomObject]@{
            Name        = $name
            Description = $description
            Script      = $scriptText
        }
        $script:Commands += $new
    }

    Save-Commands
    Refresh-CommandList
})

$btnDelete.Add_Click({
    $idx = $listCommands.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:Commands.Count) {
        return
    }

    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Delete command '$($script:Commands[$idx].Name)'?",
        "Confirm delete",
        "YesNo",
        "Warning"
    )

    if ($confirm -ne "Yes") {
        return
    }

    $newList = @()
    for ($i = 0; $i -lt $script:Commands.Count; $i++) {
        if ($i -ne $idx) {
            $newList += $script:Commands[$i]
        }
    }
    $script:Commands = $newList

    Save-Commands
    $txtName.Text        = ""
    $txtDescription.Text = ""
    $txtScript.Text      = ""
    Refresh-CommandList
    Update-VariableControls
})

$btnRun.Add_Click({
    $idx = $listCommands.SelectedIndex
    if ($idx -lt 0 -or $idx -ge $script:Commands.Count) {
        [System.Windows.Forms.MessageBox]::Show("Please select a command to run.","Info","OK","Information") | Out-Null
        return
    }

    if ($script:CurrentJob -ne $null -and $script:CurrentJob.State -eq 'Running') {
        [System.Windows.Forms.MessageBox]::Show("A command is already running. Terminate it first.","Info","OK","Information") | Out-Null
        return
    }

    $vars = @{}
    foreach ($entry in $script:VarTextBoxes.GetEnumerator()) {
        $vars[$entry.Key] = $entry.Value.Text
    }

    $scriptText = $script:Commands[$idx].Script

    $txtOutput.Text = ""

    $script:CurrentJob = Start-Job -ScriptBlock {
        param($scriptText, $vars)

        foreach ($pair in $vars.GetEnumerator()) {
            Set-Variable -Name $pair.Key -Value $pair.Value
        }

        Invoke-Expression $scriptText 2>&1
    } -ArgumentList $scriptText, $vars

    $btnRun.Enabled = $false
    $btnTerminate.Enabled = $true

    $timer.Start()
})

$btnTerminate.Add_Click({
    if ($script:CurrentJob -ne $null -and $script:CurrentJob.State -eq 'Running') {
        Stop-Job -Job $script:CurrentJob -ErrorAction SilentlyContinue
    }
})

# Start
Refresh-CommandList
Update-VariableControls
[System.Windows.Forms.Application]::EnableVisualStyles()
[void]$form.ShowDialog()
