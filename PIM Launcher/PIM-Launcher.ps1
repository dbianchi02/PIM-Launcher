#Requires -Version 7.0
<#
.SYNOPSIS
    PIM Launcher - GUI WPF per attivare i ruoli Entra ID PIM sui tenant dei clienti
    e aprire subito i portali con una sessione browser pulita (token nuovo, con i ruoli attivati).

.DESCRIPTION
    Si connette a un tenant con Microsoft Graph, elenca i ruoli PIM eleggibili, li attiva con
    giustificazione, durata e ticket, attende che risultino attivi e apre i portali del cliente
    in Microsoft Edge (profilo del cliente, sessione isolata oppure InPrivate).

.PARAMETER ConfigPath
    Percorso del file di configurazione. Default: tenants.json nella cartella dello script.

.EXAMPLE
    .\PIM-Launcher.ps1

.EXAMPLE
    .\PIM-Launcher.ps1 -ConfigPath C:\Config\tenants.json

.NOTES
    Requisiti: Windows, PowerShell 7, modulo Microsoft.Graph.Authentication, Microsoft Edge.
    Configurazione: tenants.json nella stessa cartella (vedi tenants.example.json).
#>
[CmdletBinding()]
param([string]$ConfigPath = (Join-Path $PSScriptRoot 'tenants.json'))

if (-not $IsWindows) { throw 'WPF richiede Windows.' }

# WPF vuole un thread STA: se serve, rilancia se stesso
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    & pwsh -STA -NoProfile -File $PSCommandPath -ConfigPath $ConfigPath
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    throw 'Modulo mancante: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
}
Import-Module Microsoft.Graph.Authentication
if (-not (Test-Path $ConfigPath)) { throw "Configurazione non trovata: $ConfigPath" }

$config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$Scopes = @('RoleManagement.ReadWrite.Directory', 'User.Read')
$TempRoot = Join-Path $env:TEMP 'PIMLauncher'

# Pulizia profili browser temporanei vecchi
if (Test-Path $TempRoot) {
    Get-ChildItem $TempRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt (Get-Date).AddDays(-1) |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}

if (-not ('RoleRow' -as [type])) {
    Add-Type -TypeDefinition @'
public class RoleRow {
    public bool Selected { get; set; }
    public string Role { get; set; }
    public string Scope { get; set; }
    public string State { get; set; }
    public bool IsActive { get; set; }
    public string RoleDefinitionId { get; set; }
    public string DirectoryScopeId { get; set; }
}
'@
}

# ---------------------------------------------------------------- UI (XAML)
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PIM Launcher" Width="880" Height="740" MinWidth="720" MinHeight="600"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*" MinHeight="140"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="150"/>
    </Grid.RowDefinitions>

    <DockPanel Grid.Row="0" Margin="0,0,0,8">
      <TextBlock Text="Cliente" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <ComboBox x:Name="cmbTenant" Width="260"/>
      <Button x:Name="btnConnect" Content="Connetti" Margin="8,0,0,0" Padding="14,4"/>
      <TextBlock x:Name="lblAccount" Margin="12,0,0,0" VerticalAlignment="Center" Foreground="Gray"/>
    </DockPanel>

    <DataGrid x:Name="grid" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False"
              HeadersVisibility="Column" GridLinesVisibility="Horizontal" SelectionMode="Single">
      <DataGrid.Columns>
        <DataGridCheckBoxColumn Header="" Width="34"
            Binding="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/>
        <DataGridTextColumn Header="Ruolo"  Binding="{Binding Role}"  IsReadOnly="True" Width="2*"/>
        <DataGridTextColumn Header="Scope" Binding="{Binding Scope}" IsReadOnly="True" Width="*"/>
        <DataGridTextColumn Header="Status"  Binding="{Binding State}" IsReadOnly="True" Width="*"/>
      </DataGrid.Columns>
    </DataGrid>

    <Grid Grid.Row="2" Margin="0,8,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/><ColumnDefinition Width="70"/>
        <ColumnDefinition Width="Auto"/><ColumnDefinition Width="120"/>
      </Grid.ColumnDefinitions>
      <TextBlock Text="Giustificazione" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <TextBox x:Name="txtJust" Grid.Column="1" Padding="3"/>
      <TextBlock Text="Ore" Grid.Column="2" VerticalAlignment="Center" Margin="10,0,6,0"/>
      <ComboBox x:Name="cmbDur" Grid.Column="3">
        <ComboBoxItem Content="1"/><ComboBoxItem Content="2"/>
        <ComboBoxItem Content="4"/><ComboBoxItem Content="8"/>
      </ComboBox>
      <TextBlock Text="Ticket" Grid.Column="4" VerticalAlignment="Center" Margin="10,0,6,0"/>
      <TextBox x:Name="txtTicket" Grid.Column="5" Padding="3"/>
    </Grid>

    <StackPanel Grid.Row="3" Margin="0,10,0,8">
      <StackPanel Orientation="Horizontal">
        <Button x:Name="btnActivate" Content="Attiva selezionati" Padding="14,5" FontWeight="SemiBold"/>
        <Button x:Name="btnRefresh" Content="Aggiorna" Padding="14,5" Margin="8,0,0,0"/>
        <CheckBox x:Name="chkAutoOpen" Content="Dopo l'attivazione apri i portali" IsChecked="True"
                  VerticalAlignment="Center" Margin="16,0,0,0"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
        <TextBlock Text="Sessione browser" VerticalAlignment="Center" Margin="0,0,8,0"/>
        <ComboBox x:Name="cmbMode" Width="220" SelectedIndex="2">
          <ComboBoxItem Content="Isolata e nuova"/>
          <ComboBoxItem Content="InPrivate"/>
          <ComboBoxItem Content="Profilo Edge del cliente"/>
        </ComboBox>
        <Button x:Name="btnOpenAll" Content="Apri tutti i portali" Padding="12,4" Margin="8,0,0,0"/>
      </StackPanel>
      <WrapPanel x:Name="pnlPortals" Margin="0,8,0,0"/>
    </StackPanel>

    <TextBox x:Name="txtLog" Grid.Row="4" IsReadOnly="True" TextWrapping="Wrap"
             VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12"/>
  </Grid>
</Window>
'@

$window = [Windows.Markup.XamlReader]::Load([System.Xml.XmlNodeReader]::new($xaml))
foreach ($n in 'cmbTenant','btnConnect','lblAccount','grid','txtJust','cmbDur','txtTicket',
              'btnActivate','btnRefresh','chkAutoOpen','cmbMode','btnOpenAll','pnlPortals','txtLog') {
    Set-Variable -Name $n -Value $window.FindName($n) -Scope Script
}

# ---------------------------------------------------------------- Helper
function Pump { $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Background) }

function Write-UiLog([string]$Message) {
    $line = '[{0:HH:mm:ss}] {1}' -f (Get-Date), $Message
    $txtLog.AppendText($line + "`r`n")
    $txtLog.ScrollToEnd()
    Pump
}

function Get-ErrMsg($err) {
    if ($err.ErrorDetails.Message) {
        try { return ($err.ErrorDetails.Message | ConvertFrom-Json).error.message } catch { return $err.ErrorDetails.Message }
    }
    $err.Exception.Message
}

function Invoke-Ui([scriptblock]$Action, $Argument) {
    $window.Cursor = [System.Windows.Input.Cursors]::Wait
    try { & $Action $Argument }
    catch {
        $msg = Get-ErrMsg $_
        Write-UiLog "ERRORE: $msg"
        if ($msg -match 'claims|authentication context|MultiFactor|MFA') {
            Write-UiLog 'Il ruolo richiede uno step-up (MFA/auth context). Premi Connetti per ri-autenticarti e riprova.'
        }
    }
    finally { $window.Cursor = $null }
}

function Get-CurrentTenant { @($config.Tenants)[$cmbTenant.SelectedIndex] }

function Get-GraphAll([string]$Uri) {
    $out = @()
    while ($Uri) {
        $r = Invoke-MgGraphRequest -Method GET -Uri $Uri
        $out += $r.value
        $Uri = $r.'@odata.nextLink'
    }
    $out
}

# ---------------------------------------------------------------- Graph / PIM
function Connect-Tenant($t) {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    $p = @{ TenantId = $t.TenantId; Scopes = $Scopes; NoWelcome = $true; ContextScope = 'Process' }
    if ($config.ClientId) { $p.ClientId = $config.ClientId }
    Write-UiLog "Connessione a $($t.Name)..."
    Connect-MgGraph @p | Out-Null
    $script:me = Invoke-MgGraphRequest -Method GET -Uri '/v1.0/me?$select=id,userPrincipalName'
    $lblAccount.Text = $script:me.userPrincipalName
    Write-UiLog "Connesso come $($script:me.userPrincipalName)"
}

function Update-Roles {
    $q = '?$expand=roleDefinition'
    $eligible = Get-GraphAll "/v1.0/roleManagement/directory/roleEligibilityScheduleInstances/filterByCurrentUser(on='principal')$q"
    $active   = Get-GraphAll "/v1.0/roleManagement/directory/roleAssignmentScheduleInstances/filterByCurrentUser(on='principal')"
    $rows = foreach ($e in $eligible) {
        $a = $active | Where-Object { $_.roleDefinitionId -eq $e.roleDefinitionId -and $_.directoryScopeId -eq $e.directoryScopeId } |
             Select-Object -First 1
        $row = [RoleRow]::new()
        $row.Role = $e.roleDefinition.displayName
        $row.Scope = if ($e.directoryScopeId -eq '/') { 'Tenant' } else { $e.directoryScopeId }
        $row.RoleDefinitionId = $e.roleDefinitionId
        $row.DirectoryScopeId = $e.directoryScopeId
        if ($a) {
            $row.IsActive = $true
            $row.State = if ($a.assignmentType -eq 'Activated' -and $a.endDateTime) {
                try { 'Attivo fino alle ' + ([datetime]$a.endDateTime).ToLocalTime().ToString('HH:mm') } catch { 'Attivo' }
            } else { 'Attivo (permanente)' }
        } else { $row.State = 'Eleggibile' }
        $row
    }
    $grid.ItemsSource = @($rows | Sort-Object Role)
    Write-UiLog "Ruoli eleggibili: $(@($rows).Count)"
}

function Wait-RoleActive([string[]]$RoleDefinitionIds, [int]$TimeoutSec = 60) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        Start-Sleep -Seconds 3
        $act = Get-GraphAll "/v1.0/roleManagement/directory/roleAssignmentScheduleInstances/filterByCurrentUser(on='principal')"
        $missing = $RoleDefinitionIds | Where-Object { $id = $_; -not ($act | Where-Object { $_.roleDefinitionId -eq $id }) }
        if (-not $missing) { return $true }
        Pump
    }
    $false
}

function Invoke-Activation {
    $sel = @($grid.ItemsSource | Where-Object { $_.Selected -and -not $_.IsActive })
    if (-not $sel) { Write-UiLog 'Nessun ruolo eleggibile selezionato.'; return }
    $just = $txtJust.Text.Trim()
    if (-not $just) { Write-UiLog 'Inserisci una giustificazione.'; return }
    $hours = [int]$cmbDur.SelectedItem.Content
    $ticket = $txtTicket.Text.Trim()

    $activated = @()
    foreach ($r in $sel) {
        Write-UiLog "Attivo: $($r.Role) ($hours h)..."
        $body = @{
            action           = 'selfActivate'
            principalId      = $script:me.id
            roleDefinitionId = $r.RoleDefinitionId
            directoryScopeId = $r.DirectoryScopeId
            justification    = $just
            scheduleInfo     = @{
                startDateTime = (Get-Date).ToUniversalTime().ToString('o')
                expiration    = @{ type = 'afterDuration'; duration = "PT${hours}H" }
            }
        }
        if ($ticket) { $body.ticketInfo = @{ ticketNumber = $ticket; ticketSystem = 'ITSM' } }
        try {
            $res = Invoke-MgGraphRequest -Method POST -ContentType 'application/json' `
                -Uri '/v1.0/roleManagement/directory/roleAssignmentScheduleRequests' -Body ($body | ConvertTo-Json -Depth 6)
            Write-UiLog "  -> $($res.status)"
            if ($res.status -notmatch 'Pending') { $activated += $r.RoleDefinitionId }
            else { Write-UiLog '  Richiesta in attesa di approvazione.' }
        }
        catch { Write-UiLog "  ERRORE su $($r.Role): $(Get-ErrMsg $_)" }
    }

    if ($activated) {
        Write-UiLog 'Attendo che i ruoli risultino attivi...'
        if (Wait-RoleActive $activated) { Write-UiLog 'Ruoli attivi.' }
        else { Write-UiLog 'Timeout: controlla lo stato con Aggiorna.' }
    }
    Update-Roles

    if ($activated) {
        Write-UiLog 'Nota: Exchange, SharePoint e Teams possono impiegare diversi minuti a propagare il ruolo.'
        if ($chkAutoOpen.IsChecked) { Open-Portals (Get-CurrentTenant) (Get-TenantPortals (Get-CurrentTenant)) }
    }
}

# ---------------------------------------------------------------- Portali / browser
function Get-TenantPortals($t) {
    if ($t.PortalNames) { @($config.Portals | Where-Object { $t.PortalNames -contains $_.Name }) }
    else { @($config.Portals) }
}

function Resolve-Url($t, $url) {
    $url.Replace('{tenantId}', $t.TenantId).Replace('{upn}', "$($t.Upn)").Replace('{spo}', "$($t.Spo)")
}

function Open-Portals($t, $portals) {
    if (-not $portals) { Write-UiLog 'Nessun portale configurato.'; return }
    $urls = @($portals | ForEach-Object { Resolve-Url $t $_.Url })
    $edgeArgs = @()
    switch ($cmbMode.SelectedIndex) {
        0 {
            $dir = Join-Path $TempRoot (('{0}-{1}' -f ($t.Name -replace '\W', '_'), [guid]::NewGuid().ToString('N').Substring(0, 6)))
            $edgeArgs += "--user-data-dir=`"$dir`"", '--no-first-run', '--no-default-browser-check', '--new-window'
        }
        1 { $edgeArgs += '--inprivate', '--new-window' }
        2 {
            if ($t.EdgeProfile) { $edgeArgs += "--profile-directory=`"$($t.EdgeProfile)`"" }
            $edgeArgs += '--new-window'
        }
    }
    Start-Process msedge -ArgumentList ($edgeArgs + $urls)
    Write-UiLog "Aperti $($urls.Count) portale/i per $($t.Name)."
}

function Update-PortalButtons {
    $pnlPortals.Children.Clear()
    $t = Get-CurrentTenant
    foreach ($p in (Get-TenantPortals $t)) {
        $b = [System.Windows.Controls.Button]::new()
        $b.Content = $p.Name; $b.Padding = '10,3'; $b.Margin = '0,0,6,6'; $b.Tag = $p
        $b.Add_Click({ param($s, $e) Invoke-Ui { param($portal) Open-Portals (Get-CurrentTenant) @($portal) } $s.Tag })
        [void]$pnlPortals.Children.Add($b)
    }
}

# ---------------------------------------------------------------- Eventi
foreach ($t in @($config.Tenants)) { [void]$cmbTenant.Items.Add($t.Name) }
$cmbTenant.SelectedIndex = 0
$cmbDur.SelectedIndex = 2
foreach ($i in $cmbDur.Items) { if ([int]$i.Content -eq [int]$config.DefaultDurationHours) { $cmbDur.SelectedItem = $i } }
$txtJust.Text = "$($config.DefaultJustification)"
Update-PortalButtons

$cmbTenant.Add_SelectionChanged({
    $grid.ItemsSource = $null
    $lblAccount.Text = ''
    Update-PortalButtons
})
$btnConnect.Add_Click({ Invoke-Ui { Connect-Tenant (Get-CurrentTenant); Update-Roles } })
$btnRefresh.Add_Click({ Invoke-Ui { Update-Roles } })
$btnActivate.Add_Click({ Invoke-Ui { Invoke-Activation } })
$btnOpenAll.Add_Click({ Invoke-Ui { $t = Get-CurrentTenant; Open-Portals $t (Get-TenantPortals $t) } })

[void]$window.ShowDialog()