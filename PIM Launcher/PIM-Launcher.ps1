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
        Title="PIM Launcher" Width="900" Height="780" MinWidth="760" MinHeight="640"
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI" FontSize="13"
        Background="#F3F4F6" Foreground="#1F2937">
  <Window.Resources>

    <!-- Pulsanti: angoli arrotondati, hover e stato disabilitato -->
    <Style x:Key="BaseButton" TargetType="Button">
      <Setter Property="Background" Value="White"/>
      <Setter Property="Foreground" Value="#1F2937"/>
      <Setter Property="BorderBrush" Value="#D1D5DB"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="14,6"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.85"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.7"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="Button" BasedOn="{StaticResource BaseButton}"/>
    <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource BaseButton}">
      <Setter Property="Background" Value="#0F6CBD"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="BorderBrush" Value="#0F6CBD"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>

    <!-- Campi di input -->
    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="BorderBrush" Value="#D1D5DB"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Height" Value="30"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>

    <!-- Testi -->
    <Style x:Key="MutedLabel" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#6B7280"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="SectionTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#6B7280"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>

    <!-- Card -->
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="White"/>
      <Setter Property="BorderBrush" Value="#E5E7EB"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="14"/>
      <Setter Property="Margin" Value="0,0,0,10"/>
    </Style>

    <!-- Griglia dei ruoli -->
    <Style x:Key="GridHeader" TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#F9FAFB"/>
      <Setter Property="Foreground" Value="#6B7280"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="BorderBrush" Value="#E5E7EB"/>
      <Setter Property="BorderThickness" Value="0,0,0,1"/>
    </Style>
    <Style x:Key="GridRow" TargetType="DataGridRow">
      <Setter Property="Background" Value="White"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#F5F9FF"/>
        </Trigger>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="#E8F1FB"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="GridCell" TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
      <Style.Triggers>
        <Trigger Property="IsSelected" Value="True">
          <Setter Property="Background" Value="Transparent"/>
          <Setter Property="BorderBrush" Value="Transparent"/>
          <Setter Property="Foreground" Value="#1F2937"/>
        </Trigger>
      </Style.Triggers>
    </Style>
    <Style x:Key="CellText" TargetType="TextBlock">
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="10,0"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>
    <Style x:Key="StateText" TargetType="TextBlock" BasedOn="{StaticResource CellText}">
      <Setter Property="Foreground" Value="#6B7280"/>
      <Style.Triggers>
        <DataTrigger Binding="{Binding IsActive}" Value="True">
          <Setter Property="Foreground" Value="#107C10"/>
          <Setter Property="FontWeight" Value="SemiBold"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <!-- Intestazione -->
    <Border Grid.Row="0" Background="#10253F" Padding="20,14">
      <StackPanel>
        <TextBlock Text="PIM Launcher" FontSize="20" FontWeight="SemiBold" Foreground="White"/>
        <TextBlock Text="Attiva i ruoli Entra ID e apri i portali con una sessione già aggiornata"
                   FontSize="12" Foreground="#A9BCD0" Margin="0,2,0,0"/>
      </StackPanel>
    </Border>

    <Grid Grid.Row="1" Margin="16,14,16,16">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*" MinHeight="140"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="130"/>
      </Grid.RowDefinitions>

      <!-- Cliente -->
      <Border Grid.Row="0" Style="{StaticResource Card}">
        <DockPanel>
          <TextBlock Text="Cliente" Style="{StaticResource MutedLabel}" Margin="0,0,10,0"/>
          <ComboBox x:Name="cmbTenant" Width="280"/>
          <Button x:Name="btnConnect" Content="Connetti" Margin="10,0,0,0" Style="{StaticResource PrimaryButton}"/>
          <TextBlock x:Name="lblAccount" Style="{StaticResource MutedLabel}" Margin="14,0,0,0"/>
        </DockPanel>
      </Border>

      <!-- Ruoli -->
      <Border Grid.Row="1" Style="{StaticResource Card}" Padding="0">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <TextBlock Text="RUOLI ELEGGIBILI" Style="{StaticResource SectionTitle}" Margin="14,12,14,8"/>
          <DataGrid x:Name="grid" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False"
                    CanUserResizeRows="False" HeadersVisibility="Column" GridLinesVisibility="Horizontal"
                    HorizontalGridLinesBrush="#EEF0F3" BorderBrush="#E5E7EB" BorderThickness="0,1,0,0"
                    Background="White" RowHeight="34" SelectionMode="Single"
                    ColumnHeaderStyle="{StaticResource GridHeader}"
                    RowStyle="{StaticResource GridRow}" CellStyle="{StaticResource GridCell}">
            <DataGrid.Columns>
              <DataGridTemplateColumn Header="" Width="44">
                <DataGridTemplateColumn.CellTemplate>
                  <DataTemplate>
                    <CheckBox HorizontalAlignment="Center" VerticalAlignment="Center"
                              IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}">
                      <CheckBox.Style>
                        <Style TargetType="CheckBox">
                          <Style.Triggers>
                            <DataTrigger Binding="{Binding IsActive}" Value="True">
                              <Setter Property="IsEnabled" Value="False"/>
                            </DataTrigger>
                          </Style.Triggers>
                        </Style>
                      </CheckBox.Style>
                    </CheckBox>
                  </DataTemplate>
                </DataGridTemplateColumn.CellTemplate>
              </DataGridTemplateColumn>
              <DataGridTextColumn Header="Ruolo"  Binding="{Binding Role}"  IsReadOnly="True" Width="2*" ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Scope" Binding="{Binding Scope}" IsReadOnly="True" Width="*"  ElementStyle="{StaticResource CellText}"/>
              <DataGridTextColumn Header="Status"  Binding="{Binding State}" IsReadOnly="True" Width="*"  ElementStyle="{StaticResource StateText}"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </Border>

      <!-- Attivazione -->
      <Border Grid.Row="2" Style="{StaticResource Card}" Margin="0,10,0,10">
        <StackPanel>
          <TextBlock Text="ATTIVAZIONE" Style="{StaticResource SectionTitle}"/>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="70"/>
              <ColumnDefinition Width="Auto"/><ColumnDefinition Width="120"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="Giustificazione" Style="{StaticResource MutedLabel}" Margin="0,0,8,0"/>
            <TextBox x:Name="txtJust" Grid.Column="1" Height="30"/>
            <TextBlock Text="Ore" Grid.Column="2" Style="{StaticResource MutedLabel}" Margin="12,0,8,0"/>
            <ComboBox x:Name="cmbDur" Grid.Column="3">
              <ComboBoxItem Content="1"/><ComboBoxItem Content="2"/>
              <ComboBoxItem Content="4"/><ComboBoxItem Content="8"/>
            </ComboBox>
            <TextBlock Text="Ticket" Grid.Column="4" Style="{StaticResource MutedLabel}" Margin="12,0,8,0"/>
            <TextBox x:Name="txtTicket" Grid.Column="5" Height="30"/>
          </Grid>
          <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
            <Button x:Name="btnActivate" Content="Attiva selezionati" Style="{StaticResource PrimaryButton}"/>
            <Button x:Name="btnRefresh" Content="Aggiorna" Margin="8,0,0,0"/>
            <CheckBox x:Name="chkAutoOpen" Content="Dopo l'attivazione apri i portali" IsChecked="True"
                      VerticalAlignment="Center" Margin="16,0,0,0"/>
          </StackPanel>
        </StackPanel>
      </Border>

      <!-- Portali -->
      <Border Grid.Row="3" Style="{StaticResource Card}">
        <StackPanel>
          <TextBlock Text="PORTALI" Style="{StaticResource SectionTitle}"/>
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="Sessione browser" Style="{StaticResource MutedLabel}" Margin="0,0,8,0"/>
            <ComboBox x:Name="cmbMode" Width="220" SelectedIndex="2">
              <ComboBoxItem Content="Isolata e nuova"/>
              <ComboBoxItem Content="InPrivate"/>
              <ComboBoxItem Content="Profilo Edge del cliente"/>
            </ComboBox>
            <Button x:Name="btnOpenAll" Content="Apri tutti i portali" Margin="8,0,0,0"/>
          </StackPanel>
          <WrapPanel x:Name="pnlPortals" Margin="0,10,0,0"/>
        </StackPanel>
      </Border>

      <!-- Log -->
      <Border Grid.Row="4" CornerRadius="8" Background="#0F172A" Padding="6">
        <TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                 Background="Transparent" Foreground="#D1D5DB" BorderThickness="0"
                 FontFamily="Consolas" FontSize="12" VerticalContentAlignment="Top"/>
      </Border>
    </Grid>
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
        $row.Scope = if ($e.directoryScopeId -eq '/') { 'Intero tenant' } else { $e.directoryScopeId }
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
