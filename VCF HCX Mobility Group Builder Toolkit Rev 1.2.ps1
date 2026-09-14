<#
.SYNOPSIS
  VCF 91 HCX MobilityGroup Builder Toolkit Rev 1.1a-Import-Hotfix.ps1
.DESCRIPTION
  PowerShell 7 WPF utility for connecting to an HCX 9.1 Manager, importing VM names,
  discovering source and destination inventory, mapping multiple vNICs, validating
  selections, and exporting a versioned CSV for a future Phase II mobility-group creator.

  IMPORTANT: HCX 9.x operations in this tool use REST, not VMware.VimAutomation.Hcx.
  The legacy HCX PowerCLI module shipped with current VCF.PowerCLI is not used.

  CSV input requires a VMName column. Name is also accepted.
.NOTES
  Windows and PowerShell 7 or later are required. Credentials and tokens are never exported.

#>
[CmdletBinding()]
param([switch]$NoRelaunch)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 3.0

function Ensure-SelfSignedScriptCertificate([string]$TargetPath) {
    if (-not $IsWindows -or -not (Test-Path -LiteralPath $TargetPath)) { return $false }
    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $TargetPath -ErrorAction SilentlyContinue
        if ($sig.Status -eq 'Valid') { return $true }
        $subject = 'CN=HCX91 Mobility CSV Builder Local Code Signing'
        $cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
            Where-Object { $_.Subject -eq $subject -and $_.NotAfter -gt (Get-Date).AddDays(30) } |
            Sort-Object NotAfter -Descending | Select-Object -First 1
        if (-not $cert) {
            $cert = New-SelfSignedCertificate -Subject $subject -Type CodeSigningCert `
                -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 2048 `
                -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(3)
        }
        foreach ($storeName in 'TrustedPublisher','Root') {
            $store = [Security.Cryptography.X509Certificates.X509Store]::new($storeName,'CurrentUser')
            try {
                $store.Open('ReadWrite')
                if (-not ($store.Certificates | Where-Object Thumbprint -eq $cert.Thumbprint)) { $store.Add($cert) }
            } finally { $store.Close() }
        }
        $signed = Set-AuthenticodeSignature -LiteralPath $TargetPath -Certificate $cert -HashAlgorithm SHA256
        return $signed.Status -in 'Valid','UnknownError'
    } catch {
        Write-Warning "Self-signing failed: $($_.Exception.Message)"
        return $false
    }
}

$null = Ensure-SelfSignedScriptCertificate $PSCommandPath
if ($PSVersionTable.PSVersion.Major -lt 7 -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    if (-not $NoRelaunch) {
        $pwsh = (Get-Command pwsh.exe -ErrorAction SilentlyContinue).Source
        if (-not $pwsh) { $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue).Source }
        if (-not $pwsh) { throw 'PowerShell 7 or later is required.' }
        & $pwsh -NoProfile -ExecutionPolicy Bypass -STA -File $PSCommandPath -NoRelaunch
        exit $LASTEXITCODE
    }
}
if (-not $IsWindows) { throw 'This WPF utility requires Windows.' }
if ($PSVersionTable.PSVersion.Major -lt 7) { throw 'PowerShell 7 or later is required.' }

Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Xaml,System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$script:OutputBase = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { (Get-Location).Path }
$script:RunDir = Join-Path $script:OutputBase ('HCX91-MobilityCSV-Run-' + (Get-Date -Format yyyyMMdd-HHmmss))
New-Item -ItemType Directory -Path $script:RunDir -Force | Out-Null
$script:LogFile = Join-Path $script:RunDir ('HCX91-MobilityCSV-' + (Get-Date -Format yyyyMMdd-HHmmss) + '.log')
$script:Rows = [Collections.ObjectModel.ObservableCollection[object]]::new()
$script:Hcx = [ordered]@{ BaseUri=''; Session=$null; Connected=$false; Version=''; User=''; SitePairs=@(); EndpointProfile=''; Headers=@{} }
$script:Inventory = [ordered]@{ VMs=@(); Networks=@(); Datastores=@(); Computes=@(); Folders=@(); Sites=@() ; Policies=@()}
$script:ValidationCurrent = $false
$script:SuppressEditInvalidation = $false
$script:LastCsv = $null
$script:SchemaVersion = '2.0'
$script:MobilityGroupMaximum = 50
$script:SourceVIServer=$null
$script:DestinationVIServer=$null

function DoEvents { try { [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background) } catch {} }
function Log([string]$Message,[ValidateSet('INFO','WARN','ERROR','PASS')][string]$Level='INFO') {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') [$Level] $Message"
    Add-Content -LiteralPath $script:LogFile -Value $line
    if ($script:txtLog) { $script:txtLog.AppendText($line + [Environment]::NewLine); $script:txtLog.ScrollToEnd(); DoEvents }
}
function Wait-Tcp([string]$HostName,[int]$Port=443,[int]$Seconds=15) {
    $end=(Get-Date).AddSeconds($Seconds)
    while((Get-Date) -lt $end) {
        try { $c=[Net.Sockets.TcpClient]::new(); $a=$c.BeginConnect($HostName,$Port,$null,$null); if($a.AsyncWaitHandle.WaitOne(1000)){ $c.EndConnect($a);$c.Close();return $true };$c.Close() } catch {}
    }
    return $false
}
function Has-Module([string]$Name) { [bool](Get-Module -ListAvailable -Name $Name) }
function Ensure-Module([string]$Name) {
    if (Has-Module $Name) { Import-Module $Name -ErrorAction Stop | Out-Null; return $true }
    $old=$ProgressPreference
    try {
        $ProgressPreference='SilentlyContinue'
        Install-PackageProvider NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
        Set-PSRepository PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue | Out-Null
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -AcceptLicense
        Import-Module $Name -ErrorAction Stop | Out-Null
        Log "$Name installed and imported." PASS
        return $true
    } catch { Log "$Name installation failed: $($_.Exception.Message)" ERROR; return $false }
    finally { $ProgressPreference=$old }
}
function Update-Prerequisites {
    $script:lblPS.Text=$PSVersionTable.PSVersion.ToString();$script:lblPS.Foreground='LightGreen'
    $script:lblSTA.Text=[Threading.Thread]::CurrentThread.ApartmentState.ToString();$script:lblSTA.Foreground=if([Threading.Thread]::CurrentThread.ApartmentState -eq 'STA'){'LightGreen'}else{'Tomato'}
    $found=Has-Module 'VCF.PowerCLI';$script:lblPCLI.Text=if($found){'Found'}else{'Missing'};$script:lblPCLI.Foreground=if($found){'LightGreen'}else{'Tomato'}
    $script:lblApi.Text=if($script:Hcx.Connected){'Authenticated'}else{'Not connected'};$script:lblApi.Foreground=if($script:Hcx.Connected){'LightGreen'}else{'#76C7D8'}
}
function Set-OutputBase([string]$Path) {
    $candidate=[Environment]::ExpandEnvironmentVariables($Path.Trim())
    if(-not$candidate){throw 'Select or enter an output base path.'}
    if(-not(Test-Path -LiteralPath $candidate)){New-Item -ItemType Directory -Path $candidate -Force|Out-Null}
    $candidate=(Resolve-Path -LiteralPath $candidate).Path
    $test=Join-Path $candidate ('.hcx-write-test-'+[guid]::NewGuid().ToString('N'))
    try{[IO.File]::WriteAllText($test,'test');Remove-Item -LiteralPath $test -Force} catch {throw "Output path is not writable: $candidate"}
    $oldLog=$script:LogFile
    $script:OutputBase=$candidate
    $script:RunDir=Join-Path $script:OutputBase ('HCX91-MobilityCSV-Run-'+(Get-Date -Format yyyyMMdd-HHmmss))
    New-Item -ItemType Directory -Path $script:RunDir -Force|Out-Null
    $script:LogFile=Join-Path $script:RunDir ('HCX91-MobilityCSV-'+(Get-Date -Format yyyyMMdd-HHmmss)+'.log')
    if(Test-Path -LiteralPath $oldLog){Copy-Item -LiteralPath $oldLog -Destination $script:LogFile -Force}
    $script:txtOutputPath.Text=$script:OutputBase
$script:chkP2Security.IsChecked=$false
    Log "Output base path changed to $($script:OutputBase). Active run folder: $($script:RunDir)" PASS
}
function Select-OutputBase {
    $dialog=New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description='Select the default base folder for HCX CSV files, logs, and run artifacts'
    $dialog.SelectedPath=$script:OutputBase
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){Set-OutputBase $dialog.SelectedPath}
}
function Get-Value($Object,[string[]]$Names) {
    if ($null -eq $Object) { return $null }
    foreach($n in $Names) {
        $p=$Object.PSObject.Properties[$n]
        if($p -and $null -ne $p.Value -and "$($p.Value)" -ne '') { return $p.Value }
    }
    return $null
}
function Get-ArrayFromResponse($Response) {
    if ($null -eq $Response) { return @() }
    if ($Response -is [array]) { return @($Response) }
    foreach($n in 'items','elements','data','results','content','list','vms','networks','datastores','computeContainers','folders','sites','sitePairs') {
        $p=$Response.PSObject.Properties[$n]; if($p -and $null -ne $p.Value) { return @($p.Value) }
    }
    return @($Response)
}
function Normalize-InventoryObject($Object,[string]$Kind) {
    $name=Get-Value $Object @('name','displayName','vmName','networkName','datastoreName','computeName','folderName','siteName','resourceName')
    $id=Get-Value $Object @('id','uid','uuid','objectId','vmId','networkId','datastoreId','computeId','folderId','siteId','resourceId')
    [pscustomobject]@{ Kind=$Kind; Name=[string]$name; Id=[string]$id; Raw=$Object }
}
function Invoke-HcxRest([string]$Method,[string]$Path,$Body=$null,[hashtable]$Headers=@{},[switch]$AllowFailure) {
    if (-not $script:Hcx.BaseUri) { throw 'HCX base URI is not initialized.' }
    $uri=if($Path -match '^https?://'){$Path}else{$script:Hcx.BaseUri.TrimEnd('/')+'/'+$Path.TrimStart('/')}
    $all=@{ Accept='application/json' }
    foreach($k in $script:Hcx.Headers.Keys){$all[$k]=$script:Hcx.Headers[$k]}
    foreach($k in $Headers.Keys){$all[$k]=$Headers[$k]}
    $p=@{Method=$Method;Uri=$uri;Headers=$all;SkipCertificateCheck=$true;ErrorAction='Stop';TimeoutSec=90}
    if($script:Hcx.Session){$p.WebSession=$script:Hcx.Session}
    if($null -ne $Body){$p.Body=$Body|ConvertTo-Json -Depth 50 -Compress;$p.ContentType='application/json'}
    try { Invoke-RestMethod @p }
    catch {
        $status='';$detail=$_.Exception.Message
        try{$status=[int]$_.Exception.Response.StatusCode}catch{}
        try{if($_.ErrorDetails.Message){$detail=$_.ErrorDetails.Message}}catch{}
        Log "HCX REST FAILURE: Method='$Method'; Path='$Path'; HTTP='$status'; Response='$detail'." ERROR
        if($AllowFailure){return $null}
        throw "HCX REST $Method $Path failed. HTTP=$status; Response=$detail; Exception=$($_.Exception.Message)"
    }
}
function Connect-Hcx91Rest([string]$Fqdn,[string]$User,[string]$Password) {
    $hostName=($Fqdn -replace '^https?://','').TrimEnd('/')
    if(-not$hostName -or -not$User -or -not$Password){throw 'HCX Manager FQDN, username, and password are required.'}
    if(-not(Wait-Tcp $hostName 443 15)){throw "TCP 443 is not reachable on $hostName."}
    $script:Hcx.BaseUri='https://'+$hostName
    $script:Hcx.Session=[Microsoft.PowerShell.Commands.WebRequestSession]::new()
    $script:Hcx.Headers=@{}
    $uri=$script:Hcx.BaseUri+'/hybridity/api/sessions'
    $json=@{username=$User;password=$Password}|ConvertTo-Json -Compress
    $response=$null
    $failures=[Collections.Generic.List[string]]::new()
    $pair=[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$User`:$Password"))
    foreach($profile in @(
        @{Name='HCX JSON session';Headers=@{Accept='application/json'}},
        @{Name='HCX JSON session with Basic bootstrap';Headers=@{Accept='application/json';Authorization="Basic $pair"}}
    )){
        try{
            $p=@{Method='POST';Uri=$uri;Headers=$profile.Headers;Body=$json;ContentType='application/json';WebSession=$script:Hcx.Session;SkipCertificateCheck=$true;ErrorAction='Stop';TimeoutSec=90}
            $response=Invoke-WebRequest @p
            $token=[string]$response.Headers['x-hm-authorization']
            if(-not$token){$token=[string]$response.Headers['X-HM-Authorization']}
            if(-not$token){throw 'HTTP success was returned, but the x-hm-authorization response header was missing.'}
            $script:Hcx.Headers=@{'x-hm-authorization'=$token}
            $script:Hcx.EndpointProfile=$profile.Name
            break
        }catch{
            $status='';$detail=$_.Exception.Message
            try{$status=[int]$_.Exception.Response.StatusCode}catch{}
            try{if($_.ErrorDetails.Message){$detail=$_.ErrorDetails.Message}}catch{}
            $safeDetail=($detail -replace [regex]::Escape($Password),'********')
            $failures.Add("$($profile.Name): HTTP=$status; $safeDetail")
            Log "HCX authentication attempt '$($profile.Name)' failed. HTTP=$status; $safeDetail" WARN
        }
    }
    if(-not$script:Hcx.Headers.ContainsKey('x-hm-authorization')){
        throw "HCX session creation failed at POST /hybridity/api/sessions. $($failures -join ' | ')"
    }
    $script:Hcx.User=$User;$script:Hcx.Connected=$true
    $ver=$null
    foreach($path in '/hybridity/api/about','/hybridity/api/version','/api/v1/version','/api/version'){
        try{$r=Invoke-HcxRest GET $path -AllowFailure;$ver=Get-Value $r @('version','buildVersion','productVersion','releaseVersion');if($ver){break}}catch{}
    }
    $script:Hcx.Version=[string]$ver
    if($script:Hcx.Version -and $script:Hcx.Version -notmatch '^9\.1(?:\.|$)'){
        Disconnect-HcxRest
        throw "The connected HCX Manager reports version '$ver'. This Phase I tool requires HCX 9.1.x."
    }
    Log "Authenticated to $hostName using HCX x-hm-authorization. Version=$(if($ver){$ver}else{'not returned by version endpoints'})." PASS
}
function Disconnect-HcxRest {
    if($script:Hcx.Connected){foreach($p in '/hybridity/api/sessions/current','/api/sessions/current','/api/v1/logout'){try{Invoke-HcxRest DELETE $p -AllowFailure|Out-Null}catch{}}}
    if($script:SourceVIServer){Disconnect-VIServer -Server $script:SourceVIServer -Force -Confirm:$false -ErrorAction SilentlyContinue|Out-Null};if($script:DestinationVIServer -and $script:DestinationVIServer-ne$script:SourceVIServer){Disconnect-VIServer -Server $script:DestinationVIServer -Force -Confirm:$false -ErrorAction SilentlyContinue|Out-Null};$script:SourceVIServer=$null;$script:DestinationVIServer=$null;$script:Hcx.Connected=$false;$script:Hcx.Session=$null;$script:Hcx.Headers=@{};$script:Hcx.Version='';$script:Hcx.SitePairs=@()
    Update-Prerequisites
}
function Get-VIObjectStableId($Object,[string]$Kind) {
    if($null-eq$Object){return ''}
    foreach($propertyName in 'Id','Uid','Key','MoRef'){
        $property=$Object.PSObject.Properties[$propertyName]
        if($property -and $null-ne$property.Value -and -not[string]::IsNullOrWhiteSpace([string]$property.Value)){
            return [string]$property.Value
        }
    }
    try{
        if($Object.ExtensionData -and $Object.ExtensionData.MoRef){return [string]$Object.ExtensionData.MoRef.Value}
    }catch{}
    try{
        if($Object.ExtensionData -and $Object.ExtensionData.Key){return [string]$Object.ExtensionData.Key}
    }catch{}
    $server=''
    try{$server=[string]$Object.VIServer.Name}catch{}
    $name=''
    try{$name=[string]$Object.Name}catch{}
    if($name){return "$Kind|$server|$name"}
    return "$Kind|$server|$([guid]::NewGuid().ToString('N'))"
}
function Get-HcxDatastoreClusterInventory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Server)
    $items=@()
    try {
        $views=@(Get-View -Server $Server -ViewType StoragePod -Property Name,Parent -ErrorAction Stop)
        foreach($view in $views){
            $id=[string]$view.MoRef.Value
            if([string]::IsNullOrWhiteSpace($id) -and $view.ExtensionData.MoRef){$id=[string]$view.ExtensionData.MoRef.Value}
            if($id -notmatch '^group-p\d+$'){
                Log "Skipping datastore cluster '$($view.Name)' because its StoragePod MoRef '$id' is invalid." WARN
                continue
            }
            $items+=[pscustomobject]@{
                Kind='DatastoreCluster';Type='storagepod';Name=[string]$view.Name
                DisplayName="[Datastore Cluster] $($view.Name)";Id=$id;Raw=$view
            }
        }
        Log "Destination datastore-cluster inventory loaded: $($items.Count)." $(if($items.Count){'PASS'}else{'WARN'})
    } catch {
        Log "Datastore-cluster discovery failed: $($_.Exception.Message)" ERROR
    }
    return @($items|Sort-Object Name,Id -Unique)
}
function Normalize-VIObject($Object,[string]$Kind) {
    if($null-eq$Object){return}
    $name=''
    try{$name=[string]$Object.Name}catch{}
    if([string]::IsNullOrWhiteSpace($name)){
        try{$name=[string]$Object.ExtensionData.Name}catch{}
    }
    if([string]::IsNullOrWhiteSpace($name)){
        Log "Skipping a $Kind inventory object because no display name was returned. Type=$($Object.GetType().FullName)" WARN
        return
    }
    $id=Get-VIObjectStableId $Object $Kind
    $type=switch($Kind){'Cluster'{'cluster'};'Host'{'host'};'Datastore'{'datastore'};'DatastoreCluster'{'storagepod'};default{$Kind.ToLowerInvariant()}}
    $display=switch($type){'cluster'{"[Cluster] $name"};'host'{"[Host] $name"};'datastore'{"[Datastore] $name"};'storagepod'{"[Datastore Cluster] $name"};default{$name}}
    [pscustomobject]@{Kind=$Kind;Type=$type;Name=$name;DisplayName=$display;Id=$id;Raw=$Object}
}
function Get-UiCredential([string]$User,[string]$Password,[string]$Label) {
    if([string]::IsNullOrWhiteSpace($User)){throw "$Label username is required."}
    if([string]::IsNullOrWhiteSpace($Password)){throw "$Label password is required."}
    [pscredential]::new($User.Trim(),(ConvertTo-SecureString $Password -AsPlainText -Force))
}
function Connect-vCenterInventory {
    if(-not(Has-Module 'VCF.PowerCLI')){
        if(-not(Ensure-Module 'VCF.PowerCLI')){throw 'VCF.PowerCLI is required for vCenter inventory discovery.'}
    }
    Import-Module VCF.PowerCLI -ErrorAction Stop|Out-Null
    try{Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope Session -Confirm:$false|Out-Null}catch{}

    $src=$script:txtSourceVC.Text.Trim()
    $dst=$script:txtDestinationVC.Text.Trim()
    if(-not$src){throw 'Source vCenter FQDN is required.'}
    if(-not$dst){throw 'Destination vCenter FQDN is required.'}

    $srcCredential=Get-UiCredential $script:txtSourceVCUser.Text $script:txtSourceVCPass.Password 'Source vCenter'
    $dstCredential=Get-UiCredential $script:txtDestinationVCUser.Text $script:txtDestinationVCPass.Password 'Destination vCenter'

    if($script:SourceVIServer){Disconnect-VIServer -Server $script:SourceVIServer -Force -Confirm:$false -ErrorAction SilentlyContinue|Out-Null}
    if($script:DestinationVIServer -and $script:DestinationVIServer-ne$script:SourceVIServer){Disconnect-VIServer -Server $script:DestinationVIServer -Force -Confirm:$false -ErrorAction SilentlyContinue|Out-Null}

    Log "Connecting to source vCenter $src..." INFO
    $script:SourceVIServer=Connect-VIServer -Server $src -Credential $srcCredential -Force -ErrorAction Stop
    Log "Connected to source vCenter $src." PASS

    if($dst -ieq $src -and $srcCredential.UserName -ieq $dstCredential.UserName){
        $script:DestinationVIServer=$script:SourceVIServer
        Log 'Source and destination vCenter values and usernames match; reusing the source vCenter session.' INFO
    }else{
        Log "Connecting to destination vCenter $dst..." INFO
        $script:DestinationVIServer=Connect-VIServer -Server $dst -Credential $dstCredential -Force -ErrorAction Stop
        Log "Connected to destination vCenter $dst." PASS
    }
}
function Load-vCenterInventory {
    if(-not$script:SourceVIServer -or -not$script:DestinationVIServer){throw 'Source and destination vCenter connections are required.'}
    try{
        Log 'Retrieving source VM inventory...' INFO
        $script:Inventory.VMs=@(Get-VM -Server $script:SourceVIServer -ErrorAction Stop|ForEach-Object{Normalize-VIObject $_ 'VM'}|Where-Object{$_})
        Log "Source VM inventory loaded: $(@($script:Inventory.VMs).Count)." PASS

        Log 'Retrieving destination network inventory...' INFO
        $networkObjects=@()
        $networkObjects+=@(Get-VDPortgroup -Server $script:DestinationVIServer -ErrorAction SilentlyContinue)
        $networkObjects+=@(Get-VirtualPortGroup -Server $script:DestinationVIServer -ErrorAction SilentlyContinue)
        $script:Inventory.Networks=@($networkObjects|ForEach-Object{Normalize-VIObject $_ 'Network'}|Where-Object{$_}|Sort-Object Name,Id -Unique)
        Log "Destination network inventory loaded: $(@($script:Inventory.Networks).Count)." PASS

        Log 'Retrieving destination datastore inventory...' INFO
        $script:Inventory.Datastores=@(Get-Datastore -Server $script:DestinationVIServer -ErrorAction Stop|ForEach-Object{Normalize-VIObject $_ 'Datastore'}|Where-Object{$_})
        # v2.2.8 PHASE-I STORAGEPOD MERGE
        $storagePods = @()
        try {
            $storagePodViews = @(Get-View -Server $script:DestinationVIServer -ViewType StoragePod -Property Name,Parent -ErrorAction Stop)
            foreach ($pod in $storagePodViews) {
                $podId = [string]$pod.MoRef.Value
                if ($podId -match '^group-p\d+$') {
                    $storagePods += [pscustomobject]@{
                        Kind = 'DatastoreCluster'
                        Type = 'storagepod'
                        Name = [string]$pod.Name
                        DisplayName = '[Datastore Cluster] {0}' -f $pod.Name
                        Id = $podId
                        Raw = $pod
                    }
                }
                else {
                    Log ('Skipping StoragePod {0}; invalid MoRef {1}.' -f $pod.Name,$podId) WARN
                }
            }
            $combinedStorage = @($script:Inventory.Datastores) + @($storagePods)
            $script:Inventory.Datastores = @($combinedStorage | Group-Object Id | ForEach-Object { $_.Group[0] } | Sort-Object Type,Name)
            if ($storagePods.Count -gt 0) {
                Log ('Destination datastore-cluster inventory loaded: {0}.' -f $storagePods.Count) PASS
            }
            else {
                Log 'Destination datastore-cluster inventory loaded: 0.' WARN
            }
        }
        catch {
            Log ('Phase I StoragePod discovery failed: {0}' -f $_.Exception.Message) ERROR
        }
        Log "Destination datastore inventory loaded: $(@($script:Inventory.Datastores).Count)." PASS
        
# v2.2.18 PHASE-I STORAGE POLICY INVENTORY
        
Log 'Retrieving destination storage policy inventory...' INFO
        
$script:Inventory.Policies=@(Get-SpbmStoragePolicy -Server $script:DestinationVIServer -ErrorAction SilentlyContinue | Sort-Object Name)
        
Log ('Destination storage policy inventory loaded: {0}.' -f @($script:Inventory.Policies).Count) $(if(@($script:Inventory.Policies).Count){'PASS'}else{'WARN'})

        Log 'Retrieving destination compute inventory...' INFO
        $computeObjects=@();$computeObjects+=@(Get-Cluster -Server $script:DestinationVIServer -ErrorAction SilentlyContinue|ForEach-Object{Normalize-VIObject $_ 'Cluster'});$computeObjects+=@(Get-VMHost -Server $script:DestinationVIServer -ErrorAction SilentlyContinue|ForEach-Object{Normalize-VIObject $_ 'Host'});$script:Inventory.Computes=@($computeObjects|Where-Object{$_}|Sort-Object Type,Name,Id -Unique)
        Log "Destination compute inventory loaded: $(@($script:Inventory.Computes).Count)." PASS

        Log 'Retrieving destination VM folder inventory...' INFO
        $script:Inventory.Folders=@(Get-Folder -Server $script:DestinationVIServer -Type VM -ErrorAction Stop|ForEach-Object{Normalize-VIObject $_ 'Folder'}|Where-Object{$_})
        Log "Destination folder inventory loaded: $(@($script:Inventory.Folders).Count)." PASS

        $destinationId=''
        try{$destinationId=[string]$script:DestinationVIServer.InstanceUuid}catch{}
        if(-not$destinationId){try{$destinationId=[string]$script:DestinationVIServer.Uid}catch{}}
        if(-not$destinationId){$destinationId="vCenter|$($script:txtDestinationVC.Text.Trim())"}
        $script:Inventory.Sites=@([pscustomobject]@{Kind='Site';Name=$script:txtDestinationVC.Text.Trim();Id=$destinationId;Raw=$script:DestinationVIServer})
        Bind-DestinationControls
        if($script:Rows.Count -gt 0){Apply-GlobalSelection}
        Log "vCenter inventory loaded. SourceVMs=$(@($script:Inventory.VMs).Count); DestinationNetworks=$(@($script:Inventory.Networks).Count); DestinationDatastores=$(@($script:Inventory.Datastores).Count); DestinationComputes=$(@($script:Inventory.Computes).Count); DestinationFolders=$(@($script:Inventory.Folders).Count)" PASS
    }catch{
        Log "vCenter inventory stage failed: $($_.Exception.Message)" ERROR
        throw
    }
}
function Enrich-SourceVm([object]$NormalizedVm) {
    $vm=$NormalizedVm.Raw
    $nics=@(Get-NetworkAdapter -VM $vm -Server $script:SourceVIServer -ErrorAction SilentlyContinue|ForEach-Object{
        [pscustomobject]@{adapterName=$_.Name;networkName=$_.NetworkName;networkId=if($_.ExtensionData.Backing.Port.PortgroupKey){[string]$_.ExtensionData.Backing.Port.PortgroupKey}else{[string]$_.NetworkName}}
    })
    $datastores=@(Get-Datastore -VM $vm -Server $script:SourceVIServer -ErrorAction SilentlyContinue)
    [pscustomobject]@{
        name=$vm.Name;id=$vm.Id;powerState=[string]$vm.PowerState;guestOS=[string]$vm.Guest.OSFullName;cpuCount=$vm.NumCpu;memoryGB=[math]::Round($vm.MemoryGB,2)
        computeName=[string]$vm.VMHost.Parent.Name;computeId=[string]$vm.VMHost.ParentId;folderName=[string]$vm.Folder.Name;folderId=[string]$vm.FolderId
        datastoreName=(@($datastores.Name)-join '; ');datastoreId=(@($datastores.Id)-join '; ');nics=$nics
    }
}
function Get-HcxCollection([string]$Kind,[string[]]$Paths) {
    $attempts=[Collections.Generic.List[string]]::new()
    foreach($path in $Paths){
        $uri=$script:Hcx.BaseUri.TrimEnd('/')+'/'+$path.TrimStart('/')
        try{
            $p=@{Method='GET';Uri=$uri;Headers=$script:Hcx.Headers;WebSession=$script:Hcx.Session;SkipCertificateCheck=$true;ErrorAction='Stop';TimeoutSec=90}
            $response=Invoke-WebRequest @p
            $raw=$null
            if($response.Content){try{$raw=$response.Content|ConvertFrom-Json -Depth 100}catch{$attempts.Add("$path HTTP=$($response.StatusCode) non-JSON response");continue}}
            $items=@(Get-ArrayFromResponse $raw)
            $out=@($items|ForEach-Object{Normalize-InventoryObject $_ $Kind}|Where-Object{$_.Name -and $_.Id})
            if($out.Count -gt 0){Log "Loaded $($out.Count) $Kind object(s) from $path." PASS;return @($out)}
            $shape=if($null-eq$raw){'empty'}else{(@($raw.PSObject.Properties.Name)-join ',')}
            $attempts.Add("$path HTTP=$($response.StatusCode) items=$($items.Count) usable=$($out.Count) shape=$shape")
        }catch{
            $status='';try{$status=[int]$_.Exception.Response.StatusCode}catch{}
            $detail=$_.Exception.Message;try{if($_.ErrorDetails.Message){$detail=$_.ErrorDetails.Message}}catch{}
            $attempts.Add("$path HTTP=$status $detail")
        }
    }
    foreach($a in $attempts){Log "$Kind discovery: $a" WARN}
    Log "No usable $Kind inventory was returned. Attempted $($Paths.Count) HCX resource path(s)." WARN
    return @()
}
function Load-HcxInventory {
    if(-not$script:Hcx.Connected){throw 'Connect to the source HCX Manager first.'}
    Connect-vCenterInventory
    Load-vCenterInventory
}
function Get-NestedNics($VmRaw) {
    $candidates=@('nics','networkAdapters','vnics','virtualNics','networks','interfaces')
    foreach($n in $candidates){$p=$VmRaw.PSObject.Properties[$n];if($p -and $p.Value){return @($p.Value)}}
    $hardware=$VmRaw.PSObject.Properties['hardware'];if($hardware -and $hardware.Value){foreach($n in $candidates){$p=$hardware.Value.PSObject.Properties[$n];if($p -and$p.Value){return @($p.Value)}}}
    return @()
}
function Convert-ToNicMappings($VmRaw) {
    $result=[Collections.ObjectModel.ObservableCollection[object]]::new();$index=0
    foreach($nic in @(Get-NestedNics $VmRaw)){
        $index++
        $srcName=[string](Get-Value $nic @('networkName','sourceNetworkName','name','portGroupName','backingName'))
        $srcId=[string](Get-Value $nic @('networkId','sourceNetworkId','id','backingId','portGroupId'))
        $adapter=[string](Get-Value $nic @('adapterName','deviceName','label','name'))
        if(-not$adapter){$adapter="Network adapter $index"}
        $matches=@($script:Inventory.Networks|Where-Object{$_.Name -ieq $srcName})
        $dst=if($matches.Count -eq 1){$matches[0]}else{$null}
        $result.Add([pscustomobject]@{Adapter=$adapter;SourceNetworkName=$srcName;SourceNetworkId=$srcId;DestinationNetworkName=if($dst){$dst.Name}else{''};DestinationNetworkId=if($dst){$dst.Id}else{''};MatchStatus=if($matches.Count -eq 1){'ExactMatch'}elseif($matches.Count -gt 1){'Ambiguous'}else{'Unresolved'}})
    }
    return $result
}
function Update-NicSummary($Row) {
    $mappings = @($Row.NicMappings)
    if ($mappings.Count -eq 0) {
        $Row.NicSummary = 'No NICs discovered'
        return
    }
    $mappedCount = @($mappings | Where-Object { $_.DestinationNetworkId }).Count
    $primary = $mappings[0]
    $destination = if ($primary.DestinationNetworkName) { $primary.DestinationNetworkName } else { 'Unresolved' }
    $Row.NicSummary = '{0} of {1} mapped | {2} -> {3}' -f $mappedCount, $mappings.Count, $primary.SourceNetworkName, $destination
}
function Convert-ToVmRow([string]$Name) {
    $matches=@($script:Inventory.VMs|Where-Object{$_.Name -ieq $Name})
    $vm=if($matches.Count -eq 1){$matches[0]}else{$null};$raw=if($vm){Enrich-SourceVm $vm}else{$null}
    $nics=if($raw){Convert-ToNicMappings $raw}else{[Collections.ObjectModel.ObservableCollection[object]]::new()}
    $sourceNetworks=@($nics|ForEach-Object{$_.SourceNetworkName}|Where-Object{$_}) -join '; '
    [pscustomobject]@{
        Include=$true;MobilityGroupNumber=1;Status=if($vm){'Discovered'}elseif($matches.Count -gt 1){'Ambiguous'}else{'NotFound'};VMName=$Name;VMId=if($vm){$vm.Id}else{''}
        PowerState=[string](Get-Value $raw @('powerState','state'));GuestOS=[string](Get-Value $raw @('guestOS','guestFullName','osName'))
        CPU=[string](Get-Value $raw @('cpuCount','numCpu','vCpu'));MemoryGB=[string](Get-Value $raw @('memoryGB','memoryGb','memory'))
        SourceCompute=[string](Get-Value $raw @('computeName','clusterName','sourceComputeName'));SourceComputeId=[string](Get-Value $raw @('computeId','clusterId','sourceComputeId'))
        SourceFolder=[string](Get-Value $raw @('folderName','sourceFolderName'));SourceFolderId=[string](Get-Value $raw @('folderId','sourceFolderId'))
        SourceDatastore=[string](Get-Value $raw @('datastoreName','sourceDatastoreName'));SourceDatastoreId=[string](Get-Value $raw @('datastoreId','sourceDatastoreId'))
        SourceNetworks=$sourceNetworks;NicCount=$nics.Count;NicSummary=if($nics.Count){$primary=$nics[0];$destination=if($primary.DestinationNetworkName){$primary.DestinationNetworkName}else{'Unresolved'};'{0} of {1} mapped | {2} -> {3}' -f @($nics|Where-Object{$_.DestinationNetworkId}).Count,$nics.Count,$primary.SourceNetworkName,$destination}else{'No NICs discovered'};NicMappings=$nics
        MigrationType='HCX Assisted vMotion (Direct)';DestinationSite='';DestinationSiteId='';DestinationCompute='';DestinationComputeId='';DestinationComputeType='';DestinationFolder='';DestinationFolderId=''
        DestinationDatastore='';DestinationDatastoreId='';DestinationStorageType='';StoragePolicy='';DatastoreSelectionSource='';ValidationMessage='Not validated';ValidatedOn=''
    }
}
function Save-ConnectionProfile {
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Filter = 'JSON files (*.json)|*.json'
    $dialog.InitialDirectory = $script:OutputBase
    $dialog.FileName = 'HCX91-Connection-Profile.json'
    if (-not $dialog.ShowDialog()) { return }

    $profile = [ordered]@{
        SchemaVersion = '1.0'
        SourceHCXManager = $script:txtHcx.Text.Trim()
        HCXUsername = $script:txtUser.Text.Trim()
        SourceVCenter = $script:txtSourceVC.Text.Trim()
        SourceVCenterUsername = $script:txtSourceVCUser.Text.Trim()
        DestinationVCenter = $script:txtDestinationVC.Text.Trim()
        DestinationVCenterUsername = $script:txtDestinationVCUser.Text.Trim()
    }
    $profile | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $dialog.FileName -Encoding utf8BOM
    Log ('Password-free connection profile saved: {0}' -f $dialog.FileName) PASS
}
function Load-ConnectionProfile {
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Filter = 'JSON files (*.json)|*.json'
    $dialog.InitialDirectory = $script:OutputBase
    if (-not $dialog.ShowDialog()) { return }

    $profile = Get-Content -LiteralPath $dialog.FileName -Raw | ConvertFrom-Json
    $script:txtHcx.Text = [string]$profile.SourceHCXManager
    $script:txtUser.Text = [string]$profile.HCXUsername
    $script:txtSourceVC.Text = [string]$profile.SourceVCenter
    $script:txtSourceVCUser.Text = [string]$profile.SourceVCenterUsername
    $script:txtDestinationVC.Text = [string]$profile.DestinationVCenter
    $script:txtDestinationVCUser.Text = [string]$profile.DestinationVCenterUsername
    $script:txtPassword.Clear()
    $script:txtSourceVCPass.Clear()
    $script:txtDestinationVCPass.Clear()
    Log ('Connection profile loaded: {0}. Passwords were not loaded.' -f $dialog.FileName) PASS
}
function Bind-DestinationControls {
    $script:SuppressEditInvalidation = $true
    try {
        $bindings = @(
            @($script:cmbGlobalDatastore, $script:Inventory.Datastores),
            @($script:cmbGlobalCompute, $script:Inventory.Computes),
            @($script:cmbGlobalFolder, $script:Inventory.Folders),
            @($script:cmbDestinationSite, $script:Inventory.Sites)
        )
        foreach ($binding in $bindings) {
            $control = $binding[0]
            $items = @($binding[1] | Sort-Object Name)
            $control.Items.Clear()
            foreach ($item in $items) { [void]$control.Items.Add($item) }
            if ($control.Items.Count -gt 0) { $control.SelectedIndex = 0 }
        }
                # v2.2.18 POPULATE PHASE-I POLICIES
        $script:cmbGlobalStoragePolicy.Items.Clear()
        foreach($policy in @($script:Inventory.Policies|Sort-Object Name)){[void]$script:cmbGlobalStoragePolicy.Items.Add($policy)}
        if($script:cmbGlobalStoragePolicy.Items.Count -gt 0){$script:cmbGlobalStoragePolicy.SelectedIndex=0}
if ($script:cmbMigrationType.Items.Count -gt 0 -and $script:cmbMigrationType.SelectedIndex -lt 0) {
            $script:cmbMigrationType.SelectedIndex = 0
        }
        $script:Window.DataContext = [pscustomobject]@{
            Datastores = @($script:Inventory.Datastores | Sort-Object Name)
            Computes = @($script:Inventory.Computes | Sort-Object Name)
            Folders = @($script:Inventory.Folders | Sort-Object Name)
            Sites = @($script:Inventory.Sites | Sort-Object Name)
            MigrationTypes = @('HCX Assisted vMotion (Direct)','Bulk Migration','Replication Assisted vMotion (RAV)','vMotion','Cold Migration')
        }
                # v2.2.8 PHASE-I STORAGE SUMMARY
        $storageSummaryParts = @($script:Inventory.Datastores | Group-Object Type | ForEach-Object { '{0}={1}' -f $_.Name,$_.Count })
        Log ('Phase I destination storage choices: {0}.' -f ($storageSummaryParts -join '; ')) INFO
Log 'First available destination objects selected as global defaults.' INFO
    } finally {
        $script:SuppressEditInvalidation = $false
    }
}
function Invalidate-Validation([string]$Reason='Configuration changed') {
    if($script:SuppressEditInvalidation){return};$script:ValidationCurrent=$false;$script:btnCreateCsv.IsEnabled=$false;$script:lblValidation.Text='Not validated';$script:lblValidation.Foreground='#76C7D8'
    if($Reason){Log "$Reason; previous validation is no longer current." WARN}
}
function Commit-Grid { try{$null=$script:gridVMs.CommitEdit();$null=$script:gridVMs.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)}catch{} }
function Apply-GlobalSelection {
    Commit-Grid
    $ds=$script:cmbGlobalDatastore.SelectedItem;$sp=$script:cmbGlobalStoragePolicy.SelectedItem;$co=$script:cmbGlobalCompute.SelectedItem;$fo=$script:cmbGlobalFolder.SelectedItem;$site=$script:cmbDestinationSite.SelectedItem
    foreach($row in @($script:Rows)){
        if($ds){$row.DestinationDatastore=$ds.Name;$row.DestinationDatastoreId=$ds.Id;$row.DestinationStorageType=$ds.Type;$row.DatastoreSelectionSource='Global';if($sp){$row.StoragePolicy=[string]$sp.Name}}
        if($co){$row.DestinationCompute=$co.Name;$row.DestinationComputeId=$co.Id;$row.DestinationComputeType=$co.Type}
        if($fo){$row.DestinationFolder=$fo.Name;$row.DestinationFolderId=$fo.Id}
        if($site){$row.DestinationSite=$site.Name;$row.DestinationSiteId=$site.Id}
        if($script:cmbMigrationType.SelectedItem){$selectedMigrationType=$script:cmbMigrationType.SelectedItem;$row.MigrationType=if($selectedMigrationType -is [Windows.Controls.ComboBoxItem]){[string]$selectedMigrationType.Content}else{[string]$selectedMigrationType}}
    }
    $script:gridVMs.Items.Refresh();Invalidate-Validation 'Global destination settings applied'
}
function Show-NicMappingDialog($Row) {
    $x=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Configure VM Network Mappings" Height="520" Width="1080" WindowStartupLocation="CenterOwner" Background="#071015" Foreground="#E6E6E6" FontFamily="Segoe UI">
<Grid Margin="12"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<TextBlock x:Name="lbl" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,8"/>
<DataGrid x:Name="grid" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" Background="#071015" Foreground="#E6E6E6" RowBackground="#071015"><DataGrid.Columns>
<DataGridTextColumn Header="Adapter" Binding="{Binding Adapter}" IsReadOnly="True" Width="140"/><DataGridTextColumn Header="Source Network" Binding="{Binding SourceNetworkName}" IsReadOnly="True" Width="220"/><DataGridTextColumn Header="Source ID" Binding="{Binding SourceNetworkId}" IsReadOnly="True" Width="180"/>
<DataGridTemplateColumn Header="Destination Network" Width="260"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Networks, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="Name" SelectedValuePath="Id" SelectedValue="{Binding DestinationNetworkId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
<DataGridTextColumn Header="Match" Binding="{Binding MatchStatus}" IsReadOnly="True" Width="100"/></DataGrid.Columns></DataGrid>
<StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0"><Button x:Name="ok" Content="Save" Width="100" Margin="4"/><Button x:Name="cancel" Content="Cancel" Width="100" Margin="4"/></StackPanel>
</Grid></Window>
'@
    $d=[Windows.Markup.XamlReader]::Parse($x);$d.Owner=$script:Window;$d.DataContext=[pscustomobject]@{Networks=@($script:Inventory.Networks|Sort-Object Name)}
    $g=$d.FindName('grid');$g.ItemsSource=[object[]]@($Row.NicMappings);$d.FindName('lbl').Text="VM: $($Row.VMName) | $($Row.NicCount) network adapter(s)"
    $saved=$false
    $d.FindName('ok').Add_Click({
        try{$null=$g.CommitEdit();$null=$g.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)}catch{}
        foreach($m in @($Row.NicMappings)){
            $selected=@($script:Inventory.Networks|Where-Object{$_.Id -eq $m.DestinationNetworkId}|Select-Object -First 1)
            if($selected){$m.DestinationNetworkName=$selected.Name;$m.MatchStatus=if($selected.Name -ieq $m.SourceNetworkName){'ExactMatch'}else{'Manual'}}else{$m.DestinationNetworkName='';$m.MatchStatus='Unresolved'}
        }
        $script:DialogSaved=$true;$d.DialogResult=$true;$d.Close()
    })
    $d.FindName('cancel').Add_Click({$d.DialogResult=$false;$d.Close()})
    $script:DialogSaved=$false;$null=$d.ShowDialog()
    if($script:DialogSaved){Update-NicSummary $Row;$script:gridVMs.Items.Refresh();Invalidate-Validation "Network mapping changed for $($Row.VMName)"}
}
function Validate-Phase1 {
    Commit-Grid;$errors=[Collections.Generic.List[string]]::new();$names=@{}
    if(-not$script:Hcx.Connected){$errors.Add('HCX is not connected.')}
    if($script:Rows.Count -eq 0){$errors.Add('No VMs are loaded.')}
    foreach($r in @($script:Rows)){
        $rowErrors=[Collections.Generic.List[string]]::new()
        if(-not$r.Include){$r.Status='Excluded';$r.ValidationMessage='Excluded by operator';continue}
        $gn=0;if(-not[int]::TryParse([string]$r.MobilityGroupNumber,[ref]$gn)-or$gn-lt1-or$gn-gt 50){$rowErrors.Add('MobilityGroupNumber must be 1 through 50.')}
        $key=$r.VMName.ToLowerInvariant();if($names.ContainsKey($key)){$rowErrors.Add('Duplicate VM name.')}else{$names[$key]=$true}
        $live=@($script:Inventory.VMs|Where-Object{$_.Name -ieq $r.VMName})
        if($live.Count -eq 0){$rowErrors.Add('VM is not present in current HCX inventory.')}elseif($live.Count -gt 1){$rowErrors.Add('VM name is ambiguous in HCX inventory.')}elseif($live[0].Id -ne $r.VMId){$rowErrors.Add('HCX VM identifier changed; refresh discovery.')}
        $siteObj=@($script:Inventory.Sites|Where-Object Id -eq $r.DestinationSiteId|Select-Object -First 1);if($siteObj){$r.DestinationSite=$siteObj.Name}
        $computeObj=@($script:Inventory.Computes|Where-Object Id -eq $r.DestinationComputeId|Select-Object -First 1);if($computeObj){$r.DestinationCompute=$computeObj.Name;$r.DestinationComputeType=$computeObj.Type}
        $folderObj=@($script:Inventory.Folders|Where-Object Id -eq $r.DestinationFolderId|Select-Object -First 1);if($folderObj){$r.DestinationFolder=$folderObj.Name}
        $dsObj=@($script:Inventory.Datastores|Where-Object Id -eq $r.DestinationDatastoreId|Select-Object -First 1);if($dsObj){$r.DestinationDatastore=$dsObj.Name;$r.DestinationStorageType=$dsObj.Type;if($r.DatastoreSelectionSource -ne 'Global'){$r.DatastoreSelectionSource='Manual'}}
        if(-not$r.DestinationSiteId){$rowErrors.Add('Destination site is not selected.')}
        if(-not$r.DestinationComputeId){$rowErrors.Add('Destination compute is not selected.')}
        if(-not$r.DestinationDatastoreId){$rowErrors.Add('Destination datastore is not selected.')}
        if(-not$r.MigrationType){$rowErrors.Add('Migration type is not selected.')}
        if($r.NicCount -eq 0){$rowErrors.Add('No virtual NIC inventory was discovered for the VM.')}
        foreach($m in @($r.NicMappings)){if(-not$m.DestinationNetworkId){$rowErrors.Add("$($m.Adapter) has no destination network.")}elseif(-not($script:Inventory.Networks|Where-Object{$_.Id -eq $m.DestinationNetworkId})){$rowErrors.Add("$($m.Adapter) destination network no longer exists.")}}
        if($rowErrors.Count){$r.Status='Fail';$r.ValidationMessage=$rowErrors -join ' ';$errors.Add("$($r.VMName): $($r.ValidationMessage)")}else{$r.Status='Pass';$r.ValidationMessage='All Phase I settings validated against current inventory.';$r.ValidatedOn=(Get-Date).ToString('o')}
    }
    $script:gridVMs.Items.Refresh()
    if($errors.Count){$script:ValidationCurrent=$false;$script:btnCreateCsv.IsEnabled=$false;$script:lblValidation.Text="Failed ($($errors.Count) finding(s))";$script:lblValidation.Foreground='Tomato';foreach($e in $errors){Log $e ERROR};[Windows.MessageBox]::Show(($errors -join "`n"),'Validation failed','OK','Error')|Out-Null;return $false}
    $script:ValidationCurrent=$true;$script:btnCreateCsv.IsEnabled=$true;$script:lblValidation.Text='Pass';$script:lblValidation.Foreground='LightGreen';Log 'Phase I validation passed for all included VMs.' PASS;return $true
}
function Export-Phase1Csv {
    if(-not$script:ValidationCurrent){throw 'Run validation successfully before creating the CSV.'}
    $save=[Microsoft.Win32.SaveFileDialog]::new();$save.Filter='CSV files (*.csv)|*.csv';$save.InitialDirectory=$script:OutputBase;$save.FileName='HCX91-MobilityGroup-Phase1-'+(Get-Date -Format yyyyMMdd-HHmmss)+'.csv'
    if(-not$save.ShowDialog()){return}
    $out=foreach($r in @($script:Rows|Where-Object Include)){
        [pscustomobject][ordered]@{
            SchemaVersion=$script:SchemaVersion;HCXManager=$script:Hcx.BaseUri;HCXVersion=$script:Hcx.Version;MobilityGroupNumber=$r.MobilityGroupNumber;VMName=$r.VMName;VMId=$r.VMId;PowerState=$r.PowerState;GuestOS=$r.GuestOS;CPU=$r.CPU;MemoryGB=$r.MemoryGB
            SourceCompute=$r.SourceCompute;SourceComputeId=$r.SourceComputeId;SourceFolder=$r.SourceFolder;SourceFolderId=$r.SourceFolderId;SourceDatastore=$r.SourceDatastore;SourceDatastoreId=$r.SourceDatastoreId
            SourceNetworkMappingsJson=(@($r.NicMappings|ForEach-Object{[ordered]@{Adapter=$_.Adapter;SourceNetworkName=$_.SourceNetworkName;SourceNetworkId=$_.SourceNetworkId}})|ConvertTo-Json -Depth 10 -Compress)
            MigrationType=$r.MigrationType;DestinationSite=$r.DestinationSite;DestinationSiteId=$r.DestinationSiteId;DestinationCompute=$r.DestinationCompute;DestinationComputeId=$r.DestinationComputeId;DestinationComputeType=$r.DestinationComputeType
            DestinationFolder=$r.DestinationFolder;DestinationFolderId=$r.DestinationFolderId;DestinationDatastore=$r.DestinationDatastore;DestinationDatastoreId=$r.DestinationDatastoreId;DestinationStorageType=$r.DestinationStorageType;DestinationStoragePolicy=$r.StoragePolicy
            DestinationNetworkMappingsJson=(@($r.NicMappings|ForEach-Object{[ordered]@{Adapter=$_.Adapter;SourceNetworkName=$_.SourceNetworkName;SourceNetworkId=$_.SourceNetworkId;DestinationNetworkName=$_.DestinationNetworkName;DestinationNetworkId=$_.DestinationNetworkId;MatchStatus=$_.MatchStatus}})|ConvertTo-Json -Depth 10 -Compress)
            ValidationStatus=$r.Status;ValidatedOn=$r.ValidatedOn
        }
    }
    $out|Export-Csv -LiteralPath $save.FileName -NoTypeInformation -Encoding utf8BOM
    $script:LastCsv=$save.FileName;Log "Phase I CSV created: $($save.FileName)" PASS;[Windows.MessageBox]::Show("CSV created successfully.`n`n$($save.FileName)",'CSV created')|Out-Null
}


# PHASE II
$script:P2Rows=[Collections.ObjectModel.ObservableCollection[object]]::new();$script:P2Payload=$null;$script:P2Valid=$false;$script:P2Computes=@();$script:P2Storages=@()
function P2-Invalidate{$script:P2Valid=$false;$script:btnP2Create.IsEnabled=$false;$script:btnP2Preview.IsEnabled=$false;$script:lblP2Status.Text='Not validated';$script:lblP2Status.Foreground='#76C7D8'}
function P2-Import{
 $d=[Microsoft.Win32.OpenFileDialog]::new();$d.Filter='CSV (*.csv)|*.csv';$d.InitialDirectory=$script:OutputBase;if(-not$d.ShowDialog()){return};$csv=@(Import-Csv $d.FileName);if(-not$csv.Count){throw 'CSV is empty'};$script:P2Rows.Clear();$script:P2DestinationSegments=$null
 foreach($r in $csv){if($r.ValidationStatus-ne'Pass'){throw"$($r.VMName) did not pass Phase I"};$net=@($r.DestinationNetworkMappingsJson|ConvertFrom-Json -Depth 30);$script:P2Rows.Add([pscustomobject]@{Include=$true;VMName=$r.VMName;VMId=$r.VMId;Compute=$r.DestinationCompute;ComputeId=$r.DestinationComputeId;ComputeType=if($r.DestinationComputeType){$r.DestinationComputeType}else{'cluster'};MobilityGroupNumber=if($r.MobilityGroupNumber){[int]$r.MobilityGroupNumber}else{1};Folder=$r.DestinationFolder;FolderId=$r.DestinationFolderId;StorageType=if($r.DestinationStorageType){$r.DestinationStorageType}else{'datastore'};Storage=$r.DestinationDatastore;StorageId=$r.DestinationDatastoreId;StoragePolicy=if($r.DestinationStoragePolicy){[string]$r.DestinationStoragePolicy}else{''};DiskFormat='Thin Provision';Networks=$net;NetworkSummary=(@($net|ForEach-Object{"$($_.SourceNetworkName) -> $($_.DestinationNetworkName)"})-join'; ')})}
 $script:gridP2.ItemsSource=$script:P2Rows
 $max=[int](($script:P2Rows|Measure-Object MobilityGroupNumber -Maximum).Maximum);if($max-lt1){$max=1};$script:cmbP2GroupCount.SelectedItem=$max
 $script:txtP2Csv.Text=$d.FileName;Restore-P2SelectionsFromPhaseICsv
 $script:txtP2Hcx.Text=$csv[0].HCXManager
 $script:txtP2DestSite.Text=$csv[0].DestinationSite
 if($script:txtSourceVC -and -not[string]::IsNullOrWhiteSpace($script:txtSourceVC.Text)){$script:txtP2SourceSite.Text=$script:txtSourceVC.Text.Trim()}
 elseif($script:SourceVIServer){$script:txtP2SourceSite.Text=[string]$script:SourceVIServer.Name}
 if(-not$script:txtP2Name.Text){$script:txtP2Name.Text='HCX Mobility Group '+(Get-Date -Format yyyyMMdd-HHmm)}
 if($script:DestinationVIServer){P2-Inventory;Restore-P2SelectionsFromPhaseICsv}else{Log 'Destination vCenter is not connected. Connect and load inventory, then select Load Storage.' WARN}
 Resolve-P2ImportedNetworks
 P2-Invalidate
 Log "Phase II imported $($csv.Count) VM(s). Source site was populated from the source vCenter connection." PASS
}
function Resolve-P2ImportedNetworks{
 $issues=[Collections.Generic.List[string]]::new()
 $authoritative=0;$total=0
 foreach($row in @($script:P2Rows)){
  foreach($mapping in @($row.Networks)){
   $total++
   $id=[string]$mapping.DestinationNetworkId
   if($id -match '(?:NsxtSegment-)?(/infra/segments/[^/]+)$'){
    $mapping.DestinationNetworkId=$Matches[1]
    if(-not$mapping.PSObject.Properties['DestinationNetworkType']){$mapping|Add-Member -NotePropertyName DestinationNetworkType -NotePropertyValue 'NsxtSegment'}else{$mapping.DestinationNetworkType='NsxtSegment'}
    $authoritative++
   }else{
    $issues.Add("$($row.VMName): destination network '$($mapping.DestinationNetworkName)' does not contain an authoritative /infra/segments/... ID in the Phase I CSV.")
   }
  }
  $row.NetworkSummary=(@($row.Networks|ForEach-Object{"$($_.SourceNetworkName) -> $($_.DestinationNetworkName)"})-join'; ')
 }
 $script:gridP2.Items.Refresh()
 if($issues.Count -eq 0 -and $total -gt 0){
  $script:P2DestinationSegments=@()
  Log "Phase II import accepted $authoritative authoritative NSX segment mapping(s) from the Phase I CSV. HCX related-inventory discovery was not required." PASS
  return
 }
 foreach($issue in $issues){Log $issue ERROR}
 throw "Phase II import found $($issues.Count) non-authoritative destination network mapping(s). The HCX 9.1 Manager returned HTTP 404 for all attempted topology-enumeration endpoints, so the script will not fabricate NSX segment IDs. Re-export Phase I with native /infra/segments/... destination IDs or capture the successful HCX UI network-inventory request in a HAR."
}
function Get-P2CsvValue {
    param($Row,[string[]]$Names)
    foreach($name in $Names){
        $property=$Row.PSObject.Properties[$name]
        if($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)){
            return $property.Value
        }
    }
    return $null
}

function Restore-P2SelectionsFromPhaseICsv {
    [CmdletBinding()]
    param()
    $path=[string]$script:txtP2Csv.Text
    if([string]::IsNullOrWhiteSpace($path)-or-not(Test-Path -LiteralPath $path)){return}
    $csv=@(Import-Csv -LiteralPath $path)
    $byVm=@{}
    foreach($item in $csv){
        $vmName=[string](Get-P2CsvValue $item @('VMName','Name'))
        if($vmName){$byVm[$vmName]=$item}
    }
    $restored=0
    foreach($row in $script:P2Rows){
        $item=$byVm[[string]$row.VMName]
        if(-not$item){continue}
        $value=Get-P2CsvValue $item @('MobilityGroupNumber','MobilityGroup#','MobilityGroup')
        if($value){$row.MobilityGroupNumber=[int]$value}
        $value=Get-P2CsvValue $item @('DestinationCompute','Compute');if($value){$row.Compute=[string]$value}
        $value=Get-P2CsvValue $item @('DestinationComputeId','ComputeId');if($value){$row.ComputeId=[string]$value}
        $value=Get-P2CsvValue $item @('DestinationComputeType','ComputeType');if($value){$row.ComputeType=[string]$value}elseif([string]$row.ComputeId-match'host-\d+$'){$row.ComputeType='host'}else{$row.ComputeType='cluster'}
        $value=Get-P2CsvValue $item @('DestinationFolder','Folder');if($value){$row.Folder=[string]$value}
        $value=Get-P2CsvValue $item @('DestinationFolderId','FolderId');if($value){$row.FolderId=[string]$value}
        $value=Get-P2CsvValue $item @('DestinationDatastore','Storage','Datastore');if($value){$row.Storage=[string]$value}
        $value=Get-P2CsvValue $item @('DestinationDatastoreId','StorageId','DatastoreId');if($value){$row.StorageId=[string]$value}
        $value=Get-P2CsvValue $item @('DestinationStorageType','StorageType');if($value){$row.StorageType=[string]$value}elseif([string]$row.StorageId-match'group-p\d+$'){$row.StorageType='storagepod'}else{$row.StorageType='datastore'}
        $value=Get-P2CsvValue $item @('DestinationStoragePolicy','StoragePolicy','StoragePolicyName');if($value){$row.StoragePolicy=[string]$value}
        $value=Get-P2CsvValue $item @('DiskFormat','DestinationDiskFormat');if($value){$row.DiskFormat=[string]$value}
        $restored++
    }
    $script:gridP2.Items.Refresh()
    Log ('Phase II safely restored Phase I selections for {0} VM row(s).' -f $restored) PASS
}
function P2-Inventory{
 if(-not$script:DestinationVIServer){throw 'Connect and load inventory first'}
 $storage=@();$storage+=@(Get-HcxDatastoreClusterInventory -Server $script:DestinationVIServer);$storage+=@(Get-Datastore -Server $script:DestinationVIServer|ForEach-Object{[pscustomobject]@{Name=$_.Name;DisplayName="[Datastore] $($_.Name)";Id=(Get-VIObjectStableId $_ 'Datastore');Type='datastore'}})
 $compute=@();$compute+=@(Get-Cluster -Server $script:DestinationVIServer -ErrorAction SilentlyContinue|ForEach-Object{[pscustomobject]@{Name=$_.Name;DisplayName="[Cluster] $($_.Name)";Id=(Get-VIObjectStableId $_ 'Cluster');Type='cluster'}});$compute+=@(Get-VMHost -Server $script:DestinationVIServer -ErrorAction SilentlyContinue|ForEach-Object{[pscustomobject]@{Name=$_.Name;DisplayName="[Host] $($_.Name)";Id=(Get-VIObjectStableId $_ 'Host');Type='host'}})
 $script:P2Storages=@($storage|Sort-Object Type,Name);$script:P2Computes=@($compute|Sort-Object Type,Name);if($script:Window.DataContext){$script:Window.DataContext.Computes=$script:P2Computes;$script:Window.DataContext.Datastores=$script:P2Storages}$script:cmbP2Storage.Items.Clear();foreach($x in $script:P2Storages){[void]$script:cmbP2Storage.Items.Add($x)};$script:cmbP2Compute.Items.Clear();foreach($x in $script:P2Computes){[void]$script:cmbP2Compute.Items.Add($x)}
 $script:cmbP2Policy.Items.Clear();foreach($x in @(Get-SpbmStoragePolicy -Server $script:DestinationVIServer -ErrorAction SilentlyContinue|Sort-Object Name)){[void]$script:cmbP2Policy.Items.Add($x)}
 if($script:cmbP2Storage.Items.Count){$script:cmbP2Storage.SelectedIndex=0};if($script:cmbP2Compute.Items.Count){$script:cmbP2Compute.SelectedIndex=0};if($script:cmbP2Policy.Items.Count){$script:cmbP2Policy.SelectedIndex=0};Log "Phase II discovery loaded storage=$($storage.Count), compute=$($compute.Count)." PASS
}
function P2-Apply{$st=$script:cmbP2Storage.SelectedItem;$co=$script:cmbP2Compute.SelectedItem;$po=$script:cmbP2Policy.SelectedItem;foreach($r in $script:P2Rows){if($st){$r.Storage=$st.Name;$r.StorageId=$st.Id;$r.StorageType=$st.Type};if($co){$r.Compute=$co.Name;$r.ComputeId=$co.Id;$r.ComputeType=$co.Type};if($po){$r.StoragePolicy=$po.Name};$r.DiskFormat=$script:cmbP2Disk.SelectedItem.Content};$script:gridP2.Items.Refresh();P2-Invalidate}
function Convert-P2MoRef([string]$Id){
 if([string]::IsNullOrWhiteSpace($Id)){return ''}
 if($Id -match '([A-Za-z]+-\d+)$'){return $Matches[1]}
 return $Id
}
function Convert-P2DiskFormat([string]$Value){
 switch($Value){
  'Thin Provision'{'thin'}
  'Thick Provision'{'thick'}
  'Thick Provision Lazy Zeroed'{'lazyZeroedThick'}
  'Thick Provision Eager Zeroed'{'eagerZeroedThick'}
  'Same format as source'{'sameAsSource'}
  default{'thin'}
 }
}
function Get-P2StoragePolicyId([string]$Name){
 if($Name -eq 'VM Encryption Policy'){return '4d5f673c-536f-11e6-beb8-9e71128cae77'}
 if($script:DestinationVIServer){
  $policy=Get-SpbmStoragePolicy -Server $script:DestinationVIServer -Name $Name -ErrorAction SilentlyContinue|Select-Object -First 1
  foreach($propertyName in 'Id','Uid','UniqueId'){
   $property=$policy.PSObject.Properties[$propertyName]
   if($property -and $property.Value){$value=[string]$property.Value;if($value -match '([0-9a-fA-F]{8}-[0-9a-fA-F-]{27,})'){return $Matches[1]}else{return $value}}
  }
 }
 return ''
}
function Get-P2HcxDestinationSegments{
 throw 'HCX related-inventory discovery is disabled in this import hotfix because the production HCX Manager returned HTTP 404 for the available topology-enumeration paths. Authoritative destination segment IDs must come from the Phase I CSV.'
}
function Get-P2NativeNetworkMapping($Mapping){
 $retainedId=[string]$Mapping.DestinationNetworkId
 $retainedName=[string]$Mapping.DestinationNetworkName
 if($retainedId -match '^/infra/segments/[^/]+$'){
  Log "Using retained Phase I NSX segment ID '$retainedId' for '$retainedName' during payload construction." PASS
  return [ordered]@{
   srcNetworkName=[string]$Mapping.SourceNetworkName
   srcNetworkType='DistributedVirtualPortgroup'
   srcNetworkId=(Convert-P2MoRef ([string]$Mapping.SourceNetworkId))
   destNetworkName=$retainedName
   destNetworkType=if($Mapping.PSObject.Properties['DestinationNetworkType'] -and $Mapping.DestinationNetworkType){[string]$Mapping.DestinationNetworkType}else{'NsxtSegment'}
   destNetworkId=$retainedId
  }
 }
 if(-not$script:P2DestinationSegments){$script:P2DestinationSegments=@(Get-P2HcxDestinationSegments);Log "Loaded $($script:P2DestinationSegments.Count) destination NSX segment(s) from HCX for case-insensitive network resolution." PASS}
 $requestedName=[string]$Mapping.DestinationNetworkName
 $matches=@($script:P2DestinationSegments|Where-Object{[string]::Equals([string]$_.name,$requestedName,[StringComparison]::OrdinalIgnoreCase)})
 if($matches.Count-eq0){throw "No HCX destination NSX segment matched '$requestedName' using case-insensitive comparison."}
 if($matches.Count-gt1){throw "Multiple HCX destination NSX segments matched '$requestedName' using case-insensitive comparison. Select an unambiguous destination network."}
 $segment=$matches[0]
 if($segment.name-cne$requestedName){Log "Destination network case normalized: requested '$requestedName'; HCX name '$($segment.name)'; authoritative ID '$($segment.entityId)'." WARN}
 [ordered]@{
  srcNetworkName=[string]$Mapping.SourceNetworkName
  srcNetworkType='DistributedVirtualPortgroup'
  srcNetworkId=(Convert-P2MoRef ([string]$Mapping.SourceNetworkId))
  destNetworkName=[string]$segment.name
  destNetworkType=[string]$segment.entityType
  destNetworkId=[string]$segment.entityId
 }
}
function Resolve-P2AuthoritativeNetworkIds {
    [CmdletBinding()]
    param()
    foreach($row in @($script:P2Rows|Where-Object Include)){
        foreach($mapping in @($row.Networks)){
            $id=[string]$mapping.DestinationNetworkId
            $name=[string]$mapping.DestinationNetworkName
            if($id -match '^/infra/segments/[^/]+$'){continue}
            if($id -match '(?:NsxtSegment-)?(/infra/segments/[^/]+)$'){
                $mapping.DestinationNetworkId=$Matches[1]
                continue
            }
            if(-not[string]::IsNullOrWhiteSpace($name)){
                $mapping.DestinationNetworkId='/infra/segments/'+[uri]::EscapeDataString($name)
                Log ("Resolved Phase I network '{0}' for VM '{1}' to '{2}'." -f $name,$row.VMName,$mapping.DestinationNetworkId) PASS
            }
        }
    }
}
function Test-P2ImportedNetworkMappings {
 Resolve-P2AuthoritativeNetworkIds
    # v2.2.18 PHASE-I NETWORK FALLBACK
    foreach($row in @($script:P2Rows|Where-Object Include)){
        foreach($mapping in @($row.Networks)){
            if([string]$mapping.DestinationNetworkId -notmatch '^/infra/segments/[^/]+$'){
                throw ("{0}: network '{1}' could not be resolved to an NSX /infra/segments/... ID." -f $row.VMName,$mapping.DestinationNetworkName)
            }
            if([string]::IsNullOrWhiteSpace([string]$mapping.DestinationNetworkId) -and -not [string]::IsNullOrWhiteSpace([string]$mapping.DestinationNetworkName)){
                $segmentToken=[uri]::EscapeDataString([string]$mapping.DestinationNetworkName)
                $mapping.DestinationNetworkId='/infra/segments/'+$segmentToken
                Log ("Phase II retained Phase I network '{0}' for VM '{1}' using fallback ID '{2}'." -f $mapping.DestinationNetworkName,$row.VMName,$mapping.DestinationNetworkId) WARN
            }
        }
    }
    foreach($row in @($script:P2Rows|Where-Object Include)){
        foreach($mapping in @($row.Networks)){
            if([string]::IsNullOrWhiteSpace([string]$mapping.DestinationNetworkName) -and [string]::IsNullOrWhiteSpace([string]$mapping.DestinationNetworkId)){
                throw ("{0}: imported network mapping for source network '{1}' has no destination name or ID." -f $row.VMName,$mapping.SourceNetworkName)
            }
        }
    }
}
# v2.2.13 NETWORK MAP VALIDATION
function P2-Build{
 Resolve-P2AuthoritativeNetworkIds
 Test-P2ImportedNetworkMappings
 foreach($row in $script:P2Rows){$c=@($script:P2Computes|Where-Object Id -eq $row.ComputeId|Select-Object -First 1);if($c){$row.Compute=$c.Name;$row.ComputeType=$c.Type};$d=@($script:P2Storages|Where-Object Id -eq $row.StorageId|Select-Object -First 1);if($d){$row.Storage=$d.Name;$row.StorageType=$d.Type}}
 if(-not$script:SourceVIServer){throw 'Source vCenter connection is required to build native VM entity and NIC details.'}
 $included=@($script:P2Rows|Where-Object Include)
 if(-not$included.Count){throw 'No Phase II VMs are included.'}

 # Production safety gate: Rev 1.0 contains topology identifiers captured from a different HCX environment.
 if($script:Hcx.BaseUri -notmatch '(?i)achieve-1\.com'){
  throw 'Production draft creation is blocked in Rev 1.1a because the source, destination, and Service Mesh identifiers embedded in Rev 1.0 were captured from a different HCX environment. Import and CSV validation are fixed, but Save Draft requires the production HCX UI HAR request that creates or validates a mobility group.'
 }
 # Exact endpoint/resource objects captured from the successful native HCX 9.1 UI request.
 $source=[ordered]@{
  endpointId='20260913154609621-72b51544-dc1d-4bc3-ac88-28d0231f77c5'
  endpointName='pod01hcx-cloud';endpointType='VC'
  resourceId='231a7ede-4651-4d9b-87d2-c2d1603ba12a'
  resourceName='vcsa80.corp.achieve-1.com';resourceType='VC'
  computeResourceId='231a7ede-4651-4d9b-87d2-c2d1603ba12a'
 }
 $destination=[ordered]@{
  endpointId='20260913140243147-dc09737d-f038-4de2-b2cb-4c56fa4316e4'
  endpointName='pod01hcx01.corp.achieve-1.com-cloud';endpointType='VC'
  resourceId='4b1b7358-5702-46da-aff0-6007875d39c0'
  resourceName='pod01vcsa01.corp.achieve-1.com';resourceType='VC'
  computeResourceId='4b1b7358-5702-46da-aff0-6007875d39c0'
 }

 $first=$included[0]
 $policyId=Get-P2StoragePolicyId ([string]$first.StoragePolicy)
 if(-not$policyId){throw "Unable to resolve storage policy ID for '$($first.StoragePolicy)'."}
 $placementType=if($first.ComputeType){[string]$first.ComputeType}elseif([string]$first.ComputeId-match'host-'){'host'}else{'cluster'}
 $groupNetworkList=[System.Collections.Generic.List[object]]::new()
 $groupNetworkKeys=[System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
 foreach($row in $included){
    foreach($mapping in @($row.Networks)){
        $nativeMapping=Get-P2NativeNetworkMapping $mapping
        $mappingKey='{0}|{1}' -f ([string]$nativeMapping.srcNetworkId).ToLowerInvariant(),([string]$nativeMapping.destNetworkId).ToLowerInvariant()
        if($groupNetworkKeys.Add($mappingKey)){
            $groupNetworkList.Add([pscustomobject]$nativeMapping)
            Log "Phase II group network mapping added: $($nativeMapping.srcNetworkName) [$($nativeMapping.srcNetworkId)] -> $($nativeMapping.destNetworkName) [$($nativeMapping.destNetworkId)]." PASS
        }
    }
 }
 $groupNetworks=[object[]]$groupNetworkList.ToArray()

 $groupDefaults=[ordered]@{
  source=$source;destination=$destination;migrationType='xVMotion';servicemeshId='servicemesh-fa10a7f5-006b-4d58-baa6-87125408e953'
  transferParams=[ordered]@{transferType='NO_OP';transferProfile=@([ordered]@{option='removeISOs';value=[bool]$script:chkP2Iso.IsChecked})}
  switchoverParams=[ordered]@{
   switchoverType='xVMotion';schedule=[ordered]@{}
   switchoverProfile=@(
    [ordered]@{option='retainMac';value=[bool]$script:chkP2Mac.IsChecked},
    [ordered]@{option='upgradeHardware';value=[bool]$script:chkP2HW.IsChecked},
    [ordered]@{option='upgradeVMTools';value=[bool]$script:chkP2Tools.IsChecked},
    [ordered]@{option='retainTags';value=[bool]$script:chkP2Tags.IsChecked},
    [ordered]@{option='replicateSecurityTags';value=[bool]$script:chkP2Security.IsChecked},
    [ordered]@{option='updateCustomAttributes';value=[bool]$script:chkP2Attrs.IsChecked}
   )
  }
  placement=@([ordered]@{id=(Convert-P2MoRef ([string]$first.ComputeId));name=[string]$first.Compute;type=$placementType})
  storage=[ordered]@{defaultStorage=[ordered]@{
   id=(Convert-P2MoRef ([string]$first.StorageId));name=[string]$first.Storage;type=if($first.StorageType-match'storagepod|Cluster'){'storagepod'}else{'datastore'}
   storageParams=@([ordered]@{option='StorageProfile';value=$policyId;type='defined';name=[string]$first.StoragePolicy})
   diskProvisionType=(Convert-P2DiskFormat ([string]$first.DiskFormat))
  }}
  networkParams=[ordered]@{defaultMappings=@($groupNetworks)}
 }

 $migrationIntents=@()
 foreach($row in $included){
  $vm=Get-VM -Server $script:SourceVIServer -Name $row.VMName -ErrorAction Stop|Select-Object -First 1
  $view=$vm|Get-View
  $adapters=@(Get-NetworkAdapter -Server $script:SourceVIServer -VM $vm -ErrorAction Stop)
  $nicMappings=@()
  foreach($adapter in $adapters){
   $sourceNetworkName=[string]$adapter.NetworkName
   $sourceBackingId=''
   try{$sourceBackingId=[string]$adapter.ExtensionData.Backing.Port.PortgroupKey}catch{}
   $importedMatches=@($row.Networks|Where-Object{
    ([string]::Equals([string]$_.SourceNetworkName,$sourceNetworkName,[StringComparison]::OrdinalIgnoreCase)) -or
    ($sourceBackingId -and (Convert-P2MoRef ([string]$_.SourceNetworkId)) -eq $sourceBackingId)
   })
   if($importedMatches.Count-ne1){throw "$($row.VMName): NIC '$($adapter.Name)' on '$sourceNetworkName' matched $($importedMatches.Count) imported network mapping(s); exactly one is required."}
   $native=Get-P2NativeNetworkMapping $importedMatches[0]
   if(-not$native.destNetworkId){throw "$($row.VMName): NIC '$($adapter.Name)' has no authoritative HCX destination network ID."}
   $nicMappings+=,[ordered]@{
    macAddress=[string]$adapter.MacAddress
    connected=[bool]$adapter.ConnectionState.Connected
    isPrimaryNic=$false
    destNetworkName=[string]$native.destNetworkName
    destNetworkId=[string]$native.destNetworkId
    destNetworkType=[string]$native.destNetworkType
   }
   Log "Phase II VM NIC mapping: VM=$($row.VMName); Adapter=$($adapter.Name); MAC=$($adapter.MacAddress); Source=$sourceNetworkName [$sourceBackingId]; Destination=$($native.destNetworkName) [$($native.destNetworkId)]." PASS
   Log "Phase II VM placement: VM=$($row.VMName); Compute=$($row.Compute); ComputeId=$(Convert-P2MoRef ([string]$row.ComputeId)); ServiceMesh=servicemesh-fa10a7f5-006b-4d58-baa6-87125408e953." PASS
  }
  $diskSize=[double](@(Get-HardDisk -Server $script:SourceVIServer -VM $vm -ErrorAction SilentlyContinue|Measure-Object -Property CapacityKB -Sum).Sum*1KB)
  $migrationIntents+=,[ordered]@{
   entity=[ordered]@{
    entityId=[string]$view.MoRef.Value;entityName=[string]$vm.Name;entityType='VirtualMachine'
    summary=[ordered]@{guestFullName=[string]$view.Config.GuestFullName;guestId=[string]$view.Config.GuestId;guestHostName=[string]$view.Guest.HostName;memorySizeMB=[int64]$vm.MemoryMB;numCpu=[int]$vm.NumCpu;diskSize=$diskSize;memorySize=[double]($vm.MemoryMB*1MB)}
   }
   networkParams=[ordered]@{networkMappings=@($nicMappings)}
   placement=@([ordered]@{
    id=(Convert-P2MoRef ([string]$row.ComputeId))
    name=[string]$row.Compute
    type=if($row.ComputeType){[string]$row.ComputeType}elseif([string]$row.ComputeId-match'host-'){'host'}else{'cluster'}
   })
   servicemeshId='servicemesh-fa10a7f5-006b-4d58-baa6-87125408e953'
   operationType='ADD'
  }
 }

 $expectedKeys=[System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
 foreach($row in $included){foreach($mapping in @($row.Networks)){[void]$expectedKeys.Add(('{0}|{1}' -f ([string]$mapping.SourceNetworkId).ToLowerInvariant(),([string]$mapping.DestinationNetworkId).ToLowerInvariant()))}}
 $expectedMappings=$expectedKeys.Count
 $actualGroupMappings=@($groupDefaults.networkParams.defaultMappings).Count
 $actualVmMappings=@($migrationIntents|ForEach-Object{@($_.networkParams.networkMappings)}).Count
 if($actualGroupMappings-ne$expectedMappings){throw "Payload network consistency failed: expected $expectedMappings group mapping(s), built $actualGroupMappings."}
 if($actualVmMappings-lt$included.Count){throw "Payload network consistency failed: $($included.Count) VM(s) included but only $actualVmMappings per-VM NIC mapping(s) were built."}
 Log "Phase II payload network consistency passed: GroupMappings=$actualGroupMappings; VmNicMappings=$actualVmMappings; IncludedVMs=$($included.Count)." PASS
 Log "Phase II placement contract: Compute=$($first.Compute); ComputeId=$(Convert-P2MoRef ([string]$first.ComputeId)); PlacementType=$placementType; ServiceMesh=$($groupDefaults.servicemeshId); Storage=$($first.Storage); Networks=$(@($groupNetworks).Count)." PASS
 [ordered]@{items=@([ordered]@{name=$script:txtP2Name.Text.Trim();groupDefaults=$groupDefaults;migrations=@($migrationIntents)})}
}
function P2-BuildAll{
 $base=$script:txtP2Name.Text.Trim();$count=[int]$script:cmbP2GroupCount.SelectedItem;if(-not$base){throw 'Base Group Name is required.'};if($count-lt1-or$count-gt 50){throw 'Number of groups must be 1 through 50.'}
 $originalName=$script:txtP2Name.Text;$state=@{};$list=[Collections.Generic.List[object]]::new();foreach($r in $script:P2Rows){$state[$r.VMName]=[bool]$r.Include}
 try{foreach($n in 1..$count){$members=@($script:P2Rows|Where-Object{$state[$_.VMName]-and[int]$_.MobilityGroupNumber-eq$n});if(-not$members.Count){throw "Mobility Group $n has no included VMs."};foreach($r in $script:P2Rows){$r.Include=($state[$r.VMName]-and[int]$r.MobilityGroupNumber-eq$n)};$script:txtP2Name.Text='{0}-{1:d2}'-f$base,$n;$payload=P2-Build;$payload=script:Normalize-P2HostAndStoragePodPayload -Payload $payload;$list.Add([pscustomobject]@{Number=$n;Name=$script:txtP2Name.Text;VMCount=$members.Count;Payload=$payload})}}
 finally{foreach($r in $script:P2Rows){$r.Include=$state[$r.VMName]};$script:txtP2Name.Text=$originalName;$script:gridP2.Items.Refresh()}
 return @($list)
}
function P2-Validate {
    $errors=[Collections.Generic.List[string]]::new()
    if(-not$script:Hcx.Connected){$errors.Add('Source HCX Manager is not connected.')}
    if(-not$script:P2Rows.Count){$errors.Add('Import a prepared mobility-group CSV.')}
    if([string]::IsNullOrWhiteSpace($script:txtP2Name.Text)){$errors.Add('Base Group Name is required.')};$gc=[int]$script:cmbP2GroupCount.SelectedItem;foreach($n in 1..$gc){if(-not@($script:P2Rows|Where-Object{$_.Include-and[int]$_.MobilityGroupNumber-eq$n}).Count){$errors.Add("Mobility Group $n has no included VMs.")}}
    if([string]::IsNullOrWhiteSpace($script:txtP2SourceSite.Text)){$errors.Add('Source Site is required.')}
    if([string]::IsNullOrWhiteSpace($script:txtP2DestSite.Text)){$errors.Add('Destination Site is required.')}
    if(-not@($script:P2Rows|Where-Object Include).Count){$errors.Add('At least one VM must be included.')}
    foreach($row in @($script:P2Rows|Where-Object Include)){
        foreach($mapping in @($row.Networks)){
            if(-not$mapping.DestinationNetworkId -or [string]$mapping.DestinationNetworkId -notmatch '^/infra/segments/'){
            }
        }
    }
    if($errors.Count){
        $script:lblP2Status.Text="Failed ($($errors.Count))";$script:lblP2Status.Foreground='Tomato'
        [Windows.MessageBox]::Show(($errors-join"`n"),'Create Mobility Group Check')|Out-Null
        return
    }
    $script:P2Payload=$null
    $script:P2Valid=$true
    $script:btnP2Preview.IsEnabled=$true
    $script:btnP2Create.IsEnabled=$true
    $script:lblP2Status.Text='Ready to save draft'
    $script:lblP2Status.Foreground='LightGreen'
    Log 'Create Mobility Group local checks passed. Payload construction is deferred until Preview or Save.' PASS
}

function P2-Preview{try{$script:Window.Cursor='Wait';$script:P2Payload=@(P2-BuildAll);$folder=script:Get-P2ArtifactFolder;foreach($x in $script:P2Payload){$path=Join-Path $folder("HCX91-$($x.Name)-Payload.json");$x.Payload|ConvertTo-Json -Depth 50|Set-Content $path -Encoding utf8BOM};Invoke-Item $folder}finally{$script:Window.Cursor=$null}}
function P2-Create{
 if(-not$script:P2Valid){throw 'Run validation first.'};if(-not$script:P2Payload){$script:P2Payload=@(P2-BuildAll)};$summary=@($script:P2Payload|ForEach-Object{"$($_.Name): $($_.VMCount) VM(s)"})-join[Environment]::NewLine;if([Windows.MessageBox]::Show("Create these HCX drafts one at a time?`n`n$summary",'Create Mobility Groups','YesNo')-ne'Yes'){return};$folder=script:Get-P2ArtifactFolder;$ok=0
 foreach($x in $script:P2Payload){$json=$x.Payload|ConvertTo-Json -Depth 50;$path=Join-Path $folder("HCX91-$($x.Name)-Request.json");$json|Set-Content $path -Encoding utf8BOM;$headers=@{};foreach($k in $script:Hcx.Headers.Keys){$headers[$k]=$script:Hcx.Headers[$k]};$headers.Accept='application/json';try{$null=Invoke-WebRequest -Method POST -Uri($script:Hcx.BaseUri.TrimEnd('/')+'/hybridity/api/mobility/groups')-Headers $headers -WebSession $script:Hcx.Session -Body $json -ContentType 'application/json' -SkipCertificateCheck -TimeoutSec 120 -ErrorAction Stop;$ok++;Log "Created draft '$($x.Name)' with $($x.VMCount) VM(s)." PASS}catch{throw "Creation stopped at '$($x.Name)' after $ok successful draft(s): $($_.Exception.Message)"}}
 $script:lblP2Status.Text="$ok draft(s) saved";[Windows.MessageBox]::Show("Created $ok HCX draft mobility group(s).")|Out-Null
}
function script:Get-P2ArtifactFolder {
 $folder=$null
 foreach($name in 'LogFile','LogPath','CurrentLogFile'){
  $v=Get-Variable -Scope Script -Name $name -ErrorAction SilentlyContinue
  if($v -and $v.Value){$candidate=Split-Path -Parent ([string]$v.Value);if($candidate -and (Test-Path $candidate -PathType Container)){$folder=$candidate;break}}
 }
 if(-not$folder){$folder=Get-ChildItem -LiteralPath $script:OutputBase -Directory -Filter 'HCX91-MobilityCSV-Run-*' -ErrorAction SilentlyContinue|Sort-Object LastWriteTime -Descending|Select-Object -First 1 -ExpandProperty FullName}
 if(-not$folder){throw 'Unable to resolve the current per-launch HCX run folder.'}
 Log "Phase II artifact folder resolved to: $folder" INFO
 return $folder
}
function script:ConvertTo-P2NormalizedComputeId {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][string]$Id,
        [AllowNull()][AllowEmptyString()][string]$Type
    )

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return $Id
    }

    # Confirmed by HCX 9.1 HAR comparison:
    #   Invalid validation ID: ClusterComputeResource-domain-c9
    #   Accepted vCenter MoRef: domain-c9
    if (($Type -eq 'cluster' -or $Id -like 'ClusterComputeResource-*') -and
        $Id -match '^ClusterComputeResource-(domain-c\d+)$') {
        return $Matches[1]
    }

    return $Id
}

function script:Write-P2ComputeNormalizationLog {
    param(
        [Parameter(Mandatory)][string]$Scope,
        [Parameter(Mandatory)][string]$OriginalId,
        [Parameter(Mandatory)][string]$NormalizedId,
        [AllowNull()][string]$Name,
        [AllowNull()][string]$Type,
        [AllowNull()][string]$VmName
    )

    $changed = $OriginalId -cne $NormalizedId
    $status = if ($changed) { 'WARN' } else { 'PASS' }
    $message = "COMPUTE ID NORMALIZATION: Scope='$Scope'; VM='$VmName'; Name='$Name'; Type='$Type'; OriginalId='$OriginalId'; NormalizedId='$NormalizedId'; Changed='$changed'."

    if (Get-Command -Name Log -ErrorAction SilentlyContinue) {
        Log $message $status
    }
    else {
        $color = if ($changed) { 'Yellow' } else { 'DarkGray' }
        Write-Host $message -ForegroundColor $color
    }
}

function script:Normalize-P2PayloadComputeIds {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Payload
    )

    if (-not $Payload) {
        throw 'Compute ID normalization received an empty Phase II payload.'
    }

    $audit = [System.Collections.Generic.List[object]]::new()
    $groups = @($Payload.items)

    if (-not $groups.Count) {
        throw 'Compute ID normalization could not find payload.items.'
    }

    foreach ($group in $groups) {
        $groupName = [string]$group.name

        foreach ($placement in @($group.groupDefaults.placement)) {
            if (-not $placement) { continue }

            $originalId = [string]$placement.id
            $normalizedId = script:ConvertTo-P2NormalizedComputeId `
                -Id $originalId `
                -Type ([string]$placement.type)

            script:Write-P2ComputeNormalizationLog `
                -Scope 'GroupDefault' `
                -OriginalId $originalId `
                -NormalizedId $normalizedId `
                -Name ([string]$placement.name) `
                -Type ([string]$placement.type) `
                -VmName ''

            $audit.Add([pscustomobject][ordered]@{
                MobilityGroupName = $groupName
                Scope             = 'GroupDefault'
                VM                 = $null
                PlacementName      = [string]$placement.name
                PlacementType      = [string]$placement.type
                OriginalId         = $originalId
                NormalizedId       = $normalizedId
                Changed            = ($originalId -cne $normalizedId)
            })

            $placement.id = $normalizedId
        }

        foreach ($migration in @($group.migrations)) {
            if (-not $migration) { continue }

            $vmName = [string]$migration.entity.entityName
            if ([string]::IsNullOrWhiteSpace($vmName)) {
                $vmName = [string]$migration.entityName
            }

            foreach ($placement in @($migration.placement)) {
                if (-not $placement) { continue }

                $originalId = [string]$placement.id
                $normalizedId = script:ConvertTo-P2NormalizedComputeId `
                    -Id $originalId `
                    -Type ([string]$placement.type)

                script:Write-P2ComputeNormalizationLog `
                    -Scope 'PerVM' `
                    -OriginalId $originalId `
                    -NormalizedId $normalizedId `
                    -Name ([string]$placement.name) `
                    -Type ([string]$placement.type) `
                    -VmName $vmName

                $audit.Add([pscustomobject][ordered]@{
                    MobilityGroupName = $groupName
                    Scope             = 'PerVM'
                    VM                 = $vmName
                    PlacementName      = [string]$placement.name
                    PlacementType      = [string]$placement.type
                    OriginalId         = $originalId
                    NormalizedId       = $normalizedId
                    Changed            = ($originalId -cne $normalizedId)
                })

                $placement.id = $normalizedId
            }
        }
    }

    # Fail closed if a confirmed invalid cluster prefix remains anywhere in the outgoing payload.
    $remaining = @($audit | Where-Object {
        $_.NormalizedId -match '^ClusterComputeResource-domain-c\d+$'
    })

    if ($remaining.Count) {
        $details = $remaining | ForEach-Object {
            "Scope='$($_.Scope)'; VM='$($_.VM)'; Id='$($_.NormalizedId)'"
        }
        throw ('Compute placement normalization failed. Unnormalized IDs remain: {0}' -f
            ($details -join '; '))
    }

    $folder = if (Get-Command -Name script:Get-P2ArtifactFolder -ErrorAction SilentlyContinue) {
        script:Get-P2ArtifactFolder
    }
    else {
        $PSScriptRoot
    }

    if ([string]::IsNullOrWhiteSpace([string]$folder)) {
        $folder = $PSScriptRoot
    }

    if (-not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
    $auditPath = Join-Path $folder "HCX91-ComputeId-Normalization-$stamp.json"

    [pscustomobject][ordered]@{
        CapturedAt       = (Get-Date).ToString('o')
        RecordCount      = $audit.Count
        ChangedCount     = @($audit | Where-Object Changed).Count
        UnchangedCount   = @($audit | Where-Object { -not $_.Changed }).Count
        Placements       = @($audit)
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $auditPath -Encoding utf8

    $summary = "COMPUTE ID NORMALIZATION SUMMARY: Records='$($audit.Count)'; Changed='$(@($audit | Where-Object Changed).Count)'; Audit='$auditPath'."
    if (Get-Command -Name Log -ErrorAction SilentlyContinue) {
        Log $summary PASS
    }
    else {
        Write-Host $summary -ForegroundColor Green
    }

    # Return the same payload object after in-place normalization.
    return $Payload
}
function script:ConvertTo-P2NormalizedDestinationId {
    param([AllowNull()][string]$Id,[AllowNull()][string]$Type)
    if([string]::IsNullOrWhiteSpace($Id)){return $Id}
    switch(([string]$Type).Trim().ToLowerInvariant()){
        'cluster'    {if($Id-match'^ClusterComputeResource-(domain-c\d+)$'){return $Matches[1]}}
        'host'       {if($Id-match'^HostSystem-(host-\d+)$'){return $Matches[1]}}
        'datastore'  {if($Id-match'^Datastore-(datastore-\d+)$'){return $Matches[1]}}
        'storagepod' {if($Id-match'^StoragePod-(group-p\d+)$'){return $Matches[1]}}
    }
    if($Id-match'^ClusterComputeResource-(domain-c\d+)$'){return $Matches[1]}
    if($Id-match'^HostSystem-(host-\d+)$'){return $Matches[1]}
    if($Id-match'^Datastore-(datastore-\d+)$'){return $Matches[1]}
    if($Id-match'^StoragePod-(group-p\d+)$'){return $Matches[1]}
    return $Id
}

function script:Assert-P2DestinationObject {
    param([string]$Type,[string]$Id,[string]$Context)
    $patterns=@{cluster='^domain-c\d+$';host='^host-\d+$';datastore='^datastore-\d+$';storagepod='^group-p\d+$'}
    $t=([string]$Type).Trim().ToLowerInvariant()
    if(-not$patterns.ContainsKey($t)){throw "Unsupported HCX destination type '$Type' in $Context."}
    if($Id-notmatch$patterns[$t]){throw "Invalid HCX destination object in $Context`: Type='$t'; Id='$Id'; Expected='$($patterns[$t])'."}
}

function script:Write-P2DestinationSelectionLog {
    param([string]$Category,[string]$Scope,[string]$VM,[string]$Name,[string]$Type,[string]$OriginalId,[string]$NormalizedId,[string]$EffectiveId,[string]$Profile,[string]$Provisioning)
    $changed=$OriginalId-cne$NormalizedId
    $msg="DESTINATION $Category`: Scope='$Scope'; VM='$VM'; Name='$Name'; Type='$Type'; OriginalId='$OriginalId'; NormalizedId='$NormalizedId'; EffectiveId='$EffectiveId'; Changed='$changed'"
    if($Category-eq'STORAGE'){$msg+="; StorageProfile='$Profile'; DiskProvisionType='$Provisioning'"}
    $msg+='.'
    $status=if($changed){'WARN'}else{'PASS'}
    if(Get-Command Log -ErrorAction SilentlyContinue){Log $msg $status}else{Write-Host $msg}
}

function script:Normalize-P2HostAndStoragePodPayload {
    param([Parameter(Mandatory)]$Payload)
    if(-not$Payload){throw'Host and datastore-cluster normalization received an empty payload.'}
    $groups=@($Payload.items)
    if(-not$groups.Count){throw'Payload does not contain items.'}
    $audit=[System.Collections.Generic.List[object]]::new()

    foreach($group in $groups){
        $groupName=[string]$group.name
        $gc=@($group.groupDefaults.placement)|Select-Object -First 1
        $gs=$group.groupDefaults.storage.defaultStorage

        if($gc){
            $old=[string]$gc.id;$type=([string]$gc.type).ToLowerInvariant()
            $new=script:ConvertTo-P2NormalizedDestinationId $old $type
            script:Assert-P2DestinationObject $type $new "group '$groupName' compute default"
            $gc.id=$new
            script:Write-P2DestinationSelectionLog COMPUTE GroupDefault '' ([string]$gc.name) $type $old $new $new '' ''
            $audit.Add([pscustomobject]@{Group=$groupName;VM=$null;Category='COMPUTE';Scope='GroupDefault';Name=[string]$gc.name;Type=$type;OriginalId=$old;NormalizedId=$new;EffectiveId=$new;Changed=($old-cne$new);Profile=$null;Provisioning=$null})
        }
        if($gs){
            $old=[string]$gs.id;$type=([string]$gs.type).ToLowerInvariant()
            $new=script:ConvertTo-P2NormalizedDestinationId $old $type
            script:Assert-P2DestinationObject $type $new "group '$groupName' storage default"
            $gs.id=$new
            $profile=[string](@($gs.storageParams)|Where-Object option -eq 'StorageProfile'|Select-Object -First 1).name
            $prov=[string]$gs.diskProvisionType
            script:Write-P2DestinationSelectionLog STORAGE GroupDefault '' ([string]$gs.name) $type $old $new $new $profile $prov
            $audit.Add([pscustomobject]@{Group=$groupName;VM=$null;Category='STORAGE';Scope='GroupDefault';Name=[string]$gs.name;Type=$type;OriginalId=$old;NormalizedId=$new;EffectiveId=$new;Changed=($old-cne$new);Profile=$profile;Provisioning=$prov})
        }

        foreach($m in @($group.migrations)){
            if(-not$m){continue}
            $vm=[string]$m.entity.entityName;if([string]::IsNullOrWhiteSpace($vm)){$vm=[string]$m.entityName}
            $mc=@($m.placement)|Select-Object -First 1
            if($mc){$scope='PerVMOverride';$old=[string]$mc.id;$type=([string]$mc.type).ToLowerInvariant();$new=script:ConvertTo-P2NormalizedDestinationId $old $type;script:Assert-P2DestinationObject $type $new "VM '$vm' compute override";$mc.id=$new;$name=[string]$mc.name}
            elseif($gc){$scope='InheritedGroupDefault';$old=[string]$gc.id;$new=$old;$type=([string]$gc.type).ToLowerInvariant();$name=[string]$gc.name}
            else{throw "VM '$vm' has no compute override and no group compute default."}
            script:Write-P2DestinationSelectionLog COMPUTE $scope $vm $name $type $old $new $new '' ''
            $audit.Add([pscustomobject]@{Group=$groupName;VM=$vm;Category='COMPUTE';Scope=$scope;Name=$name;Type=$type;OriginalId=$old;NormalizedId=$new;EffectiveId=$new;Changed=($old-cne$new);Profile=$null;Provisioning=$null})

            $ms=$null;if($m.PSObject.Properties['storage'] -and $m.storage -and $m.storage.PSObject.Properties['defaultStorage']){$ms=$m.storage.defaultStorage}
            if($ms){$scope='PerVMOverride';$old=[string]$ms.id;$type=([string]$ms.type).ToLowerInvariant();$new=script:ConvertTo-P2NormalizedDestinationId $old $type;script:Assert-P2DestinationObject $type $new "VM '$vm' storage override";$ms.id=$new;$name=[string]$ms.name;$profile=[string](@($ms.storageParams)|Where-Object option -eq 'StorageProfile'|Select-Object -First 1).name;$prov=[string]$ms.diskProvisionType}
            elseif($gs){$scope='InheritedGroupDefault';$old=[string]$gs.id;$new=$old;$type=([string]$gs.type).ToLowerInvariant();$name=[string]$gs.name;$profile=[string](@($gs.storageParams)|Where-Object option -eq 'StorageProfile'|Select-Object -First 1).name;$prov=[string]$gs.diskProvisionType}
            else{throw "VM '$vm' has no storage override and no group storage default."}
            script:Write-P2DestinationSelectionLog STORAGE $scope $vm $name $type $old $new $new $profile $prov
            $audit.Add([pscustomobject]@{Group=$groupName;VM=$vm;Category='STORAGE';Scope=$scope;Name=$name;Type=$type;OriginalId=$old;NormalizedId=$new;EffectiveId=$new;Changed=($old-cne$new);Profile=$profile;Provisioning=$prov})
        }
    }

    $folder=if(Get-Command script:Get-P2ArtifactFolder -ErrorAction SilentlyContinue){script:Get-P2ArtifactFolder}else{$PSScriptRoot}
    if([string]::IsNullOrWhiteSpace([string]$folder)){$folder=$PSScriptRoot}
    if(-not(Test-Path -LiteralPath $folder)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}
    $path=Join-Path $folder("HCX91-Host-StoragePod-Audit-{0}.json"-f(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    [pscustomobject]@{CapturedAt=(Get-Date).ToString('o');Records=$audit.Count;Changed=@($audit|Where-Object Changed).Count;Selections=@($audit)}|ConvertTo-Json -Depth 10|Set-Content -LiteralPath $path -Encoding utf8
    if(Get-Command Log -ErrorAction SilentlyContinue){Log "HOST/STORAGEPOD SUMMARY: Records='$($audit.Count)'; Audit='$path'." PASS}
    return $Payload
}

function script:Get-P2SupportedDestinationTypes {
    [pscustomobject]@{Compute=@('cluster','host');Storage=@('datastore','storagepod');Inventory=@('ClusterComputeResource','HostSystem','Datastore','StoragePod')}
}
$xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="HCX 9.1 Mobility Group Builder" Height="900" Width="1580" MinHeight="760" MinWidth="1200" WindowStartupLocation="CenterScreen" Background="#071015" Foreground="#E6E6E6" FontFamily="Segoe UI">
<Window.Resources>
<Style TargetType="Button"><Setter Property="Background" Value="#2B3740"/><Setter Property="Foreground" Value="#F1F4F6"/><Setter Property="BorderBrush" Value="#5F7482"/><Setter Property="Padding" Value="9,4"/><Setter Property="Margin" Value="4"/><Setter Property="MinHeight" Value="28"/></Style>
<Style TargetType="TextBlock"><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="Margin" Value="3"/></Style>
<Style TargetType="TextBox"><Setter Property="Background" Value="#071015"/><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="BorderBrush" Value="#607D8B"/><Setter Property="Padding" Value="4"/></Style>
<Style TargetType="PasswordBox"><Setter Property="Background" Value="#071015"/><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="BorderBrush" Value="#607D8B"/><Setter Property="Padding" Value="4"/></Style>
<Style TargetType="ComboBox">
<Setter Property="MinHeight" Value="27"/>
<Setter Property="Margin" Value="3"/>
<Setter Property="Background" Value="#E8EDF0"/>
<Setter Property="Foreground" Value="#101820"/>
<Setter Property="BorderBrush" Value="#6F8794"/>
<Setter Property="Padding" Value="5,2"/>
</Style>
<Style TargetType="ComboBoxItem">
<Setter Property="Background" Value="#E8EDF0"/>
<Setter Property="Foreground" Value="#101820"/>
<Setter Property="Padding" Value="6,4"/>
<Style.Triggers>
<Trigger Property="IsHighlighted" Value="True"><Setter Property="Background" Value="#3B6677"/><Setter Property="Foreground" Value="#FFFFFF"/></Trigger>
<Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#2C5364"/><Setter Property="Foreground" Value="#FFFFFF"/></Trigger>
</Style.Triggers>
</Style>
<Style TargetType="RadioButton"><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="Padding" Value="2"/></Style>
<Style TargetType="CheckBox"><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="Padding" Value="2"/></Style><Style TargetType="TabControl"><Setter Property="Background" Value="#071015"/><Setter Property="BorderBrush" Value="#5B7280"/></Style><Style TargetType="TabItem"><Setter Property="Background" Value="#25343D"/><Setter Property="Foreground" Value="#F1F4F6"/><Setter Property="BorderBrush" Value="#5B7280"/><Setter Property="Padding" Value="10,5"/><Setter Property="Margin" Value="2,0,2,0"/><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TabItem"><Border x:Name="TabBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1,1,1,0" CornerRadius="3,3,0,0" Padding="{TemplateBinding Padding}"><ContentPresenter ContentSource="Header" TextElement.Foreground="#F1F4F6" HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsSelected" Value="True"><Setter TargetName="TabBorder" Property="Background" Value="#314550"/><Setter TargetName="TabBorder" Property="BorderBrush" Value="#76C7D8"/></Trigger><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="TabBorder" Property="Background" Value="#3A4E59"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.55"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style>
<Style TargetType="GroupBox"><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="Background" Value="#0D1B22"/><Setter Property="BorderBrush" Value="#2B3740"/><Setter Property="Margin" Value="5"/><Setter Property="Padding" Value="7"/></Style>
<Style TargetType="DataGrid"><Setter Property="Background" Value="#071015"/><Setter Property="Foreground" Value="#E6E6E6"/><Setter Property="RowBackground" Value="#071015"/><Setter Property="AlternatingRowBackground" Value="#0D1B22"/><Setter Property="GridLinesVisibility" Value="All"/><Setter Property="HorizontalGridLinesBrush" Value="#607D8B"/><Setter Property="VerticalGridLinesBrush" Value="#607D8B"/><Setter Property="BorderBrush" Value="#607D8B"/><Setter Property="RowHeaderWidth" Value="0"/></Style>
<Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="#2B3740"/><Setter Property="Foreground" Value="#F1F4F6"/><Setter Property="FontWeight" Value="SemiBold"/></Style>
</Window.Resources>
<Grid Margin="10"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="165"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<Grid><Grid.ColumnDefinitions><ColumnDefinition Width="0.56*" MinWidth="360"/><ColumnDefinition Width="1.94*" MinWidth="720"/></Grid.ColumnDefinitions>
<GroupBox Header="Prerequisites"><Grid><Grid.RowDefinitions><RowDefinition/><RowDefinition Height="Auto"/></Grid.RowDefinitions><Grid Margin="4,2,4,4">
<Grid.ColumnDefinitions><ColumnDefinition Width="1*"/><ColumnDefinition Width="1*"/><ColumnDefinition Width="1.3*"/></Grid.ColumnDefinitions>
<Grid.RowDefinitions><RowDefinition Height="25"/><RowDefinition Height="25"/></Grid.RowDefinitions>
<StackPanel Grid.Row="0" Grid.Column="0" Orientation="Horizontal" HorizontalAlignment="Center"><TextBlock Text="PowerShell:"/><TextBlock x:Name="lblPS"/></StackPanel>
<StackPanel Grid.Row="0" Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Center" ToolTip="Single-Threaded Apartment mode, required by Windows Presentation Foundation controls."><TextBlock Text="STA:"/><TextBlock x:Name="lblSTA"/></StackPanel>
<StackPanel Grid.Row="0" Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Center"><TextBlock Text="VCF.PowerCLI:"/><TextBlock x:Name="lblPCLI"/></StackPanel>
<StackPanel Grid.Row="1" Grid.Column="0" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Center"><TextBlock Text="HCX API:"/><TextBlock x:Name="lblApi"/></StackPanel>
<StackPanel Grid.Row="1" Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Center"></StackPanel>
</Grid>
<StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Center"><Button x:Name="btnRecheck" Content="Recheck"/><Button x:Name="btnInstall" Content="Install VCF.PowerCLI"/></StackPanel></Grid></GroupBox>
<GroupBox Header="HCX 9.1 and vCenter Connections" Grid.Column="1">
<Grid Margin="4">
<Grid.ColumnDefinitions><ColumnDefinition Width="125"/><ColumnDefinition Width="1.18*"/><ColumnDefinition Width="100"/><ColumnDefinition Width="1.05*"/><ColumnDefinition Width="100"/><ColumnDefinition Width="1.05*"/></Grid.ColumnDefinitions>
<Grid.RowDefinitions><RowDefinition Height="32"/><RowDefinition Height="32"/><RowDefinition Height="32"/><RowDefinition Height="46"/></Grid.RowDefinitions>

<TextBlock Grid.Row="0" Grid.Column="0" Text="Source HCX Manager" VerticalAlignment="Center"/>
<TextBox x:Name="txtHcx" Grid.Row="0" Grid.Column="1"/>
<TextBlock Grid.Row="0" Grid.Column="2" Text="HCX Username" VerticalAlignment="Center"/>
<TextBox x:Name="txtUser" Grid.Row="0" Grid.Column="3"/>
<TextBlock Grid.Row="0" Grid.Column="4" Text="HCX Password" VerticalAlignment="Center"/>
<PasswordBox x:Name="txtPassword" Grid.Row="0" Grid.Column="5"/>

<TextBlock Grid.Row="1" Grid.Column="0" Text="Source vCenter" VerticalAlignment="Center"/>
<TextBox x:Name="txtSourceVC" Grid.Row="1" Grid.Column="1"/>
<TextBlock Grid.Row="1" Grid.Column="2" Text="Source Username" VerticalAlignment="Center"/>
<TextBox x:Name="txtSourceVCUser" Grid.Row="1" Grid.Column="3"/>
<TextBlock Grid.Row="1" Grid.Column="4" Text="Source Password" VerticalAlignment="Center"/>
<PasswordBox x:Name="txtSourceVCPass" Grid.Row="1" Grid.Column="5"/>

<TextBlock Grid.Row="2" Grid.Column="0" Text="Destination vCenter" VerticalAlignment="Center"/>
<TextBox x:Name="txtDestinationVC" Grid.Row="2" Grid.Column="1"/>
<TextBlock Grid.Row="2" Grid.Column="2" Text="Destination Username" VerticalAlignment="Center"/>
<TextBox x:Name="txtDestinationVCUser" Grid.Row="2" Grid.Column="3"/>
<TextBlock Grid.Row="2" Grid.Column="4" Text="Destination Password" VerticalAlignment="Center"/>
<PasswordBox x:Name="txtDestinationVCPass" Grid.Row="2" Grid.Column="5"/>

<TextBlock Grid.Row="3" Grid.Column="0" Text="Use Source HCX Manager" Foreground="#76C7D8" FontWeight="SemiBold" VerticalAlignment="Center" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="4,2"/>
<WrapPanel Grid.Row="3" Grid.Column="1" Grid.ColumnSpan="5" Orientation="Horizontal" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="4,2">
<Button x:Name="btnConnect" Content="Connect and Load Inventory" MinWidth="175" Height="30" Padding="7,2"/>
<Button x:Name="btnDisconnect" Content="Disconnect" MinWidth="96" Height="30" Padding="7,2"/>
<Button x:Name="btnSaveConnectionJson" Content="Save JSON" MinWidth="82" Height="30" Padding="7,2"/>
<Button x:Name="btnLoadConnectionJson" Content="Load JSON" MinWidth="82" Height="30" Padding="7,2"/>
</WrapPanel>
</Grid>
</GroupBox>
</Grid>
<TabControl Grid.Row="1" Margin="0,6,0,0"><TabItem Header="Prepare Mobility Group"><Grid Margin="7"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<GroupBox Header="Global Destination Settings" Margin="5,10,5,5" Padding="9">
 <Grid Margin="8">
  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="190"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="220"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="240"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="230"/></Grid.ColumnDefinitions>
  <Grid.RowDefinitions><RowDefinition Height="34"/><RowDefinition Height="34"/></Grid.RowDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="Destination Site" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbDestinationSite" Grid.Row="0" Grid.Column="1" Margin="4" DisplayMemberPath="Name" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="0" Grid.Column="2" Text="Compute" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalCompute" Grid.Row="0" Grid.Column="3" Margin="4" DisplayMemberPath="DisplayName" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="0" Grid.Column="4" Text="Datastore" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalDatastore" Grid.Row="0" Grid.Column="5" Margin="4" DisplayMemberPath="DisplayName" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="0" Grid.Column="6" Text="Storage Policy" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalStoragePolicy" Grid.Row="0" Grid.Column="7" Margin="4" DisplayMemberPath="Name" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="1" Grid.Column="0" Text="Folder" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalFolder" Grid.Row="1" Grid.Column="1" Margin="4" DisplayMemberPath="Name" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="1" Grid.Column="2" Text="Migration Type" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbMigrationType" Grid.Row="1" Grid.Column="3" Margin="4" ItemsSource="{Binding DataContext.MigrationTypes, RelativeSource={RelativeSource AncestorType=Window}}" SelectedIndex="0"/>
  <Button x:Name="btnApplyGlobal" Grid.Row="1" Grid.Column="5" Grid.ColumnSpan="3" Margin="4" Content="Apply to All VMs"/>
 </Grid>
</GroupBox>
<DataGrid x:Name="gridVMs" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" SelectionMode="Extended" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" EnableRowVirtualization="True" EnableColumnVirtualization="True" ScrollViewer.CanContentScroll="True"><DataGrid.Columns>
<DataGridCheckBoxColumn Header="Include" Binding="{Binding Include, Mode=TwoWay}" Width="60"/><DataGridTextColumn Header="Mobility Group #" Binding="{Binding MobilityGroupNumber, Mode=TwoWay}" Width="105"/><DataGridTextColumn Header="Status" Binding="{Binding Status}" IsReadOnly="True" Width="85"/><DataGridTextColumn Header="VM Name" Binding="{Binding VMName}" IsReadOnly="True" Width="160"/><DataGridTextColumn Header="Power" Binding="{Binding PowerState}" IsReadOnly="True" Width="80"/><DataGridTextColumn Header="Source Compute" Binding="{Binding SourceCompute}" IsReadOnly="True" Width="150"/><DataGridTextColumn Header="Source Datastore" Binding="{Binding SourceDatastore}" IsReadOnly="True" Width="150"/><DataGridTextColumn Header="Source Networks" Binding="{Binding SourceNetworks}" IsReadOnly="True" Width="190"/><DataGridTextColumn Header="NICs" Binding="{Binding NicCount}" IsReadOnly="True" Width="45"/>
<DataGridTemplateColumn Header="Network Mapping" Width="330"><DataGridTemplateColumn.CellTemplate><DataTemplate><StackPanel Orientation="Horizontal"><TextBlock Text="{Binding NicSummary}" Width="255" TextTrimming="CharacterEllipsis" ToolTip="{Binding NicSummary}"/><Button Content="Configure" Tag="{Binding}" Padding="5,2" MinHeight="23"/></StackPanel></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
<DataGridTemplateColumn Header="Migration Type" Width="105"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.MigrationTypes, RelativeSource={RelativeSource AncestorType=Window}}" SelectedItem="{Binding MigrationType, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Destination Site" Binding="{Binding DestinationSite}" IsReadOnly="True" Width="145"/><DataGridTemplateColumn Header="Destination Compute" Width="170"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Computes, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding DestinationComputeId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTemplateColumn Header="Destination Folder" Width="155"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Folders, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="Name" SelectedValuePath="Id" SelectedValue="{Binding DestinationFolderId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTemplateColumn Header="Destination Datastore" Width="180"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Datastores, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding DestinationDatastoreId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Validation Detail" Binding="{Binding ValidationMessage}" IsReadOnly="True" Width="300"/>
</DataGrid.Columns></DataGrid>
<Grid Grid.Row="2"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><StackPanel Orientation="Horizontal"><Button x:Name="btnImport" Content="Import VM CSV"/><Button x:Name="btnExample" Content="Download Example CSV"/><Button x:Name="btnRemove" Content="Remove Selected VM(s)"/><Button x:Name="btnClear" Content="Clear Grid"/><Button x:Name="btnRefresh" Content="Refresh Discovery"/></StackPanel><StackPanel Grid.Column="1" Orientation="Horizontal"><TextBlock Text="Validation:" FontWeight="SemiBold" VerticalAlignment="Center"/><TextBlock x:Name="lblValidation" Text="Not validated" Foreground="#76C7D8" VerticalAlignment="Center"/><Button x:Name="btnValidate" Content="Validate" Width="100"/><Button x:Name="btnCreateCsv" Content="Create Mobility Group CSV" Width="145" IsEnabled="False"/></StackPanel></Grid>
</Grid></TabItem><TabItem Header="Create Mobility Group"><Grid Margin="8"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<GroupBox Header="Workload Selection"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="100"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="110"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition Height="34"/><RowDefinition Height="34"/></Grid.RowDefinitions><TextBlock Text="Phase I CSV"/><TextBox x:Name="txtP2Csv" Grid.Column="1" IsReadOnly="True"/><Button x:Name="btnP2Import" Grid.Column="2" Content="Import CSV"/><TextBlock Grid.Column="3" Text="Base Group Name"/><StackPanel Grid.Column="4" Orientation="Horizontal"><TextBox x:Name="txtP2Name" Width="280"/><TextBlock Text="Groups (1-50)" Margin="10,3"/><ComboBox x:Name="cmbP2GroupCount" Width="65" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" Foreground="#000000" Background="#FFFFFF" FontWeight="Bold" FontSize="14">
 <ComboBox.ItemTemplate><DataTemplate><Border Background="#FFFFFF" Padding="6,3"><TextBlock Text="{Binding}" Foreground="#000000" Background="#FFFFFF" FontWeight="Bold" FontSize="14"/></Border></DataTemplate></ComboBox.ItemTemplate>
 <ComboBox.ItemContainerStyle><Style TargetType="{x:Type ComboBoxItem}"><Setter Property="Foreground" Value="#000000"/><Setter Property="Background" Value="#FFFFFF"/><Style.Triggers><Trigger Property="IsHighlighted" Value="True"><Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="#000000"/></Trigger><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#CDE8FF"/><Setter Property="Foreground" Value="#000000"/></Trigger></Style.Triggers></Style></ComboBox.ItemContainerStyle>
</ComboBox></StackPanel><TextBlock Grid.Row="1" Text="Source HCX"/><TextBox x:Name="txtP2Hcx" Grid.Row="1" Grid.Column="1" IsReadOnly="True"/><TextBlock Grid.Row="1" Grid.Column="3" Text="Source Site"/><TextBox x:Name="txtP2SourceSite" Grid.Row="1" Grid.Column="4"/></Grid></GroupBox>
<GroupBox Grid.Row="1" Header="Destination Settings"><WrapPanel><TextBlock Text="Destination Site" Width="105"/><TextBox x:Name="txtP2DestSite" Width="220" IsReadOnly="True"/><TextBlock Text="Compute" Width="70"/><ComboBox x:Name="cmbP2Compute" Width="230" DisplayMemberPath="DisplayName" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/><TextBlock Text="Storage" Width="70"/><ComboBox x:Name="cmbP2Storage" Width="230" DisplayMemberPath="DisplayName" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/><TextBlock Text="Storage Policy" Width="100"/><ComboBox x:Name="cmbP2Policy" Width="230" DisplayMemberPath="Name"/><TextBlock Text="Disk Format" Width="80"/><ComboBox x:Name="cmbP2Disk" Width="190" SelectedIndex="0"><ComboBoxItem>Thin Provision</ComboBoxItem><ComboBoxItem>Thick Provision Lazy Zeroed</ComboBoxItem><ComboBoxItem>Thick Provision Eager Zeroed</ComboBoxItem><ComboBoxItem>Same format as source</ComboBoxItem></ComboBox><Button x:Name="btnP2Inventory" Content="Load Storage"/><Button x:Name="btnP2Apply" Content="Apply to All VMs"/></WrapPanel></GroupBox>
<GroupBox Grid.Row="2" Header="Migration Settings"><StackPanel><WrapPanel><TextBlock Text="Migration Type: HCX Assisted vMotion" Width="275"/><RadioButton x:Name="rbP2Now" Content="Start after transfer is complete" IsChecked="True" Margin="8"/><RadioButton x:Name="rbP2Schedule" Content="Set switchover schedule" Margin="8"/><RadioButton x:Name="rbP2Defer" Content="Defer switchover" Margin="8"/><TextBox x:Name="txtP2Schedule" Width="180" ToolTip="Schedule value"/></WrapPanel><WrapPanel><CheckBox x:Name="chkP2Mac" Content="Retain MAC" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2HW" Content="Upgrade Virtual Hardware" Margin="8"/><CheckBox x:Name="chkP2Tools" Content="Upgrade VM Tools" Margin="8"/><CheckBox x:Name="chkP2Attrs" Content="Migrate Custom Attributes" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2Tags" Content="Migrate vCenter Tags" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2Iso" Content="Force unmount ISO images" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2Security" Content="Replicate Security Tags" IsChecked="False" Margin="8"/></WrapPanel><TextBlock Text="For a complete migration, select a storage policy. Use an encryption-capable policy for vTPM or encrypted VMs. Drafts may be completed later in HCX Manager." Foreground="#76C7D8"/></StackPanel></GroupBox>
<DataGrid x:Name="gridP2" Grid.Row="3" AutoGenerateColumns="False" CanUserAddRows="False" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"><DataGrid.Columns><DataGridCheckBoxColumn Header="Include" Binding="{Binding Include}" Width="55"/><DataGridTextColumn Header="Group #" Binding="{Binding MobilityGroupNumber}" Width="65"/><DataGridTextColumn Header="VM Name" Binding="{Binding VMName}" Width="140" IsReadOnly="True"/><DataGridTextColumn Header="Network Mapping" Binding="{Binding NetworkSummary}" Width="280" IsReadOnly="True"/><DataGridTemplateColumn Header="Compute" Width="210"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Computes, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding ComputeId, Mode=TwoWay}" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Folder" Binding="{Binding Folder}" Width="110"/><DataGridTextColumn Header="Storage Type" Binding="{Binding StorageType}" Width="125"/><DataGridTemplateColumn Header="Storage" Width="220"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Datastores, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding StorageId, Mode=TwoWay}" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Storage Policy" Binding="{Binding StoragePolicy}" Width="180"/><DataGridTextColumn Header="Disk Format" Binding="{Binding DiskFormat}" Width="150"/></DataGrid.Columns></DataGrid>
<StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right"><TextBlock Text="Phase II: "/><TextBlock x:Name="lblP2Status" Text="Not validated" Foreground="#76C7D8"/><Button x:Name="btnP2Validate" Content="Validate"/><Button x:Name="btnP2Preview" Content="Preview Payload" IsEnabled="False"/><Button x:Name="btnP2Create" Content="Save Draft to HCX" IsEnabled="False"/></StackPanel></Grid></TabItem></TabControl>
<GroupBox Grid.Row="2" Header="Log"><TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12" Background="#071015" Foreground="#E6E6E6"/></GroupBox>
<GroupBox Grid.Row="3" Header="Output Location" Margin="5"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="Default base path" VerticalAlignment="Center"/><TextBox x:Name="txtOutputPath" Grid.Column="1" Margin="4"/><StackPanel Grid.Column="2" Orientation="Horizontal"><Button x:Name="btnBrowseOutput" Content="Browse..."/><Button x:Name="btnApplyOutput" Content="Apply Path"/></StackPanel></Grid></GroupBox><StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center"><Button x:Name="btnOpenOutput" Content="Open Last CSV"/><Button x:Name="btnOpenRun" Content="Open Run Folder"/><Button x:Name="btnClose" Content="Close"/></StackPanel>
</Grid></Window>
'@
$script:Window=[Windows.Markup.XamlReader]::Parse($xaml)
$names=@('lblPS','lblSTA','lblPCLI','lblApi','btnRecheck','btnInstall','txtHcx','txtUser','txtPassword','txtSourceVC','txtSourceVCUser','txtSourceVCPass','txtDestinationVC','txtDestinationVCUser','txtDestinationVCPass','btnConnect','btnDisconnect','btnSaveConnectionJson','btnLoadConnectionJson','cmbDestinationSite','cmbGlobalCompute','cmbGlobalDatastore','cmbGlobalFolder','cmbMigrationType','cmbGlobalStoragePolicy','btnApplyGlobal','gridVMs','btnImport','btnExample','btnRemove','btnClear','btnRefresh','lblValidation','btnValidate','btnCreateCsv','txtLog','txtOutputPath','btnBrowseOutput','btnApplyOutput','btnOpenOutput','btnOpenRun','btnClose','txtP2Csv','btnP2Import','txtP2Name','cmbP2GroupCount','txtP2Hcx','txtP2SourceSite','txtP2DestSite','cmbP2Compute','cmbP2Storage','cmbP2Policy','cmbP2Disk','btnP2Inventory','btnP2Apply','rbP2Now','rbP2Schedule','rbP2Defer','txtP2Schedule','chkP2Mac','chkP2HW','chkP2Tools','chkP2Attrs','chkP2Tags','chkP2Iso','chkP2Security','gridP2','lblP2Status','btnP2Validate','btnP2Preview','btnP2Create')
foreach($n in $names){Set-Variable -Scope Script -Name $n -Value $script:Window.FindName($n)}
$script:gridVMs.ItemsSource=$script:Rows
foreach($n in 1..50){[void]$script:cmbP2GroupCount.Items.Add($n)};$script:cmbP2GroupCount.SelectedIndex=0;
$script:txtOutputPath.Text=$script:OutputBase
$script:chkP2Security.IsChecked=$false
if ([string]::IsNullOrWhiteSpace($script:txtUser.Text)) { $script:txtUser.Text = 'administrator@vsphere.local' }
if ([string]::IsNullOrWhiteSpace($script:txtSourceVCUser.Text)) { $script:txtSourceVCUser.Text = 'administrator@vsphere.local' }
if ([string]::IsNullOrWhiteSpace($script:txtDestinationVCUser.Text)) { $script:txtDestinationVCUser.Text = 'administrator@vsphere.local' }

$script:Window.AddHandler([Windows.Controls.Button]::ClickEvent,[Windows.RoutedEventHandler]{param($sender,$e);if($e.OriginalSource.Name -eq '' -and $e.OriginalSource.Content -eq 'Configure' -and $e.OriginalSource.Tag){Show-NicMappingDialog $e.OriginalSource.Tag}})
$script:gridVMs.Add_CellEditEnding({Invalidate-Validation 'VM grid edited'})
$script:btnRecheck.Add_Click({Update-Prerequisites})
$script:btnInstall.Add_Click({$null=Ensure-Module 'VCF.PowerCLI';Update-Prerequisites})
$script:btnConnect.Add_Click({
    try{
        $script:btnConnect.IsEnabled=$false
        Connect-Hcx91Rest $script:txtHcx.Text.Trim() $script:txtUser.Text.Trim() $script:txtPassword.Password
        Load-HcxInventory;Update-Prerequisites;if($script:P2Rows -and $script:P2Rows.Count -gt 0){P2-Inventory;Restore-P2SelectionsFromPhaseICsv}
        [Windows.MessageBox]::Show("Connected. Inventory loading completed.`n`nVMs: $($(@($script:Inventory.VMs).Count))`nNetworks: $($(@($script:Inventory.Networks).Count))`nDatastores: $($(@($script:Inventory.Datastores).Count))`nComputes: $($(@($script:Inventory.Computes).Count))`nFolders: $($(@($script:Inventory.Folders).Count))",'HCX connected')|Out-Null
    }catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'HCX connection failed','OK','Error')|Out-Null}
    finally{$script:txtPassword.Clear();$script:txtSourceVCPass.Clear();$script:txtDestinationVCPass.Clear();$script:btnConnect.IsEnabled=$true;Update-Prerequisites}
})
$script:btnDisconnect.Add_Click({Disconnect-HcxRest;Log 'Disconnected from HCX.' INFO})
$script:btnSaveConnectionJson.Add_Click({
    try { Save-ConnectionProfile }
    catch { Log $_.Exception.Message ERROR; [Windows.MessageBox]::Show($_.Exception.Message,'Save JSON failed','OK','Error') | Out-Null }
})
$script:btnLoadConnectionJson.Add_Click({
    try { Load-ConnectionProfile }
    catch { Log $_.Exception.Message ERROR; [Windows.MessageBox]::Show($_.Exception.Message,'Load JSON failed','OK','Error') | Out-Null }
})
$script:btnApplyGlobal.Add_Click({Apply-GlobalSelection})
$script:btnImport.Add_Click({
    try{
        if(-not$script:Hcx.Connected){throw 'Connect to HCX and load inventory before importing VMs.'}
        $d=[Microsoft.Win32.OpenFileDialog]::new();$d.Filter='CSV files (*.csv)|*.csv';if(-not$d.ShowDialog()){return}
        $input=@(Import-Csv -LiteralPath $d.FileName);if(-not$input.Count){throw 'The selected CSV contains no rows.'}
        $script:Rows.Clear()
        foreach($i in $input){$name=[string](Get-Value $i @('VMName','Name'));if($name){$row=Convert-ToVmRow $name.Trim();$mg=[string](Get-Value $i @('MobilityGroupNumber','MobilityGroup#','MobilityGroup'));if($mg){$row.MobilityGroupNumber=[int]$mg};$script:Rows.Add($row)}}
        Apply-GlobalSelection;Log "Imported and discovered $($script:Rows.Count) VM row(s) from $($d.FileName)." PASS
    }catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Import failed','OK','Error')|Out-Null}
})
$script:btnExample.Add_Click({$p=Join-Path $script:OutputBase 'example-vm-import.csv';@([pscustomobject]@{VMName='appserver01';MobilityGroupNumber=1},[pscustomobject]@{VMName='appserver02';MobilityGroupNumber=2})|Export-Csv -LiteralPath $p -NoTypeInformation -Encoding utf8BOM;Log "Example CSV created: $p" PASS;Invoke-Item $p})
$script:btnRemove.Add_Click({foreach($r in @($script:gridVMs.SelectedItems)){$script:Rows.Remove($r)};Invalidate-Validation 'Selected VM rows removed'})
$script:btnClear.Add_Click({$script:Rows.Clear();Invalidate-Validation 'VM grid cleared'})
$script:btnRefresh.Add_Click({
    try{Load-HcxInventory;$names=@($script:Rows|ForEach-Object{$_.VMName});$script:Rows.Clear();foreach($n in $names){$script:Rows.Add((Convert-ToVmRow $n))};Apply-GlobalSelection;Log 'HCX inventory and VM discovery refreshed.' PASS}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Refresh failed','OK','Error')|Out-Null}
})
$script:btnValidate.Add_Click({$null=Validate-Phase1})
$script:btnCreateCsv.Add_Click({try{Export-Phase1Csv}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'CSV creation failed','OK','Error')|Out-Null}})
$script:btnBrowseOutput.Add_Click({try{Select-OutputBase}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Output path failed','OK','Error')|Out-Null}})
$script:btnApplyOutput.Add_Click({try{Set-OutputBase $script:txtOutputPath.Text}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Output path failed','OK','Error')|Out-Null}})
$script:btnOpenOutput.Add_Click({if($script:LastCsv -and(Test-Path -LiteralPath $script:LastCsv)){Invoke-Item $script:LastCsv}else{[Windows.MessageBox]::Show('No Phase I CSV has been created in this session.','No CSV')|Out-Null}})
$script:btnOpenRun.Add_Click({Invoke-Item $script:RunDir})
$script:btnClose.Add_Click({Disconnect-HcxRest;$script:Window.Close()})
$script:btnP2Import.Add_Click({try{P2-Import}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Phase II Import')|Out-Null}})
$script:btnP2Inventory.Add_Click({try{P2-Inventory}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Phase II Storage')|Out-Null}})
$script:btnP2Apply.Add_Click({P2-Apply});$script:btnP2Validate.Add_Click({P2-Validate});$script:btnP2Preview.Add_Click({P2-Preview});$script:btnP2Create.Add_Click({
    try{
        if(-not$script:P2Payload){try{$script:Window.Cursor='Wait';$script:P2Payload=@(P2-BuildAll)}finally{$script:Window.Cursor=$null}}
        P2-Create
    }catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Create Mobility Group')|Out-Null}
});$script:gridP2.Add_CellEditEnding({P2-Invalidate})

Update-Prerequisites
Log "HCX 9.1 Mobility Group CSV Builder Phase I started. Output folder: $script:RunDir" PASS
Log 'Multiple-vNIC mapping is enabled. Each adapter is validated independently before CSV creation.' INFO
$null=$script:Window.ShowDialog()