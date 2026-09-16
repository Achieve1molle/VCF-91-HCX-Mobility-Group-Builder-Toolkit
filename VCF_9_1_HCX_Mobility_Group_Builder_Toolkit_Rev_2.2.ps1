<#
.SYNOPSIS
  VCF 9.1 HCX Mobility Group Builder Toolkit Rev 2.9.1-Network-Mapping-CSV-Fixed.ps1
.DESCRIPTION
  PowerShell 7 WPF utility for connecting to an HCX 9.1 Manager, importing VM names,
  discovering source and destination inventory, mapping multiple vNICs, validating
  selections, and exporting a versioned CSV and creating HCX 9.1 mobility-group drafts using automatic current-session topology discovery.

  IMPORTANT: HCX 9.x operations in this tool use REST, not VMware.VimAutomation.Hcx.
  The legacy HCX PowerCLI module shipped with current VCF.PowerCLI is not used.

  CSV input requires VMName or Name. vTPM is detected automatically from live source vCenter inventory and is not accepted from the import CSV.
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
$script:DebugLoggingEnabled = $true
$script:DebugSequence = 0
$script:DebugArtifactDir = Join-Path $script:RunDir 'Debug-Artifacts'
New-Item -ItemType Directory -Path $script:DebugArtifactDir -Force | Out-Null
$script:TranscriptFile = Join-Path $script:RunDir ('HCX91-PowerShell-Transcript-' + (Get-Date -Format yyyyMMdd-HHmmss) + '.log')
$script:TranscriptStarted = $false
$VerbosePreference = 'Continue'
$DebugPreference = 'SilentlyContinue'
$InformationPreference = 'Continue'

$script:Rows = [Collections.ObjectModel.ObservableCollection[object]]::new()
$script:Hcx = [ordered]@{ BaseUri=''; Session=$null; Connected=$false; Version=''; User=''; SitePairs=@(); EndpointProfile=''; Headers=@{} }
$script:Inventory = [ordered]@{ VMs=@(); Networks=@(); Datastores=@(); Computes=@(); Folders=@(); Sites=@() ; Policies=@()}
$script:ValidationCurrent = $false
$script:SuppressEditInvalidation = $false
$script:LastCsv = $null
$script:SchemaVersion = '2.0'
$script:MobilityGroupMaximum = 50
$script:ImportedNetworkMappings = @{}
$script:ImportedNetworkMappingRows = @()
$script:NetworkMappingCsvPath = ''
$script:SourceVIServer=$null
$script:DestinationVIServer=$null
$script:HcxDiscovery=$null

function DoEvents { try { [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::Background) } catch {} }
function Protect-HcxDiagnosticText([AllowNull()][string]$Text) {
    if ($null -eq $Text) { return '' }
    $safe = $Text
    $safe = $safe -replace '(?i)("?(?:password|passwd|pwd|token|access_token|refresh_token|authorization|x-hm-authorization|cookie|set-cookie|xsrf-token|csrf-token)"?\s*[:=]\s*")([^"]+)(")','$1********$3'
    $safe = $safe -replace '(?i)((?:authorization|x-hm-authorization|cookie|set-cookie|xsrf-token|csrf-token)\s*[:=]\s*)([^;\r\n]+)','$1********'
    $safe = $safe -replace '(?i)(Basic|Bearer)\s+[A-Za-z0-9+/_=.-]+','$1 ********'
    foreach($secret in @($script:txtPassword.Password,$script:txtSourceVCPass.Password,$script:txtDestinationVCPass.Password)){
        if(-not[string]::IsNullOrWhiteSpace([string]$secret)){$safe=$safe -replace [regex]::Escape([string]$secret),'********'}
    }
    return $safe
}
function ConvertTo-HcxSanitizedObject($Object) {
    if ($null -eq $Object) { return $null }
    try { return (Protect-HcxDiagnosticText ($Object | ConvertTo-Json -Depth 60 -Compress)) | ConvertFrom-Json -AsHashtable -Depth 60 }
    catch { return (Protect-HcxDiagnosticText ([string]$Object)) }
}
function Write-HcxDebug([string]$Message,[string]$Category='GENERAL') {
    if(-not$script:DebugLoggingEnabled){return}
    $safe=Protect-HcxDiagnosticText $Message
    $line="$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') [DEBUG] [$Category] $safe"
    Add-Content -LiteralPath $script:LogFile -Value $line
    if($script:txtLog){$script:txtLog.AppendText($line+[Environment]::NewLine);$script:txtLog.ScrollToEnd();DoEvents}
}
function Save-HcxDebugArtifact([string]$Category,[string]$Operation,$Data,[hashtable]$Metadata=@{}) {
    if(-not$script:DebugLoggingEnabled){return $null}
    $script:DebugSequence++
    $safeCategory=($Category -replace '[^A-Za-z0-9_-]','_')
    $safeOperation=($Operation -replace '[^A-Za-z0-9_-]','_')
    $path=Join-Path $script:DebugArtifactDir ('{0:d4}-{1}-{2}-{3}.json' -f $script:DebugSequence,(Get-Date -Format 'yyyyMMdd-HHmmss-fff'),$safeCategory,$safeOperation)
    $envelope=[ordered]@{CapturedAt=(Get-Date).ToString('o');Category=$Category;Operation=$Operation;Metadata=(ConvertTo-HcxSanitizedObject $Metadata);Data=(ConvertTo-HcxSanitizedObject $Data)}
    $envelope|ConvertTo-Json -Depth 70|Set-Content -LiteralPath $path -Encoding utf8BOM
    Write-HcxDebug "Artifact saved: $path" 'ARTIFACT'
    return $path
}
function Write-HcxExceptionDiagnostic($ErrorRecord,[string]$Context='Unhandled') {
    $detail=[ordered]@{Context=$Context;Message=$ErrorRecord.Exception.Message;ExceptionType=$ErrorRecord.Exception.GetType().FullName;FullyQualifiedErrorId=$ErrorRecord.FullyQualifiedErrorId;CategoryInfo=[string]$ErrorRecord.CategoryInfo;ScriptName=$ErrorRecord.InvocationInfo.ScriptName;ScriptLineNumber=$ErrorRecord.InvocationInfo.ScriptLineNumber;PositionMessage=$ErrorRecord.InvocationInfo.PositionMessage;Line=$ErrorRecord.InvocationInfo.Line;ScriptStackTrace=$ErrorRecord.ScriptStackTrace;PowerShellStack=($ErrorRecord.Exception.StackTrace)}
    $artifact=Save-HcxDebugArtifact -Category 'EXCEPTION' -Operation $Context -Data $detail
    Write-HcxDebug ((Protect-HcxDiagnosticText ($detail|ConvertTo-Json -Depth 10 -Compress))+"; Artifact=$artifact") 'EXCEPTION'
    return $artifact
}
function Start-HcxDiagnosticTranscript {
    if(-not$script:DebugLoggingEnabled -or $script:TranscriptStarted){return}
    try{Start-Transcript -LiteralPath $script:TranscriptFile -IncludeInvocationHeader -Force|Out-Null;$script:TranscriptStarted=$true}
    catch{Write-HcxDebug "Transcript could not start: $($_.Exception.Message)" 'TRANSCRIPT'}
}
function Stop-HcxDiagnosticTranscript {
    if($script:TranscriptStarted){try{Stop-Transcript|Out-Null}catch{};$script:TranscriptStarted=$false}
}
function Log([string]$Message,[ValidateSet('INFO','WARN','ERROR','PASS')][string]$Level='INFO') {
    $Message = Protect-HcxDiagnosticText $Message
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
    $script:DebugArtifactDir=Join-Path $script:RunDir 'Debug-Artifacts';New-Item -ItemType Directory -Path $script:DebugArtifactDir -Force|Out-Null
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
function Test-HcxTransientTransportError($ErrorRecord) {
    $messages=[System.Collections.Generic.List[string]]::new()
    $exception=$ErrorRecord.Exception
    while($exception){$messages.Add([string]$exception.Message);$exception=$exception.InnerException}
    $joined=$messages -join ' | '
    return ($joined -match '(?i)response ended prematurely|ResponseEnded|connection.*closed|forcibly closed|unexpected end|HTTP/2.*error|transport connection|request was aborted|error occurred while sending')
}
function Invoke-HcxRest([string]$Method,[string]$Path,$Body=$null,[hashtable]$Headers=@{},[switch]$AllowFailure,[int]$MaxAttempts=3) {
    if (-not $script:Hcx.BaseUri) { throw 'HCX base URI is not initialized.' }
    $uri=if($Path -match '^https?://'){$Path}else{$script:Hcx.BaseUri.TrimEnd('/')+'/'+$Path.TrimStart('/')}
    $all=@{Accept='application/json'};foreach($k in $script:Hcx.Headers.Keys){$all[$k]=$script:Hcx.Headers[$k]};foreach($k in $Headers.Keys){$all[$k]=$Headers[$k]}
    $requestId=[guid]::NewGuid().ToString('N');$overallStarted=Get-Date
    $requestMeta=[ordered]@{RequestId=$requestId;Method=$Method;Uri=$uri;Path=$Path;Headers=$all;Body=$Body;StartedAt=$overallStarted.ToString('o');MaxAttempts=$MaxAttempts}
    $requestArtifact=Save-HcxDebugArtifact -Category 'REST-REQUEST' -Operation $Method -Data $requestMeta
    $lastError=$null
    for($attempt=1;$attempt -le $MaxAttempts;$attempt++){
        $attemptStarted=Get-Date
        Write-HcxDebug "REST attempt start: RequestId=$requestId; Attempt=$attempt/$MaxAttempts; Method=$Method; Uri=$uri; RequestArtifact=$requestArtifact" 'REST'
        $p=@{Method=$Method;Uri=$uri;Headers=$all;SkipCertificateCheck=$true;ErrorAction='Stop';TimeoutSec=90}
        if($script:Hcx.Session){$p.WebSession=$script:Hcx.Session}
        if($null-ne$Body){$p.Body=$Body|ConvertTo-Json -Depth 60 -Compress;$p.ContentType='application/json'}
        # HCX appliances observed in testing can close a pooled connection after prior requests.
        # A fresh HTTP/1.1 connection per attempt prevents a stale keep-alive socket from poisoning discovery.
        if((Get-Command Invoke-RestMethod).Parameters.ContainsKey('DisableKeepAlive')){$p.DisableKeepAlive=$true}
        if((Get-Command Invoke-RestMethod).Parameters.ContainsKey('HttpVersion')){$p.HttpVersion='1.1'}
        try{
            $response=Invoke-RestMethod @p
            $response=ConvertFrom-HcxJsonIfNeeded $response
            $attemptMs=[math]::Round(((Get-Date)-$attemptStarted).TotalMilliseconds,0)
            $totalMs=[math]::Round(((Get-Date)-$overallStarted).TotalMilliseconds,0)
            $responseArtifact=Save-HcxDebugArtifact -Category 'REST-RESPONSE' -Operation $Method -Data $response -Metadata @{RequestId=$requestId;Uri=$uri;Attempt=$attempt;AttemptElapsedMs=$attemptMs;TotalElapsedMs=$totalMs;Success=$true}
            Write-HcxDebug "REST success: RequestId=$requestId; Attempt=$attempt/$MaxAttempts; Method=$Method; Uri=$uri; AttemptElapsedMs=$attemptMs; TotalElapsedMs=$totalMs; ResponseArtifact=$responseArtifact" 'REST'
            return $response
        }catch{
            $lastError=$_;$attemptMs=[math]::Round(((Get-Date)-$attemptStarted).TotalMilliseconds,0);$status='';$detail=$_.Exception.Message
            try{$status=[int]$_.Exception.Response.StatusCode}catch{};try{if($_.ErrorDetails.Message){$detail=[string]$_.ErrorDetails.Message}}catch{}
            $transient=Test-HcxTransientTransportError $_
            Save-HcxDebugArtifact -Category 'REST-ATTEMPT-FAILURE' -Operation $Method -Data ([ordered]@{RequestId=$requestId;Attempt=$attempt;MaxAttempts=$MaxAttempts;Uri=$uri;Status=$status;ElapsedMs=$attemptMs;Transient=$transient;Detail=$detail})|Out-Null
            if($AllowFailure){Write-HcxDebug "Optional REST request unavailable: RequestId=$requestId; Attempt=$attempt; Path=$Path; HTTP=$status; Response=$detail" 'REST-OPTIONAL';return $null}
            if($transient -and $attempt -lt $MaxAttempts){
                $delayMs=250*$attempt
                Log "Transient HCX transport failure. RequestId='$requestId'; Attempt='$attempt/$MaxAttempts'; Path='$Path'; Detail='$detail'. Retrying with a fresh connection." WARN
                Start-Sleep -Milliseconds $delayMs
                continue
            }
            break
        }
    }
    $status='';$detail=$lastError.Exception.Message
    try{$status=[int]$lastError.Exception.Response.StatusCode}catch{};try{if($lastError.ErrorDetails.Message){$detail=[string]$lastError.ErrorDetails.Message}}catch{}
    $totalMs=[math]::Round(((Get-Date)-$overallStarted).TotalMilliseconds,0)
    $exceptionArtifact=Write-HcxExceptionDiagnostic $lastError "REST-$Method"
    $failure=[ordered]@{RequestId=$requestId;Method=$Method;Uri=$uri;Status=$status;Attempts=$MaxAttempts;TotalElapsedMs=$totalMs;ResponseDetail=$detail;RequestArtifact=$requestArtifact;ExceptionArtifact=$exceptionArtifact}
    $failureArtifact=Save-HcxDebugArtifact -Category 'REST-FAILURE' -Operation $Method -Data $failure
    Log "HCX REST FAILURE: RequestId='$requestId'; Method='$Method'; Path='$Path'; HTTP='$status'; Attempts='$MaxAttempts'; TotalElapsedMs='$totalMs'; Response='$detail'; FailureArtifact='$failureArtifact'." ERROR
    throw "HCX REST $Method $Path failed after $MaxAttempts attempt(s). HTTP=$status; Response=$detail; RequestId=$requestId"
}

trap {
    try{$artifact=Write-HcxExceptionDiagnostic $_ 'GLOBAL-TRAP';Log "Unhandled PowerShell error captured. Diagnostic artifact: $artifact" ERROR}catch{}
    continue
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
    $script:Hcx.Version='9.1 session authenticated; version endpoint not queried'
    Log "Authenticated to $hostName using HCX x-hm-authorization. Optional version probes were skipped to avoid unnecessary connection churn." PASS
}
function Disconnect-HcxRest {
    if($script:Hcx.Connected){foreach($p in '/hybridity/api/sessions/current','/api/sessions/current','/api/v1/logout'){try{Invoke-HcxRest DELETE $p -AllowFailure|Out-Null}catch{}}}
    if($script:SourceVIServer){Disconnect-VIServer -Server $script:SourceVIServer -Force -Confirm:$false -ErrorAction SilentlyContinue|Out-Null};if($script:DestinationVIServer -and $script:DestinationVIServer -ne $script:SourceVIServer){Disconnect-VIServer -Server $script:DestinationVIServer -Force -Confirm:$false -ErrorAction SilentlyContinue|Out-Null};$script:SourceVIServer=$null;$script:DestinationVIServer=$null
$script:HcxDiscovery=$null;$script:Hcx.Connected=$false;$script:Hcx.Session=$null;$script:Hcx.Headers=@{};$script:Hcx.Version='';$script:Hcx.SitePairs=@()
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
        Log "Destination datastore-cluster inventory loaded: $($(Get-SafeCount $items))." $(if($(Get-SafeCount $items)){'PASS'}else{'WARN'})
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
    if($script:DestinationVIServer -and $script:DestinationVIServer -ne $script:SourceVIServer){Disconnect-VIServer -Server $script:DestinationVIServer -Force -Confirm:$false -ErrorAction SilentlyContinue|Out-Null}

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
function Set-HcxItemsSource($Control,$Items,[string]$PreferredId='') {
    if($null -eq $Control){return}
    $collection=@($Items)
    $Control.ItemsSource=$collection
    if($collection.Count -eq 0){$Control.SelectedIndex=-1;return}
    if(-not[string]::IsNullOrWhiteSpace($PreferredId)){
        $match=@($collection|Where-Object{[string]$_.Id -eq $PreferredId}|Select-Object -First 1)
        if($match.Count -eq 1){$Control.SelectedItem=$match[0];return}
    }
    if($Control.SelectedIndex -lt 0){$Control.SelectedIndex=0}
}
function Set-HcxBindingContextCollection([string]$Name,$Items) {
    if(-not$script:Window -or -not$script:Window.DataContext){return}
    $property=$script:Window.DataContext.PSObject.Properties[$Name]
    if($property){$property.Value=@($Items)}else{$script:Window.DataContext|Add-Member -NotePropertyName $Name -NotePropertyValue @($Items)}
}
function Bind-DestinationControls {
    try{
        Set-HcxBindingContextCollection 'MigrationTypes' $script:MigrationTypes
        Set-HcxBindingContextCollection 'Sites' $script:Inventory.Sites
        Set-HcxBindingContextCollection 'Computes' $script:Inventory.Computes
        Set-HcxBindingContextCollection 'Datastores' $script:Inventory.Datastores
        Set-HcxBindingContextCollection 'Folders' $script:Inventory.Folders
        Set-HcxBindingContextCollection 'Networks' $script:Inventory.Networks
        Set-HcxBindingContextCollection 'Policies' $script:Inventory.Policies
        Set-HcxItemsSource $script:cmbDestinationSite $script:Inventory.Sites
        Set-HcxItemsSource $script:cmbGlobalCompute $script:Inventory.Computes
        Set-HcxItemsSource $script:cmbGlobalDatastore $script:Inventory.Datastores
        Set-HcxItemsSource $script:cmbGlobalFolder $script:Inventory.Folders
        Set-HcxItemsSource $script:cmbGlobalStoragePolicy $script:Inventory.Policies
        Set-HcxItemsSource $script:cmbVtpmStoragePolicy $script:Inventory.Policies
        Set-HcxItemsSource $script:cmbMigrationType $script:MigrationTypes
        $script:gridVMs.Items.Refresh()
        Write-HcxDebug "Destination controls and row dropdown sources bound. MigrationTypes=$(@($script:MigrationTypes).Count); Sites=$(@($script:Inventory.Sites).Count); Computes=$(@($script:Inventory.Computes).Count); Storage=$(@($script:Inventory.Datastores).Count); Folders=$(@($script:Inventory.Folders).Count); Networks=$(@($script:Inventory.Networks).Count); Policies=$(@($script:Inventory.Policies).Count)." 'UI-BIND'
    }catch{$artifact=Write-HcxExceptionDiagnostic $_ 'BIND-DESTINATION-CONTROLS';Log "Destination control binding failed. Diagnostic=$artifact" ERROR;throw}
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
            if ($(Get-SafeCount $storagePods) -gt 0) {
                Log ('Destination datastore-cluster inventory loaded: {0}.' -f $(Get-SafeCount $storagePods)) PASS
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
        if(Get-Command Bind-DestinationControls -CommandType Function -ErrorAction SilentlyContinue){Bind-DestinationControls}else{throw 'Required function Bind-DestinationControls is not loaded.'}
        if((Get-SafeCount $script:Rows) -gt 0){Apply-GlobalSelection}
        Log "vCenter inventory loaded. SourceVMs=$(@($script:Inventory.VMs).Count); DestinationNetworks=$(@($script:Inventory.Networks).Count); DestinationDatastores=$(@($script:Inventory.Datastores).Count); DestinationComputes=$(@($script:Inventory.Computes).Count); DestinationFolders=$(@($script:Inventory.Folders).Count)" PASS
    }catch{
        Log "vCenter inventory stage failed: $($_.Exception.Message)" ERROR
        throw
    }
}
function Get-SourceVmVtpmStatus($VM){
 try{
  if(-not(Get-Command Get-VTpm -ErrorAction SilentlyContinue)){try{Import-Module VMware.VimAutomation.Security -ErrorAction SilentlyContinue|Out-Null}catch{}}
  if(Get-Command Get-VTpm -ErrorAction SilentlyContinue){$d=@(Get-VTpm -VM $VM -Server $script:SourceVIServer -ErrorAction Stop);return [pscustomobject]@{Value=if($d.Count){1}else{0};Status='Detected';Source='Get-VTpm';Detail="Get-VTpm returned $($d.Count) device(s)."}}
  $v=Get-View -Server $script:SourceVIServer -Id $VM.Id -Property Config.Hardware.Device -ErrorAction Stop;$d=@($v.Config.Hardware.Device|Where-Object{$_.GetType().Name-match'VirtualTPM|VirtualTrustedPlatformModule'});return [pscustomobject]@{Value=if($d.Count){1}else{0};Status='Detected';Source='vCenter Hardware View';Detail="Hardware inventory returned $($d.Count) vTPM device(s)."}
 }catch{$m=$_.Exception.Message;Log "vTPM detection failed for '$($VM.Name)': $m" WARN;return [pscustomobject]@{Value=$null;Status='DetectionFailed';Source='vCenter Inventory';Detail=$m}}
}

function Enrich-SourceVm([object]$NormalizedVm) {
    $vm=$NormalizedVm.Raw
    $nics=@(Get-NetworkAdapter -VM $vm -Server $script:SourceVIServer -ErrorAction SilentlyContinue|ForEach-Object{
        [pscustomobject]@{adapterName=$_.Name;networkName=$_.NetworkName;networkId=if($_.ExtensionData.Backing.Port.PortgroupKey){[string]$_.ExtensionData.Backing.Port.PortgroupKey}else{[string]$_.NetworkName}}
    })
    $datastores=@(Get-Datastore -VM $vm -Server $script:SourceVIServer -ErrorAction SilentlyContinue)
    $vtpm=Get-SourceVmVtpmStatus $vm
    [pscustomobject]@{
        name=$vm.Name;id=$vm.Id;powerState=[string]$vm.PowerState;guestOS=[string]$vm.Guest.OSFullName;cpuCount=$vm.NumCpu;memoryGB=[math]::Round($vm.MemoryGB,2)
        computeName=[string]$vm.VMHost.Parent.Name;computeId=[string]$vm.VMHost.ParentId;folderName=[string]$vm.Folder.Name;folderId=[string]$vm.FolderId
        datastoreName=(@($datastores.Name)-join '; ');datastoreId=(@($datastores.Id)-join '; ');nics=$nics;vtpm=$vtpm.Value;vtpmDetectionStatus=$vtpm.Status;vtpmDetectionSource=$vtpm.Source;vtpmDetectionDetail=$vtpm.Detail
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
            if($(Get-SafeCount $out) -gt 0){Log "Loaded $($(Get-SafeCount $out)) $Kind object(s) from $path." PASS;return @($out)}
            $shape=if($null-eq$raw){'empty'}else{(@($raw.PSObject.Properties.Name)-join ',')}
            $attempts.Add("$path HTTP=$($response.StatusCode) items=$($(Get-SafeCount $items)) usable=$($(Get-SafeCount $out)) shape=$shape")
        }catch{
            $status='';try{$status=[int]$_.Exception.Response.StatusCode}catch{}
            $detail=$_.Exception.Message;try{if($_.ErrorDetails.Message){$detail=$_.ErrorDetails.Message}}catch{}
            $attempts.Add("$path HTTP=$status $detail")
        }
    }
    foreach($a in $attempts){Log "$Kind discovery: $a" WARN}
    Log "No usable $Kind inventory was returned. Attempted $($(Get-SafeCount $Paths)) HCX resource path(s)." WARN
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
function Convert-HcxNetworkIdCanonical([string]$Id) {
    if([string]::IsNullOrWhiteSpace($Id)){return ''}
    if($Id -match '(/infra/segments/[^/]+)$'){return $Matches[1]}
    if($Id -match '(dvportgroup-\d+)$'){return $Matches[1]}
    if($Id -match '(network-\d+)$'){return $Matches[1]}
    if($Id -match '(opaqueNetwork-[^/]+)$'){return $Matches[1]}
    return $Id.Trim()
}
function Get-HcxNetworkTypeRank([string]$EntityType) {
    switch -Regex ($EntityType) {
        '^DistributedVirtualPortgroup$' { return 10 }
        '^StandardPortgroup$|^Network$' { return 20 }
        '^NsxtSegment$' { return 30 }
        '^VirtualWire$' { return 40 }
        '^OpaqueNetwork$' { return 50 }
        default { return 99 }
    }
}
function Get-HcxImportedNetworkMapping([string]$SourceName){
    if([string]::IsNullOrWhiteSpace($SourceName)){return $null}
    $key=$SourceName.Trim().ToLowerInvariant()
    if($script:ImportedNetworkMappings.ContainsKey($key)){return $script:ImportedNetworkMappings[$key]}
    return $null
}
function Resolve-HcxDestinationNetworkByImportedName([string]$DestinationName){
    $matches=@($script:Inventory.Networks|Where-Object{[string]$_.Name -ieq $DestinationName})
    if($matches.Count -eq 0){return [pscustomobject]@{Network=$null;Status='UnresolvedDestination';Detail="No destination network named '$DestinationName' exists in current inventory."}}
    $collapsed=@($matches|Group-Object{Convert-HcxNetworkIdCanonical ([string]$_.Id)}|ForEach-Object{$_.Group|Sort-Object @{Expression={Get-HcxNetworkTypeRank ([string]$_.EntityType)}}|Select-Object -First 1})
    if($collapsed.Count -eq 1){return [pscustomobject]@{Network=$collapsed[0];Status='Imported Mapping';Detail='Destination resolved from imported network-mapping CSV.'}}
    $detail=@($collapsed|ForEach-Object{"$($_.Name) [$($_.Id),$($_.EntityType)]"})-join'; '
    return [pscustomobject]@{Network=$null;Status='AmbiguousDestination';Detail="Destination name '$DestinationName' resolves to multiple distinct networks: $detail"}
}
function Apply-ImportedNetworkMappingsToRows {
    $applied=0;$unresolved=[Collections.Generic.List[string]]::new()
    foreach($row in @($script:Rows)){
        foreach($mapping in @($row.NicMappings)){
            $imported=Get-HcxImportedNetworkMapping ([string]$mapping.SourceNetworkName)
            if(-not$imported){continue}
            $resolved=Resolve-HcxDestinationNetworkByImportedName ([string]$imported.DestinationNetworkName)
            if($resolved.Network){
                $mapping.DestinationNetworkName=[string]$resolved.Network.Name
                $mapping.DestinationNetworkId=Convert-HcxNetworkIdCanonical ([string]$resolved.Network.Id)
                $mapping.DestinationNetworkType=[string]$resolved.Network.EntityType
                $mapping.MatchStatus='Imported Mapping'
                $mapping.MatchDetail="Imported CSV: $($imported.SourceNetworkName) -> $($imported.DestinationNetworkName)"
                $mapping.MappingSource='Imported Mapping'
                $applied++
            }else{
                $mapping.DestinationNetworkName='';$mapping.DestinationNetworkId='';$mapping.DestinationNetworkType=''
                $mapping.MatchStatus=[string]$resolved.Status;$mapping.MatchDetail=[string]$resolved.Detail;$mapping.MappingSource='Imported Mapping - Unresolved'
                $unresolved.Add("$($row.VMName) / $($mapping.Adapter): $($resolved.Detail)")
            }
        }
        Update-NicSummary $row
    }
    if($script:gridVMs){$script:gridVMs.Items.Refresh()}
    Invalidate-Validation 'Imported network mapping CSV applied'
    Log "Imported network mappings applied to $applied VM NIC mapping(s)." PASS
    if($unresolved.Count){foreach($item in $unresolved){Log $item ERROR}}
    return [pscustomobject]@{Applied=$applied;Unresolved=@($unresolved)}
}
function Import-HcxNetworkMappingCsv {
    $dialog=[Microsoft.Win32.OpenFileDialog]::new();$dialog.Filter='CSV files (*.csv)|*.csv';$dialog.InitialDirectory=$script:OutputBase
    if(-not$dialog.ShowDialog()){return}
    $rows=@(Import-Csv -LiteralPath $dialog.FileName)
    if($rows.Count -eq 0){throw 'The selected network-mapping CSV contains no rows.'}
    $map=@{};$normalized=[Collections.Generic.List[object]]::new();$issues=[Collections.Generic.List[string]]::new();$line=2
    foreach($item in $rows){
        $source=[string](Get-Value $item @('SourceNetworkName','SourceNetwork','Source','SourcePortGroup'))
        $destination=[string](Get-Value $item @('DestinationNetworkName','DestinationNetwork','Destination','DestinationPortGroup'))
        $source=$source.Trim();$destination=$destination.Trim()
        if(-not$source){$issues.Add("CSV line $line has no SourceNetworkName.");$line++;continue}
        if(-not$destination){$issues.Add("CSV line $line for source '$source' has no DestinationNetworkName.");$line++;continue}
        $key=$source.ToLowerInvariant()
        if($map.ContainsKey($key)){
            if([string]$map[$key].DestinationNetworkName -ine $destination){$issues.Add("Conflicting mappings for source '$source': '$($map[$key].DestinationNetworkName)' and '$destination'.")}
            else{$issues.Add("Duplicate mapping for source '$source' and destination '$destination'.")}
            $line++;continue
        }
        $resolved=Resolve-HcxDestinationNetworkByImportedName $destination
        if(-not$resolved.Network){$issues.Add("CSV line ${line}: $($resolved.Detail)");$line++;continue}
        $entry=[pscustomobject]@{SourceNetworkName=$source;DestinationNetworkName=[string]$resolved.Network.Name;DestinationNetworkId=Convert-HcxNetworkIdCanonical ([string]$resolved.Network.Id);DestinationNetworkType=[string]$resolved.Network.EntityType;SourceCsvLine=$line}
        $map[$key]=$entry;$normalized.Add($entry);$line++
    }
    if($issues.Count){throw "Network-mapping CSV validation failed:`n`n$($issues -join [Environment]::NewLine)"}
    $script:ImportedNetworkMappings=$map;$script:ImportedNetworkMappingRows=@($normalized);$script:NetworkMappingCsvPath=$dialog.FileName
    $script:txtNetworkMappingCsv.Text=$dialog.FileName;$script:lblNetworkMappingStatus.Text="$($normalized.Count) mapping(s) loaded";$script:lblNetworkMappingStatus.Foreground='LightGreen'
    $result=Apply-ImportedNetworkMappingsToRows
    $summary="Imported $($normalized.Count) source-to-destination network mapping(s).`nApplied to $($result.Applied) currently loaded VM NIC(s).`nUnresolved VM NICs: $(@($result.Unresolved).Count).`n`nImported mappings take precedence over automatic exact-name matching. Per-VM Configure selections remain the final override."
    [Windows.MessageBox]::Show($summary,'Network Mapping CSV Imported')|Out-Null
}
function Clear-HcxNetworkMappingCsv {
    $script:ImportedNetworkMappings=@{};$script:ImportedNetworkMappingRows=@();$script:NetworkMappingCsvPath=''
    $script:txtNetworkMappingCsv.Text='';$script:lblNetworkMappingStatus.Text='Not loaded';$script:lblNetworkMappingStatus.Foreground='#76C7D8'
    foreach($row in @($script:Rows)){foreach($mapping in @($row.NicMappings)){if([string]$mapping.MappingSource -like 'Imported Mapping*'){$resolution=Resolve-HcxDestinationNetworkMatch -SourceName $mapping.SourceNetworkName -SourceId $mapping.SourceNetworkId;$dst=$resolution.Network;$mapping.DestinationNetworkName=if($dst){$dst.Name}else{''};$mapping.DestinationNetworkId=if($dst){Convert-HcxNetworkIdCanonical $dst.Id}else{''};$mapping.DestinationNetworkType=if($dst){$dst.EntityType}else{''};$mapping.MatchStatus=$resolution.Status;$mapping.MatchDetail=$resolution.Detail;$mapping.MappingSource=if($dst){'Automatic Exact Match'}else{'Unresolved'}}};Update-NicSummary $row}
    $script:gridVMs.Items.Refresh();Invalidate-Validation 'Imported network mapping CSV cleared';Log 'Imported network mapping CSV cleared; applicable VM NICs returned to automatic matching.' INFO
}

function Resolve-HcxDestinationNetworkMatch([string]$SourceName,[string]$SourceId='') {
    $sourceCanonical=Convert-HcxNetworkIdCanonical $SourceId
    $nameMatches=@($script:Inventory.Networks|Where-Object{[string]$_.Name -ieq $SourceName})
    if($nameMatches.Count -eq 0){return [pscustomobject]@{Network=$null;Status='Unresolved';Detail="No destination network named '$SourceName' was found."}}
    # Collapse duplicate PowerCLI and HCX representations of the same destination object.
    $collapsed=@(
        $nameMatches |
            Group-Object { Convert-HcxNetworkIdCanonical ([string]$_.Id) } |
            ForEach-Object {
                $groupItems=@($_.Group)
                $groupItems |
                    Sort-Object @{Expression={Get-HcxNetworkTypeRank ([string]$_.EntityType)}} |
                    Select-Object -First 1
            }
    )
    if($collapsed.Count -eq 1){return [pscustomobject]@{Network=$collapsed[0];Status='ExactName';Detail='Unique destination name after canonical-ID deduplication.'}}
    # If source and destination MoRefs happen to match, prefer that representation.
    if($sourceCanonical){
        $idMatches=@($collapsed|Where-Object{(Convert-HcxNetworkIdCanonical ([string]$_.Id)) -eq $sourceCanonical})
        if($idMatches.Count -eq 1){return [pscustomobject]@{Network=$idMatches[0];Status='ExactId';Detail='Destination selected by canonical network ID.'}}
    }
    # Prefer a non-uplink HCX inventory object. Uplink/trunk objects are never valid VM destinations.
    $deployable=@($collapsed|Where-Object{
        $raw=$_.Raw;$obj=if($raw){Get-HcxPropertyValue $raw 'object'}else{$null};$tags=if($obj){Get-HcxPropertyValue $obj 'hybridityTags'}else{$null}
        $uplink=$false;$trunk=$false
        if($tags){$u=Get-HcxPropertyValue $tags 'UPLINK';$t=Get-HcxPropertyValue $tags 'FLEET_TRUNK';if($u){$uplink=[bool](Get-HcxPropertyValue $u 'value')};if($t){$trunk=[bool](Get-HcxPropertyValue $t 'value')}}
        -not$uplink -and -not$trunk
    })
    if($deployable.Count -eq 1){return [pscustomobject]@{Network=$deployable[0];Status='UniqueDeployable';Detail='One non-uplink/non-trunk destination remained.'}}
    $detail=@($collapsed|ForEach-Object{"$($_.Name) [$($_.Id),$($_.EntityType)]"})-join'; '
    return [pscustomobject]@{Network=$null;Status='Ambiguous';Detail="Multiple distinct destination networks share '$SourceName': $detail"}
}
function Convert-ToNicMappings($VmRaw) {
    $result=[Collections.ObjectModel.ObservableCollection[object]]::new();$index=0
    foreach($nic in @(Get-NestedNics $VmRaw)){
        $index++
        $srcName=[string](Get-Value $nic @('networkName','sourceNetworkName','name','portGroupName','backingName'))
        $srcId=Convert-HcxNetworkIdCanonical ([string](Get-Value $nic @('networkId','sourceNetworkId','id','backingId','portGroupId')))
        $adapter=[string](Get-Value $nic @('adapterName','deviceName','label','name'));if(-not$adapter){$adapter="Network adapter $index"}
        $imported=Get-HcxImportedNetworkMapping $srcName
        if($imported){$resolution=Resolve-HcxDestinationNetworkByImportedName ([string]$imported.DestinationNetworkName)}else{$resolution=Resolve-HcxDestinationNetworkMatch -SourceName $srcName -SourceId $srcId}
        $dst=$resolution.Network
        if($dst){
            $dst.Id=Convert-HcxNetworkIdCanonical ([string]$dst.Id)
            Write-HcxDebug "Automatic network match: Source='$srcName' [$srcId]; Destination='$($dst.Name)' [$($dst.Id),$($dst.EntityType)]; Status='$($resolution.Status)'; Detail='$($resolution.Detail)'." 'NETWORK-MATCH'
        }else{
            Log "Automatic network match not selected: Source='$srcName' [$srcId]; Status='$($resolution.Status)'; Detail='$($resolution.Detail)'." WARN
        }
        $result.Add([pscustomobject]@{
            Adapter=$adapter;SourceNetworkName=$srcName;SourceNetworkId=$srcId
            DestinationNetworkName=if($dst){[string]$dst.Name}else{''};DestinationNetworkId=if($dst){[string]$dst.Id}else{''}
            DestinationNetworkType=if($dst){[string]$dst.EntityType}else{''};MatchStatus=[string]$resolution.Status;MatchDetail=[string]$resolution.Detail;MappingSource=if($imported){'Imported Mapping'}elseif($dst){'Automatic Exact Match'}else{'Unresolved'}
        })
    }
    return $result
}
function Update-NicSummary($Row) {
    $mappings = @($Row.NicMappings)
    if ((Get-SafeCount $mappings) -eq 0) {
        $Row.NicSummary = 'No NICs discovered'
        return
    }
    $mappedCount = @($mappings | Where-Object { $_.DestinationNetworkId }).Count
    $primary = $mappings[0]
    $destination = if ($primary.DestinationNetworkName) { $primary.DestinationNetworkName } else { 'Unresolved' }
    $Row.NicSummary = '{0} of {1} mapped | {2} -> {3}' -f $mappedCount, @($mappings).Count, $primary.SourceNetworkName, $destination
}
function Convert-ToVmRow([string]$Name) {
    $matches=@($script:Inventory.VMs|Where-Object{$_.Name -ieq $Name})
    $vm=if(@($matches).Count -eq 1){$matches[0]}else{$null};$raw=if($vm){Enrich-SourceVm $vm}else{$null}
    $nics=if($raw){@(Convert-ToNicMappings $raw)}else{@()}
    $sourceNetworks=@($nics|ForEach-Object{$_.SourceNetworkName}|Where-Object{$_}) -join '; '
    [pscustomobject]@{
        Include=$true;MobilityGroupNumber=1;Status=if($vm){'Discovered'}elseif(@($matches).Count -gt 1){'Ambiguous'}else{'NotFound'};VMName=$Name;VMId=if($vm){$vm.Id}else{''}
        PowerState=[string](Get-Value $raw @('powerState','state'));GuestOS=[string](Get-Value $raw @('guestOS','guestFullName','osName'))
        CPU=[string](Get-Value $raw @('cpuCount','numCpu','vCpu'));MemoryGB=[string](Get-Value $raw @('memoryGB','memoryGb','memory'))
        SourceCompute=[string](Get-Value $raw @('computeName','clusterName','sourceComputeName'));SourceComputeId=[string](Get-Value $raw @('computeId','clusterId','sourceComputeId'))
        SourceFolder=[string](Get-Value $raw @('folderName','sourceFolderName'));SourceFolderId=[string](Get-Value $raw @('folderId','sourceFolderId'))
        SourceDatastore=[string](Get-Value $raw @('datastoreName','sourceDatastoreName'));SourceDatastoreId=[string](Get-Value $raw @('datastoreId','sourceDatastoreId'))
        SourceNetworks=$sourceNetworks;NicCount=@($nics).Count;NicSummary=if(@($nics).Count -gt 0){$primary=$nics[0];$destination=if($primary.DestinationNetworkName){$primary.DestinationNetworkName}else{'Unresolved'};'{0} of {1} mapped | {2} -> {3}' -f @($nics|Where-Object{$_.DestinationNetworkId}).Count,$(Get-SafeCount $nics),$primary.SourceNetworkName,$destination}else{'No NICs discovered'};NicMappings=$nics
        MigrationType='HCX Assisted vMotion (Direct)';DestinationSite='';DestinationSiteId='';DestinationCompute='';DestinationComputeId='';DestinationComputeType='';DestinationFolder='';DestinationFolderId=''
        DestinationDatastore='';DestinationDatastoreId='';DestinationStorageType='';StoragePolicy='';Vtpm=if($raw-and$null-ne$raw.vtpm){[int]$raw.vtpm}else{$null};VtpmDetectionStatus=if($raw){$raw.vtpmDetectionStatus}else{'NotAvailable'};VtpmDetectionSource=if($raw){$raw.vtpmDetectionSource}else{'VM Not Found'};VtpmDetectionDetail=if($raw){$raw.vtpmDetectionDetail}else{'VM was not found.'};StoragePolicyAssignment=if($raw-and[int]$raw.vtpm-eq1){'vTPM Global'}else{'Standard Global'};DiskFormat='Same format as source';DiskFormatSelectionSource='Phase I Global';DatastoreSelectionSource='';ValidationMessage='Not validated';ValidatedOn=''
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
    $dialog=[Microsoft.Win32.OpenFileDialog]::new();$dialog.Filter='JSON files (*.json)|*.json';$dialog.InitialDirectory=$script:OutputBase
    if(-not$dialog.ShowDialog()){return}
    try{
        $profile=Get-Content -LiteralPath $dialog.FileName -Raw|ConvertFrom-Json -Depth 20
        $script:txtHcx.Text=[string](Get-Value $profile @('SourceHCXManager','HCXManager','HcxManager'))
        $script:txtUser.Text=[string](Get-Value $profile @('HCXUsername','HcxUsername','Username'))
        $script:txtSourceVC.Text=[string](Get-Value $profile @('SourceVCenter','SourceVC','SourceVCenterServer'))
        $script:txtSourceVCUser.Text=[string](Get-Value $profile @('SourceVCenterUsername','SourceVCUsername','SourceUsername'))
        $script:txtDestinationVC.Text=[string](Get-Value $profile @('DestinationVCenter','DestinationVC','DestinationVCenterServer'))
        $script:txtDestinationVCUser.Text=[string](Get-Value $profile @('DestinationVCenterUsername','DestinationVCUsername','DestinationUsername'))
        Reset-HcxAutomaticDiscovery
        Log "Connection profile loaded: $($dialog.FileName). Passwords were not loaded." PASS
    }catch{$artifact=Write-HcxExceptionDiagnostic $_ 'LOAD-CONNECTION-PROFILE';Log "Connection profile load failed. Diagnostic=$artifact" ERROR;throw}
}
function Invalidate-Validation {
    param([string]$Reason='Configuration changed')
    $script:ValidationCurrent=$false
    if($script:btnCreateCsv){$script:btnCreateCsv.IsEnabled=$false}
    if($script:lblValidation){$script:lblValidation.Text='Not validated';$script:lblValidation.Foreground='#76C7D8'}
    Write-HcxDebug "Phase I validation invalidated. Reason='$Reason'." 'VALIDATION'
    Log "Phase I validation invalidated: $Reason." INFO
}
function Invalidate-P2Validation {
    param([string]$Reason='Phase II configuration changed')
    $script:P2Valid=$false
    $script:P2Payload=$null
    if($script:btnP2Preview){$script:btnP2Preview.IsEnabled=$false}
    if($script:btnP2Create){$script:btnP2Create.IsEnabled=$false}
    if($script:lblP2Status){$script:lblP2Status.Text='Not validated';$script:lblP2Status.Foreground='#76C7D8'}
    Write-HcxDebug "Phase II validation invalidated. Reason='$Reason'." 'VALIDATION'
}
function Commit-Grid {
    param([System.Windows.Controls.DataGrid]$Grid=$script:gridVMs)
    if($null -eq $Grid){return}
    try{
        # Commit the active cell first and then the row so bound values are written back.
        $null=$Grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell,$true)
        $null=$Grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row,$true)
        $binding=[System.Windows.Data.BindingOperations]::GetBindingExpressionBase($Grid,[System.Windows.Controls.ItemsControl]::ItemsSourceProperty)
        if($binding){$binding.UpdateSource()}
        Write-HcxDebug 'VM grid edit state committed successfully.' 'UI-GRID'
    }catch{
        $artifact=Write-HcxExceptionDiagnostic $_ 'COMMIT-VM-GRID'
        Log "VM grid commit failed. Diagnostic=$artifact" ERROR
        throw
    }
}
function Commit-P2Grid {
    param([System.Windows.Controls.DataGrid]$Grid=$script:gridP2)
    if($null -eq $Grid){return}
    try{
        $null=$Grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Cell,$true)
        $null=$Grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row,$true)
        Write-HcxDebug 'Phase II grid edit state committed successfully.' 'UI-GRID'
    }catch{
        $artifact=Write-HcxExceptionDiagnostic $_ 'COMMIT-PHASE2-GRID'
        Log "Phase II grid commit failed. Diagnostic=$artifact" ERROR
        throw
    }
}
function Apply-GlobalSelection {
    Commit-Grid
    $ds=$script:cmbGlobalDatastore.SelectedItem;$sp=$script:cmbGlobalStoragePolicy.SelectedItem;$vsp=$script:cmbVtpmStoragePolicy.SelectedItem;$co=$script:cmbGlobalCompute.SelectedItem;$fo=$script:cmbGlobalFolder.SelectedItem;$site=$script:cmbDestinationSite.SelectedItem
    $migration=[string]$script:cmbMigrationType.SelectedItem
    $diskFormat=if($script:cmbGlobalDiskFormat.SelectedItem -is [Windows.Controls.ComboBoxItem]){[string]$script:cmbGlobalDiskFormat.SelectedItem.Content}else{[string]$script:cmbGlobalDiskFormat.Text}
    if([string]::IsNullOrWhiteSpace($migration)){$migration=[string]$script:cmbMigrationType.Text}
    foreach($row in @($script:Rows)){
        if($ds){$row.DestinationDatastore=[string]$ds.Name;$row.DestinationDatastoreId=[string]$ds.Id;$row.DestinationStorageType=[string]$ds.Type;$row.DatastoreSelectionSource='Global'}
        if($row.StoragePolicyAssignment-ne'Per-VM Override'){if([int]$row.Vtpm-eq1){if($vsp){$row.StoragePolicy=$vsp.Name;$row.StoragePolicyAssignment='vTPM Global'}}elseif($sp){$row.StoragePolicy=$sp.Name;$row.StoragePolicyAssignment='Standard Global'}}
        if($co){$row.DestinationCompute=[string]$co.Name;$row.DestinationComputeId=[string]$co.Id;$row.DestinationComputeType=[string]$co.Type}
        if($fo){$row.DestinationFolder=[string]$fo.Name;$row.DestinationFolderId=[string]$fo.Id}
        if($site){$row.DestinationSite=[string]$site.Name;$row.DestinationSiteId=[string]$site.Id}
        if($migration){$row.MigrationType=$migration}
        if($diskFormat){$row.DiskFormat=$diskFormat;$row.DiskFormatSelectionSource='Phase I Global'}
        Write-HcxDebug "Applied row defaults: VM='$($row.VMName)'; Migration='$($row.MigrationType)'; Compute='$($row.DestinationCompute)' [$($row.DestinationComputeId)]; Storage='$($row.DestinationDatastore)' [$($row.DestinationDatastoreId)]; Folder='$($row.DestinationFolder)' [$($row.DestinationFolderId)]; Policy='$($row.StoragePolicy)'." 'ROW-DEFAULTS'
    }
    $script:gridVMs.Items.Refresh()
    Invalidate-Validation 'Global destination settings applied'
    Write-HcxDebug "Global destination values applied to $(@($script:Rows).Count) VM row(s)." 'PHASE1-IMPORT'
}
function Show-NicMappingDialog($Row) {
    $x=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Configure VM Network Mappings" Height="520" Width="1080" WindowStartupLocation="CenterOwner" Background="#071015" Foreground="#E6E6E6" FontFamily="Segoe UI" WindowState="Maximized" MinHeight="720" MinWidth="1200">
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
            if($selected){$m.DestinationNetworkName=$selected.Name;$m.DestinationNetworkType=$selected.EntityType;$m.MatchStatus='Per-VM Override';$m.MatchDetail='Selected by operator in per-VM network mapping dialog.';$m.MappingSource='Per-VM Override'}else{$m.DestinationNetworkName='';$m.DestinationNetworkType='';$m.MatchStatus='Unresolved';$m.MappingSource='Unresolved'}
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
    if((Get-SafeCount $script:Rows) -eq 0){$errors.Add('No VMs are loaded.')}
    foreach($r in @($script:Rows)){
        $rowErrors=[Collections.Generic.List[string]]::new()
        if(-not$r.Include){$r.Status='Excluded';$r.ValidationMessage='Excluded by operator';continue}
        $gn=0;if(-not[int]::TryParse([string]$r.MobilityGroupNumber,[ref]$gn)-or$gn-lt1-or$gn-gt 50){$rowErrors.Add('MobilityGroupNumber must be 1 through 50.')}
        $key=$r.VMName.ToLowerInvariant();if($names.ContainsKey($key)){$rowErrors.Add('Duplicate VM name.')}else{$names[$key]=$true}
        $live=@($script:Inventory.VMs|Where-Object{$_.Name -ieq $r.VMName})
        if(@($live).Count -eq 0){$rowErrors.Add('VM is not present in current HCX inventory.')}elseif(@($live).Count -gt 1){$rowErrors.Add('VM name is ambiguous in HCX inventory.')}elseif($live[0].Id -ne $r.VMId){$rowErrors.Add('HCX VM identifier changed; refresh discovery.')}
        $siteObj=@($script:Inventory.Sites|Where-Object Id -eq $r.DestinationSiteId|Select-Object -First 1);if($siteObj){$r.DestinationSite=$siteObj.Name}
        $computeObj=@($script:Inventory.Computes|Where-Object Id -eq $r.DestinationComputeId|Select-Object -First 1);if($computeObj){$r.DestinationCompute=$computeObj.Name;$r.DestinationComputeType=$computeObj.Type}
        $folderObj=@($script:Inventory.Folders|Where-Object Id -eq $r.DestinationFolderId|Select-Object -First 1);if($folderObj){$r.DestinationFolder=$folderObj.Name}
        $dsObj=@($script:Inventory.Datastores|Where-Object Id -eq $r.DestinationDatastoreId|Select-Object -First 1);if($dsObj){$r.DestinationDatastore=$dsObj.Name;$r.DestinationStorageType=$dsObj.Type;if($r.DatastoreSelectionSource -ne 'Global'){$r.DatastoreSelectionSource='Manual'}}
        if(-not$r.DestinationSiteId){$rowErrors.Add('Destination site is not selected.')}
        if(-not$r.DestinationComputeId){$rowErrors.Add('Destination compute is not selected.')}
        if(-not$r.DestinationDatastoreId){$rowErrors.Add('Destination datastore is not selected.')}
        if(-not$r.MigrationType){$rowErrors.Add('Migration type is not selected.')}
        if([string]::IsNullOrWhiteSpace([string]$r.DiskFormat)){$rowErrors.Add('Disk format is not selected.')}
        if($r.Status-eq'Discovered'-and$r.VtpmDetectionStatus-ne'Detected'){$rowErrors.Add("Automatic vTPM detection failed: $($r.VtpmDetectionDetail)")}
        if([string]::IsNullOrWhiteSpace([string]$r.StoragePolicy)){$rowErrors.Add('Destination storage policy is not selected.')}elseif(-not@($script:Inventory.Policies|Where-Object Name -ieq $r.StoragePolicy).Count){$rowErrors.Add('Destination storage policy is not in current inventory.')}
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
            DestinationFolder=$r.DestinationFolder;DestinationFolderId=$r.DestinationFolderId;DestinationDatastore=$r.DestinationDatastore;DestinationDatastoreId=$r.DestinationDatastoreId;DestinationStorageType=$r.DestinationStorageType;DestinationStoragePolicy=$r.StoragePolicy;vTPM=if($null-ne$r.Vtpm){[int]$r.Vtpm}else{''};VtpmDetectionStatus=$r.VtpmDetectionStatus;VtpmDetectionSource=$r.VtpmDetectionSource;VtpmDetectionDetail=$r.VtpmDetectionDetail;StoragePolicyAssignment=$r.StoragePolicyAssignment;DiskFormat=$r.DiskFormat;DiskFormatSelectionSource=$r.DiskFormatSelectionSource
            DestinationNetworkMappingsJson=(@($r.NicMappings|ForEach-Object{[ordered]@{Adapter=$_.Adapter;SourceNetworkName=$_.SourceNetworkName;SourceNetworkId=$_.SourceNetworkId;DestinationNetworkName=$_.DestinationNetworkName;DestinationNetworkId=$_.DestinationNetworkId;DestinationNetworkType=$_.DestinationNetworkType;MatchStatus=$_.MatchStatus;MatchDetail=$_.MatchDetail;MappingSource=$_.MappingSource}})|ConvertTo-Json -Depth 10 -Compress)
            NetworkMappingCsvPath=$script:NetworkMappingCsvPath;ValidationStatus=$r.Status;ValidatedOn=$r.ValidatedOn
        }
    }
    $out|Export-Csv -LiteralPath $save.FileName -NoTypeInformation -Encoding utf8BOM
    $script:LastCsv=$save.FileName;Log "Phase I CSV created: $($save.FileName)" PASS;[Windows.MessageBox]::Show("CSV created successfully.`n`n$($save.FileName)",'CSV created')|Out-Null
}


# PHASE II
$script:P2Rows=[Collections.ObjectModel.ObservableCollection[object]]::new();$script:P2Payload=$null;$script:P2Valid=$false;$script:P2Computes=@();$script:P2Storages=@()
function P2-Invalidate{$script:P2Valid=$false;$script:btnP2Create.IsEnabled=$false;$script:btnP2Preview.IsEnabled=$false;$script:lblP2Status.Text='Not validated';$script:lblP2Status.Foreground='#76C7D8'}
function P2-Import{
 $d=[Microsoft.Win32.OpenFileDialog]::new();$d.Filter='CSV (*.csv)|*.csv';$d.InitialDirectory=$script:OutputBase
 if(-not$d.ShowDialog()){return}
 try{
  $csv=@(Import-Csv -LiteralPath $d.FileName)
  if((Get-SafeCount $csv)-eq 0){throw 'CSV is empty.'}
  $newRows=[Collections.Generic.List[object]]::new();$rowNumber=1
  foreach($r in @($csv)){
   try{
    if([string]$r.ValidationStatus -ne 'Pass'){throw "$($r.VMName) did not pass Phase I."}
    $net=@()
    if(-not[string]::IsNullOrWhiteSpace([string]$r.DestinationNetworkMappingsJson)){$net=@($r.DestinationNetworkMappingsJson|ConvertFrom-Json -Depth 30)}
    $newRows.Add([pscustomobject]@{Include=$true;VMName=$r.VMName;VMId=$r.VMId;Compute=$r.DestinationCompute;ComputeId=$r.DestinationComputeId;ComputeType=if($r.DestinationComputeType){$r.DestinationComputeType}else{'cluster'};MobilityGroupNumber=if($r.MobilityGroupNumber){[int]$r.MobilityGroupNumber}else{1};Folder=$r.DestinationFolder;FolderId=$r.DestinationFolderId;StorageType=if($r.DestinationStorageType){$r.DestinationStorageType}else{'datastore'};Storage=$r.DestinationDatastore;StorageId=$r.DestinationDatastoreId;StoragePolicy=if($r.DestinationStoragePolicy){[string]$r.DestinationStoragePolicy}else{''};Vtpm=if(([string]$r.vTPM).Trim()-eq'1'){1}else{0};StoragePolicyAssignment=if($r.StoragePolicyAssignment){$r.StoragePolicyAssignment}else{'Standard Global'};StorageSelectionSource='Phase I Import';DiskFormat=if($r.DiskFormat){[string]$r.DiskFormat}else{'Same format as source'};DiskFormatSelectionSource='Phase I Import';Networks=@($net);NetworkSummary=(@($net|ForEach-Object{"$($_.SourceNetworkName) -> $($_.DestinationNetworkName)"})-join'; ')})
   }catch{throw "Phase II CSV row $rowNumber failed. VM='$($r.VMName)'. $(Get-HcxDetailedError $_)"}
   $rowNumber++
  }
  $script:P2Rows.Clear();foreach($row in $newRows){$script:P2Rows.Add($row)}
  $script:P2DestinationSegments=@();$script:gridP2.ItemsSource=$script:P2Rows
  $groups=@($script:P2Rows|ForEach-Object{[int]$_.MobilityGroupNumber})
  $max=if((Get-SafeCount $groups)-gt0){[int](($groups|Measure-Object -Maximum).Maximum)}else{1};if($max-lt1){$max=1};$script:cmbP2GroupCount.SelectedItem=$max
  $script:txtP2Csv.Text=$d.FileName;Restore-P2SelectionsFromPhaseICsv
  $script:txtP2Hcx.Text=[string]$csv[0].HCXManager;$script:txtP2DestSite.Text=[string]$csv[0].DestinationSite
  if($script:txtSourceVC -and -not[string]::IsNullOrWhiteSpace($script:txtSourceVC.Text)){$script:txtP2SourceSite.Text=$script:txtSourceVC.Text.Trim()}elseif($script:SourceVIServer){$script:txtP2SourceSite.Text=[string]$script:SourceVIServer.Name}
  if(-not$script:txtP2Name.Text){$script:txtP2Name.Text='HCX Mobility Group '+(Get-Date -Format yyyyMMdd-HHmm)}
  if($script:DestinationVIServer){P2-Inventory;Restore-P2SelectionsFromPhaseICsv}else{Log 'Destination vCenter is not connected. Connect and load inventory, then select Load Storage.' WARN}
  Resolve-P2ImportedNetworks;P2-Invalidate
  Log "Phase II imported $(Get-SafeCount $csv) VM(s). Source site was populated from the source vCenter connection." PASS
 }catch{$detail=Get-HcxDetailedError $_;Log "PHASE II IMPORT FAILURE: $detail" ERROR;throw}
}
function Convert-P2NetworkMoRef([string]$Id){
 if([string]::IsNullOrWhiteSpace($Id)){return ''}
 if($Id -match '(/infra/segments/[^/]+)$'){return $Matches[1]}
 if($Id -match '(dvportgroup-\d+)$'){return $Matches[1]}
 if($Id -match '(network-\d+)$'){return $Matches[1]}
 return $Id
}
function Get-P2NetworkEntityType([string]$Id){
 $v=Convert-P2NetworkMoRef $Id
 if($v -match '^/infra/segments/[^/]+$'){return 'NsxtSegment'}
 if($v -match '^dvportgroup-\d+$'){return 'DistributedVirtualPortgroup'}
 if($v -match '^network-\d+$'){return 'Network'}
 return ''
}
function Resolve-P2ImportedNetworks{
 $issues=[Collections.Generic.List[string]]::new();$resolved=0
 foreach($row in @($script:P2Rows)){
  foreach($mapping in @($row.Networks)){
   $srcId=Convert-P2NetworkMoRef ([string]$mapping.SourceNetworkId)
   $dstId=Convert-P2NetworkMoRef ([string]$mapping.DestinationNetworkId)
   $srcType=Get-P2NetworkEntityType $srcId;$dstType=Get-P2NetworkEntityType $dstId
   if(-not$srcType){$issues.Add("$($row.VMName): unsupported source network ID '$srcId'.")}
   if(-not$dstType){$issues.Add("$($row.VMName): unsupported destination network ID '$dstId'.")}
   if($srcType -and $dstType){
    $mapping.SourceNetworkId=$srcId;$mapping.DestinationNetworkId=$dstId
    if($mapping.PSObject.Properties['SourceNetworkType']){$mapping.SourceNetworkType=$srcType}else{$mapping|Add-Member -NotePropertyName SourceNetworkType -NotePropertyValue $srcType}
    if($mapping.PSObject.Properties['DestinationNetworkType']){$mapping.DestinationNetworkType=$dstType}else{$mapping|Add-Member -NotePropertyName DestinationNetworkType -NotePropertyValue $dstType}
    $resolved++
   }
  }
  $row.NetworkSummary=(@($row.Networks|ForEach-Object{"$($_.SourceNetworkName) -> $($_.DestinationNetworkName)"})-join'; ')
 }
 $script:gridP2.Items.Refresh()
 if($issues.Count){foreach($x in $issues){Log $x ERROR};throw ($issues-join[Environment]::NewLine)}
 Log "Phase II network check passed for $(@($script:P2Rows).Count) VM(s); $resolved mapping(s) retained Phase I inventory IDs." PASS
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
        $value=Get-P2CsvValue $item @('DestinationStoragePolicy','StoragePolicy','StoragePolicyName');if($value){$row.StoragePolicy=[string]$value};$value=Get-P2CsvValue $item @('vTPM');$row.Vtpm=if(([string]$value).Trim()-eq'1'){1}else{0};$value=Get-P2CsvValue $item @('StoragePolicyAssignment');if($value){$row.StoragePolicyAssignment=$value}
        $value=Get-P2CsvValue $item @('DiskFormat','DestinationDiskFormat');if($value){$row.DiskFormat=[string]$value}else{$row.DiskFormat='Same format as source'}
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
 $script:cmbP2Policy.Items.Clear();$script:cmbP2VtpmPolicy.Items.Clear();foreach($x in @(Get-SpbmStoragePolicy -Server $script:DestinationVIServer -ErrorAction SilentlyContinue|Sort-Object Name)){[void]$script:cmbP2Policy.Items.Add($x);[void]$script:cmbP2VtpmPolicy.Items.Add($x)}
 if($script:cmbP2Storage.Items.Count){$script:cmbP2Storage.SelectedIndex=0};if($script:cmbP2Compute.Items.Count){$script:cmbP2Compute.SelectedIndex=0};if($script:cmbP2Policy.Items.Count){$script:cmbP2Policy.SelectedIndex=0};if($script:cmbP2VtpmPolicy.Items.Count){$script:cmbP2VtpmPolicy.SelectedIndex=-1};Log "Phase II discovery loaded storage=$($(Get-SafeCount $storage)), compute=$($(Get-SafeCount $compute))." PASS
}
function P2-Apply{
 $st=$script:cmbP2Storage.SelectedItem;$co=$script:cmbP2Compute.SelectedItem;$po=$script:cmbP2Policy.SelectedItem;$vpo=$script:cmbP2VtpmPolicy.SelectedItem
 foreach($r in $script:P2Rows){
  if($st -and [string]$r.StorageSelectionSource -ne 'Phase II Per-VM Override'){$r.Storage=$st.Name;$r.StorageId=$st.Id;$r.StorageType=$st.Type;$r.StorageSelectionSource='Phase II Global'}
  if($co){$r.Compute=$co.Name;$r.ComputeId=$co.Id;$r.ComputeType=$co.Type}
  if([string]$r.StoragePolicyAssignment -ne 'Phase II Per-VM Override'){
   if([int]$r.Vtpm -eq 1){if($vpo){$r.StoragePolicy=$vpo.Name;$r.StoragePolicyAssignment='vTPM Global'}}
   elseif($po){$r.StoragePolicy=$po.Name;$r.StoragePolicyAssignment='Standard Global'}
  }
  if([string]$r.DiskFormatSelectionSource -ne 'Phase II Per-VM Override'){$r.DiskFormat=$script:cmbP2Disk.SelectedItem.Content;$r.DiskFormatSelectionSource='Phase II Global'}
 }
 $script:gridP2.Items.Refresh();P2-Invalidate
}
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
 if($script:DestinationVIServer){
  $policy=Get-SpbmStoragePolicy -Server $script:DestinationVIServer -Name $Name -ErrorAction SilentlyContinue|Select-Object -First 1
  foreach($propertyName in 'Id','Uid','UniqueId'){
   $property=$policy.PSObject.Properties[$propertyName]
   if($property -and $property.Value){$value=[string]$property.Value;if($value -match '([0-9a-fA-F]{8}-[0-9a-fA-F-]{27,})'){return $Matches[1]}else{return $value}}
  }
 }
 return ''
}
function Get-P2HcxDestinationSegments{Log 'HCX related-inventory lookup skipped; using live Phase I vCenter network IDs.' INFO;return @()}
function Get-P2NativeNetworkMapping($Mapping){
 $srcId=Convert-P2NetworkMoRef ([string]$Mapping.SourceNetworkId);$dstId=Convert-P2NetworkMoRef ([string]$Mapping.DestinationNetworkId)
 $srcType=Get-P2NetworkEntityType $srcId;$dstType=Get-P2NetworkEntityType $dstId
 if(-not$srcType){throw "Unsupported source network ID '$srcId'."};if(-not$dstType){throw "Unsupported destination network ID '$dstId'."}
 [ordered]@{srcNetworkName=[string]$Mapping.SourceNetworkName;srcNetworkType=$srcType;srcNetworkId=$srcId;destNetworkName=[string]$Mapping.DestinationNetworkName;destNetworkType=$dstType;destNetworkId=$dstId}
}
function Resolve-P2AuthoritativeNetworkIds{
 foreach($row in @($script:P2Rows|Where-Object Include)){foreach($mapping in @($row.Networks)){$mapping.SourceNetworkId=Convert-P2NetworkMoRef ([string]$mapping.SourceNetworkId);$mapping.DestinationNetworkId=Convert-P2NetworkMoRef ([string]$mapping.DestinationNetworkId)}}
}
function Test-P2ImportedNetworkMappings{
 Resolve-P2AuthoritativeNetworkIds
 foreach($row in @($script:P2Rows|Where-Object Include)){foreach($mapping in @($row.Networks)){
  if(-not(Get-P2NetworkEntityType ([string]$mapping.SourceNetworkId))){throw "$($row.VMName): invalid source network ID '$($mapping.SourceNetworkId)'."}
  if(-not(Get-P2NetworkEntityType ([string]$mapping.DestinationNetworkId))){throw "$($row.VMName): invalid destination network ID '$($mapping.DestinationNetworkId)'."}
 }}
}
# HCX 9.1 Automatic Discovery Functions - Rev 2.1
# Environment-neutral. No HCX, vCenter, endpoint, resource, site-pair, or Service Mesh IDs are embedded.
# Dependencies expected from the parent application:
#   Invoke-HcxRest -Method <GET|POST|PUT> -Path <relative path> [-Body <object>] [-AllowFailure]
#   Log <message> <INFO|PASS|WARN|ERROR>
#   $script:Hcx.BaseUri, $script:SourceVIServer, $script:DestinationVIServer

function New-HcxDiscoveryState {
    [ordered]@{
        IsValid                    = $false
        Status                     = 'NotStarted'
        CurrentStage               = ''
        FailureMessage             = ''
        DiscoveredAt               = $null
        HcxManager                 = ''
        HcspUUID                   = ''
        Direction                  = ''
        SourceVCenterName          = ''
        SourceVCenterInstanceId    = ''
        DestinationVCenterName     = ''
        DestinationVCenterInstanceId = ''
        Source                     = $null
        Destination                = $null
        MigrationTopology          = $null
        ResourceContainers         = @()
        ServiceMeshes              = @()
        MobilityServiceMeshes      = @()
        EligibleServiceMeshes      = @()
        SelectedServiceMesh        = $null
        SupportedMigrationTypes    = @()
        DiscoveryArtifacts         = @()
    }
}

function Reset-HcxAutomaticDiscovery {
    $script:HcxDiscovery = New-HcxDiscoveryState
    if ($script:lblHcxTopologyStatus) { $script:lblHcxTopologyStatus.Text = 'Not discovered'; $script:lblHcxTopologyStatus.Foreground = 'Gray' }
    if ($script:lblHcxDirection) { $script:lblHcxDirection.Text = 'Not discovered'; $script:lblHcxDirection.Foreground = 'Gray' }
    if ($script:lblHcxServiceMesh) { $script:lblHcxServiceMesh.Text = 'Not discovered'; $script:lblHcxServiceMesh.Foreground = 'Gray' }
    if ($script:cmbHcxServiceMesh) { $script:cmbHcxServiceMesh.ItemsSource = $null; $script:cmbHcxServiceMesh.Visibility = 'Collapsed' }
}

function Set-HcxDiscoveryStage([string]$Stage) {
    $script:HcxDiscovery.Status = 'Discovering'
    $script:HcxDiscovery.CurrentStage = $Stage
    if ($script:lblHcxTopologyStatus) { $script:lblHcxTopologyStatus.Text = "Discovering: $Stage"; $script:lblHcxTopologyStatus.Foreground = 'Gold' }
    Log "HCX automatic discovery stage: $Stage." INFO
}

function ConvertFrom-HcxJsonIfNeeded($InputObject) {
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [string]) {
        $text=$InputObject.Trim()
        if (($text.StartsWith('{') -and $text.EndsWith('}')) -or ($text.StartsWith('[') -and $text.EndsWith(']'))) {
            try {
                # HCX 9.1 responses can contain both uuid and UUID in the same object.
                # -AsHashtable preserves both keys; PSCustomObject conversion cannot.
                return ($text | ConvertFrom-Json -AsHashtable -Depth 100 -ErrorAction Stop)
            } catch {
                Write-HcxDebug "JSON normalization failed after -AsHashtable: $($_.Exception.Message)" 'JSON'
                return $InputObject
            }
        }
    }
    return $InputObject
}
function Get-HcxDictionaryValue($Dictionary,[string]$Name) {
    if($null -eq $Dictionary){return $null}
    # Prefer exact-case key, then case-insensitive fallback only when exact key is absent.
    if($Dictionary.Contains($Name)){return (ConvertFrom-HcxJsonIfNeeded $Dictionary[$Name])}
    $matches=@($Dictionary.Keys|Where-Object{[string]$_ -ieq $Name})
    if($matches.Count -gt 0){return (ConvertFrom-HcxJsonIfNeeded $Dictionary[$matches[0]])}
    return $null
}
function Get-HcxPropertyValue($Object,[string]$Name) {
    $Object=ConvertFrom-HcxJsonIfNeeded $Object
    if($null -eq $Object){return $null}
    if($Object -is [System.Collections.IDictionary]){return (Get-HcxDictionaryValue $Object $Name)}
    $property=$Object.PSObject.Properties[$Name]
    if($property){return (ConvertFrom-HcxJsonIfNeeded $property.Value)}
    return $null
}
function Get-HcxObjectPropertyNames($Object) {
    $Object=ConvertFrom-HcxJsonIfNeeded $Object
    if($null -eq $Object){return @()}
    if($Object -is [System.Collections.IDictionary]){return @($Object.Keys|ForEach-Object{[string]$_})}
    return @($Object.PSObject.Properties.Name)
}
function Get-HcxFirstValue($Object,[string[]]$Names) {
    $Object=ConvertFrom-HcxJsonIfNeeded $Object
    if($null -eq $Object){return $null}
    foreach($name in $Names){
        $value=Get-HcxPropertyValue $Object $name
        if($null -ne $value -and -not[string]::IsNullOrWhiteSpace([string]$value)){return $value}
    }
    return $null
}
function Get-HcxResponseItems($Response) {
    $Response=ConvertFrom-HcxJsonIfNeeded $Response
    if($null -eq $Response){return @()}
    $items=Get-HcxPropertyValue $Response 'items'
    if($null -ne $items){return @($items|ForEach-Object{ConvertFrom-HcxJsonIfNeeded $_})}
    $data=Get-HcxPropertyValue $Response 'data'
    if($null -ne $data){
        $dataItems=Get-HcxPropertyValue $data 'items'
        if($null -ne $dataItems){return @($dataItems|ForEach-Object{ConvertFrom-HcxJsonIfNeeded $_})}
    }
    if($Response -is [System.Collections.IEnumerable] -and $Response -isnot [string] -and $Response -isnot [System.Collections.IDictionary]){
        return @($Response|ForEach-Object{ConvertFrom-HcxJsonIfNeeded $_})
    }
    return @($Response)
}
function Normalize-HcxHostName([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    try { if ($Value -match '^https?://') { return ([uri]$Value).Host.TrimEnd('.').ToLowerInvariant() } } catch {}
    return $Value.Trim().TrimEnd('/').TrimEnd('.').ToLowerInvariant()
}

function Get-HcxShortName([string]$Value) {
    $normalized = Normalize-HcxHostName $Value
    if (-not $normalized) { return '' }
    return ($normalized -split '\.')[0]
}

function Get-VCenterIdentity($VIServer, [string]$EnteredName = '') {
    if (-not $VIServer) { throw 'A connected VIServer is required.' }
    $serviceInstance = Get-View -Server $VIServer -Id 'ServiceInstance-ServiceInstance' -ErrorAction Stop
    [pscustomobject]@{
        EnteredName  = if ($EnteredName) { $EnteredName } else { [string]$VIServer.Name }
        ServerName   = [string]$VIServer.Name
        InstanceUuid = [string]$serviceInstance.Content.About.InstanceUuid
        FullName     = [string]$serviceInstance.Content.About.FullName
        Version      = [string]$serviceInstance.Content.About.Version
        Build        = [string]$serviceInstance.Content.About.Build
    }
}

function Get-HcxMigrationTopology {
    $response = Invoke-HcxRest -Method POST -Path '/hybridity/api/site/migration-topology' -Body ([ordered]@{}) -MaxAttempts 4
    $items = @(Get-HcxResponseItems $response)
    Write-HcxDebug "Migration-topology response normalized. RuntimeType=$($response.GetType().FullName); ItemCount=$($items.Count); FirstItemType=$(if($items.Count){$items[0].GetType().FullName}else{'none'})." 'TOPOLOGY-PARSE'
    if ($items.Count -eq 0) { throw 'HCX migration-topology returned no items.' }
    return [pscustomobject]@{ Response = $response; Items = $items }
}

function Convert-HcxTopologyItemToResource($Item, [string]$Direction = '') {
    $Item=ConvertFrom-HcxJsonIfNeeded $Item
    if($null -eq $Item){throw 'HCX topology conversion received a null root item.'}
    $appliance=Get-HcxPropertyValue $Item 'applianceDetails'
    if($null -eq $appliance){
        $props=@(Get-HcxObjectPropertyNames $Item)-join', '
        throw "HCX topology root item does not expose applianceDetails. RuntimeType='$($Item.GetType().FullName)'; Properties='$props'; Preview='$((Protect-HcxDiagnosticText ([string]$Item)).Substring(0,[Math]::Min(400,([string]$Item).Length)))'."
    }
    [pscustomobject]@{
        EndpointId=[string](Get-HcxFirstValue $appliance @('uuid','UUID','endpointId'));EndpointName=[string](Get-HcxFirstValue $appliance @('name','endpointName'))
        EndpointType=[string](Get-HcxFirstValue $appliance @('type','endpointType'));EndpointUrl=[string](Get-HcxFirstValue $appliance @('url'))
        ResourceId=[string](Get-HcxFirstValue $Item @('resourceId','infraManagerId','vcuuid','vcGuid'));ResourceName=[string](Get-HcxFirstValue $Item @('resourceName','name','url'))
        ResourceType=[string](Get-HcxFirstValue $Item @('resourceType','cloudType'));ComputeResourceId=[string](Get-HcxFirstValue $Item @('resourceId','infraManagerId','vcuuid','vcGuid'))
        Direction=$Direction;Raw=$Item
    }
}
function Convert-HcxTopologyDestinationToResource($Destination) {
    $Destination=ConvertFrom-HcxJsonIfNeeded $Destination
    if($null -eq $Destination){throw 'HCX topology conversion received a null destination item.'}
    $infra=Get-HcxPropertyValue $Destination 'infraManagerDetails';$appliance=Get-HcxPropertyValue $Destination 'applianceDetails'
    if($null -eq $infra -or $null -eq $appliance){
        $props=@(Get-HcxObjectPropertyNames $Destination)-join', '
        throw "HCX topology destination is missing infraManagerDetails or applianceDetails. RuntimeType='$($Destination.GetType().FullName)'; Properties='$props'."
    }
    [pscustomobject]@{
        EndpointId=[string](Get-HcxFirstValue $appliance @('uuid','UUID','endpointId'));EndpointName=[string](Get-HcxFirstValue $appliance @('name','endpointName'))
        EndpointType=[string](Get-HcxFirstValue $appliance @('type','endpointType'));EndpointUrl=[string](Get-HcxFirstValue $appliance @('url'))
        ResourceId=[string](Get-HcxFirstValue $infra @('uuid','UUID','resourceId','infraManagerId'));ResourceName=[string](Get-HcxFirstValue $infra @('name','resourceName'))
        ResourceType=[string](Get-HcxFirstValue $infra @('type','resourceType'));ComputeResourceId=[string](Get-HcxFirstValue $infra @('uuid','UUID','resourceId','infraManagerId'))
        Direction=[string](Get-HcxFirstValue $Destination @('migrationDirection','direction'));Raw=$Destination
    }
}
function Get-HcxTopologyResourceCandidates($TopologyItems) {
    $candidates = [System.Collections.Generic.List[object]]::new()
    foreach ($item in @($TopologyItems)) {
        $candidates.Add((Convert-HcxTopologyItemToResource $item))
        foreach ($destination in @(Get-HcxPropertyValue $item 'destinations')) { $candidates.Add((Convert-HcxTopologyDestinationToResource $destination)) }
    }
    $unique = @($candidates | Group-Object { "$($_.EndpointId)|$($_.ResourceId)" } | ForEach-Object { $_.Group[0] })
    return $unique
}

function Get-HcxResourceMatchScore($Candidate, $Identity) {
    $score = 0
    if ($Identity.InstanceUuid -and $Candidate.ResourceId -and $Identity.InstanceUuid -ieq $Candidate.ResourceId) { $score += 100 }
    $candidateName = Normalize-HcxHostName $Candidate.ResourceName
    foreach ($name in @($Identity.EnteredName,$Identity.ServerName)) {
        $normalized = Normalize-HcxHostName $name
        if ($normalized -and $candidateName -eq $normalized) { $score += 50 }
        elseif ($normalized -and (Get-HcxShortName $candidateName) -eq (Get-HcxShortName $normalized)) { $score += 10 }
    }
    return $score
}

function Resolve-HcxResourceForVCenter($Candidates, $Identity, [string]$Role) {
    $ranked = @($Candidates | ForEach-Object { [pscustomobject]@{ Candidate=$_; Score=(Get-HcxResourceMatchScore $_ $Identity) } } | Where-Object Score -gt 0 | Sort-Object Score -Descending)
    if ($ranked.Count -eq 0) { throw "No HCX topology resource matched the $Role vCenter '$($Identity.EnteredName)' with instance UUID '$($Identity.InstanceUuid)'." }
    $topScore = $ranked[0].Score
    $top = @($ranked | Where-Object Score -eq $topScore)
    if ($top.Count -ne 1) { throw "HCX topology resource matching for the $Role vCenter is ambiguous. $($top.Count) candidates have score $topScore." }
    return $top[0].Candidate
}

function Resolve-HcxDirection($TopologyItems, $Source, $Destination) {
    foreach ($item in @($TopologyItems)) {
        $root = Convert-HcxTopologyItemToResource $item
        foreach ($dest in @(Get-HcxPropertyValue $item 'destinations')) {
            $target = Convert-HcxTopologyDestinationToResource $dest
            if ($root.ResourceId -ieq $Source.ResourceId -and $target.ResourceId -ieq $Destination.ResourceId) { return [string]$target.Direction }
        }
    }
    throw "HCX migration-topology did not contain a route from source '$($Source.ResourceName)' to destination '$($Destination.ResourceName)'."
}

function Find-HcxSitePairUuid($TopologyItems, $Source, $Destination) {
    # In both captured HCX 9.1 workflows, hcspUUID equals the connected HCX appliance UUID represented by applianceDetails.
    # Prefer the topology root matching the vCenter managed by the connected HCX Manager. Fall back only to a unique root appliance UUID.
    $baseHost = Normalize-HcxHostName $script:Hcx.BaseUri
    $byUrl = @($TopologyItems | Where-Object { $a=Get-HcxPropertyValue $_ 'applianceDetails'; (Normalize-HcxHostName (Get-HcxFirstValue $a @('url'))) -eq $baseHost })
    if ($byUrl.Count -eq 1) { return [string](Get-HcxFirstValue (Get-HcxPropertyValue $byUrl[0] 'applianceDetails') @('uuid','UUID')) }
    $sourceRoot = @($TopologyItems | Where-Object { [string]$_.resourceId -ieq $Source.ResourceId })
    if ($sourceRoot.Count -eq 1) { return [string](Get-HcxFirstValue (Get-HcxPropertyValue $sourceRoot[0] 'applianceDetails') @('uuid','UUID')) }
    $uuids = @($TopologyItems | ForEach-Object { [string](Get-HcxFirstValue (Get-HcxPropertyValue $_ 'applianceDetails') @('uuid','UUID')) } | Where-Object { $_ } | Select-Object -Unique)
    if ($uuids.Count -eq 1) { return $uuids[0] }
    throw 'The active HCX site-pair UUID could not be uniquely resolved from migration-topology.'
}

function Get-HcxResourceContainers([string]$HcspUUID) {
    $path = '/hybridity/api/service/inventory/resourcecontainer/list?hcspUUID=' + [uri]::EscapeDataString($HcspUUID)
    $body = [ordered]@{ filter = [ordered]@{ cloud = [ordered]@{ remote = $true; local = $true } } }
    $response = Invoke-HcxRest -Method POST -Path $path -Body $body -AllowFailure
    if ($null -eq $response) { Log 'Optional HCX resource-container discovery returned no response; migration-topology remains authoritative.' WARN; return @() }
    return @(Get-HcxResponseItems $response)
}

function Get-HcxServiceMeshInventory([string]$VCenterInstanceUuid) {
    $path = '/hybridity/api/interconnect/serviceMesh?vcGuid=' + [uri]::EscapeDataString($VCenterInstanceUuid)
    return @(Get-HcxResponseItems (Invoke-HcxRest -Method GET -Path $path))
}

function Get-HcxMobilityServiceMeshes($Source, $Destination) {
    $body = [ordered]@{
        source      = [ordered]@{ endpointId=[string]$Source.EndpointId; resourceId=[string]$Source.ResourceId }
        destination = [ordered]@{ endpointId=[string]$Destination.EndpointId; resourceId=[string]$Destination.ResourceId }
    }
    return @(Get-HcxResponseItems (Invoke-HcxRest -Method POST -Path '/hybridity/api/interconnect/mobility/servicemeshes' -Body $body))
}

function Get-HcxSupportedMigrationTypes {
    $body = [ordered]@{ filter = [ordered]@{ getAllTypes = $true } }
    return Invoke-HcxRest -Method POST -Path '/hybridity/api/mobility/migrations/supportedMigrationTypes?optionsRequired=true' -Body $body
}

function Convert-HcxServiceMesh($Mesh) {
    $id = [string](Get-HcxFirstValue $Mesh @('serviceMeshId','id'))
    $services = @($Mesh.services | ForEach-Object {
        if ($_ -is [string]) { $_ }
        else { [string](Get-HcxFirstValue $_ @('service','name','type','id')) }
    } | Where-Object { $_ })
    [pscustomobject]@{
        Id          = $id
        Name        = [string](Get-HcxFirstValue $Mesh @('name','displayName'))
        DisplayName = ('{0} [{1}]' -f [string](Get-HcxFirstValue $Mesh @('name','displayName')), $id)
        State       = [string](Get-HcxFirstValue $Mesh @('state','status'))
        Services    = $services
        Raw         = $Mesh
    }
}

function Resolve-HcxEligibleServiceMeshes($GeneralMeshes, $MobilityMeshes, [string]$RequiredService = 'HCX_ASSISTED_VMOTION') {
    $general = @($GeneralMeshes | ForEach-Object { Convert-HcxServiceMesh $_ })
    $mobilityIds = @($MobilityMeshes | ForEach-Object { [string](Get-HcxFirstValue $_ @('id','serviceMeshId')) } | Where-Object { $_ })
    $eligible = @($general | Where-Object {
        $stateOk = (-not $_.State) -or $_.State -match '^(?i:COMMITTED|UP|OK|HEALTHY)$'
        $mobilityOk = ($mobilityIds.Count -eq 0) -or ($mobilityIds -contains $_.Id)
        $serviceOk = ($_.Services.Count -eq 0) -or ($_.Services -contains $RequiredService)
        $stateOk -and $mobilityOk -and $serviceOk
    })
    return $eligible
}

function Select-HcxServiceMesh($EligibleMeshes) {
    $eligible = @($EligibleMeshes)
    if ($eligible.Count -eq 0) { throw 'HCX returned no eligible Service Mesh for the selected source, destination, and migration service.' }
    if ($eligible.Count -eq 1) {
        if ($script:cmbHcxServiceMesh) { $script:cmbHcxServiceMesh.Visibility = 'Collapsed' }
        return $eligible[0]
    }
    if (-not $script:cmbHcxServiceMesh) { throw "HCX returned $($eligible.Count) eligible Service Meshes, but the Service Mesh selector is not available." }
    $script:cmbHcxServiceMesh.ItemsSource = $eligible
    $script:cmbHcxServiceMesh.SelectedIndex = 0
    $script:cmbHcxServiceMesh.Visibility = 'Visible'
    return $script:cmbHcxServiceMesh.SelectedItem
}

function Initialize-HcxAutomaticDiscovery {
    Reset-HcxAutomaticDiscovery
    try {
        $script:HcxDiscovery.HcxManager = $script:Hcx.BaseUri
        Set-HcxDiscoveryStage 'vCenter identity'
        $sourceIdentity = Get-VCenterIdentity -VIServer $script:SourceVIServer -EnteredName $script:txtSourceVC.Text
        $destinationIdentity = Get-VCenterIdentity -VIServer $script:DestinationVIServer -EnteredName $script:txtDestinationVC.Text

        Set-HcxDiscoveryStage 'migration topology'
        $topology = Get-HcxMigrationTopology
        $candidates = @(Get-HcxTopologyResourceCandidates $topology.Items)

        Set-HcxDiscoveryStage 'source and destination correlation'
        $source = Resolve-HcxResourceForVCenter -Candidates $candidates -Identity $sourceIdentity -Role 'source'
        $destination = Resolve-HcxResourceForVCenter -Candidates $candidates -Identity $destinationIdentity -Role 'destination'
        $direction = Resolve-HcxDirection -TopologyItems $topology.Items -Source $source -Destination $destination
        $hcspUUID = Find-HcxSitePairUuid -TopologyItems $topology.Items -Source $source -Destination $destination

        Set-HcxDiscoveryStage 'optional resource containers'
        $containers = @(Get-HcxResourceContainers -HcspUUID $hcspUUID)

        Set-HcxDiscoveryStage 'Service Mesh inventory'
        $generalMeshes = @(Get-HcxServiceMeshInventory -VCenterInstanceUuid $sourceIdentity.InstanceUuid)
        $mobilityMeshes = @(Get-HcxMobilityServiceMeshes -Source $source -Destination $destination)
        $eligibleMeshes = @(Resolve-HcxEligibleServiceMeshes -GeneralMeshes $generalMeshes -MobilityMeshes $mobilityMeshes)
        $selectedMesh = Select-HcxServiceMesh -EligibleMeshes $eligibleMeshes

        Set-HcxDiscoveryStage 'supported migration types'
        $supported = Get-HcxSupportedMigrationTypes

        $script:HcxDiscovery.HcxManager = $script:Hcx.BaseUri
        $script:HcxDiscovery.HcspUUID = $hcspUUID
        $script:HcxDiscovery.Direction = $direction
        $script:HcxDiscovery.SourceVCenterName = $sourceIdentity.ServerName
        $script:HcxDiscovery.SourceVCenterInstanceId = $sourceIdentity.InstanceUuid
        $script:HcxDiscovery.DestinationVCenterName = $destinationIdentity.ServerName
        $script:HcxDiscovery.DestinationVCenterInstanceId = $destinationIdentity.InstanceUuid
        $script:HcxDiscovery.Source = $source
        $script:HcxDiscovery.Destination = $destination
        $script:HcxDiscovery.MigrationTopology = $topology.Response
        $script:HcxDiscovery.ResourceContainers = $containers
        $script:HcxDiscovery.ServiceMeshes = $generalMeshes
        $script:HcxDiscovery.MobilityServiceMeshes = $mobilityMeshes
        $script:HcxDiscovery.EligibleServiceMeshes = $eligibleMeshes
        $script:HcxDiscovery.SelectedServiceMesh = $selectedMesh
        $script:HcxDiscovery.SupportedMigrationTypes = $supported
        $script:HcxDiscovery.DiscoveredAt = Get-Date
        $script:HcxDiscovery.Status = 'Ready'
        $script:HcxDiscovery.IsValid = $true

        if ($script:lblHcxTopologyStatus) { $script:lblHcxTopologyStatus.Text = 'Ready'; $script:lblHcxTopologyStatus.Foreground = 'LightGreen' }
        if ($script:lblHcxDirection) { $script:lblHcxDirection.Text = $direction; $script:lblHcxDirection.Foreground = 'LightGreen' }
        if ($script:lblHcxServiceMesh) { $script:lblHcxServiceMesh.Text = $selectedMesh.DisplayName; $script:lblHcxServiceMesh.Foreground = 'LightGreen' }
        Log "HCX automatic discovery passed. Direction='$direction'; SitePair='$hcspUUID'; Source='$($source.ResourceName)'; Destination='$($destination.ResourceName)'; ServiceMesh='$($selectedMesh.Id)'." PASS
        Save-HcxDebugArtifact -Category 'DISCOVERY' -Operation 'SUCCESS' -Data $script:HcxDiscovery | Out-Null
        return $script:HcxDiscovery
    }
    catch {
        $script:HcxDiscovery.IsValid = $false
        $script:HcxDiscovery.Status = 'Failed'
        $script:HcxDiscovery.FailureMessage = $_.Exception.Message
        if ($script:lblHcxTopologyStatus) { $script:lblHcxTopologyStatus.Text = 'Failed'; $script:lblHcxTopologyStatus.Foreground = 'Tomato' }
        Log "HCX automatic discovery failed at '$($script:HcxDiscovery.CurrentStage)': $($_.Exception.Message)" ERROR
        throw
    }
}

function Assert-HcxAutomaticDiscoveryCurrent {
    if (-not $script:HcxDiscovery -or -not $script:HcxDiscovery.IsValid) { throw 'HCX automatic discovery is not valid for the active connection.' }
    if ($script:HcxDiscovery.HcxManager -ne $script:Hcx.BaseUri) { throw 'The HCX Manager changed after automatic discovery.' }
    $sourceIdentity = Get-VCenterIdentity -VIServer $script:SourceVIServer -EnteredName $script:txtSourceVC.Text
    $destinationIdentity = Get-VCenterIdentity -VIServer $script:DestinationVIServer -EnteredName $script:txtDestinationVC.Text
    if ($sourceIdentity.InstanceUuid -ine $script:HcxDiscovery.SourceVCenterInstanceId) { throw 'The source vCenter changed after automatic discovery.' }
    if ($destinationIdentity.InstanceUuid -ine $script:HcxDiscovery.DestinationVCenterInstanceId) { throw 'The destination vCenter changed after automatic discovery.' }
    if ([string]::IsNullOrWhiteSpace([string]$script:HcxDiscovery.HcspUUID)) { throw 'The active HCX site-pair UUID is missing.' }
    if (-not $script:HcxDiscovery.Source -or -not $script:HcxDiscovery.Destination) { throw 'The discovered HCX source or destination resource is missing.' }
    if (-not $script:HcxDiscovery.SelectedServiceMesh) { throw 'No eligible HCX Service Mesh is selected.' }
    return $true
}

function Get-HcxDiscoveredPayloadTopology {
    $null = Assert-HcxAutomaticDiscoveryCurrent
    [pscustomobject]@{
        HcspUUID = [string]$script:HcxDiscovery.HcspUUID
        Direction = [string]$script:HcxDiscovery.Direction
        Source = [ordered]@{
            endpointId=[string]$script:HcxDiscovery.Source.EndpointId; endpointName=[string]$script:HcxDiscovery.Source.EndpointName; endpointType=[string]$script:HcxDiscovery.Source.EndpointType
            resourceId=[string]$script:HcxDiscovery.Source.ResourceId; resourceName=[string]$script:HcxDiscovery.Source.ResourceName; resourceType=[string]$script:HcxDiscovery.Source.ResourceType
            computeResourceId=[string]$script:HcxDiscovery.Source.ComputeResourceId
        }
        Destination = [ordered]@{
            endpointId=[string]$script:HcxDiscovery.Destination.EndpointId; endpointName=[string]$script:HcxDiscovery.Destination.EndpointName; endpointType=[string]$script:HcxDiscovery.Destination.EndpointType
            resourceId=[string]$script:HcxDiscovery.Destination.ResourceId; resourceName=[string]$script:HcxDiscovery.Destination.ResourceName; resourceType=[string]$script:HcxDiscovery.Destination.ResourceType
            computeResourceId=[string]$script:HcxDiscovery.Destination.ComputeResourceId
        }
        ServiceMeshId = [string]$script:HcxDiscovery.SelectedServiceMesh.Id
    }
}

# Integration points:
# 1. Call Reset-HcxAutomaticDiscovery at the beginning of Connect, Disconnect, and when any endpoint textbox changes.
# 2. Call Initialize-HcxAutomaticDiscovery after HCX plus both vCenter connections succeed.
# 3. In P2-Build use:
#       $topology = Get-HcxDiscoveredPayloadTopology
#       $source = $topology.Source
#       $destination = $topology.Destination
#       $serviceMeshId = $topology.ServiceMeshId
# 4. Use $script:HcxDiscovery.HcspUUID for current-session related-inventory requests.
# 5. Call Assert-HcxAutomaticDiscoveryCurrent immediately before Validate, Preview, POST, and every PUT.

function Get-HcxNsxContext {
    $response = Invoke-HcxRest -Method GET -Path '/hybridity/api/metainfo/context/hcxconfig?sections=nsx' -AllowFailure
    if (-not $response) { return $null }
    $nodes = [System.Collections.Generic.List[object]]::new()
    function Add-HcxNode($Node) {
        if ($null -eq $Node -or $Node -is [string] -or $Node.GetType().IsPrimitive) { return }
        $nodes.Add($Node)
        if ($Node -is [System.Collections.IDictionary]) { foreach ($value in $Node.Values) { Add-HcxNode $value }; return }
        if ($Node -is [System.Collections.IEnumerable] -and $Node -isnot [pscustomobject]) { foreach ($value in $Node) { Add-HcxNode $value }; return }
        foreach ($property in $Node.PSObject.Properties) { if ($property.Value -isnot [string]) { Add-HcxNode $property.Value } }
    }
    Add-HcxNode $response
    $candidate = @($nodes | Where-Object {
        (Get-HcxFirstValue $_ @('uuid','UUID','infraManagerId','resourceId')) -and
        (([string](Get-HcxFirstValue $_ @('type','resourceType','name','managerType'))) -match '(?i)NSX')
    } | Select-Object -First 1)
    if ($candidate.Count -eq 0) {
        Write-HcxDebug "HCX NSX context response did not expose an NSX manager UUID. NSX related-inventory expansion is skipped; existing vCenter and previously discovered HCX network inventory remains available." 'NSX-CONTEXT'
        return $null
    }
    [pscustomobject]@{
        Id=[string](Get-HcxFirstValue $candidate[0] @('uuid','UUID','infraManagerId','resourceId'))
        Name=[string](Get-HcxFirstValue $candidate[0] @('name','resourceName','displayName'))
        Raw=$candidate[0]
    }
}
function Get-HcxRelatedInventoryPage($Destination,[string]$Category,[string]$EntityType,[string]$InventoryType,[string]$InfraManagerId,[int]$PageNumber=1,[int]$PageSize=500,[hashtable]$Filter=@{}) {
    $body=[ordered]@{
        endpointId=[string]$Destination.EndpointId;infraManagerId=$InfraManagerId;categoryType=$Category;relatedEntityType=$EntityType
        inventoryType=$InventoryType;entityId=$InfraManagerId;entityType=$InventoryType;filter=$Filter
        pageParameters=[ordered]@{pageNumber=$PageNumber;pageSize=$PageSize;sortBy=@([ordered]@{field='name';direction='ASC'})}
    }
    $path='/hybridity/api/v2/inventory/related?hcspUUID='+[uri]::EscapeDataString([string]$script:HcxDiscovery.HcspUUID)
    return Invoke-HcxRest -Method POST -Path $path -Body $body
}
function Get-HcxRelatedInventoryAll($Destination,[string]$Category,[string]$EntityType,[string]$InventoryType,[string]$InfraManagerId,[hashtable]$Filter=@{}) {
    $all=[System.Collections.Generic.List[object]]::new()
    for($page=1;$page -le 100;$page++){
        $response=Get-HcxRelatedInventoryPage -Destination $Destination -Category $Category -EntityType $EntityType -InventoryType $InventoryType -InfraManagerId $InfraManagerId -PageNumber $page -Filter $Filter
        $items=@(Get-HcxResponseItems $response);foreach($item in $items){$all.Add($item)}
        if($items.Count -lt 500){break}
    }
    return @($all)
}
function Merge-HcxAutomaticNetworkInventory {
    $null=Assert-HcxAutomaticDiscoveryCurrent
    $destination=$script:HcxDiscovery.Destination
    $vcFilter=@{entityTypeLegacy=$true;excludeLowerBackings=$true}
    $hcxNetworks=[System.Collections.Generic.List[object]]::new()
    foreach($type in @('DistributedVirtualPortgroup','StandardPortgroup','OpaqueNetwork')){
        try{foreach($item in @(Get-HcxRelatedInventoryAll -Destination $destination -Category 'NETWORK' -EntityType $type -InventoryType 'VC' -InfraManagerId $destination.ResourceId -Filter $vcFilter)){$hcxNetworks.Add($item)}}catch{Log "HCX VC network discovery for $type failed: $($_.Exception.Message)" WARN}
    }
    $nsx=Get-HcxNsxContext
    if($nsx -and $nsx.Id){
        foreach($type in @('NsxtSegment','VirtualWire')){
            try{foreach($item in @(Get-HcxRelatedInventoryAll -Destination $destination -Category 'NETWORK' -EntityType $type -InventoryType 'NSX' -InfraManagerId $nsx.Id -Filter $vcFilter)){$hcxNetworks.Add($item)}}catch{Log "HCX NSX network discovery for $type failed: $($_.Exception.Message)" WARN}
        }
    }
    $normalized=@($hcxNetworks|ForEach-Object{
        $name=[string](Get-HcxFirstValue $_ @('name','displayName','entityName'))
        $id=Convert-HcxNetworkIdCanonical ([string](Get-HcxFirstValue $_ @('entityId','id','path','networkId')))
        $type=[string](Get-HcxFirstValue $_ @('entityType','type','relatedEntityType'))
        if($name -and $id){[pscustomobject]@{Kind='Network';Name=$name;Id=$id;EntityType=$type;Raw=$_}}
    })
    $combined=@($script:Inventory.Networks)+$normalized
    foreach($network in $combined){
        $network.Id=Convert-HcxNetworkIdCanonical ([string]$network.Id)
        if(-not$network.PSObject.Properties['EntityType']){
            $inferredEntityType=''
            if($network.Id -match '^dvportgroup-'){$inferredEntityType='DistributedVirtualPortgroup'}
            elseif($network.Id -match '^network-'){$inferredEntityType='StandardPortgroup'}
            elseif($network.Id -match '^/infra/segments/'){$inferredEntityType='NsxtSegment'}
            elseif($network.Id -match '^opaqueNetwork-'){$inferredEntityType='OpaqueNetwork'}
            $network|Add-Member -NotePropertyName EntityType -NotePropertyValue $inferredEntityType
        }
    }
    Write-HcxDebug "Canonical network merge prepared. InputCount=$(@($combined).Count); TypedCount=$(@($combined|Where-Object{$_.EntityType}).Count)." 'NETWORK-MERGE'
    $script:Inventory.Networks=@(
        $combined |
            Group-Object Id |
            ForEach-Object {
                $dictionaryItems=@($_.Group|Where-Object{$_.Raw -is [System.Collections.IDictionary]})
                if($dictionaryItems.Count -gt 0){$dictionaryItems[0]}else{@($_.Group)[0]}
            } |
            Sort-Object Name,EntityType,Id
    )
    if(Get-Command Bind-DestinationControls -CommandType Function -ErrorAction SilentlyContinue){Bind-DestinationControls}else{throw 'Required function Bind-DestinationControls is not loaded.'}
    Log "Automatic HCX network inventory merged: HCX=$($normalized.Count); Combined=$(@($script:Inventory.Networks).Count)." PASS
}

# v2.2.13 NETWORK MAP VALIDATION
function P2-Build{
 Commit-P2Grid
 Resolve-P2AuthoritativeNetworkIds
 Test-P2ImportedNetworkMappings
 foreach($row in $script:P2Rows){$c=@($script:P2Computes|Where-Object Id -eq $row.ComputeId|Select-Object -First 1);if($c){$row.Compute=$c.Name;$row.ComputeType=$c.Type};$d=@($script:P2Storages|Where-Object Id -eq $row.StorageId|Select-Object -First 1);if($d){$row.Storage=$d.Name;$row.StorageType=$d.Type}}
 if(-not$script:SourceVIServer){throw 'Source vCenter connection is required to build native VM entity and NIC details.'}
 $included=@($script:P2Rows|Where-Object Include)
 if(-not$(Get-SafeCount $included)){throw 'No Phase II VMs are included.'}

 $null=Assert-HcxAutomaticDiscoveryCurrent
 $topology=Get-HcxDiscoveredPayloadTopology
 $source=$topology.Source
 $destination=$topology.Destination
 $serviceMeshId=[string]$topology.ServiceMeshId
 $first=@($included|Where-Object{[int]$_.Vtpm -ne 1 -and [string]$_.StoragePolicyAssignment -ne 'Per-VM Override'}|Select-Object -First 1);if(-not$first){$first=$included[0]}
 $policyId=Get-P2StoragePolicyId ([string]$first.StoragePolicy)
 if(-not$policyId){throw "Unable to resolve storage policy ID for '$($first.StoragePolicy)'."}
 $policyIds=@{};foreach($policyName in @($included|ForEach-Object{[string]$_.StoragePolicy}|Where-Object{$_}|Select-Object -Unique)){$resolvedPolicyId=Get-P2StoragePolicyId $policyName;if(-not$resolvedPolicyId){throw "Unable to resolve storage policy ID for '$policyName'."};$policyIds[$policyName]=$resolvedPolicyId}
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
  source=$source;destination=$destination;migrationType='xVMotion';servicemeshId=$serviceMeshId
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
   if($(Get-SafeCount $importedMatches)-ne1){throw "$($row.VMName): NIC '$($adapter.Name)' on '$sourceNetworkName' matched $($(Get-SafeCount $importedMatches)) imported network mapping(s); exactly one is required."}
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
   Log "Phase II VM placement: VM=$($row.VMName); Compute=$($row.Compute); ComputeId=$(Convert-P2MoRef ([string]$row.ComputeId)); ServiceMesh=$serviceMeshId." PASS
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
   storage=if([int]$row.Vtpm -eq 1 -or [string]$row.StoragePolicyAssignment -eq 'Per-VM Override'){[ordered]@{defaultStorage=[ordered]@{id=(Convert-P2MoRef ([string]$row.StorageId));name=[string]$row.Storage;type=if($row.StorageType-match'storagepod|Cluster'){'storagepod'}else{'datastore'};storageParams=@([ordered]@{option='StorageProfile';value=[string]$policyIds[[string]$row.StoragePolicy];type='defined';name=[string]$row.StoragePolicy});diskProvisionType=(Convert-P2DiskFormat ([string]$row.DiskFormat))}}}else{$null}
   servicemeshId=$serviceMeshId
   operationType='ADD'
  }
  if([int]$row.Vtpm -ne 1 -and [string]$row.StoragePolicyAssignment -ne 'Per-VM Override'){[void]$migrationIntents[-1].Remove('storage')}
 }

 $expectedKeys=[System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
 foreach($row in $included){foreach($mapping in @($row.Networks)){[void]$expectedKeys.Add(('{0}|{1}' -f ([string]$mapping.SourceNetworkId).ToLowerInvariant(),([string]$mapping.DestinationNetworkId).ToLowerInvariant()))}}
 $expectedMappings=$expectedKeys.Count
 $actualGroupMappings=@($groupDefaults.networkParams.defaultMappings).Count
 $actualVmMappings=@($migrationIntents|ForEach-Object{@($_.networkParams.networkMappings)}).Count
 if($actualGroupMappings-ne$expectedMappings){throw "Payload network consistency failed: expected $expectedMappings group mapping(s), built $actualGroupMappings."}
 if($actualVmMappings-lt$(Get-SafeCount $included)){throw "Payload network consistency failed: $($(Get-SafeCount $included)) VM(s) included but only $actualVmMappings per-VM NIC mapping(s) were built."}
 Log "Phase II payload network consistency passed: GroupMappings=$actualGroupMappings; VmNicMappings=$actualVmMappings; IncludedVMs=$($(Get-SafeCount $included))." PASS
 Log "Phase II placement contract: Compute=$($first.Compute); ComputeId=$(Convert-P2MoRef ([string]$first.ComputeId)); PlacementType=$placementType; ServiceMesh=$($groupDefaults.servicemeshId); Storage=$($first.Storage); Networks=$(@($groupNetworks).Count)." PASS
 [ordered]@{items=@([ordered]@{name=$script:txtP2Name.Text.Trim();groupDefaults=$groupDefaults;migrations=@($migrationIntents)})}
}
function P2-BuildAll{
 $base=$script:txtP2Name.Text.Trim();$count=[int]$script:cmbP2GroupCount.SelectedItem;if(-not$base){throw 'Base Group Name is required.'};if($count-lt1-or$count-gt 50){throw 'Number of groups must be 1 through 50.'}
 $originalName=$script:txtP2Name.Text;$state=@{};$list=[Collections.Generic.List[object]]::new();foreach($r in $script:P2Rows){$state[$r.VMName]=[bool]$r.Include}
 try{foreach($n in 1..$count){$members=@($script:P2Rows|Where-Object{$state[$_.VMName]-and[int]$_.MobilityGroupNumber-eq$n});if(-not$(Get-SafeCount $members)){throw "Mobility Group $n has no included VMs."};foreach($r in $script:P2Rows){$r.Include=($state[$r.VMName]-and[int]$r.MobilityGroupNumber-eq$n)};$script:txtP2Name.Text='{0}-{1:d2}'-f$base,$n;$payload=P2-Build;$payload=script:Normalize-P2HostAndStoragePodPayload -Payload $payload;$list.Add([pscustomobject]@{Number=$n;Name=$script:txtP2Name.Text;VMCount=$(Get-SafeCount $members);Payload=$payload})}}
 finally{foreach($r in $script:P2Rows){$r.Include=$state[$r.VMName]};$script:txtP2Name.Text=$originalName;$script:gridP2.Items.Refresh()}
 return @($list)
}
function P2-Validate {
    Commit-P2Grid
    $errors=[Collections.Generic.List[string]]::new()
    if(-not$script:Hcx.Connected){$errors.Add('Source HCX Manager is not connected.')}else{try{$null=Assert-HcxAutomaticDiscoveryCurrent}catch{$errors.Add($_.Exception.Message)}}
    if(-not$(@($script:P2Rows).Count)){$errors.Add('Import a prepared mobility-group CSV.')}
    if([string]::IsNullOrWhiteSpace($script:txtP2Name.Text)){$errors.Add('Base Group Name is required.')};$gc=[int]$script:cmbP2GroupCount.SelectedItem;foreach($n in 1..$gc){if(-not@($script:P2Rows|Where-Object{$_.Include-and[int]$_.MobilityGroupNumber-eq$n}).Count){$errors.Add("Mobility Group $n has no included VMs.")}}
    if([string]::IsNullOrWhiteSpace($script:txtP2SourceSite.Text)){$errors.Add('Source Site is required.')}
    if([string]::IsNullOrWhiteSpace($script:txtP2DestSite.Text)){$errors.Add('Destination Site is required.')}
    if(-not@($script:P2Rows|Where-Object Include).Count){$errors.Add('At least one VM must be included.')}
    foreach($row in @($script:P2Rows|Where-Object Include)){if([string]::IsNullOrWhiteSpace([string]$row.StorageId)){$errors.Add("$($row.VMName): destination storage is not selected.")};if([string]::IsNullOrWhiteSpace([string]$row.StoragePolicy)){$errors.Add("$($row.VMName): effective storage policy is not selected.")}elseif(-not @(Get-SpbmStoragePolicy -Server $script:DestinationVIServer -Name $row.StoragePolicy -ErrorAction SilentlyContinue).Count){$errors.Add("$($row.VMName): storage policy '$($row.StoragePolicy)' is not present in destination inventory.")};if([int]$row.Vtpm -eq 1 -and [string]::IsNullOrWhiteSpace([string]$row.StoragePolicy)){$errors.Add("$($row.VMName): detected vTPM requires an effective storage policy.")}}
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

function P2-Preview{try{$null=Assert-HcxAutomaticDiscoveryCurrent;$script:Window.Cursor='Wait';$script:P2Payload=@(P2-BuildAll);$folder=script:Get-P2ArtifactFolder;foreach($x in $script:P2Payload){$path=Join-Path $folder("HCX91-$($x.Name)-Payload.json");$x.Payload|ConvertTo-Json -Depth 50|Set-Content $path -Encoding utf8BOM};Invoke-Item $folder}finally{$script:Window.Cursor=$null}}
function Assert-HcxSubmissionAuthentication {
    if(-not$script:Hcx.Connected){throw 'HCX is not connected.'}
    if(-not$script:Hcx.Headers -or -not$script:Hcx.Headers.ContainsKey('x-hm-authorization')){throw 'The active HCX session does not contain x-hm-authorization.'}
    if([string]::IsNullOrWhiteSpace([string]$script:Hcx.Headers['x-hm-authorization'])){throw 'The active HCX authorization value is empty.'}
    return $true
}
function Refresh-HcxAuthentication {
    $manager=[string]$script:txtHcx.Text.Trim();$user=[string]$script:txtUser.Text.Trim();$password=[string]$script:txtPassword.Password
    if([string]::IsNullOrWhiteSpace($manager)-or[string]::IsNullOrWhiteSpace($user)-or[string]::IsNullOrWhiteSpace($password)){throw 'HCX Manager, username, and password are required to refresh authentication.'}
    Write-HcxDebug "Refreshing HCX authentication for '$manager' before mobility-group submission." 'AUTH-REFRESH'
    Connect-Hcx91Rest -Fqdn $manager -User $user -Password $password
    $null=Assert-HcxSubmissionAuthentication
    Log 'HCX authentication refreshed successfully for mobility-group submission.' PASS
}
function Test-HcxUnauthorizedError($ErrorRecord){
    $messages=[System.Collections.Generic.List[string]]::new();$current=$ErrorRecord.Exception
    while($current){$messages.Add([string]$current.Message);$current=$current.InnerException}
    if($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message){$messages.Add([string]$ErrorRecord.ErrorDetails.Message)}
    return (($messages -join ' | ') -match '(?i)HTTP=401|status code.*401|401.*Unauthorized|Full authentication is required')
}
function Submit-HcxMobilityGroup {
    param([Parameter(Mandatory=$true)][string]$GroupName,[Parameter(Mandatory=$true)]$Payload)
    $null=Assert-HcxAutomaticDiscoveryCurrent;$null=Assert-HcxSubmissionAuthentication
    $path='/hybridity/api/mobility/groups'
    $requestArtifact=Save-HcxDebugArtifact -Category 'MOBILITY-GROUP-REQUEST' -Operation $GroupName -Data $Payload
    try{
        $response=Invoke-HcxRest -Method POST -Path $path -Body $Payload -MaxAttempts 3
        Save-HcxDebugArtifact -Category 'MOBILITY-GROUP-RESPONSE' -Operation $GroupName -Data $response -Metadata @{AuthenticationRefreshed=$false;RequestArtifact=$requestArtifact}|Out-Null
        return $response
    }catch{
        if(-not(Test-HcxUnauthorizedError $_)){throw}
        Log "HCX returned HTTP 401 while creating '$GroupName'. Authentication will be refreshed and the request retried once." WARN
        Save-HcxDebugArtifact -Category 'MOBILITY-GROUP-AUTH-FAILURE' -Operation $GroupName -Data ([ordered]@{GroupName=$GroupName;Message=$_.Exception.Message;RequestArtifact=$requestArtifact;ReauthenticationAttempted=$true})|Out-Null
        Refresh-HcxAuthentication
        $null=Assert-HcxAutomaticDiscoveryCurrent;$null=Assert-HcxSubmissionAuthentication
        $retryArtifact=Save-HcxDebugArtifact -Category 'MOBILITY-GROUP-RETRY-REQUEST' -Operation $GroupName -Data $Payload
        try{
            $response=Invoke-HcxRest -Method POST -Path $path -Body $Payload -MaxAttempts 1
            Save-HcxDebugArtifact -Category 'MOBILITY-GROUP-RETRY-RESPONSE' -Operation $GroupName -Data $response -Metadata @{AuthenticationRefreshed=$true;RetryRequestArtifact=$retryArtifact}|Out-Null
            Log "Mobility group '$GroupName' was accepted after HCX authentication refresh." PASS
            return $response
        }catch{
            $diagnostic=Write-HcxExceptionDiagnostic $_ 'MOBILITY-GROUP-RETRY'
            throw "HCX rejected mobility group '$GroupName' after one authentication refresh and retry. Diagnostic='$diagnostic'. Error='$($_.Exception.Message)'"
        }
    }
}
function P2-Create{
 if(-not$script:P2Valid){throw 'Run validation first.'}
 $null=Assert-HcxAutomaticDiscoveryCurrent
 if(-not$script:P2Payload){$script:P2Payload=@(P2-BuildAll)}
 $summary=@($script:P2Payload|ForEach-Object{"$($_.Name): $($_.VMCount) VM(s)"})-join[Environment]::NewLine
 if([Windows.MessageBox]::Show("Create these HCX drafts one at a time?`n`n$summary",'Create Mobility Groups','YesNo')-ne'Yes'){return}
 $folder=script:Get-P2ArtifactFolder;$ok=0
 Refresh-HcxAuthentication
 foreach($x in $script:P2Payload){
  $path=Join-Path $folder("HCX91-$($x.Name)-Request.json")
  $x.Payload|ConvertTo-Json -Depth 50|Set-Content $path -Encoding utf8BOM
  try{
   $response=Submit-HcxMobilityGroup -GroupName ([string]$x.Name) -Payload $x.Payload
   $responsePath=Join-Path $folder("HCX91-$($x.Name)-Response.json")
   $response|ConvertTo-Json -Depth 50|Set-Content $responsePath -Encoding utf8BOM
   $ok++;Log "Created draft '$($x.Name)' with $($x.VMCount) VM(s)." PASS
  }catch{throw "Creation stopped at '$($x.Name)' after $ok successful draft(s): $($_.Exception.Message)"}
 }
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

    if (-not $(Get-SafeCount $groups)) {
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
    if(-not$(Get-SafeCount $groups)){throw'Payload does not contain items.'}
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

    $folder=[string]$script:RunDir
    if([string]::IsNullOrWhiteSpace($folder) -or -not(Test-Path -LiteralPath $folder -PathType Container)){
        if(Get-Command script:Get-P2ArtifactFolder -ErrorAction SilentlyContinue){$folder=[string](script:Get-P2ArtifactFolder)}
    }
    if([string]::IsNullOrWhiteSpace($folder)){throw 'The active HCX run folder could not be resolved for the Host/StoragePod audit.'}
    if(-not(Test-Path -LiteralPath $folder -PathType Container)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}
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
<Grid.RowDefinitions><RowDefinition Height="32"/><RowDefinition Height="32"/><RowDefinition Height="32"/><RowDefinition Height="46"/><RowDefinition Height="32"/></Grid.RowDefinitions>

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
<WrapPanel Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="6" HorizontalAlignment="Center" VerticalAlignment="Center">
<TextBlock Text="HCX Topology:" FontWeight="SemiBold"/><TextBlock x:Name="lblHcxTopologyStatus" Text="Not discovered" Foreground="Gray" Margin="4,3,16,3"/>
<TextBlock Text="Direction:" FontWeight="SemiBold"/><TextBlock x:Name="lblHcxDirection" Text="Not discovered" Foreground="Gray" Margin="4,3,16,3"/>
<TextBlock Text="Service Mesh:" FontWeight="SemiBold"/><TextBlock x:Name="lblHcxServiceMesh" Text="Not discovered" Foreground="Gray" Margin="4,3,8,3"/>
<ComboBox x:Name="cmbHcxServiceMesh" Width="290" DisplayMemberPath="DisplayName" Visibility="Collapsed"/>
</WrapPanel>
</Grid>
</GroupBox>
</Grid>
<TabControl Grid.Row="1" Margin="0,6,0,0"><TabItem Header="Prepare Mobility Group"><Grid Margin="7"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<GroupBox Header="Global Destination Settings" Margin="5,10,5,5" Padding="9">
 <Grid Margin="8">
  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="190"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="220"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="240"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="230"/></Grid.ColumnDefinitions>
  <Grid.RowDefinitions><RowDefinition Height="34"/><RowDefinition Height="34"/><RowDefinition Height="34"/></Grid.RowDefinitions>
  <TextBlock Grid.Row="0" Grid.Column="0" Text="Destination Site" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbDestinationSite" Grid.Row="0" Grid.Column="1" Margin="4" DisplayMemberPath="Name" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="0" Grid.Column="2" Text="Compute" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalCompute" Grid.Row="0" Grid.Column="3" Margin="4" DisplayMemberPath="DisplayName" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="0" Grid.Column="4" Text="Datastore" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalDatastore" Grid.Row="0" Grid.Column="5" Margin="4" DisplayMemberPath="DisplayName" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="0" Grid.Column="6" Text="Storage Policy" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalStoragePolicy" Grid.Row="0" Grid.Column="7" Margin="4" DisplayMemberPath="Name" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="1" Grid.Column="0" Text="Folder" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalFolder" Grid.Row="1" Grid.Column="1" Margin="4" DisplayMemberPath="Name" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  <TextBlock Grid.Row="1" Grid.Column="2" Text="Migration Type" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbMigrationType" Grid.Row="1" Grid.Column="3" Margin="4" SelectedIndex="0"/>
  <TextBlock Grid.Row="1" Grid.Column="4" Text="vTPM Storage Policy" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbVtpmStoragePolicy" Grid.Row="1" Grid.Column="5" Margin="4" DisplayMemberPath="Name"/>
  <TextBlock Grid.Row="1" Grid.Column="6" Text="Disk Format" Margin="4" VerticalAlignment="Center"/><ComboBox x:Name="cmbGlobalDiskFormat" Grid.Row="1" Grid.Column="7" Margin="4" SelectedIndex="0"><ComboBoxItem>Same format as source</ComboBoxItem><ComboBoxItem>Thin Provision</ComboBoxItem><ComboBoxItem>Thick Provision Lazy Zeroed</ComboBoxItem><ComboBoxItem>Thick Provision Eager Zeroed</ComboBoxItem></ComboBox><TextBlock Grid.Row="2" Grid.Column="0" Text="Network Mapping CSV" Margin="4" VerticalAlignment="Center"/><TextBox x:Name="txtNetworkMappingCsv" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="3" Margin="4" IsReadOnly="True" ToolTip="Optional CSV columns: SourceNetworkName,DestinationNetworkName"/><Button x:Name="btnImportNetworkMapping" Grid.Row="2" Grid.Column="4" Margin="4" Content="Import Mapping CSV"/><Button x:Name="btnClearNetworkMapping" Grid.Row="2" Grid.Column="5" Margin="4" Content="Clear Mapping CSV"/><StackPanel Grid.Row="2" Grid.Column="6" Orientation="Horizontal" VerticalAlignment="Center"><TextBlock Text="Status:"/><TextBlock x:Name="lblNetworkMappingStatus" Text="Not loaded" Foreground="#76C7D8"/></StackPanel><Button x:Name="btnApplyGlobal" Grid.Row="2" Grid.Column="7" Margin="4" Content="Apply to All VMs"/>
 </Grid>
</GroupBox>
<DataGrid x:Name="gridVMs" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" SelectionMode="Extended" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" EnableRowVirtualization="True" EnableColumnVirtualization="True" ScrollViewer.CanContentScroll="True"><DataGrid.Columns>
<DataGridCheckBoxColumn Header="Include" Binding="{Binding Include, Mode=TwoWay}" Width="60"/><DataGridTextColumn Header="Mobility Group #" Binding="{Binding MobilityGroupNumber, Mode=TwoWay}" Width="105"/><DataGridTextColumn Header="Status" Binding="{Binding Status}" IsReadOnly="True" Width="85"/><DataGridTextColumn Header="vTPM (Auto)" Binding="{Binding Vtpm}" IsReadOnly="True" Width="80"/><DataGridTextColumn Header="vTPM Detection" Binding="{Binding VtpmDetectionStatus}" IsReadOnly="True" Width="105"/><DataGridTextColumn Header="VM Name" Binding="{Binding VMName}" IsReadOnly="True" Width="160"/><DataGridTextColumn Header="Power" Binding="{Binding PowerState}" IsReadOnly="True" Width="80"/><DataGridTextColumn Header="Source Compute" Binding="{Binding SourceCompute}" IsReadOnly="True" Width="150"/><DataGridTextColumn Header="Source Datastore" Binding="{Binding SourceDatastore}" IsReadOnly="True" Width="150"/><DataGridTextColumn Header="Source Networks" Binding="{Binding SourceNetworks}" IsReadOnly="True" Width="190"/><DataGridTextColumn Header="NICs" Binding="{Binding NicCount}" IsReadOnly="True" Width="45"/>
<DataGridTemplateColumn Header="Network Mapping" Width="330"><DataGridTemplateColumn.CellTemplate><DataTemplate><StackPanel Orientation="Horizontal"><TextBlock Text="{Binding NicSummary}" Width="255" TextTrimming="CharacterEllipsis" ToolTip="{Binding NicSummary}"/><Button Content="Configure" Tag="{Binding}" Padding="5,2" MinHeight="23"/></StackPanel></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn>
<DataGridTemplateColumn Header="Migration Type" Width="105"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.MigrationTypes, RelativeSource={RelativeSource AncestorType=Window}}" SelectedItem="{Binding MigrationType, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Destination Site" Binding="{Binding DestinationSite}" IsReadOnly="True" Width="145"/><DataGridTemplateColumn Header="Destination Compute" Width="170"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Computes, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding DestinationComputeId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTemplateColumn Header="Destination Folder" Width="155"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Folders, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="Name" SelectedValuePath="Id" SelectedValue="{Binding DestinationFolderId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Policy Assignment" Binding="{Binding StoragePolicyAssignment}" IsReadOnly="True" Width="120"/><DataGridTextColumn Header="Disk Format" Binding="{Binding DiskFormat}" IsReadOnly="True" Width="150"/><DataGridTemplateColumn Header="Effective Storage Policy" Width="210"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox Tag="{Binding}" ItemsSource="{Binding DataContext.Policies, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="Name" SelectedValuePath="Name" SelectedValue="{Binding StoragePolicy, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTemplateColumn Header="Destination Datastore" Width="180"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Datastores, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding DestinationDatastoreId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" ScrollViewer.CanContentScroll="True"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Validation Detail" Binding="{Binding ValidationMessage}" IsReadOnly="True" Width="300"/>
</DataGrid.Columns></DataGrid>
<Grid Grid.Row="2"><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><StackPanel Orientation="Horizontal"><Button x:Name="btnImport" Content="Import VM CSV"/><Button x:Name="btnExample" Content="Download Example VM CSV"/><Button x:Name="btnExampleNetworkMapping" Content="Download Example Network Mapping CSV"/><Button x:Name="btnRemove" Content="Remove Selected VM(s)"/><Button x:Name="btnClear" Content="Clear Grid"/><Button x:Name="btnRefresh" Content="Refresh Discovery"/></StackPanel><StackPanel Grid.Column="1" Orientation="Horizontal"><TextBlock Text="Validation:" FontWeight="SemiBold" VerticalAlignment="Center"/><TextBlock x:Name="lblValidation" Text="Not validated" Foreground="#76C7D8" VerticalAlignment="Center"/><Button x:Name="btnValidate" Content="Validate" Width="100"/><Button x:Name="btnCreateCsv" Content="Create Mobility Group CSV" Width="145" IsEnabled="False"/></StackPanel></Grid>
</Grid></TabItem><TabItem Header="Create Mobility Group"><Grid Margin="8"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
<GroupBox Header="Workload Selection"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="100"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="110"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions><Grid.RowDefinitions><RowDefinition Height="34"/><RowDefinition Height="34"/></Grid.RowDefinitions><TextBlock Text="Phase I CSV"/><TextBox x:Name="txtP2Csv" Grid.Column="1" IsReadOnly="True"/><Button x:Name="btnP2Import" Grid.Column="2" Content="Import CSV"/><TextBlock Grid.Column="3" Text="Base Group Name"/><StackPanel Grid.Column="4" Orientation="Horizontal"><TextBox x:Name="txtP2Name" Width="280"/><TextBlock Text="Groups (1-50)" Margin="10,3"/><ComboBox x:Name="cmbP2GroupCount" Width="65" MaxDropDownHeight="420" ScrollViewer.VerticalScrollBarVisibility="Auto" Foreground="#000000" Background="#FFFFFF" FontWeight="Bold" FontSize="14">
 <ComboBox.ItemTemplate><DataTemplate><Border Background="#FFFFFF" Padding="6,3"><TextBlock Text="{Binding}" Foreground="#000000" Background="#FFFFFF" FontWeight="Bold" FontSize="14"/></Border></DataTemplate></ComboBox.ItemTemplate>
 <ComboBox.ItemContainerStyle><Style TargetType="{x:Type ComboBoxItem}"><Setter Property="Foreground" Value="#000000"/><Setter Property="Background" Value="#FFFFFF"/><Style.Triggers><Trigger Property="IsHighlighted" Value="True"><Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="#000000"/></Trigger><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#CDE8FF"/><Setter Property="Foreground" Value="#000000"/></Trigger></Style.Triggers></Style></ComboBox.ItemContainerStyle>
</ComboBox></StackPanel><TextBlock Grid.Row="1" Text="Source HCX"/><TextBox x:Name="txtP2Hcx" Grid.Row="1" Grid.Column="1" IsReadOnly="True"/><TextBlock Grid.Row="1" Grid.Column="3" Text="Source Site"/><TextBox x:Name="txtP2SourceSite" Grid.Row="1" Grid.Column="4"/></Grid></GroupBox>
<Expander Grid.Row="1" Header="Destination Settings (Optional: expand only to revise Phase I selections)" IsExpanded="False" BorderBrush="#76C7D8" BorderThickness="1" Margin="0,3,0,3"><StackPanel><WrapPanel><TextBlock Text="Destination Site" Width="105"/><TextBox x:Name="txtP2DestSite" Width="210" IsReadOnly="True"/><TextBlock Text="Compute" Width="65"/><ComboBox x:Name="cmbP2Compute" Width="205" DisplayMemberPath="DisplayName" MaxDropDownHeight="420"/><TextBlock Text="Storage" Width="60"/><ComboBox x:Name="cmbP2Storage" Width="210" DisplayMemberPath="DisplayName" MaxDropDownHeight="420"/><TextBlock Text="Standard Policy" Width="100"/><ComboBox x:Name="cmbP2Policy" Width="200" DisplayMemberPath="Name" MaxDropDownHeight="420"/><TextBlock Text="vTPM Policy" Width="85"/><ComboBox x:Name="cmbP2VtpmPolicy" Width="200" DisplayMemberPath="Name" MaxDropDownHeight="420" ToolTip="Applied to automatically detected vTPM VMs unless a Phase II per-VM override exists."/><TextBlock Text="Disk Format" Width="80"/><ComboBox x:Name="cmbP2Disk" Width="185" SelectedIndex="0"><ComboBoxItem>Same format as source</ComboBoxItem><ComboBoxItem>Thin Provision</ComboBoxItem><ComboBoxItem>Thick Provision Lazy Zeroed</ComboBoxItem><ComboBoxItem>Thick Provision Eager Zeroed</ComboBoxItem></ComboBox><Button x:Name="btnP2Inventory" Content="Load Storage"/><Button x:Name="btnP2Apply" Content="Apply Destination Defaults" ToolTip="Applies compute, storage, disk format, standard policy, and vTPM policy while preserving Phase II per-VM overrides."/></WrapPanel><TextBlock Text="Optional: Phase II imports the validated destination selections from Phase I. Use this section only when destination compute, storage, storage policy, vTPM policy, or disk format must be updated before payload validation and HCX draft creation; otherwise no changes are required." Foreground="#76C7D8" TextWrapping="Wrap" Margin="6,8,6,2"/></StackPanel></Expander>
<GroupBox Grid.Row="2" Header="Migration Settings"><StackPanel><WrapPanel><TextBlock Text="Migration Type: HCX Assisted vMotion" Width="275"/><RadioButton x:Name="rbP2Now" Content="Start after transfer is complete" IsChecked="True" Margin="8"/><RadioButton x:Name="rbP2Schedule" Content="Set switchover schedule" Margin="8"/><RadioButton x:Name="rbP2Defer" Content="Defer switchover" Margin="8"/><TextBox x:Name="txtP2Schedule" Width="180" ToolTip="Schedule value"/></WrapPanel><WrapPanel><CheckBox x:Name="chkP2Mac" Content="Retain MAC" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2HW" Content="Upgrade Virtual Hardware" Margin="8"/><CheckBox x:Name="chkP2Tools" Content="Upgrade VM Tools" Margin="8"/><CheckBox x:Name="chkP2Attrs" Content="Migrate Custom Attributes" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2Tags" Content="Migrate vCenter Tags" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2Iso" Content="Force unmount ISO images" IsChecked="True" Margin="8"/><CheckBox x:Name="chkP2Security" Content="Replicate Security Tags" IsChecked="False" Margin="8"/></WrapPanel><TextBlock Text="For a complete migration, select a storage policy. Use an encryption-capable policy for vTPM or encrypted VMs. Drafts may be completed later in HCX Manager." Foreground="#76C7D8"/></StackPanel></GroupBox>
<TextBlock Grid.Row="3" VerticalAlignment="Top" HorizontalAlignment="Right" Margin="0,2,18,0" Panel.ZIndex="5" Foreground="#76C7D8" Background="#071015" Text="VM grid supports vertical and horizontal scrolling; first five columns remain frozen."/><DataGrid x:Name="gridP2" Grid.Row="3" MinHeight="300" MaxHeight="560" AutoGenerateColumns="False" CanUserAddRows="False" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" EnableRowVirtualization="True" VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" ScrollViewer.CanContentScroll="True" FrozenColumnCount="5"><DataGrid.Columns><DataGridCheckBoxColumn Header="Include" Binding="{Binding Include}" Width="55"/><DataGridTextColumn Header="Group #" Binding="{Binding MobilityGroupNumber}" Width="65"/><DataGridTextColumn Header="VM Name" Binding="{Binding VMName}" Width="140" IsReadOnly="True"/><DataGridTextColumn Header="vTPM" Binding="{Binding Vtpm}" Width="55" IsReadOnly="True"/><DataGridTextColumn Header="Policy Assignment" Binding="{Binding StoragePolicyAssignment}" Width="150" IsReadOnly="True"/><DataGridTextColumn Header="Network Mapping" Binding="{Binding NetworkSummary}" Width="260" IsReadOnly="True"/><DataGridTemplateColumn Header="Compute" Width="200"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox ItemsSource="{Binding DataContext.Computes, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding ComputeId, Mode=TwoWay}" MaxDropDownHeight="420"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTextColumn Header="Folder" Binding="{Binding Folder}" Width="110"/><DataGridTemplateColumn Header="Destination Storage" Width="215"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox Tag="{Binding}" ItemsSource="{Binding DataContext.Datastores, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="DisplayName" SelectedValuePath="Id" SelectedValue="{Binding StorageId, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" MaxDropDownHeight="420"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTemplateColumn Header="Effective Storage Policy" Width="210"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox Tag="{Binding}" ItemsSource="{Binding DataContext.Policies, RelativeSource={RelativeSource AncestorType=Window}}" DisplayMemberPath="Name" SelectedValuePath="Name" SelectedValue="{Binding StoragePolicy, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" MaxDropDownHeight="420"/></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn><DataGridTemplateColumn Header="Disk Format" Width="180"><DataGridTemplateColumn.CellTemplate><DataTemplate><ComboBox Tag="{Binding}" SelectedValue="{Binding DiskFormat, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" SelectedValuePath="Content"><ComboBoxItem Content="Same format as source"/><ComboBoxItem Content="Thin Provision"/><ComboBoxItem Content="Thick Provision Lazy Zeroed"/><ComboBoxItem Content="Thick Provision Eager Zeroed"/></ComboBox></DataTemplate></DataGridTemplateColumn.CellTemplate></DataGridTemplateColumn></DataGrid.Columns></DataGrid>
<StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,6,0,0"><TextBlock Text="Phase II: "/><TextBlock x:Name="lblP2Status" Text="Not validated" Foreground="#76C7D8"/><Button x:Name="btnP2Validate" Content="Validate"/><Button x:Name="btnP2Preview" Content="Preview Payload" IsEnabled="False"/><Button x:Name="btnP2Create" Content="Save Draft to HCX" IsEnabled="False"/></StackPanel></Grid></TabItem></TabControl>
<GroupBox Grid.Row="2" Header="Log"><TextBox x:Name="txtLog" IsReadOnly="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12" Background="#071015" Foreground="#E6E6E6"/></GroupBox>
<GroupBox Grid.Row="3" Header="Output Location" Margin="5"><Grid><Grid.ColumnDefinitions><ColumnDefinition Width="120"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions><TextBlock Text="Default base path" VerticalAlignment="Center"/><TextBox x:Name="txtOutputPath" Grid.Column="1" Margin="4"/><StackPanel Grid.Column="2" Orientation="Horizontal"><Button x:Name="btnBrowseOutput" Content="Browse..."/><Button x:Name="btnApplyOutput" Content="Apply Path"/></StackPanel></Grid></GroupBox><StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Center"><Button x:Name="btnOpenOutput" Content="Open Last CSV"/><Button x:Name="btnOpenRun" Content="Open Run Folder"/><Button x:Name="btnClose" Content="Close"/></StackPanel>
</Grid></Window>
'@
$script:Window=[Windows.Markup.XamlReader]::Parse($xaml)
$script:MigrationTypes=@('HCX Assisted vMotion (Direct)')
$script:Window.DataContext=[pscustomobject]@{
    MigrationTypes=$script:MigrationTypes
    Computes=@()
    Datastores=@()
    Folders=@()
    Networks=@()
    Sites=@()
    Policies=@()
}

$names=@('lblPS','lblSTA','lblPCLI','lblApi','btnRecheck','btnInstall','txtHcx','txtUser','txtPassword','txtSourceVC','txtSourceVCUser','txtSourceVCPass','txtDestinationVC','txtDestinationVCUser','txtDestinationVCPass','btnConnect','btnDisconnect','btnSaveConnectionJson','btnLoadConnectionJson','lblHcxTopologyStatus','lblHcxDirection','lblHcxServiceMesh','cmbHcxServiceMesh','cmbDestinationSite','cmbGlobalCompute','cmbGlobalDatastore','cmbGlobalFolder','cmbMigrationType','cmbGlobalStoragePolicy','cmbVtpmStoragePolicy','cmbGlobalDiskFormat','txtNetworkMappingCsv','btnImportNetworkMapping','btnClearNetworkMapping','lblNetworkMappingStatus','btnApplyGlobal','gridVMs','btnImport','btnExample','btnExampleNetworkMapping','btnRemove','btnClear','btnRefresh','lblValidation','btnValidate','btnCreateCsv','txtLog','txtOutputPath','btnBrowseOutput','btnApplyOutput','btnOpenOutput','btnOpenRun','btnClose','txtP2Csv','btnP2Import','txtP2Name','cmbP2GroupCount','txtP2Hcx','txtP2SourceSite','txtP2DestSite','cmbP2Compute','cmbP2Storage','cmbP2Policy','cmbP2VtpmPolicy','cmbP2Disk','btnP2Inventory','btnP2Apply','rbP2Now','rbP2Schedule','rbP2Defer','txtP2Schedule','chkP2Mac','chkP2HW','chkP2Tools','chkP2Attrs','chkP2Tags','chkP2Iso','chkP2Security','gridP2','lblP2Status','btnP2Validate','btnP2Preview','btnP2Create')
foreach($n in $names){Set-Variable -Scope Script -Name $n -Value $script:Window.FindName($n)}
function Test-HcxRequiredFunctions {
    $required=@('Convert-HcxNetworkIdCanonical','Resolve-HcxDestinationNetworkMatch','Set-HcxBindingContextCollection','Bind-DestinationControls','Invalidate-Validation','Invalidate-P2Validation','P2-Invalidate','Commit-Grid','Commit-P2Grid','Apply-GlobalSelection','Validate-Phase1','Initialize-HcxAutomaticDiscovery','Assert-HcxAutomaticDiscoveryCurrent','Get-HcxDiscoveredPayloadTopology','Merge-HcxAutomaticNetworkInventory','Invoke-HcxRest','Assert-HcxSubmissionAuthentication','Refresh-HcxAuthentication','Submit-HcxMobilityGroup')
    $missing=@($required|Where-Object{-not(Get-Command $_ -CommandType Function -ErrorAction SilentlyContinue)})
    if($missing.Count){throw "Script initialization failed. Missing required function(s): $($missing -join ', ')."}
    Write-HcxDebug "Required function self-test passed: $($required -join ', ')." 'SELF-TEST'
}
$script:gridVMs.ItemsSource=$script:Rows
Test-HcxRequiredFunctions
Start-HcxDiagnosticTranscript
Write-HcxDebug "Debug logging enabled by default. Log=$script:LogFile; Transcript=$script:TranscriptFile; ArtifactFolder=$script:DebugArtifactDir" 'STARTUP'
foreach($n in 1..50){[void]$script:cmbP2GroupCount.Items.Add($n)};$script:cmbP2GroupCount.SelectedIndex=0;
$script:txtOutputPath.Text=$script:OutputBase
$script:chkP2Security.IsChecked=$false
if ([string]::IsNullOrWhiteSpace($script:txtUser.Text)) { $script:txtUser.Text = 'administrator@vsphere.local' }
if ([string]::IsNullOrWhiteSpace($script:txtSourceVCUser.Text)) { $script:txtSourceVCUser.Text = 'administrator@vsphere.local' }
if ([string]::IsNullOrWhiteSpace($script:txtDestinationVCUser.Text)) { $script:txtDestinationVCUser.Text = 'administrator@vsphere.local' }

$script:Window.AddHandler([Windows.Controls.Button]::ClickEvent,[Windows.RoutedEventHandler]{param($sender,$e);if($e.OriginalSource.Name -eq '' -and $e.OriginalSource.Content -eq 'Configure' -and $e.OriginalSource.Tag){Show-NicMappingDialog $e.OriginalSource.Tag}})
$script:gridVMs.Add_CellEditEnding({Invalidate-Validation 'VM grid edited'})
$script:gridVMs.AddHandler([Windows.Controls.ComboBox]::DropDownClosedEvent,[Windows.RoutedEventHandler]{param($s,$e);$c=$e.OriginalSource;if($c-is[Windows.Controls.ComboBox]-and$c.Tag-and$c.SelectedItem){$c.Tag.StoragePolicy=$c.SelectedItem.Name;$c.Tag.StoragePolicyAssignment='Per-VM Override';$script:gridVMs.Items.Refresh();Invalidate-Validation 'Per-VM storage policy override changed'}})
$script:btnRecheck.Add_Click({Update-Prerequisites})
$script:btnInstall.Add_Click({$null=Ensure-Module 'VCF.PowerCLI';Update-Prerequisites})
$script:btnConnect.Add_Click({
    try{
        $script:btnConnect.IsEnabled=$false
        Reset-HcxAutomaticDiscovery
        Connect-Hcx91Rest $script:txtHcx.Text.Trim() $script:txtUser.Text.Trim() $script:txtPassword.Password
        Load-HcxInventory;Initialize-HcxAutomaticDiscovery;Merge-HcxAutomaticNetworkInventory;Update-Prerequisites;if($script:P2Rows -and $(@($script:P2Rows).Count) -gt 0){P2-Inventory;Restore-P2SelectionsFromPhaseICsv}
        [Windows.MessageBox]::Show("Connected. Inventory loading completed.`n`nVMs: $($(@($script:Inventory.VMs).Count))`nNetworks: $($(@($script:Inventory.Networks).Count))`nDatastores: $($(@($script:Inventory.Datastores).Count))`nComputes: $($(@($script:Inventory.Computes).Count))`nFolders: $($(@($script:Inventory.Folders).Count))",'HCX connected')|Out-Null
    }catch{$artifact=Write-HcxExceptionDiagnostic $_ 'HCX-CONNECTION';Log "$($_.Exception.Message) Diagnostic=$artifact" ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'HCX connection failed','OK','Error')|Out-Null}
    finally{$script:btnConnect.IsEnabled=$true;Update-Prerequisites}
})
$script:btnDisconnect.Add_Click({Disconnect-HcxRest;Reset-HcxAutomaticDiscovery;Log 'Disconnected from HCX.' INFO})
$script:btnSaveConnectionJson.Add_Click({
    try { Save-ConnectionProfile }
    catch { Log $_.Exception.Message ERROR; [Windows.MessageBox]::Show($_.Exception.Message,'Save JSON failed','OK','Error') | Out-Null }
})
$script:btnLoadConnectionJson.Add_Click({
    try { Load-ConnectionProfile }
    catch { Log $_.Exception.Message ERROR; [Windows.MessageBox]::Show($_.Exception.Message,'Load JSON failed','OK','Error') | Out-Null }
})

$script:cmbHcxServiceMesh.Add_SelectionChanged({
    if($script:HcxDiscovery -and $script:cmbHcxServiceMesh.SelectedItem){
        $script:HcxDiscovery.SelectedServiceMesh=$script:cmbHcxServiceMesh.SelectedItem
        $script:lblHcxServiceMesh.Text=$script:HcxDiscovery.SelectedServiceMesh.DisplayName
        Log "Service Mesh selected: $($script:HcxDiscovery.SelectedServiceMesh.DisplayName)." PASS
        P2-Invalidate
    }
})
$script:btnApplyGlobal.Add_Click({Apply-GlobalSelection})
$script:btnImportNetworkMapping.Add_Click({try{Import-HcxNetworkMappingCsv}catch{$detail=Get-HcxDetailedError $_;Log $detail ERROR;[Windows.MessageBox]::Show($detail,'Network Mapping CSV Import','OK','Error')|Out-Null}})
$script:btnClearNetworkMapping.Add_Click({Clear-HcxNetworkMappingCsv})
function Show-DetailedVmImportSummary($Rows){
 $all=@($Rows);$items=@();foreach($r in $all){$finding=if($r.Status -eq 'NotFound'){'Missing from inventory'}elseif($r.Status -eq 'Ambiguous'){'Ambiguous inventory match'}elseif($r.VtpmDetectionStatus -ne 'Detected'){'vTPM detection failed'}elseif($r.PowerState -eq 'PoweredOff'){'Powered off'}else{'Present and powered on'};$items+=[pscustomobject]@{VMName=$r.VMName;InventoryStatus=$r.Status;PowerState=if($r.PowerState){$r.PowerState}else{'Unknown'};vTPM=if($null-eq$r.Vtpm){'Unknown'}else{[string]$r.Vtpm};VtpmDetection=$r.VtpmDetectionStatus;Finding=$finding}}
 $missing=@($items|Where-Object Finding -in @('Missing from inventory','Ambiguous inventory match'));$off=@($items|Where-Object Finding -eq 'Powered off');$bad=@($items|Where-Object Finding -eq 'vTPM detection failed');$ready=@($items|Where-Object Finding -eq 'Present and powered on')
 $x=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="VM Import Summary" Height="590" Width="950" WindowStartupLocation="CenterOwner" Background="#071015" Foreground="#E6E6E6" FontFamily="Segoe UI"><Grid Margin="14"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions><TextBlock x:Name="summary" FontSize="16" FontWeight="SemiBold" TextWrapping="Wrap"/><TextBlock Grid.Row="1" Text="Import continued. Review highlighted VMs before Phase I validation. Red indicates missing, ambiguous, or failed vTPM detection. Gold indicates powered off." Foreground="#76C7D8" TextWrapping="Wrap" Margin="0,7,0,10"/><DataGrid x:Name="grid" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" EnableRowVirtualization="True" VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" ScrollViewer.CanContentScroll="True"><DataGrid.RowStyle><Style TargetType="DataGridRow"><Style.Triggers><DataTrigger Binding="{Binding Finding}" Value="Missing from inventory"><Setter Property="Background" Value="#8B1E2D"/><Setter Property="Foreground" Value="White"/></DataTrigger><DataTrigger Binding="{Binding Finding}" Value="Ambiguous inventory match"><Setter Property="Background" Value="#8B1E2D"/><Setter Property="Foreground" Value="White"/></DataTrigger><DataTrigger Binding="{Binding Finding}" Value="vTPM detection failed"><Setter Property="Background" Value="#8B1E2D"/><Setter Property="Foreground" Value="White"/></DataTrigger><DataTrigger Binding="{Binding Finding}" Value="Powered off"><Setter Property="Background" Value="#8A6500"/><Setter Property="Foreground" Value="White"/></DataTrigger></Style.Triggers></Style></DataGrid.RowStyle><DataGrid.Columns><DataGridTextColumn Header="VM Name" Binding="{Binding VMName}" Width="220"/><DataGridTextColumn Header="Inventory" Binding="{Binding InventoryStatus}" Width="120"/><DataGridTextColumn Header="Power" Binding="{Binding PowerState}" Width="110"/><DataGridTextColumn Header="vTPM" Binding="{Binding vTPM}" Width="70"/><DataGridTextColumn Header="vTPM Detection" Binding="{Binding VtpmDetection}" Width="130"/><DataGridTextColumn Header="Finding" Binding="{Binding Finding}" Width="*"/></DataGrid.Columns></DataGrid><Button x:Name="ok" Grid.Row="3" Content="Continue" Width="110" HorizontalAlignment="Right" Margin="0,12,0,0"/></Grid></Window>
'@
 $d=[Windows.Markup.XamlReader]::Parse($x);$d.Owner=$script:Window;$d.FindName('summary').Text="Imported $($all.Count) VM(s): $($ready.Count) present and powered on, $($off.Count) powered off, $($missing.Count) missing or ambiguous, $($bad.Count) vTPM detection failure(s).";$d.FindName('grid').ItemsSource=$items;$d.FindName('ok').Add_Click({$d.DialogResult=$true;$d.Close()});$null=$d.ShowDialog()
}
function Get-SafeCount($Value){ return @($Value).Count }
function Get-HcxDetailedError($ErrorRecord){
    $line='';$position='';$stack=''
    try{$line=[string]$ErrorRecord.InvocationInfo.ScriptLineNumber}catch{}
    try{$position=[string]$ErrorRecord.InvocationInfo.PositionMessage}catch{}
    try{$stack=[string]$ErrorRecord.ScriptStackTrace}catch{}
    return "Message=$($ErrorRecord.Exception.Message)`nLine=$line`nPosition=$position`nStack=$stack"
}
$script:btnImport.Add_Click({
    try{
        if(-not$script:Hcx.Connected){throw 'Connect to HCX and load inventory before importing VMs.'}
        $d=[Microsoft.Win32.OpenFileDialog]::new();$d.Filter='CSV files (*.csv)|*.csv';if(-not$d.ShowDialog()){return}
        $input=@(Import-Csv -LiteralPath $d.FileName)
        if(@($input).Count -eq 0){throw 'The selected CSV contains no rows.'}
        $newRows=[Collections.Generic.List[object]]::new()
        $rowNumber=1
        foreach($i in @($input)){
            try{
                $name=[string](Get-Value $i @('VMName','Name'))
                if([string]::IsNullOrWhiteSpace($name)){throw "CSV row $rowNumber does not contain VMName or Name."}
                $row=Convert-ToVmRow $name.Trim()
                $mg=[string](Get-Value $i @('MobilityGroupNumber','MobilityGroup#','MobilityGroup'))
                if($mg){$row.MobilityGroupNumber=[int]$mg}
                $newRows.Add($row)
            }catch{
                throw "Phase I CSV row $rowNumber failed. VM='$([string](Get-Value $i @('VMName','Name')))'. $(Get-HcxDetailedError $_)"
            }
            $rowNumber++
        }
        $script:Rows.Clear()
        foreach($row in $newRows){$script:Rows.Add($row)}
        Apply-GlobalSelection
        if($script:ImportedNetworkMappings.Count -gt 0){$null=Apply-ImportedNetworkMappingsToRows}
        Write-HcxDebug "Phase I import grid populated and global selections applied. RowCount=$(@($script:Rows).Count)." 'PHASE1-IMPORT'
        Log "Imported and discovered $(@($script:Rows).Count) VM row(s) from $($d.FileName)." PASS
        Show-DetailedVmImportSummary -Rows @($script:Rows)
    }catch{
        $detail=Get-HcxDetailedError $_
        Log "PHASE I IMPORT FAILURE: $detail" ERROR
        [Windows.MessageBox]::Show($detail,'Import failed','OK','Error')|Out-Null
    }
})
$script:btnExampleNetworkMapping.Add_Click({$p=Join-Path $script:OutputBase 'example-network-mapping.csv';@([pscustomobject]@{SourceNetworkName='Legacy-App-Network';DestinationNetworkName='New-App-Network'},[pscustomobject]@{SourceNetworkName='Legacy-Database-Network';DestinationNetworkName='New-Database-Network'})|Export-Csv -LiteralPath $p -NoTypeInformation -Encoding utf8BOM;Log "Example network mapping CSV created: $p" PASS;Invoke-Item $p})
$script:btnExample.Add_Click({$p=Join-Path $script:OutputBase 'example-vm-import.csv';@([pscustomobject]@{VMName='appserver01';MobilityGroupNumber=1},[pscustomobject]@{VMName='secureapp01';MobilityGroupNumber=2})|Export-Csv -LiteralPath $p -NoTypeInformation -Encoding utf8BOM;Log "Example CSV created: $p" PASS;Invoke-Item $p})
$script:btnRemove.Add_Click({foreach($r in @($script:gridVMs.SelectedItems)){$script:Rows.Remove($r)};Invalidate-Validation 'Selected VM rows removed'})
$script:btnClear.Add_Click({$script:Rows.Clear();Invalidate-Validation 'VM grid cleared'})
$script:btnRefresh.Add_Click({
    try{Load-HcxInventory;Initialize-HcxAutomaticDiscovery;Merge-HcxAutomaticNetworkInventory;$names=@($script:Rows|ForEach-Object{$_.VMName});$script:Rows.Clear();foreach($n in $names){$script:Rows.Add((Convert-ToVmRow $n))};Apply-GlobalSelection;Log 'HCX inventory and VM discovery refreshed.' PASS}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Refresh failed','OK','Error')|Out-Null}
})
$script:btnValidate.Add_Click({$null=Validate-Phase1})
$script:btnCreateCsv.Add_Click({try{Export-Phase1Csv}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'CSV creation failed','OK','Error')|Out-Null}})
$script:btnBrowseOutput.Add_Click({try{Select-OutputBase}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Output path failed','OK','Error')|Out-Null}})
$script:btnApplyOutput.Add_Click({try{Set-OutputBase $script:txtOutputPath.Text}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Output path failed','OK','Error')|Out-Null}})
$script:btnOpenOutput.Add_Click({if($script:LastCsv -and(Test-Path -LiteralPath $script:LastCsv)){Invoke-Item $script:LastCsv}else{[Windows.MessageBox]::Show('No Phase I CSV has been created in this session.','No CSV')|Out-Null}})
$script:btnOpenRun.Add_Click({Invoke-Item $script:RunDir})
$script:btnClose.Add_Click({Disconnect-HcxRest;Stop-HcxDiagnosticTranscript;$script:Window.Close()})
$script:Window.Add_Closing({
    try{$script:txtPassword.Clear();$script:txtSourceVCPass.Clear();$script:txtDestinationVCPass.Clear()}catch{}
    try{Disconnect-HcxRest}catch{}
    try{Stop-HcxDiagnosticTranscript}catch{}
    $script:Hcx.Session=$null;$script:Hcx.Headers=@{}
})
$script:btnP2Import.Add_Click({try{P2-Import}catch{$detail=Get-HcxDetailedError $_;Log $detail ERROR;[Windows.MessageBox]::Show($detail,'Phase II Import')|Out-Null}})
$script:btnP2Inventory.Add_Click({try{P2-Inventory}catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Phase II Storage')|Out-Null}})
$script:btnP2Apply.Add_Click({P2-Apply});$script:btnP2Validate.Add_Click({P2-Validate});$script:btnP2Preview.Add_Click({P2-Preview});$script:btnP2Create.Add_Click({
    try{
        if(-not$script:P2Payload){try{$script:Window.Cursor='Wait';$script:P2Payload=@(P2-BuildAll)}finally{$script:Window.Cursor=$null}}
        P2-Create
    }catch{Log $_.Exception.Message ERROR;[Windows.MessageBox]::Show($_.Exception.Message,'Create Mobility Group')|Out-Null}
});$script:gridP2.Add_CellEditEnding({P2-Invalidate})
$script:gridP2.AddHandler([Windows.Controls.ComboBox]::DropDownClosedEvent,[Windows.RoutedEventHandler]{param($sender,$e);$cb=$e.OriginalSource;if($cb -isnot [Windows.Controls.ComboBox] -or -not $cb.Tag -or -not $cb.SelectedItem){return};$row=$cb.Tag;if(-not($script:P2Rows -contains $row)){return};$header=[string]$script:gridP2.CurrentColumn.Header;switch($header){'Destination Storage'{$selected=@($script:P2Storages|Where-Object Id -eq $cb.SelectedValue|Select-Object -First 1);if($selected){$row.Storage=$selected.Name;$row.StorageId=$selected.Id;$row.StorageType=$selected.Type;$row.StorageSelectionSource='Phase II Per-VM Override'}}'Effective Storage Policy'{$row.StoragePolicy=$cb.SelectedItem.Name;$row.StoragePolicyAssignment='Phase II Per-VM Override'}'Disk Format'{$row.DiskFormat=$cb.SelectedItem.Content;$row.DiskFormatSelectionSource='Phase II Per-VM Override'}};$script:gridP2.Items.Refresh();P2-Invalidate})

Update-Prerequisites
Log "HCX 9.1 Mobility Group CSV Builder Phase I started. Output folder: $script:RunDir" PASS
Log 'Multiple-vNIC mapping is enabled. Each adapter is validated independently before CSV creation.' INFO
Log 'Optional Phase I network-mapping CSV is enabled. Precedence: per-VM override, imported mapping CSV, automatic exact-name matching.' PASS
Log 'Rev 3.2 authentication refresh and centralized mobility-group submission, row dropdown binding, validation state, case-sensitive HCX JSON, resilient HTTP transport, and default verbose diagnostics active: current-session topology, direction, site-pair, Service Mesh, vCenter and NSX network discovery, scalar Count normalization, Phase I/II diagnostics, and placement normalization.' PASS
Log 'Host/StoragePod audit JSON files are written to the active per-launch HCX run folder.' PASS
Log 'Credentials remain in memory only while this PowerShell application window is open and are purged when the window closes.' PASS
$null=$script:Window.ShowDialog()
Stop-HcxDiagnosticTranscript

Log 'Phase I and Phase II disk format default is Same format as source. The import summary grid supports scrolling and row virtualization for large batches.' PASS

# SIG # Begin signature block
# MIIF7AYJKoZIhvcNAQcCoIIF3TCCBdkCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCzD+/X+2a2JVc
# OCtqbcHPrrt1GH0Q5OQdv0dfTA+YtKCCA0QwggNAMIICKKADAgECAhAvhbIkuuEX
# vkZE9AgHvpY4MA0GCSqGSIb3DQEBCwUAMDgxNjA0BgNVBAMMLUhDWDkxIE1vYmls
# aXR5IENTViBCdWlsZGVyIExvY2FsIENvZGUgU2lnbmluZzAeFw0yNjA5MTMxNjM5
# NThaFw0yOTA5MTMxNjQ5NTdaMDgxNjA0BgNVBAMMLUhDWDkxIE1vYmlsaXR5IENT
# ViBCdWlsZGVyIExvY2FsIENvZGUgU2lnbmluZzCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAL7NuhqUV3QFo3upMBnq9z6SFnT8tUPo5rMTrVsJw+49AR/1
# mmLhdYOri5aQOiJEPnpYGlltrdKQzITs6s9x67askLjUavuCvlflCUv8uHr8bBN5
# x5Zyg7yKobYd9hMsWmUIVj73hBockFza1TdVFU+HhRKM1LXJEAAucIfmUA8Wy9Lu
# chJJaGtJ7qXkGsHVWg5dQLKzDxwIJky9EeBpI03IXV37D+SVn9O+OLnN7Oo09wLG
# +gCaawl7uxeJI54ipR3kfplcz4DfK2oUifX56KgPzrjZtORWWXAueLlz63VRyxhg
# oKkHYYpIqIhFYj+ck0otl5EIum3oeGJs5ip46BECAwEAAaNGMEQwDgYDVR0PAQH/
# BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQVYe4TeDN03bYF
# j2xjYk4Q+zA4xDANBgkqhkiG9w0BAQsFAAOCAQEAqLxgbNOQ0kHRQGJavVB+qWgH
# ExOhZcupsuudmOC/ro6HI8gb5IHBjpY7xQS2qZ9oSjt8ATntBiQOJHfFEi4abcJs
# wcuMAa8dKV5qgmeUynlP8JLaNCd+RShR13kmSGYS8G3XavURWpp5SUG7KAD/YxzV
# xgG3sOadZ8CqAbFemdpmE2yn+qKQFkE1HOTHkQZ1oCuR3nyZu92l8kmnWfI43/0k
# yaJBHr0ZXeXpmKQ7lK54UOgA+jjtJfcpiKWzvU62dXuHNOwkaF/xQWiXOjqYxBLA
# J1vA2NFN6l5isjZ8Vco98h7jpOE6SnSGFzXsPWJnCZEnG46cY5/ScPcV/7C9TDGC
# Af4wggH6AgEBMEwwODE2MDQGA1UEAwwtSENYOTEgTW9iaWxpdHkgQ1NWIEJ1aWxk
# ZXIgTG9jYWwgQ29kZSBTaWduaW5nAhAvhbIkuuEXvkZE9AgHvpY4MA0GCWCGSAFl
# AwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkD
# MQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJ
# KoZIhvcNAQkEMSIEIHbDRsjKw/W/LdyHmfZvlb6zg6cPakXmIhb26ehRHKBeMA0G
# CSqGSIb3DQEBAQUABIIBABFeNFVWCZgCrYGXGJ5zW8xXnKlzGc7cPIVFQpTeO8aS
# Iju7oNVLYldqCFdTC5owA5B9FsyQNJ2bUW8xy/PfsFsWD6KozJQilbb1uLYT0O8z
# 4gQ6dX0AH+CBMGrwRon4v9VsiYbMCvg2SQ0ewutxzejNZYQgdua9+t5wtN0QXkBL
# wmqNxT98W6s2oM9eoVSlwAgrZqa2unrEN6EHb2kN4RrsoQkXQwUZtDMLMxACQAaU
# N5JJHORT3pjYxReAvl2SVkUsk5h7ZyS2AeR3wUEpT2tthCRX1LQAYj1kBp65x2xC
# 6qPrBJQQVYUzLc2/jC323O9p/UEIljWuIrRFkd3mmEI=
# SIG # End signature block
