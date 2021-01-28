$localIPs = (Get-CimInstance -ClassName 'Win32_NetworkAdapterConfiguration' | `
    Where-Object {$_.IPAddress}).ipaddress | `
        Select-Object -Unique
$localIPs += '127.0.0.1' 
$netstat = Get-NetTCPConnection -State 'Established' | `
    Where-Object {$localIPs -notcontains $_.RemoteAddress -and $_.LocalAddress -match '^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$'} | `
        Select-Object -Property LocalAddress, RemoteAddress, RemotePort, State, OwningProcess -Unique


        $processes = Get-Process
foreach ($object in $netstat)
{
    try 
    {
        Write-Verbose "Updating process information for PID $($object.OwningProcess)."
        $processName = ($processes | Where-Object {$_.Id -eq $object.OwningProcess}).ProcessName
        $object | Add-Member -Name ProcessName -MemberType NoteProperty -Value $processName
        Write-Verbose "Updated information for PID $($object.OwningProcess)."
    }
    catch 
    {
        Write-Warning "Unable to get process information for PID $($object.OwningProcess)."
        $ErrorMessage = $_.Exception.Message
        Write-Warning "$ErrorMessage"
    }             
}
#


function Get-GeoIP
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Position=0, Mandatory=$True)]
        [string]$ipv4Address
    )
    
    $method = 'GET'
    $uri = "https://ipapi.co/$ipv4Address/json/"
    
    try 
    {
        Write-Verbose "Invoking REST method, $method, to $uri."
        $response = Invoke-RestMethod -Method $method -Uri $uri
        Write-Verbose "Invoked REST method, $method, to $uri."
        return $response
    }
    catch 
    {
        Write-Warning "Unable to invoked REST method, $method, to $uri."
        $ErrorMessage = $_.Exception.Message
        Write-Warning "$ErrorMessage"
    }
}



$uniqueRemoteAddresses = $netstat.RemoteAddress | Select-Object -Unique
foreach ($address in $uniqueRemoteAddresses)
{
    try 
    {
        Write-Verbose "Getting geo-IP information for $address."
        $geoIp = Get-GeoIP -ipv4Address $address
        [array]$geoIPs += $geoIp
        Write-Verbose "Got geo-IP information for $address."
    }
    catch 
    {
        Write-Warning "Unable to get get-IP information for $address."
        $ErrorMessage = $_.Exception.Message
        Write-Warning "$ErrorMessage"
    }        
}


foreach ($connection in $netstat)
{
    $geoIpToAdd = $geoIPs | Where-Object {$_.ip -like $connection.RemoteAddress}
    $connection | Add-Member -MemberType NoteProperty -Name 'geoIp' -Value $geoIpToAdd -Force
}

$xmldoc = @"
<directedgraph>
  <page>
    <renderoptions
      usedynamicconnectors="true"
      scalingfactor="20" />
    <shapes>
    </shapes>

    <connectors>
    </connectors>

  </page>

</directedgraph>
"@

$xmldoc | Out-File -FilePath C:\temp\directedGraphModel.xml -Force
$modelFile = Get-Item -Path C:\temp\directedGraphModel.xml
$model = Import-VisioModel -FileName $modelFile.FullName

$nodes = $netstat.LocalAddress + $netstat.RemoteAddress | Select-Object -Unique
$stencil = "basic_u.vssx"
$master = "Circle"

foreach($node in $nodes)
{
    $model.Layouts.AddNode($node, $node, $stencil, $master) | Out-Null
}

foreach($connection in $netstat)
{
    $model.Layouts.AddEdge($netstat.indexof($connection), $connection.LocalAddress, $connection.RemoteAddress, "$($connection.LocalAddress)-&gt;$($connection.RemoteAddress)", 'curved') | Out-Null
}

New-VisioApplication
Out-VisioApplication($model)