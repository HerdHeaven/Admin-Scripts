param ([string]$macaddy=(Read-host -Prompt "Please Enter the MAC Address"))
    $mymac = $macaddy.split(':') | %{ [byte]('0x' + $_) }
    if ($mymac.Length -ne 6)
    {
        throw 'Mac Address Must be 6 hex Numbers Separated by : or -'
    }
    Write-Verbose "Creating UDP Packet"
    $UDPclient = new-Object System.Net.Sockets.UdpClient
    $UDPclient.Connect(([System.Net.IPAddress]::Broadcast),4000)
    $packet = [byte[]](,0xFF * 6)
    $packet += $mymac * 16
    Write-Verbose ([bitconverter]::tostring($packet))
    [void] $UDPclient.Send($packet, $packet.Length)
    Write-Host  "   - Wake-On-Lan Packet of length $($packet.Length) sent to $mymac"
