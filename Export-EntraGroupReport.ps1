#Requires -modules ImportExcel, Microsoft.Graph.Authentication, Microsoft.Graph.Groups

Import-Module Microsoft.Graph.Authentication, Microsoft.Graph.Groups, ImportExcel

function New-GraphConnection
{
    param(
        $clientId = "CLIENT ID HERE",
        $tenantId = "TENANT ID HERE",
        $CertificateThumbPrint = "CERTIFICATE THUMBPRINT HERE"
    )


    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbPrint $CertificateThumbPrint
}

function Get-GraphGroups
{
    param([array]$properties=@("Id","DisplayName","Description","GroupTypes"))
    Get-MgGroup -All -Property $properties | Select-Object -Property $properties
}

function Get-GraphGroupMembers
{
    param([string]$Id,
    [array]$properties = @("id","givenName","surname","AccountEnabled","userPrincipalName","employeeId","EmployeeType"))
    
    Get-MgGroupMemberAsUser -GroupId $ID -Property $properties | Select-object -Property $properties

}

Function Add-GroupReportLinks
{
    param([string]$filePath)

    # Load the Excel workbook
    $worksheet = Import-Excel -Path $filePath -WorksheetName "Group Names"

    # Iterate through each entry in the 2nd column
    $updated = foreach ($entry in $worksheet) {
        # Get the name of the worksheet to link to

        $groupname = $entry.DisplayName
        $worksheetName = $entry.MailNickname

        if($worksheetName.length -gt 31)
        {
            # Change the entry to a link to the worksheet
            $entry.DisplayName = "=HYPERLINK(`"#'$($worksheetName.substring(0,31))'!A1`", `"$groupname`")"
        }
        else
        {
            # Change the entry to a link to the worksheet
            $entry.DisplayName = "=HYPERLINK(`"#'$worksheetName'!A1`", `"$groupname`")"
        }
        
        $entry
    }

    # Save the modified Excel workbook
    $updated | Export-Excel -Path $filePath -WorksheetName "Group Names"

}

Function Wait-FileLocked
{
    param([string]$filePath)
    do{
        Try {
            $FileStream = [System.IO.File]::Open($filePath,'Open','Write')
        
            $FileStream.Close()
            $FileStream.Dispose()
        
            $IsLocked = $False
        } Catch [System.UnauthorizedAccessException] {
            $IsLocked = 'AccessDenied'
            Write-Warning "Locked.."
        } Catch {
            $IsLocked = $True
            Write-Warning "Locked.."
        }
    }until($IsLocked -eq $false)
}

Function New-GroupReport 
{
    param([string]$egrReportPath = $(Get-FilePath))

    #$egrReportName = $egrReportPath.Split("\") | ?{$_ -imatch "([\w]|[\d])+.xlsx"}

    $egrProperties = @(
        "Id",
        "DisplayName",
        "MailNickname",
        "Description",
        "GroupTypes",
        "IsAssignableByRole",
        "MailEnabled",
        "Mail",
        "ProxyAddresses",
        "CreatedDateTime",
        "ExpirationDateTime"
    )

    New-GraphConnection

    $egrGroups = Get-GraphGroups -properties $egrProperties | Sort-Object displayName

    # Iterate through each group
    foreach ($egrGroup in $egrGroups) {
        # Get members of the current group
        $egrMembers = Get-GraphGroupMembers -Id $egrGroup.Id | Sort-Object SurName

        #Write-Output "$($egrGroup.MailNickname):$($egrMembers.count)"

        if(Test-Path -Path $reportPath)
        {
            Wait-FileLocked -filePath $reportPath
        }

        if($null -ne $egrMembers)
        {
            do{
                try{
                    Export-Excel -Path $egrreportPath -WorksheetName $egrGroup.MailNickname -InputObject $egrMembers -AutoSize -TableStyle Light6 -Title "$($egrGroup.displayname)"

                    $IsSuccess = $true
                }catch [System.Management.Automation.MethodInvocationException] {
                    Write-Output "Error attempting to save file, trying again.."
                    $IsSuccess = $false
                }
            }until($IsSuccess -eq $true)
        }
        else {
            do{
                try{
                    Export-Excel -Path $egrreportPath -WorksheetName $egrgroup.MailNickname -Title "$($egrgroup.displayname)"

                    $IsSuccess = $true
                }catch [System.Management.Automation.MethodInvocationException] {
                    Write-Output "Error attempting to save file, trying again.."
                    $IsSuccess = $false
                }
            }until($IsSuccess -eq $true)
        }
    
        Start-Sleep -Milliseconds 400
    }

    Start-Sleep -Milliseconds 400

    # Export group names to worksheet
    do{
        try{
            $egrworkbook = Export-Excel -Path $egrreportPath -WorksheetName "Group Names" -InputObject $egrgroups -MoveToStart -AutoSize -BoldTopRow -TableStyle Light6 -PassThru

            $IsSuccess = $true
        }catch [System.Management.Automation.MethodInvocationException] {
            Write-Output "Error attempting to save file, trying again.."
            $IsSuccess = $false
        }
    }until($IsSuccess -eq $true)


    $egrworkbook.Save()

    #Add-GroupWorksheetLinks
    $egrWorksheets = ($egrworkbook.Workbook.Worksheets).count

    foreach($egrworksheet in $egrworkbook.Workbook.Worksheets[2..$egrWorksheets])
    {
        ($egrworksheet.Cells.Item("A1")).Hyperlink = "#'Group Names'!A1"
    }
    
    Start-Sleep -Milliseconds 300

    $egrworkbook.Save()

    $egrworkbook.Dispose()

    Start-Sleep -Milliseconds 400

    Add-GroupReportLinks -filePath $egrReportPath
}

Function Get-FilePath 
{   
    param([string]$initialDirectory=$env:USERPROFILE)
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.initialDirectory = $initialDirectory
        $SaveFileDialog.filter = “Excel Workbook (*.xlsx)|*.xlsx”
        $SaveFileDialog.ShowDialog() | Out-Null
        $SaveFileDialog.FileName
}

Function Start-Cleanup
{
    param([switch]$err)

    Clear-Variable egr*
    Remove-Variable egr*
    [gc]::Collect()

    if($err)
    {
        Exit 1
    }else {
        Exit 0
    }
    
}


New-GroupReport

Start-Sleep -Milliseconds 500

Start-Cleanup


