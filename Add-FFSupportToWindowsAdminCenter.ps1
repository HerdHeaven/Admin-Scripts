#Force support for Firefox in Windows Admin Center

#Find all the javascript files
foreach($file in (get-item -path $env:ProgramData\WindowsAdminCenter\UX\*.js))
{
    $content = Get-content $file

    #find the one containing the bad reference 
    if($content -ilike "*window.location.ancestorOrigins*")
    {
        #replace with the more acceptable reference
        $newcontent = $content.Replace("window.location.ancestorOrigins", "window.location.origin")

        #save it!
        Set-Content -Path $file.FullName -Value $newContent -Force
    }
    clear-variable content
}
