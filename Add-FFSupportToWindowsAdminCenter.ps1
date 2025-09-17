#Force support for Firefox in previous Windows Admin Center versions
    #This is supposedly corrected in the latest GA release

#Find all the javascript files to fix the origin reference
foreach($file in (get-childitem -path $env:ProgramData\WindowsAdminCenter\UX\* -Include "*.js","*.css"))
{
    $content = Get-content $file

    switch($content)
    {
        #find the one containing the bad reference
        {$PSItem -ilike "*window.location.ancestorOrigins*"} 
            {
                #replace with the more acceptable reference
                $newcontent = $content.Replace("window.location.ancestorOrigins", "window.location.origin") 
                Set-Content -Path $file.FullName -Value $newContent -Force
            }
        #fix a small but annoying stylesheet issue
        {$PSItem -ilike "*.sme-layout-float-left{float:left}*"}
            {
                $newcontent = $content.Replace(".sme-layout-float-left{float:left}", "") 
                Set-Content -Path $file.FullName -Value $newContent -Force
            }
            default{Continue}
    }

    clear-variable content
}
