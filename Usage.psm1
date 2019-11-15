Function Usage{
Add-PSSnapin “Microsoft.SharePoint.PowerShell" -ea Continue
$path = "D:\SearchReports\"
# Set the intranet portal location 
$WebURL = “https://teams.sharepoint.COMPANYNAME.com/sites/SITE/”

$allsites=Get-SPWebApplication 'portal'| Get-spsite -Limit all
function Get-SPSearchReports ($url, $searchreport, $path) 
{ 
  # function to run the usage reports and store locally on server 
  # Usage reports ID’s must match the environment, use View source on the page _layouts/15/Reporting.aspx?Category=AnalyticsSiteCollection to get the ID’s

  # Usage Reports for Prod 
   
   $Usage                   = "6bbf6e1c-d79a-45da-9ba0-d0c3332bf6e2" 
   $Number_of_Queries         = "df46e7fb-8ab0-4ce8-8851-6868a7d986ab" 
  
  # Search Reports forProd

  $Top_Queries_by_Day         = "06dbb459-b6ef-46d1-9bfc-deae4b2bda2d" 
  $Top_Queries_by_Month       = "8cf96ee8-c905-4301-bdc4-8fdcb557a3d3" 
  $Abandoned_Queries_by_Day   = "5dd1c2fb-6048-440c-a60f-53b292e26cac" 
  $Abandoned_Queries_by_Month = "73bd0b5a-08d9-4cd8-ad5b-eb49754a8949" 
  $No_Result_Queries_by_Day   = "6bfd13f3-048f-474f-a155-d799848be4f1" 
  $No_Result_Queries_by_Month = "6ae835fa-3c64-40a7-9e90-4f24453f2dfe" 
  $Query_Rule_Usage_by_Day    = "8b28f21c-4bdb-44b3-adbe-01fdbe96e901" 
  $Query_Rule_Usage_by_Month  = "95ac3aea-0564-4a7e-a0fc-f8fdfab333f6"

  # set the file path and name 
  $filename = $path + $site.RootWeb.Title +" "+ (Get-Variable $searchreport).Name + ".xlsx" 
  $reportid = (Get-Variable $searchreport).Value

  $TTNcontent = "&__EVENTTARGET=__Page&__EVENTARGUMENT=ReportId%3D" + $reportid

  # setup the WebRequest 
  $webRequest = [System.Net.WebRequest]::Create($url) 
  $webRequest.UseDefaultCredentials = $true 
  $webRequest.Accept = "image/jpeg, application/x-ms-application, image/gif, application/xaml+xml, image/pjpeg, application/x-ms-xbap, */*" 
  $webRequest.ContentType = "application/x-www-form-urlencoded" 
  $webRequest.Method = "POST"

  $encodedContent = [System.Text.Encoding]::UTF8.GetBytes($TTNcontent) 
    $webRequest.ContentLength = $encodedContent.length 
    $requestStream = $webRequest.GetRequestStream() 
    $requestStream.Write($encodedContent, 0, $encodedContent.length) 
    $requestStream.Close()

  # get the data 
  [System.Net.WebResponse] $resp = $webRequest.GetResponse(); 
    $rs = $resp.GetResponseStream(); 
    #[System.IO.StreamReader] $sr = New-Object System.IO.StreamReader -argumentList $rs; 
    #[byte[]]$results = $sr.ReadToEnd(); 
    [System.IO.BinaryReader] $sr = New-Object System.IO.BinaryReader -argumentList $rs; 
    [byte[]]$results = $sr.ReadBytes(10000000);

  # write the file 
  Set-Content $filename $results -enc byte

}

Function Upload-Report($URL, $DocLibName, $FilePath) 
{ 
# function to upload the locally stored reports to SharePoint 
# Get a variable that points to the folder 
$Web = Get-SPWeb $URL 
$List = $Web.GetFolder($DocLibName) 
$Files = $List.Files

# Get just the name of the file from the whole path 
$FileName = $FilePath.Substring($FilePath.LastIndexOf("\")+1)

# Load the file into a variable 
$File= Get-ChildItem $FilePath

# Upload it to SharePoint 
$Files.Add($DocLibName +"/" + $FileName,$File.OpenRead(),$true) 
$web.Dispose() 
}
$FILES = Get-ChildItem -Path D:\SearchReports

Foreach($file in $FILES){
Upload-Report $WebURL "Site Usage" “D:\SearchReports\$File”
}
 <#
Get-SPSearchReports $url "Number_of_Queries" $path 
Get-SPSearchReports $url "Top_Queries_by_Day" $path 
Get-SPSearchReports $url "Top_Queries_by_Month" $path 
Get-SPSearchReports $url "Abandoned_Queries_by_Day" $path 
Get-SPSearchReports $url "Abandoned_Queries_by_Month" $path 
Get-SPSearchReports $url "No_Result_Queries_by_Day" $path 
Get-SPSearchReports $url "No_Result_Queries_by_Month" $path 
Get-SPSearchReports $url "Query_Rule_Usage_by_Day" $path 
Get-SPSearchReports $url "Query_Rule_Usage_by_Month" $path

# upload the reports

Upload-Report $WebURL "Audit Reports\Site Collection Usage" “D:\SearchReports\Usage.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\Number_of_Queries.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\Abandoned_Queries_by_Day.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\Abandoned_Queries_by_Month.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\No_Result_Queries_by_Day.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\No_Result_Queries_by_Month.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\Query_Rule_Usage_by_Day.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\Query_Rule_Usage_by_Month.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\Top_Queries_by_Day.xlsx” 
Upload-Report $WebURL "Audit Reports\Search" “D:\SearchReports\Top_Queries_by_Month.xlsx”#>
Foreach($site in $allsites){$si=$site.Url 
#This is the URL for site collection usage reports 
$urlend = "/_layouts/15/Reporting.aspx?Category=AnalyticsSiteCollection"
$url=$si+$urlend

# This is the path to write the reports to must exist on server running the script 

# delete anything in the d:\SearchReports folder 
#Remove-item D:\SearchReports\*

# run the reports

Get-SPSearchReports $url "Usage" $path}

}