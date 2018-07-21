#  Author: Michael Curtis
#  Version:  .8
#  Date  7/16/2018
#  Purpose:  Using Daily Referral Report, check Primary Care Provider in centricity matches PRI in OCeMedconnect
#            Checks for eligability and current date.  
#            If !eligable or current date_month == this_month  then do nothing
#            Else Match names and update Current Date field  with todays date.
#
#
#
# Constant defines
# Web Pages
$idx = "http://idx/gpmsweb"
$ocemed = "oc.medeconnect.com"
$emed_web_use = "MainContent_txtUsername"
$emed_web_pass = "MainContent_txtPassword"
$emed_web_login ="MainContent_btnLogonImage"
$emed_web_RTS = "https://oc.medeconnect.com/main/RealtimeSearch.aspx"
$emed_web_SubID = "tabContainer_tabSearch_txtCertificateNumber"
$emed_web_Submit = "tabContainer_tabSearch_btnHSubmit"

#HighScope Constants
$fixedCSV = @()

# Referral_Report Location xls  on COS
$local_Ref_Rep_Dir = "$home\desktop\Local Referral Report"
$ref_Rep_name = "Referral_Report.xls"
$rrl = "\\Analyzer\Analyzer\Snapshots\cos\Excel\daily\Referral_Report.xls"  
$header = 
$occreds = Get-Cred -Whatfor "OC Credentials"
$idxCreds = Get-Cred  -Whatfor "Centricity Credentials"
$emedCreds = Get-Cred  -Whatfor "oc.emedeconnect credentials"

# .net Interfaces used in manipulating the embedded Java applet
[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")

# Creates Local Referral report folder and downloads current report
if(!(Test-Path -Path "$home\desktop\Local Referral Report")){
		New-Item -Path $home\desktop\ -Name "Local Referral Report" -ItemType dir
	}
Start-BitsTransfer -Source $rrl -Credential (Get-Credential) -Destination "$home\desktop\Local Referral Report"


function Add-Data{
Param(
[Parameter(Mandatory=$true)][PSObject]$ReturnObject, 
[Parameter(Mandatory=$true)][int]$counter,
[Parameter(Mandatory=$true)][Object]$Dataobject)
$ReturnObject | Add-Member -type NoteProperty -Name $header[$counter] -Value $DataObject."$($header[$counter])"
}

#  dynamically  name a variable with a variable === $value=$NetworkInfo."$($_.Name)"
function Format-the-fucking-data {
$shitCSV = Import-Csv $rrl -Header $header 
foreach( $line in $shitCSV){
  $temp = New-Object PSObject
  $indexer = 0
  $line.PSObject.Properties | ForEach-Object { 
   if($_.value)
     { Add-Data -ReturnObject $temp -counter $indexer -Dataobject $line}
     else{ Add-Data -ReturnObject $temp -counter $indexer  -Dataobject $last  }
     $indexer++
   }
  $FixedCSV += $temp
  $last = $temp
 }
}

function SelectData-byClinic {
Param([Parameter(Mandatory=$TRUE)][Object]$DataObject, [Parameter(Mandatory=$TRUE)][string]$SearchName, [Parameter(Mandatory=$TRUE)][string]$Field)
$filtered = $DataObject | Where-Object {$_."$Field" -like $SearchName  } 
return $Filtered
}


#Stupid function
Function New-Webpage{
Param([Parameter(Mandatory=$TRUE)][String]$page, [Parameter(Mandatory=$false)][int]$vis)
    $ie = new-object -com InternetExplorer.Application 
    if($vis -eq 0){ $ie.visible = $true} 
    else{$ie.visible = $false} 
    $ie.navigate2($page)
return $ie
}


$Shell = New-Object -COM Shell.Application
$Shell.Windows()  ## Find the right one in the list

$ie = $Shell.Windows().Item(1)  ## Grab the window
$frames = $this1.Document.frames  
$fr.document.getElementsByTagName("tr")  # $fr = $frames[0]  whree the name = topfr  (so first frame that has the tables embedded)  tr is table row,  tr[2] is the last
.getElementsByTagName("td")  #reference the right one
.getElementsByTagName("a")
$a[0].click()
$tr = $ie.document.frames[0].getElementsByTagName("tr") ; $td = $tr[2].getElementsByTagName("td") ;$a =  $td[1].getElementsByTagName("a");
$a[0].click()  

#this is the 1 liner:
$idxhome.Document.frames[0].document.getElementsByTagName("tr")[2].getElementsByTagName("td")[1].getElementsByTagName("a")[0].click()
 $frames =$idxpage.Document.frames[1]
 $wfeapplet = $frames.document.getElementById("wfeapplet")

 $iepid = (Get-Process | Where-Object { $_.MainWindowHandle -eq $idxpage.Hwnd }).Id
 [Microsoft.VisualBasic.Interaction]::AppActivate($iepid);[System.Windows.Forms.SendKeys]::SendWait("$usname")

 #Emulate Keys to IDX website
 function Send-Keys {
 Param([Parameter(Mandatory=$false)][String]$key, [Parameter(Mandatory=$false)][int]$enter)
   AppActivate-Window($iepid)
   [System.Windows.Forms.SendKeys]::SendWait($key)
   if($enter -ne 0)
   {
    [System.Windows.Forms.SendKeys]::SendWait("~")
   }
 }
 #major function for IDX java website to emulate keystrokes
 Function AppActivate-Window {
 Param([Parameter(Mandatory=$TRUE)][int]$ID)
  [Microsoft.VisualBasic.Interaction]::AppActivate($ID);
 }
 $idxpage.Document.frames[0].document.getElementsByTagName("tr")[2].getElementsByTagName("td")[1].getElementsByTagName("a")[2].click()

 # Function gets IDX webpage and recaptures it after java breaks it
 Function IDX-webpage{
 $dummypage = New-Webpage $idx 
 $Shell = New-Object -COM Shell.Application
 $idxhome = $Shell.Windows() | Where-object { $_.LocationURL -like "http://idx/gpmsweb/"}  ## Grab the window
 }

 #simple Get cred function
Function Get-Cred
{
Param([Parameter(Mandatory=$False)][string]$Whatfor)
if($whatfor) {Write-Host "Enter credentials for: $whatfor"}
$Creds = @()
$User =  Read-Host 'Enter Username'
$pass =  Read-Host 'Enter Password'
$Creds = @([pscustomObject] @{
        User   = $User 
        Pass   = $Pass 
        })
Return $creds
}



#main OC med function
Function ocemed-page{
$emed_page =New-Webpage $ocemed
Write-Output "Get Emed Credentials"
$emed_creds = Get-Cred
Input-creds -Credentials $emed_creds -WebPage $emed_page
return $emed_page
}

# Function gets Elements on webpage by ID
Function Get-ElementbyID{
Param([Parameter(Mandatory=$TRUE)][string]$WebPage,[Parameter(Mandatory=$TRUE)][string]$ID)
Return $webpage.Document.getElementById($ID)
}

#Function Input's credentials into a website
Function Input-creds{
Param([Parameter(Mandatory=$TRUE)][pscustomObject]$Credentials,[Parameter(Mandatory=$TRUE)][string]$WebPage)
(Get-ElementbyID -WebPage $WebPage -ID $emed_web_use).value = $Credentials.user
(Get-ElementbyID -WebPage $WebPage -ID $emed_web_pass).value = $Credentials.pass
}

#function searches Emed website
Function Search-Emed {
Param([Parameter(Mandatory=$TRUE)][string]$WebPage,[Parameter(Mandatory=$TRUE)][string]$Search)
(Get-ElementbyID -WebPage $WebPage -ID $emed_web_SubID).value = $Search 
(Get-ElementbyID -WebPage $WebPage -ID $emed_web_Submit).click() 
}

#function searches IDX PM by patient account number
function Patient-Manager {
Param([Parameter(Mandatory=$TRUE)][int]$patientAccountNumber)
$idxhome.Document.frames[4].document.getElementsByName("account")[0].value = $patientAccountNumber
$idxhome.document.frames[4].document.getElementsByTagName("table")[1].getElementsByTagName("td")[27].getElementsByTagName("input")[0].click() 
}



Function Main-Main {
Write-Host "Welcome to the Referral Automation Application"
$again = 0
While($again -eq 0){
	Write-Host "What would you like to do?"
	Write-host "Options:
                Get Referral         (1)
                Login Centricity     (2)
                Login medEconnect    (3)
                xx List              (4)
                xx ist               (5)"
        $input =  Read-Host "Input Selection:"
        Switch ($input){
        1{  ;break}
        }
	}
}
Function Test-VarSet{
}
Function Test-Variable($vartest) {
$return = if($vartest){$true} else {$false}
Return $return 
}
Function Get-NetDrives{
$result = get-wmiobject win32_mappedLogicalDisk -computername $env:computername | select caption, providername
Return $esult
}



Function ExportWSToCSV ($input_File, $output_Location )
{
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    $wb = $E.Workbooks.Open($input_File)
    foreach ($ws in $wb.Worksheets)
    {
        $n = "Referral_Report" + "_" + $ws.Name
        $ws.SaveAs($output_Location + $n + ".csv", 6)
    }
    $E.Quit()
}


Try{
if((Test-Variable($ocemed)))
{Write-host '$ocemed is empty'

}}
Catch
{
$_
}





# Old unfinished code for parsing PDF
<#
$ItextPath = 'C:\Users\Pride\Desktop\Hickory Coding Project\itextsharp.dll'
# Add type for .DLL usage in PDF parsing
 Add-Type -Path $ItextPATH
 $list = (Get-ChildItem -Filter *.pdf)
 foreach($_ in $list)
 {
    $counter = 0
    $reader = New-Object iTextSharp.text.pdf.pdfreader -ArgumentList $_.name
    #Clear variables
    $floor = ""; $Location = "";$Damper_Num = ""; $Asset_Num = ""; $pass= ""; $Damper_type = "";$Date = ""; $date_sub = ""


     for ($page = 1; $page -le $reader.NumberOfPages; $page++)
    {
      # extract a page and split it into lines
     $text = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader,$page).Split([char]0x000A)  
     foreach ($line in $text)
    {
    $line
    # line formatting for easier string recognition
    #  $line = $line.ToUpper()
     # Checks each line for key words,  extracts contents past key words
    # if($line.contains("PASS")){$Pass= $line.Substring(4,($line.Length -4 ))  }
    # if($line.contains("FLOOR")) {$floor = $line.Substring(5,($line.Length -5 )) }
    # if($line.contains("ASSET#")) {$Asset_Num = $line.Substring(6,($line.Length -6 )) }
    # if($line.contains("DAMPER#")) {$Damper_Num =$line.Substring(7,($line.Length -7 )) }
    # if($line.contains("DAMPER_TYPE")-or $line.contains("DAMPER TYPE")) {$Damper_type = $line.Substring(11,($line.Length -11 )) }
    # if($line.contains("DATE / TIME")) {$Date = $line.Substring(12,($line.Length -12 )) }
    # if($line.contains("DAMPER_LOCATION") -OR $line.Contains("DAMPER LOCATION")) {$Location = $line.Substring(15,($line.Length -15 )) }
     # IF no date exists, gets date from submission
    # if($line -eq $text[$text.count - 1] -and $Date -eq "") {$date_sub = $line.Substring(0,12)}
     }
   
    }
}

#>
