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
$Full_header = "Location","Count","Provider","Appt Date","Enc","Acct #"," ","Patient","Entered by","Insurance", "Error"
$Trunc_header  = "Location","Count","Provider","Appt Date","Enc","Acct #","Patient","Entered by","Insurance", "Error"
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
ExportWStoCSV -RR_FileName_NO_extention "Referral_Report" -Output_Folder_Location "$home\desktop\Local Referral Report\"
# Consolidates and completes each CSV object
function Add-Data{
Param(
[Parameter(Mandatory=$true)][PSObject]$ReturnObject, 
[Parameter(Mandatory=$true)][int]$counter,
[Parameter(Mandatory=$true)][Object]$Dataobject)
$ReturnObject | Add-Member -type NoteProperty -Name $Trunc_header[$counter] -Value $DataObject."$($Trunc_header[$counter])"
}

#  dynamically  name a variable with a variable === $value=$NetworkInfo."$($_.Name)"
function Format-the-fucking-data {

$shitCSV = (Import-Csv "$home\desktop\Local Referral Report\Referral_Report.csv" -Header $Full_header | Select-Object $Trunc_header | Select-Object -Skip 1)
$counter = 0
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
  $counter ++ 
  Write-Progress -Activity "Building Data" -Status "Progress" -PercentComplete (($counter / $shitCSV.Count)*100 )
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
	while($ie.busy) {sleep 3} 
return $ie
}

Function IDX-NavButtons{
	$Custom = New-Object psobject
	$homepage  =  $idxhome.Document.frames[0].document.getElementsByTagName("a")
	$Custom | Add-Member -type NoteProperty -Name "HomePage"  -Value  $homepage
	$Custom | Add-Member -type NoteProperty -Name "Login"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:wait_for_wfe(0);"})
	$Custom | Add-Member -type NoteProperty -Name "PM"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:top.open_patmgr();"})
	$Custom | Add-Member -type NoteProperty -Name "Back"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:do_back();"})
	$Custom | Add-Member -type NoteProperty -Name "Home"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:top.open_home();"})
	$Custom | Add-Member -type NoteProperty -Name "wfe"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:top.open_wfe();"})
	Return $Custom
}

Function Login-IDX {
$idxhome = IDX-webpage 
$iepid = (Get-Process | Where-Object { $_.MainWindowHandle -eq $idxhome.Hwnd }).Id
AppActivate-Window -ID $iepid
$IDX_Nav = IDX-NavButtons
$IDX_Nav.login.click()

Send-Keys $idxCreds.user -enter 1 
Send-Keys $idxCreds.Pass -enter 1 
Send-Keys -enter 1
Send-Keys "tst" -enter 1  # Replace TST with $practice 

 #open PM
 $IDX_Nav.PM.click()

 Format-the-fucking-data  #Fixes the CSV
 $mydata = SelectData-byClinic -DataObject $fixedCSV -SearchName (Get-Input) -Field (Get-Input)  #Fetches the Data we want to work on
	#loops through the data to do stuff here
	Foreach($This_Data in $mydata){

		Patient-Manager -patientAccountNumber ($This_Data.'Acct #')


		 Write-Host "Would you like to Continue to the next patient?"
		 $continue = (Read-Host "Input (y)es or (n)o:").tolower()
         If($continue -eq "y"){
			  $IDX_Nav.Back.click()
		 }
		 Else {break} 
	}




 #use Patient-Manager to search by Acct number
 Patient-Manager -patientAccountNumber $Referral_Report_Patient_Account_Number
 #open Insurance tab
 $idxhome.Document.frames[4].document.getElementById("radioIns").click()
 #open Primary Insurance
 ($idxhome.Document.frames[4].document.getElementsByTagName("a")| Where-Object {$_.href -like "javascript:openwindow_idet(1,'P');"} ).click()
 #gets Popup window
 $insurance_Window = ($Shell.Windows() | Where-Object { $_.LocationName -like "Group Management Edit Insurance Detail"})
 # Not sure if needed yet
 $iwinpid = (Get-Process | Where-Object { $_.MainWindowHandle -eq $insurance_Window.Hwnd }).Id
 #Get Policy #
 $policy_Value = $insurance_Window.Document.getElementsByName("policy1")[0].value   # should go to emed here and get PRI provider and Eligebility


 #Pri Care Provider
 # Name = pcp1     Search By Name
 # Search Box Href = javascript:search(theform.pcp1,'pcpx'); Tag = a   Seach by Tag

 # Opens Search Window 
 ($insurance_Window.Document.getElementsByTagName("a")| Where-Object {$_.href -like "javascript:search(theform.pcp1,'pcpx');"} ).click() #Opens New Window  wsearch.cgi

 



 # PCP name    ID = "pcpname1"   .. search by ID
	$pcp_name  = (Get-Elementsby -webpage $insurance_Window -Type "ID"  -value ([ref] "pcpname1")).textContent
	$pcp_value = (Get-Elementsby -webpage $insurance_Window -Type "name"  -value ([ref] "pcp1")).value

 # Free Text
 # Name = ft111  Free Test 1     Search By Name
 # Name = ft121  Free Test 2     Search By Name
 # Name = ft131  Free Test 3     Search By Name
	$Free_Text1  = Get-Elementsby -webpage $insurance_Window -Type "Name"  -value ([ref] "ftl11") 
	$Free_Text2  = Get-Elementsby -webpage $insurance_Window -Type "Name"  -value ([ref] "ftl21") 
	$Free_Text3  = Get-Elementsby -webpage $insurance_Window -Type "Name"  -value ([ref] "ftl31") 
}

# Gets medE variables and puts them into an object
Function GetMedE-Variables{
	$Custom = @()
	$Eligibility_Status = (((($benfits_Window.Document.frames)[0].document.getElementsByTagName("title"))[0].document.getElementsByTagName("tr") | Where-Object {$_.outertext -like "Eligibility Status:*"}).outerText).Remove(0,19)
	$Custom | Add-Member -type NoteProperty -Name "Eligibility_Status"  -Value $Eligibility_Status
	$PCP_Name = ((($benfits_Window.Document.frames)[0].document.getElementsByTagName("title"))[0].document.getElementsByTagName("tr") | Where-Object {$_.outertext -like "PCP Name*"}).outerText
	$PCP_Name = ($pcp_name.Remove(0,44)).Replace("'" ,"")
	$Custom | Add-Member -type NoteProperty -Name "PCP_Name"  -Value $PCP_Name
	$PCP_Phone_Number = ((($benfits_Window.Document.frames)[0].document.getElementsByTagName("title"))[0].document.getElementsByTagName("tr") |  Where-Object {$_.outertext -like "PCP Phone Number*"}).outerText
	$PCP_Phone_Number = ($PCP_Phone_Number.replace("PCP Phone Number PCP Phone Number is ","")).replace("'","")
	$Custom | Add-Member -type NoteProperty -Name "PCP_Number"  -Value $trimmed
	Return $Custom
}


# Search Window Function
Function Search-Window {
 Param([Parameter(Mandatory=$True)][String]$searchTerm, [Parameter(Mandatory=$true)][int]$number)
	 $search_window =($Shell.Windows() | Where-Object { $_.LocationName -like "Group Management Constants Search"})
	 #Search Text  Name:  search  #value = value to search by
	 (Get-Elementsby -webpage $search_window -Type "Name" -value ([ref]"Search")).value = $searchTerm
	 #Execute Search
	 ((Get-Elementsby  -webpage $search_window -Type "TagName" -value ([ref]"input" )) | Where-Object{$_.type -like "submit"}).click()
	 #search each row result for stuff
	 $chooser_term = (($search_window.Document.getElementsByTagName("tr") | Where-Object {$_.innertext -like "*"+$searchTerm+"*" -and $_.innertext -like "*"+$number+"*" }).outerText)[0].Remove(3)
	 #Select Appropriate Radio button to execute
	((Get-Elementsby -webpage $search_window -Type "Name" -value ([ref]"chooser"))| Where-Object{$_.value -eq $chooser_term+"/R"}).click()
}
function Get-Elementsby {
	Param([Parameter(Mandatory=$True)][object]$webpage, [Parameter(Mandatory=$True)][string]$Type, [Parameter(Mandatory=$False)][ref]$value)
	$element = New-Object -ComObject  InternetExplorer.Application 
	Switch($type){
		{"ID" -or 1}{ $element = $webpage.document.getElementByID($value);break}
		{"Name" -or 2}{$element = $webpage.document.getElementsByName($value);break}
		{"TagName"-or 3}{$element = $webpage.document.getElementByTagName($value);break}
	}
	Return $element
}

 #Emulate Keys to IDX website
 function Send-Keys {
 Param([Parameter(Mandatory=$false)][String]$key, [Parameter(Mandatory=$false)][int]$enter)
   AppActivate-Window($iepid)
   [System.Windows.Forms.SendKeys]::SendWait($key)
   if($enter -ne 0) {[System.Windows.Forms.SendKeys]::SendWait("~")}
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
 Return $idxhome
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
Param([Parameter(Mandatory=$TRUE)][object]$Credentials)
$emed_page =New-Webpage $ocemed
Input-creds -Credentials $Credentials -WebPage $emed_page
return $emed_page
}

# Function gets Elements on webpage by ID  ######## OLD NEEDS CHANGED IN CODE  #########
Function Get-ElementbyID{
Param([Parameter(Mandatory=$TRUE)][object]$WebPage,[Parameter(Mandatory=$TRUE)][string]$ID)
Return ($webpage.Document.getElementById($ID))
}

#Function Input's credentials into a website
Function Input-creds{
Param([Parameter(Mandatory=$TRUE)][object]$Credentials,[Parameter(Mandatory=$TRUE)][object]$WebPage)
(Get-ElementbyID -WebPage $WebPage -ID $emed_web_use).value = $Credentials.user
(Get-ElementbyID -WebPage $WebPage -ID $emed_web_pass).value = $Credentials.pass
}

#function searches Emed website
Function Search-Emed {
Param([Parameter(Mandatory=$TRUE)][object]$WebPage,[Parameter(Mandatory=$TRUE)][string]$Search)
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
                Enter Your Practice  (4)
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
Return $result
}


#Gets the .xls and converts to .csv 
Function ExportWSToCSV ($RR_FileName_NO_extention, $Output_Folder_Location )
{
	$myfile = $Output_Folder_Location + $RR_FileName_NO_extention + ".xls"
    $E = New-Object -ComObject Excel.Application
    $wb = $E.Workbooks.Open($myfile)
	$wb.Worksheets.Item(1)._SaveAs($Output_Folder_Location + $RR_FileName_NO_extention +".csv", 6)
    $E.Quit()
}

<#
For later use in abstracting Header information

$myheader = (Import-Csv -Path $directory +"\" +"Referral_Report.CSV")[0] | Get-member | Where-Object {$_.memberType -like "NoteProperty"} |Select-Object Name

#>


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
