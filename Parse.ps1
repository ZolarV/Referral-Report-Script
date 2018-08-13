<#  Author: Michael Curtis
	Version:  .8
    Date  7/16/2018
    Purpose:  Using Daily Referral Report, check Primary Care Provider in centricity matches PRI in OCeMedconnect
              Checks for eligability and current date.  
              If !eligable or current date_month == this_month  then do nothing
              Else Match names and update Current Date field  with todays date.
              How it feels to chew 5 gum https://www.youtube.com/watch?v=CqaAs_3azSs  || http://www.infinitelooper.com/?v=CqaAs_3azSs&p=n
	Primary Logic:
		Step 1: Get .XLS from Directory Using BITS 
			A:  Test netpath and map Analyzer drive if need be
			B:  Create local location and download .XLS to local
		Step 2: Convert .XLS to .CSV  using  ExportWStoCSV function
			A:  Save export to Local folder
		Step 3: Fix CSV data by using  Format--Data Function   
			A:	Note: Colsolidate CSV done by calling Add-Data for pscustomobject.  Need to abstract a tad more, currently hardcoded to use $header[$index]
		Step 4: Filter fixedCSV data using user input for Field to search and searchterm  using SelectData-ByClinic   ###  Bad name rename
		Step 5: Login to IDX  
		Step 6: Login to oc.medEconnect.com
		Step 7: Using Patient account number from Filtered FixCSV  search IDX 
		Step 8: Go to Insurance tab
			A:  If Date (month) in Free Text Box is same as current month , SKIP
		Step 9: Get subscriber ID 
			A:	using Subscriber ID, Search oc.medeconnect.com 
		Step 10: Get Eligibility status, PRI, and PRI number from MedEConnect
		Step 11: Logic Check IF ELIGIBILE  CONTINUE  ELSE EXIT PATIENT
		Step 12: Match PRI from MedEConnect with PRI in IDX
		Step 13: Open IDX PRI Search Box, Search PRI name 
			A:	 Using PRI name and PRI number Filter through Search and Select Radio button
		Step 14: Update Checked date in Free Text field
			FORK: Build Patient Referral in IDX   (NEED TO DO)
		Step 15: NEXT PATIENT!
		

#>

# Constant defines
# Web Pages
$idx = "http://idx/gpmsweb"
$ocemed = "oc.medeconnect.com"
$emed_web_use = "MainContent_txtUsername"
$emed_web_pass = "MainContent_txtPassword"
$emed_web_login = "MainContent_btnLogonImage"
$emed_web_RTS = "https://oc.medeconnect.com/main/RealtimeSearch.aspx"
$emed_web_SubID = "tabContainer_tabSearch_txtCertificateNumber"
$emed_web_Submit = "tabContainer_tabSearch_btnHSubmit"

#HighScope Constants
$CurrentDate = Get-Date
$IDX_Nav = New-Object PSObject

# Referral_Report Location xls  on COS
$local_Ref_Rep_Dir = "$home\desktop\Local Referral Report\"
$ref_Rep_name = "Referral_Report.xls"
$rrl = "\\Analyzer\Analyzer\Snapshots\cos\Excel\daily\Referral_Report.xls"
$Full_header = "Location", "Count", "Provider", "Appt Date", "Enc", "Acct #", " ", "Patient", "Entered by", "Insurance", "Error"
$Trunc_header = "Location", "Count", "Provider", "Appt Date", "Enc", "Acct #", "Patient", "Entered by", "Insurance", "Error"

# .net Interfaces used in manipulating the embedded Java applet
[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")

# Creates Local Referral report folder and downloads current report
if (!(Test-Path -Path $local_Ref_Rep_Dir)) {
    New-Item -Path $home\desktop\ -Name "Local Referral Report" -ItemType dir
}

#simple Get cred function
Function Get-Cred {
    Param([Parameter(Mandatory = $False)][string]$Whatfor)
    if ($whatfor) {Write-Host "Enter credentials for: $whatfor"}
    $Creds = @()
    $User = Read-Host 'Enter Username'
    $pass = Read-Host 'Enter Password'
    $Creds = @([pscustomObject] @{
            User = $User
            Pass = $Pass
        })
    Return $creds
}

# $occreds = Get-Cred -Whatfor "OC Credentials"
$idxCreds = Get-Cred  -Whatfor "Centricity Credentials"
$emedCreds = Get-Cred  -Whatfor "oc.emedeconnect credentials"

# Consolidates and completes each CSV object
function Add-Data {
    Param(
        [Parameter(Mandatory = $true)][PSObject]$ReturnObject,
        [Parameter(Mandatory = $true)][int]$counter,
        [Parameter(Mandatory = $true)][Object]$Dataobject)
        if( $Trunc_header[$counter] -eq "Appt Date"){
            $ReturnObject | Add-Member -TypeName NoteProperty -NotePropertyName $Trunc_header[$counter] -NotePropertyValue [dateTime]::Parse($DataObject."$($Trunc_header[$counter])")
        }
        Else{
            $ReturnObject | Add-Member -typeName NoteProperty -NotePropertyName $Trunc_header[$counter] -NotePropertyValue $DataObject."$($Trunc_header[$counter])"
        }

}

#  dynamically  name a variable with a variable === $value=$NetworkInfo."$($_.Name)"
function Format-data {
    $FixedCSV = @()
    $shitCSV = (Import-Csv "$home\desktop\Local Referral Report\Referral_Report.csv" -Header $Full_header | Select-Object $Trunc_header | Select-Object -Skip 1)
    $counter = 0
    foreach ( $line in $shitCSV) {
        $temp = New-Object PSObject
        $indexer = 0
        $line.PSObject.Properties | ForEach-Object {
            if ($_.value)
            { Add-Data -ReturnObject $temp -counter $indexer -Dataobject $line}
            else { Add-Data -ReturnObject $temp -counter $indexer  -Dataobject $last  }
            $indexer++
        }
        $FixedCSV += $temp
        $last = $temp
        $counter ++
        #neat little progress bar for the functions that take longer
        Write-Progress -Activity "Building Data" -Status "Progress" -PercentComplete (($counter / $shitCSV.Count) * 100 )
    }
    Return $FixedCSV
}

# Returns selected data from the csv object.    Supply both field to search and the search term
function SelectData-byClinic {
    Param([Parameter(Mandatory = $TRUE)][Object]$DataObject, [Parameter(Mandatory = $TRUE)][ref]$SearchName, [Parameter(Mandatory = $TRUE)][string]$Field)
    if($Field -eq "Appt Date"){
        $filtered = $DataObject | Where-Object {$_."$Field" -ge $SearchName.Value[0]  }
        $filtered = $filtered | Where-Object {$_."$Field" -le $SearchName.Value[1]  }
    }
    else {
    $filtered = $DataObject | Where-Object {$_."$Field" -like $SearchName.Value  }}
    return $Filtered
}


# Gets medE variables and puts them into an object
Function GetMedE-Variables {
    Param([Parameter(Mandatory = $True)][String]$Benefits_Window)
    $Custom = @()
    $Eligibility_Status = (((($benfits_Window.Document.frames)[0].document.getElementsByTagName("title"))[0].document.getElementsByTagName("tr") | Where-Object {$_.outertext -like "Eligibility Status:*"}).outerText).Remove(0, 19)
    $Custom | Add-Member -type NoteProperty -Name "Eligibility_Status"  -Value $Eligibility_Status
    $PCP_Name = ((($benfits_Window.Document.frames)[0].document.getElementsByTagName("title"))[0].document.getElementsByTagName("tr") | Where-Object {$_.outertext -like "PCP Name*"}).outerText
    $PCP_Name = ($pcp_name.Remove(0, 44)).Replace("'" , "")
    $Custom | Add-Member -type NoteProperty -Name "PCP_Name"  -Value $PCP_Name
    $PCP_Phone_Number = ((($benfits_Window.Document.frames)[0].document.getElementsByTagName("title"))[0].document.getElementsByTagName("tr") |  Where-Object {$_.outertext -like "PCP Phone Number*"}).outerText
    $PCP_Phone_Number = ($PCP_Phone_Number.replace("PCP Phone Number PCP Phone Number is ", "")).replace("'", "")
    $Custom | Add-Member -type NoteProperty -Name "PCP_Number"  -Value $trimmed
    Return $Custom
}


# Search Window Function
Function Search-Window {
    Param([Parameter(Mandatory = $True)][String]$searchTerm, [Parameter(Mandatory = $true)][int]$number)
    $Shell = New-Object -COM Shell.Application
    $search_window =  Ret-shell -page "Group Management Constants Search"
    #Search Text  Name:  search  #value = value to search by
    (Get-Elementsby -webpage $search_window -Type "Name" -value ([ref]"Search")).value = $searchTerm
    #Execute Search
    ((Get-Elementsby  -webpage $search_window -Type "TagName" -value ([ref]"input" )) | Where-Object {$_.type -like "submit"}).click()
    #search each row result for stuff
    $chooser_term = (($search_window.Document.getElementsByTagName("tr") | Where-Object {$_.innertext -like "*" + $searchTerm + "*" -and $_.innertext -like "*" + $number + "*" }).outerText)[0].Remove(3)
    #Select Appropriate Radio button to execute
    ((Get-Elementsby -webpage $search_window -Type "Name" -value ([ref]"chooser"))| Where-Object {$_.value -eq $chooser_term + "/R"}).click()
}

# New Elements function
function Get-Elementsby {
    Param([Parameter(Mandatory = $True)][object]$webpage, [Parameter(Mandatory = $True)][string]$Type, [Parameter(Mandatory = $False)][ref]$value)
    $element = New-Object -ComObject  InternetExplorer.Application
    Switch ($type) {
        {"ID" -or 1} { $element = $webpage.document.getElementByID($value); break}
        {"Name" -or 2} {$element = $webpage.document.getElementsByName($value); break}
        {"TagName" -or 3} {$element = $webpage.document.getElementByTagName($value); break}
    }
    Return $element
}

# Emulate Keys to IDX website
function Send-Keys {
    Param([Parameter(Mandatory = $false)][String]$key, [Parameter(Mandatory = $false)][int]$enter)
    AppActivate-Window($iepid)
    [System.Windows.Forms.SendKeys]::SendWait($key)
    if ($enter -ne 0) {[System.Windows.Forms.SendKeys]::SendWait("~")}
    Start-Sleep 3
}

#major function for IDX java website to emulate keystrokes
Function AppActivate-Window {
    Param([Parameter(Mandatory = $TRUE)][int]$ID)
    [Microsoft.VisualBasic.Interaction]::AppActivate($ID);
}

# Function gets IDX webpage and recaptures it after java breaks it   ##OLD  Reworked into IDX-LOGIN
Function IDX-webpage {
    $dummypage = New-Webpage $idx
    $Shell = New-Object -COM Shell.Application
    $idxhome =  Ret-shell -page  "$idx"   ## Grab the window
    Return $idxhome
}

#main OC med function
Function ocemed-page {
    Param([Parameter(Mandatory = $TRUE)][object]$Credentials)
    $emed_page = New-Webpage $ocemed
    Input-creds -Credentials $Credentials
    return $emed_page
}

# Function gets Elements on webpage by ID  ######## OLD NEEDS CHANGED IN CODE  #########
Function Get-ElementbyID {
    Param([Parameter(Mandatory = $TRUE)][object]$WebPage, [Parameter(Mandatory = $TRUE)][string]$ID)
    Return ($webpage.Document.getElementById($ID))
}

#Function Input's credentials into a website
Function Input-creds {
    Param([Parameter(Mandatory = $TRUE)][object]$Credentials)
    $WebPage = Ret-shell -page "Eligibility: Login*"
    (Get-ElementbyID -WebPage $WebPage -ID $emed_web_use).value = "$Credentials.user"
    (Get-ElementbyID -WebPage $WebPage -ID $emed_web_pass).value = "$Credentials.pass"
    (Get-ElementbyID -WebPage $WebPage -ID "MainContent_btnLogon").click()
}

#function searches Emed website
Function Search-Emed {
    Param([Parameter(Mandatory = $TRUE)][string]$Search)
    $webpage = Ret-shell -page "Eligibility: Real Time Eligibility Search*"
    (Get-ElementbyID -WebPage $WebPage -ID $emed_web_SubID).value = "$Search"
    (Get-ElementbyID -WebPage $WebPage -ID $emed_web_Submit).click()
}

#function searches IDX PM by patient account number
function Patient-Manager {
    Param([Parameter(Mandatory = $TRUE)][int]$patientAccountNumber)
    $idxhome.Document.frames[4].document.getElementsByName("account")[0].value = $patientAccountNumber
    $idxhome.document.frames[4].document.getElementsByTagName("table")[1].getElementsByTagName("td")[27].getElementsByTagName("input")[0].click() 
}

#Stupid function
Function New-Webpage {
    Param([Parameter(Mandatory = $TRUE)][String]$page, [Parameter(Mandatory = $false)][int]$vis)
    $ie = new-object -com InternetExplorer.Application
    if ($vis -eq 0) { $ie.visible = $true}
    else {$ie.visible = $false}
    $ie.navigate2($page)
    while ($ie.busy) {Start-Sleep 3}
    return $ie
}

# Returns Homepage Nav function buttons.  instead of trying to get them a billion times.
Function IDX-NavButtons {
    Param([Parameter(Mandatory = $true)][object]$page)
    $Custom = New-Object psobject
    $homepage = $page.Document.frames[0].document.getElementsByTagName("a")
    $Custom | Add-Member -type NoteProperty -Name "HomePage"  -Value  $homepage
    $Custom | Add-Member -type NoteProperty -Name "Login"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:wait_for_wfe(0);"})
    $Custom | Add-Member -type NoteProperty -Name "PM"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:top.open_patmgr();"})
    $Custom | Add-Member -type NoteProperty -Name "Back"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:do_back();"})
    $Custom | Add-Member -type NoteProperty -Name "Home"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:top.open_home();"})
    $Custom | Add-Member -type NoteProperty -Name "wfe"  -Value  ($homepage | Where-Object {$_.href -eq "javascript:top.open_wfe();"})
    Return $Custom
}


#Gets the .xls and converts to .csv
Function ExportWSToCSV ($RR_FileName_NO_extention, $Output_Folder_Location ) {
    $myfile = $Output_Folder_Location + $RR_FileName_NO_extention + ".xls"
    $E = New-Object -ComObject Excel.Application
    $wb = $E.Workbooks.Open($myfile)
    $wb.Worksheets.Item(1)._SaveAs($Output_Folder_Location + $RR_FileName_NO_extention + ".csv", 6)
    $E.Quit()
}

# Main IDX Patient Logic.  Need to rebuild better.
Function Login-IDX {
    Param([Parameter(Mandatory = $true)][String]$Practice)
    $idxhome = New-Object -com InternetExplorer.Application
    $dummypage = New-Webpage $idx
    $idxhome =  Ret-shell -page  "$idx"
    $iepid = (Get-Process | Where-Object { $_.MainWindowHandle -eq $idxhome.Hwnd }).Id
    AppActivate-Window -ID $iepid
    $IDX_Nav = IDX-NavButtons
    $IDX_Nav.login.click()
    Send-Keys $idxCreds.user -enter 1
    Send-Keys $idxCreds.Pass -enter 1
    Send-Keys -enter 1
    Send-Keys $practice -enter 1
    return $idxhome
}

function Ret-Shell{
    Param([Parameter(Mandatory=$false)][String]$Page,[Parameter(Mandatory=$false)][String]$URL)
    IF($page){
    $Retobj = ($Shell.windows() | Where-Object {$_.LocationName -like "$Page" })
    }
    if($URL) {
    $Retobj = ($Shell.windows() | Where-Object {$_.LocationURL -like "$URL" })   
    }
    Return $Retobj
}

Write-Host "Welcome to the Referral Automation Application"
    $again = 0
    While ($again -eq 0) {
        Write-Host "What would you like to do?"
        Write-host "Options:
                Get Referral_Report       (1)
                Run live Automation       (2)
                Run test Automation       (3)
                Exit                      (5)"
        $input = Read-Host "Input Selection:"
        Switch ($input) {
            1 {
                $again =0
                ############ Possibly test netpath first? ###########
                 Write-Host "Getting Referral Report and staging to $local_Ref_Rep_Dir"  # Starts Referral Report process,  Will rename old to _old If exist
                 Get-ChildItem -path "$home\Desktop\Local Referral Report\*" -include *.pdf , *.xls , *.csv | Rename-Item -NewName {$_.name -replace $_.basename , ($_.basename +"_old") }
                # Starts Bits Transfer of file
                [bool]$Analyzer_ISmapped = $false
                $a = net use
                $shell = New-Object -ComObject Shell.Application
                $Shell.Explore("\\Analyzer\")
                While(!($Analyzer_ISmapped)){
                    Foreach($line in $a ){
                        $Analyzer_ISmapped = $line.contains("Analyzer")
                        if($Analyzer_ISmapped){break}
                    }
                    $Shell.Explore("\\Analyzer\")
                    $a = net use
                }
                 Start-BitsTransfer -Source $rrl -Destination $local_Ref_Rep_Dir
                # Converts .xls to usable CSV
                ExportWStoCSV -RR_FileName_NO_extention "Referral_Report" -Output_Folder_Location $local_Ref_Rep_Dir
                 ; break}
            2 {
                $again =0
                # Build import CSV into pscustom and reformat data
                 #Fixes the CSV 
                $mydata =  Format-data # initializes unfiltered data 
                # Logs into Centricity and captures Webpage Nav buttons in GLOBAL variable $IDX_NAV
                $idxhome = New-Object -com InternetExplorer.Application
                $practice = Read-Host -Prompt "Please Enter Practice. E.G: COS, TST ..."  #TODO  Add logic to differentiate  in Login-IDX
                $idxhome = Login-IDX -Practice $practice
                #open PM
                $IDX_Nav = IDX-NavButtons -page (Ret-Shell -URL "$idx")
                $IDX_Nav.PM.click()
                #build filter variable.   as PS custom object.  multimember use each memeber as a method for itterating through and filtering sequentially
                $Field_Array = @()
                $Want_To_Filter = 0
                Write-Host "Select each filter you wish to apply"
                Write-Host "Note: Fields are filtered by the FIFO method.  First in first out"
                Write-Host "Available fields:
                            Location        Count
                            Provider        Appt Date
                            Enc             Acct #
                            Patient         Entered by
                            Insurance       Error"
                Write-host "E.G: Location Provider Date "
                Write-host "BY FIFO Location is filtered first, then Provider off of that data"
                While($Want_To_Filter -eq 0){
                    $Field_Array +=   Read-Host -Prompt "Enter Field exactly as show above"
                    $Want_To_Filter = Read-Host -Prompt "Do you want to add another Filter?  0 = yes 1 = no"
                 }
                foreach($filter in $Field_Array){
                $searchterm = Read-Host "Enter search term for field: $filter "
                if($filter -eq "Appt Date") {
                    Write-Host "Format for Date mm/dd/yyyy  E.G: 8/17/2018"
                    $searchterm = @()
                    $searchterm += ([datetime]::Parse((Read-Host "Enter Start Date for Field: $filter ")))
                    $searchterm += ([datetime]::Parse((Read-Host "Enter End Date for Field: $filter ")))
                 }
                $mydata = SelectData-byClinic -DataObject $mydata -SearchName $searchterm -Field $filter   #Fetches the Data we want to work on
                }
                # Now lets use the data selected in $mydata
                # Login to oc.medEconnect.com
                $oc_medE_page = ocemed-page -Credentials $emedCreds
                Foreach ($This_Data in $mydata) {
                    Patient-Manager -patientAccountNumber ($This_Data.'Acct #')             #Opens Patient Record
                    $idxhome.Document.frames[4].document.getElementById("radioIns").click() #open Primary Insurance
                    ($idxhome.Document.frames[4].document.getElementsByTagName("a")| Where-Object {$_.href -like "javascript:openwindow_idet(1,'P');"} ).click()  #gets Popup window
                    $insurance_Window = Ret-shell -page "Group Management Edit Insurance Detail" # Not sure if needed yet
                    $iwinpid = (Get-Process | Where-Object { $_.MainWindowHandle -eq $insurance_Window.Hwnd }).Id
                    #Get Policy  For medEconnect#
                    $policy_Value = $insurance_Window.Document.getElementsByName("policy1")[0].value   # should go to emed here and get PRI provider and Eligebility
                    # PCP name    ID = "pcpname1"   .. search by ID
                    $pcp_name = (Get-Elementsby -webpage $insurance_Window -Type "ID"  -value ([ref] "pcpname1")).textContent
                    $pcp_value = (Get-Elementsby -webpage $insurance_Window -Type "name"  -value ([ref] "pcp1")).value

                    # Free Text
                    # Name = ft111  Free Test 1     Search By Name
                    # Name = ft121  Free Test 2     Search By Name
                    # Name = ft131  Free Test 3     Search By Name
                    $Free_Text1 = Get-Elementsby -webpage $insurance_Window -Type "Name"  -value ([ref] "ftl11")
                    $Free_Text2 = Get-Elementsby -webpage $insurance_Window -Type "Name"  -value ([ref] "ftl21")
                    $Free_Text3 = Get-Elementsby -webpage $insurance_Window -Type "Name"  -value ([ref] "ftl31")
                    #Search Emed for $policy, Return Eligebility, PRI and PRI number
                    Search-Emed  -Search $policy_Value
                    $Search_Emed_page = Ret-Shell "Eligibility: Real Time Eligibility Search"
                    $Emed_Check_Vars = GetMedE-Variables -Benefits_Window $Search_Emed_page
                    #Do logic Here with EMED Variables
                    if ($Emed_Check_Vars.Eligibility_Status.value -like "Eligible") {
                        if($pcp_name -notlike $Emed_Check_Vars.PCP_Name.value){
                            Search-Window -SearchTerm $Emed_Check_Vars.PCP_Name.value -Number $Emed_Check_Vars.PCP_Number.value

                        }
                        if (!($Free_Text3.value)) {$Free_Text3.value = $CurrentDate.toShortDateString
                            
                        }
                        
                    }
                    if((Read-host "Do you want to continue to next Patient Record? (y,n): ") -like "y"){}
                    else {break}
                ;break}
        }
    }

}



$mydata = SelectData-byClinic -DataObject $fixedCSV -SearchName (Get-Input) -Field (Get-Input)  #Fetches the Data we want to work on
#loops through the data to do stuff here
Foreach ($This_Data in $mydata) {

    Patient-Manager -patientAccountNumber ($This_Data.'Acct #')  # figure out a way to filter data to only things that need done.  Currently only filtering data on one type
    <# PseudoDo
		  1: Do you want to filter by days?  
					A:Maxday = today
					B:MinDay = at least 1
					C:data using mydata from Selected data.  
		2:  Pop in patient number in Patiend-Manager,  Get Subscriber Number for Search-Med


		#>


    Write-Host "Would you like to Continue to the next patient?"
    $continue = (Read-Host "Input (y)es or (n)o:").tolower()
    If ($continue -eq "y") {
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

# Opens PRI Care prov Search Window
($insurance_Window.Document.getElementsByTagName("a")| Where-Object {$_.href -like "javascript:search(theform.pcp1,'pcpx');"} ).click() #Opens New Window  wsearch.cgi

 



    

Function Test-VarSet {
}
Function Test-Variable($vartest) {
    $return = if ($vartest) {$true} else {$false}
    Return $return
}
Function Get-NetDrives {
    $result = get-wmiobject win32_mappedLogicalDisk -computername $env:computername | Select-Object caption, providername
    Return $result
}



<#
For later use in abstracting Header information

$myheader = (Import-Csv -Path $directory +"\" +"Referral_Report.CSV")[0] | Get-member | Where-Object {$_.memberType -like "NoteProperty"} |Select-Object Name

$FixedCSV | %{$_.'Appt Date'} | Get-Unique 

PS C:\Users\Pride> $FixedCSV |  %{([datetime]::Parse($_.'Appt Date'))} | Get-Unique 

#>


Try {
    if ((Test-Variable($ocemed))) {
        Write-host '$ocemed is empty'

    }
}
Catch {
    $_
}





<# Old unfinished code for parsing PDF

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
