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
       }
    }
}

