#Purpose of this script is to rename the files in the Bank printout directory and in doing so organizing them. If the file has already been renamed then it should skip it

$CSVDirectory = 'C:\Users\Magnus\OneDrive - Aarhus universitet\Skrivebord\Calisthenics Aarhus Membership Adminstration\Bank Statement (Per Month)'
$csvfiles = Get-ChildItem -literalpath $CSVDirectory -Filter *.csv

foreach ($csvfile in $csvfiles) 
{ $CSV

    if($csvfile.Name.Contains("Foreningskonto 4399493582")) 
    {
        $int = $int + 1
        $csvdata = Import-Csv $csvfile.FullName -Delimiter ";"
        $LateDate = $csvdata | Select-Object -first 1
        $earlyDate = $csvdata | Select-Object -last 1
        
        $newFileName = "$($Earlydate.Bogføringsdato) to $($LateDate.Bogføringsdato).csv"
        Rename-Item -literalPath $csvfile.FullName -NewName $newFileName
    }
}

