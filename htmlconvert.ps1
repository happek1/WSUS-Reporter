Cd .\reports\

#Set up web styles for html reports
$a = "<style>"
$a = $a + "BODY{Font-family: Arial; font-size 10pt;background-color:White;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:White}"
$a = $a + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:White}"
$a = $a + "</style>"

$date = Get-Date
$d = $date.ToShortDateString()
$t = $date.ToShortTimeString()

Add-Type -AssemblyName System.Xml.Linq
$files = Get-ChildItem "" -Filter *.csv

ForEach ($file in $files){
$Process = $(Get-Content $file | ConvertFrom-Csv)

$xml = [System.Xml.Linq.XDocument]::Parse( "$($Process | ConvertTo-Html -Head $a -Body "<H2>WSUS Report created on $d  $T</H2>")" )
if($Namespace = $xml.Root.Attribute("xmlns").Value) {
    $Namespace = "{{{0}}}" -f $Namespace
}

# Find the index of the column you want to format:
$NeededCountIndex = [Array]::IndexOf( $xml.Descendants("${Namespace}th").Value, "NeededCount")

foreach($row in $xml.Descendants("${Namespace}tr")){
    switch(@($row.Descendants("${Namespace}td"))[$NeededCountIndex]) {
       {100 -le $_.Value } { $_.SetAttributeValue( "style", "background: red;"); continue } 
       {50  -le $_.Value } { $_.SetAttributeValue( "style", "background: orange;"); continue } 
       {10  -le $_.Value } { $_.SetAttributeValue( "style", "background: yellow;"); continue }
       
    }
}
# Save the html out to a file
$xml.Save("$pwd/$file.html")
}

#Removes the CSV from the HTML file names
dir *.csv.html | rename-item -newname { $_.name -replace ".csv", "" }
Cd..
