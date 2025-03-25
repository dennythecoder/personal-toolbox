# Usage

$path = "c:\temp\workbook.xlsx"
$content = Grab-ExcelContent $path


Add-Type -Assembly "System.IO.Compression.FileSystem"

function Grab-ExcelContent($path){
    function Grab-Col($range){
        $col = ($range -split '[0-9]')[0]
        return $col
    }

    function Grab-Row($range){
        $col = ($range -split '[0-9]')[0]
        $row = $range -replace $col,""
        return $row 
    }

    function Grab-Xml($file){
        $stream = $file.Open() 
        $reader = New-Object IO.StreamReader($stream)
        $text = $reader.ReadToEnd()
        $content = [xml]$text
        $stream.Close()
        $reader.Close()
        return $content
    }

    $base_date = (Get-Date "1/1/1900").AddDays(-2)
    function Grab-CellValue($cell){
       
        $v = $cell.v
        If($cell.t -eq "s"){
            $ss = $sharedStrings[$v].t
            If($ss."#text" -eq $null){
                return $ss
            } else {
                $ss."#text"
            } 
        } else {
            If($cell.s -ne $null){
                $xf = $xfs[$cell.s]
                $nmid = [int] $xf.numfmtid
                if($nmid -ge 14 -and $nmid -le 22){
                    return $base_date.AddDays($v)
                } else {
                    return $v
                }   
            } else { 
                return $v
            }
        }
    }
    try{
        
        $zip = [IO.Compression.ZipFile]::OpenRead($path)
        $sharedStringsFile = $zip.Entries | where FullName -EQ "xl/sharedStrings.xml" 
        $sharedStrings = (Grab-Xml $sharedStringsFile).sst.si
        $worksheets = $zip.Entries | where FullName -Like "xl/worksheets/*" 
        $workbook = Grab-Xml ($zip.Entries | Where FullName -EQ "xl/workbook.xml")
        $styles = Grab-Xml ($zip.Entries | Where FullName -EQ "xl/styles.xml")
        $xfs = $styles.stylesheet.cellXfs.xf
        
    } catch{
        return [pscustomobject]@{"error"="File is likely in use or corrupted"}
    }
    $wsi = 0
    $content = $worksheets | foreach {
        $sheetData = (Grab-Xml $_).worksheet.sheetData
        $headers = $sheetData.row[0].c | ForEach {      
            $v = Grab-CellValue $_ 
            $c = Grab-Col $_.r
            @{$c=$v}
        }
        $rix = 0 
        $sheet = $sheetData.row | foreach {
            if($rix -eq 0) { $rix++ } else{
                $row = [PSCustomObject]@{}
                $_.c | foreach {
                    $col = Grab-Col $_.r
                    $h = $headers.$col
                    $v = Grab-CellValue $_
                    $row | Add-Member -Name $h -Type NoteProperty -Value $v
                }
                Write-Output $row
            }
        }

        if($worksheets.count -eq 1) {
            $name = $workbook.workbook.sheets.sheet.name
        } else {
            $name = $workbook.workbook.sheets.sheet[$wsi].name
            $wsi++
        } 
        Write-Output ([pscustomobject]@{$name=$sheet})
    }
    $zip.Dispose() 
    return $content
}