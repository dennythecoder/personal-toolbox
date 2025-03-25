
$d_isDebug = $true
$d_debugLogPath = ""
function D-Debug{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        $text
    )

    begin {
        $ln = (Get-PSCallStack)[1].ScriptLineNumber
        $dt = (Get-Date).ToString()
    }

    process {
        If($d_isDebug){
            $msg = $text -join ", "
            $output = ($dt,"`t","Line: ",$ln,"`t",$msg)
            If($d_debugLogPath -eq ""){
                Write-Host $output
            }else {
                Add-Content -Path $debugLogPath -Value $output
            }
        }
    }

    end {
        # Cleanup code
    }
}