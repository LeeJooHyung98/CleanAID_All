
$path = "c:\WorkSpace\CleadAid\Source\ASIS_Sources\가맹점프로그램\Form\frm세탁물인도문자1.frm"
$reading = $false
Get-Content $path -Encoding Default | ForEach-Object {
    if ($_ -match "Private Sub cmdSend_Click") {
        $reading = $true
        Write-Output $_
    } elseif ($reading) {
        Write-Output $_
        if ($_ -match "End Sub") {
            $reading = $false
        }
    }
}
