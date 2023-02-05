#####################################
# Activate Windows 10 AND MS Office #
#####################################

# Set Variables
$KMS_Host = "KMS-FQDN"
$Windows_Product_Key = "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"
$Office_2019_x86_Product_Key = "YYYYY-YYYYY-YYYYY-YYYYY-YYYYY"
$Office_2019_x64_Product_Key = "ZZZZZ-ZZZZZ-ZZZZZ-ZZZZZ-ZZZZZ"
$Office_2021_x86_Product_Key = "AAAAA-AAAAA-AAAAA-AAAAA-AAAAA"
$Office_2021_x64_Product_Key = "BBBBB-BBBBB-BBBBB-BBBBB-BBBBB"

# Activate Windows
slmgr.vbs -skms $KMS_Host
slmgr.vbs -ipk $Windows_Product_Key
slmgr.vbs -ato

# Determine the version and architecture of Office installed
$Office_Version = (Get-ItemProperty HKLM:\Software\Microsoft\Office\ClickToRun\Configuration).Client
if ($Office_Version -like "*64*") {
    $Office_Architecture = "x64"
}
else {
    $Office_Architecture = "x86"
}

# Set the path to the Office installation based on the version and architecture
if ($Office_Version -like "*19*") {
    if ($Office_Architecture -eq "x64") {
        $Office_Path = "C:\Program Files\Microsoft Office\Office19"
        $Office_Product_Key = $Office_2019_x64_Product_Key
    }
    else {
        $Office_Path = "C:\Program Files (x86)\Microsoft Office\Office19"
        $Office_Product_Key = $Office_2019_x86_Product_Key
    }
}
elseif ($Office_Version -like "*21*") {
    if ($Office_Architecture -eq "x64") {
        $Office_Path = "C:\Program Files\Microsoft Office\Office21"
        $Office_Product_Key = $Office_2021_x64_Product_Key
    }
    else {
        $Office_Path = "C:\Program Files (x86)\Microsoft Office\Office21"
        $Office_Product_Key = $Office_2021_x86_Product_Key
    }
}

# Activate MS Office
Set-Location $Office_Path
cscript ospp.vbs /sethst:$KMS_Host
cscript ospp.vbs /inpkey:$Office_Product_Key
cscript ospp.vbs /act
cscript ospp.vbs /dstatusall
