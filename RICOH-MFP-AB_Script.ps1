#============================================================================================================================#
#                                                                                                                            #
#  RICOH-MFP-AB_Script.ps1                                                                                                   #
#  Ricoh Multi Function Printer (MFP) Address Book PowerShell Script                                                         #
#  Author: Jamal A. Mohamed                                                                                                  #
#  Creation Date: 10.27.2020                                                                                                 #                                                                                                 #
#                                                                                                                            #
#                                                                                                                            #
#============================================================================================================================#



$csv = Import-Csv -Path "path to a csv containing printer ips.header should be IP"
#import the module, assuming it is in the same directory
import-module ".\RICOH-MFP-AB.psm1"

#resolves some cert issues.
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy


#loop through the printers

foreach($printer in $csv) {
#create a dictionary to put the output of the script in.
$addressbook = @()
#get the IP from the csv, we are expecting the header of the IP row to be IP
$ip = $printer.ip

#Check that the IP address is Live and it's not a blank line.
if((Test-Connection -Count 1 -ComputerName $ip -Quiet ) -and ($ip -ne ''))  {
try {
#Give the username and password of the ricoh devices, if it wasn't change leave as is.
$addressbook = (Get-MFPAB -Hostname $ip  -Username 'admin' -Password '')

$addressbook | Export-csv  -Path ".\$Ip.csv"  -NoTypeInformation 
}
catch {
write-host "couldn't retrieve addressbook from $IP, please check manually"
continue

}
Else {
write-host "This IP $ip is not reachable"

}

}


}

