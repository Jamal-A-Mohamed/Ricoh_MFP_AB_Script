#============================================================================================================================#
#                                                                                                                            #
#  RICOH-MFP-AB.psm1                                                                                                         #
#  Ricoh Multi Function Printer (MFP) Address Book PowerShell Module                                                         #
#  Author: Alexander Krause                                                                                                  #
#  Creation Date: 10.04.2013                                                                                                 #
#  Modified Date: 17.04.2013                                                                                                 #
#  Version: 0.7.7                                                                                                            #
#                                                                                                                            #
#============================================================================================================================#

function ConvertTo-Base64
{
param($String)
[System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($String))
}

function Connect-MFP
{
param($Hostname,$Authentication,$Username,$Password,$SecurePassword)
$url = "http://$Hostname/DH/udirectory"
$login = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:startSession xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <stringIn></stringIn>
   <timeLimit>30</timeLimit>
   <lockMode>X</lockMode>
  </m:startSession>
 </s:Body>
</s:Envelope>
'@
if($SecurePassword -eq $NULL){$pass = ConvertTo-Base64 $Password}else{$pass = $SecurePassword; $enc = "gwpwes003"}
$login.Envelope.Body.startSession.stringIn = "SCHEME="+(ConvertTo-Base64 $Authentication)+";UID:UserName="+(ConvertTo-Base64 $Username)+";PWD:Password=$pass;PES:Encoding=$enc"
[xml]$xml = iwr $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#startSession"} -Body $login
if($xml.Envelope.Body.startSessionResponse.returnValue -eq "OK"){$script:session = $xml.Envelope.Body.startSessionResponse.stringOut}
}

function Search-MFPAB
{
param($Hostname)
$url = "http://$Hostname/DH/udirectory"
$search = [xml]@'
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:searchObjects xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
    <sessionId></sessionId>
   <selectProps xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" soap-enc:arrayType="itt:string[1]">
    <item>id</item>
   </selectProps>
    <fromClass>entry</fromClass>
    <parentObjectId></parentObjectId>
    <resultSetId></resultSetId>
   <whereAnd xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" soap-enc:arrayType="itt:queryTerm[1]">
    <item>
     <operator></operator>
     <propName>all</propName>
     <propVal></propVal>
     <propVal2></propVal2>
    </item>
   </whereAnd>
   <whereOr xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" soap-enc:arrayType="itt:queryTerm[1]">
    <item>
     <operator></operator>
     <propName></propName>
     <propVal></propVal>
     <propVal2></propVal2>
    </item>
   </whereOr>
   <orderBy xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/udirectory" soap-enc:arrayType="itt:queryOrderBy[1]">
    <item>
     <propName></propName>
     <isDescending>false</isDescending>
    </item>
   </orderBy>
    <rowOffset>0</rowOffset>
    <rowCount>50</rowCount>
    <lastObjectId></lastObjectId>
   <queryOptions xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" soap-enc:arrayType="itt:property[1]">
    <item>
     <propName></propName>
     <propVal></propVal>
    </item>
   </queryOptions>
  </m:searchObjects>
 </s:Body>
</s:Envelope>
'@
$search.Envelope.Body.searchObjects.sessionId = $script:session
[xml]$xml = iwr $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#searchObjects"} -Body $search
$xml.SelectNodes("//rowList/item") | %{$_.item.propVal} | ?{$_.length -lt "10"} | %{[int]$_} | sort
}

function Get-MFPAB
{
param($Hostname,$Authentication="BASIC",$Username="admin",$Password,$SecurePassword)
Connect-MFP $Hostname $Authentication $Username $Password $SecurePassword
$url = "http://$Hostname/DH/udirectory"
$get = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:getObjectsProps xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <sessionId></sessionId>
  <objectIdList xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="">
  </objectIdList>
  <selectProps xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:string[7]">
   <item>entryType</item>
   <item>id</item>
   <item>index</item>
   <item>name</item>
   <item>longName</item>
   <item>auth:name</item>
   <item>mail:address</item>
   <item>fax:number</item>                                            
  </selectProps>
   <options xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:property[1]">
    <item>
     <propName></propName>
     <propVal></propVal>
    </item>
   </options>
  </m:getObjectsProps>
 </s:Body>
</s:Envelope>
'@
$get.Envelope.Body.getObjectsProps.sessionId = $script:session
Search-MFPAB $Hostname | %{
$x = $get.CreateElement("item")
$x.set_InnerText("entry:$_")
$o = $get.Envelope.Body.getObjectsProps.objectIdList.AppendChild($x)
}
$get.Envelope.Body.getObjectsProps.objectIdList.arrayType = "itt:string["+$get.Envelope.Body.getObjectsProps.objectIdList.item.count+"]"
[xml]$xml = iwr $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#getObjectsProps"} -Body $get
$xml.SelectNodes("//returnValue/item") | %{
New-Object PSObject -Property @{
   EntryType = (%{$_.item} | ?{$_.propName -eq "entryType"}).propVal
   ID        = [int](%{$_.item} | ?{$_.propName -eq "id"}).propVal
   Index     = [int](%{$_.item} | ?{$_.propName -eq "index"}).propVal
   Name      = (%{$_.item} | ?{$_.propName -eq "name"}).propVal
   LongName  = (%{$_.item} | ?{$_.propName -eq "longname"}).propVal
   UserCode  = (%{$_.item} | ?{$_.propName -eq "auth:name"}).propVal
   Mail      = (%{$_.item} | ?{$_.propName -eq "mail:address"}).propVal
   Fax     = (%{$_.item} | ?{$_.propName -eq "fax:number"}).propVal
}} | sort Index 
Disconnect-MFP $Hostname
}

function Add-MFPAB
{
param($Hostname,$Authentication="BASIC",$Username="admin",$Password,$SecurePassword,$EntryType="user",$Index,$Name,$LongName,$UserCode,$Destination="true",$Sender="false",$Mail="true",$MailAddress)
Connect-MFP $Hostname $Authentication $Username $Password $SecurePassword
$url = "http://$Hostname/DH/udirectory"
$add = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:putObjects xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <sessionId></sessionId>
   <objectClass>entry</objectClass>
   <parentObjectId></parentObjectId>
   <propListList xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="">
    <item xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:property[7]">
     <item>
      <propName>entryType</propName>
      <propVal></propVal>
     </item>
     <item>
      <propName>name</propName>
      <propVal></propVal>
     </item>
     <item>
      <propName>longName</propName>
      <propVal></propVal>
     </item>
     <item>
      <propName>isDestination</propName>
      <propVal></propVal>
     </item>
     <item>
      <propName>isSender</propName>
      <propVal></propVal>
     </item>
     <item>
      <propName>mail:</propName>
      <propVal></propVal>
     </item>
     <item>
      <propName>mail:address</propName>
      <propVal></propVal>
     </item>
    </item>
   </propListList>
  </m:putObjects>
 </s:Body>
</s:Envelope>
'@
$add.Envelope.Body.putObjects.sessionId = $script:session
$add.Envelope.Body.putObjects.propListList.item.item[0].propVal = $EntryType
if($Index -ne $NULL){
$a = $add.CreateElement("item")
$a.set_InnerText("")
$b = $add.CreateElement("propName")
$b.set_InnerText("index")
$o = $a.AppendChild($b)
$c = $add.CreateElement("propVal")
$c.set_InnerText($Index)
$o = $a.AppendChild($c)
$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
$add.Envelope.Body.putObjects.propListList.item.item[1].propVal = $Name
$add.Envelope.Body.putObjects.propListList.item.item[2].propVal = $LongName
if($UserCode -ne $NULL){
$a = $add.CreateElement("item")
$a.set_InnerText("")
$b = $add.CreateElement("propName")
$b.set_InnerText("auth:name")
$o = $a.AppendChild($b)
$c = $add.CreateElement("propVal")
$c.set_InnerText($UserCode)
$o = $a.AppendChild($c)
$o = $add.Envelope.Body.putObjects.propListList.item.AppendChild($a)
}
$add.Envelope.Body.putObjects.propListList.item.item[3].propVal = $Destination
$add.Envelope.Body.putObjects.propListList.item.item[4].propVal = $Sender
$add.Envelope.Body.putObjects.propListList.item.item[5].propVal = $Mail
$add.Envelope.Body.putObjects.propListList.item.item[6].propVal = $MailAddress
$add.Envelope.Body.putObjects.propListList.arrayType = "itt:string[]["+$add.Envelope.Body.putObjects.propListList.item.item.count+"]"
[xml]$xml = iwr $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#putObjects"} -Body $add
Disconnect-MFP $Hostname
}

function Remove-MFPAB
{
param($Hostname,$Authentication="BASIC",$Username="admin",$Password,$SecurePassword,$ID)
Connect-MFP $Hostname $Authentication $Username $Password $SecurePassword
$url = "http://$Hostname/DH/udirectory"
$remove = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
 <s:Body>
  <m:deleteObjects xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
   <sessionId></sessionId>
  <objectIdList xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.w3.org/2001/XMLSchema" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="">
  </objectIdList>
   <options xmlns:soap-enc="http://schemas.xmlsoap.org/soap/encoding/" xmlns:itt="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:type="soap-enc:Array" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="http://www.ricoh.co.jp/xmlns/schema/rdh/commontypes" xsi:arrayType="itt:property[1]">
    <item>
     <propName></propName>
     <propVal></propVal>
    </item>
   </options>
  </m:deleteObjects>
 </s:Body>
</s:Envelope>
'@
$remove.Envelope.Body.deleteObjects.sessionId = $script:session
$ID | %{
$x = $remove.CreateElement("item")
$x.set_InnerText("entry:$_")
$o = $remove.Envelope.Body.deleteObjects.objectIdList.AppendChild($x)
}
$remove.Envelope.Body.deleteObjects.objectIdList.arrayType = "itt:string["+$remove.Envelope.Body.deleteObjects.objectIdList.item.count+"]"
[xml]$xml = iwr $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#deleteObjects"} -Body $remove
Disconnect-MFP $Hostname
}

function Disconnect-MFP
{
param($Hostname)
$url = "http://$Hostname/DH/udirectory"
$logout = [xml]@'
<?xml version="1.0" encoding="utf-8" ?>
 <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/" s:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
  <s:Body>
   <m:terminateSession xmlns:m="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory">
    <sessionId></sessionId>
   </m:terminateSession>
  </s:Body>
 </s:Envelope>
'@
$logout.Envelope.Body.terminateSession.sessionId = $script:session
[xml]$xml = iwr $url -Method Post -ContentType "text/xml" -Headers @{SOAPAction="http://www.ricoh.co.jp/xmlns/soap/rdh/udirectory#terminateSession"} -Body $logout
}

Export-ModuleMember Get-MFPAB,Add-MFPAB,Remove-MFPAB