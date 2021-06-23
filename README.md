# Mail User Editor
Mail Users Editor is a GUI tool that allows you to modify ProxyAddresses,TargetAddress and Mail attributes for a single user or bulk users

[quick youtube tutorial](https://youtu.be/Cg3Dv0DzX4w)

### some scenarios where MUE can be useful
* cross-forest migrations, where users authenticate against forest A but their mailboxes are migrated to forest B and no Exchange servers are available on Forest A  
* hybrid Exchange environments where users are synced from on-prem to Office 365 (on-prem AD is the source of authority) and there are no available Exchange servers On-prem and you need to edit users' attributes
* anycase where you need to modify the previously mentioned attributes but no exchange servers are available


### Requirments
* the tool is tested on powershell V3.0 and later versions
* the tool can be run on any domain joined machine that has AD powershell module (available in administrative tools)
* the AD account running the tool must have permissions to read and edit users attributes (can run :get-aduser and set-aduser)
* must be run as administrator ( required for set-aduser)

### Instructions
* to run the .ps1 version ,open a powershell window as an administrator and run the script or you can run the included .exe version
* either enter comma-seprated values of required users' samaccountname or select users from a file
* you can use CSV or TXT files that contains samaccountnames
* click "check" button to query Active Directory for the selected users and populate the list with query result
* enter the required domain name for (proxy address and mail attributes) and | or the required domain name for targetaddress attribute
* choose the required Address format
* domain validations are performed.(i.e: entering "domain" instead of "domain.com" is not allowed")
* if replace checkbox is checked (default) : 
 <div class="highlight highlight-source-shell"><pre>
  1. if ProxyAddresses and mail attributes are empty, they will be populated with the provided value
  2. if ProxyAddresses and mail attributes are not empty, they will be replaced with provided value
</pre></div>

* if replace checkbox is unchecked
<div class="highlight highlight-source-shell"><pre>
  1. if ProxyAddresses attribute is empty, the provided value will be added as a primary SMTP address (SMTP:mail@domain.com)
  2. if ProxyAddresses attribute is not empty, the provided value will be added as an additional SMTP address (smtp:mail@domain.com)
</pre></div>

* since there can be only one value for TargetAddress attribute ,TargetAddress will always be replaced with the provided value
* you don't have to modify both proxy and target values , you can modify any of them independently
* click "modify" button to modify selected AD users with the provided values
* press "clear" button to clear Users List and start a new operation
* any errors will be appended to "mailuser-log.txt" file located in script root folder

### Example of Creating required Users File
<div class="highlight highlight-source-shell"><pre>
  get-aduser -SearchBase "CN=Users,DC=target,DC=com" -Filter * | select samaccountname | Export-Csv c:\users.csv -NoTypeInformation
</pre></div>

### Extra Info
* this tool perform asynchronous operations (using runspaces) to be responsive and avoid crashing when performing bulk operations
* completed runspaces are cleaned periodically  to avoid memory leaks
* you can compile .ps1 to an excutable file using [PS2EXE](https://www.powershellgallery.com/packages/ps2exe/1.0.5) , some anti-malwares  may flag the compiled version as a malware. this is a false-positive 

