$Global:syncHash = [hashtable]::Synchronized(@{})
$newRunspace = [runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

# Load WPF assembly 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')

$psCmd = [PowerShell]::Create().AddScript( {
        [xml]$xaml = @"
    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
              Title="Mail Users Editor" Height="450" Width="1000" Background="#FFC9C8C8" FontWeight="Normal" FontFamily="Arial" WindowStartupLocation="CenterScreen" >

    <Window.Resources>
        <SolidColorBrush x:Key="TextBox.Static.Border" Color="#FFE2E3EA"/>
        <SolidColorBrush x:Key="TextBox.MouseOver.Border" Color="#FFC5DAED"/>
        <SolidColorBrush x:Key="TextBox.Focus.Border" Color="#FFB5CFE7"/>
        <Style TargetType="{x:Type TextBox}">
       
            <Setter Property="BorderThickness" Value="1"/>
            
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type TextBox}">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderThickness="{TemplateBinding BorderThickness}" BorderBrush="{TemplateBinding BorderBrush}" SnapsToDevicePixels="True" CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Focusable="false" HorizontalScrollBarVisibility="Hidden" VerticalScrollBarVisibility="Hidden"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocused" Value="true">
                                <Setter Property="BorderBrush" TargetName="border" Value="#FF0B0B0C"/>
                                <Setter Property="CornerRadius" TargetName="border" Value="0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            
        </Style>
        
       
        <Style TargetType="{x:Type Button}">
           
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="{x:Type Button}">
                        <Border x:Name="border" Background="{TemplateBinding Background}" BorderThickness="{TemplateBinding BorderThickness}" BorderBrush="{TemplateBinding BorderBrush}" SnapsToDevicePixels="true" CornerRadius="6" Margin="-2,-6,0,0">
                            <ContentPresenter x:Name="contentPresenter" Focusable="False" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" Margin="{TemplateBinding Padding}" RecognizesAccessKey="True" SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="true">
                                <Setter Property="Background" TargetName="border" Value="white"/>
                                <Setter Property="BorderBrush" TargetName="border" Value="#FF1F4ADC"/>
                                <Setter Property="FontWeight" Value="Bold"/>
                                <Setter Property="FontSize" Value="13"/>
                                <Setter Property="BorderThickness" Value="2"/>
                            </Trigger>
                            
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
  <Grid>
        <StackPanel Orientation="Vertical" HorizontalAlignment="Left" Margin="41,62,0,239" Width="354">
            <StackPanel Name="normaluse" Orientation="Horizontal" Height="33">
                <Label Content="select User(s)"/>
                <TextBox Name="users" Text="" TextWrapping="Wrap"  Width="146" Height="30" Margin="74 0 0 0"  />
            </StackPanel>
            <StackPanel Name="usefile" Orientation="Horizontal" Height="64" Margin="0 20 0 0">
                <Button Name="selectfile" Content="select file"    Height="24" Width="79" Margin="3" Background="White"/>
                <TextBlock Name="filepath"  HorizontalAlignment="Left"  Text="" TextWrapping="Wrap" Height="30"  Width="238" Margin="75 0 0 0"/>

            </StackPanel>
        </StackPanel>
        <TextBox Name="primary" HorizontalAlignment="Left" Margin="199,191,0,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top" Width="145" Height="24"/>
        <Label Name="primlabel" Content="Proxy Address Domain" HorizontalAlignment="Left" Margin="41,191,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.508,2.684"/>
        <Label Name="targetlabel" Content="Target Address Domain" HorizontalAlignment="Left" Margin="41,226,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.508,2.684"/>
        <TextBox Name="target" HorizontalAlignment="Left" Margin="199,226,0,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top" Width="145" Height="24"/>

        <Button Name="check" Content="Check Users" HorizontalAlignment="Left" Margin="49,343,0,0" VerticalAlignment="Top" Height="30" Width="79" FontFamily="Segoe UI" Background="White"/>
        <Button Name="modify" Content="Modify" HorizontalAlignment="Left" Margin="740,343,0,0" VerticalAlignment="Top" Height="29" Width="65" FontFamily="Segoe UI" Foreground="Gray" />
        <Button Name="clear" Content="clear" HorizontalAlignment="Left" Margin="642,343,0,0" VerticalAlignment="Top" Width="63" Height="30" FontFamily="Segoe UI" Foreground="Gray" />
        <ListView Name="listView" HorizontalAlignment="Left" Height="250" Margin="500,55,0,0" VerticalAlignment="Top" Width="480" BorderBrush="Black"   >
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding displayname}" Width="100"/>
                    <GridViewColumn Header="Email" DisplayMemberBinding ="{Binding Email}" Width="190"/>
                    <GridViewColumn Header="Target Email"  DisplayMemberBinding ="{Binding target}" Width="190" />
                </GridView>
            </ListView.View>
        </ListView>
        <CheckBox Name="filecheckbox" Content="use a file" HorizontalAlignment="Left" Margin="41,30,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="replace" Content="Replace" HorizontalAlignment="Left" Margin="357,195,0,0" VerticalAlignment="Top" Foreground="RED" IsChecked="True"/>
        
        <Label Name="loading" Content="Updating..." HorizontalAlignment="Left" Margin="665,145,0,0" VerticalAlignment="Top" FontSize="22" FontWeight="Bold" />
        <Label Name="formatlabel" Content="Address Format" HorizontalAlignment="Left" Margin="41,267,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="format" HorizontalAlignment="Left" Margin="199,271,0,0" VerticalAlignment="Top" Width="145">
            <ComboBoxItem  Content="first.last"></ComboBoxItem>
            <ComboBoxItem  Content="last.first"></ComboBoxItem>
            <ComboBoxItem  Content="first_last"></ComboBoxItem>
            <ComboBoxItem  Content="last_first"></ComboBoxItem>
        </ComboBox>
       
    </Grid>
</Window>
"@

 
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    
   
        #required window controls
        $syncHash.Window = [Windows.Markup.XamlReader]::Load( $reader )
        $syncHash.users = $syncHash.Window.FindName("users")
        $syncHash.primary = $syncHash.Window.FindName("primary")
        $syncHash.target = $syncHash.Window.FindName("target")
        $syncHash.check = $syncHash.Window.FindName("check")
        $syncHash.modify = $syncHash.Window.FindName("modify")
        $syncHash.list = $syncHash.Window.FindName("listView")
        $syncHash.test = $syncHash.Window.FindName("test")
        $syncHash.clear = $syncHash.Window.FindName("clear") 
        $syncHash.loading = $syncHash.Window.FindName("loading") 
        $syncHash.primlabel = $syncHash.Window.FindName("primlabel")
        $syncHash.targetlabel = $syncHash.Window.FindName("targetlabel")
        $syncHash.filecheckbox = $syncHash.Window.FindName("filecheckbox")
        $syncHash.selectfile = $syncHash.Window.FindName("selectfile")
        $syncHash.filepath = $syncHash.Window.FindName("filepath")
        $syncHash.manualselect = $syncHash.Window.FindName("manualselect")
        $syncHash.normaluse = $syncHash.Window.FindName("normaluse")
        $syncHash.usefile = $syncHash.Window.FindName("usefile")
        $syncHash.replace = $syncHash.Window.FindName("replace")
        $syncHash.format = $syncHash.Window.FindName("format")
   
        # Synchronized list to add runspaces 
        $Script:JobCleanup = [hashtable]::Synchronized(@{})
        $Script:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
        $jobCleanup.Flag = $True


        
        # disable modify button
        function disable-modify() {

            $syncHash.modify.IsEnabled = $fasle
            $syncHash.modify.Foreground = "Gray"
            $syncHash.modify.Background = "#FFC9C8C8"

        }
        # enable modify button
        function enable-modify() {

            $syncHash.modify.IsEnabled = $true
            $syncHash.modify.Background = "White"
            $syncHash.modify.Foreground = "Black"
        }
        # disable domains textboxes and labels 
        function disable-domains() {

            $syncHash.primary.IsEnabled = $false
            $syncHash.target.IsEnabled = $false
            $syncHash.primary.Background = "#FFC9C8C8"
            $syncHash.target.Background = "#FFC9C8C8"
            $syncHash.primlabel.Foreground = "#FFE0D8D8"
            $syncHash.targetlabel.Foreground = "#FFE0D8D8"
            $syncHash.replace.IsEnabled = $false
            $syncHash.replace.Foreground = "#FFE0D8D8"

  
        }
        # enable domains textboxes and labels 
        function enable-domains() {
            $syncHash.primary.IsEnabled = $true
            $syncHash.target.IsEnabled = $true
            $syncHash.primary.Background = "White"
            $syncHash.target.Background = "White"
            $syncHash.primlabel.Foreground = "Black"
            $syncHash.targetlabel.Foreground = "Black"
            $syncHash.replace.IsEnabled = $true
            $syncHash.replace.Foreground = "RED"
 
        }

        # setup app intial view
        function initial-view() {
 
            disable-modify
            disable-domains
            $syncHash.clear.IsEnabled = $false
            $syncHash.clear.Foreground = "Gray"
            $syncHash.clear.Background = "#FFC9C8C8"
            $syncHash.usefile.Visibility = "Collapsed"
            $syncHash.loading.Visibility = "Hidden"
            $synchash.format.SelectedIndex = 0


        }

        # enable modify button if text changed event occured in primary or target textboxes
        $syncHash.primary.add_TextChanged( {

                if ($syncHash.primary.Text.Length -ne 0 -or $syncHash.target.Text.Length -ne 0) {
   
                    enable-modify
                }
                else {
                    disable-modify
                }

            })

        $syncHash.target.add_TextChanged( {

                if ($syncHash.primary.Text.Length -ne 0 -or $syncHash.target.Text.Length -ne 0 ) {
                    if ($syncHash.format.SelectedIndex -gt 0) {

                        enable-modify
                    }
      
                }
                else {
                    disable-modify
                }

            })

        initial-view


        #modify the view , if file checkbox is checked
        $synchash.filecheckbox.add_Checked( {

 
                $syncHash.list.ItemsSource = $null
                $syncHash.normaluse.Visibility = "Collapsed"
                $syncHash.usefile.Visibility = "Visible"


            })

        #modify the view , if file checkbox is unchecked
        $syncHash.filecheckbox.add_Unchecked( {

                $syncHash.list.ItemsSource = $null
                $syncHash.normaluse.Visibility = "Visible"
                $syncHash.usefile.Visibility = "Collapsed"

            })



        #open file explorer to select a file (currently set to CSV and TXT files)
        $syncHash.selectfile.add_Click( {

                Add-Type -AssemblyName System.Windows.Forms
                $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
                    InitialDirectory = $PSScriptRoot
                    Filter           = 
                    'Text|*.txt|CSV files (*.csv)|*.csv'
                    #'CSV files (*.csv)|*.csv|Text|*.txt'
    
                }

                $null = $FileBrowser.ShowDialog()
                $syncHash.filepath.Text = $fileBrowser.FileName

            })


        #ASYNC check selected users (manually of from a file) against Active Directory
        $syncHash.check.add_Click( {
                # perform some input validation and empty the current List of users
                $syncHash.list.ItemsSource = $null
 
                if (!$syncHash.filecheckbox.IsChecked) {
                    $syncHash.usersList = $syncHash.users.Text.Split(",").Trim()
  
                }

                else {
  
                    try {
                        $syncHash.usersList = (Get-Content -Path $syncHash.filepath.Text).trim('"', ' ') | where { $_ -NotLike "samac*" -and $_ -notlike "----------*" -and $_.trim() -ne "" }
                        if ($syncHash.usersList[0] -eq "samaccountname") {
                            $syncHash.usersList.Remove(0)
                        }
                    }
                    catch {
                        [System.Windows.MessageBox]::Show("select a file")
          
                    }
      
      
                }
  

                if ($syncHash.usersList.Length -eq 0) {
                    [System.Windows.MessageBox]::Show("No users selected !")
                    return
      
                }

  
                $Definition = Get-Content Function:\enable-domains -ErrorAction Stop
                $SessionStateFunction = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList 'enable-domains', $Definition
                $InitialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
                $InitialSessionState.Commands.Add($SessionStateFunction)
                $newRunspace = [runspacefactory]::CreateRunspace($InitialSessionState)
                $newRunspace.ApartmentState = "STA"
                $newRunspace.ThreadOptions = "ReuseThread"          
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("usersList", $syncHash.usersList) 
        
                $newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash) 

                $syncHash.loading.Visibility = "Visible"

                $PowerShell = [PowerShell]::Create().AddScript( {
       
        

                        $finalusers = @()
                        foreach ($user in $syncHash.usersList) { 
                            try {
                                $u = Get-ADUser -Identity $user -Properties emailaddress, targetaddress
                                if ($u.EmailAddress -eq $null ) {
                                    $email = ""
           
                                }

                                else {
            
                                    $email = $u.EmailAddress
          
          
         
                                }

                                if ($u.targetaddress -eq $null) {
                                    $syncHash.targetaddr = ""
        
                                }

                                else {
            
        
                                    $syncHash.targetaddr = $u.targetAddress.Substring(5)
          
         
                                }

       
                                $finalusers += [pscustomobject]@{displayname = $u.Name; Email = $email; target = $syncHash.targetaddr; sam = $u.samaccountname; first = $u.givenname; last = $u.surname }
        
       
                            }

                            catch {
   
                                #append any errors to mailuser-log file located in script Root folder
                                $Error | Out-File "$PSScriptRoot\mailuser-log.txt" -Append
  
                                "---------------------------------------------------------------------------------------------------" | Out-File "$PSScriptRoot\mailuser-log.txt" -Append
      
                            }
 
                        }
       
       
                        $syncHash.list.Dispatcher.Invoke([action] {
                                $syncHash.list.ItemsSource = [System.Windows.Data.ListCollectionView] $finalusers
                                $syncHash.loading.Visibility = "Hidden"
                                if ($syncHash.list.items.Count -gt 0) {
             
                                    enable-domains
                                    $syncHash.clear.IsEnabled = $true
                                    $syncHash.clear.Background = "White"
                                    $syncHash.clear.Foreground = "Black"
                                    if ($syncHash.primary.Text.Length -ne 0 -or $syncHash.target.Text.Length -ne 0) {
                                        if ($syncHash.format.SelectedIndex -gt 0) {
                                            $syncHash.modify.IsEnabled = $true
                                            $syncHash.modify.Background = "White"
                                            $syncHash.modify.Foreground = "Black"
                                        }
                                    }

                                }
             
            
                            }, "Normal")


        
                    })

                $PowerShell.Runspace = $newRunspace

                [void]$Script:Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell
                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))
            })


        #clear selected users

        $syncHash.clear.add_Click( {

                $syncHash.list.ItemsSource = $null
                $syncHash.primary.IsEnabled = $false
                $syncHash.target.IsEnabled = $false
                $syncHash.users.Text = ""
                $syncHash.filepath.Text = ""
                disable-modify

            })




        # ASYNC modify selected users properties
        $syncHash.modify.add_Click( {
                # perform some input validation
                $override = $false
                if ($syncHash.target.Text -eq "" -and $syncHash.primary.Text -eq "") {
                    [System.Windows.MessageBox]::Show("enter a valid primary or target address")
       
                }

                if ($syncHash.list.Items.Count -eq 0 ) {
                    [System.Windows.MessageBox]::Show("No Users Selected")
                    return
       
                }

                $contacts = $syncHash.list.Items
                $domainvalidation = "^[a-zA-Z0-9][a-zA-Z0-9-]{1,61}[a-zA-Z0-9](?:\.[a-zA-Z]{2,})+$"
                if ($syncHash.primary.Text) {
                    if (!($syncHash.primary.Text -match $domainvalidation)) {
                        [System.Windows.MessageBox]::Show("please enter a valid primary address domain")
                        return

                    }
    
                }


                if ($syncHash.target.Text) {
                    if (!($syncHash.target.Text -match $domainvalidation)) {
                        [System.Windows.MessageBox]::Show("please enter a  valid target address domain")
                        return

                    }
                }


                if ($syncHash.replace.IsChecked -and $syncHash.primary.Text) {
   
                    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
                    $result = [System.Windows.Forms.MessageBox]::Show('Do you want to override the current primary address ?' , "Warning" , 4)
                    if ($result -eq 'Yes') {
                        $override = $true
                    }

                }


                $newRunspace = [runspacefactory]::CreateRunspace()
                $newRunspace.ApartmentState = "STA"
                $newRunspace.ThreadOptions = "ReuseThread"          
                $newRunspace.Open()
                $newRunspace.SessionStateProxy.SetVariable("contacts", $contacts)       
                $newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)
                $newRunspace.SessionStateProxy.SetVariable("override", $override) 
                $newRunspace.SessionStateProxy.SetVariable("primary", $syncHash.primary.Text) 
                $newRunspace.SessionStateProxy.SetVariable("target", $syncHash.target.Text) 
                $newRunspace.SessionStateProxy.SetVariable("format", $syncHash.format.SelectedItem.Content) 
                $syncHash.loading.Visibility = "Visible"
                $PowerShell = [PowerShell]::Create().AddScript( {
                        # a function that  return the required format based on the selected address format
                        function email-format ($firstname, $lastname, $domain, $format) {
                            switch ( $format ) {
                                "first.last" { $result = "$($firstname).$($lastname)@$($domain)" }
                                "last.first" { $result = "$($lastname).$($firstname)@$($domain)" }
                                "first_last" { $result = "$($firstname)_$($lastname)@$($domain)" }
                                "last_first" { $result = "$($lastname)_$($firstname)@$($domain)" }
                            }

                            return $result
                        }

                        try {
                            foreach ($contact in $contacts) {
     
                                # if replace checkbox is checked , replace the current email address and proxy addresses
                                if ($override) {
     
                                    set-aduser $contact.sam  -EmailAddress  "$(email-format -firstname $contact.first -lastname $contact.last -domain $primary -format $format)"
 
                                    Set-ADUser $contact.sam  -replace @{ProxyAddresses = "SMTP:$(email-format -firstname $contact.first -lastname $contact.last -domain $primary -format $format)" }
     
                                }
                                # if replace checkbox is unchecked 
                                else {
                                    #if user has a primary proxy address , add this one as an additional proxy address
                                    if ($primary) {
                                        if ((Get-ADUser $contact.sam -Properties Proxyaddresses).proxyaddresses) {
                                            Set-ADUser $contact.sam  -add @{ProxyAddresses = "smtp:$(email-format -firstname $contact.first -lastname $contact.last -domain $primary -format $format)" }
       
                                        }
                                        # if there's no primary proxy address , add this one as a primary proxy address 
                                        else {
                                    
                                            Set-ADUser $contact.sam  -add @{ProxyAddresses = "SMTP:$(email-format -firstname $contact.first -lastname $contact.last -domain $primary -format $format)" }
            
                
                                        }
                                    }
       
                                }
     
                                # since there can't be more than one value in target address attribute ,target address always gets replaced by the new value
                                if ($target) {
                                    Set-ADUser $contact.sam  -replace @{targetaddress = "SMTP:$(email-format -firstname $contact.first -lastname $contact.last -domain $target -format $format)" }
                                }
                            }

  
   

                            $syncHash.loading.Dispatcher.Invoke([action] {
                                    $syncHash.loading.Visibility = "Hidden"
             
            
                                }, "Normal")


                        }
  
 
                        catch {
                            #append any errors to mailuser-log file located in script Root folder
     
                            $Error | Out-File "$PSScriptRoot\mailuser-log.txt" -Append
  
                            "---------------------------------------------------------------------------------------------------" | Out-File "$PSScriptRoot\mailuser-log.txt" -Append
   
                        }

                    })
                $PowerShell.Runspace = $newRunspace

                [void]$Script:Jobs.Add((
                        [pscustomobject]@{
                            PowerShell = $PowerShell
                            Runspace   = $PowerShell.BeginInvoke()
                        }
                    ))
            })




        $syncHash.primary.add_TextChanged( {
                if ($syncHash.primary.Text.Length -ne 0 -or $syncHash.target.Text.Length -ne 0 ) {
                    enable-modify
                }
                else {
                    disable-modify
                }

            })

        $syncHash.target.add_TextChanged( {
                if ($syncHash.primary.Text.Length -ne 0 -or $syncHash.target.Text.Length -ne 0  ) {
                    enable-modify
      
                }
                else {
                    disable-modify
                }

            })
        # runspace that continuously checks for completed runspaces and dispose them
        $newRunspace = [runspacefactory]::CreateRunspace()
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"          
        $newRunspace.Open() 
        $newRunspace.SessionStateProxy.SetVariable("jobs", $Script:Jobs) 
        $newRunspace.SessionStateProxy.SetVariable("JobCleanup", $Script:JobCleanup) 
        $jobCleanup.PowerShell = [powershell]::Create().AddScript(
            {
                do {    
            
                    foreach ($runspace in $jobs) {      
                        if ($runspace.Runspace.isCompleted) {
                            $runspace.powershell.EndInvoke($runspace.Runspace) | Out-Null
                            $runspace.powershell.Runspace.Dispose() 
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null               
                        } 
                    }
                    $temphash = $jobs.clone()
                    $temphash | Where-Object { $_.runspace -eq $Null } | ForEach-Object { $jobs.remove($_) }        
                    Start-Sleep -Seconds 1    
                } while ($jobCleanup.Flag)
            }
        )
        $jobCleanup.PowerShell.Runspace = $newRunspace
        $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  

        # dispose of jobCleanup runspace on window close
        $syncHash.Window.Add_Closed( {
                $jobCleanup.Flag = $False
                $jobCleanup.powershell.Runspace.Dispose() 
                $jobCleanup.PowerShell.Dispose()  
                   
            })
       
        $syncHash.Window.ShowDialog() | Out-Null
       

    })
$psCmd.Runspace = $newRunspace
$data =$psCmd.BeginInvoke() | Out-Null



