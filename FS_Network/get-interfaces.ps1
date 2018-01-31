                      <#
                      .Synopsis
                         Gets Network device interfaces
                      .DESCRIPTION
                         Gets Network device interfaces information from EM7 with methods to call Switch via SSH
                      .EXAMPLE
                         Basic Use:
                              Get-Interfaces -market CLT -SearchValue PEAK10
                      .EXAMPLE
                         Assign the resultset to a variable to execute methods
                              $Interfaces = Get-Interfaces -market CLT -SearchValue Peak10
                              $interfaces.getrun()
                      .Example
                         Use this function on switches in the "Parent" Market. Example: 'CLT-access-01','CLT2-access-05','CLT3-Access-10'
                              $interfaces = Get-Interfaces -market CLT -SearchValue Peak10 -Parent $True
                                  Then just select the Switches of interest
                                  Use the Out-Gridview filter!
                      .Parameter Market
                         Inputs to this cmdlet (if any)
                      .Parameter SearchValue
                         SearchValue will look on the Alias field to find anything you put in it
                      .Parameter Parent
                          [Not Mandatory] This is a boolean field value must be set to $true or $false. True = combines all submarkets with parent market CLT = clt1,clt2,clt3,clt4
                      .NOTES
                         Methods

                      #>
                          [CmdletBinding()]
                          param (
                              [Parameter(Mandatory=$True,Position=1)]
                              [market]$market,
                              [Parameter(Mandatory=$false,Position=2)]
                              [string]$SearchValue,
                              [Parameter(Position=3)]
                              [ValidateSet('1', '2')]
                              [Int]$Parent
                          )
                          Enum Market{
                              ATL = 1
                              ATL1 = 2
                              ATL2 = 3
                              ATL3 = 4
                              CIN = 5
                              CIN1 = 6
                              CIN2 = 7
                              CLT = 8
                              CLT1 = 9
                              CLT2 = 10
                              CLT3 = 11
                              CLT4 = 12
                              FLL = 13
                              FLL1 = 14
                              FLL2 = 15
                              JAX = 16
                              JAX1 = 17
                              JAX2 = 18
                              LOU = 19
                              LOU1 = 20
                              LOU2 = 21
                              LOU3 = 22
                              LOU4 = 23
                              NAS = 24
                              NAS1 = 25
                              NAS4 = 26
                              NAS2 = 27
                              NAS3 = 28
                              NAS5 = 29
                              RAL = 30
                              RAL1 = 31
                              RAL2 = 32
                              RAL3 = 33
                              RIC = 34
                              RIC1 = 35
                              RIC2 = 36
                              TPA = 37
                              TPA1 = 38
                              TPA2 = 39
                              TPA3 = 40
                              LAB = 41
                          }
                          #EM7 URI
                          $UriPre = 'https://overlook.peak10.com/api/device/'
                          $UriPost = '/interface?limit=1000&extended_fetch=1'
                          #Credentials
                          if(!$creds){$global:creds = Get-Credential -UserName $env:USERNAME -Message 'Enter EM7 Password'}
                          $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
                          $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                          $ErrorActionPreference = "SilentlyContinue"

                          #
                          $counter = 0
                          if($parent -eq '1')
                              {
                                  $switch = $switches | where-object {$_.market -like "$($market)*"}| Sort-Object Market, Type, Name| Out-GridView -Title 'Select Switch' -PassThru

                              }else
                              {
                                  $switch = $switches | where-object {$_.market -eq $market}| Sort-Object Type| Out-GridView -Title 'Select Switch' -PassThru
                              }
                          $totalcount = $switch.count

                          $return = foreach($sw in $switch){
                              $em7_id = $sw.em7_id
                              $uri = "$($UriPre)$($em7_id)$($UriPost)"
                              Write-Progress -Activity 'Getting Interfaces' -CurrentOperation $sw.name -PercentComplete (($counter / $totalcount) * 100) -Status "#$($counter) Total:
                      $($totalcount)"
                              $Results = Invoke-RestMethod -Method Get -Uri $Uri -Credential $creds
                              $ints = $results.result_set |  GM -MemberType NoteProperty | select name
                              foreach($int in $ints) {
                                 if(!$SearchValue)
                                 {
                                      $results.result_set.$($int.name) | select @{N="Hostname";E={@($sw.Name)}}, name, alias, @{N="Sts";E={@(Switch($_.ifoperstatus){1{"Up"}2{ "Down" }})}},
                      @{N="AdminStatus";E={@(Switch($_.ifAdminStatus){1{"Up"}2{ "Down" }})}}, ifAdminstatus, ifoperstatus
                                 }else
                                 {
                                      $results.result_set.$($int.name) | select @{N="Hostname";E={@($sw.Name)}}, name, alias, @{N="Sts";E={@(Switch($_.ifoperstatus){1{"Up"}2{ "Down" }})}},
                      @{N="AdminStatus";E={@(Switch($_.ifAdminStatus){1{"Up"}2{ "Down" }})}}, ifAdminstatus, ifoperstatus| where-object {$_.alias -like "*$($searchvalue)*"}
                                 }
                              }$counter ++
                          }

                          $return | ForEach-Object{ $_ | Add-Member -MemberType ScriptProperty -name CustID -value {$this.alias.split(":")[1]}}
                          $return |ForEach-Object{ $_ | Add-Member -MemberType ScriptProperty -name Status -value {[string]"$($this.sts)/$($this.adminstatus)"}}
                          $return | Add-Member -MemberType ScriptMethod -Name ShowRun -Value {
                                  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
                                  $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                                  $result = Get-sCommand -Hostname $this.hostname  -command "show run interface $($this.name)"
                                  $this | add-member -membertype NoteProperty -name ShowRun -value $result
                                  if($this.name -like 'Vl*'){
                                      $results = ($result -split '[\r\n]') |? {$_}
                                      $ipaddLine = $results | ForEach-Object {($_ |  Select-String -Pattern 'IP Address')}
                                      $regexIPaddress = '\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b'
                                      $ip = ([regex]::Match($ipaddLine, $regexIPaddress)).value
                                      $this | add-member -MemberType NoteProperty -name IPaddress -Value $ip
                                  }
                                  return $result
                          }


                      #show interface descriptions
                          $return | Add-Member -MemberType ScriptProperty -Name ShowIntDetails -Value {
                                  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
                                  $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                                  $result = Get-sCommand -Hostname $this.hostname -command "show interface $($this.name)"
                                  return $result
                          }

                          $return | Add-member -MemberType ScriptMethod -Name SwitchDesc -value {
                                  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
                                  $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                                  $result = Get-sCommand -Hostname $this.hostname  -command "show interface description"
                                  return $result
                          }

                          #Start Putty in ConEMU from Pipeline value
                          $return | Add-member -MemberType ScriptMethod -Name cPutty -value {
                                  $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($creds.password)
                                  $pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
                                  $host = $this.hostname | Out-GridView -PassThru
                                  Start-cPutty -Hostname $host
                          }

                          #VLAN ID
                          $return | foreach {$_ | Add-Member -MemberType NoteProperty -name VLAN -value ([string]($_.alias.split(":")[2].substring(1,3)))}
                          #Status
                          #$return | ForEach-Object {$_ | add-member -MemberType NoteProperty -name 'PortStatus' -value $($this.status)"-"$($this.adminstatus)}
                          #Market
                          $return |  ForEach-Object {$_ | add-member -MemberType NoteProperty -name 'Market' -value ($($_.hostname).split("-"))[0].substring(0,3)}
                          #Default View
                          $errorActionPreference = "Continue"
                          $defaultProperties = @('CustID','Hostname','Name','Alias', 'Status','VLAN','IPaddress')
                          $defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet(‘DefaultDisplayPropertySet’,[string[]]$defaultProperties)
                          $PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
                          $return | Add-Member MemberSet PSStandardMembers $PSStandardMembers
                          return $return