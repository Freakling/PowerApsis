# PowerApsis
Powershell module to provide easy integration with Apsis Pro API

This project is a wrap of the existing Apsis API functionality to cmdlets. It is in no way affiliated with the product or company.

There are a number of requirements to use this module, which include an Apsis Pro account and API key. Contact APSIS for more information (https://apsis.se)

http://se.apidoc.anpdm.com/

# Disclaimer
This module is by any means not complete, it is just a collection of functions that I have used and needed to simplify the tasks I use
If you have any other needs, submit a Pull Request.

# Installation
This module is published automatically to PowerShell Gallery.

https://www.powershellgallery.com/packages/PowerApsis/

```powershell
Install-Module PowerApsis
```

# Example


```powershell
Get-Command -Module PowerApsis

CommandType     Name                                               Version    Source                                                     
-----------     ----                                               -------    ------                                                     
Function        Add-ApsisEventAttendee                             1.0        PowerApsis                                                 
Function        Get-ApsisEventAttendees                            1.0        PowerApsis                                                 
Function        Get-ApsisEventOptions                              1.0        PowerApsis                                                 
Function        Get-ApsisEvents                                    1.0        PowerApsis                                                 
Function        Get-ApsisMailinglists                              1.0        PowerApsis                                                 
Function        Get-ApsisSubscribers                               1.0        PowerApsis                                                 
Function        Invoke-ApsisAPI                                    1.0        PowerApsis
Function        New-ApsisMailingLists                              1.0        PowerApsis                                
Function        Register-ApsisEventAttendee                        1.0        PowerApsis                                
Function        Remove-ApsisMailingLists                           1.0        PowerApsis                                
Function        Remove-ApsisSubscriber                             1.0        PowerApsis                                
Function        Set-ApsisEventAttendeeStatus                       1.0        PowerApsis                                
Function        Set-ApsisSubscriber                                1.0        PowerApsis  
```
