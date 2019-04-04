<#
.SYNOPSIS
Get-NETFrameworkVersion.ps1

.DESCRIPTION 
Display the current .NET Framework version installed.

.EXAMPLE
Get-NETFrameworkVersion.ps1

.NOTES
Written by: Darren Lucht

* Github:	https://github.com/DarrenLucht/PowerShell/tree/master/Exchange

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

License:

The MIT License (MIT)

Copyright (c) 2019 Darren Lucht

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Change Log:
V1.00, 03/27/2019 - Initial version

The following table lists the release keys on individual operating systems for .NET Framework 4.5 and later versions.

.NET Framework 4.5	    378389  All Windows operating systems            
.NET Framework 4.5.1	378675  On Windows 8.1 and Windows Server 2012 R2
                        378758  On all other Windows operating systems
.NET Framework 4.5.2	379893  All Windows operating systems
.NET Framework 4.6	    393295  On Windows 10
                        393297  On all other Windows operating systems
.NET Framework 4.6.1	394254  On Windows 10 November Update systems
                        394271  On all other Windows operating systems
.NET Framework 4.6.2	394802  On Windows 10 Anniversary Update and Windows Server 2016
                        394806  On all other Windows operating systems
.NET Framework 4.7	    460798  On Windows 10 Creators Update
                        460805  On all other Windows operating systems
.NET Framework 4.7.1	461308  On Windows 10 Fall Creators Update and Windows Server, version 1709
                        461310  On all other Windows operating systems
.NET Framework 4.7.2	461808  On Windows 10 April 2018 Update and Windows Server, version 1803
                        461814  On all other Windows operating systems
#>

Start-Transcript Get-NETFrameworkVersion.txt
$Version = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' ).Version
$Release = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' ).Release
Write-Host ""
Write-Host ".NET Framework" $Version "Release" $Release
Write-Host ""
Stop-Transcript
