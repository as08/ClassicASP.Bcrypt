Bcrypt hashing algorithm implemented in VBscript for Classic ASP

## INSTALLATION:
Uses BCrypt.Net (A .Net port of jBCrypt implemented in C#)
https://archive.codeplex.com/?p=bcrypt

Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
Open IIS, go to the applicaton pools and open the application pool being used by your 
Classic ASP application. Check the .NET CRL version
E.g: v4.0.30319
	
Navigate to the CRL folder
E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Copy over: ClassicASP.Bcrypt.dll
	
Run CMD as administrator

Change the directory to your CRL folder
E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
Run the following command: RegAsm ClassicASP.Bcrypt.dll /tlb /codebase

## Example output from bcrypt.asp

	Password: myPassword
	
	Bcrypt Hash: $2a$10$s9THkLgv6bJU9Qio8Id2N.FpB79P5w4zdsHvzMAxHK/ht3KxQnsca

	Time to execute: 0.4854s

	Bcrypt Verified: True

	Time to execute: 0.6074
