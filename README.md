This is a Component Object Model (COM) Dynamic-link library (DLL) coded in C# that can be set in Classic ASP using VBscripts "CreateObject" method and allows you to compute Bcrypt hashes.

Bcrypt is a password hashing function designed by Niels Provos and David Mazières, based on the Blowfish cipher, and presented at USENIX in 1999. Besides incorporating a salt to protect against rainbow table attacks, bcrypt is an adaptive function: over time, the iteration count can be increased to make it slower, so it remains resistant to brute-force search attacks even with increasing computation power.

https://en.wikipedia.org/wiki/Bcrypt

Also see:

https://github.com/as08/ClassicASP.Argon2

https://github.com/as08/ClassicASP.PBKDF2

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

## Usage

	Set Bcrypt = Server.CreateObject("ClassicASP.Bcrypt")
	
	' Generate a hash with a default work factor of 10
	Bcrypt.Hash("myPassword")
	
	' Generate a hash with a custom work factor
	Bcrypt.Hash("myPassword",12) ' >=4 <=31
	
	' Verify a hash
	Bcrypt.Verify("myPassword","$2a$10$s9THkLgv6bJU9Qio8Id2N.FpB79P5w4zdsHvzMAxHK/ht3KxQnsca") ' True / False

## Example output from bcrypt.asp

	Password: myPassword
	
	Bcrypt Hash: $2a$10$s9THkLgv6bJU9Qio8Id2N.FpB79P5w4zdsHvzMAxHK/ht3KxQnsca

	Time to execute: 0.4854s

	Bcrypt Verified: True

	Time to execute: 0.6074
