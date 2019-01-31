<%
	
	' INSTALLATION:
	'************************************************************************
	
	' Uses BCrypt.Net (A .Net port of jBCrypt implemented in C#)
	' https://archive.codeplex.com/?p=bcrypt
	
	' Make sure you have the lastest .NET Framework installed (tested on .NET Framework 4.7.2)
	
	' Open IIS, go to the applicaton pools and open the application pool being used by your 
	' website. Check the .NET CRL version
	' E.g: v4.0.30319
	
	' Navigate to the CRL folder
	' E.g: C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Copy over: 
	' ClassicASP.Bcrypt.dll
	
	' Run CMD as administrator
	
	' Change the directory to your CRL folder
	' E.g: cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
	
	' Run the following command: RegAsm ClassicASP.Bcrypt.dll /tlb /codebase
	
	
	' BCRYPT IN VBSCRIPT:
	'************************************************************************
	
	Const Bcrypt_workFactor = 10
	
	function Bcrypt_hash(ByVal password)
		
		Dim Bcrypt : set Bcrypt = server.CreateObject("ClassicASP.Bcrypt")
		
			Bcrypt_hash = Bcrypt.hash(password,Bcrypt_workFactor)
		
		set Bcrypt = nothing
		
	end function
	
	function Bcrypt_verify(ByVal password, ByVal BcryptHash)
	
		Dim Bcrypt : set Bcrypt = server.CreateObject("ClassicASP.Bcrypt")
			
			Bcrypt_verify = Bcrypt.verify(password,BcryptHash)
			
		set Bcrypt = nothing
		
	end function
	
	Dim bc_hash, start, testPassword
	
	testPassword = "myPassword"
	
	response.write "<p><b>Password:</b> " & testPassword & "</p>"
	
	start = Timer()
	
	bc_hash = Bcrypt_hash(testPassword)
				
	response.write "<p><b>Bcrypt Hash:</b> " & bc_hash & "</p>"
	response.write "<p><b>Time to execute:</b> " & formatNumber(Timer()-start,4) & "s</p>"
	
	start = Timer()
	
	response.write "<p><b>Bcrypt Verified:</b> " & Bcrypt_verify(testPassword,bc_hash) & "</p>"
	response.write "<p><b>Time to execute:</b> " & formatNumber(Timer()-start,4) & "s</p>"	
	
%>
