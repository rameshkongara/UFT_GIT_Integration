Public Function gfOpenBrowser(ByVal strBrowserType, ByVal strURL)
	Err.Clear
	Dim strFuncName	'Procedure Name is stored for displaying the Fail error message details
	Call gfunCloseBrowsers()
	'call gfunCloseIEprocess()
	strFuncName = "Error in gfOpenBrowser():"

	Select Case UCase(strBrowserType)
		Case gCnstIEBrowser
			SystemUtil.Run "iexplore.exe",strURL
		Case gCnstFFBrowser
			SystemUtil.Run "firefox.exe",strURL
	End Select

	

End Function
