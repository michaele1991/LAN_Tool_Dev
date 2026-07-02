Option Explicit

    LaunchMacro

    Sub LaunchMacro() 
      Dim xl
      Dim xlBook      
      Dim sCurPath

        sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
        'msgBox sCurPath
		Set xl = CreateObject("Excel.application")
		'msgBox xl
        Set xlBook = xl.Workbooks.Open(sCurPath & "\n8_cnl_lp_nvm_map.xlsm", 0, True)   
		'msgBox xlBook
        xl.Application.Visible = False
        xl.Application.run "n8_cnl_lp_nvm_map.xlsm!Module1.genNvmCMDline"
        xl.DisplayAlerts = False        
        xlBook.Save = True
		
        xl.activewindow.close
        xl.Quit

        Set xlBook = Nothing
        Set xl = Nothing

	End Sub 