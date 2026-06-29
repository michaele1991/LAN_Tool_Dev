Option Explicit

    LaunchMacro

    Sub LaunchMacro() 
      Dim xl
      Dim xlBook      
      Dim sCurPath

        sCurPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".")
        Set xl = CreateObject("Excel.application")
        Set xlBook = xl.Workbooks.Open(sCurPath & "\ttlr_nvm_map.xlsm", 0, True)   
        xl.Application.Visible = False
        xl.Application.run "ttlr_nvm_map.xlsm!Module1.genNvmCMDline"
        xl.DisplayAlerts = False        
        xlBook.Save = True
        xl.activewindow.close
        xl.Quit
        Set xlBook = Nothing
        Set xl = Nothing
        End Sub 
