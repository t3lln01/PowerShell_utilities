sleep 1
Write-Host "`n`n"
sleep -Milliseconds 100
write-host "   *******             **    **                   **         ******             **                       **                     ********                                  **                      " -ForegroundColor Red #Red
sleep -Milliseconds 100
write-host "  **/////**           /**   /**                  /**        **////**           /**                      /**                    **//////                                  /**                      " -ForegroundColor Red #Green
sleep -Milliseconds 100
write-host " **     //** **   ** ****** /**  ******   ****** /**  **   **    //   ******   /**  *****  *******      /**  ******   ******  /**         *****   ******   ******  ***** /**       *****  ******  " -ForegroundColor Red #Yellow
sleep -Milliseconds 100
write-host "/**      /**/**  /**///**/  /** **////** **////**/** **   /**        //////**  /** **///**//**///**  ****** //////** //**//*  /********* **///** //////** //**//* **///**/******  **///**//**//*  " -ForegroundColor Red #Red
sleep -Milliseconds 100
write-host "/**      /**/**  /**  /**   /**/**   /**/**   /**/****    /**         *******  /**/******* /**  /** **///**  *******  /** /   ////////**/*******  *******  /** / /**  // /**///**/******* /** /   " -ForegroundColor Red #Cyan
sleep -Milliseconds 100
write-host "//**     ** /**  /**  /**   /**/**   /**/**   /**/**/**   //**    ** **////**  /**/**////  /**  /**/**  /** **////**  /**            /**/**////  **////**  /**   /**   **/**  /**/**////  /**     " -ForegroundColor Red #Magenta
sleep -Milliseconds 100
write-host " //*******  //******  //**  ***//****** //****** /**//**   //****** //******** ***//****** ***  /**//******//********/***      ******** //******//********/***   //***** /**  /**//******/***     " -ForegroundColor Red #Green
sleep -Milliseconds 100
write-host "  ///////    //////    //  ///  //////   //////  //  //     //////   //////// ///  ////// ///   //  //////  //////// ///      ////////   //////  //////// ///     /////  //   //  ////// ///      " -ForegroundColor Red #Red
Write-Host "`n"
write-host "Simone Filiaggi©" -ForegroundColor White
Write-Host "`n`n"

sleep 2

# Defines timeframe for the Outlook calendar search

while(1){
    Try{ [DateTime]$a = Read-Host "Select start date (ex: MM/dd/yyyy)"
    break
}
    catch{
    write-host "Not a Valid date! (ex: MM/dd/yyyy)" -ForegroundColor Red
  }
}


while(1){
    Try{ [DateTime]$b = Read-Host "Select month end date (ex: MM/dd/yyyy)"
      break
}
    catch{
    write-host "Not a Valid date! (ex: MM/dd/yyyy)" -ForegroundColor Red
  }
}

# Loads Outlook COM Object
    $outlook = New-Object -comobject Outlook.Application

# Finds current user email address
    $email = $outlook.Session.Accounts | Select -ExpandProperty SmtpAddress

# Finds user's calendar folder  
    $namespace = $outlook.GetNameSpace("MAPI")
    $folder = $namespace.GetDefaultFolder(9)

# Comment the $folder var line above and uncomment the 2 below in case DefaultFolder item is empty /  
#    $folders = $namespace.Folders.Item("$email")
#    $folder = $folders.Folders.Item("Calendar") 


# Finds items in calendar
$calendar = $folder.Items

# Finds meetings accepted within the timeframe provided
$items = $calendar | ? {( $_.Start -gt "$a") -and ( $_.Start -lt "$b" ) -and ( $_.ResponseStatus -eq "3" ) }
$meeting = $items | select -Property ConversationTopic, Start | ft -AutoSize
$list = $items.Count
Write-host `n
Write-Host "You have attended $list meeting(s) during the timeframe provided `n" -ForegroundColor Green; $meeting
