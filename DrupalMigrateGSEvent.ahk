^!d::
  ; Gets URL and Window data
  ; Separated so that GUI Progress bar does not activate before copying these, where these windows need to be active
  D7CreateEvent = Create Event | Graduate and Professional Admissions - Google Chrome
  CasURL := "https://cas.rutgers.edu/login?service=http%3A%2F%2Fgradstudy.rutgers.edu%2Flogin"
  WinGetActiveTitle D6Node
  Send ^l
  Clipboard := ""
  Send ^c
  NodeURL = %Clipboard%

  Send ^!c

  Exit

^!c::
  ; Initialize GUI Progress bar
  Gui, Add, Progress, w400 h20 -Smooth vProgressBar, 0
  Gui, Add, Text, vStatus, Migrating data
  Gui, Show, w420 h50 x100 y100
  GoSub, progBar
  Return

  progBar:
      GuiControl,, ProgressBar, 0

  ; COM Objects to extract HTML Source code from Drupal 6 edit site
  node := ComObjCreate("InternetExplorer.Application")
  node.navigate[NodeURL]
  node.visible := False

  while node.busy
    Sleep 100

  ; COM Object to login to server if not logged in
  login := ComObjCreate("InternetExplorer.Application")

  source := node.document.documentElement.outerHTML

  GuiControl,, ProgressBar, +10

  ; Checks if the correct NetID and Password are provided
  isDeniedPos := InStr(source, "Access Denied")
  nTimesDenied := 0
  while (isDeniedPos > 0 && nTimesDenied < 6) {
    login.navigate("http://gradstudy.rutgers.edu/login")
    login.visible := False

    while login.busy
      sleep 100

    ; Retrieves password information
    FileReadLine, netid, login.txt, 1
    FileReadLine, password, login.txt, 2

    currentURL := login.LocationURL
    if (currentURL == CasURL) {
      ; Inputs NetID and password information
      login.document.getElementById("username").value := netid
      login.document.getElementById("password").value := password
      login.document.getElementsByClassName("btn-submit")[0].click()
      while login.busy
        sleep 100
    }

    node.navigate[NodeURL]
    while node.busy
      Sleep 100

    source := node.document.documentElement.outerHTML

    if (nTimesDenied == 3 || nTimesDenied == 6) {
      MsgBox Unable to login to server. Please provide correct NetID and password

      InputBox, netidTemp, Credentials, NetID:,,, 120
      InputBox, passTemp, Credentials, Password:, HIDE,, 120

      IfExist, login.txt
      {
        FileDelete, login.txt
      }
      FileAppend, %netidTemp%`n%passTemp%, login.txt
    }
    isDeniedPos := InStr(source, "Access Denied")
    if (isDeniedPos > 0) {
      nTimesDenied += 1
    }
  }

  GuiControl,, ProgressBar, +10

  login.quit

  node.navigate[NodeURL]
  while node.busy
    Sleep 100

  title := ""
  category := ""
  organization := ""
  event := ""
  location := ""
  direction := ""
  body := ""
  link := ""

  title := node.document.getElementById("edit-title").value
  category := node.document.getElementById("edit-field-event-category-0-value").value
  organization := node.document.getElementById("edit-field-event-org-id-0-value").value
  event := node.document.getElementById("edit-field-event-id-0-value").value
  link := node.document.getElementById("edit-field-feature-link-0-url").value
  body := node.document.getElementsByTagName("textarea")[1].innerHTML

  GuiControl,, ProgressBar, +20

  startTagPos := InStr(body, "&lt;")
  while (startTagPos > 0) {
    StringReplace, body, body, &lt;, <
    startTagPos := InStr(body, "&lt;")
  }

  endTagPos := InStr(body, "&gt;")
  while (endTagPos > 0) {
    StringReplace, body, body, &gt;, >
    endTagPos := InStr(body, "&gt;")
  }

  ampTagPos := InStr(body, "&amp;")
  while (ampTagPos > 0) {
    StringReplace, body, body, &amp;, &
    ampTagPos := InStr(body, "&amp;")
  }

  GuiControl,, ProgressBar, +10

  locationPos := InStr(body, "Location:")
  if (locationPos > 0) {
    StringTrimLeft location, body, locationPos + 9
    locationPos := InStr(location, "<br />")
    locationPos2 := InStr(location, "<br>")

    if (locationPos > 0) {
      locationLen := StrLen(location)
      StringTrimRight, location, location, locationLen - locationPos - 5
      StringReplace, body, body, Location: %location%
      StringReplace, location, location, <br />
    } else if (locationPos2 > 0) {
      locationLen := StrLen(location)
      StringTrimRight, location, location, locationLen - locationPos - 3
      StringReplace, body, body, Location: %location%
      StringReplace, location, location, <br>
    } else {
      locationPos := InStr(location, "</p>")
      locationLen := StrLen(location)
      StringTrimRight, location, location, locationLen - locationPos + 1
      StringReplace, body, body, Location: %location%
    }
  }

  directionPos := InStr(body, "Directions:")
  if (directionPos > 0) {
    StringTrimLeft direction, body, directionPos + 11
    directionPos := InStr(direction, "<br />")
    directionPos2 := InStr(direction, "<br>")
    if (directionPos > 0) {
      directionLen := StrLen(direction)
      StringTrimRight, direction, direction, directionLen - directionPos - 5
      StringReplace, body, body, Directions: %direction%
      StringReplace, direction, direction, <br />
    } else if (directionPos2 > 0) {
      directionLen := StrLen(direction)
      StringTrimRight, direction, direction, directionLen - directionPos - 3
      StringReplace, body, body, Direction: %direction%
      StringReplace, direction, direction, <br>
    } else {
      directionPos := InStr(direction, "</p>")
      directionLen := StrLen(direction)
      StringTrimRight, direction, direction, directionLen - directionPos + 1
      StringReplace, body, body, Directions: %direction%
    }
  }

  GuiControl,, ProgressBar, +10

  ; Extracts date data from page view

  NodeURL := node.document.getElementsByClassName("tabs")[0].getElementsByTagName("li")[0].getElementsByTagName("a")[0].href
  node.quit

  GuiControl,, ProgressBar, +10

  node := ComObjCreate("InternetExplorer.Application")
  node.navigate[NodeURL]
  while node.busy
    sleep 1000
  node.visible := False

  source := node.document.documentElement.outerHTML

  singleDate := ""
  startDate := ""
  endDate := ""
  fromDate := ""
  toDate := ""
  fromTime := ""
  toTime := ""
  isSeparator := ""

  singleDatePos := InStr(source, "date-display-single")
  startDatePos := InStr(source, "date-display-start")
  endDatePos := InStr(source, "date-display-end")
  Comma := ","

  GuiControl,, ProgressBar, +10

  if (singleDatePos > 0) {
    singleDate := node.document.getElementsByClassName("date-display-single")[0].innerHTML
    isSeparator := InStr(singleDate, "|")
    StringReplace, singleDate, singleDate, %A_Space%,, All
    StringReplace, singleDate, singleDate, %Comma%, /, All
  }
  if (startDatePos > 0) {
    startDate := node.document.getElementsByClassName("date-display-start")[0].innerHTML
    StringReplace, startDate, startDate, %A_Space%,, All
    StringReplace, startDate, startDate, %Comma%, /, All
  }
  if (endDatePos > 0) {
    endDate := node.document.getElementsByClassName("date-display-end")[0].innerHTML
    StringReplace, endDate, endDate, %A_Space%,, All
    StringReplace, endDate, endDate, %Comma%, /, All
  }

  ; Both start and end dates
  if (singleDatePos == 0 && startDatePos > 0 && endDatePos > 0) {
    fromDate := ConvertDate(startDate)
    toDate := ConvertDate(endDate)
    fromTime := "12:00am"
    toTime := "12:00am"
  }
  ; One date and one start time
  else if (singleDatePos > 0 && isSeparator > 0 && endDatePos == 0) {
    singleDate := ConvertDate(singleDate)
    startTime := singleDate
    fromDate := SubStr(singleDate, 1, 10)
    StringTrimLeft, fromTime, startTime, 11
    fromTimeLen := StrLen(fromTime)
    if (fromTimeLen <> 7) {
      fromTime := "0" . fromTime
    }
  }
  ; One date and both start and end times
  else if (singleDatePos > 0 && startDatePos > 0 && endDatePos > 0) {
    singleDate := ConvertDate(singleDate)
    fromDate := SubStr(singleDate, 1, 10)
    toDate := fromDate
    startDateLen := StrLen(startDate)
    endDateLen := StrLen(endDate)
    if (startDateLen <> 7) {
      fromTime := "0" . startDate
    } else {
      fromTime := startDate
    }
    if (endDateLen <> 7) {
      toTime := "0" . endDate
    } else {
      toTime := endDate
    }
  }
  ; One date only
  else if (singleDatePos > 0 && isSeparator == 0 && startDatePos == 0 && endDatePos == 0) {
    fromDate := ConvertDate(singleDate)
    fromTime := "12:00am"
  }

  GuiControl,, ProgressBar, +10

  ; Copy data to Drupal 7 fields
  WinActivate %D7CreateEvent%
  WinWaitActive %D7CreateEvent%

  Send ^l
  Sleep 100
  Send javascript:
  Clipboard := "javascript: document.getElementById('switch_edit-body-und-0-value').click();"
  Send ^v
  Sleep 100
  Send {Enter}
  Sleep 100

  Send ^l
  Sleep 100
  Send javascript:
  Clipboard :="document.getElementById('edit-title').value = '" . title . "'; document.getElementById('edit-field-event-date-und-0-value-datepicker-popup-0').value = '" . fromDate . "'; document.getElementById('edit-field-event-date-und-0-value-timeEntry-popup-1').value = '" . fromTime . "'; document.getElementById('edit-field-event-date-und-0-value2-datepicker-popup-0').value = '" . toDate . "'; document.getElementById('edit-field-event-date-und-0-value2-timeEntry-popup-1').value = '" . toTime . "'; document.getElementById('edit-field-category-und-0-value').value = '" . category . "'; document.getElementById('edit-field-organization-id-und-0-value').value = '" . organization . "'; document.getElementById('edit-field-event-id-und-0-value').value = '" . event . "'; document.getElementById('edit-field-event-location-und-0-value').value = '" . location . "'; document.getElementById('edit-field-event-direction-und-0-value').value = '" . direction . "'; document.getElementById('edit-body-und-0-value').value = '" . body . "'; document.getElementById('edit-field-featured-link-und-0-title').value = '" . title . "'; document.getElementById('edit-field-featured-link-und-0-url').value = '" . link . "'; var temp = {title: 'Drupal', url: 'http://gradstudy.p7.drupaldev.rutgers.edu/node/add/event'}; history.pushState(temp.title, temp.url);"
  Send ^v
  Sleep 100
  Send {Enter}
  Sleep 100

  Send ^l
  Sleep 100
  Send javascript:
  Clipboard := "javascript: document.getElementById('switch_edit-body-und-0-value').click();"
  Send ^v
  Sleep 100
  Send {Enter}


  GuiControl,, ProgressBar, +10

  ConvertDate(date)
  {
    IfInString, date, january
    {
      StringReplace, date, date, january, 01/
    }
    IfInString, date, february
    {
      StringReplace, date, date, february, 02/
    }
    IfInString, date, march
    {
      StringReplace, date, date, march, 03/
    }
    IfInString, date, april
    {
      StringReplace, date, date, april, 04/
    }
    IfInString, date, may
    {
      StringReplace, date, date, may, 05/
    }
    IfInString, date, june
    {
      StringReplace, date, date, june, 06/
    }
    IfInString, date, july
    {
      StringReplace, date, date, july, 07/
    }
    IfInString, date, august
    {
      StringReplace, date, date, august, 08/
    }
    IfInString, date, september
    {
      StringReplace, date, date, september, 09/
    }
    IfInString, date, october
    {
      StringReplace, date, date, october, 10/
    }
    IfInString, date, november
    {
      StringReplace, date, date, november, 11/
    }
    IfInString, date, december
    {
      StringReplace, date, date, december, 12/
    }
    date := correctDate(date)
    return date
  }
  correctDate(date) {
    tempDate := SubStr(date, 4)
    sepPos := InStr(tempDate, "/")
    tempDateLen := StrLen(tempDate)
    StringTrimRight, tempDate, tempDate, tempDateLen - sepPos + 1
    tempDateLen := StrLen(tempDate)
    slashTempDate := "/" . tempDate
    if (tempDateLen < 2) {
      date := RegExReplace(date, slashTempDate, "/0"tempDate)
    }
    return date
  }

  ;For testing purposes
  ;Clipboard := body
  ;MsgBox Title: %title%
  ;MsgBox Category: %category%
  ;MsgBox Organization: %organization%
  ;MsgBox Event: %event%
  ;MsgBox Location: %location%
  ;MsgBox Direction: %direction%
  ;MsgBox Body: %body%
  ;MsgBox Link: %link%

  Gui, Destroy
  node.quit
  Exit

^!F10::
  ExitApp
