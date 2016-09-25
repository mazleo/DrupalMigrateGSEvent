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
  Gui, Show, w420 h50
  GoSub, progBar
  Return

  progBar:
      GuiControl,, ProgressBar, 0

  GuiControl,, ProgressBar, +50

  ; COM Objects to extract HTML Source code from Drupal 6 edit site
  node := ComObjCreate("InternetExplorer.Application")
  node.navigate[NodeURL]
  node.visible := False

  while node.busy
    Sleep 100

  ; COM Object to login to server if not logged in
  login := ComObjCreate("InternetExplorer.Application")

  source := node.document.documentElement.outerHTML

  ; Checks if the correct NetID and Password are provided
  isDeniedPos := InStr(source, "Access Denied")
  nTimesDenied := 0
  while (isDeniedPos > 0 && nTimesDenied < 3) {
    login.navigate("http://gradstudy.rutgers.edu/login")
    login.visible := False

    while login.busy
      sleep 100

    ; Retrieves password information
    FileReadLine, netid, data/login.txt, 1
    FileReadLine, password, data/login.txt, 2

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

    isDeniedPos := InStr(source, "Access Denied")
    if (isDeniedPos > 0) {
      nTimesDenied += 1
    }
    if (nTimesDenied == 3) {
      MsgBox Unable to login to server. Please provide correct NetID and password
      login.quit
      node.quit
      ExitApp
    }
  }

  login.quit

  node.navigate[NodeURL]
  while node.busy
    Sleep 100

  title := node.document.getElementById("edit-title").value
  category := node.document.getElementById("edit-field-event-category-0-value").value
  organization := node.document.getElementById("edit-field-event-org-id-0-value").value
  event := node.document.getElementById("edit-field-event-id-0-value").value
  link := node.document.getElementById("edit-field-feature-link-0-url").value
  body := node.document.getElementsByTagName("textarea")[1].innerHTML

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

  locationPos := InStr(body, "Location:")
  if (locationPos > 0) {
    StringTrimLeft location, body, locationPos + 9
    locationPos := InStr(location, "<br />")
    if (locationPos > 0) {
      locationLen := StrLen(location)
      StringTrimRight, location, location, locationLen - locationPos - 5
      StringReplace, body, body, Location: %location%
      StringReplace, location, location, <br />
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
    if (directionPos > 0) {
      directionLen := StrLen(direction)
      StringTrimRight, direction, direction, directionLen - directionPos - 5
      StringReplace, body, body, Directions: %direction%
      StringReplace, direction, direction, <br />
    } else {
      directionPos := InStr(direction, "</p>")
      directionLen := StrLen(direction)
      StringTrimRight, direction, direction, directionLen - directionPos + 1
      StringReplace, body, body, Directions: %direction%
    }
  }

  ;Clipboard := body
  ;MsgBox Title: %title%
  ;MsgBox Category: %category%
  ;MsgBox Organization: %organization%
  ;MsgBox Event: %event%
  ;MsgBox Link: %link%
  ;MsgBox Location: %location%
  ;MsgBox Directions: %direction%
  ;MsgBox Body: %body%

  ;---------------------------------------------------------------------------

  NodeURL := node.document.getElementsByClassName("tabs")[0].getElementsByTagName("li")[0].getElementsByTagName("a")[0].href
  node.quit

  node := ComObjCreate("InternetExplorer.Application")
  node.navigate[NodeURL]
  while node.busy
    sleep 1000
  node.visible := True

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

  if (singleDatePos > 0) {
    singleDate := node.document.getElementsByClassName("date-display-single")[0].innerHTML
    isSeparator := InStr(singleDate, "|")
    IfInString, singleDate, ", "
    {
      StringReplace, singleDate, singleDate, ", "
      StringReplace, singleDate, singleDate, " "
    }
  }
  if (startDatePos > 0) {
    startDate := node.document.getElementsByClassName("date-display-start")[0].innerHTML
    IfInString, startDate, ", "
    {
      StringReplace, startDate, startDate, ", "
      StringReplace, startDate, startDate, " "
    }
  }
  if (endDatePos > 0) {
    endDate := node.document.getElementsByClassName("date-display-end")[0].innerHTML
    IfInString, endDate, ", "
    {
      StringReplace, endDate, endDate, ", "
      StringReplace, endDate, endDate, " "
    }
  }



  ; Both start and end dates
  if (singleDatePos == 0 && startDatePos > 0 && endDatePos > 0) {
    MsgBox % startDate
    MsgBox % endDate
  }
  ; One date and one start time
  else if (singleDatePos > 0 && isSeparator > 0 && endDatePos == 0) {
    MsgBox % singleDate
  }
  ; One date and both start and end times
  else if (singleDatePos > 0 && startDatePos > 0 && endDatePos > 0) {
    MsgBox % singleDate
    MsgBox % startDate
    MsgBox % endDate
  }
  ; One date only
  else if (singleDatePos > 0 && isSeparator == 0 && startDatePos == 0 && endDatePos == 0) {
    MsgBox % singleDate
  }

  node.quit

^!F10::
  ExitApp
