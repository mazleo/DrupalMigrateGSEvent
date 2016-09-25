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
  Gui, Add, Text, vStatus, Copying data
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

    currentURL := login.LocationURL
    if (currentURL == CasURL) {
      ; Inputs NetID and password information
      login.document.getElementById("username").value := "jnl64"
      login.document.getElementById("password").value := "Jomarneon427"
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

  source := node.document.documentElement.outerHTML
  node.quit

  titlePos := InStr(source, "edit-title")
  StringTrimLeft, title, source, %titlePos%
  titlePos := InStr(title, "value=")
  StringTrimLeft, title, title, titlePos + 6
  titlePos := InStr(title, ">")
  titleLength := StrLen(title)
  StringTrimRight, title, title, titleLength - titlePos + 2

  categoryPos := InStr(source, "edit-field-event-category-0-value")
  StringTrimLeft category, source, %categoryPos%
  categoryPos := InStr(category, "value=")
  StringTrimLeft category, category, categoryPos + 6
  categoryPos := InStr(category, ">")
  categoryLength := StrLen(category)
  StringTrimRight, category, category, categoryLength - categoryPos + 2

  organizationPos := InStr(source, "edit-field-event-org-id-0-value")
  StringTrimLeft organization, source, %organizationPos%
  organizationPos := InStr(organization, "value=")
  StringTrimLeft organization, organization, organizationPos + 6
  organizationPos := InStr(organization, ">")
  organizationLength := StrLen(organization)
  StringTrimRight, organization, organization, organizationLength - organizationPos + 2

  eventPos := InStr(source, "edit-field-event-id-0-value")
  StringTrimLeft event, source, %eventPos%
  eventPos := InStr(event, "value=")
  StringTrimLeft event, event, eventPos + 6
  eventPos := InStr(event, ">")
  eventLength := StrLen(event)
  StringTrimRight, event, event, eventLength - eventPos + 2

  linkPos := InStr(source, "edit-field-feature-link-0-url")
  StringTrimLeft link, source, %linkPos%
  linkPos := InStr(link, "value=")
  StringTrimLeft link, link, linkPos + 6
  linkPos := InStr(link, ">")
  linkLength := StrLen(link)
  StringTrimRight, link, link, linkLength - linkPos + 2

  bodyPos := InStr(source, "form-textarea resizable ckeditor-mod  textarea-processed ckeditor-processed")
  StringTrimLeft body, source, %bodyPos%
  bodyPos := InStr(body, ">")
  StringTrimLeft body, body, bodyPos
  bodyPos := InStr(body, "</textarea>")
  bodyLength := StrLen(body)
  StringTrimRight, body, body, bodyLength - bodyPos + 1

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
    locationLen := StrLen(location)
    StringTrimRight, location, location, locationLen - locationPos - 5

    StringReplace, body, body, Location: %location%
    StringReplace, location, location, <br />
  }

  directionPos := InStr(body, "Directions:")
  if (directionPos > 0) {
    StringTrimLeft direction, body, directionPos + 11
    directionPos := InStr(direction, "<br />")
    directionLen := StrLen(direction)
    StringTrimRight, direction, direction, directionLen - directionPos - 5
    StringReplace, body, body, Directions: %direction%
    StringReplace, direction, direction, <br />
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

  ; COM Objects to extract HTML Source code from Drupal 6 view page
  node := ComObjCreate("InternetExplorer.Application")
  node.navigate[NodeURL]
  node.visible := True

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

    currentURL := login.LocationURL
    if (currentURL == CasURL) {
      ; Inputs NetID and password information
      login.document.getElementById("username").value := "jnl64"
      login.document.getElementById("password").value := "Jomarneon427"
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

  NodeURL := node.document.getElementsByClassName("tabs")[0].getElementsByTagName("li")[0].getElementsByTagName("a")[0].href

  node.navigate[NodeURL]
  while node.busy
    sleep 100

  source := node.document.documentElement.outerHTML

^!F10::
  ExitApp
