const minPercent = 35
const maxPercent = 75
const checkInterval = 300000 ' 5 minutes

set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
   iFull = oResult.FullChargedCapacity
next

while (1)
  set oResults = oServices.ExecQuery("select * from batterystatus")
  for each oResult in oResults
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining / iFull) * 100) mod 100
  if bCharging and (iPercent > maxPercent) Then msgbox "Battery is charged now more than " & maxPercent & "%. Please stop charging for optimal battery life.", _
	vbOKOnly + vbExclamation + vbSystemModal
  if not bCharging and (iPercent < minPercent) Then msgbox "Battery is discharging and is below " & minPercent & "%. Please switch on charging immediately.", _
	vbOKOnly + vbInformation + vbSystemModal
  wscript.sleep checkInterval
wend