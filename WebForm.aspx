<%@ Page Language="VB" %>

<!DOCTYPE html>

<%

ON ERROR RESUME NEXT

'Capture E-mail Address of logged in user.

' Read E-mail Text Response    
     Const ForReading = 1, ForWriting = 2, ForAppending = 8
     Dim fso, f
     fso = CreateObject("Scripting.FileSystemObject")
     'Open the file for reading
     f = fso.OpenTextFile("\\frm-web\eic\messagebody_newemployee.txt", ForReading)
     'The ReadAll method reads the entire file into the variable BodyText'
     Dim BodyText
     BodyText = f.ReadAll
     'Close the file
     f.Close
     f = Nothing
     fso = Nothing
     
 
'Set Variablels Information to write to file
     Dim Name=request.form("Employee_Name")
     Dim PrefName=request.form("Preferred_Name")
     Dim Loc=request.form("Location")
     Dim Pos=request.form("Position")
     Dim Paidby=request.form("Paidby")
     Dim StaffExplain=request.form("Staff_Explanation")
     Dim AttyAsst=request.form("Assistant")
     Dim AttyBill=request.form("Bill_Rate")
     Dim RetainBill=request.form("Retain_Bill")
     Dim PracGrp=request.form("Practice_Group")
     Dim AttyAssign=request.form("Attorney_Assignments")
     Dim StartDate=request.form("Start_Date")
     Dim EndDate=request.form("End_Date")
     Dim Phone=request.form("Phone_Number")

     Dim Comments = request.form("Comments")
     ' Test for line breaks in variable
        Dim varInfo = Comments
        If IsDBNull (varInfo) Then
		   varInfo = varInfo
        Else
	  	   Comments = replace(varInfo,vbcrlf,"<BR>")
        End if
         
     Dim Computer=request.form("Computer")
     Dim CompExist=request.form("Previous_User")
     Dim AAP=request.form("AAP")
     Dim EXECCOMM=request.form("Exec_Comm")
     Dim Immigration=request.form("Immigration")
     Dim MgrMember=request.form("Managing_Members")
     Dim OSHA=request.form("OSHA")
     Dim StratPlan=request.form("Strategic_Comm")
     Dim WageHour=request.form("Wage_Hour")
     Dim ERISA=request.form("ERISA")
     Dim Health=request.form("Healthcare")
     Dim Litig=request.form("Litigation")
     Dim OffHead=request.form("Office_Heads")
     Dim PICCOMM=request.form("PIC_Comm")
     Dim Traditional=request.form("Traditional")
     Dim WorkComp=request.form("Workers_Comp")
     
'    response.write Paidby

' Subsitute Fields values for place holders in BodyText

  BodyText=Replace(BodyText, "@@@Name@@@", Name)
  BodyText=Replace(BodyText, "@@@PrefName@@@", PrefName)
  BodyText=Replace(BodyText, "@@@Loc@@@", Loc)
  BodyText=Replace(BodyText, "@@@Pos@@@", Pos)
  
 If StaffExplain="" Then
       BodyText=Replace(BodyText, "@@@StaffExplain@@@", "")
  Else
	   BodyText=Replace(BodyText, "@@@StaffExplain@@@", "<b>Staff Explanation: </b>"+StaffExplain)	
  End if
  
 If Paidby="" Then
       BodyText=Replace(BodyText, "@@@Paidby@@@", "")
  Else
	   BodyText=Replace(BodyText, "@@@Paidby@@@", "<b>Paid by: </b>"+Paidby)	
  End if
 
  
  If AttyAsst="" Then
       BodyText=Replace(BodyText, "@@@AttyAsst@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@AttyAsst@@@", "<b>Assistant: </b>"+AttyAsst)
  End if
  
  If AttyBill="" Then
       BodyText=Replace(BodyText, "@@@AttyBill@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@AttyBill@@@", "<b>Standard Billing Rate: </b>"+AttyBill)
  End if
  
  If RetainBill="" Then
       BodyText=Replace(BodyText, "@@@RetainBill@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@RetainBill@@@", "<b>Retainer Billing Rate: </b>"+RetainBill)
  End if

  If PracGrp="SELECT PRACTICE GROUP" Then
       BodyText=Replace(BodyText, "@@@PracGrp@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@PracGrp@@@", "<b>Practice Group: </b>"+PracGrp)
  End if
  
  If AttyAssign="" Then
       BodyText=Replace(BodyText, "@@@AttyAssign@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@AttyAssign@@@", "<b>Attorney Assignment(s): </b>"+AttyAssign)
  End if
  
  BodyText=Replace(BodyText, "@@@Start@@@", StartDate)
  
  If EndDate="" Then
       BodyText=Replace(BodyText, "@@@End@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@End@@@", "<b>End Date: </b>"+EndDate)
  End if
  
  BodyText=Replace(BodyText, "@@@Phone@@@", Phone)
    
   If len(Comments)=0 Then
       BodyText=Replace(BodyText, "@@@Comments@@@", "<b>Comments:</b> No comments entered.<br>")
   Else
      BodyText=Replace(BodyText, "@@@Comments@@@", "<b>Comments:</b><br>"+Comments)
   End if
   
  BodyText=Replace(BodyText, "@@@Computer@@@", Computer)
  
  If CompExist="" then
       BodyText=Replace(BodyText, "@@@CompExist@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@CompExist@@@", "<b>Previous User: </b>"+CompExist)
  End if
  
 
  IF AAP="" or EXECCOMM="" or Immigration="" or MgrMember="" or OSHA="" or StratPlan="" or WageHour="" or ERISA="" or Health="" or Litig="" or OffHead="" or PICCOMM="" or Traditional="" or WorkComp=""  then
       BodyText=Replace(BodyText, "@@@SPECEmail@@@", "<b>Special E-mail Groups: </b>")
  Else
       BodyText=Replace(BodyText, "@@@SPECEmail@@@", "")
  End if
  
  If AAP="" then
       BodyText=Replace(BodyText, "@@@AAP@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@AAP@@@", "AAP")
  End if
  
  If EXECCOMM="" then
        BodyText=Replace(BodyText, "@@@EXEC@@@<br>", "")       
  Else
       BodyText=Replace(BodyText, "@@@EXEC@@@", "Executive Committee")
  End if
  
  If Immigration="" then
       BodyText=Replace(BodyText, "@@@Immigration@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@Immigration@@@", "Immigration")
  End if
  
  If MgrMember="" then
      BodyText=Replace(BodyText, "@@@MgrMember@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@MgrMember@@@", "Managing Members")
  End if
  
  If OSHA="" then
       BodyText=Replace(BodyText, "@@@OSHA@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@OSHA@@@","OSHA")
  End if
  
  If StratPlan="" then
       BodyText=Replace(BodyText, "@@@StratPlan@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@StratPlan@@@", "Strategic Planning Committee")
  End if
  
  If WageHour="" then
       BodyText=Replace(BodyText, "@@@WageHour@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@WageHour@@@", "Wage Hour")
  End if
  
  If ERISA="" then
      BodyText=Replace(BodyText, "@@@ERISA@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@ERISA@@@", "ERISA")
  End if
  
  If Health="" then
       BodyText=Replace(BodyText, "@@@Health@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@Health@@@", "Healthcare")
  End if
  
  If Litig="" then
       BodyText=Replace(BodyText, "@@@Litig@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@Litig@@@", "Litigation")
  End if  

  If OffHead="" then
       BodyText=Replace(BodyText, "@@@OffHead@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@OffHead@@@", "Office Heads")
  End if
  
  If PICCOMM="" then
       BodyText=Replace(BodyText, "@@@PIC@@@<br>", "")
  Else
       BodyText=Replace(BodyText, "@@@PIC@@@", "PIC Committee")
  End if
  
  If Traditional="" then
       BodyText=Replace(BodyText, "@@@Traditional@@@<br>", "")
  Else
      BodyText=Replace(BodyText, "@@@Traditional@@@", "Traditional")
  End if
 
  If WorkComp="" then
       BodyText=Replace(BodyText, "@@@WorkComp@@@", "")
  Else
       BodyText=Replace(BodyText, "@@@WorkComp@@@", "Workers Compensation")
  End if
  

' Send by connecting to port 25 of the SMTP server.

Const cdoSendUsingPort = 2

Dim iMsg = CreateObject("CDO.Message")
Dim iConf = CreateObject("CDO.Configuration")

Dim Flds = iConf.Fields

' Set the CDOSYS configuration fields to use port 25 on the SMTP server.

With Flds
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
'Replace 0.0.0.0 with the IP address of your email server
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "192.168.4.18" 
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10  
    .Update
End With

' Email New Employee Information to emp-data (Key staff)

Dim strEmailTO = "emp-data@constangy.com"
' strEmailTO = "noc@constangy.com"
Dim strSubject = "New Employee Notification - "+Loc

' Apply the settings to the message.
With iMsg
    .Configuration = iConf
    .To = strEmailTO
    .Cc = strEmailCC
    .From = EmailAddress
    .Subject = strSubject
    .HTMLBody = BodyText
    .Send
End With

' Email New Employee Information to hr-emp-data (HR Director/HR Assistant)

Dim strEmailTO = "hr-emp-data@constangy.com"
' strEmailTO = "noc@constangy.com"
Dim strSubject = "New Employee Notification - "+Loc

' Apply the settings to the message.
With iMsg
    .Configuration = iConf
    .To = strEmailTO
    .Cc = strEmailCC
    .From = EmailAddress
    .Subject = strSubject
    .HTMLBody = BodyText
    .Send
End With

' Email New Employee Information to Help

strEmailTO = "help@constangy.com"
' strEmailTO = "noc@constangy.com"
strSubject = "New Employee Notification - "+Loc

' Apply the settings to the message.
With iMsg
    .Configuration = iConf
    .To = strEmailTO
    .Cc = strEmailCC
    .From = EmailAddress
    .Subject = strSubject
    .HTMLBody = BodyText
    .Send
End With

' Email New Employee Information to Help to add user to author's .mdb

'strEmailTO = "help@constangy.com"
' strEmailTO = "noc@constangy.com"
'strSubject = "Add user to Author's Database - " + Name

' Apply the settings to the message.
'With iMsg
    'Set .Configuration = iConf
    '.To = strEmailTO
    '.Cc = strEmailCC
    '.From = EmailAddress
    '.Subject = strSubject
    '.HTMLBody = BodyText
    '.Send
'End With


' Email Diveristy Chair notice of new attorney

If Pos="Managing Partner" or Pos="Partner" or Pos="Associate" or Pos="Of Counsel" then
     strSubject = "New Attorney Notification "+Name+" in "+Loc
      ' Apply the settings to the message.
     With iMsg
         .Configuration = iConf
         .To = "mzabijaka@constangy.com"
         .From = "help@constangy.com"
         .Subject = strSubject
         .HTMLBody = BodyText
        .Send
     End With
End if

' Email Associate Liaison Committee notice of new associate

If Pos="Associate" or Pos="Of Counsel" then
     strSubject = "New Associate Notification "+Name+" in "+Loc
      ' Apply the settings to the message.
     With iMsg
         .Configuration = iConf
         .To = "assoc-liaison@constangy.com"
         .From = "help@constangy.com"
         .Subject = strSubject
         .HTMLBody = BodyText
         .Send
     End With
End if

' Email request for Teleconferencing Card

'If Pos="Managing Partner" or Pos="Partner" or Pos="Associate" or Pos="Of Counsel" then
     'strSubject = "Order Teleconferencing Card for "+Name+" in "+Loc
     ' Apply the settings to the message.
     'With iMsg
         'Set .Configuration = iConf
         '.To = "help@constangy.com"
         '.From = "help@constangy.com"
         '.Subject = strSubject
         '.HTMLBody = BodyText
         '.Send
     'End With
'End if

' Email for Training

'strSubject = "Setup New User Orientation/Training for "+Name+" in "+Loc

' Apply the settings to the message.
'With iMsg
    'Set .Configuration = iConf
    '.To = "noc@constangy.com"
    '.To = "help@constangy.com"
    '.From = "help@constangy.com"
    '.Subject = strSubject
    '.HTMLBody = BodyText
    '.Send
'End With

     ' Clean up variables.
     iMsg = Nothing
     iConf = Nothing
     Flds = Nothing
%>
 
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Constangy New Employee Form</title>
</head>

<body>

<table border="0" style="border-collapse: collapse" width="650" id="table1" cellpadding="5">
	<tr>
		<td width="55">
		<img border="0" src="images/ok-icon.png" width="128" height="128"></td>
		<td width="575">
		<img border="0" src="images/Logo.jpg" width="250" height="72"></td>
	</tr>
	<tr>
		<td colspan="2">
<p align="left"><b><font face="Arial" color="#398AB1">Thank You!</font></b></p>
<p align="left"><b><font face="Arial" color="#398AB1">The new employee information has been 
sent to the appropriate personnel in Human Resources, Marketing Accounting and 
Information Technology.</font></b></p>
<p align="left"><b><font face="Arial" color="#398AB1">If you have any questions or require any 
additional assistance, <br>
please contact the help desk support line at</font></b></p>
<p align="left"><b><font face="Arial" size="6" color="#398AB1">(404) 230-6711</font><font face="Arial" color="#398AB1"><br>
&nbsp;<br>
or E-mail the help desk at<br>
<br>
&nbsp;</font><font face="Arial" size="6" color="#398AB1">HELP</font></b></p>

		</td>
	</tr>
</table>

</body>

</html>
