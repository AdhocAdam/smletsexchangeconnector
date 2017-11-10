//Optional - Hide "Send Outlook Meeting" from End Users
//app.custom.formTasks.add('Incident', null, function (formObj, viewModel) { formObj.boundReady(function () 
//{ if (!session.user.Analyst) { $( ".taskmenu li:contains('Send Outlook Meeting')" ).hide() } }); }); 

//Email address and Display name (that you'll see in the CC: of the email address) of your SCSM Workflow Account to receive the meeting
var scsmWFAccount = "account@domain.tld"
var scsmWFAccountDisplayName = "Service Manager"

//generate dates for appoint start/end time and the created time of the ICS message
var date = new Date(),
    year = date.getFullYear(),
    month = (date.getMonth() + 1).toString(),
    formattedMonth = (month.length === 1) ? ("0" + month) : month,
    day = date.getDate().toString(),
    formattedDay = (day.length === 1) ? ("0" + day) : day,
    hour = date.getHours().toString(),
    formattedHour = (hour.length === 1) ? ("0" + hour) : hour,
    minute = date.getMinutes().toString(),
    formattedMinute = (minute.length === 1) ? ("0" + minute) : minute,
    second = date.getSeconds().toString(),
    formattedSecond = (second.length === 1) ? ("0" + second) : second;
var icsCreateDate = year + formattedMonth + formattedDay + 'T' + formattedHour + formattedMinute + formattedSecond
var stateDate = year + formattedMonth + formattedDay
var endDate = year + formattedMonth + formattedDay

//Send Outlook Meeting, Incident
app.custom.formTasks.add('Incident', "Send Outlook Meeting", function (formObj, viewModel){ 
	console.log(formObj);
	
	var meetingSubject = "[" + pageForm.viewModel.Id + "]" + " " + pageForm.viewModel.Title
	var calMeetingInvitee = pageForm.viewModel.RequestedWorkItem.UPN
	
	//Create calendar meeting based on Affected User
	if (calMeetingInvitee)
	{
		var calMeetingInviteeDisplayName = pageForm.viewModel.RequestedWorkItem.DisplayName
		var icsMSG = "BEGIN:VCALENDAR\nVERSION:2.0\nMETHOD:PUBLISH\nBEGIN:VEVENT\nATTENDEE;CN="+ calMeetingInviteeDisplayName +";RSVP=TRUE:mailto:" + calMeetingInvitee + "\nATTENDEE;CN="+ scsmWFAccountDisplayName +";RSVP=TRUE:mailto:" + scsmWFAccount + "\nDTSTART:" + stateDate + "+\nDTEND:" + endDate + "\nTRANSP:OPAQUE\nSEQUENCE:0\nUID:" + pageForm.viewModel.Id + "\nDTSTAMP:" + icsCreateDate + "\nSUMMARY:" + meetingSubject + "\nPRIORITY:5\nCLASS:PUBLIC\nBEGIN:VALARM\nTRIGGER:PT15M\nACTION:DISPLAY\nDESCRIPTION:Reminder\nEND:VALARM\nEND:VEVENT\nEND:VCALENDAR"
	}
	else
	{
		var icsMSG = "BEGIN:VCALENDAR\nVERSION:2.0\nMETHOD:PUBLISH\nBEGIN:VEVENT\nATTENDEE;CN=" + scsmWFAccountDisplayName + ";RSVP=TRUE:mailto:" + scsmWFAccount +"\nDTSTART:" + stateDate + "\nDTEND:" + endDate + "\nTRANSP:OPAQUE\nSEQUENCE:0\nUID:" + pageForm.viewModel.Id +"\nDTSTAMP:" + icsCreateDate + "\nSUMMARY:"+ meetingSubject +"\nPRIORITY:5\nCLASS:PUBLIC\nBEGIN:VALARM\nTRIGGER:PT15M\nACTION:DISPLAY\nDESCRIPTION:Reminder\nEND:VALARM\nEND:VEVENT\nEND:VCALENDAR"
	}
	
	//Open the the ICS file
	window.open( "data:text/calendar;charset=utf8," + escape(icsMSG));
});