import re
import pycurl
import StringIO
from flask import Flask
from flask import request

#
# Created by Porter Liu on 01/02/2015 in PHP
# Rewritten by Porter Liu on 06/01/2019 in Python
#
# This tool requires Flask and PycURL. I only tested on macOS. Install PycURL on macOS:
# PYCURL_SSL_LIBRARY=openssl LDFLAGS="-L/usr/local/opt/openssl/lib" CPPFLAGS="-I/usr/local/opt/openssl/include" pip install --no-cache-dir --user pycurl
#
# How do you run it:
# FLASK_APP=ews.py flask run --host=0.0.0.0
#

html_body = r'''
<html>
<header>
	<style>
		*
		{
			font-family: "Source Code Pro";
			font-size: 10pt;
		}
	</style>
	<script type="text/javascript">
	if( !String.prototype.format )
	{
		String.prototype.format = function()
		{
			var args = arguments;
			return this.replace( /{(\d+)}/g, function( match, number )
			{ 
				return typeof args[number] != 'undefined' ? args[number] : match;
			});
		};
	}
	if( !Date.prototype.format )
	{
		Date.prototype.format = function(format)
		{
			var date = {
				"M+": this.getMonth() + 1,
				"d+": this.getDate(),
				"h+": this.getHours(),
				"m+": this.getMinutes(),
				"s+": this.getSeconds(),
				"q+": Math.floor( (this.getMonth() + 3) / 3 ),
				"S+": this.getMilliseconds()
			};
			if( /(y+)/i.test( format ) )
			{
				format = format.replace( RegExp.$1, ( this.getFullYear() + '' ).substr( 4 - RegExp.$1.length ) );
			}
			for( var k in date )
			{
				if( new RegExp( "(" + k + ")" ).test( format ) )
				{
					format = format.replace( RegExp.$1, RegExp.$1.length == 1 ? date[k] : ("00" + date[k]).substr( ("" + date[k]).length ) );
				}
			}
			return format;
		}
	}
	function ewsUrlOffice365()
	{
		document.getElementById( "ewsUrl" ).value = "https://outlook.office365.com/ews/exchange.asmx";
	}
	function ewsUrlChinaOffice365()
	{
		document.getElementById( "ewsUrl" ).value = "https://partner.outlook.cn/EWS/Exchange.asmx";
	}
	function xmlFindItem()
	{
		var xml = '<?xml version="1.0" encoding="utf-8"?>\n\
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"\n\
	xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">\n\
	<soap:Body>\n\
		<FindItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"\n\
			xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"\n\
			Traversal="Shallow">\n\
			<ItemShape>\n\
				<t:BaseShape>IdOnly</t:BaseShape>\n\
			</ItemShape>\n\
			<CalendarView MaxEntriesReturned="500" StartDate="{0}" EndDate="{1}"/>\n\
			<ParentFolderIds>\n\
				<t:DistinguishedFolderId Id="calendar">\n\
					<t:Mailbox>\n\
						<t:EmailAddress>{2}</t:EmailAddress>\n\
					</t:Mailbox>\n\
				</t:DistinguishedFolderId>\n\
			</ParentFolderIds>\n\
		</FindItem>\n\
	</soap:Body>\n\
</soap:Envelope>';
		var now = new Date();
		var hoursOffset = now.getTimezoneOffset() / -60;

		var today = new Date().format( 'yyyy-MM-dd' );
		var startDate = today + "T00:00:00" + ( hoursOffset > 0 ? "+" : "-" ) + ( Math.abs( hoursOffset ) > 9 ? "" : "0" ) + Math.abs( hoursOffset ) + ":00";
		var endDate   = today + "T23:59:59" + ( hoursOffset > 0 ? "+" : "-" ) + ( Math.abs( hoursOffset ) > 9 ? "" : "0" ) + Math.abs( hoursOffset ) + ":00";
		var roomEmail = document.getElementById( "room_email" ).value;
		if( roomEmail.length == 0 )
			roomEmail = "__________@__________.onmicrosoft.com";

		document.getElementById( "xml" ).value = xml.format( startDate, endDate, roomEmail );
		document.getElementById( "requestType" ).value = "FindItem";
	}
	function xmlGetItem()
	{
		document.getElementById( "xml" ).value = '<?xml version="1.0" encoding="utf-8"?>\n\
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"\n\
	xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">\n\
	<soap:Body>\n\
		<GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">\n\
			<ItemShape>\n\
				<t:BaseShape>IdOnly</t:BaseShape>\n\
				<t:BodyType>Text</t:BodyType>\n\
				<t:AdditionalProperties>\n\
					<t:FieldURI FieldURI="item:Subject"/>\n\
					<t:FieldURI FieldURI="item:Body"/>\n\
					<t:FieldURI FieldURI="calendar:Start"/>\n\
					<t:FieldURI FieldURI="calendar:End"/>\n\
					<t:FieldURI FieldURI="calendar:Organizer"/>\n\
					<t:FieldURI FieldURI="calendar:Location"/>\n\
					<t:FieldURI FieldURI="item:Sensitivity"/>\n\
				</t:AdditionalProperties>\n\
			</ItemShape>\n\
			<ItemIds>\n\
				<t:ItemId Id="__________" ChangeKey="__________"/>\n\
			</ItemIds>\n\
		</GetItem>\n\
	</soap:Body>\n\
</soap:Envelope>';
		document.getElementById( "requestType" ).value = "GetItem";
	}
	function xmlCreateItem()
	{
		var xml = '<?xml version="1.0" encoding="utf-8"?>\n\
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"\n\
	xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"\n\
	xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"\n\
	xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">\n\
	<soap:Body>\n\
		<m:CreateItem SendMeetingInvitations="SendToAllAndSaveCopy">\n\
			<m:Items>\n\
				<t:CalendarItem>\n\
					<t:Subject>__________</t:Subject>\n\
					<t:Body BodyType="Text">__________</t:Body>\n\
					<t:ReminderIsSet>true</t:ReminderIsSet>\n\
					<t:ReminderMinutesBeforeStart>15</t:ReminderMinutesBeforeStart>\n\
					<t:Start>{0}</t:Start>\n\
					<t:End>{1}</t:End>\n\
					<t:RequiredAttendees>\n\
						<t:Attendee>\n\
							<t:Mailbox>\n\
								<t:EmailAddress>{2}</t:EmailAddress>\n\
							</t:Mailbox>\n\
						</t:Attendee>\n\
					</t:RequiredAttendees>\n\
				</t:CalendarItem>\n\
			</m:Items>\n\
		</m:CreateItem>\n\
	</soap:Body>\n\
</soap:Envelope>';
		var now = new Date();
		var hoursOffset = now.getTimezoneOffset() / -60;

		var today = new Date().format( 'yyyy-MM-dd' );
		var startTime = new Date().format( 'hh:mm:00' );
		var endTime   = new Date( new Date().getTime() + 30 * 60000 ).format( 'hh:mm:00' );
		var startDate = today + "T" + startTime + ( hoursOffset > 0 ? "+" : "-" ) + ( Math.abs( hoursOffset ) > 9 ? "" : "0" ) + Math.abs( hoursOffset ) + ":00";
		var endDate   = today + "T" + endTime   + ( hoursOffset > 0 ? "+" : "-" ) + ( Math.abs( hoursOffset ) > 9 ? "" : "0" ) + Math.abs( hoursOffset ) + ":00";
		var roomEmail = document.getElementById( "room_email" ).value;
		if( roomEmail.length == 0 )
			roomEmail = "__________@__________.onmicrosoft.com";

		document.getElementById( "xml" ).value = xml.format( startDate, endDate, roomEmail );
		document.getElementById( "requestType" ).value = "CreateItem";
	}
	function xmlDeleteItem()
	{
		document.getElementById( "xml" ).value = '<?xml version="1.0" encoding="utf-8"?>\n\
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"\n\
	xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">\n\
	<soap:Body>\n\
		<DeleteItem DeleteType="HardDelete" SendMeetingCancellations="SendOnlyToAll" AffectedTaskOccurrences="SpecifiedOccurrenceOnly" xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">\n\
			<ItemIds>\n\
				<t:ItemId Id="__________" ChangeKey="__________"/>\n\
			</ItemIds>\n\
		</DeleteItem>\n\
	</soap:Body>\n\
</soap:Envelope>';
		document.getElementById( "requestType" ).value = "DeleteItem";
	}
	function onRoomEmailChange( roomEmail )
	{
		var requestType = document.getElementById( "requestType" ).value;
		if( requestType == "FindItem" && /<FindItem/.test( document.getElementById( "xml" ).value ) )
		{
			// if user clicked FindItem and the request body is still FindItem, let's update the roomEmail
			xmlFindItem();
		}
		else if( requestType == "CreateItem" && /<m:CreateItem/.test( document.getElementById( "xml" ).value ) )
		{
			// if user clicked CreateItem and the request body is still CreateItem, let's update the roomEmail
			xmlCreateItem();
		}
	}
	</script>
</header>

<body>
	<form method="post">
		<input type="hidden" id="requestType" name="request_type" value="_REQUEST_TYPE_"/>
		<p>
			<b>EWS URL:</b> <a href="#" onclick="ewsUrlOffice365();">Office 365</a>&nbsp;&nbsp;<a href="#" onclick="ewsUrlChinaOffice365();">China Office 365</a><br/>
			<input name="url" id="ewsUrl" type="text" value="_URL_" style="width:100%"/>
		</p>

		<p>
			<b>Username/Password:</b> (test@abc.com:1234 or domain\test:1234)<br/>
			<input name="username_password" type="text" value="_USERNAME_PASSWORD_" style="width:100%"/>
		</p>

		<p>
			<b>Room email:</b><br/>
			<input id="room_email" name="room_email" type="text" value="_ROOM_EMAIL_" style="width:100%" onkeyup="return onRoomEmailChange( this.value );"/>
		</p>

		<b>XML:</b>&nbsp;&nbsp;<a href="#" onclick="xmlFindItem();">FindItem</a>&nbsp;&nbsp;<a href="#" onclick="xmlGetItem();">GetItem</a>&nbsp;&nbsp;<a href="#" onclick="xmlCreateItem();">CreateItem</a>&nbsp;&nbsp;<a href="#" onclick="xmlDeleteItem();">DeleteItem</a>
		<textarea name="xml" id="xml" style="width:100%; height:300px">_XML_</textarea>
		<input type="submit" name="doit" value="    Test    " style="background-color:green; color:white"/>&nbsp;
		<input type="checkbox" name="ntlm" value="1"_NTLM_/>NTLM Authentication
	</form>

    <textarea name="postdata" style="width:100%; height:300px">_CONTENT_</textarea>
</body>
</html>
'''

app = Flask( __name__ )

@app.route( '/', methods = ['GET', 'POST'] )
def home():
    body = html_body

    request_type      = request.form.get( 'request_type', '' )
    url               = request.form.get( 'url', '' )
    username_password = request.form.get( 'username_password', '' )
    room_email        = request.form.get( 'room_email', '__________@__________.onmicrosoft.com' )
    xml               = request.form.get( 'xml', '' )
    ntlm              = ' checked' if request.form.get( 'ntlm', '' ) == '1' else ''
    command           = ''
    server_output     = ''

    while request.method == 'POST':
        if len( url ) <= 0 or len( username_password ) <= 0 or len( xml ) <= 0:
            command = 'Please fill out all above fields!'
            break

        post_data = re.sub( r"[\r\n]+\s+", ' ', xml )
        post_data = re.sub( r">\s+<", '><', post_data )
        post_data = post_data.replace( '"', "'" )

        c = pycurl.Curl()
        c.setopt( pycurl.URL, url )
        if request.form.get( 'ntlm', '' ) == '1':
            c.setopt( pycurl.HTTPAUTH, pycurl.HTTPAUTH_NTLM )
        c.setopt( pycurl.HTTPHEADER, ["text/xml; charset=utf-8"] )
        c.setopt( pycurl.USERPWD, username_password )
        c.setopt( pycurl.POST, 1 )
        c.setopt( pycurl.POSTFIELDS, post_data )
        c.setopt( pycurl.SSL_VERIFYHOST, 0 )
        c.setopt( pycurl.SSL_VERIFYPEER, 0 )
        storage = StringIO.StringIO()
        c.setopt( pycurl.WRITEFUNCTION, storage.write )
        c.setopt( pycurl.VERBOSE, 1 )
        c.setopt( pycurl.HEADER, 1 )

        c.perform()
        c.close()

        server_output = storage.getvalue()
        storage.close()

        command = 'curl -v -k'
        if request.form.get( 'ntlm', '' ) == '1':
            command += ' --ntlm'
        command += ( ' -H "Content-Type: text/xml; charset=utf-8" -u \'' + username_password + '\' -X POST -d "' + post_data + '" ' + url )

        break

    body = body.replace( '_REQUEST_TYPE_', request_type )
    body = body.replace( '_URL_', url )
    body = body.replace( '_USERNAME_PASSWORD_', username_password )
    body = body.replace( '_ROOM_EMAIL_', room_email )
    body = body.replace( '_XML_', xml )
    body = body.replace( '_NTLM_', ntlm )
    body = body.replace( '_CONTENT_', command + '\n\n' + server_output )
    return body
