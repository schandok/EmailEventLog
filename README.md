EmailEventLog
=============

This module emails Windows event logs since the last saved time stamp. If no time stamp exists (1st time), last 10 events will be sent.

This module can be attached as a task in Windows Event Viewer to be executed when certain event is logged.

Configure the params below to macth the event that triggered the module, such as logger and severity level (in cmdSinceLastTimeStamp).

All event details will be emailed. E.g: executing below command on command line, will show last 10 messages in System log with Level = Critical, Error or Warning. This would be the eact content of the email.

Test command: wevtutil qe "System" "/q:*[System [(Level<4)]]" /f:text /rd:true /c:10"
