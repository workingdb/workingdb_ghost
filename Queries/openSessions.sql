SELECT tblWdbSessions.recordId, tblWdbSessions.username, tblWdbSessions.wdbVersion, tblWdbSessions.openForms, tblWdbSessions.sessionStart, tblWdbSessions.sessionEnd, tblWdbSessions.lastCheck, tblWdbSessions.machine
FROM tblWdbSessions
WHERE (((tblWdbSessions.sessionEnd) Is Null))
ORDER BY tblWdbSessions.sessionStart;

