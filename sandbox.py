import win32com.client
from datetime 				import datetime
from lib.outlook_searcher	import OutlookSearcher
from re						import search

##
# Query appointments
##
start			= datetime(2020, 6, 22)
end				= datetime(2020, 12, 1)
pattern			= "(david\s?hunter)|(hunter\s?,\s?david)|(yondr)|(briink)|(commtech)|(logitek)|(logi-tek)"
cal				= OutlookSearcher(OutlookSearcher.CALENDAR)
cal.help()
appointments	= cal.search(start, end, "RequiredAttendees", pattern)
appointments	+= cal.search(start, end, "subject", pattern)
appointments	+= cal.search(start, end, "subject", "GCA")
print(appointments)
appointments.unifyAttendees()
attendees		= appointments.uniqueAttendees
###
# Print stuff
##
print("Length:\t" + str(len(appointments)))
for attendee in attendees:
	print(attendee)