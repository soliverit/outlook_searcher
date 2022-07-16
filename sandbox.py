##
# Includes
##
## Native
import win32com.client
from datetime 				import datetime
from re						import search
## Project`
from lib.outlook_searcher	import OutlookSearcher
from lib.appointment_set	import AppointmentSet



##
# Query appointments
##
start			= datetime(2020, 6, 22)
end				= datetime(2020, 12, 1)
pattern			= "(david\s?hunter)|(hunter\s?,\s?david)|(yondr)|(briink)|(commtech)|(logitek)|(logi-tek)"
searcher		= OutlookSearcher(OutlookSearcher.CALENDAR, AppointmentSet)
appointments	= searcher.search(start, end, "RequiredAttendees", pattern)
appointments	+= searcher.search(start, end, "subject", pattern)
appointments	+= searcher.search(start, end, "subject", "GCA")
appointments.unifyAttendees()
attendees		= appointments.uniqueAttendees
###
# Print stuff
##
searcher.help()
print("Length:\t" + str(len(appointments)))
for attendee in attendees:
	print(attendee)
exit()