import win32com.client
from re					import search, IGNORECASE
from datetime			import datetime
from .appointment_set	import AppointmentSet
from .appointment		import Appointment

class OutlookSearcher():
	##
	# Folder enumerated values
	##
	DELETED							= 3
	OUTBOOX							= 4
	SENT							= 5
	INBOX							= 6
	CALENDAR						= 9
	CONTACTS						= 10
	JOURNAL							= 11
	NOTES							= 12
	TASKS							= 13
	REMINDERS_1						= 14
	REMINDERS_2 					= 15
	DRAFTS							= 16
	CONFLICTS_1						= 17
	ALL_PUBLIC_FOLDERS				= 18
	CONFLICTS_2						= 19
	SYNC_ISSUES						= 20
	LOCAL_FAILURES					= 21
	SERVER_FAILURES					= 22
	JUNK_EMAIL						= 23
	RSS_SUBSCRIPTIONS				= 25
	TRACKED_MAIL_PROCESSING			= 26
	TODO_LIST						= 28
	QUICK_STEP_SETTINGS				= 31
	CONTACT_SEARCH					= 33
	SOCIAL_ACTIVITY_NOTIFICATIONS	= 37
	def __init__(self, folderID):
		self.client		= win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
		self.folderID	= folderID
	def search(self, start, end, target, pattern):
		calendar 	= self.client.getDefaultFolder(self.folderID).Items	# Calnedar is 9, probably. Terrible get method
		calendar.IncludeRecurrences = True
		# Get all appointments between start and end date
		calendar.Sort('[Start]')
		restriction = "[Start] >= '" + start.strftime('%d/%m/%Y') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
		calendar 	= calendar.Restrict(restriction)
		# Filter results
		results		= []
		for appointment in calendar:
			if search(pattern, getattr(appointment, target), IGNORECASE):
				results.append(Appointment(appointment))
		return AppointmentSet(results)
	def help(self):
		print("\n### Selecting the query folder ###\n")
		print("Select a folder using the folder enumerated values")
		print("--- Enumerated folder values ---")
		print("DELETED                            = 3")
		print("OUTBOOX                            = 4")
		print("SENT                            = 5")
		print("INBOX                            = 6")
		print("CALENDAR                        = 9")
		print("CONTACTS                        = 10")
		print("JOURNAL                            = 11")
		print("NOTES                            = 12")
		print("TASKS                            = 13")
		print("REMINDERS_1                        = 14")
		print("REMINDERS_2                     = 15")
		print("DRAFTS                            = 16")
		print("CONFLICTS_1                        = 17")
		print("ALL_PUBLIC_FOLDERS                = 18")
		print("CONFLICTS_2                        = 19")
		print("SYNC_ISSUES                        = 20")
		print("LOCAL_FAILURES                    = 21")
		print("SERVER_FAILURES                    = 22")
		print("JUNK_EMAIL                        = 23")
		print("RSS_SUBSCRIPTIONS                = 25")
		print("TRACKED_MAIL_PROCESSING            = 26")
		print("TODO_LIST                        = 28")
		print("QUICK_STEP_SETTINGS                = 31")
		print("CONTACT_SEARCH                    = 33")
		print("SOCIAL_ACTIVITY_NOTIFICATIONS    = 37")
		print("\n### Query folder ###\n")
		print("Queries have 4 arguments:")
		print("\tstart:   datetime earliest")
		print("\tend:     datetime latest")
		print("\ttarget:  string folder item search attribute")
		print("\tpattern: string regex search parameters")
		
	
