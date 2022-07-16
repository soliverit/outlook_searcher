##
# Includes
##
## Native
import win32com.client
from re							import search, IGNORECASE
from datetime					import datetime
## Project
from .outlooker_com_object 		import OutlookerCOMObject
from .outlooker_com_object_set	import OutlookerCOMObjectSet

##
# Outlook Searcher: Search email folders for things
#
# The class is designed to search folders in the currently
# logged in Outlook account.
##
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
	##
	# folderID:		Integer (from enumerated values above) denoting the target Outlook folder
	##
	def __init__(self, folderID, itemSetType=OutlookerCOMObjectSet, includeResources=True):
		## Define stuff
		self.client							= win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
		self.folder							= self.client.getDefaultFolder(folderID)
		self.folderItems					= self.folder.Items	
		self.folderItems.IncludeRecurrences = includeResources
		##
		# Handling item set dependencies:
		#
		# Sets and object types are not particularly pretty in handling. Some searches, like
		# Calendar have bespoke objects for sets and instances. Rather than create a complicated
		# hash of the enumerated values and have the user modify instance and set classes as required
		# but I figure that the constructor can just be passed a set type. Bonus, set types can be mixed
		##
		self.itemSetType	= itemSetType
	##
	# Search self.folder for items with a property whose value matches a regex search pattern.
	#
	# start:	datetime first day of the query period
	# end:		datetime last day of the query period
	# target:	string property to be searched
	# pattern:	string regex pattern for searching the target property
	##
	def search(self, start, end, target, pattern):
		# Get all appointments between start and end date
		self.folderItems.Sort('[Start]')
		restriction 		= "[Start] >= '" + start.strftime('%d/%m/%Y') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
		folderItems			= self.folderItems.Restrict(restriction)
		# Filter results with regex
		results				= self.itemSetType()
		for item in folderItems:
			if search(pattern, getattr(item, target), IGNORECASE):
				results.append(item)
		return results
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
		
	
