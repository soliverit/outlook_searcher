##
# Includes
##
## Native
## Project
from .outlooker_com_object	import OutlookerCOMObject

##
# Outlooker object Set
#
# A collectio manager for objects taken from Outloo.Application
# folders. Sent, Inbox, Calendar etc.
##
class OutlookerCOMObjectSet():
	def __init__(self):
		self.objects	= []
		self.objectType	= OutlookerCOMObject
	##
	# Number of objects in the set
	##
	def __len__(self):
		return len(self.objects)
	##
	# Combine two sets.
	#
	# Warning: 	Don't be a dick and merge searches from
	# 			multiple Outlook.Application folders.
	##
	def __add__(self, otherSet):
		for object in otherSet.objects:
			self.append(object)
		return self
	
	def append(self, object):
		if object not in self.objects:
			self.objects.append(self.objectType(object))
	def unifyAttendees(self):
		attendees	= []
		for object in  self.objects:
			appointmentAttendees	= object.attendees
			for attendee in appointmentAttendees:
				if attendee not in attendees:
					attendees.append(attendee)
		for object in self.objects:
			for attendee in attendees:
				object.replaceAttendee(attendee)
	@property
	def uniqueAttendees(self):
		attendees		= []
		for object in self.objects:
			attendees += object.attendees
		uniqueAttendees	= []
		for attendee in attendees:
			if attendee not in uniqueAttendees:
				uniqueAttendees.append(attendee)
		return uniqueAttendees
			
		
			
	
	
		