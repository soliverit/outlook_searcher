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
	#
	# returns No. objects
	##
	def __len__(self):
		return len(self.objects)
	##
	# Combine two sets.
	#
	# Warning: 	Don't be a dick and merge searches from
	# 			multiple Outlook.Application folders.
	#
	# returns self. Weird feature of python
	##
	def __add__(self, otherSet):
		for object in otherSet.objects:
			self.append(object)
		return self
	##
	# Make a OutlookerCOMObject or descendant based on the Set's associated context
	#
	# comObject:	COMObject from win32com.client Outlook.Application.
	#
	# return SomeCOMObject(OutlookerCOMObject)
	##
	def makeObject(self, comObject):
		return self.objectType(comObject)
	##
	# Take a COMObject and create a self.objectType (Outlooker COMObject / Appointment etc)
	#
	# object:	A win32com COM object. Is converted to OutlokkerCOMObject or descendant.
	##
	def append(self, object):
		if object not in self.objects:
			self.objects.append(object)
	##
	# Unify participants.
	#
	# Identify Participant instances that represent the same person
	# and consolidate them into a single Participant identity.
	#
	# Warning: 	This assumes that you know all people with a name
	# 			matching Particpant.formattedName are the same person.
	#
	##
	def unifyParticipants(self):
		## Get one instance of all unique Participants
		participants	= []
		for object in  self.objects:
			for particpant in object.participants:
				if particpant not in participants:
					participants.append(particpant)
		## Replace Particpant instances with single instance identified 
		for object in self.objects:
			for particpant in participants:
				object.replaceParticipant(particpant)
	##
	# Get the unique Participants from everyone who participated
	# in at least one object's interaction.
	#
	# returns: list of unique Participants
	##
	@property
	def uniqueParticipants(self):
		## Get all Participants
		participants	= []
		for object in self.objects:
			participants += object.participants
		## Select the unique ones
		uniqueParticipants	= []
		for participant in participants:
			if participant not in uniqueParticipants:
				uniqueParticipants.append(participant)
		return uniqueParticipants
			
		
			
	
	
		