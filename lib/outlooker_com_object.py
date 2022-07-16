##
# Includes
##
## Native
## Project
from .participant	import Participant
##
# Outlooker COM Object
#
# A base class for objects returned by win32com client
# "Outlook.Application" queries which are scoped to a
# single directory in Outlook. Sent, Inbox, for example
# 
##
class OutlookerCOMObject():
	##
	# comObject:	COM Object from an Outlook folder search
	##
	def __init__(self, comObject):
		## Define stuff
		self.comObject		= comObject
		self.participants	= False
		## Process stuff
		self.parseParticipants()
	##
	# Parse participants:
	#
	# force:	Boolean if attendees have already been cached 
	#			but this is true, do it again anyway.
	##
	def parseParticipants(self, force=False):
		## Cache participants unless it's done already and not forced
		if self.participants and not force:
			return
		## Create Attendee instances for all participants
		self.participants	= []
		for name in self.comObject.RequiredAttendees.split(";"):
			self.participants.append(Participant(name))
	##
	# Find an instance of an participant and replace it with the
	# passed instance:
	#
	# Ok, here me out... Participants in the test project are known
	# to have unique names, meaninging that Dave Shoe and Shoe, Dave
	# will never refer to two or more people. By replacing the participant
	# instance created during this Appointment's constructor, Attendees
	# can be linked to all their appointments
	##
	def replaceParticipant(self, participantInstance):
		updatedParticipants	= []
		## Check all existing participant instances
		for participant in self.participants:
			##
			# If the existing participant is effecitvely the passed,
			# participant - their participant.formattedName is the
			# same - then remove the instance created for this
			# OutlookerCOMObject descendant and repalce with the passed instance.
			##
			if participant == participantInstance:
				updatedParticipants.append(participantInstance)
				participant.removeCOMObject(self)
				participantInstance.addCOMObject(self)
			else:
				updatedParticipants.append(participant)
		## Replace the Participants list
		self.participants	= updatedParticipants
	##
	# Check to see if a Participant was involved in
	#
	# particpant:	Participant who might be involved with the thing
	##
	def participantWasInvolved(self, participant):
		return participant in self.participants
		