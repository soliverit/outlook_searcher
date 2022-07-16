##
# Includes
##
## Native
## Project
from .outlooker_com_object	import OutlookerCOMObject
from .attendee				import Attendee
##
# Appointment: A wrapper for COM Appointments
##
class Appointment(OutlookerCOMObject):
	##
	# appointments:	List of COM Appointments from outlook client
	##
	def __init__(self, appointment):
		## Define studd
		super().__init__(appointment)
		self.attendees		= False			# List of attendees
		## Process stuff
		self.parseAttendees()
	##
	# Parse attendees:
	#
	# force:	Boolean if attendees have already been cached 
	#			but this is true, do it again anyway.
	##
	def parseAttendees(self, force=False):
		## Cache attendees unless it's done already and not forced
		if self.attendees and not force:
			return
		## Create Attendee instances for all participants
		self.attendees	= []
		for name in self.comObject.RequiredAttendees.split(";"):
			self.attendees.append(Attendee(name))
	##
	# Find an instance of an Attendee and replace it with the
	# passed instance:
	#
	# Ok, here me out... Attendees in the test project are known
	# to have unique names, meaninging that Dave Shoe and Shoe, Dave
	# will never refer to two or more people. By replacing the Attendee
	# instance created during this Appointment's constructor, Attendees
	# can be linked to all their appointments
	##
	def replaceAttendee(self, attendeeInstance):
		updatedAttendees	= []
		## Check all existing attendee instances
		for attendee in self.attendees:
			##
			# If the existing Attendee is effecitvely the passed,
			# Attendee - their Attendee.formattedName is the
			# same - then remove the instance created for this
			# Appointment and repalce with the passed instance.
			##
			if attendee == attendeeInstance:
				updatedAttendees.append(attendeeInstance)
				attendee.removeAppointment(self)
				attendeeInstance.addAppointment(self)
			else:
				updatedAttendees.append(attendee)
		self.attendees		= updatedAttendees
	##
	# Attended by: Was the passed Attendee either present or
	# required to attend the meeting.
	##
	def attendedBy(self, attendee):
		return attendee in self.attendees
		 