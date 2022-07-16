##
# Includes
##
## Native
## Project
from .outlooker_com_object	import OutlookerCOMObject
from .participant			import Participant
##
# Appointment: A wrapper for COM Appointments
##
class Appointment(OutlookerCOMObject):
	##
	# appointments:	List of COM Appointments from outlook client
	##
	def __init__(self, appointment):
		## Define stuff
		super().__init__(appointment)
		self.participants		= False			# List of attendees
		## Process stuff
		self.parseAttendees()
	
	##
	# Attended by: Was the passed Attendee either present or
	# required to attend the meeting.
	##
	def attendedBy(self, attendee):
		return attendee in self.participants
		 