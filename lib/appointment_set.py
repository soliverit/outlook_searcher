##
# Includes
##
## Native
## Project
from .appointment					import Appointment
from .outlooker_com_object_set		import OutlookerCOMObjectSet
##
# Outlooker Item Set
#
# A collectio manager for Items taken from Outloo.Application
# folders. Sent, Inbox, Calendar etc.
##
class AppointmentSet(OutlookerCOMObjectSet):
	def __init__(self):
		# Define stuff
		super().__init__()
		##
		# Handling item set dependencies: Copy paste from OutlookerItemSet __init__.
		#
		# Sets and object types are not particularly pretty in handling. Some searches, like
		# like Calendar have bespoke objects for sets and instances. Rather than create a complicated
		# hash of the enumerated values and have the user modify instance and set classes as required
		# but I figure that the constructor can just override these defaults
		##
		self.itemType		= Appointment

	def addAppointment(self, appointment):
		self.addItem(appointment)
	# def unifyAttendees(self):
		# attendees	= []
		# for appointment in  self.appointments:
			# appointmentAttendees	= appointment.attendees
			# for attendee in appointmentAttendees:
				# if attendee not in attendees:
					# attendees.append(attendee)
		# for appointment in self.appointments:
			# for attendee in attendees:
				# appointment.replaceAttendee(attendee)
	# @property
	# def uniqueAttendees(self):
		# attendees		= []
		# for appointment in self.appointments:
			# attendees += appointment.attendees
		# uniqueAttendees	= []
		# for attendee in attendees:
			# if attendee not in uniqueAttendees:
				# uniqueAttendees.append(attendee)
		# return uniqueAttendees
			
		
			
	
	
		