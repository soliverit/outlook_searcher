from .appointment	import Appointment
class AppointmentSet():
	def __init__(self, appointments, unifyAttendees=False):
		self.appointments	= []
		for appointment in appointments:
			self.addAppointment(appointment)
		if unifyAttendees:
			self.unifyAttendees()
	def __len__(self):
		return len(self.appointments)
	def __add__(self, otherSet):
		for appointment in otherSet.appointments:
			self.addAppointment(appointment)
		return self
	def addAppointment(self, appointment):
		if appointment not in self.appointments:
			self.appointments.append(appointment)
	def unifyAttendees(self):
		attendees	= []
		for appointment in  self.appointments:
			appointmentAttendees	= appointment.attendees
			for attendee in appointmentAttendees:
				if attendee not in attendees:
					attendees.append(attendee)
		for appointment in self.appointments:
			for attendee in attendees:
				appointment.replaceAttendee(attendee)
	@property
	def uniqueAttendees(self):
		attendees		= []
		for appointment in self.appointments:
			attendees += appointment.attendees
		uniqueAttendees	= []
		for attendee in attendees:
			if attendee not in uniqueAttendees:
				uniqueAttendees.append(attendee)
		return uniqueAttendees
			
		
			
	
	
		