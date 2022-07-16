class Attendee():
	def __init__(self, name):
		self.name			= name
		self.appointments	= []
	def __eq__(self, other):
		return self.lowerFormattedName == other.lowerFormattedName
	def __str__(self):
		return str(len(self.appointments)) + "\t| " + self.formattedName
	@property
	def formattedName(self):
		## Cleanse name
		name		= self.name
		name		= name.replace("\"", "")
		name		= name.replace("'", "") 
		name		= name.replace("(Guest)", "")
		name		= name.strip()
		## Handle <surname>, <forename> format
		segments	= name.split(",")
		if len(segments) == 2:
			name	= (segments[1] + " " + segments[0]).strip()
		return name
	@property
	def lowerFormattedName(self):
		return self.formattedName.lower()
	def addAppointment(self, appointment):
		if appointment.attendedBy(self):
			if appointment not in self.appointments:
				self.appointments.append(appointment)
	def removeAppointment(self, appointment):
		appointments	= []
		for existingAppointment in self.appointments:
			if appointment != existingAppointment:
				appointments.append(existingAppointment)
		self.appointments	= appointments
			
	