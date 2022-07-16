##
# Includes
##
## Native
## Project

##
# Participatn
#
# A Person who particiapted in a thing. Appointment, Email etc.
##
class Participant():
	##
	# name:	string person's raw name. Raw like ["David Shoe", "Shoe, David"] etc
	##
	def __init__(self, name):
		## Define stuff
		self.name			= name
		self.comObjects		= []
	##
	# Compare this Participant with another.
	#
	# Ok, basically the project this was written for has no
	# people with identical names.
	#
	# other:	Participant to compare
	#
	# Warning: This doesn't account for email address / name pairs.
	##
	def __eq__(self, other):
		return self.lowerFormattedName == other.lowerFormattedName
	##
	# Make a pretty string 
	##
	def __str__(self):
		return str(len(self.comObjects)) + "\t| " + self.formattedName
	##
	# Get a formatted version of the Person's name
	#
	# Names are cleansed of punctuation that's not whitespace
	#
	# E.g: Dickenson, David and David Dickenson are both David Dickenson
	#
	# returns	string formatted Name.
	#
	##
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
	##
	# Get the formatted name - but lower()!!!
	#
	# return string downcase of self.formattedName
	##
	@property
	def lowerFormattedName(self):
		return self.formattedName.lower()
	#############################
	### COM object management ###
	#############################
	# Every Item from an Outlook.Application folder
	# is a COMObject. Each type has different
	# properties like subject, RequiredAttendees.
	#
	### Tip / Recommendation:
	#
	# These method name's are brutally, unashamedly
	# abstracted... but that makes them unrelatable
	# in application. For example, Appointment
	# addCOMObject tells you nothing. 
	#
	# So, why not create forwarders? 
	#
	#############################
	##
	# Add COM object
	#
	# comObject:	OutlookerCOMObject to be added
	##
	def addCOMObject(self, comObject):
		if comObject.participantWasInvolved(self):
			if comObject not in self.comObjects:
				self.comObjects.append(comObject)
	##
	# Remove COM object
	#
	# comObject:	OutlookerCOMObject for removal
	##
	def removeCOMObject(self, comObject):
		comObjects	= []
		for comObject in self.comObjects:
			if appointment != existingCOMobject:
				comObjects.append(existingCOMobject)
		self.comObjects	= comObjects
			
	