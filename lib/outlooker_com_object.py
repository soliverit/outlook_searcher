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
		self.comObject	= comObject