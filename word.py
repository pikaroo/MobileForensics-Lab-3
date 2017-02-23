# Mobile Forensics - Lab 3
# name: Alan Pike
# Student Number: D16124621
# Usage: python word.py "c:\location\of\file.docx"

import sys
import win32com.client as win32
# Declare vaiable openedDoc to be of type word document and use module from the win32com library
openedDoc = win32.gencache.EnsureDispatch('Word.Application')
# Declare variable of filename to be passed into script when it is ran from command prompt
filename= sys.argv[1]
# Declare Variable of dictionary_file, and open the dictionary.txt file, using argument 'r' to read it
dictionary_file = open ( 'dictionary.txt', 'r' )
# Declare variable passwords to read each line of the password_file (dictionary.txt. file)
passwords = dictionary_file.readlines()
# Close the dictionary_file
dictionary_file.close()
# Set passwords to be an array containing all items in dictionary file
passwords = [item.rstrip('\n') for item in passwords]
# Declare variable "results" to open a file called "password_result.txt", and use "w" to set it to write to the file (this will create a file called password_result.txt in the same directory)
results = open('Password_result.txt', 'w')
# set variable password_found to be 0.  This will be set to 1 if the file has been opened, to allow for the program to end
password_found = 'false'
# set variable password to be the next item in array "passwords", print the password to the screen
for password in passwords:
	# If the password_found remains set as = 0
	if(password_found == 'false' ):
		print(password)
		try:
			# Try open the docment using the filename specified when the script has ran.  It will also pass 
			# The following are pre-defined for use in Documents.Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument) 
			# In this case, it is passing the filename, setting ConfirmConversion=False, ReadOnly=False, AddToRecentFiles=None and setting PasswordDocument=password
			wb = openedDoc.Documents.Open(filename, False, False, None, password)
			# If the password is correct, it will print the following to the screen and send the value of password to the variable "result", which will write this to the password_result.txt file
			print("Password identified. Password is: "+password)
			results.write(password)
			# it will close the password_result.txt file and exit out of the script
			results.close()
			# update variable "password_found" to be true.  This will be checked on the next iteration within the password array and, as it won't be equal to false, it will go to else statement to exit the program
			password_found = 'true'
		except:
			# if the password is incorrect and it cannot open the file, it will output "Incorrect password to the screen"
			print("Incorrect password")
			pass
	else:
		sys.exit(0)