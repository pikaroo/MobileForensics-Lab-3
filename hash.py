import sys
import hashlib
# Set block size to 65MB, to deal with large files
BLOCKSIZE = 65536
# Set variable my_file from file name inputted by user
my_file = raw_input("Enter the name of the file you wish to check: ")
# declare variables for each hash algorithm
sha1 = hashlib.sha1()
sha256 = hashlib.sha256()
sha512 = hashlib.sha512()
sha384 = hashlib.sha384()
sha3_512 = hashlib.sha3_512()

with open(my_file, 'rb') as file_to_check:
    # read the contents of the file and store it in variable called "contents"
    # rb specifies to read the file in binary mode
    contents = file_to_check.read(BLOCKSIZE)
    # update each variable state with variable "contents"
    sha1.update(contents)
    sha256.update(contents)
    sha512.update(contents)
    sha384.update(contents)
    sha3_512.update(contents)
    
print("                                ")
print("  Mobile Forensics - Lab 2 (modified)     ")
print("                               ")
print("   Name:             Alan Pike ")
print("   Student Number:   D16124621 ")
print("                               ")
print("File name: " +my_file)
print("                                ")
# work out the hex digest of each variable by calling hexdigest function and output each one
print("SHA1: {0}".format(sha1.hexdigest()))
print("SHA256: {0}".format(sha256.hexdigest()))
print("SHA512: {0}".format(sha512.hexdigest()))
print("SHA384: {0}".format(sha384.hexdigest()))
print("SHA3_512: {0}".format(sha3_512.hexdigest()))
