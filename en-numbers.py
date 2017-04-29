import sys
import time
import random
import platform
import subprocess

if 'Darwin' != platform.system():
	print 'Only support macOS for now!'
	exit( 1 )

min = 1
max = 20

if len( sys.argv ) > 1:
	if sys.argv[1].isdigit():
		max = int( sys.argv[1] )
	else:
		print 'Usage: ' + sys.argv[0] + ' number'
		exit( 1 )

print 'We will work on ' + str( min ) + ' ~ ' + str( max )

while True:
	command = ['say', '-r', '240']
	number = random.randint( min, max )
	command.append( str( number ) )
	if 0 != subprocess.call( command ):
		exit( 1 )

	answer = raw_input( '=> ' )
	if len( answer ) == 0 or not answer.isdigit():
		break

	if int( answer ) == number:
		subprocess.call( ['say', 'great'] )
	else:
		print "it's actually " + str( number )
		subprocess.call( ['say', 'wrong'] )

	time.sleep( 0.5 )

print 'bye'
