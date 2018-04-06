#03/09/16
fname = raw_input("Enter file name: ")
samplenumber = raw_input("Enter number of wells used: ")
wellsused = int(samplenumber)
fh = open(fname)
wr = open('BCAResults.txt', 'a')
count = 0
for line in fh:
	count = count + 1
	if not count%2 == 0 : continue
	if count > wellsused +1 : break
	line = line.rstrip()
	nline = line.split()
	value = nline[5]
	strvalue = str(value)
	wr.write(strvalue)
	wr.write('\n')


wr.write('\n')
wr.write('Second set:')
wr.write('\n')
wr.write('\n')
count = 0

sfh = open(fname)

for line in sfh:
	count = count + 1
	if count == 1 : continue
	if count%2 == 0 : continue
	if count > wellsused +1 : break
	line = line.rstrip()
	nline = line.split()
	value = nline[5]
	strvalue = str(value)
	wr.write(strvalue)
	wr.write('\n')

print "Done!"
something = raw_input("")
