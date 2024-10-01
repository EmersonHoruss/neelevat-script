import datetime

file = open(r'C:\Users\emerson\Downloads\projects\script-neelevat\task.txt', 'a')

file.write(f'{datetime.datetime.now()} - the script ran\n')
