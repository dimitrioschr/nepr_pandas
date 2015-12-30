import os

path = 'C:\\Users\\tech5\\Desktop\\PHOENIX RISING workout'
path2 = 'C:\\Users\\tech5\\Google Drive\\NEPR Actual'

listing = os.listdir(path2)

files = [l for l in listing if os.path.isfile(os.path.join(path2, l))]
dirs = [l for l in listing if l not in files]

print('......... Files:')
for f in files:
    print(f)
print()

print('......... Directories:')
for d in dirs:
    print(os.path.join(path2, d))
print()

print('.........')

# osstruct = list(os.walk(path2))
#
# for element in osstruct:
#     print(element)
#
# print('.........')
#
# for d in osstruct[0][1]:
#     dir = os.path.join(osstruct[0][0], d)
#     print(dir)


