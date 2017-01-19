import os, sys, decimal, datetime, easygui, os, csv, time, string

#CLASSES
class FileType(object):
    def __init__(self, typ, size):
        self.typ = typ
        self.size = size
        self.count = 1
    def getTyp(self):
        return self.typ
    def getTypSize(self):
        return self.size
    def getCount(self):
        return self.count
    def setSize(self, new_size):
        self.size = self.size + new_size
    def incCount(self):
        self.count = self.count + 1

#------------------------------------------------------------------------------
#FUNCTIONS

#function that outputs the types information to a .csv file
def print_types(file_types, csv_writer):

    csv_writer.writerow(["File Type", "Number of Files", "Size (B)", "Size (KB)", "Size (MB)", "Size(GB)", "Report run on : " + current_time])

    for ls in file_types:
        sizes = bytes(ls[1])
        new_list = [ls[0], ls[2], ls[1], sizes["kb"], sizes["mb"], sizes["gb"]]
        csv_writer.writerow(new_list)

#------------------------------------------------------------------------------
#function to check if the output file type entered is valid
def type_check(valid_name):
    output_name = easygui.filesavebox("Where would you like to save the report? Please save as a .csv file", title, ".csv")
    try:
        period_index = output_name.index(".")
    except ValueError:
        return False
    except AttributeError:
        sys.exit()
    output_type = str(output_name[period_index:])
    if output_type == ".csv":
        valid_name = True
        return output_name
    else:
        return False

#------------------------------------------------------------------------------
#function to convert one size into kb, mb, and gb
def bytes(size):
    kb_conversion = decimal.Decimal(1024)
    mb_conversion = decimal.Decimal(1048576)
    gb_conversion = decimal.Decimal(1073741824)
    kb_size = round((size / kb_conversion), 0)
    mb_size = round((size / mb_conversion), 0)
    gb_size = round((size / gb_conversion), 0)
    
    return {"kb":kb_size, "mb":mb_size, "gb":gb_size, "b":size}

#------------------------------------------------------------------------------
#function that creates a new FileTyp object if one does not already exist, otherwise it calls setSize to add to the FileType's size
def create_FileType(file_path, type_list):
    filename, typ = os.path.splitext(file_path)
    if isinstance(typ, str):
        return
    typ = typ.lower()
    size = os.path.getsize(file_path)
    if typ == ".jpeg":
        typ = ".jpg"
    elif typ == ".doc":
        typ = ".docx"
    elif typ == ".xls":
        typ = ".xlsx"
    elif typ == ".ppt":
        typ = ".pptx"
    for sublist in type_list:
        if sublist[0] == typ:
            sublist[2] = sublist[2] + 1
            sublist[1] = sublist[1] + size
            return
    type_list.append([typ, size, 1])


#------------------------------------------------------------------------------

#SETUP SCRIPT WITH USER INFORMATION AND INITIATE UI WINDOWS

#set title for easyGui windows
title = "Drive File Summary"

#display messgae in the console window
print "THIS WINDOW WILL BE USED FOR STATUS MESSAGES. CLOSING THIS WINDOW WELL END THE PROGRAM."

#display intro message, and allow the user to continue or cancel
if easygui.buttonbox("Welcome to Drive Inventory", title, ("Report on a Drive", "Cancel")) is "Report on a Drive":
    pass
else:
    sys.exit()

#ask the user to set the directory to be searched 
root = easygui.diropenbox("Please select the FOLDER you want to report on", title, "C:\\")
if root is not None:
    pass
else:
    sys.exit()

#pull out the name of the directory to be searched 
dir_index = root.rfind("\\") + 1
dir_name = root[dir_index:]
cap_name = dir_name.upper()

#ask where they want to save the report, and make sure the specify a .csv file
valid_name = False

output_name = type_check(valid_name)

while output_name is False:
    easygui.msgbox("Please save the report as a .csv file.", title)
    output_name = type_check(valid_name)
    
#create the csv writer to make sure they don't have the file already open
file_open = True
while file_open is True:
    try:
        output_file = open(output_name, "wb")
        file_open = False
    except IOError:
        cc = easygui.ccbox(output_name + " is open. \n\nPlease close the file and click \"Continue\". \n\n NOTE: THE FILE WILL BE OVERWRITTEN IF YOU CONTINUE.", title)
        if cc == 1:
            pass
        else:
            sys.exit()

#create the denied writer name
denied_name = output_name[:len(output_name)-4] + "Permissions Denied List.txt"

#show a holding window while the script is running

#record a time stamp to later calculate total runtime
startTime = datetime.datetime.now()

#convert startTime to a time formatted string
startTimeTime = startTime.strftime("%I:%M %p")

#convert startTime to a date formatted string
date = startTime.strftime("%m-%d-%Y")

#calculate date 30 days ago to later calculate list of > 1 month old folders
date_check = datetime.datetime.now() - datetime.timedelta(30)

print "\n\nPlease wait while the " + root + " file summary is prepared. \r\n\r\nFor larger directories this may take over an hour. \r\n\r\nClosing this window will cancel the report.\n\n\nREPORT WAS STARTED AT: " + startTimeTime 

#------------------------------------------------------------------------------

#SCAN DIRECTORY FOR CONTENTS

#creates a list of the names of all files and folders in root
root_contents = os.listdir(root)

#creates a list of the paths to all of the entries in root_contents
root_paths = []
for x in root_contents:
    root_paths.append(os.path.join(root, x))

#creates separate lists of the paths to the immediate subdirectories, and of the files at the root level, in root
subdir_paths = []
subfile_paths = []
for x in root_paths:
    if os.path.isdir(x):
        subdir_paths.append(x)
    else:
        subfile_paths.append(x)

#creates a list of the names of the directories whose paths are in subdir_paths
subdir_names = []
for x in subdir_paths:
    subdir_names.append(os.path.basename(x))

#creates a list of the names of the files whose paths are in subfile_paths
subfile_names = []
for x in subfile_paths:
    subfile_names.append(os.path.basename(x))

#creates a list of lists, where each sublist is [0] = subdirectory path, [1] = subdirectory name
subdir_list = []
subdir_list = zip(subdir_paths, subdir_names)

#creates a list of lists, where each sublist is [0] = file path, [1] = file name
subfile_list = []
subfile_list = zip(subfile_paths, subfile_names)

#------------------------------------------------------------------------------

#COMPILE INFORMATION ABOUT ROOT

#list to hold FileTyp objects
file_types = []

##
for x in subfile_list:
    while True:
        try: 
            if x[1] == "Thumbs.db":
                break
            else:
                create_FileType(x[0], file_types)
                break
        except WindowsError:
            break
#------------------------------------------------------------------------------

#COMPILE INFORMATION ABOUT SUBFOLDERS AND FILES

#loop through subdirectories
for x in subdir_list:
    #walk through the subdirectory
    for dirpath, dirname, filenames in os.walk(x[0]):

        #create a path to each file within the subdirectory
        ##
        for f in filenames:
            file_path = os.path.join(dirpath, f)
            while True:
                try: 
                    if f == "Thumbs.db":
                        break
                    else:
                        create_FileType(file_path, file_types)
                        break
                except WindowsError:
                    break


#------------------------------------------------------------------------------

#CSV OUTPUT

#create a formatted string of startTime
current_time = startTime.strftime("%m-%d-%Y %I:%M %p")

csv_writer = csv.writer(output_file, delimiter = ",")
print_types(file_types, csv_writer)

#------------------------------------------------------------------------------

#FINISHED MESSAGE

#calculate runtime 
runtime = datetime.datetime.now() - startTime 

#set a string to be printed in the final notification popup and print to the console window
#print "\n\nTHE FILE SUMMARY IS FINISHED. HIT ENTER, THEN CLICK OK IN THE MESSAGE WINDOW TO QUIT."

gui_string = root + " file summary finished. \n\nRuntime: " + str(runtime) + "\n\nThe file summary was saved as a .csv file which can be opened in Excel. To use formatting and Excel's features, resave the file as .xlsx.\n\nPlease go to the console window, and hit enter to quit."

easygui.msgbox(gui_string, title)

raw_input("\n\nfile summary finished! \n\nPRESS ENTER TO QUIT, THEN CLICK OK IN THE MESSAGE WINDOW")
