
import os, sys, decimal, datetime, easygui, os, csv, time, string

#------------------------------------------------------------------------------
#FUNCTIONS

#function that outputs the information to a .csv file
def print_csv(csv_list, size_entered, csv_writer):

    csv_writer.writerow(["Folder Name", "Size (KB)", "Size (MB)", "Size(GB)", "Number of Files", "Number of Files with Denied Permissions", "Most Recent Modification", "Created in the Last Month?", " ", "Report run on : " + current_time])

    if size_entered is True:
        output_list = [["", "","", "", cap_name + " SPACE USED:", str(gb_round_size), "GB"], ["", "", "", "", "TOTAL " + cap_name + " SIZE:", total_size, "GB"], ["", "", "", "", "PERCENT OF G DRIVE USED:", str(percent_rounded), "%"], ["", "", "", "", "1 KB =", "1,024", "bytes"], ["", "", "", "", "1 MB =", "1,048,576", "bytes"], ["", "", "", "", "1 GB =", "1,073,741,824", "bytes"], ["", "", "", "", "1 TB =", "1,099,511,627,776", "bytes"], ["", "", "", "", "NOTE: Files with denied permissions are included in file count, but NOT in folder size or modification date."], ["", "", "", "", "NOTE: Hidden Operating System Files are included in file count and folder size, but NOT in modification date."]]
    else:
        output_list = [["", "", "", "", cap_name + " SPACE USED:", str(gb_round_size), "GB"], [""], [""], ["", "", "", "", "1 KB =", "1,024", "bytes"], ["", "", "", "", "1 MB =", "1,048,576", "bytes"], ["", "", "", "", "1 GB =", "1,073,741,824", "bytes"], ["", "", "", "", "1 TB =", "1,099,511,627,776", "bytes"], ["", "", "", "", "NOTE: Files with denied permissions are included in file count, but NOT in folder size or modification date."], ["", "", "", "", "NOTE: Hidden Operating System Files are included in file count and folder size, but NOT in modification date."]]

    if len(csv_list) > 8:
        i = 0
        while i <= 8:
            write_list = [csv_list[i][0], csv_list[i][1], csv_list[i][2], csv_list[i][3], csv_list[i][4], csv_list[i][5], csv_list[i][6], csv_list[i][7]]
            write_list.extend(output_list[i])
            csv_writer.writerow(write_list)
            i += 1
        while 8 < i < len(csv_list):
            csv_writer.writerow([csv_list[i][0], csv_list[i][1], csv_list[i][2], csv_list[i][3], csv_list[i][4], csv_list[i][5], csv_list[i][6], csv_list[i][7]])
            i += 1

    else:
        i = 0
        while i < len(csv_list):
            write_list = [csv_list[i][0], csv_list[i][1], csv_list[i][2], csv_list[i][3], csv_list[i][4], csv_list[i][5], csv_list[i][6], csv_list[i][7]]
            write_list.extend(output_list[i])
            csv_writer.writerow(write_list)
            i += 1
        while len(csv_list) <= i <= 8:
            write_list = ["", "", "", "", "", "", ""]
            write_list.extend(output_list[i])
            csv_writer.writerow(write_list)
            i += 1

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
    size = decimal.Decimal(size)
    kb_conversion = decimal.Decimal(1024)
    mb_conversion = decimal.Decimal(1048576)
    gb_conversion = decimal.Decimal(1073741824)
    kb_size = round((size / kb_conversion), 0)
    mb_size = round((size / mb_conversion), 0)
    gb_size = round((size / gb_conversion), 0)

    return {"kb":kb_size, "mb":mb_size, "gb":gb_size, "b":size}

#------------------------------------------------------------------------------
#SETUP SCRIPT WITH USER INFORMATION AND INITIATE UI WINDOWS

#set title for easyGui windows
title = "Drive Reporter"

#display messgae in the console window
print "THIS WINDOW WILL BE USED FOR STATUS MESSAGES. CLOSING THIS WINDOW WELL END THE PROGRAM."

#display intro message, and allow the user to continue or cancel
if easygui.buttonbox("Welcome to Drive Reporter", title, ("Report on a Drive", "Cancel")) is "Report on a Drive":
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

if not dir_name:
    dir_index = root.find("\\")
    dir_name = root[0:dir_index]
    
cap_name = dir_name.upper()

#ask for the total available size of the drive if they know it and make sure then enter a number
result = easygui.enterbox("If you know it, please enter the total available size of " + dir_name + " in GB.\n\n(Select cancel to skip)", title)

num = False
size_attempted = False
size_entered = False

if result is None:
    pass
else:
    try:
        print result
        total_size = float(result)
        num = True
        size_entered = True
        size_attempted = True
    except:
        size_attempted = True

while size_attempted is True and num is False:
    result = easygui.enterbox("PLEASE ENTER A NUMBER VALUE. \n\nIf you know it, please enter the total available size of " + dir_name + " in GB.\n\n(Select cancel to skip)", title)
    if result is None:
        break
    else:
        try:
            total_size = float(result)
            num = True
            size_entered = True
            size_attempted = True
        except decimal.InvalidOperation: 
            size_attempted = True

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

print "\n\nPlease wait while the " + root + " report is prepared. \r\n\r\nFor larger directories this may take over an hour. \r\n\r\nClosing this window will cancel the report.\n\n\nREPORT WAS STARTED AT: " + startTimeTime 

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

#set an integer to calculate the size of the files stored in the root folder
root_size = 0

#creates an integer to track how many files in the root folder have denied permissions
root_permission_count = 0

#set an integer to count the number of files at the root level
root_num = 0

#create a sublist to hold the output information about the root folder's file contents
root_list = ["Root Folder"]

#create a list to store the modification dates of the root files
root_mod_list = []

#set an integer to hold the size of space used in the whole directory 
used_size = 0

#calculate root_size unless we don't have permissions, in which case append file_path to permission_denied_list
#find modification date if not a Thumbs.db file, and append to root_mod_list
permission_denied_list = []
count = 0
for x in subfile_list:
    while True:
        try: 
            root_size += os.path.getsize(x[0])
            root_num += 1
            if x[1] == "Thumbs.db":
                break
            else:
                root_mod = datetime.datetime.fromtimestamp(os.stat(x[0])[8])
                root_mod_list.append(root_mod)
                break
        except WindowsError:
            permission_denied_list.append(file_path)
            root_permission_count += 1
            root_num += 1
            break

root_sizes = bytes(root_size)

#sort root_mod, the last value will be the most recent date
root_sorted = sorted(root_mod_list)

#append root folder information to root_list
root_list.append('{0:,}'.format(round(root_sizes["kb"]), 2))
root_list.append('{0:,}'.format(round(root_sizes["mb"]), 2))
root_list.append('{0:,}'.format(round(root_sizes["gb"]), 2))
root_list.append('{0:,}'.format(root_num))
root_list.append('{0:,}'.format(root_permission_count))
if len(root_sorted) > 0:
    root_list.append(root_sorted[-1].strftime("%m-%d-%Y"))
else:
    root_list.append("NA")


root_created = datetime.datetime.fromtimestamp(os.path.getctime(root))
if root_created > date_check:
    root_recent = "yes"
else:        
    root_recent = "no"
    
root_list.append(root_recent)

#add root size to used_size
used_size += root_sizes["b"]

#------------------------------------------------------------------------------

#COMPILE INFORMATION ABOUT SUBFOLDERS AND FILES

#create list to hold output in
csv_list = []

#create a list to hold folders created in the last month
recent_list = []

#loop through subdirectories
for x in subdir_list:
    #create a sublist to hold the information about x before it is appended to csv_list
    sublist = [x[1]]

    #set an integer to track the number of files with denied permissions in each directory
    permission_count = 0

    #set an integer to calculate the size of the subdirectory 
    dir_size = 0
    
    #set an integer to count the number of files in a directory 
    dir_num = 0

    #create a list to hold the modification dates of the files in x
    mod_list = []

    #find subdirectories that were created within the last month
    created_date = datetime.datetime.fromtimestamp(os.path.getctime(x[0]))
    if created_date > date_check:
        recent = "yes"
    else:
        recent = "no"

    #walk through the subdirectory
    for dirpath, dirname, filenames in os.walk(x[0]):

        #create a path to each file within the subdirectory, and add to dir_size if we have permissions, otherwise append file_path to permission_denied_list
        #find modified date if not a Thumbs.db file, and append to mod_list
        for f in filenames:
            file_path = os.path.join(dirpath, f)
            while True:
                try: 
                    dir_size += os.path.getsize(file_path)
                    dir_num += 1
                    if f == "Thumbs.db":
                        break
                    else:
                        mod_file = datetime.datetime.fromtimestamp(os.stat(file_path)[8])
                        mod_list.append(mod_file)
                        break
                except WindowsError:
                    permission_denied_list.append(file_path)
                    permission_count += 1
                    dir_num += 1
                    break

    sizes = bytes(dir_size)

    #sort mod_list, the last value will be the most recent date
    mod_sorted = sorted(mod_list)

    #add dir_size to used_size
    used_size += dir_size

    #append the information about x to sublist, and then append sublist to a list
    #sublist.append('{0:,}'.format(dec_size))
    sublist.append('{0:,}'.format(round(sizes["kb"]), 2))
    sublist.append('{0:,}'.format(round(sizes["mb"]), 2))
    sublist.append('{0:,}'.format(round(sizes["gb"]), 2))
    sublist.append('{0:,}'.format(dir_num))
    sublist.append('{0:,}'.format(permission_count))
    if len(mod_sorted) > 0: 
        sublist.append(mod_sorted[-1].strftime("%m-%d-%Y"))
    else:
        sublist.append("NA")
    try:
        sublist.append(recent)
    except:
        pass
    csv_list.append(sublist)
    
#insert root_list to the beginning of csv_list
csv_list.insert(0, root_list)

final_sizes = bytes(used_size)
gb_round_size = round(final_sizes["gb"], 2)

if size_entered is True:
    one_hundred = decimal.Decimal(100)
    percent = ((final_sizes["gb"]/total_size)*100)
    percent_rounded = round(percent, 2)

#------------------------------------------------------------------------------

#CSV OUTPUT

#create a formatted string of startTime
current_time = startTime.strftime("%m-%d-%Y %I:%M %p")

#if the directory wasn't empty, write information to a csv file, otherwise end script with message
try:
    test = len(csv_list)
except NameError:
    output_file.close()
    os.remove(output_name)
    easygui.msgbox(dir_name + " is empty.", title)
    sys.exit()

csv_writer = csv.writer(output_file, delimiter = ",")
print_csv(csv_list, size_entered, csv_writer)

#------------------------------------------------------------------------------

#DENIED REPORT

#if there are files with denied permissions, write out files where permission was denied to the permissions denied report
if len(permission_denied_list) > 0:
    denied_writer = open(denied_name, "w")
    denied_writer.write("G Drive Report Run At: " + current_time + "\n" + "\n" + str(len(permission_denied_list)) + " files with denied permissions were not included in analysis.")
    denied_writer.write("\n" + "\n" + "--------------------------------------------------------------------------------------" + "\n" + "\n")
    for x in permission_denied_list:
        try:
            denied_writer.write("*" + x + "\n" + "\n")
        except:
            continue
    #close denied_writer
    denied_writer.close()

#------------------------------------------------------------------------------

#FINISHED MESSAGE

#calculate runtime 
runtime = datetime.datetime.now() - startTime 

#testWriter.write("runtime: " + str(runtime))

#set a string to be printed in the final notification popup and print to the console window
#print "\n\nTHE REPORT IS FINISHED. CLICK OK IN THE MESSAGE WINDOW TO QUIT."

#testWriter.write("afer finished report message")

if len(permission_denied_list) > 0:
    gui_string = root + " report finished. \n\nRuntime: " + str(runtime) + "\n\nFiles were found in " + dir_name + " that you do not have permissions to read. A list of these files has been saved here: " + denied_name + "\n\nThe report was saved as a .csv file which can be opened in Excel. To use formatting and Excel's features, resave the file as .xlsx."

else:
    gui_string = root + " report finished. \n\nRuntime: " + str(runtime) + "\n\nThe report was saved as a .csv file which can be opened in Excel. To use formatting and Excel's features, resave the file as .xlsx."

easygui.msgbox(gui_string, title)

raw_input(root + " report finished! \n\nPRESS ENTER TO QUIT, THEN CLICK OK IN THE MESSAGE WINDOW")
