import win32api
import keyboard
import os
from fpdf import FPDF
pdf = FPDF('P', 'mm', (192, 108))

#https://stackoverflow.com/questions/58523628/pyinstaller-exe-file-doesnt-take-any-input

print("""
*** Welcome to powerpoint meeting stealer ***
*App description: This app is intended to help you snip the screen and turn it into pdf
so you can grab meeting presentation easily
*How to use: just enter the title, then press "`" to copy & "shift + `" to stop
**Developer: draharjo
""")

#get all .png file that has been snipped ================
def get_all_file_paths(directory):
    # initializing empty file paths list
    file_paths = []
  
    # crawling through directory and subdirectories
    for root, directories, files in os.walk(directory):
        for filename in files:
            if filename.split(".")[1] == "png":
                # join the two strings in order to form the full filepath.
                filepath = os.path.join(root, filename)
                file_paths.append(filepath)
    # returning all file paths
    return file_paths

#select drive to save ====================================
drives = win32api.GetLogicalDriveStrings()
drives = drives.split('\000')[:-1]
print("select drive to save the file: ")
bool_drive = False
selected_drive = -1
cd = 1
for drive in drives:
    print(str(cd) + ". " + drive)
    cd = cd + 1   
while bool_drive is False:
    selected_drive = int(input("drive : "))
    if selected_drive > 0 and selected_drive <= len(drives):
        bool_drive = True

the_drive = drives[selected_drive-1]
root = the_drive + "ppt stealer"     
print("Your snipping will be saved on " + root)

#snip and save to pdf ===================================
loop = True
while loop == True:
    title = input("meeting apa? : ")
    folder = root + "\\" + title
    os.chdir(the_drive)
    CHECK_FOLDER = os.path.isdir(folder)

    if CHECK_FOLDER is False:
        os.makedirs(folder)
        loop = False
        print("folder created on "+ folder + " , to snip press (`) and to stop (shift + `)")
    else:
        print("there is same folder already exist")

c = 1
while True:
    if keyboard.read_key() == "`":
        from PIL import ImageGrab
        im = ImageGrab.grab()
        im.save(folder+ "\\" + str(c)+ "_" + title + ".png")
        print("capture " + str(c)+ "_" + title + ".png")
        c = c + 1
    if keyboard.read_key() == "~":
        imagelist = get_all_file_paths(folder)
        for image in imagelist:
            pdf.add_page()
            pdf.image(image,x=0,y=0,w=192,h=108)
        pdf.output(folder + "\\" + title + ".pdf", "F")
        print("save capture into " + folder + "\\" + title + ".pdf")
        break
    

