import os
import sys
import ctypes
from ctypes import wintypes
import win32con
import time
import threading
import concurrent.futures
import subprocess


#Implement Hotkey registration
"""
byref = ctypes.byref
user32 = ctypes.windll.user32

Hotkeys = {
    1 : (win32con.VK_P, win32con.MOD_CTRL)
}

def handle_cancel():
    sys.exit('Process aborted by user')

Hotkey_Actions = {
    1 : handle_cancel
}

for id, (vk, modifiers) in HOTKEYS.items ():
  print("Registering id " + str(id) + " for key " + str(vk))
  if not user32.RegisterHotKey (None, id, modifiers, vk):
    print("Unable to register id " + str(id))

#
# Home-grown Windows message loop: does
#  just enough to handle the WM_HOTKEY
#  messages and pass everything else along.
#
try:
  msg = wintypes.MSG ()
  while user32.GetMessageA (byref (msg), None, 0, 0) != 0:
    if msg.message == win32con.WM_HOTKEY:
      action_to_take = HOTKEY_ACTIONS.get (msg.wParam)
      if action_to_take:
        action_to_take ()

    user32.TranslateMessage (byref (msg))
    user32.DispatchMessageA (byref (msg))

finally:
  for id in HOTKEYS.keys ():
    user32.UnregisterHotKey (None, id)
"""


#Import dependencies. If not already downloaded & installed, call subprocess 'pip install <module>' to install it
try:
    import pyautogui
except ImportError:
    subprocess.call('pip install pyautogui')
    time.sleep(5)

ver = None

#Due to an issue with pulling pixels right after importing 'pyautogui', run a while loop to help mitigate the issue
while ver is None:
    try:
        ver = pyautogui.pixel(80, 60)
    except:
        None    

#Parse the MS Teams theme by comparing pixel colors
if ver == (45, 44, 43):
    ver = 'Dark'
elif ver == (255, 255, 252):
    ver = 'Light'
else:
    ver = 'INCORRECT STYLE'

#Set picture comparison files based on theme
if ver == 'Dark':
    #Gather the necessary image files and assign them to the necessary variables
    directory, file = os.path.split(os.path.abspath(sys.argv[0]))
    directory = directory + '\\Dark-ImageMatches\\'
    recentPic = os.path.join(directory, 'Recent.png')
    hidePic = os.path.join(directory, 'HideButton.png')
    hidePic2 = os.path.join(directory, 'HideButton2.png')
    hidePic3 = os.path.join(directory, 'HideButton3.png')
    leavePic = os.path.join(directory, 'LeaveButton.png')
    leavePic2 = os.path.join(directory, 'LeaveButton2.png')    
elif ver == 'Light':
    #Gather the necessary image files and assign them to the necessary variables
    directory, file = os.path.split(os.path.abspath(sys.argv[0]))
    directory = directory + '\\Light-ImageMatches\\'
    recentPic = os.path.join(directory, 'Recent.png')
    hidePic = os.path.join(directory, 'HideButton.png')
    hidePic2 = os.path.join(directory, 'HideButton2.png')
    hidePic3 = os.path.join(directory, 'HideButton3.png')
    leavePic = os.path.join(directory, 'LeaveButton.png')
    leavePic2 = os.path.join(directory, 'LeaveButton2.png')

#Non-discriminate function to remove all chats
def clearAll(count):
    while count > 0:
        #Find the location of the 'Recent' tag in MS Teams
        recentLoc = pyautogui.locateCenterOnScreen(recentPic)
        recentx, recenty = recentLoc
    
        #Set the coords for the first entry (a.k.a 20 over & 30 down)
        firstx = recentx + 20
        firsty = recenty + 30
    
        hide(firstx, firsty)

        count -= 1

#Perform a search to see if the text color matches (255, 255, 252) (identifier for white text which means the chat is new)
def isNew(firstx, firsty, ver):
    l = []
    l.clear()

    repeat = True

    while repeat == True:
        repeat = False
        for x in range(0, 21):
            try:
                if ver == 'Dark':
                    l.append(pyautogui.pixelMatchesColor(firstx - 50, firsty + x, (255, 255, 252)))
                elif ver == 'Light':
                    l.append(pyautogui.pixelMatchesColor(firstx - 50, firsty + x, (37, 36, 34)))
            except:
                repeat = True
                break
        
    if l.count(True) >= 1:
        return True
    else:
        return False

def isGroup(firstx, firsty):          
    #Right-click to bring up context menu
    pyautogui.rightClick(firstx, firsty)

    #Find the location of the 'Leave' button
    with concurrent.futures.ThreadPoolExecutor() as executor:
        th1 = executor.submit(checkLeave, firstx, firsty, leavePic)
        th2 = executor.submit(checkLeave, firstx, firsty, leavePic2)

        ret1 = th1.result()
        ret2 = th2.result()

        pyautogui.rightClick(firstx, firsty)
        
        if ret1 is not None or ret2 is not None:
            return True
        else:
            return False
    
#Threaded task to find the location of the 'Hide' button
def checkPos(firstx, firsty, pic):
    img = pyautogui.screenshot(region=(firstx, firsty, 300, 300))
    hideLoc = pyautogui.locate(pic, img)
    return hideLoc

#Threaded task to find the location of the 'Leave' button that indicates a group chat
def checkLeave(firstx, firsty, pic):
    img = pyautogui.screenshot(region=(firstx, firsty, 300, 300))
    leaveLoc = pyautogui.locate(pic, img)
    return leaveLoc

#Function to hide the chat in MS Teams (used for option '2')
def hide(firstx, firsty):
    
    #Left-click the hide button
    def click(hidex, hidey):
        orgPos = pyautogui.position()    
        pyautogui.click(hidex, hidey)

        orgPosx, orgPosy = orgPos
        pyautogui.moveTo(orgPosx, orgPosy)

        #time.sleep(0.5)
    
    #Right-click to bring up context menu
    orgPos = pyautogui.position()

    pyautogui.rightClick(firstx, firsty)

    orgPosx, orgPosy = orgPos
    pyautogui.moveTo(orgPosx, orgPosy)

    #Find the location of the 'Hide' button
    with concurrent.futures.ThreadPoolExecutor() as executor:
        th1 = executor.submit(checkPos, firstx, firsty, hidePic)
        th2 = executor.submit(checkPos, firstx, firsty, hidePic2)
        th3 = executor.submit(checkPos, firstx, firsty, hidePic3)

        ret1 = th1.result()
        ret2 = th2.result()
        ret3 = th3.result()
        
        if ret1 is not None:
            hidex, hidey, hidew, hideh = ret1
            click((firstx + hidex) + (hidew / 2), (firsty + hidey) + (hideh / 2))
        if ret2 is not None and ret1 is None:
            hidex, hidey, hidew, hideh = ret2
            click((firstx + hidex) + (hidew / 2), (firsty + hidey) + (hideh / 2))
        if ret3 is not None and ret2 is None and ret1 is None:
            hidex, hidey, hidew, hideh = ret3
            click((firstx + hidex) + (hidew / 2), (firsty + hidey) + (hideh / 2))
                  
        
#Function to start the process for clearing all chats that are NOT new (incl. group chats)
def clearNotNew(count, ver):
    while count > 0:
        #Find the location of the 'Recent' tag in MS Teams
        recentLoc = pyautogui.locateCenterOnScreen(recentPic)
        recentx, recenty = recentLoc
    
        #Set the coords for the first entry (a.k.a 20 over & 30 down)
        firstx = recentx + 20
        firsty = recenty + 25
    
        repeat = True
    
        while repeat == True:
            new = isNew(firstx, firsty, ver)
            if new == False:
                repeat = False
                hide(firstx, firsty)
            else:
                repeat = True
                firsty += 50

        count -= 1

def clearOldNotGroup(count, ver):
    while count > 0:
        #Find the location of the 'Recent' tag in MS Teams
        recentLoc = pyautogui.locateCenterOnScreen(recentPic)
        recentx, recenty = recentLoc
    
        #Set the coords for the first entry (a.k.a 20 over & 30 down)
        firstx = recentx + 20
        firsty = recenty + 25

        repeat = True
    
        while repeat == True:
            new = isNew(firstx, firsty, ver)
            if new == False:
                group = isGroup(firstx, firsty)
                if group == False:
                    repeat = False
                    hide(firstx, firsty)
                else:
                    repeat = True
                    firsty += 50 
            else:
                repeat = True
                firsty += 50

        count -= 1

def start(ver):
    #print('Press 1 to remove all chats, including new & group chats \nPress 2 to remove all non-new chats including groups \nPress 3 to remove all non-new & non-group chats')
    #_in = input()

    _in = pyautogui.confirm('Please select an option below', 'Select', buttons=['Hide all chats', 'Hide all read (singular & group chats)', 'Hide all read (only singular chats)'])
    if _in is None:
        sys.exit('No input supplied, closing...')
        
    #print('\nHow many runs do you want to perform (a.k.a how many chats do you want to hide?)')
    #count = input()

    count = pyautogui.prompt('How many loops do you want the script to perform (i.e. how many chats to remove)', 'Count', '1')
    if count is None or int(count) < 1:
        sys.exit('Input is null or less than 1, closing...')
    elif int(count) > 50:
        pyautogui.alert('The script will only loop 50 times for security reasons. Please re-run script if necessary.')
        count = 50

    print(str(_in))
    print(str(count))
    
    if _in == 'Hide all chats':
        clearAll(int(count))
    elif _in == 'Hide all read (singular & group chats)':
        clearNotNew(int(count), ver)
    elif _in == 'Hide all read (only singular chats)':
        clearOldNotGroup(int(count), ver)
    else:
        print('Incorrect value\n\n')
        start()
        
start(ver)
