# Teams Auto Hide
**This script is still in develpment, but this version is functional! See the 'Usage' section**

Python-based script made to automate the process of hiding old chat messages / conversations in Microsoft Teams.

## Dependencies
The dependencies below are automatically downloaded if not already installed. They use `subprocess` to run a `pip install <module>` command in the background (admin credentials are not required for this, however, your IT department may have other locks / blocking tools in place to prevent this). Once all dependencies are downloaded and installed, you **may** need to restart the application for it to launch. This is due to a timing issue I'm currently working on. [Issue #001](https://github.com/VoltaicGRiD/Teams-Auto-Hide/issues/1#issue-534361861)

- [pyautogui](https://github.com/asweigart/pyautogui)
- concurrent.futures



## Installation
**D**ownload a copy of th**e** zip file and extract it to **a** folder somewhere on your machine. Once the files are downloaded, you can create a shortcut to the `.pyw` file to pin to your **t**askbar, desktop, or w**h**erever you see fit.

If the dependencie**s** for the application have not ye**t** been installed on your machine, the application will try to install them. Due to a timing issue I'm currently working on, you may need to run the application a second time to see GUI and be able to use it.



## Usage
***PLEASE NOTE:*** This version of the script is built for Microsoft Teams running in 'Dark' mode ***ONLY***, changes are planned to implement 'Light' mode and 'High Contrast' mode.

Please ensu**r**e Microsoft Teams is at `100%` zoom, **and** it is open in full-screen on your main monitor. Currently, you are unable to perform background tasks while th**i**s script is runni**n**g due to the way it imitates mouse movement and clicks.

Upon start-up the application will prompt you with a **G**UI to allow you to pick which messages you want to hide (Microsoft Teams currently (12-05-2019) doesn't allow you to delete conversations like Skype for Business). The options are described below

- Hide all
  - This will remove a number of conversations, **regardless of what they are**. This includes new and unread conversations, group conversations, and anything else (**with the exception of pinned conversations**)
- Hide read (singular & group chats)
  - This will remove a number of conversations, **ignoring unread messages**, regardless of whether the conversation is a group chat or if its an individual conversation.
- Hide read (only singular chats)
  - This will remove a number of conversations, **ignoring unread _AND_ group conversations**. 
  
After selecting which option you'd like to perform, the application will prompt you to input a number. This will determine how many conversations the script will attempt to remove. Due to security purposes (this will not change) the program is limited to 50. While the script will allow an entry greater than 50, it will prompt you with a warning before proceeding. The 50-conversation limit can be changed by modifying the code if you desire.



## How it works
`pyautogui` has a built in funtion to search the screen for certain UI buttons (in this case images, you'll find the ones I used in the ImageMatches folder). This allows the tool to automatically navigate through the entries in Teams, right-click, then find the hide button. In the case of detecting group conversations, the tool will first search for a 'Leave' button within the context menu, if found, it will skip the conversation and continue on.

`pyautogui` is also capable of matching a screen pixel's color to a specified one. In the case of detecting unread chat messages, the tool looks the left side of the conversation, where a dot is located, indicating the chat has unread messages. If the dot matches the color profile in the code, it returns that conversation as 'New', thus, skipping over it.

Tons and tonnes (yes on purpose) of credit to the authors of [pyautogui](https://github.com/asweigart/pyautogui)



## Planned changes
- [ ] Add implementation for users running MS Teams in 'Light' mode
- [ ] Improve script speed
- [ ] Implement 'break-out' hotkey
- [ ] Find a way to allow for background / secondary monitor activites while script runs
- [ ] [Issue #002](https://github.com/VoltaicGRiD/Teams-Auto-Hide/issues/2#issue-534362157)


## Fun Things
##### Quotes
> Finally! A way to clear all this worthless clutter! ~ Me

> Such an easy way to sit back for a minute while it runs! ~ Also me

> That's really impressive. ~ Colleague also frustrated with no 'Hide All' button in MS Teams
