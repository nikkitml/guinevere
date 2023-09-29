import os

# ---------------------------- FUNCTIONS ------------------------------- #
exec( open( f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\functions_v3.py", encoding = "utf-8" ).read() )



# ---------------------------- CONSTANTS ------------------------------- #
automations_filepath = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\data\\automations.xlsx"
permissions_filepath = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\data\\permissions.xlsx"
userid = os.getlogin().lower()

Auto = Automation()
process = None
workingfiles = []  # uploaded files by user

automations = pd.read_excel( automations_filepath )
permissions = pd.read_excel( permissions_filepath )

permissions = permissions[ permissions["USERID"].str.lower() == userid ]
automations = automations[ automations["DISPLAY_NAME"].isin( permissions["AUTOMATION"] ) ]

pythonautomations = automations[ automations["TYPE"] == "Python" ][["TYPE", "DISPLAY_NAME", "CODE_PATH", "WORKING_FOLDER", "ICON"]].copy( deep = True )
pwrappautomations = automations[ automations["TYPE"] == "App"    ][["TYPE", "DISPLAY_NAME", "CODE_PATH", "WORKING_FOLDER", "ICON"]].copy( deep = True )

guiconfig = configparser.ConfigParser()
guiconfig.read( f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\settings.ini" )
# colors = getColorTheme( guiconfig )

# colors = ["#CAEDFF", "#E3DFFD", "#D8B4F8", "#E3DFFD"]  # purple blue
# colors = ["#DEF5E5", "#BCEAD5", "#9ED5C5", "#E3DFFD"]  # seafoam
# colors = ["#FFFFFF", "#FFFFFF", "#FFFFFF", "#FFFFFF"]  # white
# colors = ["#FDF7C3", "#FFDEB4", "#FFB4B4", "#D0BFFF"]  # macaron
colors = ["#2CD3E1", "#A459D1", "#F266AB", "#FFB84C"]  # solid bright

# ---------------------------- UI SETUP ------------------------------- #

# +--------+
# | WINDOW |
# +--------+

window = Tk()
window.title( "Billing Management" )
window.config( padx = 10, pady = 10, bg = colors[0] )
window.geometry( "715x460" )
window.minsize( 715, 460 )
window.iconphoto( False, PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\images\\singtel_logo.png" ) )
window.option_add( "*Font", "Tahoma 10")

Label( text = "Welcome!", bg = colors[0], font = ("Gill Sans", 20, "bold") ).pack( anchor = W )


# +--------+
# | IMAGES |
# +--------+

checkmark_img  = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\images\\checked.png" )
crossmark_img  = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\images\\cancel.png" )
browse_img     = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\images\\browse_folder_2.png" )
creditnote_img = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\- Shared\\GUInevere\\images\\diamond.png" )



# +----------+
# | NOTEBOOK |
# +----------+

notebook = ttk.Notebook( window, width = 570, height = 345, style = "lefttab.TNotebook" )

python_frame = ttk.Frame( notebook )
notebook.add( python_frame, text = "    Python    ", underline = 4 ) #  pad with spaces

pwrapp_frame = ttk.Frame( notebook )
notebook.add( pwrapp_frame, text = "      Apps      ", underline = 6 ) #  pad with spaces

dshbrd_frame = ttk.Frame( notebook )
notebook.add( dshbrd_frame, text = "Dashboards", underline = 0 ) #  text should be 10 or less characters

status_frame = ttk.Frame( notebook )
notebook.add( status_frame, text = "     Status     ", underline = 5 )

notebook.enable_traversal() # use Alt and the underlined char to select the tab

style = ttk.Style( window )
style.configure( "lefttab.TNotebook" , tabposition = "wn"         )
style.configure( "TNotebook.Tab"     , padding = [20, 10, 20, 10] )
style.configure( "TFrame"            , background = colors[2]     )
style.configure( "TNotebook"         , background = colors[2]     )

# style.theme_use('default')
# style.configure('TNotebook.Tab', background="Red")
# style.map("TNotebook", background= [("selected", "red")])
# cannot change background color of notebook tabs unless 

notebook.pack( side = TOP, expand = YES, fill = BOTH, pady = 10 )



# +--------------+
# | PYTHON FRAME |
# +--------------+

# STEP 1
Label( python_frame, text = "1.", bg = colors[2] ).grid( row = 0, column = 0 )
Label( python_frame, text = "Select your automation", bg = colors[2] ).grid( row = 0, column = 1, sticky = W, columnspan = 3 )

# PYTHON COMBO BOX
python_list = list( pythonautomations["DISPLAY_NAME"] )
python_combobox = ttk.Combobox( python_frame, values = python_list )
python_combobox.focus_set()
python_combobox.bind( "<<ComboboxSelected>>", pythonComboboxSelected )
python_combobox.bind("<Return>", pythonComboboxTyped )
python_combobox.bind("<KeyRelease>", checkKey)
python_combobox.grid( row = 1, column = 1, sticky = EW, columnspan = 3, pady = 4 )

# OPEN WORKING FOLDER BUTTON
workfolder_button = Button( python_frame, text = "Work Folder", font = ("Tahoma", 9), command = openWorkFolder, bg = colors[2] )
workfolder_button.grid( row = 1, column = 4, padx = 5 )
Hovertip( workfolder_button, "Open working folder" )

python_frame.columnconfigure(5, weight = 1)  # Set weight to 1 to allow expansion
python_frame.rowconfigure(4, weight = 1)  # Set weight to 1 to allow expansion

# CHECKMARK INDICATOR (SELECTED AUTOMATION)
width = 25  # Set the desired width of the image
height = 25  # Set the desired height of the image
checkmark_img = checkmark_img.subsample( checkmark_img.width() // width, checkmark_img.height() // height )
crossmark_img = crossmark_img.subsample( crossmark_img.width() // width, crossmark_img.height() // height )

automation_checkmark = Label( python_frame, image = None, bg = colors[2] )
automation_checkmark.grid( row = 1, column = 7, sticky = (N, W), padx = 5 )

# INVALID AUTOMATION TEXT
automation_comment = Label( python_frame, text = "", bg = colors[2] )
automation_comment.grid( row = 2, column = 1, sticky = W, columnspan = 3 )

# STEP 2
Label( python_frame, text = "2.", bg = colors[2] ).grid( row = 3, column = 0 )
Label( python_frame, text = "Upload files (if any)", bg = colors[2] ).grid( row = 3, column = 1, sticky = W, columnspan = 3 )

# UPLOAD LISTBOX
uploaded_listbox = Listbox( python_frame, width = 60, height = 10 )
uploaded_listbox.grid( row = 4, column = 1, rowspan = 2, columnspan = 5, sticky = NSEW )

# UPLOAD SCROLLBAR
upload_scrollbar = Scrollbar( python_frame, orient = VERTICAL, command = uploaded_listbox.yview )
upload_scrollbar.grid( row = 4, column = 6, rowspan = 2, sticky = NS )

uploaded_listbox.configure( yscrollcommand = upload_scrollbar.set )

# CHECKMARK INDICATOR (UPLOADED FILES)
upload_checkmark = Label( python_frame, image = None, bg = colors[2] )
upload_checkmark.grid( row = 4, column = 7, sticky = (N, W), padx = 5 )

# BROWSE BUTTON
width = 20
height = 20
browse_img = browse_img.subsample( browse_img.width() // width, browse_img.height() // height )
browse_button = Button( python_frame, width = 76, height = 24, image = browse_img, highlightthickness = 0, command = uploadFile, state = "disabled", bg = colors[1] )
browse_button.grid( row = 6, column = 1, pady = 5, sticky = EW )
Hovertip( browse_button, "Browse" )

# REMOVE FILE BUTTON
remove_button = Button( python_frame, width = 10, text = "Remove", command = removeFile, bg = colors[2] )
remove_button.grid( row = 6, column = 2, sticky = EW, padx = 5 )

# CLEAR ALL BUTTON
clear_button = Button( python_frame, width = 10, text = "Clear All", command = clearUploadedFiles, bg = colors[2] )
clear_button.grid( row = 6, column = 3, sticky = EW )

# RENAME BUTTON
rename_button = Button( python_frame, width = 10, text = "Rename", command = renameFile, bg = colors[2] )
rename_button.grid( row = 6, column = 4, sticky = EW, padx = 5 )

Frame( python_frame, width = 90, bg = colors[2] ).grid( row = 6, column = 5 )

# RUN BUTTON
run_button = Button( python_frame, text = "RUN", width = 7, font = ("Tahoma", 10, "bold"), command = runAndLaunch, state = "disabled", bg = colors[1] )
run_button_menu = Menu( run_button, tearoff = 0 )
run_button_menu.add_command( label = "Bypass checks", command = clickBypassChecks )
run_button.bind('<Button-3>', rightClickRunButton)
run_button.grid( row = 6, column = 7, padx = 10 )

# CANCEL BUTTON
cancel_button = Button( python_frame, text = "CANCEL", width = 7, font = ("Tahoma", 10, "bold"), command = cancelProcess, state = "disabled", bg = colors[1] )
cancel_button.grid( row = 7, column = 7, padx = 10, pady = (0, 10) )

# REQUEST ACCESS
Button( window, text = "Request Access", command = openRequestURL, bg = colors[0] ).pack( anchor = W )



# +------------+
# | APPS FRAME |
# +------------+

app_canvas_width = 559
l = drawAppsFrame()
pwrapp_frame.bind("<Control-r>", refreshAppsShortcut)



# +------------------+
# | DASHBOARDS FRAME |
# +------------------+

# Button( dshbrd_frame, text = "dashboard", command = launchDashboard ).grid( row = 0, column = 0 )



# +---------------------+
# | REPORT STATUS FRAME |
# +---------------------+

# vFullPath = r"C:\Users\P1347618\OneDrive - Singtel\Digital Team\Automations\Reports Downloading\VisualizeReports - - Outlook.htm"
# with   open( vFullPath , 'r' , encoding='utf-8' )   as   oFile   :   vFileContents = oFile.read()

# html_content = """
# <html>
# <head></head>
# <body>
# <h1>Hello, HTML in Tkinter!</h1>
# <p>This is a demonstration of displaying HTML content in a Tkinter canvas.</p>
# </body>
# </html>
# """

# canvas = Canvas( status_frame )
# canvas.pack( fill = BOTH )
# html_label = HTMLLabel(canvas, html=html_content)
# canvas.create_window(10, 10, window=html_label, anchor="nw")

window.mainloop()