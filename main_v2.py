import os

# ---------------------------- FUNCTIONS ------------------------------- #
exec( open( f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\functions_v2.py", encoding = "utf-8" ).read() )



# ---------------------------- CONSTANTS ------------------------------- #
automations_filepath = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\data\\automations.csv"
permissions_filepath = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\data\\permissions.csv"
userid = os.getlogin().lower()

Auto = Automation()
process = None
workingfiles = []  # uploaded files by user

automations = pd.read_csv( automations_filepath )
permissions = pd.read_csv( permissions_filepath )

permissions = permissions[ permissions["USERID"].str.lower() == userid ]
automations = automations[ automations["DISPLAY_NAME"].isin( permissions["AUTOMATION"] ) ]

pythonautomations = automations[ automations["TYPE"] == "Python" ].copy( deep = True )
pwrappautomations = automations[ automations["TYPE"] == "App"    ].copy( deep = True )



# ---------------------------- UI SETUP ------------------------------- #

# +--------+
# | WINDOW |
# +--------+

window = Tk()
window.title( "Billing Management" )
window.config( padx = 10, pady = 10, bg = "white" )
window.geometry( "715x460" )
window.iconphoto( False, PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\images\\singtel_logo.png" ) )
window.option_add("*Font", "Tahoma 10")

Label( text = "👋Welcome!", bg = "white", font = ("Tahoma", 20, "bold") ).grid( row = 0, column = 0, sticky = W )



# +--------+
# | IMAGES |
# +--------+

checkmark_img  = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\images\\checked.png" )
crossmark_img  = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\images\\cancel.png" )
browse_img     = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\images\\browse_folder.png" )
creditnote_img = PhotoImage( file = f"{os.environ['USERPROFILE']}\\OneDrive - Singtel\\Digital Team\\Automations\\GUInevere\\images\\diamond.png" )



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

notebook.enable_traversal() # use Alt and the underlined char to select the tab

style = ttk.Style( window )
style.configure( "lefttab.TNotebook", tabposition = "wn" )
style.configure( "TNotebook.Tab", padding = [20, 10, 20, 10] )

notebook.grid( row = 1, column = 0, pady = 10 )


# +--------------+
# | PYTHON FRAME |
# +--------------+

# STEP 1
Label( python_frame, text = "1." ).grid( row = 0, column = 0 )
Label( python_frame, text = "Select your automation" ).grid( row = 0, column = 1, sticky = W, columnspan = 3 )

# PYTHON COMBO BOX
python_list = list( pythonautomations["DISPLAY_NAME"] )
python_combobox = ttk.Combobox( python_frame, values = python_list )
python_combobox.bind( "<<ComboboxSelected>>", pythonComboboxSelected )
python_combobox.bind("<Return>", pythonComboboxTyped )
python_combobox.bind("<KeyRelease>", checkKey)
python_combobox.grid( row = 1, column = 1, sticky = EW, columnspan = 2, pady = 4 )

# OPEN WORKING FOLDER BUTTON
workfolder_button = Button( python_frame, text = "Work Folder", font = ("Tahoma", 9), command = openWorkFolder )
workfolder_button.grid( row = 1, column = 3, padx = 5 )
Hovertip( workfolder_button, "Open working folder" )

# CHECKMARK INDICATOR (SELECTED AUTOMATION)
width = 25  # Set the desired width of the image
height = 25  # Set the desired height of the image
checkmark_img = checkmark_img.subsample( checkmark_img.width() // width, checkmark_img.height() // height )
crossmark_img = crossmark_img.subsample( crossmark_img.width() // width, crossmark_img.height() // height )

automation_checkmark = Label( python_frame, image = None )
automation_checkmark.grid( row = 1, column = 7, sticky = (N, W), padx = 5 )

# INVALID AUTOMATION TEXT
automation_comment = Label( python_frame, text = "" )
automation_comment.grid( row = 2, column = 1, sticky = W, columnspan = 3 )

# STEP 2
Label( python_frame, text = "2." ).grid( row = 3, column = 0 )
Label( python_frame, text = "Upload files (if any)" ).grid( row = 3, column = 1, sticky = W, columnspan = 3 )

# UPLOAD LISTBOX
uploaded_listbox = Listbox( python_frame, width = 60, height = 10 )
uploaded_listbox.grid( row = 4, column = 1, rowspan = 2, columnspan = 5, sticky = NSEW )

# UPLOAD SCROLLBAR
upload_scrollbar = Scrollbar( python_frame, orient = VERTICAL, command = uploaded_listbox.yview )
upload_scrollbar.grid( row = 4, column = 6, rowspan = 2, sticky = NS )

uploaded_listbox.configure( yscrollcommand = upload_scrollbar.set )

# CHECKMARK INDICATOR (UPLOADED FILES)
upload_checkmark = Label( python_frame, image = None )
upload_checkmark.grid( row = 4, column = 7, sticky = (N, W), padx = 5 )

# BROWSE BUTTON
width = 20
height = 20
browse_img = browse_img.subsample( browse_img.width() // width, browse_img.height() // height )
browse_button = Button( python_frame, width = 76, height = 24, image = browse_img, highlightthickness = 0, command = uploadFile )
browse_button.grid( row = 6, column = 1, pady = 5, sticky = EW )
Hovertip( browse_button, "Browse" )

# REMOVE FILE BUTTON
remove_button = Button( python_frame, width = 10, text = "Remove", command = removeFile )
remove_button.grid( row = 6, column = 2, sticky = EW, padx = 5 )

# CLEAR ALL BUTTON
clear_button = Button( python_frame, width = 10, text = "Clear All", command = clearUploadedFiles )
clear_button.grid( row = 6, column = 3, sticky = EW )

# RENAME BUTTON
rename_button = Button( python_frame, width = 10, text = "Rename", command = renameFile )
rename_button.grid( row = 6, column = 4, sticky = EW, padx = 5 )

Frame( python_frame, width = 90 ).grid( row = 6, column = 5 )

# RUN BUTTON
run_button = Button( python_frame, text = "RUN", width = 7, font = ("Tahoma", 10, "bold"), command = runAndLaunch, state = "disabled" )
run_button.grid( row = 6, column = 7, padx = 10 )

# CANCEL BUTTON
cancel_button = Button( python_frame, text = "CANCEL", width = 7, font = ("Tahoma", 10, "bold"), command = cancelProcess, state = "disabled" )
cancel_button.grid( row = 7, column = 7, padx = 10 )

# REQUEST ACCESS
Button( window, text = "Request Access", command = openRequestURL ).grid( row = 2, column = 0, sticky = W )



# +------------+
# | APPS FRAME |
# +------------+

app_scrollbar = ttk.Scrollbar( pwrapp_frame, orient = VERTICAL)
app_canvas = Canvas( pwrapp_frame, yscrollcommand = app_scrollbar.set )
app_scrollbar["command"] = app_canvas.yview

app_canvas.grid( row = 0, column = 0, columnspan = 4, sticky = NSEW )
app_scrollbar.grid( row = 0, column = 5, sticky = (N, S) )
pwrapp_frame.grid_columnconfigure( 0, weight = 1 )
pwrapp_frame.grid_rowconfigure( 0, weight = 1 )

app_list = processApps( pwrappautomations )

# Label( app_canvas, text = "Credit Note" ).grid( row = 1, column = 0 )



# +------------------+
# | DASHBOARDS FRAME |
# +------------------+

# Button( dshbrd_frame, text = "dashboard", command = launchDashboard ).grid( row = 0, column = 0 )

window.mainloop()