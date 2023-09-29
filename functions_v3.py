import configparser
from idlelib.tooltip import Hovertip
import pandas as pd
import re
import shutil
import subprocess
import sys
import time
import threading
# from tkhtmlview import HTMLLabel
from tkinter import *
from tkinter import messagebox, ttk, filedialog
from tkinter.filedialog import askopenfile
from tkinter.simpledialog import askstring
import webbrowser



# When "Request access" button is clicked, open URL to Project Tracker App
def openRequestURL():

    webbrowser.open( "https://apps.powerapps.com/play/e/default-beb276ac-6e9f-498e-8e31-019ee666decd/a/f6198d31-5862-444c-a83f-e18b8ccaa526?tenantId=beb276ac-6e9f-498e-8e31-019ee666decd&sourcetime=2023-08-16%2007%3A11%3A47Z" )

    return None



# +--------------+
# | PYTHON FRAME |
# +--------------+

class Automation:

    def __init__( self ):

        self.name = ""          # Name of automation
        self.inipath = ""       # Path to ini file with regex of file names
        self.config = None      # Regex of file names
        self.codepath = ""      # Path to py file / url
        self.workpath = ""      # Path to working folder

    def getName( self ):

        return self.name

    def getIniPath( self ):

        return self.inipath

    def getCodePath( self ):

        return self.codepath

    def getWorkPath( self ):

        return self.workpath

    def getConfig( self ):

        return self.config

    def setName( self, name ):

        self.name = name

    def setIniPath( self, inipath ):

        self.inipath = inipath

    def setCodePath( self, codepath ):

        self.codepath = codepath

    def setWorkPath( self, workpath ):

        self.workpath = workpath


    def setConfig( self, config ):

        d = {}

        if int(config["n"]["n"]) == 0: pass

        else:

            for i in range(1, int(config["n"]["n"]) + 1):

                t = []

                for j in range( 1, int(config[str(i)]["i"]) + 1 ):

                    t.append(config[str(i)]["file" + str(j)])

                d[i] = t

        self.config = d


    def setProperties( self, name, inipath, codepath, workpath ):

        self.setName( name )
        self.setIniPath( inipath )
        self.setCodePath( codepath )
        self.setWorkPath( workpath )

        # Set file requirements
        config = configparser.ConfigParser()
        config.read( self.getIniPath() )

        try:

            self.setConfig( config )

        except KeyError: pass

        return None


    def runAutomation( self ):

        global workingfiles, process

        # Move uploaded files to working folder
        for file in workingfiles:

            src = file.replace("/", "\\\\")
            dst = os.path.join( self.getWorkPath(), file.rsplit( "/", 1 )[1] )

            try:

                shutil.copy( src, dst )

            except FileExistsError: pass

            except shutil.SameFileError: pass

            except PermissionError:

                messagebox.showwarning(title = "Error", message = "Please close all files used by the RPA." )

        if self.getCodePath() == "":

            print( "Code path not initialized" )

        else:

            codepath = os.path.join( os.environ["USERPROFILE"], self.getCodePath() )
            process = subprocess.Popen( ['python', codepath] )

            process.wait()

            # Get the status
            status = process.poll()
            print(f"Status: {status}")

            if status is None:

                print("Process is still running.")

            else:

                print("Process is done.")
                cancel_button.config( state = "disabled", bg = colors[1] )
                run_button.config( state = "normal", bg = colors[2] )

        return None



def pythonComboboxSelected( event ):

    selection = python_combobox.get()

    if selection == "":

        automation_checkmark.config( image = crossmark_img )
        upload_checkmark.config( image = "" )
        browse_button.config( state = "disabled", bg = colors[1] )

    else:

        inipath  = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["CODE_PATH"])[0].rsplit("\\", 1)[0] + "\\filecheck.ini")
        codepath = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["CODE_PATH"])[0] )

        try:

            workpath = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["WORKING_FOLDER"])[0] )

        except TypeError:

            workpath = ""

        Auto.setProperties( selection, inipath, codepath, workpath )

        # Green tick
        automation_checkmark.config( image = checkmark_img )
        automation_comment.config( text = "" )
        browse_button.config( state = "normal", bg = colors[2] )
        requiredFilesCheck( workingfiles, Auto.getConfig() )

    return None



def pythonComboboxTyped( event ):

    selection = python_combobox.get()

    if selection == "":

        automation_checkmark.config( image = "" )
        automation_comment.config( text = "" )
        upload_checkmark.config( image = "" )
        browse_button.config( state = "disabled", bg = colors[1] )

        Auto.setName( selection )

    # check for match when all lower cased
    elif sum(pythonautomations["DISPLAY_NAME"].str.lower() == selection.lower()) == 1:

        selection = list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["DISPLAY_NAME"])[0]
        python_combobox.set( selection )
        python_combobox.icursor( END )
        automation_checkmark.config( image = checkmark_img )
        automation_comment.config( text = "" )
        browse_button.config( state = "normal", bg = colors[2] )

        inipath  = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["CODE_PATH"])[0].rsplit("\\", 1)[0] + "\\filecheck.ini")
        codepath = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["CODE_PATH"])[0] )

        try:

            workpath = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["WORKING_FOLDER"])[0] )

        except TypeError:

            workpath = ""

        Auto.setProperties( selection, inipath, codepath, workpath )

    elif sum( pythonautomations["DISPLAY_NAME"].str.lower().str.contains( selection.lower() ) ) == 1:

        selection = list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["DISPLAY_NAME"])[0]
        python_combobox.set( selection )
        python_combobox.icursor( END )
        automation_checkmark.config( image = checkmark_img )
        automation_comment.config( text = "" )
        browse_button.config( state = "normal", bg = colors[2] )

        inipath  = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["CODE_PATH"])[0].rsplit("\\", 1)[0] + "\\filecheck.ini")
        codepath = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["CODE_PATH"])[0] )

        try:

            workpath = os.path.join( os.environ["USERPROFILE"], list(pythonautomations[pythonautomations["DISPLAY_NAME"] == selection]["WORKING_FOLDER"])[0] )

        except TypeError:

            workpath = ""

        Auto.setProperties( selection, inipath, codepath, workpath )

    elif sum(pythonautomations["DISPLAY_NAME"] == selection) == 0:

        automation_checkmark.config( image = crossmark_img )
        browse_button.config( state = "disabled", bg = colors[1] )
        automation_comment.config( text = "Uh oh! No such automation.", fg = "#FFA8A8" )

    else: pass

    requiredFilesCheck( workingfiles, Auto.getConfig() )

    return None



def requiredFilesCheck( uploads, d ):

    if ((d == {}) or (d == None)) and (Auto.getName() != ""):

        run_button.config( state = "normal", bg = colors[2] )
        upload_checkmark.config( image = checkmark_img )
        automation_comment.config( text = "" )

    elif (Auto.getName() == ""):

        run_button.config( state = "disabled", bg = colors[1] )
        upload_checkmark.config( image = "" )

    else:

        for key in d.keys():

            flag = 1

            for pat in d[key]:

                if list( filter( lambda x: re.match(pat, x.rsplit("/", 1)[1]), uploads ) ) == []:

                    flag = 0

            if flag == 1:

                run_button.config( state = "normal", bg = colors[2] )
                upload_checkmark.config( image = checkmark_img )
                automation_comment.config( text = "" )
                break

            else:

                run_button.config( state = "disabled", bg = colors[1] )
                upload_checkmark.config( image = crossmark_img )

    return None



def checkKey( event ):

    selection = python_combobox.get()

    l = pythonautomations[pythonautomations["DISPLAY_NAME"].str.lower().str.contains( selection.lower() )]["DISPLAY_NAME"].copy( deep = True )

    if len(list(l)) == 0:

        python_combobox.configure( values = python_list )

    else:

        python_combobox.configure( values = list(l) )

    return None



def openWorkFolder():

    if Auto.getName() == "":

        messagebox.showwarning(title = "Error", message = "No automation selected" )

    else:

        t = Auto.getWorkPath().replace("\\\\", "\\")

        subprocess.Popen(f'explorer {t}')

    return None



# What to do when run_button is clicked
def runAndLaunch():

    global process

    if run_button.cget( "text" ) == "Launch":

        if Auto.getCodePath() == "":

            messagebox.showinfo( title = "Oops", message = "No application selected" )

        else:

            webbrowser.open( Auto.getCodePath() )

    else:

        if Auto.getCodePath() == "":

            messagebox.showinfo( title = "Oops", message = "No Python automation selected" )

        else:

            threading.Thread( target = Auto.runAutomation ).start()
            cancel_button.config( state = "normal", bg = colors[2] )
            run_button.config( state = "disabled", bg = colors[1] )

    return None



def rightClickRunButton( event ):

    try: run_button_menu.tk_popup(event.x_root, event.y_root)

    finally: run_button_menu.grab_release()

    return None



def clickBypassChecks():

    is_ok = messagebox.askokcancel( title = "Bypass file checks", message = "Incorrect/ no files were uploaded. Do you still wish to run the script?" )

    if is_ok:

        runAndLaunch()



# What to do when browse button is clicked
def uploadFile():

    global workingfiles

    file_paths = filedialog.askopenfilenames( title = "Select a file" )

    for item in file_paths:

        if item in workingfiles: pass

        else:

            workingfiles.append( item )
            uploaded_listbox.insert( END, item.rsplit("/", 1)[1] )

    requiredFilesCheck( workingfiles, Auto.getConfig() )

    return None



def removeFile():

    global workingfiles

    # Upload listbox is empty
    if uploaded_listbox.size() == 0: pass

    # Upload listbox is not empty
    else:

        # Remove selected file
        try:

            to_remove = uploaded_listbox.curselection()[0]

        except: pass

        else:

            temp = []

            for item in workingfiles:

                if item.rsplit( "/", 1 )[1] != uploaded_listbox.get( to_remove ):

                    temp.append( item )

            workingfiles = temp

            uploaded_listbox.delete( to_remove, to_remove )

            requiredFilesCheck( workingfiles, Auto.getConfig() )

    return None



def clearUploadedFiles():

    global workingfiles

    # Upload listbox is empty
    if uploaded_listbox.size() == 0: pass

    # Upload listbox is not empty
    else:

        is_ok = messagebox.askokcancel( title = "Clear uploads", message = "Do you wish to remove all the files you have uploaded so far?" )

        if is_ok:

            uploaded_listbox.delete( 0, END )
            workingfiles = []
            requiredFilesCheck( workingfiles, Auto.getConfig() )

    return None



def renameFile():

    try:

        index = uploaded_listbox.curselection()[0]
        to_rename = workingfiles[index]

    except IndexError:

        messagebox.showwarning(title = "Error", message = "No file selected!" )

        return None

    s0 = to_rename.rsplit("/", 1)[0]
    s1 = to_rename.rsplit("/", 1)[1]

    while True:

        newfilename = askstring( "Rename", f"Old file name: {s1}\nNew file name:" )

        if newfilename is None:

            return None

        else:

            newfilename = s0 + "/" + newfilename


        if bool(re.search(fr"\.{s1.rsplit('.', 1)[1]}$", newfilename)) == False:

            messagebox.showwarning(title = "Error", message = "Invalid file name" )
            continue

        else:

            os.rename( to_rename, newfilename )

            # Update workingfiles list with new file name
            workingfiles[index] = newfilename

            # Update uploaded_listbox with new file name
            uploaded_listbox.delete( index )
            newfilename = newfilename.rsplit("/", 1)[1]
            uploaded_listbox.insert( index, newfilename )
            uploaded_listbox.selection_set( index )
            break

    return None



def cancelProcess():

    global process

    try:

        process.terminate()
        process.wait( timeout = 2 )

        messagebox.showinfo( title = "RPA Cancelled", message = "Please close all automated Chrome browsers")

    except:

        # If the subprocess hasn't exited, kill it forcefully
        if process.poll() is None:

            process.kill()

    finally:

        print("\nProcess terminated by user")
        cancel_button.config( state = "disabled", bg = colors[1] )
        requiredFilesCheck( workingfiles, Auto.getConfig() )

        # Remove files copied by RPA to working folder
        for file in workingfiles:

            file = os.path.join( Auto.getWorkPath(), file.rsplit( "/", 1 )[1] )

            if os.path.exists( file ):

                os.remove( file )
                print(f"Removed: {file}")

    return None



# +------------+
# | APPS FRAME |
# +------------+

class App:

    global pwrapp_frame

    def __init__( self ):

        self.name = ""          # Name of app
        self.html = ""          # HTML link to site
        self.button = Button()  # Button object
        self.row = 0            # Row index
        self.col = 0            # Column index
        self.icon = None        # Icon

    def getName( self ):

        return self.name

    def getHTML( self ):

        return self.html

    def getButton( self ):

        return self.button

    def getRow( self ):

        return self.row

    def getCol( self ):

        return self.col

    def getIcon( self ):

        return self.icon

    def setName( self, name ):

        self.name = name

    def setHTML( self, html ):

        self.html = html

    def setButton( self, button, canvas ):

        self.button = button
        canvas.grid( row = self.getRow(), column = self.getCol() )
        button.pack( side = TOP, padx = 10, pady = 20 )
        Label( canvas, text = self.getName(), bg = colors[2] ).pack( side = TOP )

    def setRow( self, row ):

        self.row = row

    def setCol( self, col ):

        self.col = col

    def setIcon( self, icon ):

        self.icon = icon

    def openURL( self ):

        webbrowser.open( self.getHTML() )



def refreshApps():

    global app_canvas
    global app_scrollbar
    global refresh_apps_button
    global l

    app_canvas.destroy()
    app_scrollbar.destroy()
    refresh_apps_button.destroy()
    l = drawAppsFrame()

    return None



def refreshAppsShortcut( event ):

    refreshApps()

    return None



def drawAppsFrame():

    global app_canvas
    global app_scrollbar
    global refresh_apps_button
    global app_canvas_width
    global frame

    app_canvas = Canvas( pwrapp_frame, bg = colors[2] , highlightthickness = 0 )
    app_scrollbar = ttk.Scrollbar( pwrapp_frame, orient = "vertical", command = app_canvas.yview )
    refresh_apps_button = Button( pwrapp_frame, text = "Refresh" , command = refreshApps, bg = colors[2] )
    app_canvas.config( yscrollcommand = app_scrollbar.set )

    refresh_apps_button.pack( side = BOTTOM, anchor = E )
    app_scrollbar.pack( side = RIGHT, fill = Y )
    app_canvas.pack( side = LEFT, expand = YES, fill = BOTH )

    # Create a frame to contain the buttons
    frame = Frame( app_canvas, bg = colors[2] )
    app_canvas.create_window( (0, 0), window = frame, anchor = "nw" )

    app_list = displayApps( pwrappautomations )

    app_canvas.bind("<Configure>", onResize)
    frame.update_idletasks()
    app_canvas.config( scrollregion = app_canvas.bbox("all") )
    app_canvas.bind( "<MouseWheel>", onCanvasScroll )
    frame.bind("<MouseWheel>", onCanvasScroll)

    return app_list



def onCanvasScroll(event):

    app_canvas.yview_scroll(-1 * (event.delta // 120), "units")
    app_canvas.configure( scrollregion = app_canvas.bbox( "all" ) )

    return None



def onResize(event):

    global app_canvas_width
    global app_canvas

    app_canvas_width = event.width
    app_canvas.configure( scrollregion = app_canvas.bbox( "all" ) )

    return None



def displayApps( df ):

    global app_canvas_width

    num_columns = 5
    gridspace   = (app_canvas_width // num_columns) - 30

    width  = gridspace  # Set the desired width of the image
    height = gridspace  # Set the desired height of the image

    l = []

    for i in range(df.shape[0]):

        temp_canvas = Canvas( frame, width = gridspace, height = gridspace, bg = colors[2], highlightthickness = 0 )
        temp_canvas.bind("<MouseWheel>", onCanvasScroll)

        a = App()

        a.setName( df.iloc[i, 1] )
        a.setHTML( df.iloc[i, 2] )
        a.setRow( i // num_columns )
        a.setCol( i %  num_columns )

        if str(df.iloc[i, 4]) == "nan":

            print("no icon")
            icon = None

        else:

            icon = PhotoImage( file = os.environ["USERPROFILE"].replace("\\", "\\\\") + "\\\\" + df.iloc[i, 4] )
            icon = icon.subsample( icon.width() // width, icon.height() // height )

        a.setIcon( icon )

        # ADD COMMAND
        b = Button( temp_canvas, text = a.getName(), image = a.getIcon(), command = a.openURL )
        b.bind("<MouseWheel>", onCanvasScroll)
        temp_canvas.grid( row = a.getRow(), column = a.getCol() )
        b.pack( side = TOP, padx = 10, pady = (20, 0) )
        appname = Label( temp_canvas, text = a.getName(), bg = colors[2] )
        appname.bind("<MouseWheel>", onCanvasScroll)
        appname.pack( side = TOP, anchor = "n" )

        l.append( a )

    return l



# +------------------+
# | DASHBOARDS FRAME |
# +------------------+

def launchDashboard():

    webbrowser.open( "https://app.powerbi.com/groups/me/reports/886e9511-78d4-4174-8b1a-2c7e364ad374?ctid=beb276ac-6e9f-498e-8e31-019ee666decd&pbi_source=linkShare" )

    return None