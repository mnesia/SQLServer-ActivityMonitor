

import pyodbc
import re
import gtk
import os
import sys
import xml.etree.ElementTree as et

# determine if application is a script file or frozen exe
#Use embedded gtkrc
if getattr(sys, 'frozen', False):
    basedir = os.path.dirname(sys.executable)
else:
    basedir = sys.path[0]

gtkrc = os.path.join(basedir, 'gtkrc')
gtk.rc_set_default_files([gtkrc])
gtk.rc_reparse_all_for_settings(gtk.settings_get_default(), True)

connectionString = ""
server = ""
getProcessesQuery = "select distinct spid, \
                            blocked,\
                            DB_NAME(dbid), \
                            cast(loginame as varchar), \
                            cast(hostname as varchar),\
                            cast(PROGRAM_NAME as varchar), \
                            convert(varchar(19),login_time,121), \
                            convert(varchar(19),last_batch,121), \
                            lastwaittype,  \
                            status \
            from master..sysprocesses \
            where hostprocess != '' \
                and spid <> @@spid"

columnList = ["Spid", "Blocked", "Database", "Login Name","Host Name", "Program Name","Login Time","Last Batch", "Last Wait Type", "Status"]


def displayPopup(message, description):
    em = gtk.Window()                                                   # create Window
    box = gtk.VBox(spacing=6)                                           # create vertical box
                                                                        #  - homogeneous = default = True => all child widgets are given equal space allocations
                                                                        #  - spacing = 6 =>  vertical space between child widgets in pixels    
    em.add(box)                                                         # add box to window
    md = gtk.MessageDialog(em,0,gtk.MESSAGE_WARNING,gtk.BUTTONS_CLOSE, message)
    md.format_secondary_text(description)
    md.run()
    md.destroy()    

def executeQuery(sqlQuery):
    global server
    global connectionString

    server = entryServer.get_text()
    connectionString = None

    # open xml config file
    fileParam = os.path.join(os.getcwd(), "param.xml")
    tree = et.parse(fileParam)
    root = tree.getroot()
    domain = root.find("domain").text

    # add full domain name
    if not re.match('.*'+re.escape(domain), server):
        server = server + '.' + domain

    # look for server in config file
    for s in root.findall('server'):
        if (s.get('name') + '.' + domain).upper() == server.upper():
            login = s.find("login").text
            password = s.find("password").text
            connectionString = 'DRIVER={SQL Server};SERVER='+server+';UID='+login+';PWD='+password+';'

    # if we don't find the server, let's use Windows authentication
    if not connectionString:
        connectionString = 'DRIVER={SQL Server};SERVER='+server+';Trusted_Connection=yes'

    try:
        queryResults = []
        conn = pyodbc.connect(connectionString)
        cur = conn.cursor()
        cur.execute(sqlQuery)
        
        if cur.description is not None:
            queryResults = cur.fetchall()
       
        cur.close ()
        conn.close ()
    except pyodbc.Error , err:
        print(err)
        displayPopup("Cannot execute query !",err[1])
        
    return queryResults

def create_columns(treeView):
    for c in range(len(columnList)):
    
        rendererText = gtk.CellRendererText()
        column = gtk.TreeViewColumn(columnList[c], rendererText, text = c)
        column.set_sort_column_id(c)    
        treeView.append_column(column)
    if connectionString != "":
        treeView.connect("row-activated", getProcessDetails)

def create_model(results):
    store = gtk.ListStore(str, str, str, str, str, str, str, str, str,str)

    for act in results:
        store.append([act[i] for i in range(len(columnList))])
    return store

def getProcessDetails(widget, row, col):
    
    model = widget.get_model()

    spid = str(model[row][columnList.index("Spid")])
    
    getProcessDetailsQuery = 'dbcc inputbuffer ('+spid+')'
    processDetails = executeQuery(getProcessDetailsQuery)
    print("["+server+"] : get details for process "+spid)

    textbuffer.set_text(processDetails[0][2])
    killbutton.set_sensitive(True)
    killbutton.set_label('kill '+spid+' ?')

    
def killButtonClicked(button):

    killProcessQuery = killbutton.get_label()[:-1]

    executeQuery(killProcessQuery)
    print("["+server+"] : "+killProcessQuery)
    fillProcessesList()
    
def refreshButtonClicked(button):
    if connectionString != "":
        if runnableButton.get_active():
            print("["+server+"] : refresh runnable...")
        else:
            print("["+server+"] : refresh all...")
        fillProcessesList()

def fillProcessesList():
    killbutton.set_sensitive(False)
    killbutton.set_label('Click Here')
    
    textbuffer.set_text('')
    store.clear()
    results = executeQuery(getProcessesQuery)

    for act in results:
        if runnableButton.get_active():
            if "runnable" in act[columnList.index("Status")] or "suspended" in act[columnList.index("Status")]:
                store.append([act[i] for i in range(len(columnList))])
        else:
            store.append([act[i] for i in range(len(columnList))])
	

def entryServerValidated(widget, entry):
    fillProcessesList()
    print ("\nServer selected : [" + server +"]")
    return connectionString


############ main window  ############

#   [am] (Window)
#       [vbox] (VBox)
#           [swsi] (Table]
#               [entryServer] (Entry)
#               [refreshButton] (Button)    |   [runnableButton] (CheckButton)
#           [vpaned] (VPaned)
#                [swam] (ScrolledWindow)
#                   [treeView] (TreeView)
#               [swsql] (ScrolledWindow)
#                   [textview] (TextView)
#           [killbutton] (Button)


am = gtk.Window()
am.set_size_request(900, 500)                                           # initial size
am.set_position(gtk.WIN_POS_CENTER)                                     # initial position
am.set_title("Processes...")                                            # window title
am.connect("destroy", gtk.main_quit)

vbox = gtk.VBox(False, 8)                                               # create vertical box
                                                                        #  - homogeneous = False => all child widgets are NOT given equal space allocations
                                                                        #  - spacing = 8 =>  vertical space between child widgets in pixels

vpaned = gtk.VPaned()                                                   # create a vertical paned window
vpaned.set_position(300)                                                # ets the position of the divider between the two panes
############ sub-window for server info ############
swsi = gtk.Table(2, 2, False)                                           # create a 2 x 2 table [swsi]
                                                                        #  - rows = 2 => the number of rows
                                                                        #  - columns = 2 => the number of columns
                                                                        #  - homogeneous = False => all table cells will NOT be the same size as the largest cell)

## server entry
entryServer = gtk.Entry(500)                                             # create entry (size = 50)
entryServer.set_size_request(100,-1)                                    # define size
                                                                        #  - width  = 100
                                                                        #  - height = -1 => no specification
entryServer.connect("activate", entryServerValidated, entryServer)      # if we validate => call 'entryServerValidated'
entryServer.set_text("(servername)")                                    # default text in entry
swsi.attach(entryServer, 0, 1, 0, 1, gtk.SHRINK, gtk.SHRINK)            # position within table [swsi] (= top left)

## refresh button
refreshButton =  gtk.Button(label="Refresh...")                         # create button (with label 'Refresh')
refreshButton.connect("clicked", refreshButtonClicked) # if we click => call 'refreshButtonClicked'
swsi.attach(refreshButton, 0, 1, 1, 2, gtk.FILL, gtk.FILL)              # position within table [swsi] (= bottom left)
                                  
## runnable check button
runnableButton = gtk.CheckButton("runnable")                            # create checkbox (with label 'runnable')
swsi.attach(runnableButton, 1, 2, 1, 2)                                 # position within table [swsi]

############ sub-window for processes ############
## describe
swam = gtk.ScrolledWindow()                                             # create scrolled window [swam]
swam.set_shadow_type(gtk.SHADOW_ETCHED_IN)                          
swam.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)             # scrolled bars appear only if needed

## fill
emptyResults = [("","","","","","","","","","")]
store = create_model(emptyResults)
treeView = gtk.TreeView(store)

treeView.connect("row-activated", getProcessDetails)                    # if we click => call 'getProcessDetails'

swam.add(treeView)                                                      # add TextView to ScrolledWindow
create_columns(treeView)

############ sub-window for sql queries description ############
## describe
swsql = gtk.ScrolledWindow()                                            # create scrolled window [swam]
swsql.set_shadow_type(gtk.SHADOW_ETCHED_IN)
swsql.set_policy(gtk.POLICY_AUTOMATIC, gtk.POLICY_AUTOMATIC)            # scrolled bars appear only if needed

## add textview
textview = gtk.TextView()                                               # create TextView
textbuffer = textview.get_buffer()
textview.set_editable(False)                                            # text is not editable
textview.set_wrap_mode(gtk.WRAP_WORD)                                   # wrap text using word
swsql.add(textview)                                                     # add TextView to ScrolledWindow

## add kill button
killbutton = gtk.Button(label="Click Here")                             # create button (with label 'Click Here')
killbutton.set_sensitive(False)                                         # by default, button not clickable
killbutton.connect("clicked", killButtonClicked)                        # if we click => call 'killButtonClicked'


# pack everybody

vbox.pack_start(swsi, False, False,  0)
vbox.pack_start(vpaned, True, True,  0)
vpaned.add1(swam)
vpaned.pack1(swam, True, True)
vpaned.add2(swsql)
vpaned.pack2(swsql, False, False)


vbox.pack_start(killbutton, False, False,  0)


am.add(vbox)
am.show_all()

gtk.main()
