#!/usr/bin/env python

#Standard Selenium imports
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoAlertPresentException, NoSuchElementException, ElementNotVisibleException
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.common.keys import Keys

#Needed to set particular Firefox profile
from selenium.webdriver.firefox.webdriver import FirefoxProfile

#Needed to control Excel
from win32com.client import Dispatch
import pythoncom#used to determine whether Excel file is already open (part of pywin32)
import os

#Used to identify dates read from spreadsheet (pywintypes.TimeType)
import pywintypes

#Generally useful
import os,time, datetime

#Tkinter for alert messages
import Tkinter
import tkMessageBox

"""
Submits new applications in CCMS using data from specially formatted spreadsheet (ccms_application_and_page_data.xlsx)
 Alan Maydwell
 v0.1 - initial version

 v0.2   CCMS.treadmill updated so that when error encountered preventing progress
        the page source is added to the message returned.

        Fixed error in CCMS.excel_read_test_data(). Now selects chosen tab.
        Previously ignored tab argument and used hard coded "Page Details" tab.

        Added end-row reading to CCMS.excel_read_test_data()

 v0.3   Modified CCMS.excel_read_test_data() to remove leading/trailing spaces
        from data read from spreadsheet that are destined to be dictionary keys
        within CCMS.pi. Should mean fewer problems from leading/trailing spaces
        in spreadsheet data. Handling of values from spreadsheet still leaves
        leading/trailing spaces in place.

        Modified CCMS.choose_section() to return success/failure message rather
        than boolian. "ok" denotes success

        Modified CCMS.excel_run() to only call CCMS.treadmill() if the preceding
        call to CCMS.choose_section() returned "ok". This should give improved
        handling of when section attempted before its available.

        Added CCMS.pi_show() to export test data currently loaded. Modified
        CCMS.excel_run() to support new "expo" action that uses CCMS.pi_show()
        to export test data to file.


v0.4    New version, uses new spreadsheet
        CCMS.excel_run() now logs in and out to CCMS using details associated with
        new logi and logo actions. Reading of "Parameter 3" and "Parameter 4" added
        to support this activity.

        Modified CCMS.logout() to check for presence of logout link and to
        report whether it was found.

        Modified CCMS.login_cmms() (a) to accept optional Portal link parameter.
        (b) to return info message

        Added logged_in boolian attibute to CCMS class.

v 0.5   Added new "radio" radio button section to CCMS.fill_page(). Old ckick section
        can still be used but this radio-button specific section could be userful.

v 0.6   (1) Postcode search. Added CCMS.postcode_search() method which completes
        postcode search using details already entered. Modified CCMS.new_application_treadmill()
        to call postcode search when address screen reached and UK has been selected.

        (2) New applicant search parts of CCMS.new_application() modified because
        presence of "Register New Client" button cannot be relied upon to determine whether
        search completed. Need to handle situation when search returns
        more than 1 page of results (>20 records) as button will then
        be on final sub-page. Modified method to access the final sub-page when
        search returns more than one page of results.

v 0.7   Updated CCMS.treadmill() to be able to complete New Application section
        of PUI. Reorganised CCMS.treadmill() a bit.
        Removed CCMS.new_application_treadmill() as no longer needed.
        Removed CCMS.define_page_info() as it was not being used for anything.

        Increased timeout delays in CCMS.login_ccms()

v 0.8   Screenshot capture added
        CCMS.screenshot() method added
        CCMS.excel_run() responds to screenshot flag read from spreadsheet
        CCMS.treasmill() takes screenshots when screenshot folder set

        Excel filename can now be specified at the bottom of the script.

        Improved login error handling in CCMS.login_ccms()


v 0.9   (a)Other party "Organisation" now supported
        Added Special handling in treadmill for (u'Create Application', u'Opponents and Other Parties')
        because it has non-standard progress button choice.
        Updated CCMS.choose_section() to be able to select 'Opponents and Other Parties' section
        (b) Log file now includes extracted details of elements on each page. New CCMS.analyse_page()
        method added to get the details.


v 1.0   Modified part of CCMS.treadmill() that responds to 'class="buttonArea btnConfirm">'
        in page source. Previously this was taken to mean we'd reached the
        end of a section. However, this is not true for the Opponent Organisation
        section. Have modified script to not end iteration for this particular type.

        Restored "Already logged in." message to CCMS.login_ccms()

        Added support for selecting Firefox profile. CCMS.__init__() now
        accepts profile parameter. Example path near bottom of script.


v 1.1NP Adapted for new style PUI (but breaks compatibility with old)
        Updated CCMS.choose_section(), CCMS.login_ccms(), CCMS.excel_run(),
        CCMS.search() and CCMS.get_headings to work with new PUI design.
        Some of these changes break compatibility with the old PUI.
        This version is only expected to work with new.

v 1.2   Adapted to work with both old and new PUI
        * CCMS.logged_in changed from boolian to str. "" indicates not
        logged in, other values indicate different versions of PUI. Updated
        refrences to self.logged in to use designations. Value set by CCMS.login_ccms()
        and cleared by CCMS.logout_ccms()
        * Updated CCMS.choose_section() to use CCMS.logged_in valueto decide
        how to start new application
        Updated CCMS.search() to behave differently depending upon CCMS.logged_in

        Updated CCMS.login_ccms() Means of identifying PUI version changed to look
        for text in page source rather than using page title, as recent release
        made new and old PUIs have same page title.

        Modified CCMS.treadmill() to cope with "Submission Pending" page that appears when
        applicant submission fails.

        Modifed CCMS.fill_page() to print more details concerning success/failure
        of ddlist completion

        Updated CCMS.excel_read_test_data() to use .encode("utf-8") rather than str()
        when unicode data used to construct dictionary k1 and k2 values.

v1.3    Tweaked CCMSlogin_ccms(). Altered way Home page is recognised for both
        old and new PUI.

        New method CCMS.fill_page_auto(). Automatically identifies fields
        on page and can also complete them (only using very simple data choice).
        CCMS.excel_run() modified to read auto completion mode from spreadsheet
        and to pass it to CCMS.treadmill(). CCMS.treadmill() modified to call
        CCMS.fill_page_auto().

        Now new case creation can select existing applicant.
        Updated CCMS.treadmill() to support selecting an existing applicant. New
        CCMS.existing_applicant (int) attribute added, defaults to 0. Updated
        by CCMS.excel_run()to update CCMS.existing_applicant

        Added some NoSuchElementException handling to CCMS.fill_page()

v1.4    Added "auto" action which is used to change auto_mode value.
        CCMS.excel_run() modified to do this.

        HTML file now created to accompany screenshots.
        Class Logger added to help with screenshot HTML creation.
        Screenshot setup now moved from CCMS.excel_run() to new method
        CCMS.screenshot_setup(). Tweaks to CCMS.excel_run() and CCMS.treadmill()
        to add details to HTML file.

        Added two WebDriverWait()s for "Application Summary" in driver.page_source
        to CCMS.treadmill() to improve sync when returning to Application summary
        after completing section.


v1.5    Tick box completion by CCMS.fill_page() now takes existing state into account.
        Automatically identifies tick-boxes and will only tick if existing state
        does not match requested state.

v1.6
        Changed to work better with Beta PUI (but still not fully compatible)

        (i) Now supports direct login, bypasing Portal. CCMS.login_cmms() updated with
        direct Flag. Moved identification of CCMS version from CCMS.login_ccms() to new
        method CCMS.ccms_identify(). Added "dlog" action to CCMS.excel_run()

        (ii) Script now also idenfies beta PUI in addition to "new" and "old".
        CCMS.ccms_identify() updated to do this.Tweaked CCMS.choose_section()
        and CCMS.search() to handle beta PUI better

        (iii) iFrame Support for Beta - some Beta pages include iframes, which
        complicates navigation

        Added method CCMS.owd_frame() which checks for presence of iframe
        "owdFrame" and selects it if found.

        Modified CCMS.choose_section() to use this method

        Modified CCMS.treadmill()
        Checks whether page is "owd" page. Selects owd frame if present.
        Also has different "Next" button handling for owd pages.

        (iv) Sub-heading changes. Some Beta PUI pages present the page sub
        headings in a different way. We rely upon these to identify the treadmil
        pages. CCMT.get_headings() modified to be able to read this type of sub
        heading (in addition to the old style). BETA sub headings also can hold
        extra tabs, carriage returns and trailing/leading spaces. These
        are stippred out.

        (v) Re-aranged order of element type completion in CCMT.fill_page()
        so that text fields are completed only after drop-downs and buttons
        This might help script cope with dynamic elements of pages.

        Additional Fix
        --------------
        Modified CCMS.excel_read_test_data() to automatically convert dates
        to text
        if type(value) is pywintypes.TimeType:
            value = str(value)
        However, probably won't help.

        Added some ElementNotVisibleException handling to CCMT.fill_page()

1.7     Re-arranged handling of page data CCMT.pi dictionary so that it only
        has three levels of dictionary. Different element types are no longer
        stored separately and middle layer is now a list instead of dict.
        dictionaries for each element type. Outer dictionary still has
        same keys based on page headings but associated value is now a list.
        Each list contains separate dictionaries for each element within the page.
        For example

        self.pi[("Means","Income")] = [{"id":"salary","type":"field",value:"100"},{"id":"self_emp_false","type":"click",value:"this"}]

        Because lists retain order, page filling order should now reflect item order in spreadsheet.
        CCMT.excel_read_test_data(), CCMT.fill_page(),CCMT.pi_show()

        CCMT.fill_page() further updated so (a) it returns list of items it could
        not update because item was "hidden" at the time (b) Now has update_enabled
        bool argument which when True prevents page update so we just get new
        report of which items are hidden.

        CCMT.treadmill() modified so that it can call CCMT.fill_page() repeatedly
        for the same page depending upon details of hidden items returned by
        CCMT.fill_page()

        To do - replace owd_frame_check() with a new method that is used to
        select any frame or return to main part of page. Also add new attribute
        to CCMT class which holds name of currently selected frame. Modify
        CCMT.fill_page to  be able to use the new radio button true/false
        types in new iFrame pages.

1.8     Modified CCMS.read_summary() to read details in different way so it copes
        with new way details displayed in BETA. No longer assumes only one link
        per row in HTML table


        Modified CCMS.get_headings because some BETA pages have H1 and H2 tags
        reversed.

        Added driver.execute_script("return arguments[0].scrollIntoView();",ele)
        to CCMS.fill_page() otherwise iframe elements below page bottom aren't
        updated by Webdriver (even though no exception/error raised)

1.9     Added CCMS.radio_suffix_set() to set right suffix to radio button as
        new Beta pages use "t" and "f" rather than "_true" and "_false"

        Added new enter pass to access passported link in Beta.

        Modfied Next button pressing part of CCMS.treadmill() to look for
        multiple ("owd-submit") matches and click whichever is last (Keith
        found page where button press failed, seemingly because two matches
        one of which isn't the button). Also switched button press to send
        return key rather than simply click the button as it is supposed to be
        more reliable.

2.0     Added support for new "link" date type. This can be used to click hyperlinks
        in a page using data from the spreadsheet. Main changes in CCMS.fill_page()

2.1     KV Added new link for Means Oriinal & Means Upgraded in the section_mapping table

2.2     KV Caluclated and displayed elapsed time for each section


2.5     Added CCMS.stuck_page_retries attribute and corresponding argument.

        Modified CCMS.treadmill() so when struck on same page, further presses
        of the "Next" button will be made if stuck_page_retries value not reached.
        page_retry_count

2.6     CCMS.login_cmms modified to work with the new style portal (as well as the old)
"""

class CCMS:
    """Webdriver script that completes new application details in CCMS using parameters from a spreadsheet

    Treadmill part of script would originally stop processing a section if it
    found the same page returned after pressing the "Next" button. Now have
    attribute self.stuck_page_retries to specify number of retry attempts
    Args:
        excel_filename - filename of spreadsheet with test data
        ffprofile - Path to Firefox profile. Can be None for default Webdriver option.
        stuck_page_retries (int) - number of times to retry moving to next page
                            when stuck on same page
    """

    def __init__(self,excel_filename="ccms_application_and_page_data(lv0.2).xlsx",
                ffprofile=None,
                stuck_page_retries=0):
        #Tracks whether logged into CCMS. Updated by login_ccms() and logout_ccms() methods
        #empty string for not logged in, value for logged in to version
        self.logged_in = ""

        print"Start"
        self.pi = {}#stores detials to be entered in each page
        self.logfile = ""#logfile filename
        self.screenshot_folder = ""#folder name for screenshots
        self.today = time.strftime("%d/%m/%Y")#Today's date
        self.existing_applicant = 0 #When >0, used to identify existing from search to be used
        #Set Firefox profile
        self.firefox_profile = None
        if ffprofile:
            self.firefox_profile=FirefoxProfile(ffprofile)
        self.stuck_page_retries = stuck_page_retries
        #self.experimental_run()
        self.excel_run(filename = excel_filename)
        print "End"


    def experimental_run(self):
        """Used to try out stuff. Not used during typlical run."""

        #Read Excel data
        ##self.excel_run()
        ##return

        #Login
        url = "https://lsce11intport.lab.gov/"
        username = "Willylawyer061114"
        pw = "welcome03"
        self.login_ccms(url,username,pw)

        driver = self.driver

        #Search For Case
        case_ref = "300000045555"
        found = self.search(case_ref)
        print found

        #Provider Details
        self.choose_section('Provider Details')
        self.treadmill()

        #General Details
        self.choose_section('General Details')
        self.treadmill()

        #Proceedings and costs
        self.choose_section('Proceedings and Costs')
        ##WebDriverWait(self.driver,10).until(lambda driver:"Proceedings and Costs" in driver.page_source)
        #Add proceeding
        ##driver.find_element_by_link_text("Add Proceeding").click()
        ##WebDriverWait(self.driver,10).until(lambda driver:"Proceeding Details - Step 1 of 4" in driver.page_source)
        self.treadmill()

        #Opponents and Other Parties
        self.choose_section('Opponents and Other Parties','Individual')
        #Add party
        ##driver.find_element_by_link_text("Add Individual").click()
        ##WebDriverWait(self.driver,10).until(lambda driver:"Individual Contact Details" in driver.page_source)
        self.treadmill()

        """
        #Means Assessment
        self.choose_section('Means Assessment')
        self.treadmill()

        #Merits Assessment
        self.choose_section('Merits Assessment')
        self.treadmill()
        """
        #End
        self.logout_ccms()
        self.driver.quit()


    def excel_run(self,filename="ccms_application_and_page_data(lv0.1).xlsx"):
        """Run test using data from specially formatted Excel spreadsheet.

        e.g. "ccms_application_and_page_data.xlsx"
        """
        #Open Spreadsheet
        filename = os.path.join(os.getcwd(),filename)
        xl = Dispatch("Excel.Application")
        xl.Visible = True
        self.wb = xl.Workbooks.Open(filename)
        #select first worksheet in the spreadsheet
        ws1 = self.wb.Worksheets["Run Data"]

        #Test Data Spreadsheet Columns
        #Row with column names in spreadsheet
        headrow = 10
        #Expected column names from spreadsheet
        headnames = ["Date/Time","Message","Reusable","Skip?","Action","Parameter 1","Parameter 2","Parameter 3","Parameter 4"]
        #Find positions of requred columns in spreadsheet
        datacols = dict.fromkeys(headnames, -1)#Dictionary to hold mapping between heading name and column number
        for col in range(1,30):
            val = str(ws1.Cells(headrow,col).Value)
            if val in headnames:
                datacols[val]=col

        #Warn if any required columns not found
        not_found = [k for k in datacols if datacols[k]==-1]
        if not_found:
            print "WARNING - may fail as required heading(s) "+",".join(not_found)+" not found in spreadsheet row "+str(headrow)+"."

        #Setup screenshot taking if screenshots active
        if str(ws1.Cells(7,9).Value).strip().lower()[:1]=="y":
            screenshot_folder=time.strftime("Screenshots_%Y-%m-%d_%H.%M.%S")
            #Add folder name to spreadsheet
            ws1.Cells(6,9).Value = screenshot_folder
            #Setup screenshot folder and accompanying HTML file
            self.screenshot_setup(screenshot_folder,heading="CCMS PUI Screen Capture")
            #Holds info to be added to nav in html
            nav_data = []

        #Start log file if enabled in spreadsheet
        if str(ws1.Cells(5,9).Value).strip().lower()[:1]=="y":
            #Auto make filename
            self.logfile = time.strftime("ccms_log_%Y-%m-%d_%H.%M.%S.txt")
            #Add filename to spreadsheet
            ws1.Cells(4,9).Value = self.logfile
            #Add start time to log file
            message = "CCMS Application Selenium Script\n"
            message = message + "Start: "+time.strftime("%d-%m-%Y [%H:%M:%S]")+"\n"
            self.logadd(message,"w")

        #Read start and end row values
        startrow = int(ws1.Cells(4,6).Value)
        endrow = int(ws1.Cells(5,6).Value)
        print "Spreadsheet row range:",startrow,"-",endrow

        #Read "auto complete" mode
        try:
            auto_mode = int(ws1.Cells(7,6).Value)
        except (ValueError,TypeError) as e:
            print "Problem with auto_mode value from spreadsheet. Defaulting to 1"
            auto_mode = 1

        #Warn if endrow not > start row, or either is less than 1
        if endrow<startrow or endrow<1 or startrow<1:
            print"Problem with end row (",endrow,") and/or start row, (",startrow,") values."
            #Display alert message
            root = Tkinter.Tk().withdraw() #stop Tk main window appearing
            tkMessageBox.showwarning("CCMS Selenium Script","Cannot proceed."
                                     +"\nInvalid start row and end row values in spreadsheet."
                                     +"\nStart row: "+str(startrow)+"\nEnd row: "+str(endrow))

        #Loop through row range holding actions
        for row in range(startrow,endrow+1):
            #Default message text (empty)
            message =""
            #Only do something if row not skipped
            skip = str(ws1.Cells(row,datacols["Skip?"]).Value)
            if skip[:1].lower()!="y":
                #For row read action and associated paramters
                action = ws1.Cells(row,datacols["Action"]).Value
                action = str(ws1.Cells(row,datacols["Action"]).Value).strip()
                action = action.lower()[:4]
                param1 = ws1.Cells(row,datacols["Parameter 1"]).Value
                param2 = ws1.Cells(row,datacols["Parameter 2"]).Value
                param3 = ws1.Cells(row,datacols["Parameter 3"]).Value
                param4 = ws1.Cells(row,datacols["Parameter 4"]).Value
                ##print row, action, param1, param2

                #Run appropriate activity for action if it is a known type
                print "action:",action,param1

                #Read page-related test data from spreadsheet
                if action == "read":
                    #Defaults column F
                    if not param1:
                        param1="F"
                    #Default tab to second tab
                    if not param2:
                        param2=1
                    endrow = self.excel_read_test_data(value_column=param1,tab=param2)
                    message = "Data from tab "+str(param2)+", column "+str(param1)+", up to row "+str(endrow)

                #Set "Auto Mode"
                elif action =="auto":
                    try:
                        auto_mode = int(param1)
                    except (ValueError,TypeError) as e:
                        auto_mode = 1
                        message = "Auto mode conversion to integer failed. Defaulting to:"+str(auto_mode)
                    else:
                        message = "Auto mode set to :"+str(auto_mode)

                #Clear - wipe page-related test data already loaded
                elif action =="clea":
                    self.pi={}
                    message = "Cleared"

                #Export current test data using self.pi_show() method
                elif action == "expo":

                    #Separator - default to tab when None
                    sep = str(param2)
                    if sep=="None":
                        sep="\t"

                    #Filename
                    filename = str(param1)
                    #Make empty values (None) become "" rather than "None"
                    if filename =="None":
                        filename=""

                    message = self.pi_show(sep,filename)

                #Login to Portal
                elif action == "logi":

                    #Handle optional Portal link value
                    if type(param4) in (str, unicode):
                        link = param4
                    else:
                        link = "Client and Cost Management System (CCMS)"
                    #Call login method
                    message = self.login_ccms(url=param3,username=param1,password=param2,link=link)

                #Direct login, not using Portal
                elif action == "dlog":
                     message = self.login_ccms(url=param3,username=param1,password=param2,direct=True)

                #Search for application using case reference
                elif action =="sear":
                    if self.logged_in:
                        param1 = str(param1)
                        message = self.search(param1)
                    else:
                        message = "Not executed because not logged in to CCMS."

                #Logout from CCMS
                elif action =="logo":
                    if self.logged_in:
                        message = self.logout_ccms()
                    else:
                        message = "Not executed because not logged in to CCMS."

                #* Enter details into CCMS section*
                elif action == "ente":
                    #Only execute if logged in
                    if self.logged_in:
                    #Maps abreivated section names from spreadsheet to corresponding argument name passed to self.choose_section()
                    #For new PUI, mapping of "star" changed from "Start a New Application" to "New Application"
                        section_mapping={
                            "star":("New Application",""),
                            "prov":("Provider Details",""),
                            "clie":("Client Details",""),
                            "gene":("General Details",""),
                            "proc":("Proceedings and Costs",""),
                            "orga":("Opponents and Other Parties","Organisation"),
                            "indi":("Opponents and Other Parties","Individual"),
                            "meri":("Merits Assessment",""),
                            "mori":("Original",""),
                            "mupg":("Upgraded",""),
                            "mean":("Means Assessment",""),
                            "pass":("Passported","")
                            }

                        #Ensure sections are four character, lower-case strings
                        sections = [t.lower().strip()[:4] for t in str(param1).lower().split(",")]

                        for s in sections:
                            #Update CCMS.existing_applicant when param1 is "star" and there's a param2
                            if s =="star":
                                if param2:
                                    self.existing_applicant = 1
                                else:
                                    self.existing_applicant = 0

                            #Navigate to each section and complete it
                            if s in section_mapping:
                                #Try to select chosen section
                                cs_result = self.choose_section(*section_mapping[s])
                                self.logadd("\nSECTION: "+s+"\n")
                                message = message + "Section access: "+cs_result+"\n"
                                Started_At = datetime.datetime.now()

                                #Add to screenshot info, if screenshots active
                                if self.screenshot_folder:
                                    id = str(row)+s
                                    self.scrlog.heading(section_mapping[s][0],id=id,size="2")
                                    #Add in-page link to nav data (link within page, text lable take from section_mappings)
                                    nav_data.append((id,section_mapping[s][0]))

                                #If section selection was successful, begin completing pages within the section
                                if cs_result=="ok":
                                    #Complete section using treadmill method
                                    #sft=s is just additional text for screenshot filenames
                                    result = self.treadmill(sft=s,auto_mode=auto_mode)
                                    message = message + result
                                    #Capture case refernce when new application created
                                    if s=="star":
                                        #Write caseref to "Reusable" column of spreadsheet
                                        ws1.Cells(row,datacols["Reusable"]).Value = result

                                Ended_At = datetime.datetime.now()
                                message = message + "\n"+"Elasped "+str(Ended_At-Started_At)

                            else:
                                message = message + "Section '"+s+"' not recognised."
                    else:
                        message = "Not executed because not logged in to CCMS."


                #* Submit Complete application - can be very slow *
                elif action == 'subm':
                    #default timeout for submission
                    timeout = 120
                    #Only execute if logged in
                    if self.logged_in:

                        #Replace timout with value from spreadsheet, if set
                        if param1:
                            try:
                                timeout = int(param1)
                            except Exception as e:
                                print "Problem with submission timeout value. Reverting to default.",e
                            #Don't allow negative values
                            else:
                                if timeout<0:
                                    timout=0
                        #Try to submit application
                        message = self.complete_application(submission_timeout=timeout)
                    else:
                        message = "Not executed because not logged in to CCMS."

                #Write time/date to spreadsheet
                ws1.Cells(row,datacols["Date/Time"]).Value = time.strftime("%d-%m-%Y [%H:%M:%S]")
                #Write message (if exists) to spreadsheet
                if message:
                    ws1.Cells(row,datacols["Message"]).Value = message
                    #Also add to logfile if logging enabled
                    if self.logfile:
                        self.logadd(message+"\n")

                #Could add saving spreadsheet at this point to save for each row?

        #Close things at end
        #Save spreadsheet
        self.wb.Save()

        #Final message to logfile, if being used
        if self.logfile:
            self.logadd("End: "+time.strftime("%d-%m-%Y [%H:%M:%S]")+"\n")

        #Add final tags to screenshot report, if being used
        if self.screenshot_folder:
            #Add nav
            self.scrlog.nav("Links",nav_data)
            #HTML end-tags
            self.scrlog.add("""\n</body>\n</html>\n""")

        #Close Webdriver
        ##self.driver.quit()
        print "Finished"


    def excel_read_test_data(self,tab = "Page Details",value_column = 'F',clear = True):
        """Reads PUI page data from speciied column and tab in specially formatted
        spreadsheet and stores it in self.pi dictionary
        self.pi has keys constructed from page headings, each with associated value
        that's a list. Each of these lists contains dictionaries with details
        of each element within the page to be completed
        self.pi[("Heading","Subheading")]=[{id:"ap_dob",type:"field",value:"01/12/70"},{id:"under18",type:"click",value:"this"}]

        Args:
            tab - tab (name or number) with data to be read
            value_ column - column letter with values to be read
            clear - (bool) if True, existing details cleared before reading new ones.

        Returns:
            endrow value (int)
        """
        #Clear existing values if requested
        if clear:
            self.pi={}

        #Test Data Spreadsheet Columns
        #Row with column names in spreadsheet
        headrow = 4
        #Expected column names from spreadsheet
        headnames = ["Heading","Sub Heading","Element Type","Element ID"]

        #select first worksheet in the spreadsheet
        ws_pd = self.wb.Worksheets[tab]

        #Find positions of requred columns in spreadsheet
        datacols = dict.fromkeys(headnames, -1)#Dictionary to hold mapping between heading name and column number
        for col in range(1,15):
            val = str(ws_pd.Cells(headrow,col).Value)
            if val in headnames:
                datacols[val]=col
        #Warn if any required columns not found
        not_found = [k for k in datacols if datacols[k]==-1]
        if not_found:
            print "WARNING - may fail as required heading(s) "+",".join(not_found)+" not found in spreadsheet row "+str(headrow)+"."


        message = "Reading page details from tab "+str(tab)+", column "+str(value_column)+". "
        #Start row
        startrow = headrow+1

        #Read end row from spreadsheet
        er = ws_pd.Cells(2,"B").Value
        #Use it as endrow value if it can be used as valid integer
        try:
            er = int(er)
        except (ValueError,TypeError) as e:
            #Use default end row if value from spreadsheet not valid
            endrow = 200
            print message+"Default end row used:",endrow
        else:
            endrow = er
            #endrow can't be less than startrow
            if endrow<startrow:
                endrow=startrow
                print message+"End row from spreadsheet too low. Changed to:",endrow
            else:
                print message+"End row taken from spreadsheet:",endrow

        #Read values
        for row in range(startrow,endrow+1):

            #Construct self.pi key. If component not None, ensure it's a str with no leading/trailing spaces
            k1 = ws_pd.Cells(row,datacols["Heading"]).Value
            #Proceess unicode value
            if type(k1) is unicode:
                k1 = k1.encode("utf-8").strip()
            #Process all other types except None, which isn't modified
            elif k1 is not None:
                k1 = str(k1).strip()
            k2 = ws_pd.Cells(row,datacols["Sub Heading"]).Value
            #Proceess unicode value
            if type(k2) is unicode:
                k2 = k2.encode("utf-8").strip()
            #Process all other types except None, which is ignored
            elif k2 is not None:
                k2 = str(k2).strip()
            key = (k1,k2)

            #Only do stuff is key is valid - currently can't contain None
            if None not in key:

                #Add key if its not already present
                if key not in self.pi:
                    self.pi[key]=[]

                #Read element id
                id = ws_pd.Cells(row,datacols["Element ID"]).Value

                #If ID not None, store it along with its assocated type and value
                if id is not None:
                    #Ensure id is string with no leading/trailing spaces
                    id = str(id).strip()
                    #Read element type for id
                    etype = ws_pd.Cells(row,datacols["Element Type"]).Value
                    #Read associated value from chosen value column (value column can vary)
                    value = ws_pd.Cells(row,value_column).Value
                    #Conversion of dates to string (but time type hard to catch)
                    if type(value) is pywintypes.TimeType:
                        value = str(value)

                    #Key to page dictionary, key to element type dictionary, key to element id
                    self.pi[key].append({"id":id,"type":etype,"value":value})

        ##for k,v in self.pi.iteritems():
        ##    print k,v
        return endrow


    def choose_section(self,section_name,subsection=None):
        """Navigate to start new section of application for legal aid.

        For parts that expect to start from Application Summary, checks if
        chosen link is present and available on Application Summary
        screen. If it is, click it.

        In many cases, can be followed up with call to self.treadmill()


        Args:
            section_name - name of link to be clicked, eg "Merits Assessment"
            or "New Application" to start a new application
            subsection - (optional) subsection only used by Opponent which can
            be organisation or individual

        Returns:
            information message which is "ok"  if screen successfully accessed,
            otherwise contains error message

        """
        driver = self.driver

        #Known key contents of page source for known sections. Values used by
        #WebDriverWait() to see if page has loaded
        known_idents={
        'Application Type':'Create Application',
        'Provider Details':'Name of solicitor',
        'Client Details':'Summary of Client Details',
        'General Details':'Correspondence Address',
        'Proceedings and Costs':'Please enter the proceeding details',
        'Opponents and Other Parties':'Please enter the details for any opponents',
        'Merits Assessment':'interview is designed to collect all relevant',
        'Means Assessment':'interview is designed to collect information',
        'Non-Passported':'This interview is designed to collect information',
        'Passported':'The interview is designed to collect information'
        }

        #Holds message to be returned
        message = ""

        #"Start new application" - just makes sure home-page on display (new version)
        #Diferent steps for old, new and Beat PUIs
        if section_name=="New Application":
            #New PUI
            if self.logged_in =="new":
                #Navigate to Home Screen if we're not there
                if "Civil legal aid applications, amendments and billing." not in driver.page_source:
                    containing_div = driver.find_element_by_class_name("ccmsPageHeaderLogo")
                    links = containing_div.find_elements_by_tag_name("a")
                    if links:
                        links[0].click()
                    WebDriverWait(self.driver,30).until(lambda driver:"Civil legal aid applications, amendments and billing." in driver.page_source)
            #Beta PUI
            elif self.logged_in =="beta":
                #Navigate to Home Screen if we're not there
                if "Civil legal aid application, amendments and billing." not in driver.page_source:
                    containing_div = driver.find_element_by_class_name("ccmsPageHeaderLogo")
                    links = containing_div.find_elements_by_tag_name("a")
                    if links:
                        links[0].click()
                    WebDriverWait(self.driver,30).until(lambda driver:"Civil legal aid application, amendment and billing." in driver.page_source)
            #Any other PUI
            else:
                #Navigate to Home Screen if we're not there
                if "Welcome to the Client and Cost Management System" not in driver.page_source:
                    driver.find_element_by_link_text("Home").click()
                    WebDriverWait(self.driver,30).until(lambda driver:"Welcome to the Client and Cost Management System" in driver.page_source)
            message = "ok"

        #All other links (start from application summary page)
        else:
            #Check statuses
            statuses = self.read_summary()

            print "Statuses:",statuses

            #Automatically change 'Means Assessment' to 'Non-Passported' when
            if section_name=='Means Assessment' and section_name not in statuses and 'Non-Passported' in statuses:
                section_name = 'Non-Passported'
                print "Section choice automatically changed from 'Means Assessment' to 'Non-Passported'."

            #Can only do stuff if we've actually found statuses
            if not statuses:
                message = "No statuses found. Not on Application Summary page?"
            #See if section name not on page
            elif section_name not in statuses:
                message = "Section '"+section_name+"' not found on Application Summary page."
            #If on page, check see if its status is "Not Available"
            elif statuses[section_name] == "Not Available":
                message = "Section '"+section_name+"' has status 'Not Available' on  Application Summary page."

            #If choice seems OK for page, click it
            else:

                #Click the link
                driver.find_element_by_link_text(section_name).click()

                #Wait for page, if we know what to wait for
                if section_name in known_idents:

                    #"Beta" has iframes which need to be selected in order to check the wanted value
                    #Check if owd iframe present and switch to it if found
                    owd_selected = self.owd_frame_check()

                     #Wait for expected text to be found in page source
                    wanted = known_idents[section_name]
                    WebDriverWait(self.driver,40).until(lambda driver: wanted in driver.page_source)

                    #Return to default part of page if we'd previously switched to iframe
                    if owd_selected:
                        driver.switch_to_default_content()
                        print "(iFrame) returning to main page."

                #Extra clicks for Proceedings and Opponents
                #Proceeding
                if section_name=='Proceedings and Costs':
                    driver.find_element_by_link_text('Add Proceeding').click()
                    WebDriverWait(self.driver,30).until(lambda driver:"Proceeding Details - Step 1 of 4" in driver.page_source)

                #Opponent individual
                if section_name=='Opponents and Other Parties' and subsection=="Individual":
                    driver.find_element_by_link_text("Add Individual").click()
                    WebDriverWait(self.driver,30).until(lambda driver:"Individual Contact Details" in driver.page_source)

                #Opponent organisation
                if section_name=='Opponents and Other Parties' and subsection=="Organisation":
                    driver.find_element_by_link_text("Add Organisation").click()
                    WebDriverWait(self.driver,30).until(lambda driver:"Please enter the Organisation details" in driver.page_source)

                message = "ok"

        return message


    def treadmill(self,sft="",auto_mode=1):
        """Completes sections of application.

        Generic - should work for most sections as long as first page opened.

        The first page of the section should already be on display before
        this method is called.
        For "new application" home page should be open.
        For other sections, the first page of the section should be open
        (self.choose_section() can be used to access relevant start page)

        Completes values on any page if any defined in self.pi.
        Also completes values for specific pages defined wihtin this method
        (*Special handling for some pages*)
        If page not known will simply press the "next" button.


        Will stop if (a) Reaches a "confirm" button (will press this button)
        (b) is stuck on same page (c) "System Busy" error encountered (d) The
        final "new application" page is reached (will submit the application
        if no errors have been encountered)
        (e) A page with no "Next" button is reached and no alternative
        page-advance button has been specified.

        Previous script versions had separate method "new_application_treadmill()"
        for "new application" section but this functionality now included here

        Args:
            sft - screenshot filename text. Optional extra text that's included
            in screenshot filenames. Only relevant if screenshots active.
            (self.screenshot_folder has a value)
            auto_mode (int) - run mode passed to self.fill_page_auto()

        Returns:
            success failure message (case reference for successful new application)
        """

        driver = self.driver
        #Text to be returned
        message = ""

        # *Initial Navigation* - getting to first screen in section
        #New application (starts from home-page). Need to navigate to start.
        if 'value="CCMS_AH01"' in driver.page_source:
            #Press start new application button
            containing_div = driver.find_element_by_class_name("btnStartNewApplication")
            button = containing_div.find_element_by_class_name("button")
            button.click()
            #Wait for "Select Office" page
            WebDriverWait(self.driver,10).until(lambda driver:"Select Office" in driver.page_source)
        #Other sections have no preliminary navigation. self.choose_section() should
        #have already been used to get to right place. Checking could be added here.
        else:
            pass

        # *Treadmill Start*
        #Dummy page headings values to get started
        headings = ("","")
        #Counter used in log file and screenshot capture
        i = 0
        #Used to cope with being stuck on same page
        page_retry_count = 0
        while True:
            #Read headings from page
            new_headings = self.get_headings()

            #STOP if stuck on same page
            if new_headings == headings:
                print "Same page!", new_headings
                message = "Stuck on page: "+str(headings[0])+","+str(headings[1])+"."+" Retry count:"+str(page_retry_count)
                #Extract warning messages
                message_list = driver.find_elements_by_id("messages")
                if message_list:
                    messages = [m.text for m in message_list[0].find_elements_by_tag_name("li")]
                    print "<Warnings>",messages
                    message = message + ",".join(messages)
                    #Add page source to warning message
                    message =   (message
                                + "\n\n***** Page Source Start *****"
                                + driver.page_source.encode("ascii","ignore")+"\n"
                                + "\n***** Page Source End *****\n\n"
                                )

                #Save screenshot if they're active
                if self.screenshot_folder:
                    filename = time.strftime("[%H.%M.%S]")+sft+" STUCK ("+str(i)+")_"+str(headings[0])+"_"+str(headings[1])+".png"
                    self.screenshot(filename,heading=headings[1]+" (stuck on page)")

                #Stop if we're stuck on samge page
                if page_retry_count >= self.stuck_page_retries:
                    user_response = raw_input("Seem to be stuck on same page. Type Q to quit or anything else to continue (beyond retry count).")
                    if user_response.upper() =="Q":
                        break
                #Increament page retry count
                page_retry_count +=1

            #Reset page retry count if we have new headings
            else:
                page_retry_count = 0

            #New headings are now headings for current page
            headings = new_headings
            print "\nNew page:", headings[0],headings[1]

            #Add headings to log file (if active)
            if self.logfile:
                line = "("+str(i)+") "+",".join(new_headings)+"\n"
                self.logadd(line,"a")
                details = self.analyse_page().encode("utf-8")
                self.logadd(details,"a")

            #STOP if "System Busy" error on page
            if "System Busy" in driver.page_source:
                message = message +'"System Busy" error'
                break

            #Check for presence of "owd" iframe and switch to it if found
            owd_selected = self.owd_frame_check()

            # *Complete Standard Page Items*

            #Try to auto complete mandatory fields
            self.fill_page_auto(auto_mode)

            #If headings correspond with a page that has defined data in self.pi
            #complete the data by calling self.fill_page()
            if headings in self.pi:

                #Need to control mulitple update of page to handle dynamic content
                coninue_update = True
                #Limit number of updates to prevent stuck in infinite loop
                max_tries = 10
                tries = 0
                while coninue_update:
                    #Default to updating the page only once. Later condition can re-enable
                    coninue_update = False
                    #Fill page with specified details. Capture list of items which could not be updated due to being hidden
                    hidden1 = self.fill_page(headings,update_enabled=True)

                    #If some items hidden, check again to see which are currently hidden
                    if hidden1:
                        #***Might need to add a delay or sync point here ***
                        hidden2 = self.fill_page(headings,update_enabled=False)
                        if hidden1==hidden2:
                            print "[Dynamic] Some requested items hidden but have remained hidden after updating page."
                        elif len(hidden2)<len(hidden1):
                            #Need to have another go at updating page if number of hidden items has reduced
                            print "[Dynamic] Number of hidden items has reduced after updating page. A further update will be attempted."
                            coninue_update = True
                    else:
                        print "[Dynamic] No requested item has hidden status."
                    #Count attempts and give up if maximum number reached
                    tries = tries +1
                    if tries > max_tries:
                        coninue_update = False
                        print "[Dynamic] Stopping page update as maximum attempts reached.",tries

            #If page lacks defined data, don't do anything special
            else:
                pass
                ##print "No handling for this page"

            #Default "next" button (is changed for certain specified pages below)
            #Two options depending upon whether page is "Classic" or Beta with "owd" frame
            if owd_selected:
                bclass = "owd-submit"
            else:
                bclass = "btnNext"

            # *Special handling for some pages*

            #Applicant Search - uses different button
            if headings == (u"Start New Application",u"Client Search"):
                bclass = "btnSearchClient"

            #Register new client when search failed to find existing applicant(s)
            #Uses different button
            elif headings == (u"Start New Application",u"Client Search - No Search Results"):
                bclass = "btnRegisterNewClient"

            #Register other party organisation
            elif headings == (u'Create Application', u'Opponents and Other Parties'):
                bclass = "btnCreateNewOrganisation"

            #Choose existing applicant or create a new one
            elif headings == (u"Start New Application",u"Client Search - Results"):

                #Find the "Use this record" links in search results
                table = driver.find_element_by_class_name("searchresults")
                links = table.find_elements_by_tag_name("a")

                #Stop if no search results displayed (there always should be some
                #for this page title. If none, something's gone wrong)
                if not links:
                    message = message +"Problem with applicant search. Heading indicates successful search but no results found."
                    #STOP
                    break

                #Use existing applicant if wanted
                if self.existing_applicant:
                    #Click topmost "Use this record" link
                    links[0].click()
                    #Wait for confirmation page
                    WebDriverWait(self.driver,10).until(lambda driver:"Client Confirmation" in driver.page_source)
                    #Press "confirm" button
                    containing_div = driver.find_element_by_class_name("btnConfirm")
                    containing_div.find_element_by_class_name("button").click()
                    print "Case creation selected existing applicant (topmost search result)."
                    #STOP - an end-point reached
                    break

                #Existing applicant not wanted - create new one
                else:
                    bclass = "btnRegisterNewClient"
                    #If multiple pages of results, move to last page as it's the one with the confirm button.
                    if "paginationPanel" in driver.page_source:
                        #Find all the page navigation links
                        page_link_area = driver.find_element_by_class_name("paginationPanel")
                        plinks = [e.text for e in page_link_area.find_elements_by_tag_name("a")]
                        #Press the penultimate link (Last is "Next" link, so we want the one before it)
                        driver.find_element_by_link_text(plinks[-2]).click()
                    print "Case creation created new client record."

            #Client Address (new application) - mandatory postcode search when applicant has UK address
            elif headings == (u'Client Registration', u'Client Address Details'):
                #Find the value from the country list and only do search if set to UK
                country_list = driver.find_element_by_id("client_mainAddress_country")
                list_details = self.list_values(country_list,10)
                if list_details[0]==u"United Kingdom":
                    #Carry out postcode search
                    pcode_result = self.postcode_search()
                    print pcode_result
                    #Add postcode result to logfile
                    if self.logfile:
                        self.logadd(pcode_result+"\n")

            # Got to end of applicant creation (page after "Summary of Client Information")
            # Needs special handling as different WebDriverWait needed
            # due to temporary auto-closing "Submission in progress" page.
            elif headings == (u"Client Registration", u"Summary of Client Information"):
                #If message still empty, there were no errors so we should be at the end and so
                #can wait for Confirmation Page

                #If no error messages, submit the applicant
                if message == "":
                    #Press button to proceed to next page
                    containing_div = driver.find_element_by_class_name("btnNext")
                    containing_div.find_element_by_class_name("button").click()
                    # System can get stuck here with "Submission Pending" error page
                    WebDriverWait(self.driver,240).until(lambda driver:"Submission Confirmation" in driver.page_source or "Submission Pending" in driver.page_source)
                    #If submission stuck due to technical issue, quit
                    if "Submission Pending" in driver.page_source:
                        #STOP technical fault
                        message = "Applicant submission failed. 'Submission Pending' encountered."
                        break
                    #Next Button
                    containing_div = driver.find_element_by_class_name(bclass)
                    containing_div.find_element_by_class_name("button").click()
                    #Wait for Application Summary
                    WebDriverWait(self.driver,240).until(lambda driver:"Application Summary" in driver.page_source)
                    #Read case reference
                    case_ref_div = driver.find_element_by_id("lscReference")
                    span_tags_text = [e.text for e in case_ref_div.find_elements_by_tag_name("span")]
                    case_ref=span_tags_text[-1]
                    message = case_ref
                #STOP We're at end of applicant creation now, so can leave the while loop
                #regardless of whether we've been able to submit the application
                break

            #Save screnshot (of filled in page)
            if self.screenshot_folder:
                #Some headings have characters forbidden in Windows filenames
                #filename = time.strftime("[%H.%M.%S]")+" ("+str(i)+") "+str(headings[0])+" "+str(headings[1])+".png"
                filename = time.strftime("[%H.%M.%S]")+sft+"("+str(i)+").png"
                self.screenshot(filename,heading=headings[1])


            #STOP if we've got to page with "Confirm" button (Except during Opponent Creation)
            #But only if we're on a non-owb page
            if not owd_selected and 'class="buttonArea btnConfirm">' in driver.page_source:
                #For organisation creation we just click the button to proceed to next page
                if headings == (u"Create Application",u"Opponents and Other Parties - Organisation"):
                    bclass = "btnConfirm"
                #Click confirm button and end in other cases as "Confirm" is last page
                else:
                    #Click Confirm button
                    containing_div = driver.find_element_by_class_name("btnConfirm")
                    containing_div.find_element_by_class_name("button").click()
                    message = 'Ended with "Confirm" button.'
                    #Wait for Application Summary page
                    WebDriverWait(self.driver,260).until(lambda driver:"Application Summary" in driver.page_source)
                    break
                    #Should automatically return to Application Summary when confirm clicked

            # *Advance to next page*
            print "Page advance button:",bclass
            #Capture unique ids from current page - used to determine if page updated
            u = self.unique_ids()
            #Press button to proceed to next page
            try:
                #Direct button pressing for "owd" pages
                if owd_selected:
                    #This button click doesn't work if button off bottom of page. Line below ensures it's scrolled into view
                    driver.execute_script("return arguments[0].scrollIntoView();",driver.find_element_by_class_name(bclass))
                    #Click button
                    #Keith found problem where button click would fail on iframe page.
                    #Seems driver was finding two elements with class name "owd-submit"
                    #Now look for multiple matches and click whichever is last
                    matches = driver.find_elements_by_class_name(bclass)
                    ##matches[-1].click()
                    #Might be more reliable than using click
                    matches[-1].send_keys(Keys.RETURN)
                    ##driver.find_element_by_class_name(bclass).click()

                #Button pressing via div for non "owd" pages
                else:
                    containing_div = driver.find_element_by_class_name(bclass)
                    button = containing_div.find_element_by_class_name("button")
                    ##button.click()
                    button.send_keys(Keys.RETURN)

            #STOP if button not present. This is standard way to end some sections
            #Also return to "Create Application" screen
            except NoSuchElementException as e:
                message = message +  "Stopped because page without  no 'next' or 'confirm' button reached. "
                #Return to Create Application
                driver.find_element_by_link_text("Return to Create Application").click()
                #Wait for the page
                WebDriverWait(self.driver,90).until(lambda driver:"Application Summary" in driver.page_source)
                break

            #Wait for page to update after button press
            #Special wait for applicant search during case creation
            #because it (a) can fail, (b) can be slow
            #STOP if search gets stuck
            if headings == (u"Start New Application",u"Client Search"):
                try:
                    WebDriverWait(self.driver,90).until(lambda driver:"Your search has returned" in driver.page_source)
                except TimeoutException:
                    message = "Client search seems slow or broken. Results not displayed after 90 seconds."
                    break
            #General wait for all pages
            #Wait for unique ids to change - can mean (a) new page or (b) old page refreshed with error message
            WebDriverWait(self.driver,60).until(lambda u: u != self.unique_ids())
            i = i+1
            #Should now be on new page or refreshed existing page. Loop back to "While True"

            #If owd frame had been selected, return focus back main part of page
            if owd_selected:
                driver.switch_to_default_content()
                print "(iFrame) returning to main page."

        #Treadmill end
        return message


    def fill_page(self,headings,update_enabled=True):
        """Completes details on typical PUI page using values from self.pi
        dictionary

        self.pi needs to have following structure
        (1) Outer dictionary with key for each page using tuple of heading and sub-heading
        (2) Value associated with key is a list
        (3) List contains dictionary for each element to be completed
        (4) element dictionary has keys "id","type" and "value"


        For example
        self.pi[("Means","Income")] = [{"id":"salary","type":"field",value:"100"},{"id":"self_emp_false","type":"click",value:"this"}]

        Args:
            headings - tuple key to self.pi outer dictionary - used to identify the page
            e.g. ("Means Assessment","Address Entry")
            update_enabled - (bool) when true page will be updated, when false
            elements checked to see if present and not-hidden but not updated.

        Returns:
            list of items which could not be updated because they were hidden
            at time update was attempted
            """
        driver = self.driver

        #stores elements which could not be updated because they have status
        #aria-hidden=true
        hidden_items=[]

        #Select details for specified page
        page_details = self.pi[headings]
        if update_enabled:
            print "<Completing>",headings
        else:
            print "<Status Checking>",headings

        #Loop through each element
        for item in self.pi[headings]:
            #Only do something if element has type and id
            if item["type"] and item["id"]:
                print "<",headings,">|<"+item["type"]+">|",item["id"],"|",item["value"]

                #Construct css selector value for each type (except "link" which doesn't use css selector)
                if item["type"] in ["field","click"]:
                    css_sel ="input[id*='"+item["id"]+"']"

                elif item["type"]=="area":
                    css_sel ="textarea[id*='"+item["id"]+"']"

                elif item["type"]=="ddlist":
                    css_sel = "select[id*='"+item["id"]+"']"

                elif item["type"]=="radio":
                    v = str(item["value"]).lower()
                    if v[:1] in ["t","y"]:
                        css_sel = self.radio_suffix_set(item["id"],["_true","t"])
                    else:
                        css_sel = self.radio_suffix_set(item["id"],["_false","f"])
                else:
                    css_sel = ""

                #Boolian that becomes false if find not valid to update element
                ok_to_update = True

                #Try to select element
                try:
                    #Hyperlink
                    if item["type"] =="link":
                        ele = driver.find_element_by_link_text(item["id"])
                    #Everything else
                    else:
                        ele = driver.find_element_by_css_selector(css_sel)
                    #Try to scroll element into view. Can get problems if element in iframe and below visible page area.
                    driver.execute_script("return arguments[0].scrollIntoView();",ele)

                #Not found on page
                except NoSuchElementException as e:
                    ok_to_update = False
                    print "*FAILED* element not found on page"
                except ElementNotVisibleException as e:
                    print "*FAILED* element present but not visible"

                #Can't update element if it's hidden
                if ok_to_update:
                    if ele.get_attribute("aria-hidden")=="true":
                        ok_to_update = False
                        print"*Element Hidden*"
                        hidden_items.append(item)

                #Element both present and not hidden

                if update_enabled and ok_to_update:
                    #Complete the item - text field and text area

                    if item["type"] in ["field","area"]:
                        v=item["value"]
                        if v is None:
                            v=""
                        ele.clear()
                        ele.send_keys(Keys.CONTROL + "a")
                        ele.send_keys(Keys.BACK_SPACE)
                        ele.send_keys(v)
                        print "typed",v

                    #Buttons, tick-boxes
                    elif item["type"] in  ["click"]:
                        #Special handling for tick-boxes as they might already be in desired state
                        if ele.get_attribute("type")=="checkbox":
                            #find state of tick-box
                            state = ele.is_selected()

                            #To aid comparison, convert supplied value to lower-case string with white-space removed
                            v=item["value"]
                            if type(v) is unicode:
                                v = v.encode("utf-8").strip().lower()
                            else:
                                v = str(v).strip().lower()

                            #Only click tick-box if current state doesn't match requested state
                            #(assuming v value of "n" means "unticked", any other value means "ticked")
                            if (v[:1] == "n" and state == True) or (v[:1]!="n" and state == False):
                                ele.click()
                                newstate = ele.is_selected()
                                print "tick-box:",item["id"],"changed from",state,"to",newstate
                            else:
                                print "tick-box:",item["id"],"not updated as it's already in requested state:",state

                        #Non tick-box, click as long as there's a value
                        else:
                            if item["value"]:
                                ele.click()
                                print "Clicked non tick-box:",item["id"]

                    #Click radio button if a value set
                    elif item["type"]=="radio" and item["value"]:
                        ele.click()
                        print "Clicked radio Y/N"

                    #Select ddlist
                    elif item["type"]=="ddlist":
                        v = item["value"]
                        #Select by position
                        if type(v) in (int,float):
                            v=int(v)
                            Select(ele).select_by_index(v)
                            print "Selected by position:",v
                        #Select by name
                        elif v is not None:
                            found = self.list_select_by_name(ele,v)
                            print "selected by text:",v
                            if not found:
                                print "*FAILED* to find by text:",v

                    #Hyperlink - only click if value set
                    elif item["type"] =="link" and item["value"]:
                            ele.click()
                            print "Clicked hyperlink:",item["id"]


        return hidden_items


    def radio_suffix_set(self,id,endings):
        """Work out the right true/false suffix for radio button. Generates
        css_sel string used by self.fill_page

        Args:
            id - partial element id used to construct CSS selector
            endings - list containing endings to try
        Returns:
            constructed cssl_selector for radio button
        """
        driver=self.driver
        for end in endings:
            css_sel = "input[id*='"+id+end+"']"

            try:
                ele = driver.find_element_by_css_selector(css_sel)
            #Not found on page
            except NoSuchElementException as e:
                pass
            else:
                #Found,so stop
                print "Using radio button:", css_sel
                break

        return css_sel


    def fill_page_auto(self, mode=1):
        """Attempts to identify and automatically complete elements
        wihtin a PUI page.

        Args:
            mode (int)  - 0 do nothing
                        - 1 examine page but type nothing
                        - 2 examine page and complete mandatory fields
                        - 3 examine page and complete all fields
                        - any other number, analyse page but type nothing
        """
        #Do nothing when mode is 0
        if mode ==0:
            return

        driver=self.driver

        #Display/log info message
        message = "Automatic element identification and completion. Mode:"+str(mode)
        print message
        if self.logfile:
            self.logadd(message+"\n")

        #Find input-field containing divs by looking for class names "inputField","inputFieldDate","confirmCheckBox"
        #Doing one at a time, so we can set text to hold a date for "inputFieldDate"
        for class_name in ["inputField","inputFieldDate","confirmCheckBox"]:
            input_class_divs = driver.find_elements_by_class_name(class_name)

            #Loop through each div found
            # i just used in dummy data generation
            for i,div in enumerate(input_class_divs):

                #Set standard text to be typed (only two options for now)
                if class_name=="inputFieldDate":
                    typin = "01/01/1990"
                else:
                    #gernerate text, partly based on iteration number,i
                    typin = "ABC"+"".join(["abcdefghij"[int(c)] for c in str(i)])


                #See if div contains mandatory elements by checking if
                #it contains an image with the  alt text "Required"
                mand = [e for e in div.find_elements_by_tag_name("img") if e.get_attribute("alt")=="Required"]

                #Identify elements of interest within the div
                fields = div.find_elements_by_tag_name("input")# buttons and text fields
                ddlists = div.find_elements_by_tag_name("select")# drop-down lists
                areas = div.find_elements_by_tag_name("textarea")# Text areas (multi-line fileds)

                message = ""
                #Examine each element and display/log info.
                #Complete it if: (a) mode is 4 or (b) we think it's mandatory and mode is 3
                #i used in data generation
                for e in fields+ddlists+areas:

                    ##print e.get_attribute("class")
                    id = e.get_attribute("id")
                    etype = e.get_attribute("type")

                    #Text field and area
                    if etype in ["text","textarea"]:

                        if etype=="text":
                            message =  "field\t"+id
                        else:
                            message =  "area\t"+id
                        #Complete fields if mode right
                        if mode == 3  or (mode == 2 and mand):
                            #Type the data in
                            e.clear()
                            e.send_keys(typin)
                            message = message +"\t"+typin

                    #Radio button and tick boxes
                    elif etype in  ["checkbox","radio"]:
                        message = "click\t"+id
                        if mode == 3  or (mode == 2 and mand):
                            e.click()
                            message = message +"\tthis"

                    #Drop-down list
                    elif etype == "select-one":
                        message = "ddlist\t"+id
                        if mode == 3  or (mode == 2 and mand):
                            Select(e).select_by_index(1)
                            message = message+"\t1"

                    #Display and log details
                    print message
                    if self.logfile:
                        self.logadd(message+"\n")

        print "Automatic element identify and completion end."
        if self.logfile:
            self.logadd("Auto identify and completion end."+"\n")


    def complete_application(self,submission_timeout=120):
        """Presses the "Complete Application" button and completed the application.
        Expects to start from Application summary page

        Args:
            submission_timeout - submission timeout time in seconds

        Returns:
            success/failure message
        """
        driver = self.driver

        #See if we're on the right page and return if we're not
        if "Application Summary" not in driver.page_source:
            message = "Failed - Application Summary page not open."
            return message

        #Press "complete" application button
        containing_div = driver.find_element_by_class_name("btnSubmitApplication")
        button = containing_div.find_element_by_class_name("button")
        button.click()

        #A different "Application Summary" page
        WebDriverWait(self.driver,60).until(lambda driver:"The information you have entered in your Application is listed below" in driver.page_source)
        #Next Button
        containing_div = driver.find_element_by_class_name("btnNext")
        containing_div.find_element_by_class_name("button").click()

        #Submit application
        WebDriverWait(self.driver,60).until(lambda driver:"Please complete the declaration" in driver.page_source)
        #tickbox
        driver.find_element_by_id("declaration_checkboxes_0__optionalValue").click()
        #Continue button
        containing_div = driver.find_element_by_class_name("btnConfirm")
        containing_div.find_element_by_class_name("button").click()

        #Confirmation Page
        #This bit can be very slow to update
        try:
            WebDriverWait(self.driver,submission_timeout).until(lambda driver:"Submission Confirmation" in driver.page_source)
        except TimeoutException:
            message = "Submission did not complete within timeout time ("+str(submission_timeout)+" seconds)."
            #Need to add clicking a link here
        else:
            message = "Submission completed"

        #Next Button - might need to tweak this for timeout situation
        containing_div = driver.find_element_by_class_name("btnNext")
        containing_div.find_element_by_class_name("button").click()

        #Now on top page
        WebDriverWait(self.driver,40).until(lambda driver:"Welcome to the Client and Cost Management System" in driver.page_source)

        return message

    def login_ccms(self,url,username,password,link="Client and Cost Management System (CCMS)",direct=False):
        """Login to Portal
        Creates webdriver instance (self.driver) and logs into Portal but
        not if already logged in.

        Updates self.logged_in if login was successful.

        Args:
            url - portal url
            username
            password
            direct - (bool) when true assume Portal not being used. Instead login
            using directly URL without username and password required.
        Return:
            message
        """

        if not self.logged_in:
            message = "Not logged in."
            self.driver = webdriver.Firefox(self.firefox_profile)#Start webdriver
            driver = self.driver
            print "Accessing:",url
            print "Username:",username

            #Open URL
            self.driver.get(url)

            if direct:
                #Dodgy direct login, bypassing portal
                #Wait for page
                WebDriverWait(driver,90).until(lambda x: driver.title[:5]=="LAA C")
                #Find CCMS version
                self.ccms_identify()
                message = "Dodgy direct login "+url+". Found PUI version:"+self.logged_in

            else:
                #Portal login (new or old)
                #Wait for page
                WebDriverWait(driver,10).until(lambda driver: driver.title=="LAA Online Portal"
                                            or "By logging in to this Portal" in driver.page_source)
                #New portal login
                if driver.title == "LAA Online Portal":
                    #Username and password, if values supplied
                    if username is not None: driver.find_element_by_name("username").send_keys(username)
                    if password is not None: driver.find_element_by_name("password").send_keys(password)
                    driver.find_element_by_class_name("button-start").click()
                #Old portal login
                else:
                    if username is not None:self.driver.find_element_by_name("ssousername").send_keys(username)
                    if password is not None:self.driver.find_element_by_name("password").send_keys(password)
                    self.driver.find_element_by_name("submit").click()

                #Handling different login outcomes
                #Successful login to portal
                if ">Logged in as:" in self.driver.page_source:
                    #Click CCMS link (or alterntive link)
                    driver.find_element_by_link_text(link).click()

                    #Find out which CCMS version
                    self.ccms_identify()
                    message = "Logged in to Portal "+url+" and clicked link: "+str(link)+". Found PUI version:"+self.logged_in

                #CCMS found without having clicked link in portal
                elif "Civil legal aid applications, amendments and billing" in self.driver.page_source:
                    self.ccms_identify()
                    message = "Logged in to CCMS without clicking link in Portal. Found PUI version"+self.logged_in

                #Authentification Fails
                elif '<span class="errorText">Authentication failed. Please try again.</span>' in self.driver.page_source:
                    message="Portal - access denied"

                #Username/password missing
                elif '<h2>To sign in to the Online Portal please enter your User Name and Password.</h2>' in self.driver.page_source:
                    message="Portal - login page still present. Username and/or password blank?"

                #This bit should not happen
                else:
                    message="Portal - unexpected thing happened."
        else:
            message = "Already logged in."

        return message


    def ccms_identify(self):
        """Identifies CCMS version based upon characteristic values on Home Page
        Records value in self.logged_in (as one of following "new","old","beta"
        or "unknown")
        Need to have Home Page open for this to work.
        """
        driver = self.driver

        #Wait for CCMS Home page to be displayed "LAA C" should work for old and new PUI
        identifiers = {
         "Civil legal aid applications, amendments and billing":"new"
        ,"Welcome to the Client and Cost Management System":"old"
        ,"BETA":"beta"}

        time.sleep(2)

        ##WebDriverWait(driver,90).until(lambda x: driver.title[:5]=="LAA C")
        ##WebDriverWait(self.driver,90).until(lambda driver:newwords in driver.page_source or oldwords in driver.page_source)
        WebDriverWait(self.driver,90).until(lambda driver: [words in driver.page_source for words in identifiers])

        #Identify version by comparing known text with content of page
        self.logged_in = "unknown"
        for ident in identifiers:
            if ident in driver.page_source:
                self.logged_in = identifiers[ident]
                print "Version:",self.logged_in

    def read_summary(self):
        """Reads Aplication Summary details
        This screen must be open.

        Returns:
            dictionary with "Application sections" as keys and status as values
        """
        driver = self.driver

        if "Application Summary" not in driver.page_source:
            #print "Not on application summary page?"
            return

        #with open("source.txt","w") as fout:
        #    fout.write(driver.page_source.encode("utf-8"))

        #Identify the table
        table = driver.find_element_by_class_name("applicationsections")
        #Read the headings
        headings = [th.text for th in  table.find_elements_by_tag_name("th")]
        ##print "Headings: "+",".join(headings)
        #Find the rows
        rows = table.find_elements_by_tag_name("tr")
        #Extract details for each row

        statuses={}
        for row in rows:
            #Capture all the values from the HTML table row (e.encode("utf-8") to convert from unicode)
            vals = [td.text.encode("utf-8") for td in row.find_elements_by_tag_name("td")]

            #New Dec15 - read links to cope with multipe links in first columns
            links = [e.text for e in row.find_elements_by_tag_name("a")]

            #Extract the Application sections and status details
            if len(vals)>=2:
                #New way to cope with mutlipe links in first columns
                for link in links:
                    statuses[link]=vals[1]


        #xpath below is other way of reading details but doesn't work (unicode?)
        """
        for tr in driver.find_elements_by_xpath('//table[@class="applicationsections"]//tr'):
            print "*"
            tds = tr.find_elements_by_tag_name('td')
            if tds:
                row_data = [td.text for td in tds]
                print row_data
        """
        return statuses


    def search(self,case_ref):
        """Search for case reference
        If search successful, selects the case so Application Summary displayed

        Starts from CCMS home page.

        Args:
            case_ref - case reference to search for
        Returns:
            "found" if found
            "not found" if search failed
            "busy" if  "system busy"

        """
        driver = self.driver

        #Click search link and wait for page
        #Link depends upon PUI version
        if self.logged_in in ['new','beta']:
            #New PUI version
            driver.find_element_by_link_text("Cases and Applications").click()
        else:
            driver.find_element_by_link_text("Your Cases and Applications").click()

        WebDriverWait(self.driver,40).until(lambda driver:"Case and Application Search" in driver.page_source)
        #Case reference
        driver.find_element_by_id("searchCase_lscCaseReference").clear()
        driver.find_element_by_id("searchCase_lscCaseReference").send_keys(case_ref)
        #Search button
        containing_div = driver.find_element_by_class_name("buttonArea")
        containing_div.find_element_by_class_name("button").click()
        #Wait for search result
        WebDriverWait(self.driver,10).until(lambda driver:"Your search has returned" in driver.page_source or "System Busy" in driver.page_source)

        #Below not used. Could be used to extract details from HTML table but not currently needed
        """
        for i,tr in enumerate(driver.find_elements_by_xpath('//table[@class="searchresults"]//tr')):
            tds = tr.find_elements_by_tag_name('td')
            if tds:
                row_data = [td.text for td in tds]
        """

        message = "Search failed - record not found"

        if "System Busy" in driver.page_source:
            message = "Search failed - 'system busy' error"

        #Determine if search was successful by seeing if there's
        #now a hyperlink on the page with the Case Ref as the link
        #text
        links = self.driver.find_elements_by_tag_name("a")
        linknames = [l.text for l in links]
        if case_ref in linknames:
            message = "Search successfull"
            #Click the link
            driver.find_element_by_link_text(case_ref).click()
            WebDriverWait(self.driver,90).until(lambda driver:"Application Summary" in driver.page_source)

        return message

    def postcode_search(self):
        """Uses "Find Address" button to perform postcode search.
        Relies on right page being open with required address details already entered.
        If search successful the topmost result is automatically selected.

        Returns:
            success/fail message

        """
        driver = self.driver

        #Get current page headings (needed later in check that we've returned to this page)
        headings = self.get_headings()

        #Capture unique ids (will be used to determine if page has updated)
        u = self.unique_ids()

        #Click the "Find Address" button
        try:
            containing_div = driver.find_element_by_class_name("btnFindAddress")
            containing_div.find_element_by_class_name("button").click()
        #Give up if button not present
        except NoSuchElementException as e:
            return "Postcode search: 'Find Address' button not found or could not be pressed."

        #Wait for page to update in response to button press
        WebDriverWait(self.driver,60).until(lambda u: u != self.unique_ids())

        #Handle Successful and Unsuccssful search result
        #Search unsuccessful
        if "This page contains one or more errors or warnings" in driver.page_source:
            #Extract any warning messages
            message_list = driver.find_elements_by_id("messages")
            if message_list:
                messages = [m.text for m in message_list[0].find_elements_by_tag_name("li")]
                print "<Postcode Search Warnings>",messages
                message = "Postcode search failed. Errors/warnings:" + ",".join(messages)
        #Search successful
        elif "Address Search Results" in driver.page_source:
            #Click the topmost result
            driver.find_element_by_id("selectAddress_1").click()
            #Click "Confirm" button
            containing_div = driver.find_element_by_class_name("btnConfirm")
            containing_div.find_element_by_class_name("button").click()
            #Wait for address page to return by having original headings returned
            #Home-made delay used rather than WebDriverWait
            for i in range(0,60):
                h = self.get_headings()
                ##print "H:",headings,h, h==headings
                if h==headings:
                    break
                else:
                    time.sleep(1)

            #WebDriverWait(self.driver,50).until(lambda headings: self.get_headings()==headings)

            message = "Postcode search successful. Top result selected."
        #Something else happened. Should not get the below.
        else:
            message = "Postcode Search: unexpected response."

        return message


    def unique_ids(self):
        """Return ezgov unique ID values from page

        Many (but not all) PUI pages contain tags such as:
        <input type="hidden" name="ezgov_private_hiddenData_uniqueId" value="33HZ3KE2PZ5HZ6R4UW9BL48FA276EZ3Q" />
        There can be more than element with name="ezgov_private_hiddenData_uniqueId" per page

        This method extracts all the values from elements with name="ezgov_private_hiddenData_uniqueId"
        and returns them in a list. This could be a way of uniquely identifying the page. Note the values
        change when next button presssed but valdiation prevents the page from advancing. This is
        usefull as can be used to check wether page has updated as result of "next" button being pressed
        regardless of wether the page actually advances.

        Returns:
            list containing values of elements with name "ezgov_private_hiddenData_uniqueId"
        """
        ##ids = [e.get_attribute("value") for e in self.driver.find_elements_by_name("ezgov_private_hiddenData_uniqueId")]

        ids=[]
        elements=[]
        for retries in range(10):
            try:
                elements = self.driver.find_elements_by_name("ezgov_private_hiddenData_uniqueId")
            except Exception as e:
                time.sleep(1)
            else:
                break

        if not elements:
            print "unique_ids problem!"

        for e in elements:
            ids.append(e.get_attribute("value"))
        return ids


    def get_headings(self):
        """Read main and sub headings from PUI page and return them.
        Returns empty string if element not present.

        Some BETA PUI pages have sub heading as h1 inside owdframe iframe
        rather than h2 in main page. Need to handle this.

        Returns:
            main heading, sub heading
        """
        driver = self.driver

        #Read main heading
        try:
            containing_div = driver.find_element_by_class_name("pageTitleRef")
        except NoSuchElementException:
            main_heading = ""
        else:
            #Selenium only interacts with visible elements. As "h1" heading present but hidden
            # in new GUI, can't access its content using ".text". Need to use
            # "get_attribute("textContent")" instead. However this seems to include
            #spaces and carriage returns, so need to strip them out.
            #Old way - doesn't work when title hidden
            ##main_heading = containing_div.find_element_by_tag_name("h1").text
            #New way

            try:
                main_heading = containing_div.find_element_by_tag_name("h1").get_attribute("textContent")
            except NoSuchElementException:
                main_heading = containing_div.find_element_by_tag_name("h2").get_attribute("textContent")

            main_heading = main_heading.strip()
            main_heading = main_heading.replace("\n","")
            ##print "Main Heading:",main_heading

        #Read sub-heading
        #Location depends upon whether page has iframes

        #Check for "owd" iframe and switch to it if present.
        owd_selected = self.owd_frame_check()

        if owd_selected:
            #New iframe page
            try:
                sub_heading = driver.find_element_by_class_name("screen-title").get_attribute("textContent")
            except NoSuchElementException:
                sub_heading = ""

            #Remove anoying characters
            if sub_heading:
                sub_heading = sub_heading.replace("\t","").replace("\n","").strip()

            #Return focus to main part of screen
            driver.switch_to_default_content()
            print "(iFrame) returning to main page."
        else:
            #Old non iframe heading
            try:
                containing_div = driver.find_element_by_class_name("panelHeaderWithLink")
            except NoSuchElementException:
                sub_heading = ""
            else:
                #Old Way - doesn't work when heading hidden
                ##sub_heading = containing_div.find_element_by_tag_name("h2").text
                #New Way - copes when heading present but hidden
                sub_heading = containing_div.find_element_by_tag_name("h2").get_attribute("textContent")
                sub_heading = sub_heading.strip()
                sub_heading = sub_heading.replace("\n","")

            #Some Beta PUI pages have sub-heading inside the

        ##print "Main Heading:", main_heading,"Sub Heading", sub_heading
        return (main_heading, sub_heading)


    def logout_ccms(self):
        """Logout from CCMS by clicking "Logout" link.
        Updates self.logged_in attribute if logout link clicked
        Also closes Webdriver instance (self.driver)

        Returns:
            Message reporting whether link  clicked
        """

        #Logout in E11 currently goes wrong tries to connect to
        # http://www.justice.gov.uk/legal-aid

        logout_links = self.driver.find_elements_by_link_text("Logout")
        if logout_links:
            logout_links[0].click()
            message = "Logout clicked."
            self.logged_in = ""
            #Close webdriver
            self.driver.quit()
        else:
            message = "Logout link not found on page."
        return message


    def list_values(self,field,max =  None):
        """Return current value of drop-down list and list of all values
        Args:
            element - Webdriver element corresponding with of field to be examined
            max (int or None) - optional maximum item number to examine
                        used because reading lists can be slow for long ones
                        . Might only want to examine first few options to save time.

        Returns: current - current drop down list value
                names - text of all items from list
                values - values of all items from list"""

        options = field.find_elements_by_tag_name("option")
        names = []
        values = []
        current = ""
        for o in options[:max]:#slicing with None seems to be OK
            ##print o.is_selected(),o.text
            name = o.text
            names.append(name)
            value = o.get_attribute("value")
            values.append(value)
            if o.is_selected()==True:
                current=name

        return (current,names,values)


    def list_select_by_name(self,dd_list,wanted):
        """Select named item in dropdown list on web page

        Args:
            dd_list - drop down list (Webdriver element)
            wanted - (str) name of option to be selected
                     (int) position of option to be selected

        Returns - True if option slected, otherwise false
        """
        found = False

        for oi, option in enumerate(dd_list.find_elements_by_tag_name("option")):
            ##print option,option.text
            if type(wanted) in [str,unicode]:
                if option.text==wanted:
                    found = True
                    break
            elif type(wanted) is int:
                if oi==wanted:
                    found = True
                    break
        if found:
            option.click()

        return found

    def logadd(self,text,mode="a"):
        """Write text to log file"""
        if self.logfile:
            with open(self.logfile,mode) as f:
                f.write(text)

    def owd_frame_check(self):
        """Looks for presence iframe with id 'owdFrame'. If found, switches
        focus to this frame. Needed by some aspects of Beta because
        elements of some pages moved to within the this frame and can only be accessed
        if frame selected.
        Returns:
            True if gone_to_frame, otherwise False
        """
        driver = self.driver
        gone_to_frame = False
        #Find iframes on page
        iframes = {e.get_attribute("id"):e for e  in driver.find_elements_by_tag_name("iframe")}
        #Iframe with id="owdFrame" is one that contains details we want to examine
        #select it if it is present
        if "owdFrame" in iframes:
            print "(iframe)'owdFrame' found. Switching to it."
            driver.switch_to_frame(iframes["owdFrame"])
            gone_to_frame = True

        return gone_to_frame


    def pi_show(self,sep="\t",filename="test.txt"):
        """Display test data held in self.pi in alphabetical order by Page title
        Can optionally export details to file.
        Separator can be specified to create SV file

        Args:
            sep - separator between values (default is tab)
            filename - (optional) name of file to export to.

        Returns:
            information message
        """

        #Start file if set
        if filename:
            with open(filename,"w") as f_out:
                f_out.write("CCMS Test Data\n")

        #Below magically sorts all the keys alphabetically (I thought I'd have to write something to do this!)
        sorted_keys = sorted(self.pi)

        #Standard field types
        types = ["area","click","ddlist","field","radio"]

        #Loop over sorted keys
        for key in sorted_keys:

            #Add top key and space to file for readability
            if filename:
                with open(filename,"a") as f_out:
                    f_out.write("\n"+key[0]+sep+key[1]+"\n")

            #Extract each element
            for element in self.pi[key]:
                #Value
                value = element["value"]
                #Convert to str (show_val) for writing to file but stop None becoming "None"
                if value is not None:
                    show_val = str(value)
                else:
                    show_val = ""
                #type and id
                t = element["type"]
                item = element["id"]
                #Construct line to be displayed
                line = key[0]+sep+key[1]+sep+t+sep+item+sep+show_val
                print line
                #Add to file if required
                if filename:
                    with open(filename,"a") as f_out:
                       f_out.write(line+"\n")

        #Construct return message
        if filename:
            message = "Details saved to file: "+filename
        else:
            message = "Details only printed to Python console."
        return message


    def screenshot_setup(self,screenshot_folder,heading=""):
        """Setups up screenshot folder and associated HTML report. Supplied
        folder name will be created within the current working directory.

        Args:
            screenshot_folder - name to be given to the screenshot folder.
            heading - heading text for file
        """

        self.screenshot_folder=os.path.join(os.getcwd(),screenshot_folder)
        #Create the folder if it doesn't exist
        if not os.path.exists(self.screenshot_folder):
            os.makedirs(self.screenshot_folder)

       #CSS Style section for HTML report
        style="""
        <style>
        h1 {font-size:200%;background-color:#111131;margin-bottom:3;margin-top:0;color:#EEEE22;clear:both;}
        h2 {font-size:140%;background-color:#111151;margin-bottom:8;margin-top:8;color:#FEFEFE;clear:both;}
        h3 {font-size:120%;background-color:#CDCDDD;margin-bottom:8;margin-top:8;color:#111151;clear:both;}

        table {border-collapse:collapse;cellspacing:1; cellpadding:1;}
        table, th, td {border:1px solid #555555;}
        th {font-size:100%;font-family:verdana,arial,'sans serif';background-color:#C0C0D0}
        tr {font-size:100%;font-family:verdana,arial,'sans serif';background-color:#FFFFFF;vertical-align:top;}

        #bad {background-color:#FFF366;}
        #skip {background-color:#DDFFDD;}

        body{
        font-size:75%;
        font-family:verdana,arial,'sans serif';
        background-color:#EFEFFF;
        color:#000040;
        margin:10px
        }

        nav{
        background-color:#EEEEAA;
        position: fixed;
        right: 2%;
        top: 5%;
        padding:8px;
        border:1px black solid;
        opacity: 0.9;
        }

        img{
        width: 60%;
        height: auto;
        }

        .left_block{
        float:left;
        width: 48%;
        overflow:hidden;
        }

        .right_block{
		float:left;
		overflow:hidden;
		width: 48%;
		##padding-right:30px
        }

        </style>
        """

        #Start HMTL file that accompanies screenshots
        html_log_file=os.path.join(self.screenshot_folder,"report.html")
        self.scrlog = Logger(html_log_file)
        self.scrlog.add("<!DOCTYPE html>\n<html>\n")
        self.scrlog.add("<head>\n")
        self.scrlog.add(style)
        self.scrlog.add("</head>\n<body>\n")
        self.scrlog.heading(heading,size="1")
        #Log the date/time
        self.scrlog.tagger(time.strftime("Start: %H:%M:%S on %d/%m/%Y"),["<b>"])


    def screenshot(self,filename="screenshot.png",heading="",captions=[]):
        """Takes selenium screenshot and adds caption text to foot.
        (First creates screenshot temporary file, then loads it
        back in, adds text using pygame, saves new file, deletes original)

        Args:
            filename - filename of screenshot file to be created
            captions - list containing captions as strings.
            Each will be written to a separate line.

        """
        driver=self.driver
        #Take screenshot
        full_filename = os.path.join(self.screenshot_folder,filename)
        driver.get_screenshot_as_file(full_filename)

        #If heading, add it to the HTML report
        if heading:
            self.scrlog.heading(heading.encode("utf-8"))

        #Add url to HTML report
        self.scrlog.add("\n<p>"+driver.current_url+"</p>\n")
        #Add impage to HTML report
        self.scrlog.image(filename,border="2")

        """
        #save temporary screenshot
        temp_filename=os.path.join(self.screenshot_folder,"temp.png")
        self.driver.get_screenshot_as_file(temp_filename)


        #Add captions to screenshot and save new version
        in_img = pygame.image.load(temp_filename)
        in_img_r = in_img.get_rect()
        font = pygame.font.Font(None,20)

        eh=0#extra height to be added to new image
        lines=[]#holds new image lines to be added
        for c in captions:
            line=font.render(c,False,(255,255,100))
            eh=eh+line.get_height()
            lines.append(line)

        #Create new image
        out_img = pygame.Surface((in_img_r.width,in_img_r.height+eh))
        out_img.fill((0,0,80))#Set background colour
        #Add original image to it
        out_img.blit(in_img,(0,0))

        #Add captions to new image
        y = in_img_r.height
        for l in lines:
            out_img.blit(l,(0,y))
            y = y+l.get_height()

        #Save
        pygame.image.save(out_img,os.path.join(self.screenshot_folder,filename))
        #delete temporary file
        os.remove(temp_filename)
        """
    def analyse_page(self):
        """Extract details of key elements on the current page.

        Returns:
            page details in one string, one line per item
        """
        driver = self.driver

        #Turned off idea to identify mandatory fields but not finished
        """
        print "########################################"
        for mand in mandatory_spans:
            es = mand.find_elements_by_tag_name("input")
            for e in es:
                print e.get_attribute("type")
                print e.get_attribute("class")
                print e.text
                try:
                    value = e.get_attribute("value")
                except Exception as ex:
                    pass
                else:
                    print value
        print "########################################"
        """

        #Fields and buttons
        fields = driver.find_elements_by_tag_name("input")

        #Text areas
        areas = driver.find_elements_by_tag_name("textarea")

        #dropdown lists
        ddlists = driver.find_elements_by_tag_name("select")

        #Holds extracted details
        details =""

        #Extract details for each element
        for e in fields+areas+ddlists:
            #Read type
            type_name = e.get_attribute("type")
            #Read id
            id_name =  e.get_attribute("id")
            #Read class
            class_name = e.get_attribute("class")
            #Read text
            text = e.text
            #Read value (Exception handling because elements might lack "value" attribute)
            try:
                value = e.get_attribute("value")
            except Exception as ex:
                value = ""
            line =  "type:"+type_name+"\ttclass:"+class_name+"\tid:"+id_name+"\tvalue:"+value+"\ttext:"+text
            details = details+line+"\n"
        return details


class Logger:
    """Used to log information to a file, mainly in HTML format
    Includes adding HTML tables.

    THis one mainly uses add() method
    """
    def __init__(self,name="log.txt"):
        self.name=name
        #Create new, empty file
        self.open(mode="w")
        self.close()

    def add(self,text,mode="a"):
        """Appends text to file
        args:
            text - text to be added
            mode - file open mode"""
        with open(self.name,mode) as f:
            f.write(text)

    def open(self,mode):
        """Open file for writing."""
        try:
            self.file=open(self.name,mode)
        except IOError as e:
            print "Failed to open file:",self.name,"(",e,")"

    def close(self):
        """close file"""
        self.file.close()

    def timewrite(self,words=""):
        """Write date/timestamp with text"""
        text = time.strftime("[%d/%m/%Y, %H:%M:%S] ")+words+"\n"
        self.add(text)

    def table(self,headings,rows):
        """Writes whole table. headings - headings, rows - list of lists, one for each row"""

        self.tabstart(headings)
        for row in rows:
            self.tabwholerow(row)
        self.tabend()

    def tabstart(self,headings):
        """Start HTML table with heading row"""
        text = '\n<table>\n'
        text = text +'<tr>\n'
        for word in headings:
            text = text + '<th>'+str(word)+"</th>\n"
        text = text + '</tr>\n'
        self.add(text)

    def tabwholerow(self,data,id=""):
        """Add complete HTML table row"""
        self.tabrow(1)
        for word in data:
            self.tabrow(2,word,id)
        self.tabrow(3)

    def tabrow(self,mode=1,datum="",id=""):
        """Add elements of table row to a file
        id - (optional) css id value"""

        if mode==1:
            text = '<tr>\n'#New row
        elif mode==2:
            if id: text='<td id="'+id+'">'
            else: text='<td>'
            text = text +str(datum)+'</td>\n'#add element to row
        elif mode==3:
            text='</tr>\n'#row end
        self.add(text)

    def tabend(self):
        """End HTML Table"""
        text = '</table>\n\n' #+ '<br>\r\n\r\n'
        self.add(text)

    def heading(self,text,id="",size='3'):
        """Add heading with id for hyperlinking"""
        size=str(size)#size needs to be a string, not a number
        text = '\n<h'+size+' id="'+str(id)+'">'+str(text)+'</h'+size+'>\n'
        self.add(text)

    def link(self,text,id):
        """Add hyperlink"""
        text = '<a href="#'+id+'">'+str(text)+'</a><br>\r\n'
        self.add(text)

    def tagger(self,text="",tags=["<p>"],mode="w"):
        """Encloses supplied string with HTML tags from supplied list.
        Can include some extra attributes with the tag eg <span id="hoot">
        These are automatically stripped from the end-tag"""
        ##endtags=["</"+t[1:] for t in reversed(tags)]
        endtags=["</"+t.split(" ")[0][1:].replace(">","")+">" for t in reversed(tags)]#this complicated to cope with removing attributes
        string= "".join(tags)+text+"".join(endtags)
        if mode=="w":
            self.add(string+"\n")
        else:
            return string

    def cssline(self,lable,values):
        """Creates a CSS style line"""
        return lable+"{"+"".join([v+";" for v in values])+"}"

    def image(self,filename,border="",scale="",hyper=True):
        """Adds image
        Args:
            filename - image filename
            border - optional border width
            scale - optional width & height scale as % e.g. "50%"
            hyper (bool) - when true make image a hyperlink to the image file

        """
        if hyper:
            text = '<a href="'+filename+'">\n'
        else:
            text = ''

        text = text+'<img src="'+filename+'" alt="'+filename+'"'

        if border:
            text = text + ' border="'+border+ '"'
        if scale:
            text = text + ' width="'+scale+ '"'
            text = text + ' height="'+scale+ '"'
        #end of image
        text = text + '>\n'
        #end of hyperlink
        if hyper:
            text = text + "</a>\n"

        self.add(text)

    def nav(self,title="Links",items=[("1","Item 1"),("2","Item 2")]):
        """Adds a nav containing in-page links

        Args:
            title (heading text for nav)

            items - items for nav as list of paris (list or tuple)
            left of each paid is the link name (without #), right of each pair
            is link text to display, eg [("1","Item 1"),("2","Item 2")]
        """
        text = "\n\n<nav>\n<b>"+title+"</b>\n<br>\n"
        for item in items:
            text = text+ '<a class="nav" href="#'+item[0]+'">'+item[1]+'</a><br>\n'
        text = text+"</nav>\n"
        self.add(text)


#Firefox Profile path. Set to None for default Webdriver profile
ffp = None

#Alan's profile
##ffp = r"C:\Documents and Settings\mayd-a\Application Data\Mozilla\Firefox\Profiles\5inetlzq.default"

#copy of Keith's profile. Works with Beta PUI bodged login
##ffp = r"C:\Documents and Settings\bky41l\Application Data\Mozilla\Firefox\Profiles\q1ah1cwy.default"

#Copy of Alan's profile
##ffp = r"E:\Test\Products\Technical Infrastructure\Redcentric Hosting\02TestDesign\Automation\SeleniumScripts\CCMS\RedcentricSTG10attempt\firefox_profile"

#Same as above but path obtained automatically from cwd.
##ffp = os.path.join(os.getcwd(),"firefox_profile")

#Excel Filename
##excel_filename = "FAM_3 Housing(1.0).xlsx"
##excel_filename = "ccms_application_and_page_data(lv0.2)run_options.xlsx"
excel_filename = "FAM_1(1.0).xlsx"
#Run script
myrun = CCMS(excel_filename=excel_filename, ffprofile=ffp, stuck_page_retries=1)