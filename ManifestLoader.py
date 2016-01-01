from tkinter import *
from tkinter import ttk
import time
import string
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlrd

def go(*args):
    # Global variable assignments.
    languages = ['', 'en', 'es', 'fr', 'zh', 'zht', 'id', 'th', 'ar', 'ja', 'ko', 'pt', 'vi', 'help']
    lms_names = ['gmt', 'gmio', 'urbsci', 'delco', 'sch', 'ad','ray', 'help']
    vendor_names = ['adl', 'blackdot', 'blupert', 'carlson', 'crognos crown',
                    'cdk', 'digitalthink', 'emeas', 'gailrice', 'gp_mx', 'ibm',
                    'intermezzon', 'jacksondawson', 'sgmit', 'biworldwide',
                    'jmorton', 'gmarg', 'alteris', 'bcanvas', 'xiangda', 'overlap',
                    'jdpa', 'jdpower', 'kms,' 'leadingway', 'manpower', 'maritz',
                    'maritz_ca', 'mcgill', 'mcm', 'mediag', 'mediawurks',
                    'micropower', 'ntl', 'pod', 'primedia', 'raytheon', 'rtsc',
                    'rps texas', 'sandycorp', 'scp', 'sewells', 'sgmntwk', 'sgmw',
                    'sigmatech', 'skillsoft', 'tta', 'vision', 'vucom', 'webaula', 'help']

    language_help = "'ar' for Arabic, 'es' for Spanish, 'fr' for French, 'id' for Indonesian, 'ja' for Japanese, 'ko' for Korean, 'pt' for Portuguese, 'th' for Thai, 'vi' for Vietnamiese, 'zh' for Chinese, or 'zht' for Traditional Chinese."
    lms_help = "'delco' for AC Delco, 'ad' for ALLDATA, 'gmt' for GM Training, 'gmio' for GM IO, 'sch' for Schwab, 'ray' for RayLMS, or 'urbsci' for Urban Science."
    vendor_help = 'List of potential vendors to enter: adl, blackdot, blupert, carlson, crognos crown, digitalthink, emeas, gailrice, gp_mx, ibm, intermezzon, jacksondawson, sgmit, biworldwide, jmorton, gmarg, alteris, bcanvas, xiangda, overlap, jdpa, jdpower, kms, leadingway, manpower, maritz, maritz_ca, mcgill, mcm, mediag, mediawurks, micropower, ntl, pod, primedia, raytheon, rtsc, sandycorp, scp, sewells, sgmntwk, sgmw, sigmatech, skillsoft, tta, vision, vucom, or webaula'

    course_number = GUIcourseNumber.get()
    language = GUIlanguage.get()
    lms = GUILMS.get()
    vendor = GUIvendor.get()

    # Function defined to concatenate the manifest file path.
    def buildManifest(course_number, language, lms, vendor):
        if lms.lower() == 'gmt':
            if language == '' or language == 'en':
                _manifest = 'http://courses.gmtraining.com/GM/' + vendor + '/' + course_number + '/imsmanifest.xml'
            else:
                _manifest = 'http://courses.gmtraining.com/GM/' + vendor + '/' + course_number + '-' + language + '/' + 'imsmanifest.xml'

        elif lms.lower() == 'gmio':
            if language == '' or language == 'en':
                _manifest = 'https://www.gmiotraining.com/gmtscormcourses/' + vendor + '/' + course_number + '/imsmanifest.xml'
            else:
                _manifest = 'https://www.gmiotraining.com/gmtscormcourses/' + vendor + '/' + course_number + '-' + language + '/' + 'imsmanifest.xml'

        elif lms.lower() == 'urbsci':
            if language == '' or language == 'en':
                _manifest = 'https://www.urbanscienceuniversity.com/scormcourses/' + vendor + '/' + course_number + '/imsmanifest.xml'
            else:
                _manifest = 'https://www.urbanscienceuniversity.com/scormcourses/' + vendor + '/' + course_number + '-' + language + '/' + 'imsmanifest.xml'

        elif lms.lower() == 'delco':
            if language == '' or language == 'en':
                _manifest = 'https://www.acdelcotraining.com/scormcourses/courses/' + vendor + '/' + course_number + '/' + 'imsmanifest.xml'
            else:
                _manifest = 'https://www.acdelcotraining.com/scormcourses/courses/' + course_number + '-' + language + '/' + 'imsmanifest.xml'

        elif lms.lower() == 'ray':
            if language == '' or language == 'en':
                _manifest = 'http://raylms.com/scormcourses/Raytheon/' + course_number + '/' + 'imsmanifest.xml'
            else:
                _manifest = 'http://raylms.com/scormcourses/Raytheon/' + course_number + '-' + language + '/' + 'imsmanifest.xml'

        else:
            print('The manifest failed to concatenate!')

        return _manifest

    manifest = buildManifest(course_number, language, lms, vendor)

    # Function defined to read login information for each LMS from a text file title logins.
    # The function creates a dictionary (lms_list) where the keys are the LMS name and each key contains two values one for user name and one for password.


    # Function defined to determine which LMS manifest loader the browser should navigate to.
    def getManifestUrl(lms):
        manifest_url = ''

        if lms == 'gmt':
            manifest_url = 'https://www.centerlearning.com/Scorm/ManifestLoader/Login.aspx'

        elif lms == 'gmio':
            manifest_url = 'https://www.gmiotraining.com/asp/Scorm/ManifestLoader/Login.aspx'

        elif lms == 'delco':
            manifest_url = 'https://www.acdelcotraining.com/scorm/manifestloader/login.aspx'

        elif lms == 'urbsci':
            manifest_url = 'https://www.urbanscienceuniversity.com/asp/scorm/ManifestLoader/Login.aspx'

        elif lms == 'ray':
            manifest_url = 'https://www.raylms.com/asp/scorm/manifestloader/login.aspx'

        elif lms == 'ad':
            manifest_url = 'https://training.alldata.com/asp/scorm/manifestloader/login.aspx'

        else:
            print('Something must be wrong with logins.txt! (getManifestURL)')

        return manifest_url

    url = getManifestUrl(lms)

    # The only way select the correct radio button is by value. This function determines the coresponding radio button value for each language.
    def getLanguageValue(language):
        if language == '' or language == 'en':
            language_value = ".//input[@type='radio' and @value='9']"
            #language_value = 9

        elif language == 'es':
            language_value = ".//input[@type='radio' and @value='10']"
            #language_value = 10

        elif language == 'fr':
            language_value = ".//input[@type='radio' and @value='12']"
            #language_value = 12

        elif language == 'ar':
            language_value = ".//input[@type='radio' and @value='1']"
            #language_value = 1

        elif language == 'id':
            language_value = ".//input[@type='radio' and @value='33']"
            #language_value = 33

        elif language == 'ja':
            language_value = ".//input[@type='radio' and @value='17']"
            #language_value = 17

        elif language == 'ko':
            language_value = ".//input[@type='radio' and @value='18']"
            #language_value = 18

        elif language == 'zh':
            language_value = ".//input[@type='radio' and @value='4']"
            #language_value = 4

        elif language == 'zht':
            language_value = ".//input[@type='radio' and @value='31748']"
            #language_value = 31748

        elif language == 'th':
            language_value = ".//input[@type='radio' and @value='30']"
            #language_value = 30

        elif language == 'vi':
            language_value = ".//input[@type='radio' and @value='42']"
            #language_value = 42

        else:
            print('The language value could not be generated!')

        return language_value

    language_value = getLanguageValue(language)

    # these functions will grab the username and password from the excel file depending on which LMS is chosen
    usr =''
    pwd =''
    book = xlrd.open_workbook('Passwords.xls')
    sheet = book.sheet_by_index(0)
    def getUser(l):
        if l == 'gmt':
            usrn = str(sheet.cell(1,1).value)
            if ".0" in usrn:
                usrn = usrn[:-2]
        elif l == 'gmio':
            usrn = str(sheet.cell(2,1).value)
            if ".0" in usrn:
                usrn = usrn[:-2]
        elif l == 'urbsci':
            usrn = str(sheet.cell(3,1).value)
            if ".0" in usrn:
                usrn = usrn[:-2]
        elif l == 'delco':
            usrn = str(sheet.cell(4,1).value)
            if ".0" in usrn:
                usrn = usrn[:-2]
        elif l == 'sch':
            usrn = str(sheet.cell(5,1).value)
            if ".0" in usrn:
                usrn = usrn[:-2]
        elif l == 'ad':
            usrn = str(sheet.cell(6,1).value)
            if ".0" in usrn:
                usrn = usrn[:-2]
        elif l == 'ray':
            usrn = str(sheet.cell(7,1).value)
            if ".0" in usrn:
                usrn = usrn[:-2]
        return usrn

    usr = getUser(lms)

    def getPass(p):
        if p == 'gmt':
            fpassword = str(sheet.cell(1,2).value)
            if ".0" in fpassword:
                fpassword = fpassword[:-2]
        elif p == 'gmio':
            fpassword = str(sheet.cell(2,2).value)
            if ".0" in fpassword:
                fpassword = fpassword[:-2]
        elif p == 'urbsci':
            fpassword = str(sheet.cell(3,2).value)
            if ".0" in fpassword:
                fpassword = fpassword[:-2]
        elif p == 'delco':
            fpassword = str(sheet.cell(4,2).value)
            if ".0" in fpassword:
                fpassword = fpassword[:-2]
        elif p == 'sch':
            fpassword = str(sheet.cell(5,2).value)
            if ".0" in fpassword:
                fpassword = fpassword[:-2]
        elif p == 'ad':
            fpassword = str(sheet.cell(6,2).value)
            if ".0" in fpassword:
                fpassword = fpassword[:-2]
        elif p == 'ray':
            fpassword = str(sheet.cell(7,2).value)
            if ".0" in fpassword:
                fpassword = fpassword[:-2]
        return fpassword
    pwd = getPass(lms)



    # This is a function that uses selenium webdriver to create a broswer instance that automates
    # all of the navigation for the user then enters the path to the course manifest file.
    def navigateToManifestUrl(url, usr, pswd, course_number, lms, language_value, manifest_path):
        browser = webdriver.Firefox()
        browser.get(url)
        username_field = browser.find_element_by_id('txtLogin')
        username_field.send_keys(usr)
        password_field = browser.find_element_by_id('txtPassword')
        password_field.send_keys(pswd)
        submitBtn = browser.find_element_by_id('btnSubmit')
        submitBtn.click()
        search_field = browser.find_element_by_id('txtSearch')
        search_field.send_keys(course_number)
        searchBtn = browser.find_element_by_id('btnSubmit')
        searchBtn.click()
        course_link = browser.find_element_by_class_name('linkBlue')
        course_link.click()
        radioBtns = browser.find_element_by_xpath(language_value)
        radioBtns.click()
        manifest_field = browser.find_element_by_id('txtManifestPath')
        manifest_field.send_keys(manifest_path)

    navigateToManifestUrl(url, usr, pwd, course_number, lms, language_value, manifest)

# Below is the code for the GUI

root = Tk()
root.title("Manifest Loader")
root.geometry('350x300')

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

GUIcourseNumber = StringVar()
GUILMS = StringVar()
GUIvendor = StringVar()
GUIlanguage = StringVar()
GUIloadOption = StringVar()

# Go Button Creation
ttk.course_entry = Entry(mainframe, width=15, textvariable=GUIcourseNumber)
ttk.course_entry.grid(column=1,row=2)
ttk.Button(mainframe, text="Go", command=go).grid(column=3, row=7, sticky=W)

# Label Creation
ttk.Label(mainframe, text="Course Number").grid(column=1,row=1,pady=(30,0))
ttk.Label(mainframe, text="LMS").grid(column=1,row=3,sticky=W,padx=(50,50),pady=(30,0))
ttk.Label(mainframe, text="Vendor").grid(column=1,row=5,pady=(30,0))
ttk.Label(mainframe, text="Translated Language").grid(column=3,row=1,pady=(30,0))
ttk.Label(mainframe, text="Loader Options").grid(column=3,row=3,pady=(30,0))

# Spacer Creation
#ttk.Label(mainframe, text="|").grid(column=2, row=1)
#ttk.Label(mainframe, text="|").grid(column=2, row=2)
#ttk.Label(mainframe, text="|").grid(column=2, row=3)
#ttk.Label(mainframe, text="|").grid(column=2, row=4)
#ttk.Label(mainframe, text="|").grid(column=2, row=5)
#ttk.Label(mainframe, text="|").grid(column=2, row=6)

# Drop-Down Menu Creation
w = OptionMenu(mainframe, GUILMS, 'gmt', 'gmio', 'urbsci', 'delco', 'sch', 'ad', 'ray').grid(column=1, row=4)
w = OptionMenu(mainframe, GUIvendor, 'adl', 'blackdot', 'blupert', 'carlson', 'crognos crown',
                    'cdk', 'digitalthink', 'emeas', 'gailrice', 'gp_mx', 'ibm',
                    'intermezzon', 'jacksondawson', 'sgmit', 'biworldwide',
                    'jmorton', 'gmarg', 'alteris', 'bcanvas', 'xiangda', 'overlap',
                    'jdpa', 'jdpower', 'kms,' 'leadingway', 'manpower', 'maritz',
                    'maritz_ca', 'mcgill', 'mcm', 'mediag', 'mediawurks',
                    'micropower', 'ntl', 'pod', 'primedia', 'raytheon', 'rtsc',
                    'rps texas', 'sandycorp', 'scp', 'sewells', 'sgmntwk', 'sgmw',
                    'sigmatech', 'skillsoft', 'tta', 'vision', 'vucom', 'webaula').grid(column=1,row=6)
w = OptionMenu(mainframe, GUIlanguage, 'en', 'es', 'fr', 'zh', 'zht', 'id', 'th', 'ar', 'ja', 'ko', 'pt', 'vi').grid(column=3, row=2)
w = OptionMenu(mainframe, GUIloadOption, 'Someday').grid(column=3, row=4)

root.mainloop()