#imports:
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
from selenium import webdriver

#stores all the courses from the specified file into the courses array - note that the file needs to be in the same directory in order for this to work
def readCourses(courses, docName):
    #checks to see if the input is an actual integer
    def isInt(n):
        try:
            int(n)
            return True
        except ValueError:
            return False
    
    with open(docName, 'r') as text:
        lines = text.readlines()

        for line in lines:
            lLength = len(line)

            #checks to see if a line has 4 or more characters - since courses need 3 characters for the number and at least 1 character for the course code (eg CSC)
            if lLength >= 4:
                #stores the last 3 letters of a valid (having more than 4 characters) line as a variable
                courseNum = line[lLength - 4:]

                #this checks to see if courseNum is greater than 100 since there were some lines in the document that were just "100000", and though the last 3 characters there are ints, that obviously isn't a course code
                if isInt(courseNum) and int(courseNum) >= 100:
                    courses.append(line[:lLength - 1])

#this takes the x and y positions and returns a string that we can use to specify a specific excel cell - I just made it it's own function so the getWebsiteCourses function doesn't get too cluttered
def getSheetPos(sheetX, sheetY):
    if sheetX == 1:
        returnX = "A"
    elif sheetX == 2:
        returnX = "B"
    elif sheetX == 3:
        returnX = "C"
    elif sheetX == 4:
        returnX = "D"
    elif sheetX == 5:
        returnX = "E"
    elif sheetX == 6:
        returnX = "F"
    elif sheetX == 7:
        returnX = "G"
    elif sheetX == 8:
        returnX = "H"
    elif sheetX == 9:
        returnX = "I"
    elif sheetX == 10:
        returnX = "J"
    elif sheetX == 11:
        returnX = "K"
    elif sheetX == 12:
        returnX = "L"
    elif sheetX == 13:
        returnX = "M"
    elif sheetX == 14:
        returnX = "N"
    elif sheetX == 15:
        returnX = "O"
    elif sheetX == 16:
        returnX = "P"
    elif sheetX == 17:
        returnX = "Q"
    elif sheetX == 18:
        returnX = "R"
    elif sheetX == 19:
        returnX = "S"
    elif sheetX == 20:
        returnX = "T"
    elif sheetX == 21:
        returnX = "U"
    elif sheetX == 22:
        returnX = "V"
    elif sheetX == 23:
        returnX = "W"
    elif sheetX == 24:
        returnX = "X"
    elif sheetX == 25:
        returnX = "Y"
    elif sheetX == 26:
        returnX = "Z"
    elif sheetX == 27:
        returnX = "AA"

    return returnX + str(sheetY)

#scrapes UVic's website for the courses, and stores them into an excel sheet
def getWebsiteCourses(courses, driver, worksheet, username, password):
    #sets up 2 "global" elements that can't are needed in the loop below, but can't be reset by said loop -- the array that stores the names of the courses that we can't find, and the y position of the sheet that we have to write the elements on
    coursesNotFound = []
    sheetY = 1

    #logs into the UVic website
    driver.get("https://www.uvic.ca/tools/student/index.php")
    
    username = driver.find_element_by_id("username")
    username.send_keys(username)

    password = driver.find_element_by_id("password")
    password.send_keys(password)
    
    submit = driver.find_element_by_id("form-submit")
    submit.click()

    
    #dismisses the cookies icon so we won't have to deal w/ it in the loop
    driver.get("https://www.uvic.ca/tools/student/registration/look-up-classes/index.php")

    cookiesButton = driver.find_element_by_id("cookies-btn")
    cookiesButton.click()

    #looks for each course on the site, and puts the course's information on an excel doc
    for course in courses:
        driver.get("https://www.uvic.ca/tools/student/registration/look-up-classes/index.php")

        #switches to the needed iframe so that we can actually interact with the elements that we want to
        driver.switch_to.frame("SSBFrame")

        #finds the newest term from the drop-down menu, and clicks on it
        termInput = driver.find_element_by_css_selector("#term_input_id")
        terms = termInput.find_elements_by_tag_name("option")
        terms[1].click()

        submit = driver.find_element_by_xpath("/html/body/div[3]/form/input[2]")
        submit.click()

        #clicks on the "advanced search" option - this will help us find each specific course much easier
        adv = driver.find_element_by_xpath("/html/body/div[3]/form/input[18]")
        adv.click()

        #defines the course code and course number as seperate variables, since those are entered seperatley in advanced search
        divPoint = len(course)-3
        courseCode = course[:divPoint]
        courseNum = course[divPoint:]

        #finds and clicks on the course code
        subjectList = driver.find_element_by_id("subj_id")
        subjects = subjectList.find_elements_by_tag_name("option")
        for subject in subjects:
            if subject.get_attribute("value") == courseCode:
                subject.click()

        #enters the course number
        enterNum = driver.find_element_by_id("crse_id")
        enterNum.send_keys(courseNum)

        submit = driver.find_element_by_xpath("/html/body/div[3]/form/span/input[1]")
        submit.click()

        #now we should be on the page containing all the information about the course formatted in a table - from here we are able to find each cell in the table since they all have the same class name
        table = driver.find_element_by_xpath("/html/body/div[3]/form/table/tbody")
        cells = table.find_elements_by_class_name("dddefault")

        #if there is table content, (meaning we didn't get an error messaage when looking up courses) then put it in the spreadsheet
        if len(cells) > 0:
            sheetX = 1
            for cell in cells:
                #gets the next appropriate sheet position
                sheetPos = getSheetPos(sheetX, sheetY)

                worksheet[sheetPos] = str(cell.text)

                #checks to see if we need a new line (done by incrementing sheetY and resetting sheetX) - there's 22 cells per row on UVic's website
                if sheetX < 22:
                    sheetX += 1
                else:
                    sheetX = 1
                    sheetY += 1
        #if there isn't table content, then add it to the coursesNotFound array
        else:
            coursesNotFound.append(course)
    
    #if we weren't able to find some courses, then add a row on the bottom of the spreadsheet specifying which courses we couldn't find
    if len(coursesNotFound) > 0:
        #I'm inserting this message at the start of the coursesNotFound array since I thought it would be easier to do that and put every entry of the array on the sheet than just put the message down first and then put down every entry of the array of the spreadsheet down as well
        coursesNotFound.insert(0, "could not find these courses:")
        sheetX = 1
        for course in coursesNotFound:
            sheetPos = getSheetPos(sheetX, sheetY)
            worksheet[sheetPos] = coursesNotFound[sheetX-1]
            sheetX += 1

def seleniumScraperMain(username, password):
    #sets up our webdriver - note that this needs the driver to be in the same directory as this file
    driver = webdriver.Chrome()

    #creates and fills the courses array using the specified text document
    courses = []
    readCourses(courses, 'm4.v21Jan7.famc2.txt')

    #sets up a workbook for us to do work in inside the getWebsiteCourses function
    workbook = Workbook()
    worksheet = workbook.active

    #scrapes the UVic website and dumps the results into an excel document
    getWebsiteCourses(courses, driver, worksheet, username, password)

    #saves the workbook in this file's directory
    workbook.save("cutDownCourses.xlsx")

    driver.close()