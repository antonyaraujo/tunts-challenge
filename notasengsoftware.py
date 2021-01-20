import gspread
from oauth2client.service_account import ServiceAccountCredentials

# function responsible by the authentication and import of the spreadsheet
def importSheet():
    scope = ['https://spreadsheets.google.com/feeds'] #Used scope
    credentials = ServiceAccountCredentials.from_json_keyfile_name('engsoftware-d33263fc40c6.json', scope) #API Credentials
    gc = gspread.authorize(credentials)
    return gc.open_by_key('1XEaMipbo4zPXLO-jZ5LmIFFwTUf5lYWmgu9zE6uXbEY') # returns an open sheet

# function that returns the status of the student by grade
def studentStatus(grade):
    if (grade<50):
        return "Reprovado por Nota"
    elif (50<=grade<70):
        return "Exame final"
    elif (grade>=70):
        return "Aprovado"

# function that sum the three grades and returns the average between them
def computeAverages():
    for p in range(len(P1)):
        if (p+1 > 3):
            average = (int(P1[p]) + int(P2[p]) + int(P3[p]) )/ 3
            averages.append(average)

# function that process the student situation by average and absences and make the returns to the sheet
def updateCells():
    for i in range(3, len(P1)):
        print("Aluno: %s [m = %s]" %(students[i], "{0:.2f}".format(round(float(averages.__getitem__(i-3)), 2))))
        if(int(absences[i]) > 60*0.25): #checks if the amount of absences is bigger than 25% of the total classes
            grades.update_cell(i+1, 7, "Reprovado por Falta") #if it's true the cell is written with correspondence message
            grades.update_cell(i+1, 8, 0) # and the cell of naf receive the value of zero
        else:
            average = float(averages.__getitem__(i-3))
            grades.update_cell(i + 1, 8, 0) # it puts naf value as zero
            if(50 <= average < 70): # checks if the average its between 50 and 70
                naf = 100-average
                decimal = naf - int(naf)
                if (decimal != 0):
                    naf = naf + (1-decimal)
                grades.update_cell(i+1, 8, "{0:.2f}".format(round(naf, 2))) #updates the cell value of naf
            grades.update_cell(i+1, 7, studentStatus(average)) #update the student status
        print("Situaçao: %s [Naf = %s] \n" %(grades.cell(i+1, 7).value, grades.cell(i+1, 8).value))

wks = importSheet() #gets the sheet into  variable
grades = wks.get_worksheet(0) # gets the sheet's page into a variable
absences = grades.col_values(3) #gets the column with absences
students = grades.col_values(2) # gets the students's names column
P1 = grades.col_values(4) # gets the column with P1 grades
P2 = grades.col_values(5) # gets the column with P2 grades
P3 = grades.col_values(6) # gets the column with P3 grades
averages = [] #initialize a list for the averages
computeAverages() #calls the function that compute the averages and insert into the average's list
updateCells() # put the results processed into the spreadsheet's cells