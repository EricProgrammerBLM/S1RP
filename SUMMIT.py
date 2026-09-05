import openpyxl
from time import sleep
from datetime import date
from openpyxl import Workbook, load_workbook
import xlsxwriter 
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string #Helps us get the Column Number
import string
from openpyxl.utils.dataframe import dataframe_to_rows
import random
from natsort import natsorted #Will help put list in order with letters and numbers



'''
Copyright (c) 2026 Eric Byam. All Rights Reserved.

This software and its source code are proprietary. Unauthorized copying, 
distribution, modification, or execution of this software, via any medium, 
is strictly prohibited without the express written permission of the copyright owner.
'''


print ('As you Add Changes, Dont forget to literally SAVE FILE inside the code, or you wont notice.... SAVE THE FILE')

File_Name = 'C:/Users/ericb/OneDrive/Desktop/Rotation_Test_V1.xlsx'
#            C:\Users\ericb\OneDrive\Desktop
#            'C:/Users/ericb/OneDrive/Desktop/Rotation_Test_V1.xlsx' - For Work Laptop
#            'C:/Users/John Doe/Desktop/Code/Python/Rotation_Test_V1.xlsx' - For Home Laptop

wb = xlsxwriter.Workbook(File_Name)
wb.close()

wb = load_workbook(File_Name)
ws = wb.active

#Cell Style
border = Border(
    left=Side(border_style='thin', color='040005'),  # Adjust border style and color as needed
    right=Side(border_style='thin', color='040005'),
    top=Side(border_style='thin', color='040005'),
    bottom=Side(border_style='thin', color='040005')
)



def Create_Post(cells_to_merge, the_row, the_col, post):
    # ADDED THIS CONDITION:
    # Only merge if ":" is present. This prevents "A100" or "Header" from merging.
    if ":" in cells_to_merge:
        ws.merge_cells(cells_to_merge)
        
        # Apply borders to the whole range
        merged_cells_range = ws[cells_to_merge]
        for row in merged_cells_range:
            for cell in row:
                cell.border = border
    else:
        # If there is no colon, it's just a single cell.
        # We just apply the border to that one specific cell.
        cell = ws.cell(row=the_row, column=the_col)
        cell.border = border

    # This part handles the text and centering for BOTH merged and single cells.
    cell = ws.cell(row=the_row, column=the_col)
    cell.value = post
    cell.alignment = Alignment(horizontal='center', vertical='center')




Create_Post('A1:V1', 1, 1, 'GUEST EXPERIENCE DAILY POST ROTATIONS PM')
Create_Post('A2:V2', 2, 1, str(date.today()))
Create_Post('A3:D3', 3, 1, 'GE TEAM')
Create_Post('E3:F3', 3, 5, 'DOORS OPEN')

wb.save(File_Name)




'''


#-----------------Best for finding Merged Cells; May be useful later
# Iterate through each worksheet
for sheet_name in ws:
    sheet = ws
    
    # Create a list to store merged cell ranges
    merged_cells = []
    
    # Iterate through all rows and columns in the worksheet
    for row in sheet.iter_rows():
        for cell in row:
            # Check if the cell is part of a merged cell
            if cell.coordinate in sheet.merged_cells:
                # If the cell is in a merged range, add the merged range to the list
                merged_cells.append(cell.coordinate)
    

print('')
#Regular for loops wouldn't work with this list to remove items, so these statements were used instead
merged_cells = [item for item in merged_cells if '1' not in item]
merged_cells = [item for item in merged_cells if '2' not in item]
merged_cells = [item for item in merged_cells if '3' not in item]

print (merged_cells)
'''








#Cell Colors Based on Floors and Rotation Post @ SUMMIT
B1 = 'F8CBAD'
OB1 = 'B4C6E7'
OB2 = 'FFD966'
OB3 = 'C6E0B4'
BREAK = 'FFFF00'
BRIEF = 'AEAAAA'
FLOAT = '00B0F0'
LEAD_SHIFT_COLOR = '92D050'
REG_SHIFT_COLOR = 'D6EAF8'           
FT_SHIFT_COLOR = 'F1C0F2'
TILL_DOORS_CLOSE_SHIFT_COLOR = 'B4C6E7'
BREAKER_SHIFT_COLOR = 'FFFACD'
BLACK = '060300'
FREIGHT_COLOR = 'BD82C7'
PREP = 'C2371E'
RADIO = 'EBBE32'
CLEAR_COLOR = 'E3CDF7'

#Cell Styles
Center_Text = Alignment(horizontal= 'center', vertical='center')




'''
List of Post in 2 seperate Tiers. Tiers are based on level of importance. 
Tier 1's Post can never go without someone being there

Tier 2's will be in order from most expendable post, to least expandle post, 
but overall every post in Tier 2 can go on without someone being there, if their are massive call outs 
'''

#Update for this began 6/17/25
B1_Tier_1 = ['RISE TABLET','RISE LINE', "HELLO 1", 'SHOES 1', 'CELEBRATE', 'PLAZA VANDY', 'PLAZA MAD', 'HELLO 2**', 'QUEUE', 'KIOSK', 'TURNSTILE', 'WELCOME', 'TICKET CHECK'] #Updated
B1_Tier_2 = ['HELLO 3**', 'TURNSTILE 2**','SHOES 2**', 'SHOE PREP**', 'RISE LINE 2**', 'FACESCAN**', 'SHOE PREP**', 'SHOE PREP**', 'SHOE PREP**' ] #Updated Last 10/25
B1_Tier_3 = ['SHOES 2**', 'SHOE PREP**', 'LAUNCH**'] #Shouldn't be a Tier 3, only added for the purpose of OB1 and OB2. May keep this when we re-write B1
Mix_Of_Both_B1_Tiers = []
Mix_Of_Both_B1_Tiers.extend(B1_Tier_1 + B1_Tier_2 + B1_Tier_3) 
#If this updated list above somehow causes an error, dig this up from one of our previous versions of this code. This is needed specifically for when we borrow a B1 Post and Use Celebrat to help with upstairs rotations


LEAD_ENDING_POST_B1 = ['LAUNCH**', 'QUEUE', 'RISE LINE'] #Can possibly delete this after B1 is REWRITTEN

OB1_Tier_1 = ['AFF 1', 'TR1**', 'REFL', 'ESC', 'AFF 2'] #Never Edit... Unless an Important Post is needed to be Added. At some point we can add that AFF1 starts rotations under conditions it goes to Ref and Ref goes to the other Aff and that person relieves ESC
OB1_Tier_2 = ['S ELE**', 'TR2**', 'TR3**', 'TR4**', 'TR5**', 'AFF 3**'] #If you ever edit, remember to check through the code that involves AFF 3, I forgot which line so search for the string AFF3

OB2_Tier_1 = ['UNITY', 'SHOE REMOVAL', 'EXIT LINE', 'EXIT TAB', 'SKYBRIDGE', 'OB3 MOMENTO', 'CAB 1', 'CAB 2**', 'ASCENT DESK'] #Add: 'CAB 1', 'CAB 2', 'ASCENT DESK' - Never Edit... Unless an Important Post is needed to be Added
OB2_Tier_2 = ['LEVITATION**', 'UNITY 2**', 'OB2 SERVICE**', 'EXIT LINE 2**', 'APRES**', 'OB3 SERVICE**'] #Could potentially be an issue, since it may start the rotations for this floor.

OB3_Tier_1 = [] 

Full_Timers_4PM_Post_On_B1 = [ 'PREP**', 'PREP**', 'PREP**', 'PREP**', 'PREP**', 'PREP**']
Full_Timers_4PM_Post_Upstairs = ['E TAB', 'CAB 1', 'CAB 2**', 'ASC DESK', 'AFF 3**', 'PREP', 'PREP', 'PREP', 'PREP']
#This will be filled with post to give full timers for the PM Briefing.

Full_Timers_Opening = []
Full_Timers_Opening_Tier_1 = []
Full_Timers_Opening_Tier_2 = []

Overall_OB1_Post = OB1_Tier_1 + OB1_Tier_2
Overall_OB2_Post = OB2_Tier_1 + OB2_Tier_2

FREIGHT = ['B1 BOH FREIGHT', 'FREIGHT PULLER', 'FREIGHT COUNTER', 'FREIGHT RIDER'] #Old List Order - ['B1 BOH FREIGHT', 'FREIGHT PULLER', 'FREIGHT RIDER', 'FREIGHT RIDER', 'FREIGHT COUNTER']

FLOAT_LIST = ['FLOAT', 'FLOAT', 'FLOAT', 'FLOAT', 'FLOAT', 'FLOAT', 'FLOAT', 'FLOAT', 'FLOAT', 'FLOAT']

EVERYTHING_ELSE = ['FLOAT', 'BRIEF', 'CKLST', 'BREAK']

WORK_TIMES = ['4:00pm-12:00am', '4:00pm-11:00pm', '6pm-10:30pm', 'LEAD', '3:30pm-12:00pm'] #This will get fixed last....
ROTATION_STARTERS = ['SHOE PREP**', 'LAUNCH**', 'SHOES 2**', 'PREP', 'ESC**', 'TR1**', 'S ELE**', 'TR2**', 'TR3**','TR4**', 'TR5**','AFF 3**', 'LEVITATION**', 'UNITY 2**', 'BREAK', 'BRIEF', EVERYTHING_ELSE[0], 'FREIGHT PULLER', 'FREIGHT RIDER', 'FREIGHT RIDER', 'FREIGHT COUNTER', 'OB2 SERVICE**', 'EXIT LINE 2**', 'SKIP', 'RISE LINE 2**', 'HELLO 3**', 'SHOES 2**', 'SHOE PREP**', 'FACESCAN**', 'B1 BOH FREIGHT', 'HELLO 2**', 'TURNSTILE 2**', 'RADIO', 'CAB 2**']
#POTENTIAL ERROR READ: B1 FREIGHT CAN NOT BE ADDED TO THIS YET....... Leads to an ERROR on B1's Rotation relief Function for some reason... not SURE WHY. Add it when we re-do B1

RANDOMNESS = [1, 0, 0] #This list will literally be used just to add randomness to certain post being picked out

END_ROTATION_POST = []
#A list made up of all rotation post that should end it.

#Makes sure everything in FREIGHT and EVERYTHING ELSE list are recognized as Rotation enders, even if they're updated
for a in EVERYTHING_ELSE:
	ROTATION_STARTERS.append(a) #Can possibly delete later

for post in EVERYTHING_ELSE + FREIGHT + ROTATION_STARTERS:
	END_ROTATION_POST.append(post)




AM = ['8:00AM', '8:30AM', '9:00AM', '9:30AM', '10:00AM', '10:30AM', '11:00AM', '11:30AM', '12:00PM', '12:30PM',
'1:00PM', '1:30PM', '2:00PM', '2:30PM', '3:00PM', '3:30PM', '4:00PM', '4:30PM', '5:00PM']
PM = ['3:00PM','3:30PM', '4:00PM', '4:30PM', '5:00PM', '5:30PM', '6:00PM', '6:30PM', '7:00PM', '7:30PM', '8:00PM', '8:30PM',
'9:00PM', '9:30PM', '10:00PM', '10:30PM', '11:00PM', '11:30PM', '12:00PM'] #Added 3PM and removed , '12:30PM'
ALPHABET = list(string.ascii_lowercase)
ALPHABET = [letter.upper() for letter in ALPHABET]


#Creates and Colors the floor in the rotation
B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
ws['A4'].fill = B1_Cell
ws['A5'].fill = B1_Cell
Create_Post('A4:C4', 4, 1, 'B1')
Create_Post('A5:C5', 5, 1, 'B1')

OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
ws['A6'].fill = OB1_Cell
Create_Post('A6:C6', 6, 1, 'OB1')

OB2_Cell = PatternFill(patternType = 'solid', fgColor = OB2)
ws['A7'].fill = OB2_Cell
Create_Post('A7:C7', 7, 1, 'OB2')

OB3_Cell = PatternFill(patternType = 'solid', fgColor = OB3)
ws['A8'].fill = OB3_Cell
Create_Post('A8:C8', 8, 1, 'OB3')




Times = 19
Alpha = 2 #Short for Alphabet, which is a list with every letter in the alphabet. We will be starting at C
Rotate = 0
Floor_Rows = '4'
Scheduled_Rotations = 0
#Also we can add all of these numbers into 1 list to make the code look cleaner instead of multiple variables
#Go back to the old format if a solution cant be figured out
#Update: It works

Time_Periods = [19, 3, 0, '4', 0] #3 in this list was 2 Before; this controlled which column the time stamps for rotation times would start
while Time_Periods[0] != 0 and Time_Periods[4] != 5: #Possibly turn this into a function
	ws[ALPHABET[Time_Periods[1]] + Time_Periods[3]].value = PM[Time_Periods[2]] #Had .upper here at the end of both ALPHABETs, no longer needed
	ws[ALPHABET[Time_Periods[1]] + Time_Periods[3]].alignment = Center_Text
	ws[ALPHABET[Time_Periods[1]] + Time_Periods[3]].border = border
	#wb.save(File_Name)
	Time_Periods[2] += 1
	Time_Periods[1] += 1
	Time_Periods[0] -= 1
	if Time_Periods[0] == 0:
		Time_Periods[0] = 19
		Time_Periods[2] = 0
		Time_Periods[1] = 3 #Was 2 before; Ended up extending Floor text to 3 Cells so we had to push all the times in the row over by 1
		Time_Periods[4] += 1
		Time_Periods[3] = int(Time_Periods[3]) + 1 #This is the starting row where we will be telling the program to write in the times for the rotations Ex. 'A4'
		Time_Periods[3] = str(Time_Periods[3])     #

# Full Timers; FULL TIMERS; SHIFTS; ADP - All searchable hash tags to help find this later
Full_Time_Closers = 9 #3
Closers = 10 #Was 14 Before        
Closers = Closers + Full_Time_Closers
Closers_Untill_11 = 12 #6
#The Amount of people coming in with Regular Closing/Opening Shifts
Leads = 4 #4              #Any Leads more than 7 is an error; 8 or more is an error
Leads_2 = Leads        #Will be needed later on
#Lead shifts. Almost no different than Regular shifts except slightly different post (Floating, Radio, CHECKLIST etc)
Breakers = 8 #8,anything over 9 is a bug  
#The 6:00 - 10:30's or 10:30 - 3:00

#Priorty Should Be: 1.Leads, 2.Closers_Untill_11, Closers, Breakers
Ambassador = Closers + Closers_Untill_11 + Leads
Total = Closers + Closers_Untill_11 + Leads + Breakers #Delete Later
print ('Total All Shifts: ', Total)




'''
This function will be used to help locate index coridinates on the excel sheet. Will make mapping out
the post rotations alot easier, because we can call upon cordinates as we go along.
'''
def Locate(rotation_post_as_a_string):
	# Define the target value you want to search for
	target_value = rotation_post_as_a_string  # Replace with the value you're looking for

	# Create a list to store the coordinates of cells with the target value
	cell_coordinates = []

	# Iterate through all rows and columns in the worksheet
	for row in ws.iter_rows():
		for cell in row:
			if cell.value == target_value:
				# Get the cell's coordinate (cell address)
				cell_coordinate = cell.coordinate
				cell_coordinates.append(cell_coordinate)
	#wb.save(File_Name) Causing Permission Errors

	# Print the coordinates of cells with the target value, as well as the list
	for coordinate in cell_coordinates:
		return(cell_coordinates)


#Code Below should be the one that determines how many rows will be created under each floor based on the amount of shifts their are working. (Leads, Breakers etc)
#Create a List of the Floors we work on
#Turn this into a function later
Floors = ['B1', 'B1', 'OB1', 'OB2', 'OB3']
Floors.reverse()
'''
We reverse the list containing the Floors so that we can properly locate the cordinates of each floor and add rows below them.
Only reason why we can't add rows without the list being reversed, is due to the fact that once rows are added under B1 first
it creates an inaccuracy on where the following floors would be at, resulting in rows being added in places they shouldnt be.

It Adds Rows from the Bottom up, starting with OB3
'''


def Rotation_Rows(type_of_shift):
	Added_Rows = 0
	while type_of_shift > 0:
		Add_Rows = 0
		while type_of_shift > 0:     #Before was: while type_of_shift > 0 and Added_Rows != type_of_shift                         
			Asign_Floor = Locate(Floors[Add_Rows])                                        
			while (len(Asign_Floor) != 0):       
				ws.insert_rows(int(Asign_Floor[-1][1:]) + 1, 1) #This statement tells it where to add a row at, and how many rows to add
				Asign_Floor.remove(Asign_Floor[-1])
				Add_Rows += 1
				Added_Rows += 1
				type_of_shift -= 1
				if (Add_Rows) == len(Floors): #Was originally 5 before.
					Add_Rows = 0
				wb.save(File_Name)
	#print ('Rows Added: ', Added_Rows)


#Priorty Should Be: 1.Leads, 2.Closers_Untill_11, Closers, Breakers
#If All_Shifts are odd, it'll add an extra row; Edit: May have been fixed, no issues so far
All_Shifts = Leads + Closers + Closers_Untill_11 + Breakers
#print (All_Shifts)
Rotation_Rows(All_Shifts)

Names = All_Shifts + len(Floors) + 4
#print (Names)

Shift_Rows = [] #This list will be used for all the rows where a rotation line will be for a shift
for i in range(1, Names):
	Result = ws['A' + str(i)].value
	if Result == None:
		Shift_Rows.append('A' + str(i))
		for cell in list(ws.merged_cells.ranges): #Gives us a list of Merged Cells
			if ('A' + str(i)) in cell:      #Unmerges every cell where employee names would be. It would merge over with the next cell for an odd reason
				ws.merged_cells.remove('A' + str(i) + ':C' + str(i))

		ws['A' + str(i)].value = 'Name'     #Prints the names of employees that will work during this shift
		ws['A' + str(i)].alignment = Center_Text #Centers the our names inside the cell
		ws['A' + str(i)].border = border
		#wb.save(File_Name)



#Every Loop Below will be used to determine certain rows for Leads, Breakers, Closers/Openers, and those that may leave at 11
Lead_Rows = [Shift_Rows[0]] #Lead Row Cordinates Go in this List
Leads -=1
num = 1
Lead_Max = Leads - len(Lead_Rows)
if Leads_2 == 2: #This if statement is specifically built for it its only 2 leads. The variable name and number are not connected in any way
	Only_Other_Lead = 'A' + str(int(Locate('OB1')[0][1:]) + 1)
	print (Only_Other_Lead)
	Leads -= 1
	Lead_Max -= 1
	Lead_Rows.append(Only_Other_Lead)


#This Loop is to determine the lines/rows that belong strictly for the leads. List Name: Lead_Rows
while Lead_Max > 0: 
	for i in Shift_Rows:
		#print ('Lead Max: ', Lead_Max)
		print ('Lead Row: ', Lead_Rows)
		#print (int(i[1:]))
		#print (int(Shift_Rows[num][1:]))
		#print (int(Shift_Rows[num][1:]) - int(i[1:]))
		if int(Shift_Rows[num][1:]) - int(i[1:]) >= 2:
			Leads -= 1
			Lead_Max -= 1
			Lead_Rows.append(Shift_Rows[num])
		elif Leads <= 0:
			break
		elif Lead_Max <= 0 and Leads > 0 and len(Lead_Rows) == 5:
			break
		num += 1
	
		if len(Lead_Rows) == 5:
			Lead_Max = 0
#print (Shift_Rows)
#print (Lead_Rows)
#sleep(1998)
#This is used to continue to create more Lead Lines/Rows after the 5th one is created
while Leads > 0:
	Extra_Lead_Row = int(Lead_Rows[-1][1:]) + 1
	Lead_Rows.append('A' + str(Extra_Lead_Row))
	Leads = Leads - 1

#This list is created for a use case that we will use later on line: ; It will help us create balance on B1 with adding more rows by pulling from OB3
Create_Balance = Shift_Rows
#print ('All Rows and Shifts: ', Create_Balance)


#Removing all Lead/Breakers Line/Row's cordinates from the shift_rows list, so it'll be easier to use later
Shift_Rows = [x for x in Shift_Rows if x not in Lead_Rows]

#A List that will be used for later that shouldnt include the rows with the LEADS
Last_Row_Cordinates = Shift_Rows



#This Loop is to determine the lines/rows that belong strictly for the leads. List Name: Breaker_Rows
Breaker_Rows = [] #Breaker Row Cordinates Go in this List
num = 1
while Breakers > 0:
	for i in Shift_Rows:
		#print (int(i[1:]))
		#print (int(Shift_Rows[num][1:]))
		#print (int(Shift_Rows[num][1:]) - int(i[1:]))
		if Breakers > 8:
			if int(Shift_Rows[num][1:]) - int(i[1:]) >= 3:
				Breakers -= 1
				Breaker_Rows.append(i)
				#print ('Breakers-r: ', Breakers)
			elif int(Shift_Rows[num + 1][1:]) - int(i[1:]) >= 3:
				Breakers -= 1
				Breaker_Rows.append(i)
				#print ('Breakers-e: ', Breakers)
			elif int(Shift_Rows[num + 2][1:]) - int(i[1:]) >= 5: #This is only useful if Breakers are 8 and more
				Breakers -= 1
				Breaker_Rows.append(i)
				#print ('Breakers-p: ', Breakers)
			elif Breakers <= 0: #was <= before
				#print ('Breakers-c: ', Breakers)
				break
			num += 1
			#print ('Breakers: ', Breakers)
			#print ('Breaker Rows', Breaker_Rows)
		elif Breakers < 9:
			if int(Shift_Rows[num][1:]) - int(i[1:]) >= 3:
				Breakers -= 1
				Breaker_Rows.append(i)
				#print ('Breakers-k: ', Breakers)
			elif int(Shift_Rows[num + 1][1:]) - int(i[1:]) >= 3:
				Breakers -= 1
				Breaker_Rows.append(i)
				#print ('Breakers-w: ', Breakers)
			elif Breakers <= 0: #was <= before
				break
			num += 1

#Minor bug that sometimes causes an extra Breaker Row to be created, code below prevents that
if Breakers < 0:
	Breakers = 0
	index_breaker = Breaker_Rows[-2] #Has to be the 2nd to last one, else a 4 - 12 shift will be placed unver the breaker shift
	Breaker_Rows.remove(index_breaker)

#Removing all Lead/Breakers Line/Row's cordinates from the shift_rows list, so it'll be easier to use later
Shift_Rows = [x for x in Shift_Rows if x not in Breaker_Rows]



Closer_11_Row = [] #Closing Row Cordinates Go in this List
for i in Shift_Rows:
	if Closers_Untill_11 > 0:
		Closer_11_Row.append(i)
		Closers_Untill_11 -= 1
	else:
		break

Shift_Rows = [x for x in Shift_Rows if x not in Closer_11_Row]




#print ('Regular Shifts: ', Shift_Rows)
#print ('Lead Shifts: ', Lead_Rows)
#print ('Breaker Shifts: ', Breaker_Rows)
#print ('Closing Until 11 Shifts: ', Closer_11_Row)

Color_For_Regular_Shift = PatternFill(patternType = 'solid', fgColor = REG_SHIFT_COLOR)
Color_For_Lead_Shift = PatternFill(patternType = 'solid', fgColor = LEAD_SHIFT_COLOR)
Color_For_11_Shift = PatternFill(patternType = 'solid', fgColor = TILL_DOORS_CLOSE_SHIFT_COLOR)
Color_For_Breaker_Shift = PatternFill(patternType = 'solid', fgColor = BREAKER_SHIFT_COLOR)


def TimeStamps(time_of_the_shift_var, shift_list, cell_color):
	for i in shift_list:
	#index = ALPHABET.index(i[0])
	#Next_Col = ALPHABET[index + 1]
	#print (Next_Col)
	#print ('B' + i[1:])
	#ws[Next_Col + i[1:]].value = WORK_TIMES[0]
		ws['B' + i[1:]].alignment = Center_Text
		ws['B' + i[1:]].fill = cell_color
		ws['B' + i[1:]].font = Font(size=9)
		col_num = int(i[1:])
		row_num = 'B' + i[1:]
		time_cell = 'B' + i[1:] + ':' + 'C' + str(col_num)
		Create_Post(time_cell,int(row_num[1:]) ,2, WORK_TIMES[time_of_the_shift_var])

		#wb.save(File_Name)

	#If ever an issue, replace every 'B' with Next_Col and uncomment the first 2 lines inside this for loop

TimeStamps(0, Shift_Rows, Color_For_Regular_Shift)
TimeStamps(1, Closer_11_Row, Color_For_11_Shift)
TimeStamps(2, Breaker_Rows, Color_For_Breaker_Shift)
TimeStamps(3, Lead_Rows, Color_For_Lead_Shift)


#------------Blacking Out Every Box and Setting the Brief

# Create a Font object with black color
Dark = PatternFill(patternType = 'solid', fgColor = '000000') 
Blue = PatternFill(patternType = 'solid', fgColor = FLOAT)
Brief = PatternFill(patternType = 'solid', fgColor = BRIEF)

# Loop through the specified range of rows in the column and set the font color
for i in range(len(Last_Row_Cordinates)):
	if ws['D' + Last_Row_Cordinates[i][1:]].value == None:
		ws['D' + Last_Row_Cordinates[i][1:]].fill = Dark
# Loop through the rows where the list will be and adding 'Lead' text
for i in range(len(Lead_Rows)):
	if ws['D' + Lead_Rows[i][1:]].value == None:
		ws['D' + Lead_Rows[i][1:]].value = WORK_TIMES[3]		#EVERYTHING_ELSE[1]           
		ws['D' + Lead_Rows[i][1:]].fill = Blue 					#Brief          
		ws['D' + Lead_Rows[i][1:]].border = border
		ws['D' + Lead_Rows[i][1:]].alignment = Center_Text #changed

# Loop through the rows in column E and add the BRIEF'ing text
for i in range(len(Last_Row_Cordinates)):
	if ws['E' + Last_Row_Cordinates[i][1:]].value == None:
		ws['E' + Last_Row_Cordinates[i][1:]].value = EVERYTHING_ELSE[1]
		ws['E' + Last_Row_Cordinates[i][1:]].fill = Brief
		ws['E' + Last_Row_Cordinates[i][1:]].border = border
		ws['E' + Last_Row_Cordinates[i][1:]].alignment = Center_Text

for i in range(len(Lead_Rows)):
	if ws['E' + Lead_Rows[i][1:]].value == None:
		ws['E' + Lead_Rows[i][1:]].value = EVERYTHING_ELSE[1]
		ws['E' + Lead_Rows[i][1:]].fill = Brief
		ws['E' + Lead_Rows[i][1:]].border = border
		ws['E' + Lead_Rows[i][1:]].alignment = Center_Text


'''
Minor Mistake made in the code, B1 doesnt have enough rows sometimes because OB3 has way more than needed.
I will write code below that counts every 4-12 and 4-11 thats on B1 - OB1, then one that counts 0OB1 - OB3
First it will look for 4-11's, if any are left over we will add them to B1 and delete them from OB3'''





'''
Purpose: We need to count every none breaker row on B1 to make sure it has at least 11 Regular Shift rows in order for that floor to operate.
Turn the Code below into functions, also add if statement below them that if cell_count is less than 11 then do this

'''
#--------------------------Checking to see if B1, OB1, and OB2 have enough rows for a shift, in order for the floor to operate.
#SOLUTION: Make if statements at the start of this large amount of code that if Leads and Both Closers equal over a certain amount, then make the floors 4 - 6 etc


def Regular_Shift_Row_Counter(start_ro, end_ro, count):
	start_row = start_ro  # Adjust the starting row as needed

	end_row = end_ro  # Adjust the ending row as needed 

	# Initialize a variable to count the cells
	count = 0

	# Loop through the rows within the specified range and count the cells in the column
	for row_number in range(start_row, end_row + 1):
		#print ('Starting Row: ', start_row)
		cell = ws['B' + str(row_number)]

		if cell.value is not None and '4:' in cell.value:  # Check if the cell is not empty
			count += 1
		elif cell.value == WORK_TIMES[3]:
			count += 1
		#else:
		#	print ('Breaker Row')
		#print ('End Count: ', count)
		#print ('Inside the Cell', cell.value)
	return (count)

cell_count = 0
cell_count = Regular_Shift_Row_Counter(5, int(Locate('OB1')[0][1:]), cell_count)


# Print the cell count
print("B1 Regular Shift Rows: ", cell_count)


extra_cell_count = 0
extra_cell_count = Regular_Shift_Row_Counter(int(Locate('OB3')[0][1:]), ws.max_row, extra_cell_count)
print("OB3 Extra Regular Shift Rows: ", extra_cell_count)


# Custom sorting function to extract the numeric part
def custom_sort_key(item):
    # Use a regular expression to find and extract the numeric part
    import re
    match = re.search(r'\d+', item)
    if match:
        return int(match.group())  # Convert the matched digits to an integer
    else:
        return item  # If no numeric part found, sort by the full string

#This variable will determine how many Ambassadors/ Shifts to have on OB1 and OB2
#Should prob edit this later, like raise the number from 22. Should def raise all the min_ambassadors numbers at some point
#ERROR / error/ fix/ FIX: we need to link all these numbers to a set var with a number/integer 

B1_Full_Shift_Rows = 0 
#This var is brand new and unrelated to the comment above ^. That comment above is meant for the mini code right below.
#This var will literally DETERMINE how many B1 Full Shift rows are on B1, thats INCLUDING Leads as well, not including 6 - 10's.

'''
if Ambassador > 22:
	B1_Full_Shift_Rows = 15
	Min_Ambassador = 9 # APRES ; OB3 Post; Was Originally 8 Before OB3 Momento was Added, Make it 9 for Better OB1 and OB2 Results. Currently only 8 to Force a bug so I can fix...
elif Ambassador < 20:
	B1_Full_Shift_Rows = 13 #Most likely needs to be played around with, can prob increase or decrease later
	Min_Ambassador = 4
elif Ambassador >= 20:
	B1_Full_Shift_Rows = 12 #Most likely needs to be played around with, can prob increase or decrease later
	Min_Ambassador = 5
'''


# --- DYNAMIC STAFFING DISTRIBUTION (ZERO-WASTE) ---
# This ensures (B1 + OB1 + OB2) == Total Ambassadors. 
# Physical OB3 rows will be 0, as those posts are now in the OB2/OB1 lists.

# 1. Set B1 Priority (The Foundation)
if Ambassador >= 33:
    B1_Full_Shift_Rows = 18
elif Ambassador >= 28:
    B1_Full_Shift_Rows = 14
else:
    B1_Full_Shift_Rows = 13

# 2. Calculate the Remaining Staff for Upstairs
Upstairs_Pool = Ambassador - B1_Full_Shift_Rows

# 3. Split the Pool between OB1 and OB2
# We give OB2 the 'extra' person if the number is odd, 
# because OB2 handles the OB3 overflow posts.
Min_Ambassador_OB1 = Upstairs_Pool // 2 
Min_Ambassador_OB2 = Upstairs_Pool - Min_Ambassador_OB1

print(f"Logic Check: B1({B1_Full_Shift_Rows}) + OB1({Min_Ambassador_OB1}) + OB2({Min_Ambassador_OB2}) = {B1_Full_Shift_Rows + Min_Ambassador_OB1 + Min_Ambassador_OB2}")


Move_To_B1 = int(Create_Balance[2][1:])
Replacement_Rows = (B1_Full_Shift_Rows - cell_count) #Basically B1 Cant have less than 10 full shifts on B1, if so it will pull from OB3
#The if statement will be created here. if replacement not equal to 0 or less, create more rows and color it in etc
if Replacement_Rows > 0 and extra_cell_count > 0 and Replacement_Rows <= extra_cell_count:
	ws.insert_rows(Move_To_B1, Replacement_Rows)


	for i in Create_Balance[-Replacement_Rows:]:
		if ws[i].value is None or ws[i].value == 'OB3' or ws[i].value == 'OB2' or '6pm' in ws['B' + i[1:]].value or '11' in ws['B' + i[1:]].value or 'LEAD' == ws['B' + i[1:]].value:
			Replacement_Rows -= 1
		else:
			print ('Deleting: ', ws['B' + i[1:]].value)
			ws.delete_rows(int(i[1:]))
			del Create_Balance[-Replacement_Rows:]
			del Shift_Rows[-Replacement_Rows:]




	start_row = int(Locate('B1')[0][1:])
	end_row = int(Locate('B1')[1][1:])
	for row_number in range(start_row, end_row + 1):
		if ws['A' + str(row_number)].value == None:
			Create_Balance.append('A' + str(row_number))
			Shift_Rows.append('A' + str(row_number))


	# Create a new sorted list using the custom sorting key to arrange every item in proper order
	Create_Balance = sorted(Create_Balance, key=custom_sort_key)
	Shift_Rows = sorted(Shift_Rows, key=custom_sort_key)


'''
cell_count_2 = 0
cell_count_2 = Regular_Shift_Row_Counter(int(Locate('OB1')[0][1:]), int(Locate('OB2')[0][1:]), cell_count_2)
# Print the cell count
print("OB1 Regular Shift Rows: ", cell_count_2)

extra_cell_count = 0
extra_cell_count = Regular_Shift_Row_Counter(int(Locate('OB3')[0][1:]), ws.max_row, extra_cell_count)
print("OB3 Extra Regular Shift Rows: ", extra_cell_count)



Move_To_OB1 = (int(Locate('OB1')[0][1:]) + 2)    #Move_To var's have to be an integer to represent the row number we will start with
Replacement_Rows = (Min_Ambassador_OB1 - cell_count_2)            #4 - 6 is the minimum amount required for this floor to flow properly
#print (Move_To_OB1)
if Replacement_Rows > 0 and extra_cell_count > 0 and Replacement_Rows <= extra_cell_count: #May need to add another and statement to make sure Replacement_Rows are less than extra_cell_count
	ws.insert_rows(Move_To_OB1, Replacement_Rows)
	#Ended up deleting OB3 By Accident so we created an if statement below to subtract Replacement_Row by 1

	print ('Value: ', ws[Create_Balance[-Replacement_Rows:][0]].value)
	print ('Replacement Row: ', Replacement_Rows)


	
	for i in Create_Balance[-Replacement_Rows:]:
		#print ('Test', ws[i].value)
		if ws[i].value is None or ws[i].value == 'OB3' or ws[i].value == 'OB2' or '6pm' in ws['B' + i[1:]].value or '11' in ws['B' + i[1:]].value or 'LEAD' == ws['B' + i[1:]].value:
			Replacement_Rows -= 1
		else:
			print ('Deleting: ', ws['B' + i[1:]].value)
			ws.delete_rows(int(i[1:]))
			del Create_Balance[-Replacement_Rows:]
			del Shift_Rows[-Replacement_Rows:]




	start_row = int(Locate('OB1')[0][1:])
	end_row = int(Locate('OB2')[0][1:])
	for row_number in range(start_row, end_row + 1):
		if ws['A' + str(row_number)].value == None:
			Create_Balance.append('A' + str(row_number))
			Shift_Rows.append('A' + str(row_number))


	# Create a new sorted list using the custom sorting key to arrange every item in proper order
	Create_Balance = sorted(Create_Balance, key=custom_sort_key)
	Shift_Rows = sorted(Shift_Rows, key=custom_sort_key)

print ('')


cell_count_3 = 0
cell_count_3 = Regular_Shift_Row_Counter(int(Locate('OB2')[0][1:]), int(Locate('OB3')[0][1:]), cell_count_3)
# Print the cell count
print("OB2 Regular Shift Rows: ", cell_count_3)


extra_cell_count = 0
extra_cell_count = Regular_Shift_Row_Counter(int(Locate('OB3')[0][1:]), 50, extra_cell_count) #Before 50 was int(Create_Balance[-1][1:])
print("OB3 Extra Regular Shift Rows: ", extra_cell_count)

print ('Cell Count 3: ', cell_count_3)

Move_To_OB2 = (int(Locate('OB2')[0][1:]) + 2)    #Move_To var's have to be an integer to represent the row number we will start with
Replacement_Rows = (Min_Ambassador_OB2 - cell_count_3)            #6 is the minimum amount required for this floor to flow properly
#print (Move_To_OB2)
if Replacement_Rows > 0 and extra_cell_count > 0 and Replacement_Rows <= extra_cell_count: #May need to add another and statement to make sure Replacement_Rows are less than extra_cell_count
	ws.insert_rows(Move_To_OB2, Replacement_Rows)
	print (Create_Balance[-Replacement_Rows:])
	print ('Replacement Row: ', Replacement_Rows)
	#SOLUTION: Solution for these if statements below is to add a function that counts


	for i in Create_Balance[-Replacement_Rows:]:
		#print ('Error: ', ws[i].value)
		#print ('Error: ', ws['B' + i[1:]].value)
		#print (i)

		if ws[i].value is None or ws[i].value == 'OB3' or ws[i].value == 'OB2' or '6pm' in ws['B' + i[1:]].value or '11' in ws['B' + i[1:]].value or 'LEAD' == ws['B' + i[1:]].value:
			Replacement_Rows -= 1
		else:
			print ('Deleting: ', ws['B' + i[1:]].value)
			ws.delete_rows(int(i[1:]))
			del Create_Balance[-Replacement_Rows:]
			del Shift_Rows[-Replacement_Rows:]





	#Accounts for any changes between these 2 cordinates / floors (OB2 - OB3)
	start_row = int(Locate('OB2')[0][1:])
	end_row = int(Locate('OB3')[0][1:])
	for row_number in range(start_row, end_row + 1):
		if ws['A' + str(row_number)].value == None:
			Create_Balance.append('A' + str(row_number))
			Shift_Rows.append('A' + str(row_number))


	# Create a new sorted list using the custom sorting key to arrange every item in proper order
	Create_Balance = sorted(Create_Balance, key=custom_sort_key)
	Shift_Rows = sorted(Shift_Rows, key=custom_sort_key)

'''

# --- OB1 PHYSICAL BALANCE --- AI Assisted Version
cell_count_2 = Regular_Shift_Row_Counter(int(Locate('OB1')[0][1:]), int(Locate('OB2')[0][1:]), 0)
Move_To_OB1 = (int(Locate('OB1')[0][1:]) + 2)
Replacement_Rows = (Min_Ambassador_OB1 - cell_count_2)

if Replacement_Rows > 0:
    ws.insert_rows(Move_To_OB1, Replacement_Rows)
    # This pulls rows from the very bottom (OB3) and moves them to OB1
    for i in range(Replacement_Rows):
        ws.delete_rows(ws.max_row) 

# --- OB2 PHYSICAL BALANCE ---
cell_count_3 = Regular_Shift_Row_Counter(int(Locate('OB2')[0][1:]), int(Locate('OB3')[0][1:]), 0)
Move_To_OB2 = (int(Locate('OB2')[0][1:]) + 2)

if len(Lead_Rows) >= 5:
	Replacement_Rows = ((Min_Ambassador_OB2 - 1) - cell_count_3)
	#Basically tells the program to not delete the last row on OB3 if its a LEAD. This Lead count has to be accurate
else:
	Replacement_Rows = (Min_Ambassador_OB2 - cell_count_3)

if Replacement_Rows > 0:
    ws.insert_rows(Move_To_OB2, Replacement_Rows)
    # This empties the rest of OB3
    for i in range(Replacement_Rows):
        ws.delete_rows(ws.max_row)





#Used to fill the names in for any new rows that may have been created
Shift_Rows = [] #This list will be used for all the rows where a rotation line will be for a shift
for i in range(1, Names):
	Result = ws['A' + str(i)].value
	if Result == None:
		Shift_Rows.append('A' + str(i))
		for cell in ws.merged_cells.ranges: #Gives us a list of Merged Cells
			if ('A' + str(i)) in cell:      #Unmerges every cell where employee names would be. It would merge over with the next cell for an odd reason
				ws.merged_cells.remove('A' + str(i) + ':C' + str(i))

		ws['A' + str(i)].value = 'Name'     #Prints the names of employees that will work during this shift
		ws['A' + str(i)].alignment = Center_Text #Centers the our names inside the cell
		ws['A' + str(i)].border = border
		#wb.save(File_Name)
'''
#Used to fill the shift times for any new rows but also, any times rows are created, 
#previous rows tend to unmerge cells that were already merged cells. Could be random also.

#for Loop below is used to remerged any pre-existing cells
'''

#print ('Updated All Shifts and Rows: ', Create_Balance)
#print ('Updated Regular Shift Rows: ', Shift_Rows)
#print ('Updated Lead Rows: ', Lead_Rows)
#print ('Updated Breaker Rows: ', Breaker_Rows)
print ('Updated Close Until 11 Rows: ', Closer_11_Row)
#print ('Error is in this Area')
#print ('Regular Closers: ', Closers)
#print ('Leads: ', Leads)
#print ('Breakers: ', Breakers)
#print ('Close_Untill_11: ', Closers_Untill_11)

TimeStamps(0, Shift_Rows, Color_For_Regular_Shift)


start_row = 5 
end_row = int(Create_Balance[-1][1:])

for i in range(start_row, end_row + 1):
	ws.merge_cells('B' + str(i) + ':C' + str(i))

#If Shift times unmerge ever again, just uncomment this



# Loop through the specified range of rows in the column and set the font color
for i in range(len(Shift_Rows)):
	if ws['D' + Shift_Rows[i][1:]].value == None:
		ws['D' + Shift_Rows[i][1:]].fill = Dark

# Loop through the rows in column E and add the BRIEF'ing text
for i in range(len(Shift_Rows)):
	if ws['E' + Shift_Rows[i][1:]].value == None:
		ws['E' + Shift_Rows[i][1:]].value = EVERYTHING_ELSE[1]
		ws['E' + Shift_Rows[i][1:]].fill = Brief
		ws['E' + Shift_Rows[i][1:]].border = border
		ws['E' + Shift_Rows[i][1:]].alignment = Center_Text





#This function will be used to recount every shift lift, to recreate accuracy inside all shift list
'''
Old Code, can Delete later if it works

def Recount(the_shift_list, char_to_look_for):
	start_row = 5
	end_row = 50
	column_letter = 'B'  # Replace 'A' with the desired column letter
	for row_number in range(start_row, end_row + 1):
		cell_value = ws[column_letter + str(row_number)].value
		
		# Check if the cell value contains the target character '12' Which represents regular closing shifts
		if cell_value and char_to_look_for in cell_value:
			# If yes, add the cell coordinates to the list
			the_shift_list.append((column_letter + str(row_number)))

	return (the_shift_list)
'''


def Recount(the_shift_list, char_to_look_for):
    start_row = 5
    end_row = ws.max_row # Dynamically check every row created
    column_letter = 'B'
    for row_number in range(start_row, end_row + 1):
        cell_value = ws[column_letter + str(row_number)].value
        if cell_value and char_to_look_for in str(cell_value):
            the_shift_list.append((column_letter + str(row_number)))
    return (the_shift_list)


#Reorganizing the list full of cordinates with accurate cordinates by looping through the B coulmn and searching for a specific character
# --- MASTER SHIFT SYNC BLOCK ---
# This wipes the old inaccurate lists and re-scans the Excel sheet
Shift_Rows = []
Shift_Rows = Recount(Shift_Rows, '12') # Scans for 12:00 and 12:30/3:30

Breaker_Rows = []
Breaker_Rows = Recount(Breaker_Rows, '6') # Scans for 6:00 Breakers

Closer_11_Row = []
Closer_11_Row = Recount(Closer_11_Row, '11') # Scans for 11:00 Closers

Lead_Rows = []
Lead_Rows = Recount(Lead_Rows, 'LEAD') # Scans for LEADS

# Rebuild the 'Create_Balance' Master List
# We use natsorted to ensure the coordinates are in order (A5, A6, A10, A11...)
Create_Balance = natsorted(Shift_Rows + Breaker_Rows + Closer_11_Row + Lead_Rows)


'''
# Double Tap: Ensure Lead rows have the correct labels in Column D
for i in Lead_Rows:
    ws['D' + i[1:]].value = WORK_TIMES[3] # Sets it to 'LEAD'
    ws['D' + i[1:]].fill = Blue
    ws['D' + i[1:]].border = border
    ws['D' + i[1:]].alignment = Center_Text


#Will be used to update the row with Leads so we can then
num_L = Leads_2 - len(Lead_Rows)
#print (num_L)
if num_L != 0:
	while num_L != 0:
		random.shuffle(Shift_Rows)
		Lead_Rows.append(Shift_Rows[num_L])
		Shift_Rows.pop(num_L)
		num_L -= 1

#Putting them back in order
Shift_Rows = sorted(Shift_Rows, key=custom_sort_key)
Lead_Rows = sorted(Lead_Rows, key=custom_sort_key)
Create_Balance = sorted(Create_Balance, key=custom_sort_key)

#Rewriting the Lead Rows since it has been updated
for i in Lead_Rows:
	if ws[i].value == WORK_TIMES[0]:
		ws[i].value = WORK_TIMES[3]
		ws[i].fill = Color_For_Lead_Shift
		ws[i].border = border
		ws[i].alignment = Center_Text
'''

# Loop through the rows where the list will be and adding 'Lead' text
for i in range(len(Lead_Rows)):
	ws['D' + Lead_Rows[i][1:]].value = EVERYTHING_ELSE[1]
	ws['D' + Lead_Rows[i][1:]].fill = Brief
	ws['D' + Lead_Rows[i][1:]].border = border
	ws['D' + Lead_Rows[i][1:]].alignment = Center_Text

#Was Here Before

# Remerges all time slots/ shift times in Column B
start_row = 5 
end_row = int(Create_Balance[-1][1:])

for i in range(start_row, end_row + 1):
	ws.merge_cells('B' + str(i) + ':C' + str(i))


# An Extra Row is left in some insances; Clean Up!
if len(Create_Balance) - Total > 0:
	Row_To_Delete = (Create_Balance[Total:])
	for i in (Row_To_Delete):
		#print (i)
		ws.delete_rows(int(i[1:]))
		ws.delete_rows(int(Row_To_Delete[0][1:])) #Double Tap
	Shift_Rows = [x for x in Shift_Rows if x not in Row_To_Delete]
	Create_Balance = [x for x in Create_Balance if x not in Row_To_Delete]


#Adding in the Black Boxes for Breaker Shifts and Briefing Column/ Also Black Boxes for 4-11
for i in Breaker_Rows:
	#ws.merge_cells('D' + i[1:] + ':H' + i[1:]) #Merging all cells for breaker shifts until 6PM
	Letter = 3             #Starts at the letter D, which is where 11PM starts
	for color in range(5): #Instead of merging the cells we just painted them all black to help make things easier as we move forward
		ws[ALPHABET[Letter] + i[1:]].fill = Dark
		ws[ALPHABET[Letter] + i[1:]].value = 'SKIP'
		Letter += 1
	wb.save(File_Name)

	#Writing Briefing in the code below for the Column I which is the start of breaker shifts
	ws['I' + i[1:]].value = EVERYTHING_ELSE[1] #Adding 'BRIEF' on the start of briefing shifts
	ws['I' + i[1:]].fill = Brief
	ws['I' + i[1:]].border = border
	ws['I' + i[1:]].alignment = Center_Text

	#Coloring the end of the shift black below
	#ws.merge_cells('R' + i[1:] + ':V' + i[1:])
	Letter = 17            #Starts at the letter R, which is where 10:30PM starts
	for color in range(5): #Instead of merging the cells we just painted them all black to help make things easier as we move forward
		ws[ALPHABET[Letter] + i[1:]].fill = Dark
		ws[ALPHABET[Letter] + i[1:]].value = 'SKIP'
		Letter += 1
	wb.save(File_Name)


for i in Closer_11_Row:
	Letter = 18            #Starts at the letter S, which is where 11PM starts
	for color in range(4): #Instead of merging the cells we just painted them all black to help make things easier as we move forward
	#ws.merge_cells('S' + i[1:] + ':V' + i[1:])
		ws[ALPHABET[Letter] + i[1:]].fill = Dark
		ws[ALPHABET[Letter] + i[1:]].value = 'SKIP'
		Letter += 1
		#wb.save(File_Name)

#Adding Dark Boxes and the word "SKIP" to the end of 4-12 Shifts
for i in Shift_Rows:
	Letter = 20           #Starts at the letter U, which is where 12PM starts
	for color in range(2):
		ws[ALPHABET[Letter] + i[1:]].fill = Dark
		ws[ALPHABET[Letter] + i[1:]].value = 'SKIP'
		Letter += 1
		#wb.save(File_Name)

#Adding Check List at the end of Lead Rows
for i in Lead_Rows:
	FLOAT_CELL = PatternFill(patternType = 'solid', fgColor = FLOAT)
	ws['U' + i[1:]].fill = FLOAT_CELL
	Cell = 'U' + i[1:] + ':' + 'V' + i[1:]
	index = ALPHABET.index('U') + 1
	Create_Post(Cell, int(i[1:]), index, EVERYTHING_ELSE[2])



#Newly Added Coded for Full Timers and Leads Only. It will re-write and add post for them at 4:00PM Instead of a BRIEFING cell
for a in Lead_Rows:
	ws['E' + a[1:]].value = EVERYTHING_ELSE[0]
	ws['E' + a[1:]].fill = FLOAT_CELL
	ws['E' + a[1:]].border = border
	ws['E' + a[1:]].alignment = Center_Text




print (B1_Full_Shift_Rows)
print ('Done testing')
wb.save(File_Name)
#sleep(99999)



print ('')
print ('Updated All Shifts and Rows: ', Create_Balance)
print (len(Create_Balance))

#All Shifts that end at 12 are Shift - Rows
print ('Updated Full Timers and Closer Rows: ', Shift_Rows)
print (len(Shift_Rows)) 
print ('Updated Lead Rows: ', Lead_Rows)
print (len(Lead_Rows))
print (Leads)
print ('Updated Breaker Rows: ', Breaker_Rows)
print (len(Breaker_Rows))
print ('Updated Close Until 11 Rows: ', Closer_11_Row)
print (len (Closer_11_Row))
	



print ('Total Shifts: ', Total)
print ('Total Cord Shifts: ', len(Create_Balance))

# Change the old print line to this:
print ('MIN GEA OB1: ', Min_Ambassador_OB1)
print ('MIN GEA OB2: ', Min_Ambassador_OB2)
print ('Ambassadors in Total: ', Ambassador)
print ('IDEA: Have it do Column by Column and do Floor by Floor first')
print ('Anything that cant fit, place into a seperate list attached to their Column Letter so we can hand them out later for anyone on OB3 or on other floors that have replaceable post')


print ('')
print (Ambassador)
print ('End of Row Creation. Line 1113')
wb.save(File_Name)
sleep(0)

#---------------------------------- Below is the Code to add Post in

Freight_Times = [] #Make it self adjustable
Break_Times = ['J{eat_time}:K{eat_time}','J{eat_time}', 'K{eat_time}', 'L{eat_time}:M{eat_time}', 'L{eat_time}', 'M{eat_time}', 'N{eat_time}:O{eat_time}', 'N{eat_time}', 'N{eat_time}' ] #K - N is 7:00PM - 8:30PM
#Captured the empty cells starting at letter F until OB1. That would be
#everyone that needs a break. Remember anything filled in dark isnt an empty cell the word 'SKIP' is written in them with black font
#eat_time will be replaced with the row number for the break time

#Dictionary attaching time intervals with the letter column that best associate with that time.

#Below is Before:
Letter_Times = {
	'3:30': 'D',
	'4:00': 'E',
	'4:30': 'F', 
	'5:00': 'G',
	'5:30': 'H',
	'6:00': 'I',
	'6:30': 'J',
	'7:00': 'K',
	'7:30': 'L',
	'8:00': 'M',
	'8:30': 'N', 
	'9:00': 'O',
	'9:30': 'P',
	'10:00': 'Q',
	'10:30': 'R',
	'11:00': 'S',
	'11:30': 'T',
	'12:00': 'U',
	'12:30': 'V'

}

#Should be the updated Version when the time is right
'''
Letter_Times = {
	'3:00': 'D',
	'3:30': 'E',
	'4:00': 'F', 
	'4:30': 'G',
	'5:00': 'H',
	'5:30': 'I',
	'6:00': 'J',
	'6:30': 'K',
	'7:00': 'L',
	'7:30': 'M',
	'8:00': 'N', 
	'8:30': 'O',
	'9:00': 'P',
	'9:30': 'Q',
	'10:00': 'R',
	'10:30': 'S',
	'11:00': 'T',
	'11:30': 'U',
	'12:00': 'V'
}'''

Freight_Start = Letter_Times['7:30'] #This will be the column Freight officially starts. It should be purplish and have Freight Text
Freight_End = Letter_Times['9:00']
Freight_Break_Check = Letter_Times['7:00']   #This will be the column after Freight is officially over. Should not be apart of Fright cell. Used to check and make sure this person's break aligns with Freight



#Recently Moved Here
'''
YES DELETE LATER.....
def Temporary_Value(Cord_In_Merged_Cell, the_cell_value):

		# Specify the cell you want to check
		target_cell = ws[Cord_In_Merged_Cell] #-----> Will be used for the parameter of the function

		if the_cell_value is None:

			# Check if the cell is part of a merged cell
			if target_cell.coordinate in ws.merged_cells:
			    # If it is, find the merged cell range
			    merged_cell_range = None
			    for merged_range in ws.merged_cells.ranges:
			        if target_cell.coordinate in merged_range:
			        	merged_cell_range = merged_range
			        	Father_Value = str(merged_cell_range)
			        	#print (Father_Value)
			        	Father_Value = Father_Value[0] + Cord_In_Merged_Cell[1:]
			        	#print (Father_Value)
			        	Father_Value = ws[Father_Value].value
			        	#print (Father_Value)
			        	return (Father_Value)
			        	#sleep(999)
			        	break

			    #if merged_cell_range:
			    	#print(f"The cell {target_cell.coordinate} is part of the merged cell range {merged_cell_range}.... should return the value of the merged cell")
			    #else:
			        #print(f"The cell {target_cell.coordinate} is part of a merged cell, but the range couldn't be determined.")
			#else:
				
			    #print(f"The cell {target_cell.coordinate} is not part of a merged cell.")

			    #print ('In order for this future function to be complete, have it return/print the value of the merged cell its originally apart of')
			    
		elif the_cell_value is not None:
			return (the_cell_value)		     #Keep this Code Alive
'''




'''
The FloorCellCount function below will partially be used to count how many Non Breaker Shifts are on B1, but it keeps count in general
But also will help us determine later on if we need a post covered on a different floor
- I had to fill in all the black boxes with the word "SKIP" to make it easier to count, the font should be black so no one can see it
'''



B1_Floor_Rows = []      #Only a list for rows that go from 4 - 11 or 4 - 12 on B1
#B1_Overall_Floor_Rows = [] #A list made up of Floor Rows including Breaker Shifts
def FloorCellCount(column_letter, floorlist, end_row_floor, start_row_floor):
	global Floor_Column_Count, Floor_OB1, Floor_OB2, Leads
	Floor_OB1 = [] #Will be used for OB1
	Floor_OB2 = [] #Will be used for OB2
	ignore_list = [EVERYTHING_ELSE[3], 'FREIGHT', 'PREP'] #Hasnt been used yet

	Floor_Column_Count = 0  #Will be used to determine how many post we will need in this column..... HERE
	start_row = start_row_floor #For B1 its the number 5

	#end_row = int(end_row_floor) #int(Locate(end_row_floor)[0][1:]) #Will have to make this a parameter - For B1 its 'OB1'
	
	if end_row_floor.isdigit():
        # If it's a number, use it directly
		end_row = int(end_row_floor)
	else:
        # If it's a word, search for it
		end_row = int(Locate(end_row_floor)[0][1:])





	for i in range(start_row, end_row + 1):
		cell = ws[column_letter + str(i)]
		
		''' Can be used as the solution later
		# 1. Iterate through the merged cell ranges on the worksheet
		is_merged = False
		cell_to_check = column_letter + str(i)
		for merged_range in ws.merged_cells.ranges:
		    # 2. Check if the cell's coordinates are within the current merged range
		    if cell_to_check in merged_range:
		        is_merged = True
		        break  # Found it, no need to check other ranges

		if is_merged:
		    print(f"Cell {cell_to_check} is part of a merged cell range.")
		else:
		    print(f"Cell {cell_to_check} is NOT part of a merged cell range.")

		# You can also get the merged range it belongs to:
		if is_merged:
		    print(f"It belongs to the range: {merged_range.coord}")
		'''
		




		if (cell.value) is None and (column_letter + str(i)) not in ws.merged_cells and start_row < int(Locate('OB1')[0][1:]):
			Floor_Column_Count += 1
			floorlist.append(column_letter + str(i))

			#Previous If Statement: (cell.value) is None and (column_letter + str(i)) not in ws.merged_cells and start_row > int(Locate('B1')[0][1:]):
		elif (cell.value) is None and (column_letter + str(i)) not in ws.merged_cells and start_row > int(Locate('B1')[0][1:]):
			Floor_Column_Count += 1
			floorlist.append(column_letter + str(i))
			if i > int(Locate('OB2')[0][1:]): #Was here before:  and i < int(Locate('OB3')[0][1:]) | Any rows thats on OB2 will be accounted for and added to the OB2 list
				Floor_OB2.append(column_letter + str(i))
			elif i > int(Locate('OB1')[0][1:]) and i < int(Locate('OB2')[0][1:]): # Any rows thats on OB1 will be accounted for and added to the OB1 list
				Floor_OB1.append(column_letter + str(i))

		#Had to add this new statement. Made a mistake and created FREIGHT and BREAKS while all B1 Post were in the middle of being created.
		elif (cell.value) is None and (column_letter + str(i)) in ws.merged_cells or (cell.value) == 'PREP' or (cell.value) == 'BREAK' or 'FREIGHT' in (cell.value): #Should only ever be FREIGHT, BREAK OR PREP. May have to make this specific to a floor later at some point. Use the above elif statements as a reference
			floorlist.append(column_letter + str(i))
			Floor_Column_Count += 1





FloorCellCount('F', B1_Floor_Rows, 'OB1', 5)






#--------------------------------------------------------------------------------------------------------
#Old Code Below and In Between these Lines





def B1_Rotations_Creation(column_letter, col_num):
	Post_Tier = 0 #Will be used to help loop through the tier list and pull out the post string
	global B1_Tier_1, B1_Tier_2, Floor_Column_Count
	B1_Length = len(B1_Tier_1)
	B2_Length = len(B1_Tier_2)
	B1_Tier_1 = ['RISE TABLET','RISE LINE', "HELLO 1", 'SHOES 1', 'CELEBRATE', 'PLAZA VANDY', 'PLAZA MAD', 'HELLO 2', 'QUEUE', 'LAUNCH**']


	if Floor_Column_Count > len(B1_Tier_1) and Floor_Column_Count <= len(Mix_Of_Both_B1_Tiers) and col_num < 18:
		print ('Starting 1st Statement')
		random.shuffle(Mix_Of_Both_B1_Tiers)
		#The Last 2 of the new larger B1 List are B1 Tier 2 Now
		for i in B1_Floor_Rows:
			Cell = i + ':' + column_letter + i[1:]
			Create_Post(Cell, int(i[1:]), col_num, Mix_Of_Both_B1_Tiers[Post_Tier])
			B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
			ws[i].fill = B1_Cell
			print ('Post Added: ', Mix_Of_Both_B1_Tiers[Post_Tier])
			print (Post_Tier)
			Post_Tier += 1
		print ('Done..................-')
		#Just resetting the B1 Tiers to what they used to be
	elif Floor_Column_Count <= len(B1_Tier_1) and col_num < 18:
		print ('Starting 2nd Statement')
		random.shuffle(B1_Tier_1)
		for i in B1_Floor_Rows:
			Cell = i + ':' + column_letter + i[1:]
			Create_Post(Cell, int(i[1:]), col_num, B1_Tier_1[Post_Tier])
			B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
			ws[i].fill = B1_Cell
			print ('Post Added: ', B1_Tier_1[Post_Tier])
			print (Post_Tier)
			Post_Tier += 1
		print ('Done..................')
	elif Floor_Column_Count > len(Mix_Of_Both_B1_Tiers) and col_num < 18:
		print ('Starting 3rd Statement')
		leads_found = 0 #Will be used to subtract from the final Floor_Column_Count
		for i in B1_Floor_Rows: #Will give leads float positions
			for o in Lead_Rows:
				if i[1:] == o[1:]:
					leads_found += 1
					FLOAT_CELL = PatternFill(patternType = 'solid', fgColor = FLOAT)
					ws[i].fill = FLOAT_CELL
					index = ALPHABET.index(i[0]) + 1
					Cell = i + ':' + ALPHABET[index] + i[1:]
					Create_Post(Cell, int(i[1:]), index, EVERYTHING_ELSE[0])
					B1_Floor_Rows.remove(i)
					Floor_Column_Count -= 1
					print ('Lead Found: ', i)
		if len(Mix_Of_Both_B1_Tiers) <= len(B1_Floor_Rows):
			floats_to_add = len(B1_Floor_Rows) - len(Mix_Of_Both_B1_Tiers)
			Mix_Of_Both_B1_Tiers.extend(FLOAT_LIST[:floats_to_add])
			random.shuffle(Mix_Of_Both_B1_Tiers)
			#The Last 2 of the new larger B1 List are B1 Tier 2 Now
			for i in B1_Floor_Rows:
				Cell = i + ':' + column_letter + i[1:]
				Create_Post(Cell, int(i[1:]), col_num, Mix_Of_Both_B1_Tiers[Post_Tier])
				B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
				ws[i].fill = B1_Cell
				print ('Post Added: ', Mix_Of_Both_B1_Tiers[Post_Tier])
				print (Post_Tier)
				Post_Tier += 1
		wb.save(File_Name)
		print ('Mix Count: ', len(Mix_Of_Both_B1_Tiers), ', B1 Floor Count: ',  len(B1_Floor_Rows))
		print ('Done..................2')
		##sleep(2000)
	elif column_letter == 'S' and Floor_Column_Count <= 12 and col_num == 18:
		print ('Starting Closing Post')
		Total_Closers = 0
		#num_Closers = 0
		Total_Leads = 0
		#num_Leads = 0
		for i in B1_Floor_Rows:
			for o in Shift_Rows:
				if i[1:] == o[1:]:
					Total_Closers += 1 #Needed to know how many 4 - 12's Work first

		print ('Total Closers: ', Total_Closers)

		if Total_Closers > 2:
			while ('RISE TABLET' not in B1_Tier_1[:Total_Closers - 1]) or ('RISE LINE' not in B1_Tier_1[:Total_Closers - 1]): 
			#Makes no sense, but 'or' works better than 'and' and does what and is supposed to do in if statements
				random.shuffle(B1_Tier_1) #Also we mainly only need to give RISE post to 4 - 12'ers. Celebrate can be given to 11'ers. May change later. 
				print (B1_Tier_1)
				#sleep(0)
				#It loops tho until these 3 Post are in the front of this list

		elif Total_Closers == 2:
			print ('Total Closers on B1 are 2... initiating')
			# Specify the two strings you want to move to the front
			string1 = "RISE LINE"
			string2 = "RISE TABLET"

			# Check if the specified strings are in the list
			if string1 in B1_Tier_1 and string2 in B1_Tier_1:
    			# Remove the strings from their original positions
				B1_Tier_1.remove(string1)
				B1_Tier_1.remove(string2)

    			# Insert the strings at the front of the list
				B1_Tier_1 = [string1, string2] + B1_Tier_1
		elif Total_Closers == 1: #If theirs ever just 1 Closer on B1
			print ('Total Closers on B1 are 2... initiating')
			# Specify the one string you want to move to the front
			string1 = "RISE TABLET"

			# Check if the specified strings are in the list
			if string1 in B1_Tier_1:
    			# Remove the strings from their original positions
				B1_Tier_1.remove(string1)

    			# Insert the strings at the front of the list
				B1_Tier_1 = [string1] + B1_Tier_1

			print (B1_Tier_1)
			#sleep(2)
			


		index = ALPHABET.index(i[0]) + 1
		print (B1_Tier_1)

		for i in B1_Floor_Rows:
			for o in Shift_Rows:
				if i[1:] == o[1:]:
					#Start creating post for the 4-12er's, they should get celebrate and rise tab first, then we do leads next, then 4 - 11'ers, also remove from
					#B1_Tier list as we go along
					index = ALPHABET.index(i[0]) + 1
					if Total_Closers >= 2: #Built for 2 or more Closers on B1; Everything was [Total_Closers - 1] before
						if 'RISE' in B1_Tier_1[0]:
							Cell = i + ':' + ALPHABET[index + 1] + i[1:]
							Create_Post(Cell, int(i[1:]), index, B1_Tier_1[0])
							print ('Post Made and Removed: ', B1_Tier_1[0])
							B1_Tier_1.remove(B1_Tier_1[0])
							#Total_Closers -= 1
						elif 'RISE' not in B1_Tier_1[0]:
							Cell = i + ':' + ALPHABET[index + 1] + i[1:]
							#Create_Post(Cell, int(i[1:]), index, 'SHOE PREP**')
							random.shuffle(RANDOMNESS)
							if RANDOMNESS[0] == 0: # 66% Chance it will be clear
								Create_Post(Cell, int(i[1:]), index, 'CLEAR')
							else:
								Create_Post(Cell, int(i[1:]), index, 'SHOE PREP**')
							#Cell = ALPHABET[index] + i[1:] + ':' + ALPHABET[index + 1] + i[1:] #If ever needs to be FIXED BACK, please refer to older version code, can literally copy and paste, the commented out code is not it. Can also use the code commented out below for reference
							#Create_Post(Cell, int(i[1:]), index + 1, 'SHOE PREP**')
							B1_Tier_1.remove(B1_Tier_1[0])
							print ('Post Made and Removed: ', B1_Tier_1[0])
							#Total_Closers -= 1
					elif Total_Closers == 1: #If theirs ever only 1 Closer on B1 and 7 4 - 11's
						if 'TAB' in B1_Tier_1[0]:
							Cell = i + ':' + ALPHABET[index + 1] + i[1:]
							Create_Post(Cell, int(i[1:]), index, B1_Tier_1[0])
							print ('Post Made and Removed: ', B1_Tier_1[0])
							B1_Tier_1.remove(B1_Tier_1[0])
							#Total_Closers -= 1
						elif 'RISE' not in B1_Tier_1[0]:
							Cell = i + ':' + ALPHABET[index + 1] + i[1:]
							#Create_Post(Cell, int(i[1:]), index, 'SHOE PREP**')
							random.shuffle(RANDOMNESS)
							if RANDOMNESS[0] == 0: # 66% Chance it will be clear
								Create_Post(Cell, int(i[1:]), index, 'CLEAR')
							else:
								Create_Post(Cell, int(i[1:]), index, 'SHOE PREP**')
							#Cell = i + ':' + i
							#Create_Post(Cell, int(i[1:]), index, B1_Tier_1[0])
							#Cell = ALPHABET[index] + i[1:] + ':' + ALPHABET[index + 1] + i[1:]
							#Create_Post(Cell, int(i[1:]), index + 1, 'SHOE PREP**')
							B1_Tier_1.remove(B1_Tier_1[0])
							print ('Post Made and Removed: ', B1_Tier_1[0])
							#Total_Closers -= 1


		for i in B1_Floor_Rows:
			for L in Lead_Rows:
				if i[1:] == L[1:]:
					Total_Leads += 1 #How many Leads do we have..
					print (i)

		#while ('HELLO' not in B1_Tier_1[0]) or ('PLAZA' not in B1_Tier_1[-1]) or ('U' not in B1_Tier_1[1]):
		#	random.shuffle(B1_Tier_1) #Filtering out time consuming post for Leads, so getting their checklist done is easier


		for i in B1_Floor_Rows:
			for L in Lead_Rows:
				if i[1:] == L[1:]:
					print ('Total Leads: ', Total_Leads)
					Cell = i + ':' + i
					Create_Post(Cell, int(i[1:]), index, FLOAT_LIST[Total_Leads - 1]) #Before it was Float_List it was B1_Tier_1
					Cell = ALPHABET[index] + i[1:] + ':' + ALPHABET[index + 1] + i[1:]
					Create_Post(Cell, int(i[1:]), index + 1, 'SHOE PREP**')
					#print ('Post Made and Removed: ', FLOAT_LIST[Total_Leads - 1])
					#B1_Tier_1.remove(B1_Tier_1[Total_Leads - 1])
					Total_Leads -= 1

		for i in B1_Floor_Rows:
			for u in Closer_11_Row:
				if i[1:] == u[1:]:
					Cell = i + ':' + i
					Create_Post(Cell, int(i[1:]), index, ROTATION_STARTERS[3])
					#print ('Post Made and Removed: ', B1_Tier_1[Total_Leads - 1])
					#B1_Tier_1.remove(B1_Tier_1[Total_Leads - 1])
					




		print ('Done-')
		print (B1_Tier_1)
		wb.save(File_Name)

	B1_Tier_1 = ['RISE TABLET','RISE LINE', "HELLO 1", 'SHOES 1', 'CELEBRATE', 'PLAZA VANDY', 'PLAZA MAD', 'HELLO 2', 'QUEUE', 'LAUNCH**'] #Delete later after its been updated officially, first search it through out all the code



#The function below was Perfectly Fine Here, 
#Just needed to place it earlier for other pre-made functions. 
#If it effects nothing, you may delete this whole block of code



def Temporary_Value(Cord_In_Merged_Cell, the_cell_value):

		# Specify the cell you want to check
		target_cell = ws[Cord_In_Merged_Cell] #-----> Will be used for the parameter of the function

		if the_cell_value is None:

			# Check if the cell is part of a merged cell
			if target_cell.coordinate in ws.merged_cells:
			    # If it is, find the merged cell range
			    merged_cell_range = None
			    for merged_range in ws.merged_cells.ranges:
			        if target_cell.coordinate in merged_range:
			        	merged_cell_range = merged_range
			        	Father_Value = str(merged_cell_range)
			        	#print (Father_Value)
			        	Father_Value = Father_Value[0] + Cord_In_Merged_Cell[1:]
			        	#print (Father_Value)
			        	Father_Value = ws[Father_Value].value
			        	#print (Father_Value)
			        	return (Father_Value)
			        	#sleep(999)
			        	break

			    #if merged_cell_range:
			    	#print(f"The cell {target_cell.coordinate} is part of the merged cell range {merged_cell_range}.... should return the value of the merged cell")
			    #else:
			        #print(f"The cell {target_cell.coordinate} is part of a merged cell, but the range couldn't be determined.")
			#else:
				
			    #print(f"The cell {target_cell.coordinate} is not part of a merged cell.")

			    #print ('In order for this future function to be complete, have it return/print the value of the merged cell its originally apart of')
			    
		elif the_cell_value is not None:
			return (the_cell_value)		     #Keep this Code Alive


				

def Check_Rotations(tier_list, letter_1, letter_2, b1_column_to_redo_post, b1_column_number):
	#print (tier_list)
	#sleep(5)

	

	#letter 1 & 2 Will represent the column of rotations that we're checking for and making sure goes through.
	num = 0
	test = 0

	Relieved_Post = []
	#All post that have been relieved will go inside this list


	'''
	This for loop will help us identify which part of the if Statement we want to use. 
	If its a regular rotation with just 1 post leading the rotation and doesnt have anyone coming back from break, or prep/float, or briefing 
	it will choose the first option, based on whats found in the for loop. The for Loop will check for this. 
	Looping from the very first post of this floor to the last 1. '''
	Letter_Next_To_1 = ALPHABET.index(letter_1)
	Letter_Next_To_1 = ALPHABET[Letter_Next_To_1 + 1]

	Letter_Next_To_2 = ALPHABET.index(Letter_Next_To_1)
	Letter_Next_To_2 = ALPHABET[Letter_Next_To_2 + 1]

	'''
	def Temporary_Value(Cord_In_Merged_Cell, the_cell_value):

		# Specify the cell you want to check
		target_cell = ws[Cord_In_Merged_Cell] #-----> Will be used for the parameter of the function

		if the_cell_value is None:

			# Check if the cell is part of a merged cell
			if target_cell.coordinate in ws.merged_cells:
			    # If it is, find the merged cell range
			    merged_cell_range = None
			    for merged_range in ws.merged_cells.ranges:
			        if target_cell.coordinate in merged_range:
			        	merged_cell_range = merged_range
			        	Father_Value = str(merged_cell_range)
			        	#print (Father_Value)
			        	Father_Value = Father_Value[0] + Cord_In_Merged_Cell[1:]
			        	#print (Father_Value)
			        	Father_Value = ws[Father_Value].value
			        	#print (Father_Value)
			        	return (Father_Value)
			        	#sleep(999)
			        	break

			    if merged_cell_range:
			    	print(f"The cell {target_cell.coordinate} is part of the merged cell range {merged_cell_range}.... should return the value of the merged cell")
			    else:
			        print(f"The cell {target_cell.coordinate} is part of a merged cell, but the range couldn't be determined.")
			else:
			    print(f"The cell {target_cell.coordinate} is not part of a merged cell.")

			    print ('In order for this future function to be complete, have it return/print the value of the merged cell its originally apart of')
		elif the_cell_value is not None:
			return (the_cell_value)
	'''



		
		

	
	




	'''
	B1_Floor_Rows is inaccurate for some odd reason, so we're going to redefine the for loop/function that counts the rows inside this function.
	Replace All B1_Floor_Rows with the new list we will use
	Code Below
	'''

	if letter_1 == 'F': #Why F? Should this change or...?
		Floor_Rows_Recount = []
		start_row = 5
		end_row = int(Locate('OB1')[0][1:])
		for i in range(start_row, end_row):
			cell = ws[letter_1 + str(i)]
			if (cell.value).isupper() and (cell.value) != 'SKIP' and '0' not in (cell.value):
				Floor_Rows_Recount.append(letter_1 + str(i))
	else: #Can Possibly Rewrite this later and delete the upper if statement, not sure atm why I need 2 statements.
		Floor_Rows_Recount = []
		Extra_Rotation_Starters = [] #People coming off break, coming in during the briefing
		None_Vars = 'Cordinates that are apart of larger merged post like FREIGHT, Go Here. We will add there Cord Here, then replace the None with the text of their much larger post'
		start_row = 5
		end_row = int(Locate('OB1')[0][1:])
		for i in range(start_row, end_row):
			cell = ws[letter_1 + str(i)]
			cell2 = ws[Letter_Next_To_1 + str(i)]
			cell3 = ws[Letter_Next_To_2 + str(i)]
			

			'''
			The first if statement below is made to catch FREIGHT post and what comes after. It is designed to determine what merged cell it is apart of
			then write in the post text, where all the None data types will be, so it no longer comes up blank and the rest of the code can read it.
			Might turn it into a function later
			'''
			if cell.value is None and cell2.value is None and cell3.value != 'SKIP' and cell3.value != '0': #May have to mention cell3 Later on here to make sure a post is coming next at least
				Extra_Rotation_Starters.append(Letter_Next_To_1 + str(i)) #Built for after FREIGHT
			elif (cell.value).isupper() and (cell.value) != 'SKIP' and '0' not in (cell.value) and (cell2.value) != EVERYTHING_ELSE[3] and (cell2.value) != 'PREP': #This is Designed specifically for regular Post to be added
				Floor_Rows_Recount.append(letter_1 + str(i))
			elif (cell2.value).upper() == 'PREP' or (cell2.value).upper() == EVERYTHING_ELSE[0] or (cell2.value).upper() == EVERYTHING_ELSE[1] or (cell2.value).upper() == EVERYTHING_ELSE[3]:
				Extra_Rotation_Starters.append(Letter_Next_To_1 + str(i)) #This is designed specifically for Brief or 30 Min Break/Prep Post to be added
			

				
	#The last column of these rotation sheets are always tricky due to people ending their shift; 
	#Usually starts around column 18. Will probably add a remove SKIP at the end or just completely add it to the ROTATION STARTERS list.
	if b1_column_number >= 18:
		ROTATION_STARTERS.append('SKIP')
		print (ROTATION_STARTERS)
		print ('Line 1291: Comment this out and find an alternative for the very last rotation, or make this permanant')
		#sleep(9999)


	#Should be Used for Any Line that has more than just 1 Rotational Post: So far its just F; 
	#Could later on also loop through them lines and have it determine that with a list
	if letter_1 != 'F':
		Floor_Rows_Recount.extend(Extra_Rotation_Starters)
		#print (Extra_Rotation_Starters)
		print ('We are going down this code and making sure prep and short breaks are being counted for in post rotation starters')
		#sleep(9999)


	
	#Referenced Here
	Post_Pushing_Rotations_Cordinates = [] #Will be used for columns that have multiple people pushing the rotations such as briefing, breaks etc
	Rotation_Starting_Post = 0 #Variable used to count how many people will start rotations for this hour
	for i in Floor_Rows_Recount:
		cell = ws[i].value #Will check for the leading post during this rotation containing **
		cell2 = ws[Letter_Next_To_1 + i[1:]].value #Will check for Briefing, Breaks, or Prep or Float going down the column next to letter_1
		cell3 = ws[Letter_Next_To_2 + i[1:]]
		#Rewrite ; May be able to delete cell2 later, if cell checks for briefing and breaks etc due to us combining both list with post and the letter next to them.
		#                              Error Here that has never caused a problem: cell2 == ROTATION_STARTERS???????  
		if cell in ROTATION_STARTERS or cell2 == ROTATION_STARTERS or cell == None and cell2 == None and cell3.value != 'SKIP' and cell3.value != '0':
			Rotation_Starting_Post += 1
			cell = Temporary_Value(i, cell)
			print ('Starting Rotation: '+ cell + ' - ' + i) #Go back inside the function and have them print the original cell if its not apart of a merged cell
			print (Letter_Next_To_1 + i[1:])
			Post_Pushing_Rotations_Cordinates.append(i)
			#sleep(9999)

	print ('Rotation Starters: ', Rotation_Starting_Post)
	


	
	#This if statement will be for Columns or Rotational hours that are started by more than 1 post
	
	if Rotation_Starting_Post > 2:
		print ('Rotation Starters: ', Rotation_Starting_Post)
		print (Floor_Rows_Recount)


		#This will create multiple list inside this list, to keep track of each line of rotation; it will also be used later to make sure every post has been relieved
		#The amount of list created depends on the value of Rotation Starting Post
		Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]
		Relieved_Post_Inner_List_Var = 0 #Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for next
		Loop_Num = 0 #Might can Delete This
		print ('Relieved Post List: ', Relieved_Post)
		sleep(0)
	


		Collapsed_List_Of_Relieved_Post = [] #Will be a collapse list of every list that was relieved during a rotational hour with multiple breaks, briefings etc so we can compare it to the list of post that it should normally be relieving....
		Redo = 0              #This determines if we're going to redo the next rotational hour, if it reachs 10. Refer to line 1288
		Started_Over = 0 	  #Keeps count of how many times we had to do the next rotational hour
		All_Post_Relieved = 0 #Making sure every post, by string, is accounted for. If not, will redo the list; This is specifically made because for some reason, LAUNCH** is being relieved 3 times
		Finished = 0 #A variable that detemines to end the while loop only if everyone is relieved. It should depict a close accuracy to the length of len(Collapsed_List_Of_Relieved_Post); Their will be increments added to this variable for certain situations
		while Finished != len(tier_list) or All_Post_Relieved == 0: #If an issue, change it to <= and ADD the (- 1) next to len(tier list); Change this at some point, make it so that Relieved Post and tier_list have every post checked off in order to stop the loop

			#This Block of code will be used to just redo the whole next column if rotations dont go through effectivly.
			Redo += 1
			index = ALPHABET.index(b1_column_to_redo_post)
			if Redo > 10: #Can make this much less if needed
				print ('Redoing Column Line 1278: ', ALPHABET[index - 1])
				B1_Rotations_Creation(b1_column_to_redo_post, b1_column_number)
				#If the rotations dont properly go through, this piece of code here will redo that column. Its due to the fact a 4th or 5th post during the rotations, ends up going to LAUNCH** or any other Post starter
				Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]
				Redo = 0
				Relieved_Post_Inner_List_Var = 0
				Started_Over += 1
				All_Post_Relieved = 0

			Next_Rotator = 0
			'''
			This variable will be used to help us accurately cut the while loop below, to move on to the next briefing rotation dominoe effect
			'''


			Post_Rotation_Done = 0
			#Will be used to determine when the rotations are done, to help break the for/while loop and decide whether to reshuffle the rotation hour, or not


			print ('Before: ', (Post_Pushing_Rotations_Cordinates))
			#May become useless later, but also cause issues

			if 'LAUNCH' not in ws[Post_Pushing_Rotations_Cordinates[0]].value:
				# Loop through the list to find the index of the item to bring to the front
				for i in range(len(Post_Pushing_Rotations_Cordinates)):
				    if 'LAUNCH' in Temporary_Value(Post_Pushing_Rotations_Cordinates[i], ws[Post_Pushing_Rotations_Cordinates[i]].value): # Was ws[Post_Pushing_Rotations_Cordinates[i]].value Before
				        # Swap the item with the first item, if it has an at risk in it.
				        Post_Pushing_Rotations_Cordinates[i], Post_Pushing_Rotations_Cordinates[0] = Post_Pushing_Rotations_Cordinates[0], Post_Pushing_Rotations_Cordinates[i]
				        break  # Stop the loop once the first occurrence of 5 is moved to the front
				print ('After: ', (Post_Pushing_Rotations_Cordinates))
				print('')
			'''
			This loop helps the code push 'at risk' Post to the front, so its easier to relieve later down the code. 
			Sometimes that at risk post will be relieved twice, therefore breaking the code since 2 exact post cant be within the same Relieved, list of list
			'''

			



			for Post_Starter in Post_Pushing_Rotations_Cordinates: #This list has all the cords that push rotations; Should be around 5 or less
				print ('In The Post Starting Rotations For Loop: ')
				print (cell, cell2)
				print (Post_Pushing_Rotations_Cordinates)
				#sleep(999)
				print('')
				print('')
				print('')
				print ('Post Rotation: ', Post_Rotation_Done)
				print ('Relieved Post Inner Var: ', Relieved_Post_Inner_List_Var)
				print('')
				print('')
				print('')
				print('')
				print ('Majority of Code is right, its just skipped over a Post Starter Row, and Doesnt know what to stop when Relieved Post Var goes over the ;ength of the complete rotation list within list')
				#sleep(99999)
				print ('Rotating: ', Post_Starter)
				print ('Current Rotating Post: ', ws[Post_Starter].value)
				print ('')
				print ('')
				print ('Rotation Post Var: ', Relieved_Post_Inner_List_Var)
				print ('Post Pushing Cord #s: ', len(Post_Pushing_Rotations_Cordinates))
				sleep(0)


				if Relieved_Post_Inner_List_Var == len(Post_Pushing_Rotations_Cordinates):
					print ('Rotations have been fully pushed.... Lets determine if everyone was relieved...')
					break
				elif Relieved_Post_Inner_List_Var > 0:
					print ('Previous List: ', Relieved_Post[Relieved_Post_Inner_List_Var - 1])
					print (Relieved_Post[Relieved_Post_Inner_List_Var - 1][-1])
					if len(Relieved_Post[Relieved_Post_Inner_List_Var - 1]) > 1 and '*' in Relieved_Post[Relieved_Post_Inner_List_Var - 1][-1] and 'LAUNCH' not in Relieved_Post[Relieved_Post_Inner_List_Var - 1][-1]:
						(Relieved_Post[Relieved_Post_Inner_List_Var - 1])[-1] = 'Rotation Starter'
						if '**' in Relieved_Post[Relieved_Post_Inner_List_Var - 1][0] and 'LAUNCH' not in Relieved_Post[Relieved_Post_Inner_List_Var - 1][0]:
							(Relieved_Post[Relieved_Post_Inner_List_Var - 1])[0] = 'Rotation Starter'
					elif len(Relieved_Post[Relieved_Post_Inner_List_Var - 1]) > 1 and '*' in Relieved_Post[Relieved_Post_Inner_List_Var - 1][0] and 'LAUNCH' not in Relieved_Post[Relieved_Post_Inner_List_Var - 1][0]:
						if '**' in Relieved_Post[Relieved_Post_Inner_List_Var - 1][0] and 'LAUNCH' not in Relieved_Post[Relieved_Post_Inner_List_Var - 1][0]:
							(Relieved_Post[Relieved_Post_Inner_List_Var - 1])[0] = 'Rotation Starter'
					elif len(Relieved_Post[Relieved_Post_Inner_List_Var - 1]) == 1 and '*' in Relieved_Post[Relieved_Post_Inner_List_Var - 1][-1] and 'LAUNCH' not in Relieved_Post[Relieved_Post_Inner_List_Var - 1][-1]:
						(Relieved_Post[Relieved_Post_Inner_List_Var - 1])[0] = 'Rotation Starter'


					print ('Solution to Infinite While Loop... Delete the last and 1st post starter thats added if it has an atrisk.... Make sure it works with the if statements thats at the end of this code first, because if it delete launch, can be an issue')
					print ('Still not Replacing the post at the front....')
					print ('Previous List Updated: ', (Relieved_Post[Relieved_Post_Inner_List_Var - 1]))

					#if len(Relieved_Post[Relieved_Post_Inner_List_Var - 1]) == 1 and Relieved_Post[Relieved_Post_Inner_List_Var - 1][0] == 'SKIP':
					#	sleep(9999)
					

				'''
				^ Above if statements are writing over certain 'post starting rotations' at the 1st and last part of a list in order for multiple 
				post starting shifts like SHOE PREP** or SHOES 2**, if this doesnt happen, it will forever loop.

				Alternative, we can add up these extra post in both columns and have it tally up every times its relieved.

				'''

						
						#Removing any post that ends with an at risk besides launch, so it doesnt confuse other rotation post later on.
						#Most likely will have to add to this code later on

					
				


				Rotation_Complete = 0
				Endless_Loop = 0 #For list where Rotation Starters relieve one another

				if '**' in Temporary_Value(Post_Starter, ws[Post_Starter].value): #Was ws[Post_Starter].value Before, but changed to combat 'NoneType' values. This function returns a value if its None
					Post_Rotation_Done += 1
					
					while Rotation_Complete == 0: #Should be connected to the amount of post rotation starters there are in this hour rotation
						print ('Loopinggg')
						Endless_Loop += 1
						
						sleep(0)

						for Floor_Cords in Floor_Rows_Recount:
							cell = ws[Floor_Cords].value                 #Will be the value of the first cell in letter_1
							cell2 = ws[letter_2 + Floor_Cords[1:]].value #Will be the value of the 2nd cell in letter_2, literally right next to each other
							

							if Floor_Cords == Post_Starter or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1: #NEW; Shouldnt effect much honestly. Should make sure the code is more accurate
								#This stays, keeps the for loop on balance for correct rotations. Make sure the first cell post in this loop is where its supposed to be and cell2 is what ever post thats being relieved.

							



								if '*' in ws[Post_Starter].value: #Change this to make it more secure maybe?; If statement may be useless here honestly
									print ('Passed Here')
									print ('Currently Relieved: ', Relieved_Post[Relieved_Post_Inner_List_Var])
									if Endless_Loop > 35:
										sleep(4)

									Has_Post_Been_Relieved = any(cell in sublist for sublist in Relieved_Post) 
									#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for

									cell = Temporary_Value(Floor_Cords, cell)
									cell2 = Temporary_Value(letter_2 + Floor_Cords[1:], cell2)
									print (cell, cell2, Has_Post_Been_Relieved)

									#Future Issue: the False value below can be bad if their are 2 of the same rotation starters within the same rotation hour, such as Shoe Prep or Launch. Can maybe have it count how many to plan for?
									if cell in ROTATION_STARTERS and Has_Post_Been_Relieved == False or len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1  :
										if '*' in cell and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
											Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) #Relieved_Post_Inner_List_Var - Determines what list we are using first inside of Relieved_Post; Starting with 0
											Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
											print (Relieved_Post[Relieved_Post_Inner_List_Var])
											
										elif cell in ROTATION_STARTERS and cell2 in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1  :
												if cell in EVERYTHING_ELSE: #If its a full BREAK.... kinda thinking this code is useless too because of the at risk if statement
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													print (Relieved_Post[Relieved_Post_Inner_List_Var])
													
												elif cell not in EVERYTHING_ELSE: #Basically if its LAUNCH or has an at risk
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													print (Relieved_Post[Relieved_Post_Inner_List_Var])

													


												print ('This ROTATION is Ending EARLY')
												print (cell)
												print (cell2)
												print (Relieved_Post)
												print ('Cell Cord: ', Floor_Cords)
												print ('Cell2 Cord: ', letter_2 + Floor_Cords[1:])
												#sleep(9999)
												Relieved_Post_Inner_List_Var += 1
												Rotation_Complete += 1
												print ('Adding Relieved Post Var + 1')
												print('')
												print('')
												print('')
												print('')
												print('')
												print('')
												print('')
												print('')
												print ('Redo: ', Redo)
												print ('---')
												break
										elif Has_Post_Been_Relieved == True: #Error here somewhere, causes it to skip this if statement
											#Just make sure it doesnt repeat itself and recount the first post starter
											print ('Already Been Relieved.')
											print (Relieved_Post[Relieved_Post_Inner_List_Var])
											print (cell, Relieved_Post[Relieved_Post_Inner_List_Var][-1])
											print ('-')
											sleep(0)


											if cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1] and cell2 in FREIGHT or cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1] and cell2 in ROTATION_STARTERS: #Basically any rotation that doesnt need a relief
												print ('This Rotation is complete, lets move on to the next: ', cell2)
												#Redo += 1
												Relieved_Post_Inner_List_Var += 1
												Rotation_Complete += 1
												#Post_Pushing_Rotations_Cordinates.remove(Post_Starter)
												print ('Redo: ', Redo)
												print ('---')
												break
											elif cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 2  : #If the list has 2 or more post and
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												print (cell + ' Just Relieved ' + cell2)
												
												#Floor_Rows_Recount.remove(Floor_Cords)
												#^ Removing the Cordinate that has been relieved so it has less of a list to loop through. Hopefully this provides a solution
												print ('----') #Made to catch lines/post that are apart of the FLOOR TIER POST
										elif len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 2 and Endless_Loop == 4:
											for i in Relieved_Post[Relieved_Post_Inner_List_Var]:
												if i in ROTATION_STARTERS or i in FREIGHT:
													print ('Post needs to END...........')
													sleep(0)
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													break


	



							
				elif  Temporary_Value(Post_Starter, ws[Post_Starter].value) in ROTATION_STARTERS: #Made to catch post that arnt apart of the FLOOR TIER LIST
					print ('Next is Briefing.......')
					print ('')
					Post_Rotation_Done += 1
					sleep(0)


					


					#Uncomment if you want the old code back
					#Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
					#Basically adding the person being relieved by a briefer, automatically to the Relieved_Post list thats in the list. Reason being, the while loop below wouldnt catch it, also its a one time thing anyway. Not deservant of a loop


					Last_Post_Relieved = cell2
					'''
					This variable will be used as a place holder, to literally help with rotations going through. 
					For some reason, the last post relieved isnt being recognized in the list as -1, so lets just make a var
					'''
					End_Loop = 0 #Will be used as a sefety measure to end the while loop below
					Next_Rotator += 1 #Should be based and determine on how many Rotation Starters are in this hour rotation or even if hour breakers are coming back from break
					while Rotation_Complete == 0: #Def may have to change later, could actually just use the break statement down in the while loop to break it instead of variable increments
						
						print ('Looping')
						print ('Rotation Complete Var: ', Rotation_Complete)
						sleep(0)

						for Floor_Cords in Floor_Rows_Recount:
							cell = ws[Floor_Cords].value                 #Will be the value of the first cell in letter_1
							cell2 = ws[letter_2 + Floor_Cords[1:]].value #Will be the value of the 2nd cell in letter_2, literally right next to each other

							if Floor_Cords == Post_Starter or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1: 
							#This stays, keeps the for loop on balance for correct briefing rotations. Make sure the first cell post in this loop is BRIEF and cell2 is what ever post thats being relieved.

								

								if cell in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
									Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
									print (Relieved_Post[Relieved_Post_Inner_List_Var])
									print (cell)
									#Basically adding the person being relieved by a briefer, automatically to the Relieved_Post list thats in the list. Reason being, the while loop below wouldnt catch it, also its a one time thing anyway. Not deservant of a loop




								if ws[Post_Starter].value in ROTATION_STARTERS: # If statement may be useless here honestly, could also change to Rotation Starters in general

									Has_Post_Been_Relieved = any(cell in sublist for sublist in Relieved_Post) 
									#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for

									print ("Whats Left: ", Floor_Rows_Recount)
									print ('Relieved List: ',Relieved_Post)
									print ('Currently on Cord: ', Floor_Cords)
									print ('')
									#This will help us find and determine where the problem lyes. Rows are being skipped for some odd reason

									search_item = cell2
									is_present = any(search_item in sublist for sublist in Relieved_Post)
									#Checking to see if cell2 has been relieved already

									print ('List Inside of List Value: ', Relieved_Post_Inner_List_Var)
									if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
										print ('Very Last Post: ', Relieved_Post[Relieved_Post_Inner_List_Var][-1])

									
										#Add a if relieved list is None statement here
									if is_present == False and cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1]  :
										Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
										print (Relieved_Post[Relieved_Post_Inner_List_Var])
										Last_Post_Relieved = cell2
										print('added')
										print (Last_Post_Relieved)
										
										
										End_Loop = 0
										#^ Removing the Cordinate that has been relieved so it has less of a list to loop through. Hopefully this provides a solution
										if Last_Post_Relieved in ROTATION_STARTERS:
											Relieved_Post_Inner_List_Var += 1
											Rotation_Complete += 1
											print ('ABORTING......... ROTATION Line Should have Ended After This Post: ', Last_Post_Relieved)
											break
										
									elif is_present == True and cell != Relieved_Post[Relieved_Post_Inner_List_Var][-1]: 
										#Just make sure it doesnt repeat itself and recount the first post starter
										#Should end the rotation cycle if the last post added to the list is a post rotation starter
										print ('Already Been Relieved.')
										print (Relieved_Post[Relieved_Post_Inner_List_Var])
										print (Floor_Cords, cell, Relieved_Post[Relieved_Post_Inner_List_Var][-1])
										print ('-')
										End_Loop += 1

										if End_Loop > len(Floor_Rows_Recount) + 3: #Added the 3 for Safety Measures
											Rotation_Complete += 1 #This should help end the while Loop
										#This if statement was made because the while loop would continue if the very last post in the list was a rotation starter.


									
									elif is_present == True and cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1] or Relieved_Post[Relieved_Post_Inner_List_Var][-1] in ROTATION_STARTERS or Relieved_Post[Relieved_Post_Inner_List_Var][-1] in FREIGHT:
										print ('Perfect Brief Rotation Done.......')
										print ('Last Post Relieved: ', Last_Post_Relieved)
										Relieved_Post_Inner_List_Var += 1
										Rotation_Complete += 1
										#Post_Pushing_Rotations_Cordinates.remove(Post_Starter)
										print (Relieved_Post_Inner_List_Var)
										print ('Last Cell That Loop is Left On: ', cell2)
										print ('')
										break
										

							else:
								pass


			
			print (Relieved_Post)
			print ('')
			print ('')



			'''
			-----CHECKING----- 

			Checking Rotations Have Gone Through.....


			Any code below here, checks, and double checks some more, just to make sure every post has been relieved, if it hasnt, it will redo
			the rotations for the next hour.
			'''



			#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
			Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
			#print ('Before Filter: ', Collapsed_List_Of_Relieved_Post)

			#We have to filter this list, to remove any post that are rotation starters besides the on that belongs to a Tier 1 List
			Junk_List = [] #List thats going to have post that we dont want
			for Post in Collapsed_List_Of_Relieved_Post:
				if Post in FREIGHT or Post not in tier_list: 
					Junk_List.append(Post)

			#print ('Junk List: ', Junk_List)

			for Junk_Post in Junk_List:
				Collapsed_List_Of_Relieved_Post.remove(Junk_Post)

			print (Collapsed_List_Of_Relieved_Post)
			print ('# of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post))
			print (tier_list)
			#sleep(59999)

			for post in Collapsed_List_Of_Relieved_Post:

				# Count the occurrences of the item 1 in the list
				count_of_ones = Collapsed_List_Of_Relieved_Post.count(post)

				# Print the result
				#print(count_of_ones)

				if count_of_ones > 2:
					break


			#The Last Rotation Only Works if Collapse length is 10 and 2 Post are Back to Back. We must not make them plaza post, 
			#or we can make it like a 5 percent chance

			'''
			Make this specific to the last rotation row only; This is the exception to the NoBack_To_Back Post function, 
			where 2 post can be back to back only because it completes the rotation
			Back to Back EXCEPTION!

			9 Can work for len(Collapse) only if their are no duplicate post (such as LAUNCH**)
			10 can also work if their is 2 LAUNCHES
			'''
			Post_Are_The_Same = 0
			if b1_column_number >= 18: 
				#Post Are The Same var is used to represent 2 post being back to back at the last hour of the rotation. Ex: TAB for an Hour, then the person has to do TAB again in order for the roations to run smooth
				

				start_row = 5
				end_row = int(Locate('OB1')[0][1:])
				for i in range(start_row, end_row):
					if ws[letter_1 + str(i)].value == ws[letter_2 + str(i)].value:
						print ('Post is the Same: ', ws[letter_2 + str(i)].value)
						print ('')
						Post_Are_The_Same += 1
						#sleep(3)
						#if 'TAB' in ws[letter_2 + str(i)].value:
						#	Post_Are_The_Same += 1
						


				'''
				#Filters the Collapse list for LAUNCH twice, then removes only 1.
				Search_For_Same_Post = list(filter(lambda x: x == 'LAUNCH**', Collapsed_List_Of_Relieved_Post))
				if len(Collapsed_List_Of_Relieved_Post) == 9 and Post_Are_The_Same == 1:
					pass
				elif "**" in Search_For_Same_Post[0] and len(Search_For_Same_Post) == 2:
					Collapsed_List_Of_Relieved_Post.remove(Search_For_Same_Post[0])
					print ('Updated # Of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post)) #Only used for the last Rotation Hour of B1
				'''

			print ('Line 1767: One of the errors is correcting this line and having it fit all columns, same with Line 1796. Can also add an if statement there to account for 11 Relieved Post')
			print ('Making solution... if it works, be sure to edit the if statement directly above to work for the last rotation')
			#sleep(9999)




			#Filters the Collapse list for LAUNCH twice, then removes only 1.

			Search_For_Same_Post = list(filter(lambda x: x == 'LAUNCH**', Collapsed_List_Of_Relieved_Post))
			#print (Search_For_Same_Post)
			if len(Search_For_Same_Post) < 2:
				pass
			elif "LAUNCH" in Search_For_Same_Post[0] and len(Search_For_Same_Post) == 2:
				print ('Before # Of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post))
				Collapsed_List_Of_Relieved_Post.remove(Search_For_Same_Post[0])
				print ('Updated # Of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post)) #Only used for the last Rotation Hour of B1
				#sleep(9999)




			result = all(item in Collapsed_List_Of_Relieved_Post for item in tier_list)
			Finished = len(Collapsed_List_Of_Relieved_Post)

			if result and count_of_ones < 3: #May have to edit this
			    print("All Post Have Been Relieved.")
			    All_Post_Relieved += 1
			elif Finished == 10 and count_of_ones < 3 and Post_Are_The_Same == 1:
				print("All Post Have Been Relieved...... 2 Post will be Back to Back to Complete Rotations")
				All_Post_Relieved += 1
			#elif Finished == 9 and count_of_ones < 3 and Post_Are_The_Same == 1:
			#	print("All Post Have Been Relieved...... 2 Post will be Back to Back to Complete Rotations")
			#	All_Post_Relieved += 1
			#	Finished += 1
			else:
			    print("Not All Post Have Been Relieved.....")
			    Redo += 10
			    print ('Skipped is being added to some brackets are full on rotatios')
			    sleep(0)

			#if b1_column_number >= 18:
			#	sleep(0)

			print ('')
			print ('')




							



		#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
		#Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
		print ('Completely Done... Everything Relieved')
		print ('Redid the Rotation Hour: ', Started_Over)
		print ('All Post Relieved: ', All_Post_Relieved)
		#sleep(9999)
		


	









	if Rotation_Starting_Post == 1:



	
		while len(Relieved_Post) != len(tier_list):
			sleep(0)
			#print (Floor_Rows_Recount)

			index = ALPHABET.index(b1_column_to_redo_post)
			if num > 5:
				print ('Redoing Column: ', ALPHABET[index - 1])
				B1_Rotations_Creation(b1_column_to_redo_post, b1_column_number)
				#If the rotations dont properly go through, this piece of code here will redo that column. Its due to the fact a 4th or 5th post during the rotations, ends up going to LAUNCH** or any other Post starter
				Relieved_Post = []
				num = 0
			


			#We will be looping through B1_Floor_Row, or if its renamed later. to basically only loop through cells that have a post inside of it
			for i in Floor_Rows_Recount:
				cell = ws[i].value
				cell2 = ws[letter_2 + i[1:]].value

				print (cell)


				if cell in ROTATION_STARTERS and cell not in Relieved_Post:
					Relieved_Post.append(cell)
					Relieved_Post.append(cell2)
					print ('Starting Rotations With: ', cell)
					print ('Relieving Post: ', cell2)
					#We're going to have it now search for the value that cell 2 gave it, possibly loop it untill everything in Relieved Post is in B1_TIER_1
					#Cant use the same if statements cause it could be confusing for column with other rotation starters
				elif cell in Relieved_Post: #Error here somewhere, causes it to skip this if statement
					#Just make sure it doesnt repeat itself and recount the first post starter
					print (Relieved_Post)
					print (cell, Relieved_Post[-1])
					print (i)
					test += 1
					sleep(0)


					if '*' in cell2:
						print ('This Post Starts the Rotation and has already been Accounted for: ', cell2)
						print ('Might have to redo this list...')
						num += 1
						print ('Num: ', num)
					elif cell == Relieved_Post[-1] and len(Relieved_Post) >= 2:
						Relieved_Post.append(cell2)
						print (cell + ' Just Relieved ' + cell2)
						test = 0
							

		if num < 5:
			print ('ATTENTION:')
			print ('Rotations Went through properly')
			print (Relieved_Post)
		else:
			print ('Rotations didnt go through properly')



#Old Retired Code Above and In between these lines
#-----------------------------------------------------------------------------------------------------
#Below this line is the updated version of these functions

def Temporary_Value(Cord_In_Merged_Cell, the_cell_value):

		# Specify the cell you want to check
		target_cell = ws[Cord_In_Merged_Cell] #-----> Will be used for the parameter of the function

		if the_cell_value is None:

			# Check if the cell is part of a merged cell
			if target_cell.coordinate in ws.merged_cells:
			    # If it is, find the merged cell range
			    merged_cell_range = None
			    for merged_range in ws.merged_cells.ranges:
			        if target_cell.coordinate in merged_range:
			        	merged_cell_range = merged_range
			        	Father_Value = str(merged_cell_range)
			        	#print (Father_Value)
			        	Father_Value = Father_Value[0] + Cord_In_Merged_Cell[1:]
			        	#print (Father_Value)
			        	Father_Value = ws[Father_Value].value
			        	#print (Father_Value)
			        	return (Father_Value)
			        	#sleep(999)
			        	break

			    #if merged_cell_range:
			    	#print(f"The cell {target_cell.coordinate} is part of the merged cell range {merged_cell_range}.... should return the value of the merged cell")
			    #else:
			        #print(f"The cell {target_cell.coordinate} is part of a merged cell, but the range couldn't be determined.")
			#else:
				
			    #print(f"The cell {target_cell.coordinate} is not part of a merged cell.")

			    #print ('In order for this future function to be complete, have it return/print the value of the merged cell its originally apart of')
			    
		elif the_cell_value is not None:
			return (the_cell_value)		     #Keep this Code Alive


#Will be Rewriting B1_Rotations_Creation #Search Terms: Updated, New Code, newB1, #NewB1, #newB1, 

Previous_Post_Pushing_Rotations_Cordinates = []
#A List created to hold the previous Cords of post that start the Rotations on B1 for the previous hour.

def B1_Rotations_Creation_Updated(column_letter, col_num):
	global Previous_Post_Pushing_Rotations_Cordinates

	B1_Max_Length = len(B1_Tier_1) + len(B1_Tier_2)
	B1_Min_Length = len(B1_Tier_1)

	B1_Temporary_Post = []
	B1_Temporary_Post.extend(B1_Tier_1)

	B1_Leads = []
	#Will be a list for B1 Leads only.

	Post_Pushing_Rotations_Cordinates = []
	#This will be a list of Cords Starting the Rotation

	Floor_B1 = B1_Floor_Rows.copy()
	#This count will hold the cords for all GEA/Ambassadors. Will also use this to gather the amount of GEA on the floor

	B1_All_Rows_For_Relief = B1_Floor_Rows.copy()
	#AI Assisted Solution

	OB1_Row_Num = int(Locate('OB1')[0][1:])

	#This for statement places B1 Leads into a list
	for i in Lead_Rows:
		#print (i)
		if int(i[1:]) < OB1_Row_Num:
			B1_Leads.append(i)




	# --- UPDATED: Determine if Leads Float or cover Tier 1 Posts --- AI Assisted Code
	# We check if we have enough regular staff to cover the Tier 1 requirements


	# --- SURGICAL LEAD LOGIC START (FG FIXED) ---
	# Identify which cells are empty in the current column (e.g., Column G)
	Available_GSA = [i for i in B1_Floor_Rows if Temporary_Value(i, ws[i].value) == None]

	# Filter for Active Leads (those NOT on BREAK right now)
	B1_Active_Leads_List = []
	for a in B1_Leads:
		current_cell_check = column_letter + a[1:]
		if Temporary_Value(current_cell_check, ws[current_cell_check].value) != EVERYTHING_ELSE[3]:
			B1_Active_Leads_List.append(a)

	# Calculate regular GSA count
	regular_staff_count = len(Available_GSA) - len(B1_Active_Leads_List)
	
	# How many leads are mandatory to reach the 13 required for B1_Tier_1
	shortage = len(B1_Tier_1) - regular_staff_count

	Leads_To_Work = []
	Leads_To_Float = []

	if shortage <= 0:
		# Staffing is good! All active leads float.
		Leads_To_Float = B1_Active_Leads_List
	else:
		# Low staffing: Draft from the bottom of the list first to keep Lead[0] floating
		Leads_To_Work = B1_Active_Leads_List[-shortage:] 
		Leads_To_Float = B1_Active_Leads_List[:-shortage] 
		print(f"Hour {column_letter} | Shortage: {shortage} | {len(Leads_To_Work)} Lead working | {len(Leads_To_Float)} Lead floating")

	# Assign Floating Leads
	for a in Leads_To_Float:
		row_num = a[1:]

		# NEW: Skip this lead if they are doing Freight in this column
		current_cell = ws[column_letter + str(row_num)]
		if current_cell.coordinate in ws.merged_cells:
			# If the cell is merged, check if the merge contains the word 'FREIGHT'
			# This prevents writing 'FLOAT' over 'FREIGHT'
			if "FREIGHT" in str(Temporary_Value(current_cell.coordinate, current_cell.value)):
				continue 
		
		# Identify current and previous columns
		col_idx = ALPHABET.index(column_letter) # G is 6
		letter_before = ALPHABET[col_idx - 1]   # F is 5
		
		# Create the FG Merge (e.g., F10:G10)
		# We start the merge from the letter_before
		Cell_Merge = letter_before + row_num + ':' + column_letter + row_num
		
		# The 'index' for Create_Post should be the numeric column of 'F'
		# index = (ALPHABET index of F is 5) + 1 = 6
		start_col_index = col_idx 

		if a == B1_Leads[0] and column_letter == 'G':
			Create_Post(Cell_Merge, int(row_num), start_col_index, 'RADIO')
			ws[letter_before + row_num].fill = PatternFill(patternType='solid', fgColor=RADIO)
			ws[column_letter + row_num].fill = PatternFill(patternType='solid', fgColor=RADIO)
			print(f"RADIO assigned to {letter_before + row_num}:{column_letter + row_num}")
		else:
			Create_Post(Cell_Merge, int(row_num), start_col_index, EVERYTHING_ELSE[0])
			ws[letter_before + row_num].fill = PatternFill(patternType='solid', fgColor=FLOAT)
			ws[column_letter + row_num].fill = PatternFill(patternType='solid', fgColor=FLOAT)

		# !!! CRITICAL: Remove the Floating Lead row from the Ambassador pool
		# This ensures GSAs never try to fill any cell in Row {row_num}
		Floor_B1 = [c for c in Floor_B1 if c[1:] != row_num]

	# Refresh the pool: GSAs + the Leads who were drafted to work
	Floor_B1_Create_Post_Cords = [c for c in Floor_B1 if Temporary_Value(c, ws[c].value) == None]
	# --- SURGICAL LEAD LOGIC END ---


	


	'''
	Delete this chunk of code when ready
	#Old Code Below 7/10/26 Revert back to this if you need too. Updated Code Above. Also that else statement was incomplete, never got to finish it. 
	That detemined if just 1 or both leads got post.
	Available_GSA = []
	for i in B1_Floor_Rows:
		if Temporary_Value(i, ws[i].value) == None:
			Available_GSA.append(i)
	# List of more accurate available cells.

	# --- 2. Filter for Active Leads (NOT on break) ---
	# We use your lead list and check if they have a 'BREAK' in the current column
	B1_Active_Leads_List = [] 
	for lead_row_cord in B1_Leads: 
		current_cell_to_check = column_letter + lead_row_cord[1:]
		val = Temporary_Value(current_cell_to_check, ws[current_cell_to_check].value)
		
		# If the lead's cell for THIS hour is NOT a break, they are active
		if val != EVERYTHING_ELSE[3]: 
			B1_Active_Leads_List.append(lead_row_cord)

	# --- 3. Calculate regular GSA count ---
	# This tells us if we have enough ambassadors to cover Tier 1 (usually 13)
	regular_staff_count = len(Available_GSA) - len(B1_Active_Leads_List)

	#regular_staff_count = len(Available_GSA) - len(B1_Leads)
	Which_Lead [0, 1]
	random.shuffle(Which_Lead)

	if regular_staff_count >= len(B1_Tier_1): 
		# CASE 1: Plenty of staff. Leads can FLOAT.
		print ('Staffing is Good: LEADS WILL FLOAT')
		for i in Floor_B1[:]: # Use [:] to safely iterate while removing
			for a in B1_Leads:
				if i[1:] == a[1:]:
					Column_Number = ALPHABET.index(i[0]) + 1
					index = ALPHABET.index(i[0]) + 1
					Cell = i + ':' + ALPHABET[index] + i[1:]

					if a == B1_Leads[0] and Floor_B1[0][0] == 'F':
						Create_Post(Cell, int(i[1:]), Column_Number, 'RADIO')
						ws[i].fill = PatternFill(patternType='solid', fgColor=RADIO)
						#Post_Pushing_Rotations_Cordinates.append(i)
					elif Temporary_Value(i, ws[i].value) != EVERYTHING_ELSE[3]:
						Create_Post(Cell, int(i[1:]), Column_Number, EVERYTHING_ELSE[0])
						ws[i].fill = PatternFill(patternType='solid', fgColor=FLOAT)
						#Post_Pushing_Rotations_Cordinates.append(i) 
					
					if i in Floor_B1:
						Floor_B1.remove(i) # Lead is handled, remove from post pool
	else:
		# CASE 2: Short staffed. Leads MUST take Tier 1 posts.
		print ('Low Staffing: LEADS WILL COVER TIER 1 POSTS')
		# In this case, we do NOT remove the leads from Floor_B1.
		# They stay in the pool so they are assigned a job like everyone else.
		if len(B1_Leads) == 1 and regular_staff_count == len(B1_Tier_1) - 1:
			Post_Pushing_Rotations_Cordinates += (B1_Leads)
		elif len(B1_Leads) == 2 and regular_staff_count == len(B1_Tier_1) - 1:
			Post_Pushing_Rotations_Cordinates.append(B1_Leads[Which_Lead[0]])
		else:
			pass 
	'''



	# --- Creating B1 Post ---
	Floor_B1_Create_Post_Cords = Floor_B1.copy()


	#Creating B1 Post

	Floor_B1_Create_Post_Cords = Floor_B1.copy()
	#Used for Creating POST only.


	#This helps filters the new copy list and removes any break, prep, or freight cords.
	for i in Floor_B1:
		if Temporary_Value(i, ws[i].value) != None:
			print (i)
			Floor_B1_Create_Post_Cords.remove(i)

		


	Amount_Of_Post_To_Add = len(Floor_B1_Create_Post_Cords) - len(B1_Temporary_Post)
	print ('Add this # of Post: ', Amount_Of_Post_To_Add)

	if Amount_Of_Post_To_Add > 0:

		#This should randomize and add any extra post needed to fill in every cell for the hour
		B1_Temporary_Post += B1_Tier_2[:Amount_Of_Post_To_Add]
		random.shuffle(B1_Temporary_Post)

	#Double checks and make sure theirs enough post that match the amount of B1 Cords/ B1 GEA for this hour.
	if len(Floor_B1_Create_Post_Cords) != len(B1_Temporary_Post):
		print (Floor_B1_Create_Post_Cords)
		print (len(Floor_B1_Create_Post_Cords))
		print (B1_Temporary_Post)
		print (len(B1_Temporary_Post))
		print ('Somethings Off... Their isnt enough post to match B1 Cords for this Hour.. Line 2171')

	print ('B1 Temporary Post: ', B1_Temporary_Post)

	B1_Creation_Whole = 0
	#This var will be in control of the whole post rotations below as well as check and balances. It will determine if we need to restart everything, if every Tier 1 post hasnt been relieved yet.

	Re_Rotated = 0
	#Var will represent the amount of times the algorithm re

	Full_Shift_12PM_Closers = []
	#This will hold a list of closers that stay until 12 on B1

	Leave_At_11_Closers = []
	#This will hold a list of closers that are only here until 11 on B1

	#Only Column R will have 3 - 4 Different Post in Total for the Closing shifts (Rise Tab & Line, Prep, Clear)
	if Floor_B1[0][0] != 'R':
		while B1_Creation_Whole != 2:
			#We chose 2 because there are 2 Rotation Checks at the very end. Once each is completed perfectly, the while Loop stops completely.


			#---Will Start Creating Rotations Below---
			#PLEASE READ: I HAVENT FULLY EDITED THE CODE BELOW

			random.shuffle(B1_Temporary_Post)
			#Incase we have to restart this whole process.

			Cords_Already_Checked = []
			#To prevent the Check and Balance process from repeating the same Post Starter Cords

			#Post_Pushing_Rotations_Cordinates = [] - Delete Later if it Works
			#This will be a list of Cords Starting the Rotation

			Post_Starter_Already_Relieved = []
			#will be a list consisted of cords that have already started rotations and their full rotation line has been completed

			All_Cords_Downstairs = []
			#This will be a list that counts every cord on B1 that has a post

			Post_Tier = 0
			#This var will help us keep track of what post is next in the list of B1_Temporary_List to create down the column

			Cords_To_Ignore = []
			#A list that holds cords of BREAKS and FREIGHT already pre-made. A bit lazy of me but it solves the issue


			#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
			Relieved_Cords = []

			#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
			Relieved_Post_Inner_List_Var = 0 

			Letter_Next_To_1 = ALPHABET.index(Floor_B1[0][0]) #Will represent the first Letter/Column originally in the Floor B1 List
			Letter_Exactly_Next_To_1 = ALPHABET[Letter_Next_To_1 - 1] #Will represent the first letter Left in the original Floor B1 List
			Letter_Next_To_1 = ALPHABET[Letter_Next_To_1 - 2] #Will represent the second Letter/Column Left in original Floor B1 List

			for i in Floor_B1_Create_Post_Cords:
				#Will collect all cords of pre-made BREAKS and FREIGHT and add them to a list.

				if Temporary_Value(i, ws[i].value) == None:
					print ('SKIP')
				#elif 'FREIGHT' in Temporary_Value(i, ws[i].value) or Temporary_Value(i, ws[i].value) == 'PREP' or Temporary_Value(i, ws[i].value) == EVERYTHING_ELSE[3]:
				#	Cords_To_Ignore.append(i)

				#AI Assisted Solution
				elif 'FREIGHT' in Temporary_Value(i, ws[i].value) or Temporary_Value(i, ws[i].value) == 'PREP' or Temporary_Value(i, ws[i].value) == EVERYTHING_ELSE[3] or Temporary_Value(i, ws[i].value) == EVERYTHING_ELSE[0] or Temporary_Value(i, ws[i].value) == 'RADIO':
					Cords_To_Ignore.append(i)



			#Creating the B1 post
			for i in Floor_B1_Create_Post_Cords:
				if i not in Cords_To_Ignore:
					Column_Number = ALPHABET.index(i[0]) + 1 #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
					Cell = i + ':' + ALPHABET[Column_Number] + i[1:] #Not merging cells, so we dont relly need the cell var but we can keep it
					Create_Post(Cell, int(i[1:]), Column_Number, B1_Temporary_Post[Post_Tier])
					B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
					ws[i].fill = B1_Cell
					print ('Post Added: ', B1_Temporary_Post[Post_Tier])
					print (Post_Tier)
					print (i)
					print ('')
					if "**" in B1_Temporary_Post[Post_Tier] or B1_Temporary_Post[Post_Tier] in ROTATION_STARTERS:
						Post_Pushing_Rotations_Cordinates.append(i) #Possibly will need to be a global list thats seperate from this, thats from the previous rotations that were made
						#Pre Adding Cords of Post that start the Rotations on B1
				else:
					print ('Cord Already Filled: ', i)
					sleep(0)


				''' AI Assisted Revision - Delete Later if it works.
				#Will catch the Briefing Cords and add them to the Post Starter List. Very Important
				Letter_For_Briefing_Column = (ALPHABET.index(i[0]) - 1) #Should be Letter/ Column i
				Letter_For_Briefing_Column = ALPHABET[Letter_For_Briefing_Column]
				
				if ws[Letter_For_Briefing_Column + i[1:]].value != None and ws[Letter_For_Briefing_Column + i[1:]].value == 'BRIEF':
					Previous_Post_Pushing_Rotations_Cordinates.append(Letter_For_Briefing_Column + i[1:])

				#Will catch the PREP Cords and add them to the Post Starter List. Very Important
				if ws[Letter_For_Briefing_Column + i[1:]].value != None and ws[Letter_For_Briefing_Column + i[1:]].value == 'PREP':
					Previous_Post_Pushing_Rotations_Cordinates.append(Letter_For_Briefing_Column + i[1:])

				#Will catch the mini BREAK Cords and add them to the Post Starter List. Very Important
				if ws[Letter_For_Briefing_Column + i[1:]].value != None and ws[Letter_For_Briefing_Column + i[1:]].value == 'BREAK':
					Previous_Post_Pushing_Rotations_Cordinates.append(Letter_For_Briefing_Column + i[1:])
				'''




				Post_Tier += 1

			#wb.save(File_Name)


			index = ALPHABET.index(Floor_B1_Create_Post_Cords[0][0])
			Letter_To_Check = ALPHABET[index - 2]
			index = ALPHABET.index(Letter_To_Check)
			Letter_To_Check_To_The_Right = ALPHABET[index + 1]
			#print (Letter_To_Check)
			#print (Letter_To_Check_To_The_Right)
			#print ('Done')
			#sleep(99999)
			#Delete all commented code
			Cords_We_Looked_At = []

			#Solution Code Here for the Moment - This code finds Post that start the Rotation, that may have been skipped over

			# --- MASTER SCAN: Detect movers from the column immediately to the left --- AI Assisted Code
			# NEW: We clear the list first to ensure we only care about the most recent moves
			Previous_Post_Pushing_Rotations_Cordinates = [] 

			for i in B1_All_Rows_For_Relief:
				idx = ALPHABET.index(i[0])
				prev_col_letter = ALPHABET[idx - 1]
				
				prev_post_val = Temporary_Value(prev_col_letter + i[1:], ws[prev_col_letter + i[1:]].value)

				
				
				if prev_post_val in ['BRIEF', 'PREP', 'BREAK'] or (prev_post_val and 'FLOAT' in str(prev_post_val)):
					# We only add it if it's the column directly to the left
					Previous_Post_Pushing_Rotations_Cordinates.append(prev_col_letter + i[1:])

			'''
			if Floor_B1[0][0] == 'H':
				print ('')
				print (Previous_Post_Pushing_Rotations_Cordinates)
				print ('Take a Look')
				sleep(99999)
				# Delete Later if Fixed
			'''




			if Floor_B1[0][0] != 'F':
				start_row = 5 #For B1 its the number 5
				end_row = int(Locate('OB1')[0][1:]) #End Row its 'OB1'
				for i in range(start_row, end_row):
					Cords_We_Looked_At.append(i)
					List_Of_Post_To_Ignore_For_The_Moment = [None, 'SKIP', 'FLOAT'] #Removed FLOAT - Delete this comment later if it works
					List_Of_Post_On_The_Right_Only = ['PREP', EVERYTHING_ELSE[3]]
					the_main_cord = Letter_To_Check + str(i)
					cell = ws[Letter_To_Check + str(i)]
					cell_Right = ws[Letter_To_Check_To_The_Right + str(i)]

					if Temporary_Value(the_main_cord, ws[Letter_To_Check + str(i)].value) in List_Of_Post_To_Ignore_For_The_Moment :
						pass
					elif 'FREIGHT' in Temporary_Value(the_main_cord, ws[Letter_To_Check + str(i)].value) and cell.value == None and cell_Right.value == None:
						Previous_Post_Pushing_Rotations_Cordinates.append(Letter_To_Check_To_The_Right + str(i))
						print ('Adding Cord Option 1: ', Letter_To_Check_To_The_Right + str(i))
					elif cell.value in ROTATION_STARTERS and cell_Right.value == None:
						Previous_Post_Pushing_Rotations_Cordinates.append(Letter_To_Check + str(i))
						print ('Adding Cord Option 2: ', Letter_To_Check + str(i))
					elif cell.value in List_Of_Post_On_The_Right_Only and cell_Right.value in List_Of_Post_On_The_Right_Only:
						Previous_Post_Pushing_Rotations_Cordinates.append(Letter_To_Check_To_The_Right + str(i))
						print ('Adding Cord Option 3: ', Letter_To_Check_To_The_Right + str(i))
					#elif Temporary_Value(the_main_cord, ws[Letter_To_Check + str(i)].value) == EVERYTHING_ELSE[3] and cell_Right == None:
					#	Previous_Post_Pushing_Rotations_Cordinates.append(Letter_To_Check + str(i))
					#	print ('Adding Cord Option 4: ', Letter_To_Check + str(i))
							




			#Deleting Duplicate Cords To Prevent Extra Relieved Post List Being Created Below:
			Previous_Post_Pushing_Rotations_Cordinates = list(set(Previous_Post_Pushing_Rotations_Cordinates))

			# Track numbers we have already processed
			seen_numbers = set()
			# Store the filtered coordinates
			filtered_coordinates = []

			for cord in Previous_Post_Pushing_Rotations_Cordinates:
				# Extract the numeric part (everything from index 1 to the end, e.g., '26' from 'M26')
				cord_number = cord[1:]
			    
				# If we haven't seen this number yet, keep the coordinate
				if cord_number not in seen_numbers:
					filtered_coordinates.append(cord)
					seen_numbers.add(cord_number) # Mark this number as seen

			# Overwrite your original list variable with the clean results
			Previous_Post_Pushing_Rotations_Cordinates = filtered_coordinates

			print("Cleaned List:", Previous_Post_Pushing_Rotations_Cordinates)
			# Output will keep 'M26', 'M27', 'M31', 'M43', 'M47' 
			# and filter out 'L26' (shares 26), 'K27' (shares 27), 'L31' (shares 31)



			No_Leads_Cords_Added = 0
			#This var will represent the while loop below to ensure it filters out Lead Cords with breaks inside of the Post Pushing List

			#Lazy Coding.... Lead's Breaks are added to this list for some reason.
			while No_Leads_Cords_Added < len(B1_Leads):
				for i in Previous_Post_Pushing_Rotations_Cordinates:
					for a in Locate('LEAD'):
						if i[1:] == a[1:] and Temporary_Value(i, ws[i].value) == EVERYTHING_ELSE[3]: #Make sure its a 'BREAK''
							print ('Removing a Leads Cord: ', i)
							sleep(0)
							Previous_Post_Pushing_Rotations_Cordinates.remove(i)
					if Temporary_Value(i, ws[i].value) not in ROTATION_STARTERS:
						Previous_Post_Pushing_Rotations_Cordinates.remove(i)
				No_Leads_Cords_Added += 1

	                 
	  
			#if Floor_B1[0][0] == 'L':
			#	print (Cords_We_Looked_At)
			#	print (Previous_Post_Pushing_Rotations_Cordinates)
			#	print ('Check These Cord')
			#	print (Letter_To_Check_To_The_Right)
			#	print (Letter_To_Check)
			#	sleep(99999) #Can Delete Later ALL THIS CODE. THE WHOLE BLOCK




			


			if Floor_B1[0][0] == 'F':
				wb.save(File_Name)
				print ('Should break the while loop due to it not being needed for the first column')
				break
			

			'''
			B1 POST RELIEF CHECK - 

			This next section below will check and see if all post have been Relieved on B1.
			'''

			print ('Before: ', Previous_Post_Pushing_Rotations_Cordinates)

			

			# 1. Get the row numbers currently in the pushing list
			pushing_rows = {cord[1:] for cord in Previous_Post_Pushing_Rotations_Cordinates}

			# 2. Get the column letters being used in the pushing list (e.g., 'L' or 'M')
			# We take the first character of the first item to find the "active" previous column
			if Previous_Post_Pushing_Rotations_Cordinates:
			    # pairing one of the column letters from the list
				search_letter = Previous_Post_Pushing_Rotations_Cordinates[0][0]
			else:
			    # Fallback: if list is empty, we calculate the letter one to the left of the GSA column
				gsa_col_idx = column_index_from_string(Available_GSA[0][0])
				search_letter = get_column_letter(gsa_col_idx - 1)

			print(f"Scanning Column {search_letter} for missing Starters...")

			# 3. Loop through Available_GSA
			for gsa_cord in Available_GSA:
				row_num = gsa_cord[1:]
			    
			    # Skip if this row is already in the pushing list
				if row_num in pushing_rows:
					continue
			    
			    # 4. Construct the coordinate to check (Pairing Search Letter + GSA Row)
				check_cord = search_letter + row_num
			    
			    # 5. Use Temporary_Value to check the post at that specific coordinate
				cell_val = Temporary_Value(check_cord, ws[check_cord].value)
			    
			    # Determine if it is a Rotation Starter
				is_starter = (cell_val in ROTATION_STARTERS) or (cell_val and '**' in str(cell_val))
			    
				if is_starter:
					print(f"Found Starter {cell_val} at {check_cord}. Adding to Pushing List.")
					Previous_Post_Pushing_Rotations_Cordinates.append(check_cord)
					sleep(0)
			        
			        # Add row to set so we don't process it again
					pushing_rows.add(row_num)

			# Sort the list to keep the rotation order clean
			Previous_Post_Pushing_Rotations_Cordinates = sorted(
				Previous_Post_Pushing_Rotations_Cordinates, 
				key=lambda x: int(x[1:])
			)


			#This var will keep count of how many Post will start rotations on B1
			Rotation_Starting_Post = len(Previous_Post_Pushing_Rotations_Cordinates)
			#Wasnt Previous before... May delete later

			#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
			Relieved_Post = [[] for i in range(Rotation_Starting_Post)]


			print ('After: ', Previous_Post_Pushing_Rotations_Cordinates)

			# Will try by sorting the letters first
			Previous_Post_Pushing_Rotations_Cordinates = natsorted(Previous_Post_Pushing_Rotations_Cordinates)



			for Post_Starter in Previous_Post_Pushing_Rotations_Cordinates:
				print (Post_Starter)

				if Post_Starter not in Cords_Already_Checked:

					if Temporary_Value(Post_Starter, ws[Post_Starter].value) in B1_Tier_2 or EVERYTHING_ELSE[1] == Temporary_Value(Post_Starter, ws[Post_Starter].value) or '**' in Temporary_Value(Post_Starter, ws[Post_Starter].value) or Temporary_Value(Post_Starter, ws[Post_Starter].value) in ROTATION_STARTERS:

						Rotation_Complete = 0 #Var used to help end the while Loop below. It'll become 1 at the end of every complete rotation

						while Rotation_Complete == 0: #Should be connected to the amount of post rotation starters there are in this hour rotation
							print ('Starting With: ', Temporary_Value(Post_Starter, ws[Post_Starter].value))
							print ('Cord of Post Starter: ', Post_Starter)
							Post_Starter_Already_Relieved.append(Post_Starter)
							print ('Full List of Post Starters: ', Previous_Post_Pushing_Rotations_Cordinates)
							print ('List of Cords to Ignore: ', Cords_To_Ignore)
							print ('List of Post Starters Already Rotated: ', Cords_Already_Checked)
							print ('Inner List #: ', Relieved_Post_Inner_List_Var)
							print ('Both Floor List: ', Floor_B1) 
							print ('Actual Floor B1 Rows: ', B1_Floor_Rows)
							print ('Showing Current List Progress: ', Relieved_Post)
							print ('------')
							print ('')
							print ('Finished Test')

							for Cords in B1_All_Rows_For_Relief: # AI Assisted Change. Was Floor_B1 Before

								# ... your cell/cell2 definitions ...

							    # NEW: Check if this specific coordinate (Cords) has ALREADY been relieved by another chain
								if Cords in Relieved_Cords:
									continue # Skip this row! It's already been taken care of by another chain.

								'''
								if Temporary_Value(Post_Starter, ws[Post_Starter].value) != 'SKIP':
									if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0 and Relieved_Post[Relieved_Post_Inner_List_Var][0] != Temporary_Value(Post_Starter, ws[Post_Starter].value):
										print ('Check this')
										print (Relieved_Post[Relieved_Post_Inner_List_Var])
										print (Temporary_Value(Post_Starter, ws[Post_Starter].value))
										sleep(99999)
								'''

								# --- CASE 1: Starter is NOT a SKIP ---
								if Temporary_Value(Post_Starter, ws[Post_Starter].value) != 'SKIP' or Temporary_Value(Post_Starter, ws[Post_Starter].value) != 'BRIEF':
									if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
										if Relieved_Post[Relieved_Post_Inner_List_Var][0] != Temporary_Value(Post_Starter, ws[Post_Starter].value):
											print('Check this (Standard Case): head of chain does not match starter value')
											print(f"Starter Cord: {Post_Starter} | Value: {Temporary_Value(Post_Starter, ws[Post_Starter].value)}")
											print(f"Chain: {Relieved_Post[Relieved_Post_Inner_List_Var]}")
											sleep(0)

								# --- CASE 2: Starter IS a SKIP ---
								else:
								    # 1. Get the coordinate one letter over (e.g., F8 -> G8)
								    # We use .offset(column=1) to find the cell immediately to the right
									next_door_cord = ws[Post_Starter].offset(column=1).coordinate
									expected_value_one_over = Temporary_Value(next_door_cord, ws[next_door_cord].value)

								    # 2. In a SKIP rotation, the 2nd item (index 1) is the first "real" post being relieved
									if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
										actual_value_in_chain = Relieved_Post[Relieved_Post_Inner_List_Var][1]
								        
									if actual_value_in_chain != expected_value_one_over:
										print(f"Check this (SKIP case at {Post_Starter}): 2nd item does not match expected relief")
										print(f"Expected (from {next_door_cord}): {expected_value_one_over}")
										print(f"Actual (in chain index 1): {actual_value_in_chain}")
										print(f"Full Chain: {Relieved_Post[Relieved_Post_Inner_List_Var]}")
										sleep(99999)




								#Old Code
								cell2 = Temporary_Value(Cords, ws[Cords].value)                 	#ws[Floor_Cords].value | Will be the value of the Floor's Cord List Cells which is the most recent list in the excel and also printed out in the terminal when its ran
								cell = Temporary_Value(Letter_Next_To_1 + Cords[1:], ws[Letter_Next_To_1 + Cords[1:]].value) #ws[Letter_Next_To_1 + Floor_Cords[1:]].value | Will be the value of the 2nd cell directly to the left of the most filled in Column

								'''
								#AI Assited Code
								# 1. 'cell' is the post in the current column (e.g., Column I)
								cell = Temporary_Value(Cords, ws[Cords].value)
							    
							    # 2. Get the index of the current column (I) and move it +1 to the right (J)
								curr_col_letter = Cords[0]
								next_col_idx = column_index_from_string(curr_col_letter) + 1
								next_col_letter = get_column_letter(next_col_idx)
							    
							    # 3. 'cell2' is now looking at the NEXT column (e.g., Column J)
								cell2 = Temporary_Value(next_col_letter + Cords[1:], ws[next_col_letter + Cords[1:]].value)
							    
							    # Now, when Row 17 is checked:
							    # cell = I17 (PLAZA VANDY)
							    # cell2 = J17 (Whatever they moved to next)
							    '''

								


								print ('Makes it here: Looping...- Error is somewhere here')
								print ('Amount of GSA Avaliable', regular_staff_count)
								print (B1_Floor_Rows)
								print ('Actual Correct Amount of GSA Available: ', len(Available_GSA))
								print (Available_GSA)
								print ('B1 Tier Length: ', len(B1_Tier_1))
								print('---------------------------------------------------------')
								print (Letter_Next_To_1 + Cords[1:])
								print ('Post Starter: ', Post_Starter)
								print (Cords[1:])
								if Post_Starter[0] == 'K':
									sleep(0)


								Post_Starter_Value = Temporary_Value(Post_Starter, ws[Post_Starter].value)
								print (Post_Starter_Value)
								if Post_Starter[0] == 'I':
									print ('Slowing down to understand')
									sleep(0)
									#Delete this chunk when done
								

								#This code insures that the 1st Cord/Post being checked is a Post Starter.
								# We only enter this row logic if it's our designated starter 
								# OR if we have already started a chain and are looking for the next link.
								# Crucially, we check that this specific coordinate hasn't been used yet.
								current_coord = Letter_Next_To_1 + Cords[1:]

								if (current_coord == Post_Starter and current_coord not in Relieved_Cords) or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1 or Temporary_Value(Post_Starter, ws[Post_Starter].value) == EVERYTHING_ELSE[0] or Temporary_Value(Post_Starter, ws[Post_Starter].value) == 'RADIO' or Temporary_Value(Post_Starter, ws[Post_Starter].value) == EVERYTHING_ELSE[1] and Cords[1:] == Post_Starter[1:] or Temporary_Value(Post_Starter, ws[Post_Starter].value) == EVERYTHING_ELSE[3] and Cords[1:] == Post_Starter[1:] or Temporary_Value(Post_Starter, ws[Post_Starter].value) == 'PREP' and Cords[1:] == Post_Starter[1:] or Temporary_Value(Post_Starter, ws[Post_Starter].value) in FREIGHT and Cords[1:] == Post_Starter[1:]:

									Post_Starter_Value = Temporary_Value(Post_Starter, ws[Post_Starter].value)

									if 'FLOAT' in Post_Starter_Value:
										print ('Made it for the FLOAT')
										sleep(0)
										#delete this when done


									#Re-Inforces the rule of only checking a Rotation Line if the post starts a Rotation or ends one.
									# Added 'RADIO' to the allowed starters to prevent the infinite loop
									if '*' in Post_Starter_Value or Post_Starter_Value in EVERYTHING_ELSE or Post_Starter_Value in FREIGHT or Post_Starter_Value in B1_Tier_2 or Post_Starter_Value == 'PREP' or Post_Starter_Value == 'RADIO' or Post_Starter_Value in ROTATION_STARTERS:
									#if '*' in Post_Starter_Value or Post_Starter_Value in EVERYTHING_ELSE or Post_Starter_Value in FREIGHT or Post_Starter_Value in B1_Tier_2 or Post_Starter_Value == 'PREP':
									#if '*' in Post_Starter_Value or Post_Starter_Value in EVERYTHING_ELSE or Post_Starter_Value in FREIGHT or Post_Starter_Value in B1_Tier_1 or Post_Starter_Value == 'PREP':

										print ('Currently Relieved: ', Relieved_Post[Relieved_Post_Inner_List_Var])
													

										Has_Post_Been_Relieved = any(cell in sublist for sublist in Relieved_Post) 
										#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for

										#Was Here Last!!!!!!!!!!!!!11 - Delete Later
										cell2 = Temporary_Value(Cords, cell2)
										cell = Temporary_Value(Letter_Next_To_1 + Cords[1:], cell)

										#cell - Will be re-defined for Briefing Cords only. So the name "SKIP" won't be relieved and BRIEF will.
										#if cell == 'SKIP':
										#	cell = Temporary_Value(Letter_Exactly_Next_To_1 + Cords[1:], cell)

										#print ('Cell Were Looking to Relieve: ', cell)
										#print (Letter_Exactly_Next_To_1 + Cords[1:])
										#if cell == EVERYTHING_ELSE[1]:
										#	sleep(99999)
										

										print (cell + ' Relieved Status: ', Has_Post_Been_Relieved)


										#This if statement is for the Post Starter cell and also the Post directly after, which may not start the rotation, it will keep it going. Thats why the or statement is for if the list has a higher value than 1
										if cell in ROTATION_STARTERS or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1 or cell == EVERYTHING_ELSE[1]:

											print ('Made it This Far....')
											if '*' in cell and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0 or cell in ROTATION_STARTERS and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) #Relieved_Post_Inner_List_Var - Determines what list we are using first inside of Relieved_Post; Starting with 0
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Cords)
												print ('if statement 1')

												if cell2 in ROTATION_STARTERS or '*' in cell2 or cell2 in B1_Tier_2:
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													print ('Ending Rotations a Bit Early.....')
													#Will End Rotations Early

											elif cell in ROTATION_STARTERS and cell2 in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1:
												print ('if statement 2...')
												if cell in EVERYTHING_ELSE and cell2 in ROTATION_STARTERS or cell in EVERYTHING_ELSE and cell2 in FREIGHT:
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell)
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Cords)
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													print ('if statement 2A..... This ROTATION Will Be Ending EARLY')
													sleep(0)
													break

												elif cell in EVERYTHING_ELSE: #If its a full BREAK.... kinda thinking this code is useless too because of the at risk if statement
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Cords)
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													print ('if statement 2B')
												elif cell not in EVERYTHING_ELSE: #Basically if its LAUNCH or has an at risk
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Cords)
													
													print (Relieved_Post[Relieved_Post_Inner_List_Var])
													


													print ('This ROTATION is Ending EARLY')
													print (cell)
													print (cell2)
													print (Relieved_Post)
													print ('Cell2 Cord: ', Cords)
													print ('Cell Cord: ', Letter_Next_To_1 + Cords[1:])
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													Stop_Var = 0
													print ('Adding Relieved Post Var + 1')
													print('')
													print('')
													print('')
													print ('if statement: 3')
													break

												#Might can delete these last 2 else statements below
												elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Cords)
													print ('Ending this Rotation--')
													

													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
												elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in Relieved_Post[Relieved_Post_Inner_List_Var] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													print (cell + 'Is Relieving ' + cell2)
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Cords)
													print ('if statement 4')

											elif cell not in ROTATION_STARTERS and cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1]:
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Cords)
												print (cell + 'Is Relieving ' + cell2)
												print ('|-----------------__if statement 5__--------------|')

												if cell2 in ROTATION_STARTERS or '*' in cell2:
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													print ('Ending this rotation')
													break
											elif cell == EVERYTHING_ELSE[1] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0: 
											#This will be used for BREAKER SHIFTS after their BRIEFING

												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Cords)
												print (cell + 'Is Relieving ' + cell2)
												print ('if statement 6')

												if cell2 in ROTATION_STARTERS or '*' in cell2:
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													#This if statement ends Rotations super early

												elif cell2 not in ROTATION_STARTERS or '*' not in cell2:
													#Relieving a regular post after briefing, this if statement passes it on to the next if statement
													pass


							

							print ('Software is Good So Far... Must add code so Post dont repeat back to back')
							print ('To avoid this we can have a check balance code here where it goes through all the relieved code and make sure all the main b1 post are in it, or we can just loop through the cords for this hour and see what post are back to back')
							print ('Add down below an if statement that refreshes all the vars needed for this whole check and balance code that restarts the process if everything doesnt get relieved')
							



							#Used to end the while loop because the Relieved_Post_Inner_List_Var will represent a list value that doesnt exist. 
							#Check All Post Were Relieved Here, in general
							if Relieved_Post_Inner_List_Var == len(Relieved_Post):
								print ('----------------------------------')
								print('')
								print ('Relieved Vars value has passed the amount of list we created. Ending Rotation early to force a re-rotate')
								print ('Showing Current List Progress: \n', Relieved_Post)
								Rotation_Complete += 1
								B1_Creation_Whole = B1_Creation_Whole + 1
								print ('')
								print ('')
							elif Relieved_Post_Inner_List_Var != len(Relieved_Post):
								print ('')
								print ('Stop here... add a continue parameter here to restart the whole while loop and skip all the code below... might just have to renew all the var that contain the relief vars only')
								if Letter_Next_To_1 == 'H':
									print (Letter_Next_To_1)
									print (Floor_B1)
									print (Floor_B1_Create_Post_Cords)
									sleep(0)
								continue
								

							


							print ('Phase 2: Checking to see if all important post has gotten relieved.')

							#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
							Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
							print ('Before Filter: ', Collapsed_List_Of_Relieved_Post)

							#We have to filter this list, to remove any post that are rotation starters besides the on that belongs to a Tier 1 List
							Junk_List = [] #List thats going to have post that we dont want
							for Post in Collapsed_List_Of_Relieved_Post:
								if Post in FREIGHT or Post not in B1_Tier_1: 
									Junk_List.append(Post)

							print ('Junk List: ', Junk_List) 

							for Junk_Post in Junk_List:
								Collapsed_List_Of_Relieved_Post.remove(Junk_Post)

							Collapsed_List_Of_Relieved_Post = list(set(Collapsed_List_Of_Relieved_Post))
							#Filtering for Duplicates and Deleting Them

							print('')
							print ('After Filter: ', Collapsed_List_Of_Relieved_Post) 
							print ('# of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post)) 
							print ('B1 Tier 1 List: ', B1_Tier_1) #List of Important Post that need to be Relieved!!!

							#Potential Issue: If we ever choose to sacrifice Launch, this may raise an eye brow. Only in the making of B1 tho.

							Unrelieved_Post = list(set(B1_Tier_1) - set(Collapsed_List_Of_Relieved_Post))
							print ('Tier 1 Post Hasnt Been Relieved: ', Unrelieved_Post)
							#wb.save(File_Name)


							#Checking to see if all B1_Tier 1 Post has been relieved.
							if len(Unrelieved_Post) == 0: #Bypass / Success / True / Green
								print ('All important Post has been relieved')
								print ('')
								print ('Re-Rotated: ', Re_Rotated)
								Rotation_Complete += 1
								B1_Creation_Whole = B1_Creation_Whole + 1
								break

							elif len(Unrelieved_Post) != 0: #Failed / False / Red
								print ('Some post werent relieved')
								print ('Restart Whole While Loop, Check and Balance Process')
								print ('Relieve One Another: ', Unrelieved_Post)

								# Clearing Relieved Post List of List.... Starting Fresh Below
						
								for sublist in Relieved_Post:
								    sublist.clear()

								# Print the result to see the empty inner lists
								print ('Starting Fresh with Relieved Post List of List: ', Relieved_Post)
								print ('Full List of Post Starters: ', Previous_Post_Pushing_Rotations_Cordinates)

								Cords_Already_Checked.clear()
								#Refreshing the list of already Loop'd Post Starter Cords
								#sleep(7)
								# Output: [[], [], []]


								#Making it equal to 0 again because it will restart the whole major while loop from the start. So it will need to bypass the 1st Phase/ Check Process again.
								B1_Creation_Whole = 0
								Relieved_Post_Inner_List_Var = 0
								Re_Rotated = Re_Rotated + 1
								continue

				elif Post_Starter in Cords_Already_Checked:
					print ('This Error is Fixed... If its ever Stopped Here look into it...')
					print ('This cord has already been looked accounted for...')

				Cords_Already_Checked.append(Post_Starter)
				#To prevent Re-Doing the same Post Starter Cords
	elif Floor_B1[0][0] == 'R':
		print ('')
		print ('Last post for the PM Shift')
		print (len(B1_Temporary_Post))
		print (len(Floor_B1_Create_Post_Cords))

		Post_Tier = 0


		#Sorting all the cords that stay until 12 and those that leave at 11
		for i in Floor_B1:
			if ws['S' + i[1:]].value == None:
				Full_Shift_12PM_Closers.append(i)

			if ws[i].value == None and i not in Full_Shift_12PM_Closers:
				Leave_At_11_Closers.append(i)

		print ('')
		print ('People that all Leave at 11: ', Leave_At_11_Closers)
		print (len(Leave_At_11_Closers))
		print ('People that all Leave at 12: ', Full_Shift_12PM_Closers)
		print (len(Full_Shift_12PM_Closers))

		B1_Temporary_Post.clear()
		for i in range(len(Leave_At_11_Closers)):
			B1_Temporary_Post.append('PREP**')


		for i in Leave_At_11_Closers:
			Column_Number = ALPHABET.index(i[0]) + 1 #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
			Cell = i  #Not merging cells, so we dont relly need the cell var but we can keep it
			Create_Post(Cell, int(i[1:]), Column_Number, B1_Temporary_Post[Post_Tier])
			B1_Cell = PatternFill(patternType = 'solid', fgColor = PREP)
			ws[i].fill = B1_Cell
			print ('Post Added: ', B1_Temporary_Post[Post_Tier])
			print (Post_Tier)
			Post_Tier += 1
			print (i)
			print ('')


		B1_Temporary_Post.clear()
		Post_Tier = 0

		B1_Temporary_Post.append('RISE TABLET')
		#Manually adding a Rise Tab Post

		if len(Full_Shift_12PM_Closers) - 1 > 1:
			#Adding Clear post if its possible

			#Clear post will be given to any other free ambassador on B1
			for i in range(len(Full_Shift_12PM_Closers) - 1):
				B1_Temporary_Post.append('CLEAR')

		random.shuffle(B1_Temporary_Post)
		#Keep it random with the Clear and Rise Tab post


		for i in Full_Shift_12PM_Closers:
			Column_Number = ALPHABET.index(i[0]) + 1 #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
			Cell = i + ':' + ALPHABET[Column_Number + 1] + i[1:] #Extends the Cell to 3 Cells
			Create_Post(Cell, int(i[1:]), Column_Number, B1_Temporary_Post[Post_Tier])
			B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
			ws[i].fill = B1_Cell
			print ('Post Added: ', B1_Temporary_Post[Post_Tier])
			print (Post_Tier)
			Post_Tier += 1
			print (i)
			print ('')




			''' In case we need a refernce later... may delete tho
			for i in Floor_B1_Create_Post_Cords:
				Column_Number = ALPHABET.index(i[0]) + 1 #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
				Cell = i + ':' + ALPHABET[Column_Number] + i[1:] #Not merging cells, so we dont relly need the cell var but we can keep it
				Create_Post(Cell, int(i[1:]), Column_Number, B1_Temporary_Post[Post_Tier])
				B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
				ws[i].fill = B1_Cell
				print ('Post Added: ', B1_Temporary_Post[Post_Tier])
				print (Post_Tier)
				print (i)
				print ('')
			'''
		

						


					

	if Floor_B1[0][0] != 'R': #Put this here recently to by pass an error
		print ('Cords that Start the Rotation: ', Post_Pushing_Rotations_Cordinates)
		print ('Cords that Start the Previous Rotation: ', Previous_Post_Pushing_Rotations_Cordinates)
		print ('')
		print (B1_Leads)
		print (Floor_B1)
		print (B1_Floor_Rows)
		print (len(B1_Floor_Rows))

		#Stays at the VERY End - This allows us to continuesly save the Cords from the previous Rotation that start the current ROTATIONS
		Previous_Post_Pushing_Rotations_Cordinates.clear()
		Previous_Post_Pushing_Rotations_Cordinates = Post_Pushing_Rotations_Cordinates.copy()
		return (Previous_Post_Pushing_Rotations_Cordinates)


		



#Currently Updated
B1_Rotations_Creation_Updated('G', 6)

B1_Breaks_Needed = B1_Floor_Rows
#This list will be used later on to help determine who needs a break and who doesnt


B1_Floor_Rows = []
FloorCellCount('H', B1_Floor_Rows, 'OB1', 5)
B1_Rotations_Creation_Updated('I', 8)



#Loop through Columns F and H and if value equal each other, reshuffle and redoo the column
#Check and Balancing
#This funtion below will be used to check own 2 columns and make sure their aren't any back to back post
def NoBack_To_Back_Post(letter_col1_with_text, letter_col2_with_text, letter_column_next_to_letter_col2, num_from_previous_b1_rotat_creation, column_letter_for_the_next_post):
	num = 0
	while num <= 0:
		for i in B1_Floor_Rows: #Replace this later with the lost of the floor, make it a parameter
			cell = ws[letter_col1_with_text + i[1:]].value
			cell2 = ws[letter_col2_with_text + i[1:]].value
			cell3 = ws[column_letter_for_the_next_post + i[1:]].value

			if cell == cell2: #[:4] Add this if you want post back to back to be completely different like no more hello 1 and hello 2 being the back to back post
				print ('These 2 Post are Back to Back...Doing Rotations Again: ', cell + ' & ' + cell2)
				B1_Rotations_Creation(letter_column_next_to_letter_col2, num_from_previous_b1_rotat_creation)
				num -= 1
				wb.save(File_Name)
				#sleep(0)

			if cell2 == cell3: #[:4] Add this if you want post back to back to be completely different like no more hello 1 and hello 2 being the back to back post
				print ('These 2 Post are Back to Back...Doing Rotations Again: ', cell + ' & ' + cell2)
				B1_Rotations_Creation(letter_column_next_to_letter_col2, num_from_previous_b1_rotat_creation)
				num -= 1
				wb.save(File_Name)
				#sleep(0)


			print ('Cell 1:', cell)
			print ('Cell 2: ', cell2)

			#sleep(1000)
			if cell != None and '*' in cell2 and cell == 'SKIP': #May also add breaks and rotation starters in there; for the start of 6 - 10:30 Shifts
				print ('Breakers cant recieve Rotational Post')
				B1_Rotations_Creation(letter_column_next_to_letter_col2, num_from_previous_b1_rotat_creation)
				num -= 1
				wb.save(File_Name)



			

			if cell2[-1] == '*' and cell3 == 'SKIP' and letter_col2_with_text == 'P': #This is exclusivly for the 6 - 10:30's ENDing Shift
				print ('Someone is going home soon but starts the rotation!: ', i )
				B1_Rotations_Creation(letter_column_next_to_letter_col2, num_from_previous_b1_rotat_creation)
				num -= 1
				wb.save(File_Name)

				
		num += 1
		print ('No 2 Post are Back to Back and No one that leaves will start the rotation')
#NoBack_To_Back_Post('F', 'H', 'I', 8, 'J')
#As of 09.12.2025 - Can delete all NoBackToBack functions once B1 is updated and if we decide to include this function inside of the updated version



#Check_Rotations(B1_Tier_1, 'F', 'H', 'I', 8)
#sleep(0)



#Adding Breaks at the Letter J
Total_Breaks = len(B1_Breaks_Needed)
print ('Breaks Needed: ', Total_Breaks)
print ('People That Need Breaks: ', B1_Breaks_Needed)

def RemoveSemiColon(string):
	char_to_remove = ':'
	new_string = string.replace(char_to_remove, '')
	string = new_string
	return (string)

def BreakTime(column_letter, break_list):
	#BREAKS BREAK - Leaving this here so its easy to find later
	global Floor_Column_Count #These 2 vars shouldnt really be global, Total Breaks just holds the value of how many people will need a break...
	Total_Breaks = len(break_list)
	Floor_Column_Count = 0  #Will be used to determine how many post we will need in this column
	num = 3 #Used to help division run smoothly with breaks. Their are 3 windows where breaks can be chosen from
	break_num = 0 #Gonna be used to help issue out breaks for 4 - 11 people and help them be randomized using the Break_Time list
	Short_Breaks = [1, 2] #To help shuffle between 30 minute breaks for etheir 6:30 or 7:00 etc
	start_row = 5
	end_row = int(Locate('OB1')[0][1:])
	random.shuffle(break_list)
	for breaks in range(3):
		Breaks_Now = Total_Breaks // num #Since their are 3 windows where breaks can be issued, we divide them into 3
		temp_list = break_list[:Breaks_Now]

		index = (ALPHABET.index(column_letter))

		for i in temp_list:
			random.shuffle(Short_Breaks)
			cell = ws[ALPHABET[index] + i[1:]]
			print ('For Cell: ', ALPHABET[index] + i[1:])

			Break_Cordinates_Short = Break_Times[break_num + Short_Breaks[0]].format(eat_time=i[1:])
			Break_Cordinates = Break_Times[break_num].format(eat_time=i[1:])
			print ('Break Cord: ',Break_Cordinates_Short)
			print (Break_Cordinates[:3])
			print ('For Shift: ', i)
			the_correct_column = ALPHABET.index(Break_Cordinates[0]) + 1
			the_correct_column_short = ALPHABET.index(Break_Cordinates_Short[0]) + 1

			Left_Side_Break = ['J', 'L', 'N']
			Right_Side_Break = ['K', 'M']
			#If the short breaks are in any of these letter columns, we will add a prep rotation on the opposite side

			filtered_strings = [string for string in Closer_11_Row if len(string) == len(i) and i[1:] in string]
			#Replace the other if statements with the same filtered string

			if (cell.value) is None and filtered_strings: #Gotta make it equal exact string cause 7 and 17 are getting confused, could also do matchaing len
				print ('30 Minute Break')
				BREAK_CELL = PatternFill(patternType = 'solid', fgColor = BREAK)
				ws[Break_Cordinates_Short].fill = BREAK_CELL
				Create_Post(Break_Cordinates_Short, int(i[1:]), the_correct_column_short, 'BREAK')
				if (Break_Cordinates_Short[0]) in Left_Side_Break:
					PREP_CELL = PatternFill(patternType = 'solid', fgColor = PREP)
					PREP_Cord_Index = the_correct_column_short + 1 #The Column Index Number in Alphabet after a short Break
					PREP_Cord = ALPHABET[PREP_Cord_Index - 1] + Break_Cordinates_Short[1:] #Prep Cord should equal up to the next cell after a short break
					ws[PREP_Cord].fill = PREP_CELL
					Create_Post(PREP_Cord, int(i[1:]), PREP_Cord_Index, 'PREP')
				elif (Break_Cordinates_Short[0]) in Right_Side_Break:
					PREP_CELL = PatternFill(patternType = 'solid', fgColor = PREP)
					PREP_Cord_Index = the_correct_column_short - 1 #The Column Index Number in Alphabet after a short Break
					PREP_Cord = ALPHABET[PREP_Cord_Index - 1] + Break_Cordinates_Short[1:] #Prep Cord should equal up to the next cell after a short break
					ws[PREP_Cord].fill = PREP_CELL
					Create_Post(PREP_Cord, int(i[1:]), PREP_Cord_Index, 'PREP')


			elif (cell.value) is None and any(i[1:] in string for string in Shift_Rows):
				print ('Hour Break for Closer')
				BREAK_CELL = PatternFill(patternType = 'solid', fgColor = BREAK)
				ws[RemoveSemiColon(Break_Cordinates[:3])].fill = BREAK_CELL
				Create_Post(Break_Cordinates, int(i[1:]), the_correct_column, 'BREAK')
			elif (cell.value) is None and any(i[1:] in string for string in Lead_Rows):
				print ('Hour Break for Leads')
				BREAK_CELL = PatternFill(patternType = 'solid', fgColor = BREAK)
				ws[RemoveSemiColon(Break_Cordinates[:3])].fill = BREAK_CELL
				Create_Post(Break_Cordinates, int(i[1:]), the_correct_column, 'BREAK')

		Total_Breaks -= Breaks_Now
		num -= 1
		break_num += 3
		break_list = [item for item in break_list if item not in temp_list]
		index = (ALPHABET.index(column_letter)) + 2



	'''
	Will add a Checks and Balance here for Breaks... It will check to see if its B1, 
	if so how many leads, if more than 1 and breaks are at the same time, fix 1.

	We will then do the same for upstairs leads
	'''

	All_Breaks = Locate('BREAK')
	All_B1_Leads_Breaks = []
	Break_Col = ['J', 'L', 'N']
	random.shuffle(Break_Col)


	for i in Locate('LEAD'):
		for a in All_Breaks:
			if a[1:] == i[1:] and int(i[1:]) < int(Locate('OB1')[0][1:]):
				All_B1_Leads_Breaks.append(a)

	# Making seeperate breaks for 2 Leads on the Same Floor
	if len(All_B1_Leads_Breaks) == 2:
		if All_B1_Leads_Breaks[0][0] == All_B1_Leads_Breaks[1][0]:
			print('')
			print ('Same Break Time..... Fixing')

			Old_Break = All_B1_Leads_Breaks[1]
			#Current Break Cord we are aiming to remove

			Break_Col.remove(All_B1_Leads_Breaks[1][0])
			#Removing the Column Letter from the list of Break letters to avoid remaking the same break we are currently erasing.

			New_Break = Break_Col[0] + All_B1_Leads_Breaks[1][1:]
			#New Break Cordinate is remade as a string

			# Reset the fill (this removes the color)
			no_fill = PatternFill(fill_type=None)

			# Apply to the primary cell (A1) and clear its value
			ws[Old_Break].fill = no_fill
			ws[Old_Break].value = None


			# Unmerge the cells
			# Get the cell to the immediate RIGHT (0 rows down, 1 column right)
			cell_to_right = ws[Old_Break].offset(row=0, column=1)

			#Unmerging the combined cells for the old break Cordinate
			ws.unmerge_cells(Old_Break + ':' + cell_to_right.coordinate)

			# Creating the new Break Cell
			New_Break_Right_Cell = ws[New_Break].offset(row=0, column=1)
			New_Break_Right_Cell = New_Break_Right_Cell.coordinate
			New_Break_Merge = New_Break + ':' + New_Break_Right_Cell

			index = ALPHABET.index(New_Break_Merge[0])

			BREAK_CELL = PatternFill(patternType = 'solid', fgColor = BREAK)
			ws[New_Break].fill = BREAK_CELL
			Create_Post(New_Break_Merge, int(i[1:]), index, 'BREAK')
			ws[New_Break].value = 'BREAK'
			ws[New_Break].alignment = Alignment(horizontal='centerContinuous')

			All_Breaks.remove(Old_Break)
			All_B1_Leads_Breaks.remove(Old_Break)
			All_Breaks.append(New_Break)
			All_B1_Leads_Breaks.append(New_Break)

			wb.save(File_Name)

	#Above Code is Good. Below now we will have it catch the B1 Breaks if they are the same, delete one and replace it with the following letters
	#J, L, N

	print ('')
	print (All_Breaks)
	print (All_B1_Leads_Breaks)
	print ('Break Test Done')



BreakTime('J', B1_Breaks_Needed)

print ('Testing Break Times... Line 3051')


#Freight_Canidates = []
Freight_Cord = [] #Freight Cords will go here
B1_Floor_Rows = []
#The Row We Count From and The Row We End basically count all the rows we will consider for Freight, Amount of Freight Lines will represent how many will be created



#------------- Updated Freight Funtion Below-----------------
def Freight_Time(the_row_we_start_counting_from, the_row_we_will_end_row, amount_of_freight_lines):
	global Freight_Used, Freight_Cord
	Freight_Canidates = []
	Freight_Used = 0 #To help keep count of which Freight was used in the Freight list already
	Freight_Lines_Needed = amount_of_freight_lines
	#Will count cells from Start of B1 to the end of B1, in the column that lands immediately after FREIGHT ENDS
	start_row = the_row_we_start_counting_from
	end_row = the_row_we_will_end_row
	for i in range(start_row, end_row):
		cell = ws[Freight_Break_Check + str(i)] #Will be used to see if this person is getting a break during their FREIGHT line
		if Temporary_Value(Freight_Break_Check + str(i), cell.value) == EVERYTHING_ELSE[3] or Temporary_Value(Freight_Break_Check + str(i) ,cell.value) == 'PREP':
			Freight_Canidates.append(Freight_Break_Check + str(i))

		#Grabbing Breaker Shifts Only | 10:30 Has no correlation to 6 - 10:30's at all, and the value is just to make sure it only adds rows that are TRUE to such value.
		if Temporary_Value(Freight_Break_Check + str(i) ,cell.value) is None and ws[Letter_Times['10:30'] + str(i)].value == 'SKIP':
			Freight_Canidates.append(Freight_Break_Check + str(i))
		#Any Shift is Eligble for FREIGHT due to the elif statement below; Comment out if want it to choose only Breaker shifts
		elif Temporary_Value(Freight_Break_Check + str(i) ,cell.value) == 'PREP' and ws[Letter_Times['10:30'] + str(i)].value == None or Temporary_Value(Freight_Break_Check + str(i) ,cell.value) == 'BREAK' and ws[Letter_Times['10:30'] + str(i)].value == None:
			Freight_Canidates.append(Freight_Break_Check + str(i))

		#ws[Letter_Times['10:30'] is Pulled from a dictionary above, where the time represents a letter. The if statement above is creating a Cordinate to check for
	
	random.shuffle(Freight_Canidates)
	print ('Freight Canidates: ', Freight_Canidates)
	print ('')


	Freight_Canidates = list(set(Freight_Canidates)) #Cleaning the list for any duplicates!

	# --- LEAD RESCUE LOGIC (No try/except) ---
	ob1_limit = int(Locate('OB1')[0][1:])
	is_upstairs_mode = amount_of_freight_lines > 1
	
	# 1. Identify leads for the current floor
	if is_upstairs_mode:
		leads_on_floor = Leads_Upstairs
	else:
		leads_on_floor = [ld for ld in Lead_Rows if int(ld[1:]) < ob1_limit]

	# 2. Check if a lead is already in the candidates pool
	lead_in_pool = False
	for cand in Freight_Canidates:
		if any(cand[1:] == ld[1:] for ld in leads_on_floor):
			lead_in_pool = True
			break

	# 3. If no lead is free, find one on break and move them
	if not lead_in_pool and len(leads_on_floor) > 0:
		target_lead = leads_on_floor[0]
		target_row = target_lead[1:]
		
		# Find all breaks currently on the sheet
		break_coords = Locate('BREAK')
		
		if break_coords:
			for b_coord in break_coords:
				# Match the break coordinate to the target lead's row
				if b_coord[1:] == target_row:
					current_b_col = b_coord[0]
					
					# Identify the range to unmerge (breaks are 2 cells wide)
					# Find the specific merged range object to avoid errors
					for m_range in list(ws.merged_cells.ranges):
						if b_coord in m_range:
							ws.unmerge_cells(str(m_range))
					
					# Empty the cell and clear color
					ws[b_coord].value = None
					ws[b_coord].fill = PatternFill(fill_type=None)
					
					# Choose a new break slot (Swap J to N, or L to J)
					new_b_col = 'N' if current_b_col == 'J' else 'J'
					new_col_idx = ALPHABET.index(new_b_col) + 1
					new_break_range = f"{new_b_col}{target_row}:{ALPHABET[new_col_idx]}{target_row}"
					
					# Create the new break on the same row
					Create_Post(new_break_range, int(target_row), new_col_idx, 'BREAK')
					ws[new_b_col + target_row].fill = PatternFill(patternType='solid', fgColor=BREAK)
					
					# Add the now-available lead row to candidates
					Freight_Canidates.append(Freight_Break_Check + target_row)
					print ('')
					print(f"Moved Lead Break from {current_b_col} to {new_b_col} for Row {target_row}")
					print ('')
					break





	while Freight_Used != Freight_Lines_Needed:
		end_it_all = 0 #Helps ends both for loops down below

		#Choosing Lead for Freight: The meaning of this for Loop and if Statement is to give Leads first priority of Counter during Freight, only if they're in the Freight Canidates list. Can be commented out
		#if (Freight_Used + 1) == 4:
		if Freight_Used == 1 or Freight_Lines_Needed == 1: # Ensures we choose a Lead for both B1 and Upstairs for BOH and Counter Freight
			for i in Freight_Canidates:
				for a in Lead_Rows:
					if i[1:] == a[1:]:
						Freight_Canidates.insert(0, i)
						#Add a pop here; not really needed
						print ('Giving Lead Freight...', i)
						#print (Freight_Canidates)
						end_it_all += 1
						break
				if end_it_all == 1:
					break
		
		elif (Freight_Used + 1) < 4:
			for i in Lead_Rows:
				for a in Freight_Canidates[:3]:
					if i[1:] == a[1:] and i[1:] == Freight_Canidates[0][1:]:
						random.shuffle(Freight_Canidates)
						print ('Preserving Freight for Counter', i)
						#Can potentially add a mini while loop to shuffle until that lead row isnt in index 0 but the odds of it still being their is close to none
		

		index = ALPHABET.index(Freight_Start)
		index += 1
		Freight_Cells = Freight_Start + Freight_Canidates[0][1:] + ':' + Freight_End + Freight_Canidates[0][1:]
		Freight_Fill_In_Color = PatternFill(patternType = 'solid', fgColor = FREIGHT_COLOR)
		ws[Freight_Start + Freight_Canidates[0][1:]].fill = Freight_Fill_In_Color
		if int(Freight_Canidates[0][1:]) > int(Locate('OB1')[0][1:]): #Basically tells it to start from index 1 in the FREIGHT list so we dont reprint B1 FREIGHT
			Create_Post(Freight_Cells, int(Freight_Canidates[0][1:]), index, FREIGHT[Freight_Used + 1]) 
			Freight_Cord.append(Freight_Cells)
			print ('Freight Created: ', FREIGHT[Freight_Used + 1])
			print (Freight_Cells)
		elif int(Freight_Canidates[0][1:]) < int(Locate('OB1')[0][1:]):
			Create_Post(Freight_Cells, int(Freight_Canidates[0][1:]), index, FREIGHT[Freight_Used])
			Freight_Cord.append(Freight_Cells)
			print ('Freight Created: ', FREIGHT[Freight_Used])
			print (Freight_Cells)
			
		Freight_Used += 1 #It will keep accurate count of which Freight Line is next in the list.
		Freight_Canidates.remove(Freight_Canidates[0])
		wb.save(File_Name)
		print ('Freight Used: ', Freight_Used)

	return (Freight_Cord)


	

#Freight List is properly created. Randomize it then create post for Freight
Freight_Time(5, int(Locate('OB1')[0][1:]), 1)

B1_Floor_Rows = []
FloorCellCount('J', B1_Floor_Rows, 'OB1', 5)
B1_Rotations_Creation_Updated('K', 10)


B1_Floor_Rows = []
FloorCellCount('L', B1_Floor_Rows, 'OB1', 5)
B1_Rotations_Creation_Updated('M', 12)


wb.save(File_Name)
B1_Floor_Rows = []
FloorCellCount('N', B1_Floor_Rows, 'OB1', 5)
B1_Rotations_Creation_Updated('O', 14)


B1_Floor_Rows = []
FloorCellCount('P', B1_Floor_Rows, 'OB1', 5)
B1_Rotations_Creation_Updated('Q', 16)


B1_Floor_Rows = []
FloorCellCount('R', B1_Floor_Rows, 'OB1', 5)
B1_Rotations_Creation_Updated('S', 18)

print ('Testing Finished for B1')
#sleep(99999)



#B1_Floor_Rows = []
#FloorCellCount('R', B1_Floor_Rows, 'OB1', 5)
#B1_Rotations_Creation('S', 18)
#NoBack_To_Back_Post('P', 'R', 'S', 18, 'T')
#Check_Rotations(B1_Tier_1, 'P', 'R', 'S', 18)



#-------------------Adding Lead Rows----------------
#If we ever decided to start adding lead rows, copy this code and place it where you find best

#Counting how many leads are on B1
Leads_On_B1 = 0
List_Of_All_Leads = Locate('LEAD')
for lead_on_b1 in List_Of_All_Leads:
	if lead_on_b1[0] == 'B' and int(lead_on_b1[1:]) < int(Locate('OB1')[0][1:]):
		Leads_On_B1 += 1

#Counting how many free people we have on OB3 that can be used
Free_Ambassadors = 0 	  #Counting how many free rows are avialable on OB3
Free_Ambassador_Rows = [] #The 4 - 12 Rows we will then use that are currently free on OB3
for full_shift_ambassadors in Shift_Rows:
	if int(Locate('OB3')[0][1:]) < int(full_shift_ambassadors[1:]):
		Free_Ambassadors += 1
		Free_Ambassador_Rows.append(int(full_shift_ambassadors[1:]))

print (Free_Ambassadors)










B1_Tier_1 = ['RISE TABLET','RISE LINE', "HELLO 1", 'SHOES 1', 'CELEBRATE', 'PLAZA VANDY', 'PLAZA MAD', 'HELLO 2', 'QUEUE', 'LAUNCH**']





#Fills in the color for any cell that may be missing some for FLOAT or SHOE PREP
for row in ws.iter_rows():
    for cell in row:
        if cell.value == 'FLOAT':
        	FLOAT_CELL = PatternFill(patternType = 'solid', fgColor = FLOAT)
        	cell.fill = FLOAT_CELL
        	wb.save(File_Name)
        elif cell.value == 'SHOE PREP**':
        	SHOE_PREP_CELL = PatternFill(patternType = 'solid', fgColor = PREP)
        	cell.fill = SHOE_PREP_CELL
        	wb.save(File_Name)
        elif cell.value in B1_Tier_1:
        	B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
        	cell.fill = B1_Cell
        	wb.save(File_Name)
        elif cell.value == 'CLEAR':
        	CLEAR_CELL = PatternFill(patternType = 'solid', fgColor = CLEAR_COLOR)
        	cell.fill = CLEAR_CELL
        	wb.save(File_Name)
        elif cell.value == 'RADIO':
        	RADIO_CELL = PatternFill(patternType = 'solid', fgColor = RADIO)
        	cell.fill = RADIO_CELL
        elif cell.value == 'PREP':
        	PREP_CELL = PatternFill(patternType = 'solid', fgColor = PREP)
        	cell.fill = PREP_CELL
        	wb.save(File_Name)


#This short piece of code cuts the text in half on B1 if its in 1 Cell and to Large to be displayed fully.
start_row = 5
end_row = int(Locate('OB1')[0][1:])
for row_number in range(start_row, end_row + 1):
	if 'RISE' not in ws['R' + str(row_number)].value and ws['R' + str(row_number)].value in B1_Tier_1:
		if len(ws['R' + str(row_number)].value) >= 9:
			ws['R' + str(row_number)].value = ws['R' + str(row_number)].value[:7]



'''
---------------------------------------------FULL TIMERS CODE BELOW
Adding Full Timers 4:00 - 12:30PM Shift

The reason why its added here below and not spread throughout the code, for some odd reason it causes a major error in how B1 is created and messes up my rotations.
Made more sense to just let the code run, work properly, then relabel and add the Full Timers in after everything is said and done. When I rewrite thise code, I want to properly
impletement this block of code in the proper places
'''

#Full Timers list is completely randomized and pull from the overall list of Shift Rows. Their list are random rows/cords for 4 - 12:30 people.
#ISSUE: If they ever go away or anything changes, can comment this whole code out and it goes back to just 4 - 12'ers.


#Start - This part below splits up the shift row and chooses which 4 - 12'ers will be 4 - 12:30's

Temporary_Shift_Rows = [] #This list is made to extract any rows that are on OB3 and removed them from Shift Rows only to avoid putting full timers up there
Row_To_Avoid = Locate('OB3')[0][1:]
Row_To_Avoid = int(Row_To_Avoid) #This should be OB3's Row but as an integer
print ('Before: ', Shift_Rows)

for coordinate in Shift_Rows:
    letter, number_str = coordinate[0], coordinate[1:]
    number = int(number_str)
    
    if number > Row_To_Avoid: #If a Row is on OB3, it will not be considers for the 4 - 12:30's
        Temporary_Shift_Rows.append(coordinate)

Shift_Rows = [item for item in Shift_Rows if item not in Temporary_Shift_Rows] 
#Removing every cord from Shift rows then will add them back later, so nothing on OB3 gets selected for 4 - 12:30
print ('After: ', Shift_Rows)
print (Temporary_Shift_Rows)


random.shuffle(Shift_Rows)
#print (Shift_Rows)

Full_Time_Closers_Row = [] #FT Closing Row Cordinates Go in this List
for i in Shift_Rows:
	if Full_Time_Closers > 0:
		Full_Time_Closers_Row.append(i)
		Full_Time_Closers -= 1
	else:
		break

#Shift_Rows = [x for x in Shift_Rows if x not in Full_Time_Closers_Row]
#Not needed ^ Full Timer shifts are basically the same shifts, they just stay 30 minutes later
Shift_Rows.extend(Temporary_Shift_Rows)
Shift_Rows = sorted(Shift_Rows, key=lambda x: int(x[1:]))
print ('Reattached: ', Shift_Rows)
print ('Full Time Closers', Full_Time_Closers_Row)

#-----------------------End

#Colors in the Full Timers shift
Color_For_Full_Time_Closers = PatternFill(patternType = 'solid', fgColor = FT_SHIFT_COLOR)

#Adds in the Time Stamp for Column B
TimeStamps(4, Full_Time_Closers_Row, Color_For_Full_Time_Closers)

#Re counts the columns for the full timers row, maybe not needed here honestly.
Full_Time_Closers_Row = []
Full_Time_Closers_Row = Recount(Full_Time_Closers_Row, '3:30') #Was 12:30 before

#IMPORTANT: Colors in the Regular closers for 4 - 12 and Full Timers. This code cant be deleted unless the previous Shift Row coloring is uncommented much further above, around line 950
for i in Shift_Rows:
	if i in Full_Time_Closers_Row:
		Letter = 21
		for color in range(2):
			ws[ALPHABET[Letter] + i[1:]].fill = Dark
			ws[ALPHABET[Letter] + i[1:]].value = 'SKIP'
			wb.save(File_Name)

	else:
		Letter = 20           #Starts at the letter U, which is where 12PM starts
		for color in range(2):
			ws[ALPHABET[Letter] + i[1:]].fill = Dark
			ws[ALPHABET[Letter] + i[1:]].value = 'SKIP'
			Letter += 1
			wb.save(File_Name)

#-------------------------ENDS HERE

		



        








        	












print ('Floor Column Count: ', Floor_Column_Count)
print ('B1 is Technically done. Just need to add Float lines for leads, Hello 3 Post, as well as RADIO post')
print ('')
print ('When we start developing for this again, start with line 1202. Will need an if statement in that area based on letter 2 or B1_Floor_Rows length, then add code that eliminates post that have been properly rotated. If any hasnt, then recalculate that column for ppost')
print ('Finished B1')
#sleep(99999)



#------------------------------------The Start of OB1 & OB2-----------------------------------------

#RECOUNTING - Recounting every list, in case we make any changes to B1, everything will effectivly be accounted for in the upcoming floors regardless.

Shift_Rows = []
Shift_Rows = Recount(Shift_Rows, '12')

Breaker_Rows = []
Breaker_Rows = Recount(Breaker_Rows, '6')

Closer_11_Row = []
Closer_11_Row = Recount(Closer_11_Row, '11')

Lead_Rows = []
Lead_Rows = Recount(Lead_Rows, 'LEAD')

Full_Time_Closers_Row = []
Full_Time_Closers_Row = Recount(Full_Time_Closers_Row, '3:30') #Was 12:30 Before

All_Shifts_Rows = [item for sublist in [Shift_Rows, Breaker_Rows, Closer_11_Row, Lead_Rows, Full_Time_Closers_Row] for item in sublist]
#Formally known as the Create Balance var. Can rename them all to just "All_Shifts_Rows" later.





#OB1 & OB2 LEAD LINES - Now we should accurately count how many leads are on OB1 and OB2, as well as how many full shifts are avilable on these floors

OB1_Cordinate = int(Locate('OB1')[0][1:])
OB3_Cordinate = int(Locate('OB3')[0][1:])

#Listing any Leads on OB1 - OB3
Leads_Upstairs = []

#Currently only for OB1 and OB2; Any full shift or full time shift on OB1 and OB2
Shift_Rows_Upstairs = []

#All Full Shifts on OB1 - OB3
Every_Full_Shift_Upstairs_Rows = []

for i in All_Shifts_Rows:
	if i in Lead_Rows: #Only looking for Leads on OB1 and OB2
		if int(i[1:]) > OB1_Cordinate:
			Leads_Upstairs.append(i)

	if i in Shift_Rows:
		if int(i[1:]) > OB1_Cordinate and int(i[1:]) < OB3_Cordinate:
			Shift_Rows_Upstairs.append(i)

	if int(i[1:]) > OB1_Cordinate:
		Every_Full_Shift_Upstairs_Rows.append(i)

			


#Creating Breaks for every Full Shift upstairs
BreakTime('J', Every_Full_Shift_Upstairs_Rows)

All_Breaks = Locate('BREAK')
All_Upstairs_Leads_Breaks = []
Break_Col = ['J', 'L', 'N']
random.shuffle(Break_Col)


for i in Locate('LEAD'):
	for a in All_Breaks:
		if a[1:] == i[1:] and int(i[1:]) > int(Locate('OB1')[0][1:]):
			All_Upstairs_Leads_Breaks.append(a)



# Making seeperate breaks for 2 or more Leads on the Same Floor
if len(All_Upstairs_Leads_Breaks) > 1:
	if All_Upstairs_Leads_Breaks and all(''.join(filter(str.isalpha, c)) == ''.join(filter(str.isalpha, All_Upstairs_Leads_Breaks[0])) for c in All_Upstairs_Leads_Breaks):
		print('')
		print ('Before: ', All_Upstairs_Leads_Breaks)
		print ('Same Break Time..... Fixing')

		Old_Break = All_Upstairs_Leads_Breaks[1]
		#Current Break Cord we are aiming to remove

		Break_Col.remove(All_Upstairs_Leads_Breaks[1][0])
		#Removing the Column Letter from the list of Break letters to avoid remaking the same break we are currently erasing.

		New_Break = Break_Col[0] + All_Upstairs_Leads_Breaks[1][1:]
		#New Break Cordinate is remade as a string

		# Reset the fill (this removes the color)
		no_fill = PatternFill(fill_type=None)

		# Apply to the primary cell (A1) and clear its value
		ws[Old_Break].fill = no_fill
		ws[Old_Break].value = None


		# Unmerge the cells
		# Get the cell to the immediate RIGHT (0 rows down, 1 column right)
		cell_to_right = ws[Old_Break].offset(row=0, column=1)

		#Unmerging the combined cells for the old break Cordinate
		ws.unmerge_cells(Old_Break + ':' + cell_to_right.coordinate)

		# Creating the new Break Cell
		New_Break_Right_Cell = ws[New_Break].offset(row=0, column=1)
		New_Break_Right_Cell = New_Break_Right_Cell.coordinate
		New_Break_Merge = New_Break + ':' + New_Break_Right_Cell

		index = ALPHABET.index(New_Break_Merge[0]) + 1 # Added + 1 here

		BREAK_CELL = PatternFill(patternType = 'solid', fgColor = BREAK)
		ws[New_Break].fill = BREAK_CELL
		Create_Post(New_Break_Merge, int(i[1:]), index, 'BREAK') # Uses the adjusted index
		ws[New_Break].value = 'BREAK'
		ws[New_Break].alignment = Alignment(horizontal='centerContinuous')

		All_Breaks.remove(Old_Break)
		All_Upstairs_Leads_Breaks.remove(Old_Break)
		All_Breaks.append(New_Break)
		All_Upstairs_Leads_Breaks.append(New_Break)
		print ('After: ', All_Upstairs_Leads_Breaks)

		wb.save(File_Name)







#--------------------Create and Checking for RADIO POST


#Searching all B1 Post to see if a Lead was given the RADIO post. If not we will give it a Lead on OB1
RADIO_Post = ws['D' + i[1:]]
Radio_Post_Found = 0
for i in Lead_Rows:
	if int(i[1:]) < OB1_Cordinate:
		print (RADIO_Post.value, int(i[1:])) #Should eventually print BRIEF or RADIO; Error: Currently printing None for some reason
		if RADIO_Post.value == 'RADIO':
			Radio_Post_Found += 1
			print ('Radio Post has been given to a B1 Lead')
			sleep(9999)

#If their isn't a RADIO Lead down stairs, it will give the RADIO post to a Lead Upstairs
if Radio_Post_Found == 0 and len(Leads_Upstairs) > 0:
	print ('Leads are upstairs')
	index = ALPHABET.index('D') + 1 
	Radio_OB1_Cell = 'D' + str(Leads_Upstairs[0][1:]) + ':' + 'E' + str(Leads_Upstairs[0][1:])
	Create_Post(Radio_OB1_Cell, int(Leads_Upstairs[0][1:]), index, 'RADIO')
	RADIO_CELL = PatternFill(patternType = 'solid', fgColor = RADIO)
	ws['D' + str(Leads_Upstairs[0][1:])].fill = RADIO_CELL




#Creating OB1 and OB2 Fright Post
Freight_Time(int(Locate('OB1')[0][1:]), int(ws.max_row) + 1, 3) #Changed from 4 to 3














#-------------------------------------------  Evenly Re-Distributing Out Breaks for OB1 and OB2

'''
This area of Code will be used to evenly distrubute out BREAKs so rotations may go smoothly. Their is a certain number of empty cells we need
per rotation hour in order for all rotations to go smoothly. Currently the minimum is 9, in order to get the exact number just add OB1_Tier_1 and OB2_Tier_1, but we
want at least 11 open cells.

May enclose this under a massive if statement with the condition that if their are a column of empty cells less than 11 do this...
'''

#Counting all Empty Cells fro OB1 - OB2 During the Break Hours
FirstBreakHour = [] #Starts at 6:30/ Column J
SecondBreakHour = [] #Starts at 7:30/ Column L
ThirdBreakHour = [] #Starts at 8:30/ Column N
FloorCellCount('J', FirstBreakHour, 'OB3', int(Locate('OB1')[0][1:]))
FloorCellCount('L', SecondBreakHour, 'OB3', int(Locate('OB1')[0][1:]))
FloorCellCount('N', ThirdBreakHour, 'OB3', int(Locate('OB1')[0][1:]))

print ('')
#These are the numbers below for empty cells that can be used later to create OB1 and OB2 post, but now we will use these numbers to evenly distrubute breaks
print (len(FirstBreakHour))
print (len(SecondBreakHour))
print (len(ThirdBreakHour))

# The text you are searching for
search_text = "FREIGHT"

# List to store freight cell coordinates that match the search text
FREIGHT_CELLS = []

# Iterate through all cells in the worksheet
for row in ws.iter_rows():  # This iterates over rows
    for cell in row:  # This iterates over cells in the row
    	if cell.value == None:
    		pass
    	elif search_text in cell.value:
            # Append the coordinate (like 'A1', 'B2') to the list
            FREIGHT_CELLS.append(cell.coordinate)

#All OB1 and OB2 Break Cords will go here
BREAKS_UPSTAIRS = []

for row in ws.iter_rows():  # This iterates over rows
    for cell in row:  # This iterates over cells in the row
    	if cell.value == None: #or all(cell.coordinate not in item for item in Freight_Cord)
    		pass
    	elif EVERYTHING_ELSE[3] in cell.value and int(cell.coordinate[1:]) > int(Locate('OB1')[0][1:]) and int(cell.coordinate[1:]) < int(Locate('OB3')[0][1:]):
            # Append the coordinate (like 'A1', 'B2') to the list
            BREAKS_UPSTAIRS.append(cell.coordinate)



BREAKS_THAT_CANT_MOVE = []
#This will be a list of break cords that we wont touch specifically because they cant be move due to FREIGHT post being right next to them.


#This for loop will capture all the breaks that have freight next to them then add the to a list
for i in BREAKS_UPSTAIRS:
	if i[0] == 'J':
		index = ALPHABET.index(i[0]) + 2
		#print ('For the Js: ', ALPHABET[index])
		if Temporary_Value(ALPHABET[index] + i[1:], ws[ALPHABET[index] + i[1:]].value) == None:
			pass
		elif 'FREIGHT' in Temporary_Value(ALPHABET[index] + i[1:], ws[ALPHABET[index] + i[1:]].value):
			BREAKS_THAT_CANT_MOVE.append(i)
	elif i[0] == 'N':
		index = ALPHABET.index(i[0]) - 1
		#print ('For the Ns: ', ALPHABET[index])
		if Temporary_Value(ALPHABET[index] + i[1:], ws[ALPHABET[index] + i[1:]].value) == None:
			pass
		elif 'FREIGHT' in Temporary_Value(ALPHABET[index] + i[1:], ws[ALPHABET[index] + i[1:]].value):
			BREAKS_THAT_CANT_MOVE.append(i)
	elif i[0] == 'L':
		pass





OB1_OB2_Empty_Cells = FirstBreakHour + SecondBreakHour + ThirdBreakHour
#This List will hold all the cords in one big list of all the empty cells on OB1 and OB2 that dont have breaks

All_Free_Rows_List = [FirstBreakHour, SecondBreakHour, ThirdBreakHour]
#This list will only contain the list of list of empty cells on OB1 and OB2

print ('')
print (len(OB1_OB2_Empty_Cells))
Goal_Rows = (len(OB1_OB2_Empty_Cells) // 3)
print ('Each Row should hold this number or more', Goal_Rows)
Extra_Rows = (len(OB1_OB2_Empty_Cells) % 3)
print ('These are the extra rows: ', Extra_Rows)
#May can delete this code




# Print the distributed lists
print ('')
print("List 1:", FirstBreakHour)
print("List 2:", SecondBreakHour)
print("List 3:", ThirdBreakHour)
print ('Breaks Upstairs: ', BREAKS_UPSTAIRS)
print ('Breaks That Cant Move: ', BREAKS_THAT_CANT_MOVE)
print ('')



# Loop through the list
for cell in BREAKS_UPSTAIRS:

	# Determine if this break belongs to a 4-11 shift
    is_4_11 = any(cell[1:] == closer[1:] for closer in Closer_11_Row)
    
    # Only erase if it's NOT a 4-11 and NOT a fixed freight break
    if cell not in BREAKS_THAT_CANT_MOVE and not is_4_11:
	    # Unmerge the cell if it's part of a merged cell range
	    for merged_cell in list(ws.merged_cells.ranges):
	        if cell in str(merged_cell):
	            ws.unmerge_cells(str(merged_cell))
	    
	    # Clear the text in the cell
	    ws[cell].value = None
	    
	    # Remove the fill color (set it to 'None' or a transparent fill)
	    ws[cell].fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type=None)



#This var will hold the value of the amount of breaks thats needed upstairs
Amount_Of_Breaks_To_Issue_Upstairs = 0

Avoid_Repeated_Cords = []
#Some cords repeat for some odd reason.... we will add them here... or Possibly clean the list for duplicates at some point

''' Old
for i in All_Shifts_Rows:
	if int(i[1:]) > int(Locate('OB1')[0][1:]) and int(i[1:]) < int(Locate('OB3')[0][1:]) and all(i[1:] not in item for item in Breaker_Rows) and i not in Avoid_Repeated_Cords and all(i[1:] not in item for item in BREAKS_THAT_CANT_MOVE):
		Amount_Of_Breaks_To_Issue_Upstairs += 1
		Avoid_Repeated_Cords.append(i)
'''

for i in All_Shifts_Rows:
    # NEW: Check if this shift is a 4-11
	is_4_11 = any(i[1:] == closer[1:] for closer in Closer_11_Row)
    
    # ADDED 'and not is_4_11' to the condition below
	if int(i[1:]) > int(Locate('OB1')[0][1:]) and int(i[1:]) < int(Locate('OB3')[0][1:]) and all(i[1:] not in item for item in Breaker_Rows) and i not in Avoid_Repeated_Cords and all(i[1:] not in item for item in BREAKS_THAT_CANT_MOVE) and not is_4_11:
		Amount_Of_Breaks_To_Issue_Upstairs += 1
		Avoid_Repeated_Cords.append(i)



print ('Upstairs Breaks Needed: ', Amount_Of_Breaks_To_Issue_Upstairs)
print ('When we re make breaks, start with lines that have freight first, then subtrace from that list of cords that need break. dont think we made it so include it in the loop above')


print (Freight_Cord)


Breaks_Given = BREAKS_THAT_CANT_MOVE
#This will be a list that keeps count of how many and what rows have been given a break already.



# Define the start and end coordinates
start_cell = Locate('OB1')[0][1:]
end_cell = Locate('OB3')[0][1:]

# Extract the row numbers from the cell coordinates
start_row = int(start_cell)
end_row = int(end_cell)


Rows_That_Need_A_Break = []
#This will be a list of rows that need a break upstairs. Should include row numbers only

# Loop through the cells in the column from start_row to end_row

''' Old
for row in range(start_row + 1, end_row):
    if all(str(row) not in item for item in BREAKS_THAT_CANT_MOVE) and all(str(row) not in item for item in Breaker_Rows):
    	Rows_That_Need_A_Break.append(str(row))
    	#Capturing every Row/Line that still needs a break
'''


for row in range(start_row + 1, end_row):
    # Skip Freight breaks, Breakers, AND 4-11 shifts
	is_4_11 = any(str(row) == closer[1:] for closer in Closer_11_Row)
	if all(str(row) not in item for item in BREAKS_THAT_CANT_MOVE) and all(str(row) not in item for item in Breaker_Rows) and not is_4_11:
		Rows_That_Need_A_Break.append(str(row))
    	#Capturing every Row/Line that still needs a break


print ('Rows that still need a Break: ', Rows_That_Need_A_Break)



print ('')
print ('Updated List')
FirstBreakHour = [] #Starts at 6:30/ Column J
SecondBreakHour = [] #Starts at 7:30/ Column L
ThirdBreakHour = [] #Starts at 8:30/ Column N
FloorCellCount('J', FirstBreakHour, 'OB3', int(Locate('OB1')[0][1:]))
FloorCellCount('L', SecondBreakHour, 'OB3', int(Locate('OB1')[0][1:]))
FloorCellCount('N', ThirdBreakHour, 'OB3', int(Locate('OB1')[0][1:]))


random.shuffle(Rows_That_Need_A_Break)
#This help spreads the breaks out down below by randomizing it. Not doing so will get the same results but it will just look weird and breaks will be on top of each other...


#This while Loop will evenly remove empty cells from all 3 Columns above to keep their length around the same so its easier to create post and therefore create breaks for OB1 and OB2
while Amount_Of_Breaks_To_Issue_Upstairs != 0:

	
	All_Free_Rows_List = [FirstBreakHour, SecondBreakHour, ThirdBreakHour]
	#This list will only contain the list of list of empty cells on OB1 and OB2

	for cord in Rows_That_Need_A_Break:
		print ('Current Cord: ', cord)

		# Sort All_Free_Rows_List by length, in descending order; This helps us keep all 3 Break columns near the same length and also diversify BREAKS through out the 3 columns J, L, N
		All_Free_Rows_List = sorted(All_Free_Rows_List, key=len, reverse=True)


		# Loops through the first list in this big list so we can evenly distribute the breaks. The above code moves the list around based on length
		for i in All_Free_Rows_List[0]:

		    
			if i[1:] == str(cord):
				print ('Creating a Break for: ', i)
				Amount_Of_Breaks_To_Issue_Upstairs -= 1
				Breaks_Given.append(i)
				All_Free_Rows_List[0].remove(i)
				Rows_That_Need_A_Break.remove(cord)
				break
			elif all(str(cord) in item for item in Breaks_Given):
				pass

			

	print ('Breaks Left to Make: ', Amount_Of_Breaks_To_Issue_Upstairs)
	print ('Breaks Given: ', Breaks_Given)
	print ('')


#Before Creating Breaks below, maybe swap out some of the j column breaks with other letters? or have it count which letter has the most breaks an then do swap outs




#Creating the Breaks Down Below
for i in Breaks_Given:

	Break_Closing_Letter = ALPHABET.index(i[0]) + 1
	Cell = i + ':' + ALPHABET[Break_Closing_Letter] + i[1:]


	if ws[i].value == None:
		BREAK_CELL = PatternFill(patternType = 'solid', fgColor = BREAK)
		ws[i].fill = BREAK_CELL
		Create_Post(Cell, int(i[1:]), Break_Closing_Letter, EVERYTHING_ELSE[3])
		wb.save(File_Name)
	elif ws[i].value == EVERYTHING_ELSE[3]:
		pass
	




# Output the results
for i, lst in enumerate(All_Free_Rows_List, start=1):
	print(f"List {i}: {lst} (Length: {len(lst)})")

print ('Breaks Given: ', Breaks_Given)



wb.save(File_Name)
print ('Check the Breaks 2: Line 4428')
#sleep(99999)




#Counting Empty Cells from OB1 - OB2
OB1_OB2_Floor_Rows = []
#FloorCellCount('F', OB1_OB2_Floor_Rows, 'OB3', int(Locate('OB1')[0][1:]))

#If len Floor_Rows 9 or less, go for ob3 count, place this inside the Floor Cell Count tho. 
#Might make 2 seperate list in there just in case, also this should depend on if freight function is activated or not



def Upper_Floor_Rotation_Creation_OB1():
	global OB1_OB2_Floor_Rows, OB1_Tier_2, Floor_OB1
	Floor_OB1_Backup = [] #Will be used for OB1 Post being recreated, as a reference or copy
	Floor_OB2_Backup = [] #Will be used for OB2 Post being recreated, as a reference or copy
	#Creating a Back Ups for a for loop down below. Use find, then next. The very next time this list is brought up, thats what it will be used for



	OB1_Max_Length = len(OB1_Tier_1) + len(OB1_Tier_2)
	OB1_Min_Length = len(OB1_Tier_1)
	OB2_Max_Length = len(OB2_Tier_1) + len(OB2_Tier_2)
	OB2_Min_Length = len(OB2_Tier_1)
	Upstairs_Min_Length = len(OB1_Tier_1) + len(OB2_Tier_1) #This should be the minimum for both floors to operate, currently 11

	OB1_Temporary_Post = [] #This list will be used to help determine what post will be added for the hour. Should be temporary and always changing
	OB1_Temporary_Post.extend(OB1_Tier_1)
	OB2_Temporary_Post = [] #This list will be used to help determine what post will be added for the hour. Should be temporary and always changing
	OB2_Temporary_Post.extend(OB2_Tier_2) #Not sure why this is Tier 2????

	Post_Tier = 0
	#This var will help us keep track of what post is next in the list of OB1_Temporary_List to create down the column

	Remade_Post = 0
	#This var will count every time a column is remade due to every post not being relieved properly

	Previous_Check = 0
	#Will be used to represent if the previous column rotated properly or had any sort of hiccups. If its 1, it will ignore running the code again

	First_Column = Floor_OB1[0][0]
	#Capturing the 1st Letter of the column so we can use it later for the previous check algorithm. We have to make sure it isnt the letter F

	No_Float_Left_Behind = []
	#This will contain cords of lead rows, and we will loop through it to see if a lead cord/cell was left empty

	

	if len(OB1_OB2_Floor_Rows) > Upstairs_Min_Length: #Dont exaclty remember why I made this if statement...? Might change the len to Floor_OB1 and min length as well

		#Counting how many Leads are on OB1. This will determine if they recieve a FLOAT post or not.
		Lead_On_OB1 = 0 
		for a in Floor_OB1:
			for i in Leads_Upstairs:
				if a[1:] == i[1:]:
					Lead_On_OB1 += 1

		if (len(Floor_OB1) - Lead_On_OB1) > OB1_Min_Length:
			print ('Their are more Rows on OB1 than the Minimum, not counting the Lead. (A)')

			#This var will be used to detemine how many more post will be needed to fill up OB1 for the hour
			Amount_Of_Post_To_Add = (len(Floor_OB1) - Lead_On_OB1) - OB1_Min_Length

			#Giving this a 20% Chance to include AFF 3 in the rotation
			Random_List = [0, 0, 0, 0, 1]
			random.shuffle(Random_List)

			if Random_List[0] == 1:
				index = OB1_Tier_2.index('AFF 3**')
				item_to_move = OB1_Tier_2.pop(index)  # Remove the last item from the list; Should be the index AFF 3 is in
				desired_index = 2  # Index to insert the item

			#Adding the perfect amount needed so every empty cell is filled in for OB1
			for i in range(0, Amount_Of_Post_To_Add):
				OB1_Temporary_Post.append(OB1_Tier_2[i])

			OB1_Tier_2 = ['S ELE**', 'TR2**', 'TR3**', 'TR4**', 'TR5**', 'AFF 3**'] 
			#Restoring the original list


			OB1_Hour_Post = 0 
			#This is a var that will be used to end this while loop below for OB1. It is designed to make sure an hour rotation is created with 2 OB1 Post

			Checks = 0
			#This var will be used to determine if proper checks were out into place to make sure every post goes through properly

			
			while OB1_Hour_Post != 2 and Checks != 1:


				#Randomize the list of the post that will be distributed
				random.shuffle(OB1_Temporary_Post) 
				
				
				if len(Floor_OB1) == len(OB1_Temporary_Post) + 1 and Lead_On_OB1 > 0:
					OB1_Temporary_Post.append(EVERYTHING_ELSE[0])
					'''
					Im adding a FLOAT post to this list only because this IF STATEMENT is designed to fill up a column on the basis that their are enough
					ambassadors to have the floor running properly. This float post will be given to the lead
					'''


					#Placing the lead cord at the end of this list
					#ATTENTION: Create a list for lead cords, make it global and at the end of these functions, fill them all in just in case we ever missed one
					#SOLUTION: To make sure every float is filled in for ob1 basically
					print ('Create a list for lead cords, make it global and at the end of these functions, fill them all in just in case we ever missed one')
					#sleep(99999)

					num = 0
					print ('Before: ', Floor_OB1)
					for Lead in Floor_OB1:
						for a in Leads_Upstairs:
							if Lead[1:] == a[1:]:
								Floor_OB1.remove(Lead)  # Remove the item from its current position
								Floor_OB1.append(Lead)  # Append the item to the end of the list
								No_Float_Left_Behind.append(Lead) #Gathering Lead cords in case we need to fill them in later. Sometimes those lead cells where FLOAT should be, has a value of None
								num += 1
								break
							if num == 1:
								break
					print ('After: ', Floor_OB1)

				


				#Creating the OB1 post
				for i in Floor_OB1:
					Column_Number = ALPHABET.index(i[0]) + 1 + OB1_Hour_Post #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
					Cell = i #Not merging cells, so we dont relly need the cell var but we can keep it
					Create_Post(Cell, int(i[1:]), Column_Number, OB1_Temporary_Post[Post_Tier])
					OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
					ws[i].fill = OB1_Cell
					print ('Post Added: ', OB1_Temporary_Post[Post_Tier])
					print (Post_Tier)
					Post_Tier += 1
					wb.save(File_Name)
					




				#Error Here: This list may become empty at some point so it may need to be replaced, also make sure this doesnt get repeated so see if its attached to OB1 Check Var and if not add it
				print ('Floor OB1 List: ', Floor_OB1)
				print ('OB1 Hour Var: ', OB1_Hour_Post)
				#Previously was Floor_OB1[0][0] != 'F'  
				if OB1_Hour_Post == 0 and First_Column != 'F': #Possibly add another parameter here with the new Previous Check var and place this if statement into a while Loop
					print ('Stopping... Place the Parameter to check for brief or something. This shouldnt work until a 2nd Column is already made. After column G')
					print (Floor_OB1[0][0])
					sleep(0)
					wb.save(File_Name) 

					while Previous_Check == 0: 
						print ('Checking Routations from the previous line goes through properly')
						print (Floor_OB1[0][0])
						sleep(0)


						#Check and Balance for Previous Line goes here! 
						#Start----------------------------------------------------------------


						Post_Pushing_Rotations_Cordinates = [] 
						#Will be used for columns that have multiple people pushing the rotations such as briefing, breaks etc

						Letter_Next_To_1 = ALPHABET.index(Floor_OB1[0][0]) #Will represent the first Letter/Column originally in the Floor OB1 List
						Letter_Next_To_1 = ALPHABET[Letter_Next_To_1 - 1] #Will represent the second Letter/Column Left in original Floor OB1 List

						#Letter_Next_To_2 = ALPHABET.index(Letter_Next_To_1)
						#Letter_Next_To_2 = ALPHABET[Letter_Next_To_2 + 1]


						for i in Floor_OB1:
							cell2 = ws[i].value 
							cell = ws[Letter_Next_To_1 + i[1:]].value #Will check for the leading post during this rotation containing ** atrisk
							

							# Temporary_Value(Post_Starter, ws[Post_Starter].value)
							if cell in ROTATION_STARTERS or '**' in cell:
								Post_Pushing_Rotations_Cordinates.append(Letter_Next_To_1 + i[1:])
								#Grabbing every cell that starts the rotation on OB1's Double Rotation, and inserts it into a list. 
								#Update this


						if ws[Post_Pushing_Rotations_Cordinates[-1]].value == EVERYTHING_ELSE[0] and ws[Letter_Next_To_1 + Post_Pushing_Rotations_Cordinates[-1][1:]].value == EVERYTHING_ELSE[0]:
							#Both Values in the above if statement should be FLOAT and apart of a Lead line

							Post_Pushing_Rotations_Cordinates.pop()
							Floor_OB1.pop() 
							#This if statement is responsible for removing Lead cord from the OB1 Floor list and Post Pushing list so its easier to create a check and balance.
							#Lead cords DONT need to be in this list since all their post will mostly be Float and RADIO.
							#This only will take effect if both the post created are float that are back to back.
							



						#Checks and Balance Starts Here
						


						#This var will keep count of how many Post will start rotations on OB1
						Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates)

						#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
						Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]

						#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
						Relieved_Cords = []

						#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
						Relieved_Post_Inner_List_Var = 0 

						Stop_Var = 0
						#Can delete later, only used for testing
						
			
						while Previous_Check == 0:
							

							#Looping through all the starting post in the OB1 Post starter list.
							for Post_Starter in Post_Pushing_Rotations_Cordinates:

								if '**' in Temporary_Value(Post_Starter, ws[Post_Starter].value) or Temporary_Value(Post_Starter, ws[Post_Starter].value) in ROTATION_STARTERS:

									Rotation_Complete = 0 #Var used to help end the while Loop below. It'll become 1 at the end of every complete rotation

									while Rotation_Complete == 0: #Should be connected to the amount of post rotation starters there are in this hour rotation
										print ('Starting With: ', ws[Post_Starter].value)
										print ('Cord of Post Starter: ', Post_Starter)
										print ('Full List of Post Starters: ', Post_Pushing_Rotations_Cordinates)
										print ('Inner List #: ', Relieved_Post_Inner_List_Var)
										print ('Floor OB1 List: ', Floor_OB1)
										print ('Showing Current List Progress: ', Relieved_Post)
										print ('------')
										print ('')
										sleep(0)

										for Floor_Cords in Floor_OB1:
											cell2 = ws[Floor_Cords].value                 		 #Will be the value of the Floor OB1 List Cells which is the most recent list in the excel and also printed out in the terminal when its ran
											cell = ws[Letter_Next_To_1 + Floor_Cords[1:]].value #Will be the value of the 2nd cell directly to the left of the most filled in Column
											#Used to end the while loop because the Relieved_Post_Inner_List_Var will represent a list alue that doesnt exist
											if Relieved_Post_Inner_List_Var == len(Relieved_Post):
												print ('Relieved Vars value has passed the amount of list we created. Ending Rottion early to force a re-rotate')
												Rotation_Complete += 1
												break


											if (Letter_Next_To_1 + Floor_Cords[1:]) == Post_Starter or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1: 
												#This is to prevent the code from relieving any Post that isnt a rotator or has an atrisk

												
											


												if '*' in ws[Post_Starter].value: #Change this to make it more secure maybe?; If statement may be useless here honestly
													print ('Currently Relieved: ', Relieved_Post[Relieved_Post_Inner_List_Var])

													

													#Has_Post_Been_Relieved = any(cell in sublist for sublist in Relieved_Post) 
													#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for
													source_cord_to_check = Letter_Next_To_1 + Floor_Cords[1:]
													Has_Post_Been_Relieved = source_cord_to_check in Relieved_Cords


													#Was Here Last!!!!!!!!!!!!!11 - Delete Later
													cell2 = Temporary_Value(Floor_Cords, cell2)
													cell = Temporary_Value(Letter_Next_To_1 + Floor_Cords[1:], cell)

													print (cell + ' Relieved Status: ', Has_Post_Been_Relieved)
													print ('')
													
													

													if cell in ROTATION_STARTERS or len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
														if '*' in cell and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) #Relieved_Post_Inner_List_Var - Determines what list we are using first inside of Relieved_Post; Starting with 0
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
															Relieved_Cords.append(Floor_Cords)
															Floor_OB1_Backup.append(Floor_Cords) #Replace to fit OB2
															print (Relieved_Post[Relieved_Post_Inner_List_Var])
															print ('if statement: 1')

															Stop_Var = 0
															#Can Delete

														

														elif cell in ROTATION_STARTERS and cell2 in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1  :
															sleep(0)
															if cell in EVERYTHING_ELSE: #If its a full BREAK.... kinda thinking this code is useless too because of the at risk if statement
																Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
																Relieved_Cords.append(Floor_Cords)
																print (Relieved_Post[Relieved_Post_Inner_List_Var])
																Relieved_Post_Inner_List_Var += 1
																Rotation_Complete += 1

																Stop_Var = 0
																#Can Delete
																
															elif cell not in EVERYTHING_ELSE: #Basically if its LAUNCH or has an at risk
																Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
																Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
																Relieved_Cords.append(Floor_Cords)
																Floor_OB1_Backup.append(Floor_Cords) #Replace to fit OB2
																print (Relieved_Post[Relieved_Post_Inner_List_Var])
																


																print ('This ROTATION is Ending EARLY')
																print (cell)
																print (cell2)
																print (Relieved_Post)
																print ('Cell2 Cord: ', Floor_Cords)
																print ('Cell Cord: ', Letter_Next_To_1 + Floor_Cords[1:])
																#sleep(9999)
																Relieved_Post_Inner_List_Var += 1
																Rotation_Complete += 1
																Stop_Var = 0
																print ('Adding Relieved Post Var + 1')
																print('')
																print('')
																print('')
																print('')
																print('')
																print('')
																print('')
																print('')
																print ('if statement: 2')
																break
															#xy10
														elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
															Relieved_Cords.append(Floor_Cords)
															print ('Ending this Rotation--')
															

															Relieved_Post_Inner_List_Var += 1
															Rotation_Complete += 1

															Stop_Var = 0
															#Can Delete
														elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in Relieved_Post[Relieved_Post_Inner_List_Var] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
															print (cell + 'Is Relieving ' + cell2)
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
															Relieved_Cords.append(Floor_Cords)
													
													

													Stop_Var += 1
													#Can Delete ^
													for cord in Floor_OB1:
														if cord in Relieved_Cords:
															Floor_OB1.remove(cord)
															Floor_OB1_Backup.append(cord)

													if Stop_Var == 80:
														print ('Chose no if statement')
														print ('Relieved so Far: ', Relieved_Post)
														print ('Current List We Are On: ', Relieved_Post_Inner_List_Var)
														print (Stop_Var)
														print ('Cords Left to be Relieved: ', Floor_OB1)
														print ('Cords Relieved and Removed from the List: ', Relieved_Cords)
														print ('Most likely this stopped because it remade the rotations. Sleeping Now.....')
														sleep(399999)




							'''
							-----CHECKING----- 

							Checking Rotations Have Gone Through.....


							Any code below here, checks, and double checks some more, just to make sure every post has been relieved, if it hasnt, it will redo
							the rotations for the next hour.
							'''


							print('')
							print ('')
							print ('---------------------------------')
							#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
							Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
							print ('Before Filter: ', Collapsed_List_Of_Relieved_Post)

							#We have to filter this list, to remove any post that are rotation starters besides the on that belongs to a Tier 1 List
							Junk_List = [] #List thats going to have post that we dont want
							for Post in Collapsed_List_Of_Relieved_Post:
								if Post in FREIGHT or Post not in OB1_Tier_1: 
									Junk_List.append(Post)

							print ('Junk List: ', Junk_List)

							#Taking away any post that was added from the Junk list; Post that aren't Tier 1 (Important Post)
							for Junk_Post in Junk_List:
								Collapsed_List_Of_Relieved_Post.remove(Junk_Post)


							Collapsed_List_Of_Relieved_Post = list(set(Collapsed_List_Of_Relieved_Post))
							#Removing Duplicate Post

							print ('After Filter: ', Collapsed_List_Of_Relieved_Post)
							print ('# of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post))
							print ('Tier 1 List: ', OB1_Tier_1)
							

							if len(Collapsed_List_Of_Relieved_Post) >= len(OB1_Tier_1):
								print ('')
								print ('----------------------------------------------------------------------------------------')
								print ('OB1 Post has been Relieved Properly')
								print ('Checked Previous Rotations: Success')
								print ('Re-Rotated: ', Remade_Post)

								#To End it All
								Previous_Check += 1
								sleep(0)
								#Checks += 1


								print ('Piecing Floor OB1 List Back Together..')
								Floor_OB1.extend(Relieved_Cords)
								Floor_OB1.extend(Relieved_Post)
								Floor_OB1.extend(Floor_OB1_Backup)
								Floor_OB1 = [item for index, item in enumerate(Floor_OB1) if item not in Floor_OB1[:index]]
								Floor_OB1 = [item for item in Floor_OB1 if isinstance(item, str)]
							elif len(Collapsed_List_Of_Relieved_Post) < len(OB1_Tier_1):
								print ('---------------------------------------------------------------------------------------------------------------------')
								print ('Redoing OB1 Floor Rotation')
								sleep(0)
								Remade_Post += 1

								#Redoing Floor Count Function to get an accurate Floor OB1 List:
								#FloorCellCount('F', OB1_OB2_Floor_Rows, 'OB3', int(Locate('OB1')[0][1:]))

								#Resetting The Main Variables and List -

								#This var will keep count of how many Post will start rotations on OB1
								Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates)

								#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
								Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]

								#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
								Relieved_Cords = []

								#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
								Relieved_Post_Inner_List_Var = 0 

								Stop_Var = 0
								#Can delete later, only used for testing




								#Recreating Floor OB1 list so the For Loop below can recreate the 2ns Column
								
								Floor_OB1.extend(Relieved_Cords)
								Floor_OB1.extend(Relieved_Post)
								#Floor_OB1.extend(Post_Pushing_Rotations_Cordinates) #This list is Cords with another letter in it and causing an error with the for loop below as well as increasing the length of the Floor_OB1 List
								Floor_OB1.extend(Floor_OB1_Backup)
								#print (Floor_OB1)
								#print ('Printing New Floor List')
								Floor_OB1 = [item for index, item in enumerate(Floor_OB1) if item not in Floor_OB1[:index]]
								Floor_OB1 = [item for item in Floor_OB1 if isinstance(item, str)]
								#At somepoint create a floor ob1 back up at the start of this code and just make floor ob1 = to it so its less confusing


								#Re-Creating the OB1 post for the 2nd Column
								Post_Tier = 0

								#Depending on how many times it chooses to redo the OB1 rotation, we just need this to remove the FLOAT 1 time
								if 'FLOAT' in OB1_Temporary_Post:
									OB1_Temporary_Post.remove(EVERYTHING_ELSE[0]) 
									#Removing FLOAT because a lead cord isnt in the FLOOT_OB1 list below. It messes up the new rotaions beind made as well when we redo them.


								random.shuffle(OB1_Temporary_Post)
								for i in Floor_OB1:
									Column_Number = ALPHABET.index(i[0]) + 1 + OB1_Hour_Post #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
									Cell = i #Not merging cells, so we dont relly need the cell var but we can keep it
									Create_Post(Cell, int(i[1:]), Column_Number, OB1_Temporary_Post[Post_Tier])
									OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
									ws[i].fill = OB1_Cell
									print ('Post Added: ', OB1_Temporary_Post[Post_Tier])
									print ('Current Cord: ', i)
									print (Post_Tier)
									print (Floor_OB1) #This is having G and H cords when it shouldnt have. Compared it to the original one above and see whats different
									print ('')
									Post_Tier += 1
									wb.save(File_Name)


								



										

						print ('Starting Cords: ', Post_Pushing_Rotations_Cordinates)
						print ('Relieved so Far: ', Relieved_Post)
						#print ('Current List We Are On: ', Relieved_Post_Inner_List_Var)
						print (Stop_Var)
						print ('Cords Left to be Relieved: ', Floor_OB1)
						print ('Cords Relieved and Removed from the List: ', Relieved_Cords)
						print ('OB1 Hour Post: ', OB1_Hour_Post)
						#Checks += 1 #This will help end the While Loop
						#sleep(99999)



					# or OB1 Rotation Check needs to happen here

					#OB1_Hour_Post += 1 #Might Delete this
					#Post_Tier = 0
					#if 'FLOAT' in OB1_Temporary_Post:
					#	OB1_Temporary_Post.remove(EVERYTHING_ELSE[0]) 
						#Removing FLOAT because a lead cord isnt in the FLOOT_OB1 list below. It messes up the new rotaions beind made as well when we redo them.














						#End ------------------------------------------------------------------


				

				if OB1_Hour_Post > 0:
					Post_Pushing_Rotations_Cordinates = [] 
					#Will be used for columns that have multiple people pushing the rotations such as briefing, breaks etc

					Letter_Next_To_1 = ALPHABET.index(Floor_OB1[0][0]) #Will represent the first Letter/Column originally in the Floor OB1 List
					Letter_Next_To_1 = ALPHABET[Letter_Next_To_1 + 1] #Will represent the second Letter/Column originally in the Floor OB1 List

					#Letter_Next_To_2 = ALPHABET.index(Letter_Next_To_1)
					#Letter_Next_To_2 = ALPHABET[Letter_Next_To_2 + 1]

					for i in Floor_OB1:
						cell = ws[i].value #Will check for the leading post during this rotation containing ** atrisk
						cell2 = ws[Letter_Next_To_1 + i[1:]].value 
						#cell3 = ws[Letter_Next_To_2 + i[1:]]


						if cell in ROTATION_STARTERS or '**' in cell:
							Post_Pushing_Rotations_Cordinates.append(i)
							#Grabbing every cell that starts the rotation on OB1's Double Rotation, and inserts it into a list. 
							#Update this

					if ws[Post_Pushing_Rotations_Cordinates[-1]].value == EVERYTHING_ELSE[0] and ws[Letter_Next_To_1 + Post_Pushing_Rotations_Cordinates[-1][1:]].value == EVERYTHING_ELSE[0]:
						#Both Values in the above if statement should be FLOAT and apart of a Lead line

						Post_Pushing_Rotations_Cordinates.pop()
						Floor_OB1.pop() 
						#This if statement is responsible for removing Lead cord from the OB1 Floor list and Post Pushing list so its easier to create a check and balance.
						#Lead cords DONT need to be in this list since all their post will mostly be Float and RADIO.
						#This only will take effect if both the post created are float that are back to back.
						



					#Checks and Balance Starts Here


					#This var will keep count of how many Post will start rotations on OB1
					Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates)

					#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
					Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]

					#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
					Relieved_Cords = []

					#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
					Relieved_Post_Inner_List_Var = 0 

					Stop_Var = 0
					#Can delete later, only used for testing
					
		
					while OB1_Hour_Post == 1:

						#Looping through all the starting post in the OB1 Post starter list.
						for Post_Starter in Post_Pushing_Rotations_Cordinates:

							if '**' in Temporary_Value(Post_Starter, ws[Post_Starter].value) or Temporary_Value(Post_Starter, ws[Post_Starter].value) in ROTATION_STARTERS:

								Rotation_Complete = 0 #Var used to help end the while Loop below. It'll become 1 at the end of every complete rotation

								while Rotation_Complete == 0: #Should be connected to the amount of post rotation starters there are in this hour rotation
									print ('Starting With: ', ws[Post_Starter].value)
									print ('Cord of Post Starter: ', Post_Starter)
									print ('Full List of Post Starters: ', Post_Pushing_Rotations_Cordinates)
									print ('Inner List #: ', Relieved_Post_Inner_List_Var)
									print ('Floor OB1 List: ', Floor_OB1)
									print ('Showing Current List Progress: ', Relieved_Post)
									print ('------')
									print ('')

									for Floor_Cords in Floor_OB1:
										cell = ws[Floor_Cords].value                 		 #Will be the value of the first cell in letter_1
										cell2 = ws[Letter_Next_To_1 + Floor_Cords[1:]].value #Will be the value of the 2nd cell directly next to the first cell

										#Used to end the while loop because the Relieved_Post_Inner_List_Var will represent a list alue that doesnt exist
										if Relieved_Post_Inner_List_Var == len(Relieved_Post):
											print ('Relieved Vars value has passed the amount of list we created. Ending Rottion early to force a re-rotate')
											Rotation_Complete += 1
											break


										if Floor_Cords == Post_Starter or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1: 

											
										


											if '*' in ws[Post_Starter].value: #Change this to make it more secure maybe?; If statement may be useless here honestly
												print ('Currently Relieved: ', Relieved_Post[Relieved_Post_Inner_List_Var])

												

												Has_Post_Been_Relieved = any(cell in sublist for sublist in Relieved_Post) 
												#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for

												cell = Temporary_Value(Floor_Cords, cell)
												cell2 = Temporary_Value(Letter_Next_To_1 + Floor_Cords[1:], cell2)

												print (cell + ' Relieved Status: ', Has_Post_Been_Relieved)
												print ('')
												#Can Delete later, will be used to see if its the reason if statement below is being skipped

												# and Has_Post_Been_Relieved == False | Took this out right before the or statement below
												if cell in ROTATION_STARTERS or len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													if '*' in cell and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) #Relieved_Post_Inner_List_Var - Determines what list we are using first inside of Relieved_Post; Starting with 0
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
														Relieved_Cords.append(Floor_Cords)
														Floor_OB1_Backup.append(Floor_Cords) #Replace to fit OB2
														print (Relieved_Post[Relieved_Post_Inner_List_Var])
														print ('if statement: 1')

														Stop_Var = 0
														#Can Delete

													elif cell in ROTATION_STARTERS and cell2 in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1  :
														sleep(0)
														if cell in EVERYTHING_ELSE: #If its a full BREAK.... kinda thinking this code is useless too because of the at risk if statement
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
															Relieved_Cords.append(Floor_Cords)
															print (Relieved_Post[Relieved_Post_Inner_List_Var])
															Relieved_Post_Inner_List_Var += 1
															Rotation_Complete += 1

															Stop_Var = 0
															#Can Delete
															
														elif cell not in EVERYTHING_ELSE: #Basically if its LAUNCH or has an at risk
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
															Relieved_Cords.append(Floor_Cords)
															Floor_OB1_Backup.append(Floor_Cords) #Replace to fit OB2
															print (Relieved_Post[Relieved_Post_Inner_List_Var])
															


															print ('This ROTATION is Ending EARLY')
															print (cell)
															print (cell2)
															print (Relieved_Post)
															print ('Cell Cord: ', Floor_Cords)
															print ('Cell2 Cord: ', Letter_Next_To_1 + Floor_Cords[1:])
															#sleep(9999)
															Relieved_Post_Inner_List_Var += 1
															Rotation_Complete += 1
															Stop_Var = 0
															print ('Adding Relieved Post Var + 1')
															print('')
															print('')
															print('')
															print('')
															print('')
															print('')
															print('')
															print('')
															print ('if statement: 2')
															break
														# cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 |   Before
													elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
														Relieved_Cords.append(Floor_Cords)
														print ('Ending this Rotation--')
														#Error is Here! Might Just add a Counter, if less than pass if more than, end rotation
														#Or every cord thats relieve, add to a list and 
														#Or we can have it officialy end when the last cord relieved has an atrisk
														#Or have it - post rotators

														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1

														Stop_Var = 0
														#Can Delete
													elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in Relieved_Post[Relieved_Post_Inner_List_Var] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
														print (cell + 'Is Relieving ' + cell2)
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
														Relieved_Cords.append(Floor_Cords)
												
												#elif cell in ROTATION_STARTERS and Has_Post_Been_Relieved == True and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1 :
												#	Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
												#	Stop_Var = 0
												#	print ('3: ')

												Stop_Var += 1
												#Can Delete ^
												for cord in Floor_OB1:
													if cord in Relieved_Cords:
														Floor_OB1.remove(cord)
														Floor_OB1_Backup.append(cord)

												if Stop_Var == 50:
													print ('Chose no if statement')
													print ('Relieved so Far: ', Relieved_Post)
													print ('Current List We Are On: ', Relieved_Post_Inner_List_Var)
													print (Stop_Var)
													print ('Cords Left to be Relieved: ', Floor_OB1)
													print ('Cords Relieved and Removed from the List: ', Relieved_Cords)
													print ('Most likely this stopped because it remade the rotations. Sleeping Now.....')
													sleep(399999)




						'''
						-----CHECKING----- 

						Checking Rotations Have Gone Through.....


						Any code below here, checks, and double checks some more, just to make sure every post has been relieved, if it hasnt, it will redo
						the rotations for the next hour.
						'''


						print('')
						print ('')
						print ('---------------------------------')
						#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
						Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
						print ('Before Filter: ', Collapsed_List_Of_Relieved_Post)

						#We have to filter this list, to remove any post that are rotation starters besides the on that belongs to a Tier 1 List
						Junk_List = [] #List thats going to have post that we dont want
						for Post in Collapsed_List_Of_Relieved_Post:
							if Post in FREIGHT or Post not in OB1_Tier_1: 
								Junk_List.append(Post)

						print ('Junk List: ', Junk_List)

						#Taking away any post that was added from the Junk list; Post that aren't Tier 1 (Important Post)
						for Junk_Post in Junk_List:
							Collapsed_List_Of_Relieved_Post.remove(Junk_Post)


						Collapsed_List_Of_Relieved_Post = list(set(Collapsed_List_Of_Relieved_Post))
						#Removing Duplicate Post

						print ('After Filter: ', Collapsed_List_Of_Relieved_Post)
						print ('# of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post))
						print ('Tier 1 List: ', OB1_Tier_1)
						

						if len(Collapsed_List_Of_Relieved_Post) >= len(OB1_Tier_1):
							print ('')
							print ('----------------------------------------------------------------------------------------')
							print ('OB1 Post has been Relieved Properly')
							print ('Re-Rotated: ', Remade_Post)

							#To End it All
							OB1_Hour_Post += 1
							Checks += 1
						elif len(Collapsed_List_Of_Relieved_Post) < len(OB1_Tier_1):
							print ('---------------------------------------------------------------------------------------------------------------------')
							print ('Redoing OB1 Floor Rotation')
							Remade_Post += 1

							#Redoing Floor Count Function to get an accurate Floor OB1 List:
							#FloorCellCount('F', OB1_OB2_Floor_Rows, 'OB3', int(Locate('OB1')[0][1:]))

							#Resetting The Main Variables and List -

							#This var will keep count of how many Post will start rotations on OB1
							Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates)

							#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
							Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]

							#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
							Relieved_Cords = []

							#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
							Relieved_Post_Inner_List_Var = 0 

							Stop_Var = 0
							#Can delete later, only used for testing




							#Recreating Floor OB1 list so the For Loop below can recreate the 2ns Column
							
							Floor_OB1.extend(Relieved_Cords)
							Floor_OB1.extend(Relieved_Post)
							Floor_OB1.extend(Post_Pushing_Rotations_Cordinates)
							Floor_OB1.extend(Floor_OB1_Backup)
							#print (Floor_OB1)
							#print ('Printing New Floor List')
							Floor_OB1 = [item for index, item in enumerate(Floor_OB1) if item not in Floor_OB1[:index]]
							Floor_OB1 = [item for item in Floor_OB1 if isinstance(item, str)]
							#print (Floor_OB1)
							#sleep(99999)
							#At somepoint create a floor ob1 back up at the start of this code and just make floor ob1 = to it so its less confusing


							#Re-Creating the OB1 post for the 2nd Column
							Post_Tier = 0

							#Depending on how many times it chooses to redo the OB1 rotation, we just need this to remove the FLOAT 1 time
							if 'FLOAT' in OB1_Temporary_Post:
								OB1_Temporary_Post.remove(EVERYTHING_ELSE[0]) 
								#Removing FLOAT because a lead cord isnt in the FLOOT_OB1 list below. It messes up the new rotaions beind made as well when we redo them.

							random.shuffle(OB1_Temporary_Post)
							for i in Floor_OB1:
								Column_Number = ALPHABET.index(i[0]) + 1 + OB1_Hour_Post #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
								Cell = i #Not merging cells, so we dont relly need the cell var but we can keep it
								Create_Post(Cell, int(i[1:]), Column_Number, OB1_Temporary_Post[Post_Tier])
								OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
								ws[i].fill = OB1_Cell
								print ('Post Added: ', OB1_Temporary_Post[Post_Tier])
								print (Post_Tier)
								Post_Tier += 1
								wb.save(File_Name)


							



									

					print ('Starting Cords: ', Post_Pushing_Rotations_Cordinates)
					print ('Relieved so Far: ', Relieved_Post)
					#print ('Current List We Are On: ', Relieved_Post_Inner_List_Var)
					print (Stop_Var)
					print ('Cords Left to be Relieved: ', Floor_OB1)
					print ('Cords Relieved and Removed from the List: ', Relieved_Cords)
					print ('OB1 Hour Post: ', OB1_Hour_Post)
					#Checks += 1 #This will help end the While Loop
					#sleep(99999)



				# or OB1 Rotation Check needs to happen here

				OB1_Hour_Post += 1 #Might Delete this
				Post_Tier = 0
				if 'FLOAT' in OB1_Temporary_Post:
					OB1_Temporary_Post.remove(EVERYTHING_ELSE[0]) 
					#Removing FLOAT because a lead cord isnt in the FLOOT_OB1 list below. It messes up the new rotations being made as well when we redo them.



		print ('Done (A)')
		
		index = ALPHABET.index(No_Float_Left_Behind[0][0])
		#This will pull the index of the letter for each Lead FLOAT Cell. We will then use it to loop and find any Lead FLOAT cell that was left with a value of None using the No_Float_Left_Behind list

		for letter in ALPHABET[index : (index + 2)]: #This is looping through only 2 letters because the No_Float_Left_Behind list should include 2 cords, which are next to each other
			Cell = letter + No_Float_Left_Behind[0][1:]

			if ws[Cell].value == None: #Allows any empty cell on a lead line to be given a FLOAT post.
				print ('This Lead Cord Is Missing a Float: ',Cell)
				Column_Number = ALPHABET.index(letter) + 1 
				Create_Post(Cell, int(Cell[1:]), Column_Number, EVERYTHING_ELSE[0])
				OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
				ws[i].fill = OB1_Cell
				print ('Post Added: ', EVERYTHING_ELSE[0])
				wb.save(File_Name)


def Upper_Floor_Rotation_Creation_OB2():
	global Floor_OB2
	Floor_OB2_Backup = [] #Will be used for OB2 Post being recreated, as a reference or copy
	#Creating a Back Ups for a for loop down below. Use find, then next. The very next time this list is brought up, thats what it will be used for

	OB2_Max_Length = len(OB2_Tier_1) + len(OB2_Tier_2)
	OB2_Min_Length = len(OB2_Tier_1)
	Upstairs_Min_Length = len(OB1_Tier_1) + len(OB2_Tier_1) #This should be the minimum for both floors to operate, currently 9

	OB2_Temporary_Post = [] #This list will be used to help determine what post will be added for the hour. Should be temporary and always changing
	OB2_Temporary_Post.extend(OB2_Tier_1)

	Post_Tier = 0
	#This var will help us keep track of what post is next in the list of OB1_Temporary_List to create down the column

	Remade_Post = 0
	#This var will count every time a column is remade due to every post not being relieved properly

	Previous_Check = 0
	#Will be used to represent if the previous column rotated properly or had any sort of hiccups. If its 1, it will ignore running the code again


	if len(OB1_OB2_Floor_Rows) > Upstairs_Min_Length:

		#Counting how many Leads are on OB2. This will determine if they recieve a FLOAT post or not.
		Lead_On_OB2 = 0 
		for a in Floor_OB2:
			for i in Leads_Upstairs:
				if a[1:] == i[1:]:
					Lead_On_OB2 += 1

		if (len(Floor_OB2) - Lead_On_OB2) > OB2_Min_Length:
			print ('Their are more Rows on OB2 than the Minimum, not counting the Lead. (A)')

			#This var will be used to detemine how many more post will be needed to fill up OB2 for the hour
			Amount_Of_Post_To_Add = (len(Floor_OB2) - Lead_On_OB2) - OB2_Min_Length


			#Adding the perfect amount needed so every empty cell is filled in for OB2
			for i in range(0, Amount_Of_Post_To_Add):
				OB2_Temporary_Post.append(OB2_Tier_2[i])

			#Shuffling the list for OB2 so its random everytime. Has to be done before FLOAT is added
			random.shuffle(OB2_Temporary_Post)


			if len(Floor_OB2) == len(OB2_Temporary_Post) + 1 and Lead_On_OB2 > 0: 
				OB2_Temporary_Post.append(EVERYTHING_ELSE[0])
				'''
				Im adding a FLOAT post to this list only because this IF STATEMENT is designed to fill up a column on the basis that their are enough
				ambassadors to have the floor running properly. This float post will be given to the lead
				'''


				#Placing the lead cord at the end of this list

				num = 0
				print ('Before: ', Floor_OB2)
				for Lead in Floor_OB2:
					for a in Leads_Upstairs:
						if Lead[1:] == a[1:]:
							Floor_OB2.remove(Lead)  # Remove the item from its current position
							Floor_OB2.append(Lead)  # Append the item to the end of the list
							num += 1
							break
						if num == 1:
							break
				print ('After: ', Floor_OB2)

			


			#Creating the OB2 post
			for i in Floor_OB2:
				Column_Number = ALPHABET.index(i[0]) + 1 #f its a value of 1 it redoes the 2nd Column
				index = ALPHABET.index(i[0]) + 1
				Cell = i + ':' + ALPHABET[index] + i[1:]
				Create_Post(Cell, int(i[1:]), Column_Number, OB2_Temporary_Post[Post_Tier])
				OB2_Cell = PatternFill(patternType = 'solid', fgColor = OB2)
				ws[i].fill = OB2_Cell
				print ('Post Added: ', OB2_Temporary_Post[Post_Tier])
				print (Post_Tier)
				Post_Tier += 1
				wb.save(File_Name)



			#Check and Balance for the previous Rotation will be created here | CB OB2
			if Floor_OB2[0][0] != 'F': #Possibly add another parameter here with the new Previous Check var and place this if statement into a while Loop
				print ('Stopping... Place the Parameter to check for brief or something. This shouldnt work until a 2nd Column is already made. After column G')
				print (Floor_OB2[0][0])
				sleep(0)
				wb.save(File_Name)

				while Previous_Check == 0:
					print ('Checking Routations from the previous line goes through properly')
					print (Floor_OB2[0][0])
					sleep(0)


					#Check and Balance for Previous Line goes here! 
					#Start----------------------------------------------------------------


					Post_Pushing_Rotations_Cordinates = [] 
					#Will be used for columns that have multiple people pushing the rotations such as briefing, breaks etc

					Letter_Next_To_1 = ALPHABET.index(Floor_OB2[0][0]) #Will represent the first Letter/Column originally in the Floor OB2 List
					Letter_Next_To_1 = ALPHABET[Letter_Next_To_1 - 1] #Will represent the second Letter/Column Left in original Floor OB2 List

					Letter_Next_To_2 = ALPHABET.index(Letter_Next_To_1)
					Letter_Next_To_2 = ALPHABET[Letter_Next_To_2 + 1]


					for i in Floor_OB2:
						cell2 = ws[i].value 
						cell = ws[Letter_Next_To_1 + i[1:]].value #Will check for the leading post during this rotation containing ** atrisk
						

						#Temporary_Value(Post_Starter, ws[Post_Starter].value)
						if Temporary_Value(Letter_Next_To_1 + i[1:], ws[Letter_Next_To_1 + i[1:]].value) in ROTATION_STARTERS or '**' in Temporary_Value(Letter_Next_To_1 + i[1:], ws[Letter_Next_To_1 + i[1:]].value):
							Post_Pushing_Rotations_Cordinates.append(Letter_Next_To_1 + i[1:])
							#Grabbing every cell that starts the rotation on OB2's Double Rotation, and inserts it into a list. 
							#Update this


					if ws[Post_Pushing_Rotations_Cordinates[-1]].value == EVERYTHING_ELSE[0] and ws[Letter_Next_To_1 + Post_Pushing_Rotations_Cordinates[-1][1:]].value == EVERYTHING_ELSE[0]:
						#Both Values in the above if statement should be FLOAT and apart of a Lead line

						Post_Pushing_Rotations_Cordinates.pop()
						Floor_OB2.pop() 
						#This if statement is responsible for removing Lead cord from the OB2 Floor list and Post Pushing list so its easier to create a check and balance.
						#Lead cords DONT need to be in this list since all their post will mostly be Float and RADIO.
						#This only will take effect if both the post created are float that are back to back.



					for the_cord in Post_Pushing_Rotations_Cordinates:
						Letter_1_Cord = the_cord
						Letter_2_Cord = Letter_Next_To_2 + the_cord[1:]
						#print (Temporary_Value(Letter_1_Cord, ws[Letter_1_Cord].value))
						#print (Temporary_Value(Letter_2_Cord, ws[Letter_2_Cord].value))
						#print(Post_Pushing_Rotations_Cordinates)
						#print ('')
						#print (Letter_1_Cord)
						#print (Letter_2_Cord)
						#sleep(0)

						#Basically if both Post are FLOAT it will remove it from the

						if Temporary_Value(Letter_1_Cord, ws[Letter_1_Cord].value) == EVERYTHING_ELSE[0] and Temporary_Value(Letter_2_Cord, ws[Letter_2_Cord].value) == EVERYTHING_ELSE[0]:
							Post_Pushing_Rotations_Cordinates.remove(the_cord)
							Floor_OB2.remove(Letter_2_Cord) 
							#This if statement is responsible for removing Lead cord from the OB2 Floor list and Post Pushing list so its easier to create a check and balance.
							#Lead cords DONT need to be in this list since all their post will mostly be Float and RADIO.
							#This only will take effect if both the post created are float that are back to back.

					#sleep(9999)

						



					#Checks and Balance Starts Here


					#This var will keep count of how many Post will start rotations on OB2
					Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates)

					#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
					Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]

					#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
					Relieved_Cords = []

					#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
					Relieved_Post_Inner_List_Var = 0 

					Stop_Var = 0
					#Can delete later, only used for testing
					
		
					while Previous_Check == 0:

						#Looping through all the starting post in the OB2 Post starter list.
						for Post_Starter in Post_Pushing_Rotations_Cordinates:

							if '**' in Temporary_Value(Post_Starter, ws[Post_Starter].value) or Temporary_Value(Post_Starter, ws[Post_Starter].value) in ROTATION_STARTERS:

								Rotation_Complete = 0 #Var used to help end the while Loop below. It'll become 1 at the end of every complete rotation

								while Rotation_Complete == 0: #Should be connected to the amount of post rotation starters there are in this hour rotation
									print ('Starting With: ', Temporary_Value(Post_Starter, ws[Post_Starter].value))
									print ('Cord of Post Starter: ', Post_Starter)
									print ('Full List of Post Starters: ', Post_Pushing_Rotations_Cordinates)
									print ('Inner List #: ', Relieved_Post_Inner_List_Var)
									print ('Floor OB2 List: ', Floor_OB2)
									print ('Showing Current List Progress: ', Relieved_Post)
									print ('------')
									print ('')
									sleep(0)

									for Floor_Cords in Floor_OB2:
										cell2 = ws[Floor_Cords].value                 		 #Will be the value of the Floor OB2 List Cells which is the most recent list in the excel and also printed out in the terminal when its ran
										cell = ws[Letter_Next_To_1 + Floor_Cords[1:]].value #Will be the value of the 2nd cell directly to the left of the most filled in Column
										#Used to end the while loop because the Relieved_Post_Inner_List_Var will represent a list alue that doesnt exist
										if Relieved_Post_Inner_List_Var == len(Relieved_Post):
											print ('Relieved Vars value has passed the amount of list we created. Ending Rottion early to force a re-rotate')
											Rotation_Complete += 1
											break


										if (Letter_Next_To_1 + Floor_Cords[1:]) == Post_Starter or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1: 

											
										


											if '*' in Temporary_Value(Post_Starter, ws[Post_Starter].value): #Change this to make it more secure maybe?; If statement may be useless here honestly
												print ('Currently Relieved: ', Relieved_Post[Relieved_Post_Inner_List_Var])

												

												#Has_Post_Been_Relieved = any(cell in sublist for sublist in Relieved_Post) 

												# Fixed: No generator/any() to avoid scoping errors
												source_cord_to_check = Letter_Left_To + Floor_Cords[1:]
												Has_Post_Been_Relieved = source_cord_to_check in Relieved_Cords
												#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for

												#Was Here Last!!!!!!!!!!!!!11 - Delete Later
												cell2 = Temporary_Value(Floor_Cords, cell2)
												cell = Temporary_Value(Letter_Next_To_1 + Floor_Cords[1:], cell)

												print (cell + ' Relieved Status: ', Has_Post_Been_Relieved)
												print ('')
												#Can Delete later, will be used to see if its the reason if statement below is being skipped

												# and Has_Post_Been_Relieved == False | Took this out right before the or statement below
												if cell in ROTATION_STARTERS or len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													if '*' in cell and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) #Relieved_Post_Inner_List_Var - Determines what list we are using first inside of Relieved_Post; Starting with 0
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
														Relieved_Cords.append(Floor_Cords)
														Floor_OB2_Backup.append(Floor_Cords) #Replace to fit OB2
														print (Relieved_Post[Relieved_Post_Inner_List_Var])
														print ('if statement: 1')

														Stop_Var = 0
														#Can Delete

													elif cell in ROTATION_STARTERS and cell2 in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1  :
														sleep(0)
														if cell in EVERYTHING_ELSE: #If its a full BREAK.... kinda thinking this code is useless too because of the at risk if statement
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
															Relieved_Cords.append(Floor_Cords)
															print (Relieved_Post[Relieved_Post_Inner_List_Var])
															Relieved_Post_Inner_List_Var += 1
															Rotation_Complete += 1

															Stop_Var = 0
															#Can Delete
															
														elif cell not in EVERYTHING_ELSE: #Basically if its LAUNCH or has an at risk
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
															Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
															Relieved_Cords.append(Floor_Cords)
															Floor_OB2_Backup.append(Floor_Cords) #Replace to fit OB2
															print (Relieved_Post[Relieved_Post_Inner_List_Var])
															


															print ('This ROTATION is Ending EARLY')
															print (cell)
															print (cell2)
															print (Relieved_Post)
															print ('Cell2 Cord: ', Floor_Cords)
															print ('Cell Cord: ', Letter_Next_To_1 + Floor_Cords[1:])
															#sleep(9999)
															Relieved_Post_Inner_List_Var += 1
															Rotation_Complete += 1
															Stop_Var = 0
															print ('Adding Relieved Post Var + 1')
															print('')
															print('')
															print('')
															print('')
															print('')
															print('')
															print('')
															print('')
															print ('if statement: 2')
															break
														# cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 |   Before
													elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
														Relieved_Cords.append(Floor_Cords)
														print ('Ending this Rotation--')
														#Error is Here! Might Just add a Counter, if less than pass if more than, end rotation
														#Or every cord thats relieve, add to a list and 
														#Or we can have it officialy end when the last cord relieved has an atrisk
														#Or have it - post rotators

														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1

														Stop_Var = 0
														#Can Delete
													elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in Relieved_Post[Relieved_Post_Inner_List_Var] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
														print (cell + 'Is Relieving ' + cell2)
														Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
														Relieved_Cords.append(Floor_Cords)
												
												#elif cell in ROTATION_STARTERS and Has_Post_Been_Relieved == True and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1 :
												#	Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
												#	Stop_Var = 0
												#	print ('3: ')

												Stop_Var += 1
												#Can Delete ^
												for cord in Floor_OB2:
													if cord in Relieved_Cords:
														Floor_OB2.remove(cord)
														Floor_OB2_Backup.append(cord)

												if Stop_Var == 50:
													print ('Chose no if statement')
													print ('Relieved so Far: ', Relieved_Post)
													print ('Current List We Are On: ', Relieved_Post_Inner_List_Var)
													print (Stop_Var)
													print ('Cords Left to be Relieved: ', Floor_OB2)
													print ('Cords Relieved and Removed from the List: ', Relieved_Cords)
													print ('Most likely this stopped because it remade the rotations. Sleeping Now.....')
													sleep(399999)




						'''
						-----CHECKING----- 

						Checking Rotations Have Gone Through.....


						Any code below here, checks, and double checks some more, just to make sure every post has been relieved, if it hasnt, it will redo
						the rotations for the next hour.
						'''


						print('')
						print ('')
						print ('---------------------------------')
						#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
						Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
						print ('Before Filter: ', Collapsed_List_Of_Relieved_Post)

						#We have to filter this list, to remove any post that are rotation starters besides the on that belongs to a Tier 1 List
						Junk_List = [] #List thats going to have post that we dont want
						for Post in Collapsed_List_Of_Relieved_Post:
							if Post in FREIGHT or Post not in OB2_Tier_1: 
								Junk_List.append(Post)

						print ('Junk List: ', Junk_List)

						#Taking away any post that was added from the Junk list; Post that aren't Tier 1 (Important Post)
						for Junk_Post in Junk_List:
							Collapsed_List_Of_Relieved_Post.remove(Junk_Post)


						Collapsed_List_Of_Relieved_Post = list(set(Collapsed_List_Of_Relieved_Post))
						#Removing Duplicate Post

						print ('After Filter: ', Collapsed_List_Of_Relieved_Post)
						print ('# of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post))
						print ('Tier 1 List: ', OB2_Tier_1)
						

						if len(Collapsed_List_Of_Relieved_Post) >= len(OB2_Tier_1):
							print ('')
							print ('----------------------------------------------------------------------------------------')
							print ('OB2 Post has been Relieved Properly')
							print ('Checked Previous Rotations: Success')
							print ('Re-Rotated: ', Remade_Post)

							#To End it All
							Previous_Check += 1
							sleep(0)
							#Checks += 1
						elif len(Collapsed_List_Of_Relieved_Post) < len(OB2_Tier_1):
							print ('---------------------------------------------------------------------------------------------------------------------')
							print ('Redoing OB2 Floor Rotation')
							sleep(0)
							Remade_Post += 1

							#Redoing Floor Count Function to get an accurate Floor OB2 List:
							#FloorCellCount('F', OB2_OB2_Floor_Rows, 'OB3', int(Locate('OB2')[0][1:]))

							#Resetting The Main Variables and List -

							#This var will keep count of how many Post will start rotations on OB2
							Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates)

							#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
							Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]

							#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
							Relieved_Cords = []

							#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
							Relieved_Post_Inner_List_Var = 0 

							Stop_Var = 0
							#Can delete later, only used for testing




							#Recreating Floor OB2 list so the For Loop below can recreate the 2ns Column
							
							Floor_OB2.extend(Relieved_Cords)
							Floor_OB2.extend(Relieved_Post)
							#Floor_OB2.extend(Post_Pushing_Rotations_Cordinates) #This list is Cords with another letter in it and causing an error with the for loop below as well as increasing the length of the Floor_OB2 List
							Floor_OB2.extend(Floor_OB2_Backup)
							#print (Floor_OB2)
							#print ('Printing New Floor List')
							Floor_OB2 = [item for index, item in enumerate(Floor_OB2) if item not in Floor_OB2[:index]]
							Floor_OB2 = [item for item in Floor_OB2 if isinstance(item, str)]
							#print (Floor_OB2)
							#sleep(99999)
							#At somepoint create a floor ob1 back up at the start of this code and just make floor ob1 = to it so its less confusing


							#Re-Creating the OB2 post for the 2nd Column
							Post_Tier = 0

							#Depending on how many times it chooses to redo the OB2 rotation, we just need this to remove the FLOAT 1 time
							if 'FLOAT' in OB2_Temporary_Post:
								OB2_Temporary_Post.remove(EVERYTHING_ELSE[0]) 
								#Removing FLOAT because a lead cord isnt in the FLOOT_OB2 list below. It messes up the new rotaions beind made as well when we redo them.

							random.shuffle(OB2_Temporary_Post)
							for i in Floor_OB2:
								Column_Number = ALPHABET.index(i[0]) + 1 #1 + OB2_Hour_Post #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
								Cell = i #Not merging cells, so we dont relly need the cell var but we can keep it
								Create_Post(Cell, int(i[1:]), Column_Number, OB2_Temporary_Post[Post_Tier])
								OB2_Cell = PatternFill(patternType = 'solid', fgColor = OB2)
								ws[i].fill = OB2_Cell
								print ('Post Added: ', OB2_Temporary_Post[Post_Tier])
								print ('Current Cord: ', i)
								print (Post_Tier)
								print (Floor_OB2) #This is having G and H cords when it shouldnt have. Compared it to the original one above and see whats different
								print ('')
								Post_Tier += 1
								wb.save(File_Name)


							



									

					print ('Starting Cords: ', Post_Pushing_Rotations_Cordinates)
					print ('Relieved so Far: ', Relieved_Post)
					#print ('Current List We Are On: ', Relieved_Post_Inner_List_Var)
					print (Stop_Var)
					print ('Cords Left to be Relieved: ', Floor_OB2)
					print ('Cords Relieved and Removed from the List: ', Relieved_Cords)
					#print ('OB2 Hour Post: ', OB2_Hour_Post)
					#Checks += 1 #This will help end the While Loop
					#sleep(99999)



				# or OB2 Rotation Check needs to happen here

				#OB2_Hour_Post += 1 #Might Delete this
				#Post_Tier = 0
				#if 'FLOAT' in OB2_Temporary_Post:
				#	OB2_Temporary_Post.remove(EVERYTHING_ELSE[0]) 
					#Removing FLOAT because a lead cord isnt in the FLOOT_OB2 list below. It messes up the new rotaions beind made as well when we redo them.














					#End ------------------------------------------------------------------









			


		elif (len(Floor_OB2) - Lead_On_OB2) < OB2_Min_Length:
			print ('OB2 Rows are Less than OB2 Min Length')
			sleep(99999)
		elif (len(Floor_OB2) - Lead_On_OB2) == OB2_Min_Length:
			print ('OB2 Rows are equal too OB2 Min Length')
			sleep(99999)

	print ('Done (A) - OB2')



#Upper_Floor_Rotation_Creation_OB1()
#Upper_Floor_Rotation_Creation_OB2()




#-------Creating Another Rotation Column

#Counting Empty Cells from OB1 - OB2
#OB1_OB2_Floor_Rows = []
#FloorCellCount('H', OB1_OB2_Floor_Rows, 'OB3', int(Locate('OB1')[0][1:]))
#Upper_Floor_Rotation_Creation_OB1()
#Upper_Floor_Rotation_Creation_OB2()

#Recently commented these out for the Funtion Below involving the letter 'H'



Borrowed_From_B1_Post = 0 #7/24/26 We can delete any existence of this var when the time is right. No longer use it.
'''
#-b1relief
This var was created for the purposes of the function below. Sometimes we will use a Post/Cord from B1 and this var will be used as a Global Var indicator to tell the program that we used a B1 Post in the previous hour.
'''




#-------------------------Creating a OB1 and OB2 Check and Balance Function....
def Upper_Floor_Rotation_Creation_Both_Floors():
	global Floor_OB1, Floor_OB2, OB1_OB2_Floor_Rows, All_Shifts_Rows, Borrowed_From_B1_Post, Full_Time_Closers_Row, Overall_OB1_Post, Overall_OB2_Post
	#OB1_OB2_Floor_Rows is a list made up of OB1 and OB2 Open Cells

	OB1_OB2_Floor_Rows_Backup = []
	#Will be used to Trouble Shoot the original broken and out of order list later down the line when we go to check previous rotations

	OB1_OB2_Floor_Rows_Relief_List = []
	#This list will contain cords that are only important for relief when checking rotations. A bit different than the regular OB1_OB2 original list. I will count 'NONE' cells as well as cells that have a BREAK. Should not be used for Post creation since it may overwrite some breaks
	#Do not use for Post Creation

	'''
	breakerninefix = Name of the comments where all the breaker's leaving at 9:30 or half an hour rotation, solution code will be. 
	Delete it all and the code will work just fine. Code below
	'''
	if OB1_OB2_Floor_Rows[0][0] == 'P':
		print ('Before: ', OB1_OB2_Floor_Rows)
		OB1_Row_Num = int(Locate('OB1')[0][1:])
		Upstairs_Breaker_Shifts_Cords = []
		for a in OB1_OB2_Floor_Rows:
			for i in Locate('6pm-10:30pm'):
				if i[1:] == a[1:]:
					Upstairs_Breaker_Shifts_Cords.append(a)
		Upstairs_Breaker_Shifts_Cords = set(Upstairs_Breaker_Shifts_Cords)
		OB1_OB2_Floor_Rows[:] = [item for item in OB1_OB2_Floor_Rows if item not in Upstairs_Breaker_Shifts_Cords]
		print ('After: ' , OB1_OB2_Floor_Rows)
		print ('Remove Breaker Shifts or Place them in a seperate list since they leave at 9:30')
		print (Locate('6pm-10:30pm'))
		wb.save(File_Name)
		#sleep(99999)

	if Borrowed_From_B1_Post != 0:
		OB1_OB2_Floor_Rows_Relief_List.append(OB1_OB2_Floor_Rows[0][0] + Borrowed_From_B1_Post)
		ROTATION_STARTERS.extend(Mix_Of_Both_B1_Tiers)

		Borrowed_From_B1_Post = 0 
		#Resetting the Global Var back to 0 so it doesn't stop here when the function is used again..
		#Can delete this var completely at some point. No longer being used. Actually being replaced
		#-b1relief

	Using_Celebrate_From_B1 = 0
	# We will use this var later to let the code know not to remove CAB 2 if we borrowed from Celebrate on B1, or are using the PREP Post instead.



	OB3_Free_Rows_Being_Used = []
	#Will be a list of OB3 Rows that are being used to cover a post; Will be used to determine when to add the cord to the relief list... later on in the code


	Free_Leads_Upstairs = []
	#This list will contain the cords of leads that're free during this rotation hour. It will give us a more accurate count of free rows/GEA's during the post creation process

	OB3_Free_Rows = []
	#Rows on OB3 that we can use to make sure rotations go through soothly. They will contain cords to cells that havent been filled in yet

	What_Should_We_Do = []
	#A semi important list. Meant to be re-written later in the code but will be used to determine contingency plans for low staffing upstairs.
	#Placed this here so its no longer known as a local var when called later regardless. Its mainly brought up in an if statement

	OB1_Max_Length = len(OB1_Tier_1) + len(OB1_Tier_2)
	OB1_Min_Length = len(OB1_Tier_1)
	OB2_Max_Length = len(OB2_Tier_1) + len(OB2_Tier_2)
	OB2_Min_Length = len(OB2_Tier_1)
	Upstairs_Min_Length = len(OB1_Tier_1) + len(OB2_Tier_1) #This should be the minimum for both floors to operate, currently 13

	OB1_Temporary_Post = [] #This list will be used to help determine what post will be added for the hour. Should be temporary and always changing
	OB1_Temporary_Post.extend(OB1_Tier_1)
	OB2_Temporary_Post = [] #This list will be used to help determine what post will be added for the hour. Should be temporary and always changing
	OB2_Temporary_Post.extend(OB2_Tier_1)


	Fixed_Repeated_Post = []
	#This will be a list that catches every repeated post, their cord, and just captures it for later just so that I can see what was fixed or "what could have been"

	'''
	This list will contain column letters mainly towards the end where it will reverse the way we create post. 
	It starts out as OB1 Post on OB1 and the same for OB2, but the letters in the list below will have OB2 Post on OB1 and OB1 Post on OB2 just to mix it up a bit
	'''
	Column_Letters_To_Reverse_Post = ['N', 'P', 'R'] #Add L and add a feature that deletes it from the list based on randomness
	#Add L and we will make it random if we want it to start at Column L or N




	
	
	if OB1_OB2_Floor_Rows[0][0] in Column_Letters_To_Reverse_Post:
		
		#Switching Floor cordinates in order to reverse Floor Post to have OB2 Post on OB1 vice verse
		Floor_OB1, Floor_OB2 = Floor_OB2, Floor_OB1
		
		#sleep(99999)
		#newly added
	





	'''
	Calculations - 
	The Math needed in order to make sure rotations standards can be met will be included here. Minimum for OB1 is 4 and Minimum for OB2 is 5
	The 'Standards' variables below for OB1 and OB2 are only values and not attached to specific floors. It will be used to determine how many post we will need for the hour and how many rows they will take

	'''

	for i in Leads_Upstairs:
		Cord = OB1_OB2_Floor_Rows[0][0] + i[1:]
		if Temporary_Value(OB1_OB2_Floor_Rows[0][0] + i[1:], ws[Cord].value) == None and Temporary_Value(OB1_OB2_Floor_Rows[0][0] + i[1:], ws[Cord].value) != EVERYTHING_ELSE[3]:
			Free_Leads_Upstairs.append(Cord)
	#This will add any lead lines to a list that we can use later to help cover some post if needed. It only checks for lead lines that are free and empty


	#Newly Added Code: Removing Breaks from this.... This sort of Code should have been added. May have to come to this code later and copy all the Cords before hand. Not sure

	# "Keep everything that is NOT a break value"
	OB1_OB2_Floor_Rows = [
	    coord for coord in OB1_OB2_Floor_Rows 
	    if Temporary_Value(coord, ws[coord].value) != EVERYTHING_ELSE[3]
	]
	
	FlLength = 0
	#We will use this var to help me know which if statement below got chosen later on in the code. Can Choose to delete later if you want

	ResultOfTheMath = len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs)
	#Delete later, will need to this value to see where a fix is needed

	SaveData = OB1_OB2_Floor_Rows
	SaveData2 = Free_Leads_Upstairs
	#Delete Later

	# We need to account for if Freight is involved.
	FreightPostNow = 0
	for i in OB1_OB2_Floor_Rows:
		if Temporary_Value(i, ws[i].value) in FREIGHT:
			FreightPostNow += 1


	#Adding FLOAT Post to Lead Lines in the 1st If Statement | Add Later: Can possibly subtract 1 too depening on if its around the time freight is happening
	if (len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs)) - FreightPostNow >= Upstairs_Min_Length:
		FlLength += 1
		#This if Statement allows Leads to Keep their FLOAT

		print ('Leads GET FLOAT POST')

		print ('')
		print ('Before')
		print (len(OB1_OB2_Floor_Rows), len(Floor_OB1), len(Floor_OB2))
		print(OB1_OB2_Floor_Rows)
		print(Floor_OB1)
		print(Floor_OB2)

		#Basically removing a lead line from the Free Empty Cells and reserving a FLOAT for them
		for i in OB1_OB2_Floor_Rows:
			for a in Free_Leads_Upstairs:
				if i[1:] == a[1:]:
					print (i)
					#print ('Create Float Post here')
					Column_Number = ALPHABET.index(i[0]) + 1 #f its a value of 1 it redoes the 2nd Column
					index = ALPHABET.index(i[0]) + 1
					Cell = i + ':' + ALPHABET[index] + i[1:]
					Create_Post(Cell, int(i[1:]), Column_Number, EVERYTHING_ELSE[0])
					FLOAT_Cell = PatternFill(patternType = 'solid', fgColor = FLOAT)
					ws[i].fill = FLOAT_Cell
					print ('FLOAT Added For Added: ', i)
					OB1_OB2_Floor_Rows.remove(i)

				for b in Floor_OB1:
					if b[1:] == a[1:]:
						Floor_OB1.remove(b)

				for c in Floor_OB2:
					if c[1:] == a[1:]:
						Floor_OB2.remove(c)


		print ('')
		print ('After')
		print (len(OB1_OB2_Floor_Rows), len(Floor_OB1), len(Floor_OB2))
		print(OB1_OB2_Floor_Rows)
		print(Floor_OB1)
		print(Floor_OB2)
		print ('----------------------------------------') #Kinda have no idea why I added the +1......?
	elif (len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs)) - FreightPostNow < Upstairs_Min_Length: #was <= before, add to the if statement, on the left side a count of free cells including the ones on ob3, The subtraction condition should change depending on how many people are coming back from break, because if we have people coming back from break then we have rotation starters. Will also include a letter in this function.... maybe
		FlLength += 2
		print ('Currently (-+): Low on rows for this rotation hour, NO FLOAT POST for Leads; Revise the code here below... 2 solutions that can be used')
		print ('Rows Avialiable While Counting Lead Rows: ', len(OB1_OB2_Floor_Rows))
		print ('Rows Avialiable Before Counting Lead Rows: ', len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs))
		print ('Rows Avialiable Before Counting Lead Rows and Removing Freight: ', (len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs)) - FreightPostNow)
		print ('Len of OB1_OB2 Floor Rows have to be More than this: ', Upstairs_Min_Length + 1)
		#sleep(99999)

		print (len(OB1_Tier_1))
		print (len(OB2_Tier_1))
		print ('Upstairs Minimum Length: ', Upstairs_Min_Length)
		print (len(OB1_OB2_Floor_Rows))
		print (OB1_OB2_Floor_Rows)
		print (len(Free_Leads_Upstairs))

		#if OB1_OB2_Floor_Rows[0][0] == 'N':
		#	print ("stopped here")
		#	sleep(99999)

		#READ ME: At some point come back and revisit the math on these if statements below. Re-Run the program with different Ambassasdors/Lead numbers
		#READ ME: That +2 allows both leads to recieve float post... can make it dependent upon a variable thats tied to the length of OB1_OB2_FLOOR_ROWS
		#Update... Will be replacing +2 with a var that has a more accurate depiction of how many leads are actually available upstairs. If one is on break
		# then 1 lead out of 2 leads on OB1 and OB2 are actually available instead of a total of 2.

		#PURPOSE OF IF Statements: They determine whether leads get float post or ambassador like lines as well as choosing what cells get a post
		#Must include Freight Post into this math. Will add FreightPostNow var
		#Delete Large Code Below

		'''
		if len(OB1_OB2_Floor_Rows) - FreightPostNow == Upstairs_Min_Length + len(Free_Leads_Upstairs) and len(Free_Leads_Upstairs) == 1:
			#-reference

			#This will be a list of random numbers and will determine which solution we will go with
			What_Should_We_Do = [1, 1, 1, 1, 3, 3, 3, 3, 3, 3, 3, 2, 2] #Change this percentage later on
			# 53% Chance it will Delete a Freight, 15% Chance it will Choose a B1 Post, 31% Chance it will have a Lead cover a post

			#Uncomment the above list when done ^

			#Test List
			#What_Should_We_Do = [1]


			#This var will help us end the while loop once a solution is chosen
			Solution_Chose = 0


			while Solution_Chose == 0:
				random.shuffle(What_Should_We_Do)
				#Randomly Shuffles to the solution is random.

				if What_Should_We_Do[0] == 1:
					#SOLUTION 1: 1st If Statement Gives Leads a FLOAT POST, and the 2nd If Statement Gives LEADS Regular Post

					if len(OB1_OB2_Floor_Rows) > Upstairs_Min_Length:

						print ('Before: ', OB1_OB2_Floor_Rows)
						print (len(OB1_OB2_Floor_Rows))
						print (Free_Leads_Upstairs)

						for i in OB1_OB2_Floor_Rows:
							for a in Free_Leads_Upstairs:
								if i[1:] == a[1:] and len(OB1_OB2_Floor_Rows) > Upstairs_Min_Length:
									print (i)
									print ('we will remove ', i)


									#print ('Create Float Post here')
									Column_Number = ALPHABET.index(i[0]) + 1 #f its a value of 1 it redoes the 2nd Column
									index = ALPHABET.index(i[0]) + 1
									Cell = i + ':' + ALPHABET[index] + i[1:]
									Create_Post(Cell, int(i[1:]), Column_Number, EVERYTHING_ELSE[0])
									FLOAT_Cell = PatternFill(patternType = 'solid', fgColor = FLOAT)
									ws[i].fill = FLOAT_Cell
									print ('FLOAT Added For Added: ', i)
									OB1_OB2_Floor_Rows.remove(i)

									for b in Floor_OB1:
										if b[1:] == a[1:]:
											Floor_OB1.remove(b)

									for c in Floor_OB2:
										if c[1:] == a[1:]:
											Floor_OB2.remove(c)


						print ('After: ', OB1_OB2_Floor_Rows)
						print (len(OB1_OB2_Floor_Rows))
						Solution_Chose += 1
					elif len(OB1_OB2_Floor_Rows) <= Upstairs_Min_Length:

						print ('Letting Lead Get a Post for this Rotation Hour...')
						Solution_Chose += 1

				elif What_Should_We_Do[0] == 2: #Uses a B1 Post to cover the hour FIX/ERROR
					print ('Choosing a B1 Line to Cover this Hour') 
					print ('Choosing this.... Line 4489')
					print ('Left Off Here!!! Find out why the code is choosing this solution to grab a B1 post when it can easily turn a lead float into a post instead...')
					#sleep(99999) #1

					EndRange = Locate('OB1')
					B1_Post_To_Borrow = []
					#This will be a list for all Tier 2 B1 Post Cords that we can use for this hour to help cover OB1 and OB2

					for i in range(5, int(EndRange[0][1:]) + 1):
						B1_Post_Name = ws[OB1_OB2_Floor_Rows[0][0] + str(i)].value
						B1_Post_Cord = OB1_OB2_Floor_Rows[0][0] + str(i)
						print (B1_Post_Name)

						if B1_Post_Name in B1_Tier_3:
							B1_Post_To_Borrow.append(B1_Post_Cord)

					random.shuffle(B1_Post_To_Borrow)

					print ('')
					print (B1_Post_To_Borrow)

					if len(B1_Post_To_Borrow) > 0:
						#Basically removing a lead line from the Free Empty Cells and reserving a FLOAT for them
						for i in OB1_OB2_Floor_Rows:
							for a in Free_Leads_Upstairs:
								if i[1:] == a[1:]:
									print (i)
									#print ('Create Float Post here')
									Column_Number = ALPHABET.index(i[0]) + 1 #f its a value of 1 it redoes the 2nd Column
									index = ALPHABET.index(i[0]) + 1
									Cell = i + ':' + ALPHABET[index] + i[1:]
									Create_Post(Cell, int(i[1:]), Column_Number, EVERYTHING_ELSE[0])
									FLOAT_Cell = PatternFill(patternType = 'solid', fgColor = FLOAT)
									ws[i].fill = FLOAT_Cell
									print ('FLOAT Added For Added: ', i)
									OB1_OB2_Floor_Rows.remove(i)

								for b in Floor_OB1:
									if b[1:] == a[1:]:
										Floor_OB1.remove(b)

								for c in Floor_OB2:
									if c[1:] == a[1:]:
										Floor_OB2.remove(c)




						# Clear content and style of the top-left cell of the merged range
						cell = ws[B1_Post_To_Borrow[0]]
						cell.value = None  # Clear content
						cell.fill = PatternFill(fill_type=None)  # Clear fill (background color)
						wb.save(File_Name)

						# Unmerge cells
						#Re-Merging the Left Half
						#Letter_To_Merge = ALPHABET.index(Freight_Cell_To_Unmerge[0][0]) + 1
						#Letter_To_Merge = ALPHABET[Letter_To_Merge] #Which is one Letter to the right of the original Freight Cord
						Right_Cell_Letter_index = ALPHABET.index(B1_Post_To_Borrow[0][0]) + 1
						Right_Cell = ALPHABET[Right_Cell_Letter_index]
						Cells_To_Unmerge = B1_Post_To_Borrow[0] + ':' + Right_Cell + B1_Post_To_Borrow[0][1:]
						ws.unmerge_cells(Cells_To_Unmerge)  # Replace with the actual merged range
						print ('Cells Unmerged: ', Cells_To_Unmerge)
						wb.save(File_Name)


						#sleep(99999)
						
						OB1_OB2_Floor_Rows.append(B1_Post_To_Borrow[0])
						OB1_OB2_Floor_Rows_Relief_List.append(B1_Post_To_Borrow[0])

						Mix_Of_Both_B1_Tiers.append('PREP') 
						
						#Lazy Coding... Can delete later when we re-do B1. 
						#For some odd reason Prep isn't recognized as a Post Starter on my OB1/PB2 Check and Balance Algorithm when ever we use this 
						#if statement to grab a B1 Post... so I wrote this in..

						#Mix_Of_Both_B1_Tiers isnt used anymore anyway when the program even makes it this far and also it isnt a global var/list so it will reset after this 
						#function is over with
						
						
						ROTATION_STARTERS.append('B1 BOH FREIGHT')
						#Adding B1 FREIGHT to the main ROTATION STARTER list bugs B1 for some reason..... Will fix when I recode B1
						ROTATION_STARTERS.extend(Mix_Of_Both_B1_Tiers)
						#Since we are borrowing a post from B1, any B1 post before that OB1/OB2 post on B1 should be considered as a Post Rotater for the time being

						Borrowed_From_B1_Post = B1_Post_To_Borrow[0][1:]
						#This will be a Global Var that tells the program, "hey we borrowed a post from B1, please remember we did that!" for the next time hour this function is used, so it knows to check B1's cordinates to complete its relief check
						Solution_Chose += 1
						#-b1relief Main Solution If Statement
					 
						

				elif What_Should_We_Do[0] == 3 and FreightPostNow > 3:

					#SOLUTION 3: Will Delete Freight to Make Room if their is no OB3. Only if we have more than the neccassary amount of Freight People Needed, which is more than 3.
					#--------------------------------------
					#Below is deleting one FREIGHT Rider and creating a space for a new post.... will make it cut frieght in half... 
					#if STATEMENT Condition should only be if we need an extra row and if it matches up with the same rotation hour as FREIGHT!!! this is important. May have to gather all Freight cords in 1 list for this

					print ('Lets also tap into 1 Freight post and possibly break it in half, preferbly a rider. We will also set up if statements on what to add like if its 1 needed, just borrow a freight person, if 2 borrow a freight person and someone from ob3.')

					
					Freight_Rows = []
					#This will be a list of Rows that have Freight post on OB1 and OB2. We will use this to hold extra Freight post cords that can be converted into regular post in case empty cells are low
					
					for i in All_Shifts_Rows:
						if Temporary_Value(OB1_OB2_Floor_Rows[0][0] + i[1:], ws[OB1_OB2_Floor_Rows[0][0] + i[1:]].value) == None:
							pass
						elif 'RIDER' in Temporary_Value(OB1_OB2_Floor_Rows[0][0] + i[1:], ws[OB1_OB2_Floor_Rows[0][0] + i[1:]].value):
							Freight_Rows.append(OB1_OB2_Floor_Rows[0][0] + i[1:])

					print (Freight_Rows)
					random.shuffle(Freight_Rows)

					#Making sure Freight Starting Column properly matches up first, so everything runs smooth when Erasing it
					if len(Freight_Rows) > 1 and any(item[0] == OB1_OB2_Floor_Rows[0][0] for item in Freight_Cord): #ERASES ENTIRE FREIGHT ROW 
						Last_Freight_Column = ALPHABET.index(Freight_Rows[0][0]) + 3 
						Last_Freight_Column = ALPHABET[Last_Freight_Column] #Should be the last letter in a Freight Post; Will be used to properly unmerge the Freight cells
						Freight_Cell_To_Unmerge = Freight_Rows[0] + ':' + Last_Freight_Column + Freight_Rows[0][1:]
						ws.unmerge_cells(Freight_Cell_To_Unmerge)
						for row in ws[Freight_Cell_To_Unmerge]:
						    for cell in row:
						        # Clear text
						        cell.value = None
						        
						        # Clear color
						        cell.fill = PatternFill(start_color=None, end_color=None, fill_type=None)


						print ('')
						print ('Choosing 1 FREIGHT RIDER')
						print ('Before')
						print (len(OB1_OB2_Floor_Rows), len(Floor_OB1), len(Floor_OB2))
						print(OB1_OB2_Floor_Rows)
						print(Floor_OB1)
						print(Floor_OB2)


						#Basically removing a lead line from the Free Empty Cells and reserving a FLOAT for them
						for i in OB1_OB2_Floor_Rows:
							for a in Free_Leads_Upstairs:
								if i[1:] == a[1:]:
									print (i)
									#print ('Create Float Post here')
									Column_Number = ALPHABET.index(i[0]) + 1 #f its a value of 1 it redoes the 2nd Column
									index = ALPHABET.index(i[0]) + 1
									Cell = i + ':' + ALPHABET[index] + i[1:]
									Create_Post(Cell, int(i[1:]), Column_Number, EVERYTHING_ELSE[0])
									FLOAT_Cell = PatternFill(patternType = 'solid', fgColor = FLOAT)
									ws[i].fill = FLOAT_Cell
									print ('FLOAT Added For Added: ', i)
									OB1_OB2_Floor_Rows.remove(i)

								for b in Floor_OB1:
									if b[1:] == a[1:]:
										Floor_OB1.remove(b)

								for c in Floor_OB2:
									if c[1:] == a[1:]:
										Floor_OB2.remove(c)


						print ('')
						print ('After')
						print (len(OB1_OB2_Floor_Rows), len(Floor_OB1), len(Floor_OB2))
						print(OB1_OB2_Floor_Rows)
						print(Floor_OB1)
						print(Floor_OB2)
						print ('----------------------------------------')
						









						OB1_OB2_Floor_Rows.append(Freight_Rows[0])
						Solution_Chose += 1 #ERASES ENTIRE FREIGHT POST

					elif len(Freight_Rows) > 1: #CUTS FREIGHT POST IN HALF

						#Pulling the full Fright Cords from start to finish into a List Below
						Freight_Cell_To_Unmerge = []

						#Freight_Cord is also the original list created in the main FREIGHT Function with all Freight Coordinates from start to finish. Referenced again here

						for i in Freight_Cord[1:]: #Assuming All Freight Post after the first one has a double number coordinate: Ex: F10:G10, J17:L17
							if len(i) == 7:
								if 'RIDER' in ws[i[:3]].value:
									Freight_Cell_To_Unmerge.append(i)
							elif len(i) != 7:
								print (i)
								print ('Pausing....Freight Cordinate is Off..... Length: ', len(i) == 7)
								sleep(99999)


						#Fully Unmerging The Cell
						ws.unmerge_cells(Freight_Cell_To_Unmerge[0]) #ERROR: We need to fix out how UNMERGE FRIGHT, Back to Original Idea and Break it in Half. I see loop through all merge list and match with value or save the original cords in freight list if havent already and match with cord number here
						for row in ws[Freight_Cell_To_Unmerge[0]]:
						    for cell in row:
						        # Clear text
						        cell.value = None
						        
						        # Clear color
						        cell.fill = PatternFill(start_color=None, end_color=None, fill_type=None)

						#Re-Merging the Left Half
						Letter_To_Merge = ALPHABET.index(Freight_Cell_To_Unmerge[0][0]) + 1
						Letter_To_Merge = ALPHABET[Letter_To_Merge] #Which is one Letter to the right of the original Freight Cord

						ws.merge_cells(Freight_Cell_To_Unmerge[0][:4] + Letter_To_Merge + Freight_Cell_To_Unmerge[0][1:3])

						#Re-Writing in the Text for the FREIGHT CELL that was SPLIT in Half
						ws[Freight_Cell_To_Unmerge[0][:3]].value = 'FREIGHT'

						#Filling In the Color
						FREIGHT_HALF_CELL = PatternFill(patternType = 'solid', fgColor = FREIGHT_COLOR)
						ws[Freight_Cell_To_Unmerge[0][:3]].fill = FREIGHT_HALF_CELL

						#Re-Saving the File
						wb.save(File_Name)
						print ('ERASED 2nd Half of FREIGHT Post... Right side should be empty and used for a seperate POST')
						

						OB1_OB2_Floor_Rows.append(OB1_OB2_Floor_Rows[0][0] + Freight_Cell_To_Unmerge[0][1:3])

						#print (Freight_Cell_To_Unmerge[0][:3])
						#print (OB1_OB2_Floor_Rows)
						#print ('Check the above accuracys... should also remove it from Freight_Cord... Tested this once.... Seems to be working fine!')
						Solution_Chose += 1 



					else:
						print ('Freight Isnt Active During this Hour... Will Continue While Loop')



			wb.save(File_Name) #If only 1 More Row is Only Needed #Deleted the +2 next to Upstairs var.... still dont know why I originally added that 
		elif len(OB1_OB2_Floor_Rows) - FreightPostNow > Upstairs_Min_Length + len(Free_Leads_Upstairs):
			print ('')
			print ('More than Enough Rows Needed.. Including Lead Rows.. Sleeping')
			print ('- Updated Numbers Below - ')
			print ('Rows Avialiable While Counting Lead Rows: ', len(OB1_OB2_Floor_Rows))
			print ('Rows Avialiable Before Counting Lead Rows: ', len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs))
			print ('Rows Avialiable Before Counting Lead Rows and Removing Freight: ', (len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs)) - FreightPostNow)
			print ('Len of OB1_OB2 Floor Rows have to be More than this: ', Upstairs_Min_Length + 1)
			#sleep(99999)

			print (len(OB1_Tier_1))
			print (len(OB2_Tier_1))
			print ('Upstairs Minimum Length: ', Upstairs_Min_Length)
			print (len(OB1_OB2_Floor_Rows))
			print (OB1_OB2_Floor_Rows)
			print ('Avaliable Leads During This Rotation: ', len(Free_Leads_Upstairs))
			sleep(99999)

		elif len(OB1_OB2_Floor_Rows) - FreightPostNow < Upstairs_Min_Length: 
			#-reference
			
			#This if Statement will capture the program if their isnt enough open cells on OB1 and OB2 to fill for every upstairs Tier1 Post
			#and we are low on ambassadors to the point where no one can start a rotation
			
			
			if len(OB1_OB2_Floor_Rows) - FreightPostNow == Upstairs_Min_Length - 1:
				print ('This if Statement will build out the rotations if its just 1 less than the Upstairs Minimum... ')
				print ('We will have a post starter rotation from ob1 or ob2 go down to B1')
				print ('Before: ', OB1_OB2_Floor_Rows)
				print ('')

				#This will be a list of random numbers and will determine which solution we will go with
				What_Should_We_Do = [2, 2, 3] 
				#Will Add 2 Later On.. Only adding a 66% Chance for B1 Solution due to not having more Freight than 3 atm


				
				#No Reason for '1' to be in the above list. '1' Basically tells the program to give a Lead a Post... it already does that if it made it this far...


				#This var will help us end the while loop once a solution is chosen
				Solution_Chose = 0

				print ('')
				print ('Before: ', OB1_OB2_Floor_Rows)
				wb.save(File_Name)

				while Solution_Chose == 0 and len(What_Should_We_Do) != 0:
					random.shuffle(What_Should_We_Do)
					#Randomly Shuffles to the solution is random.

					print ('READ ME: Dont forget that when we circle back to this... to add the .remove of the number its shuffled to, in every solution if it isnt avialiable')
					

					#if What_Should_We_Do[0] == 1: #Gives Leads a Post instead of FLOAT
					#	print ('Letting Lead Get a Post for this Rotation Hour...')
					#	Solution_Chose += 1
						#SOLUTION 1: DO NOTHING! and Ignore the Code below and allow the Lead to Cover this Post

						#REASON FOR COMMENTTING: This solution is useless because we already turn a lead's rotation hour into post

					if What_Should_We_Do[0] == 2: #Choosing B1 Tier 2 Post or B1 Post Started to Cover OB1/OB2 Post
						print ('Choosing a B1 Line to Cover this Hour') 
						#sleep(99999)

						EndRange = Locate('OB1')
						B1_Post_To_Borrow = []
						#This will be a list for all Tier 2 B1 Post Cords that we can use for this hour to help cover OB1 and OB2

						for i in range(5, int(EndRange[0][1:]) + 1):
							B1_Post_Name = ws[OB1_OB2_Floor_Rows[0][0] + str(i)].value
							B1_Post_Cord = OB1_OB2_Floor_Rows[0][0] + str(i)
							print (B1_Post_Name)

							if B1_Post_Name in B1_Tier_3:
								B1_Post_To_Borrow.append(B1_Post_Cord)

						random.shuffle(B1_Post_To_Borrow)

						print ('')
						print (B1_Post_To_Borrow)

						if len(B1_Post_To_Borrow) > 0:

							# Clear content and style of the top-left cell of the merged range
							cell = ws[B1_Post_To_Borrow[0]]
							cell.value = None  # Clear content
							cell.fill = PatternFill(fill_type=None)  # Clear fill (background color)
							wb.save(File_Name)

							# Unmerge cells
							#Re-Merging the Left Half
							#Letter_To_Merge = ALPHABET.index(Freight_Cell_To_Unmerge[0][0]) + 1
							#Letter_To_Merge = ALPHABET[Letter_To_Merge] #Which is one Letter to the right of the original Freight Cord
							Right_Cell_Letter_index = ALPHABET.index(B1_Post_To_Borrow[0][0]) + 1
							Right_Cell = ALPHABET[Right_Cell_Letter_index]
							Cells_To_Unmerge = B1_Post_To_Borrow[0] + ':' + Right_Cell + B1_Post_To_Borrow[0][1:]
							ws.unmerge_cells(Cells_To_Unmerge)  # Replace with the actual merged range
							print ('Cells Unmerged: ', Cells_To_Unmerge)
							wb.save(File_Name)


							#sleep(99999)
							
							OB1_OB2_Floor_Rows.append(B1_Post_To_Borrow[0])
							OB1_OB2_Floor_Rows_Relief_List.append(B1_Post_To_Borrow[0])

							Mix_Of_Both_B1_Tiers.append('PREP') 
							
							#Lazy Coding... Can delete later when we re-do B1. 
							#For some odd reason Prep isn't recognized as a Post Starter on my OB1/PB2 Check and Balance Algorithm when ever we use this 
							#if statement to grab a B1 Post... so I wrote this in..

							#Mix_Of_Both_B1_Tiers isnt used anymore anyway when the program even makes it this far and also it isnt a global var/list so it will reset after this 
							#function is over with
							
							
							ROTATION_STARTERS.append('B1 BOH FREIGHT')
							#Adding B1 FREIGHT to the main ROTATION STARTER list bugs B1 for some reason..... Will fix when I recode B1
							ROTATION_STARTERS.extend(Mix_Of_Both_B1_Tiers)
							#Since we are borrowing a post from B1, any B1 post before that OB1/OB2 post on B1 should be considered as a Post Rotater for the time being

							Borrowed_From_B1_Post = B1_Post_To_Borrow[0][1:]
							#This will be a Global Var that tells the program, "hey we borrowed a post from B1, please remember we did that!" for the next time hour this function is used, so it knows to check B1's cordinates to complete its relief check
							Solution_Chose += 1
							#-b1relief
						elif len(B1_Post_To_Borrow) == 0:
							print ('No Post to Pull from on B1... Trying Next Solution....')

							#Adding this 'remove' so it speeds up the process of elimination with the list of solutions..
							What_Should_We_Do.remove(What_Should_We_Do[0])
							

						



							
							

					elif What_Should_We_Do[0] == 3 and FreightPostNow > 3: #Will Delete a Freight Post

						#SOLUTION 3: Will Delete Freight to Make Room if their is no OB3
						#--------------------------------------
						#Below is deleting one FREIGHT Rider and creating a space for a new post.... will make it cut frieght in half... 
						#if STATEMENT Condition should only be if we need an extra row and if it matches up with the same rotation hour as FREIGHT!!! this is important. May have to gather all Freight cords in 1 list for this

						print ('Lets also tap into 1 Freight post and possibly break it in half, preferbly a rider. We will also set up if statements on what to add like if its 1 needed, just borrow a freight person, if 2 borrow a freight person and someone from ob3.')

						
						Freight_Rows = []
						#This will be a list of Rows that have Freight post on OB1 and OB2. We will use this to hold extra Freight post cords that can be converted into regular post in case empty cells are low
						
						for i in All_Shifts_Rows:
							if Temporary_Value(OB1_OB2_Floor_Rows[0][0] + i[1:], ws[OB1_OB2_Floor_Rows[0][0] + i[1:]].value) == None:
								pass
							elif 'RIDER' in Temporary_Value(OB1_OB2_Floor_Rows[0][0] + i[1:], ws[OB1_OB2_Floor_Rows[0][0] + i[1:]].value):
								Freight_Rows.append(OB1_OB2_Floor_Rows[0][0] + i[1:])

						print (Freight_Rows)
						random.shuffle(Freight_Rows)

						#Making sure Freight Starting Column properly matches up first, so everything runs smooth when Erasing it
						if len(Freight_Rows) > 1 and any(item[0] == OB1_OB2_Floor_Rows[0][0] for item in Freight_Cord): #ERASES ENTIRE FREIGHT ROW 
							Last_Freight_Column = ALPHABET.index(Freight_Rows[0][0]) + 3 
							Last_Freight_Column = ALPHABET[Last_Freight_Column] #Should be the last letter in a Freight Post; Will be used to properly unmerge the Freight cells
							Freight_Cell_To_Unmerge = Freight_Rows[0] + ':' + Last_Freight_Column + Freight_Rows[0][1:]
							ws.unmerge_cells(Freight_Cell_To_Unmerge)
							for row in ws[Freight_Cell_To_Unmerge]:
							    for cell in row:
							        # Clear text
							        cell.value = None
							        
							        # Clear color
							        cell.fill = PatternFill(start_color=None, end_color=None, fill_type=None)


							print ('')
							print ('Choosing 1 FREIGHT RIDER')
							print ('Before')
							print (len(OB1_OB2_Floor_Rows), len(Floor_OB1), len(Floor_OB2))
							print(OB1_OB2_Floor_Rows)
							print(Floor_OB1)
							print(Floor_OB2)


							#Basically removing a lead line from the Free Empty Cells and reserving a FLOAT for them if enough ambassadors can cover the floor now that a FREIGHT post has been taken away
							if len(OB1_OB2_Floor_Rows) >= Upstairs_Min_Length:
								for i in OB1_OB2_Floor_Rows:
									for a in Free_Leads_Upstairs:
										if i[1:] == a[1:]:
											print (i)
											#print ('Create Float Post here')
											Column_Number = ALPHABET.index(i[0]) + 1 #f its a value of 1 it redoes the 2nd Column
											index = ALPHABET.index(i[0]) + 1
											Cell = i + ':' + ALPHABET[index] + i[1:]
											Create_Post(Cell, int(i[1:]), Column_Number, EVERYTHING_ELSE[0])
											FLOAT_Cell = PatternFill(patternType = 'solid', fgColor = FLOAT)
											ws[i].fill = FLOAT_Cell
											print ('FLOAT Added For Added: ', i)
											OB1_OB2_Floor_Rows.remove(i)

										for b in Floor_OB1:
											if b[1:] == a[1:]:
												Floor_OB1.remove(b)

										for c in Floor_OB2:
											if c[1:] == a[1:]:
												Floor_OB2.remove(c)


							print ('')
							print ('After')
							print (len(OB1_OB2_Floor_Rows), len(Floor_OB1), len(Floor_OB2))
							print(OB1_OB2_Floor_Rows)
							print(Floor_OB1)
							print(Floor_OB2)
							print ('----------------------------------------')
							









							OB1_OB2_Floor_Rows.append(Freight_Rows[0])
							Solution_Chose += 1 #ERASES ENTIRE FREIGHT POST

						elif len(Freight_Rows) > 1: #CUTS FREIGHT POST IN HALF

							#Pulling the full Fright Cords from start to finish into a List Below
							Freight_Cell_To_Unmerge = []

							#Freight_Cord is also the original list created in the main FREIGHT Function with all Freight Coordinates from start to finish. Referenced again here

							for i in Freight_Cord[1:]: #Assuming All Freight Post after the first one has a double number coordinate: Ex: F10:G10, J17:L17
								if len(i) == 7:
									if 'RIDER' in ws[i[:3]].value:
										Freight_Cell_To_Unmerge.append(i)
								elif len(i) != 7:
									print (i)
									print ('Pausing....Freight Cordinate is Off..... Length: ', len(i) == 7)
									sleep(99999)


							#Fully Unmerging The Cell
							ws.unmerge_cells(Freight_Cell_To_Unmerge[0]) #ERROR: We need to fix out how UNMERGE FRIGHT, Back to Original Idea and Break it in Half. I see loop through all merge list and match with value or save the original cords in freight list if havent already and match with cord number here
							for row in ws[Freight_Cell_To_Unmerge[0]]:
							    for cell in row:
							        # Clear text
							        cell.value = None
							        
							        # Clear color
							        cell.fill = PatternFill(start_color=None, end_color=None, fill_type=None)

							#Re-Merging the Left Half
							Letter_To_Merge = ALPHABET.index(Freight_Cell_To_Unmerge[0][0]) + 1
							Letter_To_Merge = ALPHABET[Letter_To_Merge] #Which is one Letter to the right of the original Freight Cord

							ws.merge_cells(Freight_Cell_To_Unmerge[0][:4] + Letter_To_Merge + Freight_Cell_To_Unmerge[0][1:3])

							#Re-Writing in the Text for the FREIGHT CELL that was SPLIT in Half
							ws[Freight_Cell_To_Unmerge[0][:3]].value = 'FREIGHT'

							#Filling In the Color
							FREIGHT_HALF_CELL = PatternFill(patternType = 'solid', fgColor = FREIGHT_COLOR)
							ws[Freight_Cell_To_Unmerge[0][:3]].fill = FREIGHT_HALF_CELL

							#Re-Saving the File
							wb.save(File_Name)
							print ('ERASED 2nd Half of FREIGHT Post... Right side should be empty and used for a seperate POST')
							

							OB1_OB2_Floor_Rows.append(OB1_OB2_Floor_Rows[0][0] + Freight_Cell_To_Unmerge[0][1:3])

							#print (Freight_Cell_To_Unmerge[0][:3])
							#print (OB1_OB2_Floor_Rows)
							#print ('Check the above accuracys... should also remove it from Freight_Cord... Tested this once.... Seems to be working fine!')
							Solution_Chose += 1 



						else:
							print ('Freight Isnt Active During this Hour... Will Continue While Loop')

							#Adding this 'remove' so it speeds up the process of elimination with the list of solutions..
							What_Should_We_Do.remove(What_Should_We_Do[0])
					else:
						print ('Delete the sleep when we figure out whats the hold up (Line 6387)... Stops Here, and Chose This: ', What_Should_We_Do[0])

						#Adding this 'remove' so it speeds up the process of elimination with the list of solutions..
						What_Should_We_Do.remove(What_Should_We_Do[0])
						sleep(0)


				print ('')
				print ('After: ', OB1_OB2_Floor_Rows)

				

			elif len(OB1_OB2_Floor_Rows) - FreightPostNow < Upstairs_Min_Length - 1:
				print ('Less rows Avialiable than planned for.... find a new solution')
				#for this solution, we will delete tr1, delete a freight if its 4 freights, and possible borrow a not important post from B1
				sleep(99999)
			#sleep(99999)
			#Length += 3 Can delete later. Think this is supposed to be FlLength var
		'''
		#Delete Large Code Above
		# --- STAFFING HIERARCHY & LEAD PROTECTION LOGIC ---
	
		# Initial Count: GSAs (non-leads) vs Mandatory Tier 1 Requirements
		GSAs_Available = len(OB1_OB2_Floor_Rows) - len(Free_Leads_Upstairs)
		Tier1_Requirement = len(OB1_Tier_1) + len(OB2_Tier_1)
		
		# Shortage > 0 means we are missing bodies for Tier 1
		#Shortage = Tier1_Requirement - GSAs_Available

		Leads_To_Work = []
		Leads_To_Float = []


		#Freight Accuracy is needed for the 2nd if Statement. May possibly delete it out of the first if statement overall at some point.
		current_col = OB1_OB2_Floor_Rows[0][0]
		Freight_Rider_Cords = []
		Freight_Census = 0
		
		# Census: Find existing Freight upstairs for this hour
		for r in range(int(Locate('OB1')[0][1:]) + 1, int(Locate('OB3')[0][1:])):
			cv = Temporary_Value(current_col + str(r), ws[current_col + str(r)].value)
			if cv and 'FREIGHT' in str(cv):
				Freight_Census += 1
				if 'RIDER' in str(cv): Freight_Rider_Cords.append(current_col + str(r))

		# Shortage > 0 means we are missing bodies for Tier 1
		Shortage = Tier1_Requirement - (GSAs_Available - Freight_Census)
		print (f'Tier 1 Requirement for Upstairs is {Tier1_Requirement} and out of the GSAs that are Free which would be {GSAs_Available}, a few are assigned to Freight: {Freight_Census}')



		# 1. PRIORITY 1: SACRIFICE FREIGHT (MIN 3 POLICY)
		if Shortage > 0:
			current_col = OB1_OB2_Floor_Rows[0][0]
			Freight_Rider_Cords = []
			Freight_Census = 0
			
			# Census: Find existing Freight upstairs for this hour
			for r in range(int(Locate('OB1')[0][1:]) + 1, int(Locate('OB3')[0][1:])):
				cv = Temporary_Value(current_col + str(r), ws[current_col + str(r)].value)
				if cv and 'FREIGHT' in str(cv):
					Freight_Census += 1
					if 'RIDER' in str(cv): Freight_Rider_Cords.append(current_col + str(r))

			print(f"HIERARCHY 1: Freight Census is {Freight_Census}. Shortage is {Shortage}.")
			
			# Draft Riders only if we have a luxury (Census > 3)
			while Shortage > 0 and Freight_Census > 3 and len(Freight_Rider_Cords) > 0:
				target_f = Freight_Rider_Cords.pop()
				# Clean cell for new assignment
				for mr in list(ws.merged_cells.ranges):
					if target_f in mr: ws.unmerge_cells(str(mr))
				ws[target_f].value = None
				ws[target_f].fill = PatternFill(fill_type=None)
				
				OB1_OB2_Floor_Rows.append(target_f) # Add body to pool
				Freight_Census -= 1
				Shortage -= 1
				print(f"ACTION: Sacrificed Freight Rider at {target_f}. Shortage now {Shortage}.")

		# 2. PRIORITY 2: SURGICAL LEAD DRAFTING
		if Shortage <= 0:
			# Staffing resolved by GSAs or Freight: All Leads FLOAT
			Leads_To_Float = Free_Leads_Upstairs
			Leads_To_Work = []
			print(f"HIERARCHY 2: Staffing Healthy. All {len(Leads_To_Float)} Leads will FLOAT.")
			print (Shortage)
		else:
			 # Draft exactly the number of Leads we need to cover the shortage
			Leads_To_Work = Free_Leads_Upstairs[:Shortage]  # These stay in the pool
			Leads_To_Float = Free_Leads_Upstairs[Shortage:] # These get FLOAT now

		# ACTION: Physically assign FLOAT to the floaters and kick them out of the pool
		for lead_cord in Leads_To_Float:
			col_idx = ALPHABET.index(lead_cord[0]) + 1
		    # Check if we need a 1-cell or 2-cell merge (Column R is the end of shift)
			m_limit = 2 if lead_cord[0] == 'R' else 1
			Cell_Range = lead_cord + ':' + ALPHABET[col_idx + m_limit - 1] + lead_cord[1:]
		    
		    # 1. Assign the FLOAT text and color
			Create_Post(Cell_Range, int(lead_cord[1:]), col_idx, EVERYTHING_ELSE[0])
			ws[lead_cord].fill = PatternFill(patternType='solid', fgColor=FLOAT)
		    
		    # 2. REMOVE them from the pool so the script doesn't give them a Tier 1 job later
			if lead_cord in OB1_OB2_Floor_Rows:
				OB1_OB2_Floor_Rows.remove(lead_cord)

		# Now Shortage is effectively 0 because the drafted leads are still in OB1_OB2_Floor_Rows
		Shortage = 0


		'''
		May Delete Later
		else:
			# Still short: Draft only enough Leads to cover the gap
			if Shortage <= len(Free_Leads_Upstairs):
				print ('Current Shortage Before: ', Shortage)
				Leads_To_Work = Free_Leads_Upstairs[-Shortage:] # Take from bottom
				Leads_To_Float = Free_Leads_Upstairs[:-Shortage] # Rest stay floating
				print(f"HIERARCHY 2: Drafting {len(Leads_To_Work)} Lead(s). {len(Leads_To_Float)} stay Floating.")
				Shortage = 0 
			else:
				# Critical: All leads drafted, shortage remains
				Leads_To_Work = Free_Leads_Upstairs
				Leads_To_Float = []
				Shortage = Shortage - len(Leads_To_Work)
				print(f"HIERARCHY 2: ALL Leads drafted. Still missing {Shortage} people.")
			'''



		# --- HIERARCHY 3: B1 PREP-SWAP & CELEBRATE SACRIFICE ---
		if Shortage > 0: #REMINDER: Can use this to host more upstairs post if we have more preps avaliable
			
			current_col = OB1_OB2_Floor_Rows[0][0]
			col_idx = ALPHABET.index(current_col)
			next_col = ALPHABET[col_idx + 1]
			b1_start_row = 5
			b1_end_row = int(Locate('OB1')[0][1:])
			print(f"HIERARCHY 3: Detecting B1 Preps for hour block {current_col}/{next_col}...")
			
			prep_first_staff = [] # List of rows doing [PREP, BREAK]
			break_first_staff = [] # List of rows doing [BREAK, PREP]

			# 1. PHYSICAL SCAN: Identify all staff involved in PREP rotations this hour
			for r in range(b1_start_row, b1_end_row):
				cell1_val = Temporary_Value(current_col + str(r), ws[current_col + str(r)].value)
				cell2_val = Temporary_Value(next_col + str(r), ws[next_col + str(r)].value)

				if cell1_val == 'PREP' and cell2_val == 'BREAK':
					prep_first_staff.append(str(r))
				elif cell1_val == 'BREAK' and cell2_val == 'PREP':
					break_first_staff.append(str(r))

			# --- REBALANCE AND SELECTION LOGIC ---
			if len(prep_first_staff) >= 1 and len(break_first_staff) >= 1:
				# Scenario: We have both. No swap needed.
				# Row A covers 1st half, Row B covers 2nd half.
				all_prep_rows = [prep_first_staff[0], break_first_staff[0]]

			elif len(prep_first_staff) >= 2:
				# Scenario: All Prep-First. Swap the 2nd person to Break-First.
				target_row = prep_first_staff[1]
				row_int = int(target_row)
				
				# Physically swap them on the Excel sheet: [BREAK -> PREP]
				Create_Post(current_col + target_row, row_int, col_idx + 1, 'BREAK')
				ws[current_col + target_row].fill = PatternFill(patternType='solid', fgColor=BREAK)
				Create_Post(next_col + target_row, row_int, col_idx + 2, 'PREP')
				ws[next_col + target_row].fill = PatternFill(patternType='solid', fgColor=PREP)
				
				# Now person 0 covers 1st half, person 1 covers 2nd half.
				all_prep_rows = [prep_first_staff[0], target_row]
				print(f"HIERARCHY 3: Rebalanced Row {target_row} to BREAK-First for coverage.")

			elif len(break_first_staff) >= 2:
				# Scenario: All Break-First. Swap the 2nd person to Prep-First.
				target_row = break_first_staff[1]
				row_int = int(target_row)

				# Physically swap them on the Excel sheet: [PREP -> BREAK]
				Create_Post(current_col + target_row, row_int, col_idx + 1, 'PREP')
				ws[current_col + target_row].fill = PatternFill(patternType='solid', fgColor=PREP)
				Create_Post(next_col + target_row, row_int, col_idx + 2, 'BREAK')
				ws[next_col + target_row].fill = PatternFill(patternType='solid', fgColor=BREAK)

				# Now person 1 covers 1st half, person 0 covers 2nd half.
				all_prep_rows = [target_row, break_first_staff[0]]
				print(f"HIERARCHY 3: Rebalanced Row {target_row} to PREP-First for coverage.")
			else:
				# Low staff fallback
				all_prep_rows = prep_first_staff + break_first_staff
				if len(all_prep_rows) >= 2:
					all_prep_rows = all_prep_rows[0:2]

			print(f"DEBUG: Found {len(prep_first_staff)} Prep-First and {len(break_first_staff)} Break-First staffers.")

			# 2. COVERAGE CHECK & SCHEDULE FLIP
			# We need at least 2 people in the Prep pool to sacrifice a Celebrate staffer safely.
			if len(all_prep_rows) >= 2:
				# If everyone is taking a break at the same time (e.g., everyone is Prep-First)
				# we flip one person to Break-First to ensure 100% floor coverage.
				if len(prep_first_staff) >= 2 and len(break_first_staff) == 0:
					target_row = prep_first_staff[0]
					print(f"LOGISTICS: Flipping Row {target_row} to [BREAK -> PREP] for coverage.")
					
					# Update Excel: Change Prep to Break
					Create_Post(current_col + target_row, int(target_row), col_idx + 1, 'BREAK')
					ws[current_col + target_row].fill = PatternFill(patternType='solid', fgColor=BREAK)
					
					# Update Excel: Change Break to Prep
					Create_Post(next_col + target_row, int(target_row), col_idx + 2, 'PREP')
					ws[next_col + target_row].fill = PatternFill(patternType='solid', fgColor=PREP)
					
					# Adjust our local tracking lists
					break_first_staff.append(prep_first_staff.pop(0))

				# 3. RE-ASSIGN PREP STAFF TO COVER CELEBRATE DUTIES
				# (Since the Prep people are staying on B1, they will cover the Celebrate post)
				for r_str in all_prep_rows:
					# If they aren't on break, their duty is now 'CELEBRATE'
					if ws[current_col + r_str].value == 'PREP':
						ws[current_col + r_str].value = 'CELEB'
					if ws[next_col + r_str].value == 'PREP':
						ws[next_col + r_str].value = 'CELEB'

				# 4. SACRIFICE THE ORIGINAL CELEBRATE PERSON
				b1_celebrates = Locate('CELEBRATE')
				if b1_celebrates:
					for cel_cord in b1_celebrates:
						cel_row = cel_cord[1:]
						# Find the person who was ALREADY celebrate (not the ones we just re-labeled)
						if cel_cord[0] == current_col and int(cel_row) < b1_end_row and cel_row not in all_prep_rows:
							print(f"ACTION: Sacrificing original Celebrate at {cel_cord}. Moving to upstairs pool.")
							
							# Clean the cell for its new upstairs assignment
							for mr in list(ws.merged_cells.ranges):
								if cel_cord in mr: ws.unmerge_cells(str(mr))
							
							ws[cel_cord].value = None
							ws[cel_cord].fill = PatternFill(fill_type=None)
							
							# Add body to upstairs pool
							ROTATION_STARTERS.extend(Mix_Of_Both_B1_Tiers) #Important
							ROTATION_STARTERS.append('CELEB') #Important
							OB1_OB2_Floor_Rows.append(cel_cord)
							OB1_OB2_Floor_Rows_Relief_List.append(cel_cord)
							Celeb_B1_Cord = cel_cord #Important
							# This var is important to be used later to let our code know not to delete CAB 2.
							
							# Logic balancing for floor split
							if int(cel_row) < int(Locate('OB2')[0][1:]):
								Floor_OB1.append(cel_cord)
							else:
								Floor_OB2.append(cel_cord)
							
							Shortage -= 1
							break # One sacrifice per shortage unit
				print ('')
				print ('See if it worked.. Line 6735')
				print ('After Adding B1 Cord: ', OB1_OB2_Floor_Rows)
				print(f" Found {len(prep_first_staff)} Prep-First and {len(break_first_staff)} Break-First staffers.") #b1 preps is not accurately reading all the 30 min preps for the hour in the current letter column and for the column for the right.
				print (Shortage)
				Using_Celebrate_From_B1 += 1
				wb.save(File_Name)
				#sleep(10)

			Celebs_Made = Locate('CELEB')
            # Extract all column letters
			cols = [''.join(filter(str.isalpha, c)) for c in Celebs_Made]

            # If total length != unique count, at least 2 share a column letter
			has_matching_column = len(cols) != len(set(cols))
			if has_matching_column == True:
				print ('Before: ', Celebs_Made)
				print ('Inspect')
				print("HIERARCHY 3: Column collision detected for CELEB posts. Rebalancing...")
				
				# Group coords by column letter to find exactly which ones are clashing
				# Example: {'L': ['L30', 'L42'], 'M': ['M30']}
				from collections import Counter
				col_counts = Counter(cols)
				
				for cl_col, count in col_counts.items():
					if count > 1:
						# Find all coordinates in the clashing column
						clashing_cords = [c for c in Celebs_Made if c.startswith(cl_col)]
						
						# We only need to fix one person to resolve the clash
						cord_to_fix = clashing_cords[0] 
						row_str = ''.join(filter(str.isdigit, cord_to_fix))
						col_idx = ALPHABET.index(cl_col)
						
						# Determine the neighbor (Partner cell in the same hour block)
						# Blocks are J/K, L/M, N/O. If index is odd (J, L, N), partner is to the right.
						if col_idx % 2 == 1: 
							neighbor_idx = col_idx + 1
						else: # If index is even (K, M, O), partner is to the left
							neighbor_idx = col_idx - 1
							
						neighbor_letter = ALPHABET[neighbor_idx]
						neighbor_cord = neighbor_letter + row_str
						
						# Perform the swap only if the neighbor is a BREAK
						if Temporary_Value(neighbor_cord, ws[neighbor_cord].value) == 'BREAK':
							# 1. Move CELEB to the neighbor block
							ws[neighbor_cord].value = 'CELEB'
							ws[neighbor_cord].fill = PatternFill(patternType='solid', fgColor=B1)
							
							# 2. Move BREAK to the original block
							ws[cord_to_fix].value = 'BREAK'
							ws[cord_to_fix].fill = PatternFill(patternType='solid', fgColor=BREAK)
							
							print(f"REBALANCE SUCCESS: Swapped Row {row_str} from {cord_to_fix} to {neighbor_cord} to cover both half-hours.")
							wb.save(File_Name)
							
						else:
							print(f"REBALANCE WARNING: Row {row_str} neighbor {neighbor_cord} was not a BREAK. Could not swap.")
				
				# After fixing, refresh the list for the next check
				Celebs_Made = Locate('CELEB')
				#print ('After: ', Celebs_Made)
			else:
				print(f"NOTICE: Only {len(all_prep_rows)} Preps found. Hierarchy 3 requires 2. Skipping sacrifice.")
				print("NOTICE: Not enough Preps on B1 to execute swap. Skipping Hierarchy 3.")
				print ('See if it worked.. Line 6739')
				#sleep(99999)

				# 4. FINALIZATION: Assign FLOAT to protected leads and remove from pool

				for lead_cord in Leads_To_Float:
					col_idx = ALPHABET.index(lead_cord[0]) + 1
					m_limit = 2 if lead_cord[0] == 'R' else 1
					Cell_Range = lead_cord + ':' + ALPHABET[col_idx + m_limit - 1] + lead_cord[1:]
					
					Create_Post(Cell_Range, int(lead_cord[1:]), col_idx, EVERYTHING_ELSE[0])
					ws[lead_cord].fill = PatternFill(patternType='solid', fgColor=FLOAT)
					
					# Remove from coordinate lists to prevent double-assignment
					if lead_cord in OB1_OB2_Floor_Rows: OB1_OB2_Floor_Rows.remove(lead_cord)
					if lead_cord in Floor_OB1: Floor_OB1.remove(lead_cord)
					if lead_cord in Floor_OB2: Floor_OB2.remove(lead_cord)

				# 5. ZERO-WASTE MATH: All remaining staff fill the split
				Total_Spots = len(OB1_OB2_Floor_Rows)
				OB1_Standard = Total_Spots // 2
				OB2_Standard = Total_Spots - OB1_Standard

				# Safety check for OB1 Floor minimum
				if OB1_Standard < OB1_Min_Length and Total_Spots >= OB1_Min_Length:
					OB1_Standard = OB1_Min_Length
					OB2_Standard = Total_Spots - OB1_Standard

				print(f"Final Count for Posts: OB1({OB1_Standard}) | OB2({OB2_Standard})")

				


				#sleep(99999)
		
				



	#Removing Cords that are for Freight out of the main list. This can be a possible problem later... not sure
	OB1_OB2_Floor_Rows = [
	    coord for coord in OB1_OB2_Floor_Rows 
	    if Temporary_Value(coord, ws[coord].value) not in FREIGHT
	]
	



	Total_Rows_Both_Floors = len(OB1_OB2_Floor_Rows) 
	#The total of free cells for both OB1 and OB2; The Code directly below evenly devices rows for both floors due to the division

	OB2_Standard = OB2_Min_Length
	#print ('OB2 Free Rows', OB2_Standard) - Can delete later
	#Should be 9 or more. It will be what ever value is left over after we subtract OB1's standard from the Total amount of free rows on Both Floors

	OB1_Standard = OB1_Min_Length
	#print ('OB1 Free Rows', OB1_Standard) - Can delete later
	#Should be 4 or more. It will be the first value after division, since dividing in python returns the lowest value of the 2 numbers when divided by 2.

	# Finding the Surplus
	if (Total_Rows_Both_Floors - (OB1_Standard + OB2_Standard)) > 0:
		Surplus = Total_Rows_Both_Floors - (OB1_Standard + OB2_Standard)

		if Surplus > 0:
		    # random.random() returns a float between 0.0 and 1.0
		    if random.random() < 0.5:
		        # OB1 gets the smaller half, OB2 gets the larger remainder
		        OB1_Standard += (Surplus // 2)
		        OB2_Standard = Total_Rows_Both_Floors - OB1_Standard
		        print ('OB2 Gets the Extra Rows')
		    else:
		        # OB2 gets the smaller half, OB1 gets the larger remainder
		        OB2_Standard += (Surplus // 2)
		        OB1_Standard = Total_Rows_Both_Floors - OB2_Standard
		        print ('OB1 Gets the Extra Rows')

		print ('Surplus Was: ', Surplus)
	
	print ('OB1 Free Rows', OB1_Standard)
	print ('OB2 Free Rows', OB2_Standard)
	






	'''
	#Old Formula Below:

	OB1_Standard = Total_Rows_Both_Floors // 2
	print ('OB1 Free Rows', OB1_Standard)
	#Should be 4 or more. It will be the first value after division, since dividing in python returns the lowest value of the 2 numbers when divided by 2.

	OB2_Standard = Total_Rows_Both_Floors - OB1_Standard
	print ('OB2 Free Rows', OB2_Standard)
	#Should be 5 or more. It will be what ever value is left over after we subtract OB1's standard from the Total amount of free rows on Both Floors
	'''


	if OB1_Standard < OB1_Min_Length and OB2_Standard < OB2_Min_Length:
		print ('Not enough GEAs for Upstairs')
		sleep(99999)

	#print (int(Locate('OB1')[0][1:]))
	#print (int(Locate('OB3')[0][1:]))
	#sleep(99999)


	#Creating the list of cords only to be used to check for relief on OB1 and OB2.... strangely this works, may have to fix this later on.

	#Only for Column R
	Column_Letter_Before = ALPHABET.index(OB1_OB2_Floor_Rows[0][0]) - 1
	Column_Letter_Before = ALPHABET[Column_Letter_Before]




	for i in range(int(Locate('OB1')[0][1:]), int(Locate('OB3')[0][1:])):
		print (OB1_OB2_Floor_Rows[0][0])
		if OB1_OB2_Floor_Rows[0][0] != 'R' or ws[OB1_OB2_Floor_Rows[0][0] + str(i)].value == None or ws[OB1_OB2_Floor_Rows[0][0] + str(i)].value in EVERYTHING_ELSE or 'FREIGHT' in ws[OB1_OB2_Floor_Rows[0][0] + str(i)].value:
			OB1_OB2_Floor_Rows_Relief_List.append(OB1_OB2_Floor_Rows[0][0] + str(i))
		
		
		elif OB1_OB2_Floor_Rows[0][0] == 'R' or ws[Column_Letter_Before + str(i)].value == None or ws[Column_Letter_Before + str(i)].value in EVERYTHING_ELSE or 'FREIGHT' in ws[Column_Letter_Before + str(i)].value:
			OB1_OB2_Floor_Rows_Relief_List.append(OB1_OB2_Floor_Rows[0][0] + str(i))
			#print ('Chose This If Statement...')

		'''
		This 2nd if Statement above ^ is only made to help find the proper relief cords for the last hour of rotations, Column R, at the end of the night/shift. 
		If this isn't added, it wouldnt create a proper Relief List with all the needed cords because it wouldnt recognize the 6 - 10 shift cords and their
		would be no way to see if they were relieved or not, so it would result in a infinite loop when checking for post relief...
		'''

	

	if OB1_OB2_Floor_Rows[0][0] == 'R':
		print (OB1_OB2_Floor_Rows)
		print (OB1_OB2_Floor_Rows_Relief_List)
		#sleep(99999)
	#print (int(Locate('OB1')[0][1:]) - 1)
	#print (int(Locate('OB3')[0][1:]))
	print('')
	

	


	#--Creating the list of Post for each Floor below--


	#OB1 Post Created--
	Amount_Of_Post_To_Add = OB1_Standard - len(OB1_Temporary_Post)
	print ('Add this # of Post: ', Amount_Of_Post_To_Add)


	#If Statements Explained: Amount_Of_Post var becomes negative, it adds way to many Tier 2 post to the list....
	if Amount_Of_Post_To_Add > 0:

		#Giving this a 20% Chance to include AFF 3 in the rotation
		Random_List = [0, 0, 0, 0, 1]
		random.shuffle(Random_List)

		#Adding Aff 3 if this if Statement ever becomes true
		if Random_List[0] == 1:
			OB1_Temporary_Post.append(OB1_Tier_2[-1]) 
			Amount_Of_Post_To_Add = Amount_Of_Post_To_Add - 1

	if Amount_Of_Post_To_Add > 0:

		#This should randomize and add any extra post needed to fill in every cell for the hour
		OB1_Temporary_Post += OB1_Tier_2[:Amount_Of_Post_To_Add]
		random.shuffle(OB1_Temporary_Post)

	print (OB1_Temporary_Post)
	print ('---------')



	'''
			READ ME: This chunk of code was made to turn TR1** into a post without a at-risk

	


	OB1_Post_Starters = 0
	for a in OB1_Temporary_Post:
			if a[-1] == '*' or a in ROTATION_STARTERS:
				OB1_Post_Starters += 1

	Columns_To_Ignore = ['F', 'G', 'H', 'I']

	if len(OB1_Temporary_Post) > OB1_Min_Length and OB1_Post_Starters > 1 and 'TR1**' in OB1_Temporary_Post and OB1_OB2_Floor_Rows[0][0] not in Columns_To_Ignore:


		print ('Before: ', OB1_Temporary_Post)
		#print (OB1_Post_Starters)
		OB1_Temporary_Post.remove('TR1**')
		OB1_Temporary_Post.append('TR1')
		#print ('Take at-risk away from TR1')
		print ('After: ', OB1_Temporary_Post)
		#sleep(99999)

	'''


	#--------------------------



	#OB2 Post Created--
	Amount_Of_Post_To_Add = OB2_Standard - len(OB2_Temporary_Post)
	print ('Add this # of Post: ', Amount_Of_Post_To_Add)


	if Amount_Of_Post_To_Add > 0:

		#This should randomize and add any extra post needed to fill in every cell for the hour
		OB2_Temporary_Post += OB2_Tier_2[:Amount_Of_Post_To_Add]
		random.shuffle(OB2_Temporary_Post)
		print (OB2_Temporary_Post)

	'''
	print (OB1_Temporary_Post)
	print (len(OB1_Temporary_Post))
	print (OB2_Temporary_Post)
	print (len(OB2_Temporary_Post))
	print ('Look Here Line 6180')
	sleep(99999)
	Delete Later'''

	#Remaking Floor OB1 and OB2 List Incase the length isnt equal to the OB1 and OB2 Standard Var's
	Floor_OB1 = OB1_OB2_Floor_Rows[:OB1_Standard] #This picks from the front half of OB1_OB2 List, not exactly half
	Floor_OB2 = OB1_OB2_Floor_Rows[-OB2_Standard:] #This picks from the back half of OB1_OB2 List, not exactly half


	# 1. Check if there is any overlap (Intersection)
	# Sets are much faster and more reliable for "membership" checks
	overlap = set(Floor_OB1).intersection(set(Floor_OB2))

	# 2. Check the total length
	total_len = len(Floor_OB1) + len(Floor_OB2)
	expected_len = len(OB1_OB2_Floor_Rows)

	if overlap or total_len != expected_len:
	    print(f"Error found!")
	    if overlap:
	        print(f"Shared items found: {overlap}")
	    if total_len != expected_len:
	        print(f"Length mismatch: Combined ({total_len}) vs Expected ({expected_len})")
	    
	    # Your debug prints
	    print(Floor_OB1)
	    print(Floor_OB2)
	    print ('Upstairs Min Length: ', Upstairs_Min_Length)
	    print (ResultOfTheMath)
	    print (FlLength)
	    print ('OB1 n 2 Free Cords: ', SaveData)
	    print ('Upstairs Leads Cords: ', SaveData2)
	    print ('How many Freight Post Currently: ', FreightPostNow)
	    #Delete all these prints later
	    #sleep(99999)


	'''
	Old Code Below, New Code Above
	#This checks the code above, just to make sure Floor OB1 and OB2 List dont have any of the same cords. Length of both list should also always equal OB1_OB2_Floor_Rows
	for item1, item2 in zip(Floor_OB1, Floor_OB2):
	    if item1 == item2 or (len(Floor_OB1) + len(Floor_OB2)) != len(OB1_OB2_Floor_Rows):
	        print(f"Items are the same: {item1} and {item2}, or length of the list is off. Sleeping for 99999 seconds...")
	        print (Floor_OB1)
	        print (Floor_OB2)
	        print (OB1_OB2_Floor_Rows)
	        sleep(99999)
	        '''


	print ('')
	print ('Double Checking the #s Add Up..' )
	print(OB1_OB2_Floor_Rows)
	print(Floor_OB1) #Have OB1 Standard redefine these numbers using OB1_OB2_Floor var directly above
	print(Floor_OB2) #Have OB2 Standard redefine these numbers using OB1_OB2_Floor var directly above



	#This is placed here so its brings FLOOR_OB1 and OB2 back to its original list.
	if OB1_OB2_Floor_Rows[0][0] in Column_Letters_To_Reverse_Post:
		
		#Switching Floor cordinates in order to reverse Floor Post to have OB2 Post on OB1 vice verse
		Floor_OB1, Floor_OB2 = Floor_OB2, Floor_OB1

		#newly added




	#---Will Start Creating Rotations Below---

	Post_Pushing_Rotations_Cordinates = []
	#This will be a list of Cords Starting the Rotation

	All_Cords_Upstairs = []
	#This will be a list that counts every cord on OB1 and OB2 that has a post

	Letter_Left_To_1 = ALPHABET.index(OB1_OB2_Floor_Rows[0][0]) 
	Letter_Left_To_1 = ALPHABET[Letter_Left_To_1 - 1] 
	#This will represent the letter of the column just before the current one in OB1_OB2 Var. It will be used to accurately find Rotation pushing post in the previous rotation hour

	Letter_Currently = OB1_OB2_Floor_Rows[0][0]
	#Will represent the letter of the cords in the OB1_OB2 List. This will always be the letter of the main column for this function

	

	#Block of Code Below loops from OB1 - OB3 and finds all post cords that start the rotation for the next hour. Ex: 'BRIEF', 'TR2**' etc
	wb.save(File_Name)
	Start = int(Locate('OB1')[0][1:]) + 1 
	End = int(Locate('OB3')[0][1:])
	for i in range(Start, End):
		Cell_To_The_Left = Letter_Left_To_1 + str(i)
		Current_Cell = Letter_Currently + str(i)
		print ('')
		print ('Currently On Cord: ', Current_Cell)
		print ('Cell to the Left Cord: ', Cell_To_The_Left)
		print (Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value))
		print (i)
		if Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in EVERYTHING_ELSE and Temporary_Value(Letter_Currently + str(i), ws[Current_Cell].value) not in EVERYTHING_ELSE or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in ROTATION_STARTERS and Temporary_Value(Letter_Currently + str(i), ws[Current_Cell].value) not in ROTATION_STARTERS or '**' in Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) and '**' not in Temporary_Value(Letter_Currently + str(i), ws[Current_Cell].value):
			Post_Pushing_Rotations_Cordinates.append(Cell_To_The_Left)

		if Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB1_Tier_1 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB1_Tier_2 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB2_Tier_1 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB2_Tier_2 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in EVERYTHING_ELSE:
			All_Cords_Upstairs.append(Letter_Left_To_1 + str(i))

			#Add comments


	# Force the sacrificed B1 row to be recognized as a rotation starter
	if Using_Celebrate_From_B1 > 0:
	    # Use the coordinate from the Celebrate person on B1
		#b1_starter_cord = Letter_Left_To_1 + Celeb_B1_Cord[1:] 
		#if b1_starter_cord not in Post_Pushing_Rotations_Cordinates:
		#	Post_Pushing_Rotations_Cordinates.append(b1_starter_cord)

		# We must add the CURRENT column coordinate to the scan list
		current_b1_relief_cord = Letter_Currently + Celeb_B1_Cord[1:]
		if current_b1_relief_cord not in OB1_OB2_Floor_Rows_Relief_List:
			OB1_OB2_Floor_Rows_Relief_List.append(current_b1_relief_cord)






	'''
	This Newly Added For Loop will be used only if we are low on people for OB1 and OB2
	So we turn a B1 Tier 2 Post into a OB1/OB2 post. Here this For Loop will turn the exact Coordinate right before
	that OB1/OB2 Post on B1, into a Post Starter when we go to check for rotations on OB1 and OB2

	Reason Being: B1 has perfect Rotations already if it made it this for into the program, and that B1 Post right before it will start the 
	rotations on OB1/OB2 due to that person garuntee to be relieved at some point. Search hash tag below to follow this code thread. 

	Add this for loop into an if statement so it doesn't waste time looping if its not even needed. Make it codependent on a low attendance day on OB1 and OB2
	#-b1relief
	'''
	for i in OB1_OB2_Floor_Rows:
		if int(i[1:]) < int(Locate('OB1')[0][1:]):
			print ('---------------------------')
			print (int(Locate('OB1')[0][1:]) + 1)
			print (i)
			Post_Pushing_Rotations_Cordinates.append(Post_Pushing_Rotations_Cordinates[0][0] + i[1:])
			print ('Post Added: ', Post_Pushing_Rotations_Cordinates[0][0] + i[1:])
			print (Post_Pushing_Rotations_Cordinates)
			#sleep(99999)


	print ('')
	print (Post_Pushing_Rotations_Cordinates)
	print (All_Cords_Upstairs)
	print (OB1_OB2_Floor_Rows)
	print ('Below are Floor OB1 and OB2 Cord List')
	print(Floor_OB1) #Have OB1 Standard redefine these numbers using OB1_OB2_Floor var directly above
	print(Floor_OB2) #Have OB2 Standard redefine these numbers using OB1_OB2_Floor var directly above
	print ('Below would be the post for OB1 and OB2')
	print (OB1_Temporary_Post)
	print (OB2_Temporary_Post)



	#---------------------------------  Creating Post for OB1 and OB2  ---------------------------------

	#Create the Post for each cell down here

	#--Thought Process--
	#May have to make an If Statement based on the column we may be currently on. 
	#During the 6:30 time, OB1 and OB2 Rotations can stick to each floor, for the most part...
	#During the 7:30 Rotation hour, lets have people coming off break start switching floors

	# Creating OB1 and OB2 Post - Copy and Paste from Previous Function Here. Use Both Temp Post and Floor OB1 only for when the 2nd column needs to be made 



	Post_Tier = 0
	#This var will help us keep track of what post is next in the list of OB1_Temporary_List to create down the column

	Remade_Post = 0
	#This var will count every time a column is remade due to every post not being relieved properly

	Previous_Check = 0
	#Will be used to represent if the previous column rotated properly or had any sort of hiccups. If its 1, it will ignore running the code again

	First_Column = Floor_OB1[0][0]
	#Capturing the 1st Letter of the column so we can use it later for the previous check algorithm. We have to make sure it isnt the letter F

	No_Float_Left_Behind = []
	#This will contain cords of lead rows, and we will loop through it to see if a lead cord/cell was left empty

	OB1_Hour_Post = 0 
	#This is a var that will be used to end this while loop below for OB1. It is designed to make sure an hour rotation is created with 2 OB1 Post

	Checks = 0
	#This var will be used to determine if proper checks were out into place to make sure every post goes through properly

	Restart_While_Loop = 0
	#Var will be used to directly restart the while loop exactly below and help restart some vars in the if statement

	Repeated_Post = 0
	#This var will be used to determine if post repeat back to back, and if so will redo the whole rotations

	Cab_2_Removed = 0
	# This var will help us determine whether to check for CAB 2 during the check and balance process. If its 1 then we shouldnt care to check if its been relieved.

	#Floor_OB1_P2 = [] 
	#This will just be a list that contains the cords for Floor_OB1's 2nd Half. Same cords, just 1 more letter to the right.


	#for i in Floor_OB1:
	#	Letter_To_The_Right = ALPHABET.index(i[0])
	#	Letter_To_The_Right = ALPHABET[Letter_To_The_Right + 1]
	#	Cord = Letter_To_The_Right + i[1:]
	#	Floor_OB1_P2.append(Cord)


	while OB1_Hour_Post != 2 and Checks != 2:

		if Restart_While_Loop != 0:
			Previous_Check = 0
			Post_Tier = 0
			print ('Restarting While Loop...................')
			#sleep(9999)

		#READ THIS: Can create a BIG if statement here based on ob1_hour_post == 0 here, so when we go to use the floor ob1 it will ignore this code

		#Randomize the list of the post that will be distributed for OB1 and OB2
		random.shuffle(OB1_Temporary_Post)
		random.shuffle(OB2_Temporary_Post)

		#This if Statement will create post for the first half of the OB1 Rotations and the entire OB2 Rotations for the Hour.
		if OB1_Hour_Post == 0:


			Both_Floor_Temporary_Post = OB1_Temporary_Post + OB2_Temporary_Post
			#This is a Large List containing post for OB1 and OB2 in a specific order.
			#May need to make an IF STATEMENT for this here. Right now its OB1 then OB2 in that order due to it being early in the Rotation


			'''
			Adding these For Loops Below. Without them, the if Statement Below them pause the code because CAB Post haven't been properly removed from 
			Temp Post list, 
			and thats only because the code is re-rotating all post just to see if everyone has been relieved. If you comment them out and run, 
			the code will pause
			
			Comment this Whole if Statement Out below, the one only with the for loops, if we ever make CAB Post for everyone, and not just Full Timers.

			New: Important: Here we will add and z != Celeb_B1_Cord. Couldnt add that cause Celeb var isnt created unless we borrowed a cord from B1. 
			Use the z less than OB1 Row num logic instead The error only happens cause the borrowed B1 Cord isnt in the OB1_OB2 List yet
			'''
			print ('')
			if Restart_While_Loop != 0 and OB1_OB2_Floor_Rows[0][0] != 'F':
				for z in Locate('CAB 1'): 
					if z[0][0] == OB1_OB2_Floor_Rows[0][0] and z not in OB1_OB2_Floor_Rows and int(z[1:]) > int(Locate('OB1')[0][1:]):
						print ('CAB 1 Post have been created already. Rotations just need to be re-done, and CAB Post need to be removed from the list.')
						print (Both_Floor_Temporary_Post)
						print (OB1_OB2_Floor_Rows)
						print (z)
						Both_Floor_Temporary_Post.remove('CAB 1')


				for z in Locate('CAB 2**'):
					if z[0][0] == OB1_OB2_Floor_Rows[0][0] and z not in OB1_OB2_Floor_Rows and int(z[1:]) > int(Locate('OB1')[0][1:]):
						print ('CAB 2 Post have been created already. Rotations just need to be re-done, and CAB Post need to be removed from the list.')
						print (Both_Floor_Temporary_Post)
						print (OB1_OB2_Floor_Rows)
						print (z)
						Both_Floor_Temporary_Post.remove('CAB 2**')


			#PURPOSE: The code below is meant to catch the accuracy of all available cords, but also if the length of coords dont match the post we have
			# their is a possible issue.


			#Floor Rows Recount
			# --- 1. ACCURATE RECOUNT & SEPARATION ---
			current_col_letter = OB1_OB2_Floor_Rows[0][0]
			start_row_upstairs = int(Locate('OB1')[0][1:]) + 1
			end_row_upstairs = ws.max_row + 1

			OB1_OB2_Floor_Rows = []        # Humans available for the general shuffle
			Locked_Cab_Cords = []          # Humans currently locked into Elevator posts

			if Using_Celebrate_From_B1 > 0:
				OB1_OB2_Floor_Rows.append(Celeb_B1_Cord)
				#Choosing not to forget to add that B1 Cord we borrowed due to short staff

			for r in range(start_row_upstairs, end_row_upstairs):
				coord = current_col_letter + str(r)
				cell_val = Temporary_Value(coord, ws[coord].value)

				# Category A: Locked humans (Elevators)
				if cell_val == 'CAB 1' or cell_val == 'CAB 2**':
					Locked_Cab_Cords.append(coord)
				
				# Category B: Available humans (Empty or previously failed attempts)
				elif cell_val is None or cell_val in Overall_OB1_Post or cell_val in Overall_OB2_Post:
					OB1_OB2_Floor_Rows.append(coord)

			#breakerninefix
			if OB1_OB2_Floor_Rows[0][0] == 'P':
				OB1_OB2_Floor_Rows[:] = [item for item in OB1_OB2_Floor_Rows if item not in Upstairs_Breaker_Shifts_Cords]

			# --- 2. SYNC THE JOB POOL WITH LOCKED HUMANS ---
			# Remove jobs from the pool if a human is already "Locked" into that job on the sheet
			for cab_coord in Locked_Cab_Cords:
				job_title = Temporary_Value(cab_coord, ws[cab_coord].value)
				if job_title in Both_Floor_Temporary_Post:
					print(f"Sync: {job_title} already assigned at {cab_coord}. Removing from pool.")
					Both_Floor_Temporary_Post.remove(job_title)

			# --- 3. THE SURGICAL STAFFING DECISION (The While Loop) ---
			if len(Both_Floor_Temporary_Post) > len(OB1_OB2_Floor_Rows):
				print(f"Staffing Deficit: {len(Both_Floor_Temporary_Post)} jobs vs {len(OB1_OB2_Floor_Rows)} humans.")
				
				while len(Both_Floor_Temporary_Post) > len(OB1_OB2_Floor_Rows):
					print ('')
					print ('Error Here: Line 7099')
					print (Both_Floor_Temporary_Post)
					print (len(Both_Floor_Temporary_Post))
					print (OB1_OB2_Floor_Rows)
					print (len(OB1_OB2_Floor_Rows))
					removed = False
					
					# STEP A: Try to remove expendable Tier 2 posts first
					for p in Both_Floor_Temporary_Post:
						if p in OB1_Tier_2 or p in OB2_Tier_2:
							print(f"Short Staffed: Removing Tier 2 post {p}")
							Both_Floor_Temporary_Post.remove(p)
							removed = True
							break
					
					# STEP B: If Tier 2 is empty, sacrifice CAB 2** specifically
					if not removed:
						if 'CAB 2**' in Both_Floor_Temporary_Post:
							# 1. Remove the job string
							Both_Floor_Temporary_Post.remove('CAB 2**')
							Cab_2_Removed += 1
							
							# 2. Reset the Excel Grid and move the human back to the pool
							for i in Locked_Cab_Cords[:]: # [:] to safely remove while looping
								if Temporary_Value(i, ws[i].value) == 'CAB 2**':
									# Unmerge previous CAB merge to avoid Read-Only error
									for merged_range in list(ws.merged_cells.ranges):
										if i in merged_range:
											ws.unmerge_cells(str(merged_range))
									
									# Reset cell styles and values
									ws[i].value = None
									ws[i].fill = PatternFill(fill_type=None, stop_color='FFFFFF')
									
									# Move human to general pool
									OB1_OB2_Floor_Rows.append(i)
									Locked_Cab_Cords.remove(i)
									print(f"CRITICAL SHORTAGE: Unmerged and sacrificed CAB 2** at {i}")
							removed = True

					# STEP C: Sacrifice Celebrate on B1 by swapping Prep/Break logic
					if not removed:
						current_col_letter = OB1_OB2_Floor_Rows[0][0]
						b1_limit = int(Locate('OB1')[0][1:])
						b1_preps = []

						# 1. Find all PREP posts on B1 for this specific hour
						all_preps = Locate('PREP')
						if all_preps:
							for p_cord in all_preps:
								# Check if it matches current column and is on floor B1
								if p_cord[0] == current_col_letter and int(p_cord[1:]) < b1_limit:
									b1_preps.append(p_cord)

						# 2. If we have 2 or more, we can afford to flip one and sacrifice Celebrate
						random.shuffle(b1_preps)
						if len(b1_preps) >= 2:
							print(f"CONTINGENCY C: Found {len(b1_preps)} Preps on B1. Executing swap...")
							
							# Target the first Prep found to reverse their schedule
							target_prep_cord = b1_preps[0]
							row_num = int(target_prep_cord[1:])
							col_idx = ALPHABET.index(current_col_letter)
							
							# Identify the neighbor cell (the second half of the hour)
							neighbor_col = ALPHABET[col_idx + 1]
							neighbor_cord = neighbor_col + str(row_num)

							# REVERSE LOGIC: Instead of [Prep -> Break], make it [Break -> Prep]
							# Handle Current Cell (Change Prep to Break)
							Create_Post(target_prep_cord, row_num, col_idx + 1, 'BREAK')
							ws[target_prep_cord].fill = PatternFill(patternType='solid', fgColor=BREAK)

							# Handle Neighbor Cell (Change Break to Prep)
							Create_Post(neighbor_cord, row_num, col_idx + 2, 'PREP')
							ws[neighbor_cord].fill = PatternFill(patternType='solid', fgColor=PREP)
							
							# 3. SACRIFICE CELEBRATE
							b1_celebrates = Locate('CELEBRATE')
							if b1_celebrates:
								for c_cord in b1_celebrates:
									if c_cord[0] == current_col_letter and int(c_cord[1:]) < b1_limit:
										print(f"CONTINGENCY C: Sacrificing Celebrate at {c_cord} for upstairs.")
										
										# Unmerge and clear the Celebrate cell
										for merged_range in list(ws.merged_cells.ranges):
											if c_cord in merged_range:
												ws.unmerge_cells(str(merged_range))
										
										ws[c_cord].value = None
										ws[c_cord].fill = PatternFill(fill_type=None, end_color='FFFFFF')
										
										# Add the Celebrate person to the upstairs pool
										OB1_OB2_Floor_Rows.append(c_cord)
										
										# If reversing a reversed floor, ensure it's added to the right sub-list
										if int(c_cord[1:]) < int(Locate('OB2')[0][1:]):
											Floor_OB1.append(c_cord)
										else:
											Floor_OB2.append(c_cord)
											
										removed = True
										print ('See if it works.. Error possibly here')
										wb.save(File_Name)
										#sleep(1401)
										break # Only sacrifice one Celebrate per trigger
						else:
							print("CONTINGENCY C: Not enough Preps on B1 to execute Celebrate sacrifice. Line 7671")
							wb.save(File_Name)
							sleep(1401)

						
	

				# Final count check for debugging
				print ('')
				print(f"Final Count Hour {current_col_letter}: Jobs ({len(Both_Floor_Temporary_Post)}) | Humans ({len(OB1_OB2_Floor_Rows)})")


				print ('OB1 n OB2 Temp Post: ', Both_Floor_Temporary_Post)
				print (len(Both_Floor_Temporary_Post))
				print ('OB1 n OB2 Coords: ', OB1_OB2_Floor_Rows)
				print (len(OB1_OB2_Floor_Rows))
				print ('Updated Floor Rows: ', OB1_OB2_Floor_Rows)
				print (len(OB1_OB2_Floor_Rows))
				print ('Cabs We Made: ', Locate('CAB 1'))
				print ('Cabs We Made: ', Locate('CAB 2**'))
				print ('Locked Cab Cord: ', len(Locked_Cab_Cords))
			elif len(Both_Floor_Temporary_Post) < len(OB1_OB2_Floor_Rows):
				print ('OB1 n OB2 Temp Post: ', Both_Floor_Temporary_Post)
				print (len(Both_Floor_Temporary_Post))
				print ('OB1 n OB2 Coords: ', OB1_OB2_Floor_Rows)
				print (len(OB1_OB2_Floor_Rows))
				print ('Cabs We Made: ', Locate('CAB 1'))
				print ('Cabs We Made: ', Locate('CAB 2**'))
				print ('Locked Cab Cord: ', len(Locked_Cab_Cords))
				wb.save(File_Name)
				print ('Problem here 7427: May delete this, possibly solved issue.... not sure yet')
				sleep(55555)
			elif len(Both_Floor_Temporary_Post) == len(OB1_OB2_Floor_Rows):
				pass

			if len(Both_Floor_Temporary_Post) == 11: #Replace 11 with a var attached to ob1_ob2_floor_rows - 1
				print('')
				print ('problem here: 7440')
				sleep(14011)


				#The amount of post should always equal the amount of floor rows. in this example its 17 post and 16 rows.... delete one of the tier 2 post for ob1 if its multiple
				sleep(0)
				#Can delete this if statement at some point. No longer makes sense to have it.


			if len(Locked_Cab_Cords) > 0:
				print ('')
				print ('OB1 n OB2 Temp Post: ', Both_Floor_Temporary_Post)
				print (len(Both_Floor_Temporary_Post))
				print ('OB1 n OB2 Coords: ', OB1_OB2_Floor_Rows)
				print (len(OB1_OB2_Floor_Rows))
				print ('Before combining both List into OB1_OB2_Floor_Rows before moving on')
				print (OB1_OB2_Floor_Rows)
				print (Locked_Cab_Cords)
				#OB1_OB2_Floor_Rows.extend(Locked_Cab_Cords)
				#wb.save(File_Name)
				



			#--- Reverse the List --- 
			#We will reverse the list of floor cordinates only so we can reverse the floor post and have OB2 Post on OB1 and OB1 Post on OB2

			if OB1_OB2_Floor_Rows[0][0] in Column_Letters_To_Reverse_Post:
				Both_Floor_Temporary_Post.reverse()
				#newly added




			Overall_OB1_Post = OB1_Tier_1 + OB1_Tier_2
			Overall_OB2_Post = OB2_Tier_1 + OB2_Tier_2


			print (Both_Floor_Temporary_Post)
			print ('Sleeping at Line 5336')
			print ('')
			#sleep(99999)
			#-b1relief

			

			if OB1_OB2_Floor_Rows[0][0] != 'F':
				# Meant to fix closed rotations on OB2 if the previous hour doesn't have anybody to start the rotations. 
				#It scrambles the Post for the upcoming hour between both floors.
				print ('Find if this is where we can find out of their are any OB2 Starter Post')
				print (Both_Floor_Temporary_Post)
				print (OB1_OB2_Floor_Rows)

				# Assuming OB1_OB2_Floor_Rows = ['H22', 'H23', 'I10', ...]

				Previous_Post_Rotations_Values = []
				#This will be the list of post in general from the previous rotation hour. 
				#This will help determine if we have post in the correct spot on OB2 mainly that are able to start rotations, 
				#and not get stuck along the way. Before fixing, 
				#only OB1 had post from the previous rotation hour that pushed rotations only on OB1, which led to OB2 not being able to be pushed.

				for coord in OB1_OB2_Floor_Rows:
				    # 1. Grab the first character (e.g., 'H')
				    current_letter = coord[0]
				    
				    # 2. Grab everything after the first character (e.g., '23')
				    suffix = coord[1:]
				    
				    # 3. Calculate the letter before it in the alphabet (e.g., 'G')
				    letter_to_the_left = chr(ord(current_letter) - 1)
				    
				    # 4. Combine them to show the neighbor on the left
				    coord_to_the_left = letter_to_the_left + suffix
				    
				    # 5. Print the pair
				    if Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value): #Not sure why this is an if statement...? Can possible erase later
				    	#print(f"Main Coord: {coord} - Current Value: {Temporary_Value(coord, ws[coord].value)} ----------|---------- Main Coord Left: {coord_to_the_left} - Left Value: {Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value)}")
				    	print(f"Main Coord Left: {coord_to_the_left} - Left Value: {Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value)}")
				    	Previous_Post_Rotations_Values.append(Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value))
			
				print ('')
				print ('Total Post from Previous Rotation Hour: ', Previous_Post_Rotations_Values)


				# Method 1: Get a list of all matches
				matches = [item for item in OB2_Tier_2 if item in Previous_Post_Rotations_Values]

				# Method 2: Check if at least one item exists (True/False)
				has_any_match = any(item in Previous_Post_Rotations_Values for item in OB2_Tier_2)

				print(f"Items found: {matches}")
				print(f"Was a match found? {has_any_match}")

				if has_any_match == False:
					random.shuffle(Both_Floor_Temporary_Post)

				#Basically if their isnt a Post on OB2 that starts the rotation, it will have OB1 Post Starters rotate to an OB2 Post.


			

			# --- CAB PRIORITY ASSIGNMENT (TRANSACTION-SAFE) --- AI Assisted Code
			current_letter = OB1_OB2_Floor_Rows[0][0]
			col_idx = column_index_from_string(current_letter)
			
			# 1. Identify which CAB posts exist in the pool, but DON'T remove them yet
			cabs_to_assign = [p for p in Both_Floor_Temporary_Post if 'CAB' in p]
			
			# 2. Build priority pool (FTs first, then Leads)
			ft_candidates = [current_letter + a[1:] for a in Full_Time_Closers_Row]
			lead_backup = [current_letter + ld[1:] for ld in Leads_Upstairs]
			candidate_pool = ft_candidates + lead_backup

			for cab_job in cabs_to_assign:
				assigned = False
				for candidate_cord in candidate_pool:
					# Check if this priority person is actually available this hour
					if candidate_cord in OB1_OB2_Floor_Rows:
						# Back-to-back check
						has_recent_cab = False
						if col_idx > 7: 
							v1 = Temporary_Value(ws[candidate_cord].offset(column=-1).coordinate, ws[candidate_cord].offset(column=-1).value)
							if v1 and 'CAB' in str(v1): has_recent_cab = True
						
						if not has_recent_cab or len(candidate_pool) <= 2:
							# SUCCESS: Assign the post
							m_limit = 2 if current_letter == 'R' else 1
							target_range = candidate_cord + ':' + get_column_letter(col_idx + m_limit) + candidate_cord[1:]
							
							Create_Post(target_range, int(candidate_cord[1:]), col_idx, cab_job)
							ws[candidate_cord].fill = PatternFill(patternType='solid', fgColor=OB2)
							
							# CRITICAL: Only remove from lists NOW that we have a match
							if cab_job in Both_Floor_Temporary_Post:
								Both_Floor_Temporary_Post.remove(cab_job)
							OB1_OB2_Floor_Rows.remove(candidate_cord)
							
							assigned = True
							print(f"CAB SUCCESS: {cab_job} assigned to {candidate_cord}.")
							break # Move to next CAB job
				
				if not assigned:
					print(f"CAB NOTICE: No priority staff available for {cab_job}. Leaving in pool for GSAs.")







			#Creating the OB1 and OB2 post Below

			'''
			breakerninefix - Here we will make post for the last Breaker Post of the night. It will be all Tier 2 Post
			 The specific if Statement below is only used to make the last post for the breaker shifts. Search all #breakerninefix hashtags in the code
			#Delete them, and things will run smooth regardless. For some odd reason it messes up the rest of the cords getting post if i place this chunk
			of code within the for loop that actually creates the OB1 and OB2 post. Chose to relocate it right above.
			'''


			if OB1_OB2_Floor_Rows[0][0] == 'P' and OB1_Hour_Post == 0:

				Random_Tier_2_Post = ['PREP',]
				Random_Tier_2_Post.extend(OB2_Tier_2)
				#This will be a list of random Tier 2 Post to give

				Scroll = 0
				#Will dictate what post to give out for the last rotation using the Random Tier 2 List.

				Random_Tier_2_Post[:] = [item for item in Random_Tier_2_Post if item not in Both_Floor_Temporary_Post]
				random.shuffle(Random_Tier_2_Post)

				print ('')
				print (Random_Tier_2_Post)
				print (Both_Floor_Temporary_Post)


				for i in Upstairs_Breaker_Shifts_Cords:
					Column_Number = ALPHABET.index(i[0]) + 1 #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
					Cell = i #Not merging cells, so we dont relly need the cell var but we can keep it
					Create_Post(Cell, int(i[1:]), Column_Number, Random_Tier_2_Post[Scroll])
					OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
					ws[i].fill = OB1_Cell
					Scroll += 1


				#Creating the Dark Boxes to End the Breaker Shifts for 9:30
				for i in Upstairs_Breaker_Shifts_Cords:
					#ws.merge_cells('D' + i[1:] + ':H' + i[1:]) #Merging all cells for breaker shifts until 6PM
					Letter = 16             #Starts at the letter Q, which is where 9:30PM starts
					ws[ALPHABET[Letter] + i[1:]].fill = Dark
					ws[ALPHABET[Letter] + i[1:]].value = 'SKIP'
					Letter += 1

				wb.save(File_Name)
				print (OB1_Hour_Post)
				print ('Check here')
				#sleep(9999)



			# For loop below is the Main Area that creates the OB1 and OB2 Floor Post.
			for i in OB1_OB2_Floor_Rows:
				

		
				







				if Both_Floor_Temporary_Post[Post_Tier] in Overall_OB1_Post and OB1_Hour_Post <= 1: #i in Floor_OB1

					
					ranges_to_unmerge = set()

					for merged_range in list(ws.merged_cells.ranges):
						if i in merged_range:
							ranges_to_unmerge.add(merged_range.coord)
							# We store the string representation (e.g., "A1:C3")

					# Now unmerge them
					if len(ranges_to_unmerge) > 0:
						for r in ranges_to_unmerge:
						    ws.unmerge_cells(r)
						    print(f"Unmerged range: {r}")
						    #Delete this print at some point
						    




					Column_Number = ALPHABET.index(i[0]) + 1 + OB1_Hour_Post #If this is 0 it re-makes the 1st Column, if its a value of 1 it redoes the 2nd Column
					Cell = i #Not merging cells, so we dont relly need the cell var but we can keep it
					Create_Post(Cell, int(i[1:]), Column_Number, Both_Floor_Temporary_Post[Post_Tier])
					OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
					ws[i].fill = OB1_Cell
					print ('Post Added: ', Both_Floor_Temporary_Post[Post_Tier])
					#print (Post_Tier)
					#print (i)
					#print ('')
					#print ('Check These Below:')
					#print (Cell)
					#print (int(i[1:]))
					#print (Column_Number)
					#print (Both_Floor_Temporary_Post[Post_Tier])
					#print ('___________________________________________________________________________________')
					Post_Tier += 1
					#wb.save(File_Name)


				#Adding the != R so the if Statement below this that == R can create OB2 Post that Extend for 3 Columns.
				elif OB1_OB2_Floor_Rows[0][0] != 'R' and Both_Floor_Temporary_Post[Post_Tier] in Overall_OB2_Post and OB1_Hour_Post == 0: #i in Floor_OB2  |  #This only needs to Create Post 1 time, to ensure that, OB1_Hour_Post has to be 0
					Column_Number = ALPHABET.index(i[0]) + 1 #if its a value of 1 it redoes the 2nd Column
					index = ALPHABET.index(i[0]) + 1
					Cell = i + ':' + ALPHABET[index] + i[1:]
					Create_Post(Cell, int(i[1:]), Column_Number, Both_Floor_Temporary_Post[Post_Tier])
					OB2_Cell = PatternFill(patternType = 'solid', fgColor = OB2)
					ws[i].fill = OB2_Cell
					print ('Post Added: ', Both_Floor_Temporary_Post[Post_Tier])
					print (Post_Tier)
					print (i)
					print ('')
					Post_Tier += 1
					#wb.save(File_Name) 

				#Adding the == R so this if Statement can create OB2 Post that Extend for 3 Columns.
				elif OB1_OB2_Floor_Rows[0][0] == 'R' and Both_Floor_Temporary_Post[Post_Tier] in Overall_OB2_Post and OB1_Hour_Post == 0: #i in Floor_OB2  |  #This only needs to Create Post 1 time, to ensure that, OB1_Hour_Post has to be 0
					Column_Number = ALPHABET.index(i[0]) + 1 #if its a value of 1 it redoes the 2nd Column
					index = ALPHABET.index(i[0]) + 2
					Cell = i + ':' + ALPHABET[index] + i[1:]
					Create_Post(Cell, int(i[1:]), Column_Number, Both_Floor_Temporary_Post[Post_Tier])
					OB2_Cell = PatternFill(patternType = 'solid', fgColor = OB2)
					ws[i].fill = OB2_Cell
					print ('Post Added: ', Both_Floor_Temporary_Post[Post_Tier])
					print (Post_Tier)
					print (i)
					print ('')
					Post_Tier += 1
					#wb.save(File_Name)


			
			Upstairs_Post = []
			Upstairs_Post.extend(OB1_Tier_1  + OB2_Tier_1)
			if Locate('CELEB') != None: #You can also use Previous_Post_Rotations_Values created below this code, as an indicator.
				for i in Upstairs_Post:
					for a in Locate(i):
						if int(a[1:]) < int(Locate('OB1')[0][1:]):
							print (a)
							OB1_OB2_Floor_Rows_Relief_List.append(OB1_OB2_Floor_Rows_Relief_List[0][0] + a[1:])
				OB1_OB2_Floor_Rows_Relief_List = list(dict.fromkeys(OB1_OB2_Floor_Rows_Relief_List))
				#Removing Duplicates


			#Important Code Above: When you find the time, revise this for loop to only activate if we ever borrow from B1. This adds any B1 Cords we may have borrowed
			# To the relief list. Very important code.

			

			
			''' #Moving this Code Before the OB1_OB2 Post Creation Post - Delete this big chunk of code later if you need too.
			if OB1_OB2_Floor_Rows[0][0] == 'H':
				print ('Find if this is where we can find out of their are any OB2 Starter Post')
				print (Both_Floor_Temporary_Post)
				print (OB1_OB2_Floor_Rows)

				# Assuming OB1_OB2_Floor_Rows = ['H22', 'H23', 'I10', ...]

				Previous_Post_Rotations_Values = []
				#This will be the list of post in general from the previous rotation hour. 
				#This will help determine if we have post in the correct spot on OB2 mainly that are able to start rotations, 
				#and not get stuck along the way. Before fixing, 
				#only OB1 had post from the previous rotation hour that pushed rotations only on OB1, which led to OB2 not being able to be pushed.

				for coord in OB1_OB2_Floor_Rows:
				    # 1. Grab the first character (e.g., 'H')
				    current_letter = coord[0]
				    
				    # 2. Grab everything after the first character (e.g., '23')
				    suffix = coord[1:]
				    
				    # 3. Calculate the letter before it in the alphabet (e.g., 'G')
				    letter_to_the_left = chr(ord(current_letter) - 1)
				    
				    # 4. Combine them to show the neighbor on the left
				    coord_to_the_left = letter_to_the_left + suffix
				    
				    # 5. Print the pair
				    if Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value): #Not sure why this is an if statement...? Can possible erase later
				    	#print(f"Main Coord: {coord} - Current Value: {Temporary_Value(coord, ws[coord].value)} ----------|---------- Main Coord Left: {coord_to_the_left} - Left Value: {Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value)}")
				    	print(f"Main Coord Left: {coord_to_the_left} - Left Value: {Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value)}")
				    	Previous_Post_Rotations_Values.append(Temporary_Value(coord_to_the_left, ws[coord_to_the_left].value))
			
				print ('')
				print ('Total Post from Previous Rotation Hour: ', Previous_Post_Rotations_Values)

				# The larger list you want to check against
				#Previous_Post_Rotations_Values = ["Admin", "Clinical", "ER", "Research", "Outreach"]

				# The smaller list of items you're looking for
				#current_items = ["ER", "Radiology", "Research"]

				# Method 1: Get a list of all matches
				matches = [item for item in OB2_Tier_2 if item in Previous_Post_Rotations_Values]

				# Method 2: Check if at least one item exists (True/False)
				has_any_match = any(item in Previous_Post_Rotations_Values for item in OB2_Tier_2)

				print(f"Items found: {matches}")
				print(f"Was a match found? {has_any_match}")

				if has_any_match == True:
					random.shuffle(Both_Floor_Temporary_Post)
				'''


			
			#if OB1_OB2_Floor_Rows[0][0] == 'H': - Delete this post later if you need too
			#	print ('Test is doneeeeeee-')
			#	sleep(5555)


			for i in OB1_OB2_Floor_Rows:
				Letter_To_The_Left_Check = ()
				if i in OB2_Tier_2: 
				#i should print the value then the cord from the letter to the left and find out if any of them have an OB2 Post Starter or atrisk. 
				#If not we can randomize all these post that we are creating to the right, to ensure rotations go through.
					print (i)
					print ('Their is a post on OB2 that starts the rotation')
			print ('Sleeping at Line 6444')
			print ('')
			#sleep(99999)
			#-b1relief
			#Delete later


			

			



			'''
			#Searching for an ERROR Here where post only have OB1 Floors and not enough OB2 Floor post...

			#The post are there, the amount of both floor post are just bigger than what its supposed to be. it was 15 post, and only 11 spots
			if len(Both_Floor_Temporary_Post) > len(OB1_OB2_Floor_Rows):
				print ('There is a major difference in amount of post and open floors here')
				print ('Post to be Added: ', Both_Floor_Temporary_Post)
				print (len(Both_Floor_Temporary_Post))
				print ('Open Cells: ', OB1_OB2_Floor_Rows)
				print (len(OB1_OB2_Floor_Rows))

				print ('OB1 Standard: ', OB1_Standard)
				print ('OB2 Temporary Post: ', OB1_Temporary_Post)
				print ('OB2 Standard: ', OB2_Standard)
				print ('OB2 Temporary Post: ', OB2_Temporary_Post)
				sleep(99999)
				'''




				
			#Checking for REPEATING Back to Back Post. 
			#If its an OB1 Post we will Check 3 the last 3 Cells, if its and OB2 Post we will check the last 2 Cells

			OB1_List_Of_Repeated_Post = [] #If this List has 3 Post that shouldnt be back to back, we will redo rotations by adding 1 to the restart var
			OB2_List_Of_Repeated_Post = [] #If this List has 2 Post that shouldnt be back to back, we will redo rotations by adding 1 to the restart var
			for i in OB1_OB2_Floor_Rows:
				cell = i 
				print (Temporary_Value(cell, ws[cell].value))
				#First Coordinate Value
				
				Cell_Left_1 = ALPHABET.index(i[0])
				Cell_Left_1 = ALPHABET[Cell_Left_1 - 1]
				cell2 = Cell_Left_1 + i[1:]
				print (Temporary_Value(cell2, ws[cell2].value))
				#2nd Coordinate to the left by 1, Value

				Cell_Left_2 = ALPHABET.index(i[0])
				Cell_Left_2 = ALPHABET[Cell_Left_2 - 2]
				cell3 = Cell_Left_2 + i[1:]
				print (Temporary_Value(cell3, ws[cell3].value))
				#3rd Coordinate to the left by 2, Value

				print ('')

				keywords = ['FLOAT', 'BREAK']
				#All post that should be ignored if repeated in the repeated list post below, mainly for OB2 
				#['S ELE**', 'TR2**', 'TR3**', 'TR4**', 'TR5**', 'AFF 3**']

				keywords_for_OB1 = ['FLOAT', 'BREAK', 'TR1**', OB1_Tier_2[0], OB1_Tier_2[1], OB1_Tier_2[2], OB1_Tier_2[3], OB1_Tier_2[4], OB1_Tier_1[3], None] #Recently added the None's if an issue delete it
				#All post that should be ignored if repeated in 2's, in the repeated list post below, for OB1 Specifically

				if Temporary_Value(cell, ws[cell].value) in Overall_OB1_Post:
					OB1_List_Of_Repeated_Post.extend([ws[cell].value, ws[cell2].value, ws[cell3].value])


					if OB1_List_Of_Repeated_Post[0] == OB1_List_Of_Repeated_Post[1] or OB1_List_Of_Repeated_Post[1] == OB1_List_Of_Repeated_Post[2]:
						if all(item == OB1_List_Of_Repeated_Post[0] for item in OB1_List_Of_Repeated_Post):
							print ('Last 3 Cells: ', OB1_List_Of_Repeated_Post)
							print ('All 3 Of These Post Are The Same..')
							print ('Check Cord: ', i)
							print ('Later can add mini if statement that passes this if the 2 post are small post like tr1, esc, etc... can also make a seperate list for those post')
							Repeated_Post += 1
							Fixed_Repeated_Post.append(OB1_List_Of_Repeated_Post)
							Fixed_Repeated_Post.append(i)
							break
					    	#Checks for 3 OB1 Post that are exactly Back to Back.......

						"""
						PLEASE READ: 

						Nothing is wrong with this code, but sometimes it makes rotations a bit harder and longer, so im in between using this
						and not using this. It basically forces the code to rerotate if their are 2 Annoying OB1 Post that repeat back to back.

						elif not any(keywords_for_OB1 in OB1_List_Of_Repeated_Post for keywords_for_OB1 in keywords_for_OB1):
							print ('Last 3 Cells: ', OB1_List_Of_Repeated_Post)
							print ('2 Of These Post Are The Same..')
							print ('Check Cord: ', i)
							print ('Later can add mini if statement that passes this if the 2 post are small post like tr1, esc, etc... can also make a seperate list for those post')
							Repeated_Post += 1
							Fixed_Repeated_Post.append(OB1_List_Of_Repeated_Post)
							Fixed_Repeated_Post.append(i)
							break
							#If the 1st 2 post within the list are the same, or last 2, then its a repeated post. Will force it to remake post based on what post they are
						"""


				elif Temporary_Value(cell, ws[cell].value) in Overall_OB2_Post:
					OB2_List_Of_Repeated_Post.extend([Temporary_Value(cell, ws[cell].value), Temporary_Value(cell2, ws[cell2].value), Temporary_Value(cell3, ws[cell3].value)])
					

					if all(item == OB2_List_Of_Repeated_Post[0] for item in OB2_List_Of_Repeated_Post):
						if not any(keyword in OB1_List_Of_Repeated_Post for keyword in keywords):
							#If all post are the same, then its a repeated post. Will force it to remake post based on what post they are
							print ('Last 3 Cells: ', OB2_List_Of_Repeated_Post)
							print ('All Of These Post Are The Same..')
							print ('Check Cord: ', i)
							Repeated_Post += 1
							Fixed_Repeated_Post.append(OB2_List_Of_Repeated_Post)
							Fixed_Repeated_Post.append(i)
							break


				OB1_List_Of_Repeated_Post.clear()
				OB2_List_Of_Repeated_Post.clear()




			#print ('Test Done...')
			#sleep(99999)


			#Checking Rotations Below - Code below will make sure everyone is relieved. Eventually may have to mini while loop this.. or maybe not

		#This if Statement will create post for the 2nd half of the OB1 Rotations for the Hour. This activates after OB1 and OB2 has already been check for relief.
		elif OB1_Hour_Post == 1:
			print (OB1_OB2_Floor_Rows)
			print ('Floor OB1 Cords: ', Floor_OB1)
			print (OB1_Standard)
			print (Overall_OB1_Post)
			print ('Floor OB2 Cords: ', Floor_OB2)
			print (OB2_Standard)
			print (Overall_OB2_Post)

			#Re-Organizing and/or Double Checking Floor OB1 and OB2 Based on Post that are Made Already
			#Their is a mini error here where Floor OB1's or OB2's Length don't properly add up with the Post that are already made at this point. 
			#Also, we can use OB1_Standard and OB2_Standard as a reference for how many cords is supposed to be in each list (Floor_OB1/Floor_OB2)

			
			#if len(Floor_OB1) != OB1_Standard and len(Floor_OB2) != OB2_Standard:
			#Commenting the Above Code out... To many errors, no if statement needed and we need to reconfigure the Floor_OB1 and Floor OB2 list regardless.
			Floor_OB1.clear()
			Floor_OB2.clear()

			for i in OB1_OB2_Floor_Rows:
				if ws[i].value in Overall_OB1_Post:
					Floor_OB1.append(i)
				elif ws[i].value in Overall_OB2_Post:
					Floor_OB2.append(i)

			print ('')
			print ('Corrected OB1 and OB2 Floor List..')
			print ('')
			print (OB1_OB2_Floor_Rows)
			print ('Floor OB1 Cords: ', Floor_OB1)
			print (OB1_Standard)
			print ('Floor OB2 Cords: ', Floor_OB2)
			print (OB2_Standard)
			#newly added


			



			#This will just be a list that contains the cords for Floor_OB1's 2nd Half. Same cords, just 1 more letter to the right.
			Floor_OB1_P2 = [] 

			Post_Tier = 0 #Might could replace this some where else

			Previous_Check = 2 #This var accidently resets to 0.... I believe its because of an if statement at the start just below while OB1_Hour_Post != 2 and Checks != 2:

			for i in Floor_OB1:
				#Please Read: Replace any code with Letter_To_Right with offset function. So much easier and may increase code execution time
				Letter_To_The_Right = ALPHABET.index(i[0])
				Letter_To_The_Right = ALPHABET[Letter_To_The_Right + 1]
				Cord = Letter_To_The_Right + i[1:]
				Floor_OB1_P2.append(Cord) #Might can delete this

				print ('Starting With: ', Cord) 
				print ('Error Watching - Floor_OB1: ', Floor_OB1) #Error - Lead Cord should not be in here
				print ('Floor OB1 - 2nd Half Cords, ', Floor_OB1_P2) # Should be every cord to the right of Floor OB1
				print ('Post Added: ', OB1_Temporary_Post[Post_Tier])
				print ('OB1_Hour_Post: ', OB1_Hour_Post)
				print (i)
				print ('Sleeping only if OB1 Hour is 1')
				



				#Can also create post somewhere here
				#Before: ALPHABET.index(i[0]) + 1 + OB1_Hour_Post - if Fix, then Delete this

				#if OB1_Hour_Post == 0:
				#	Column_Number = ALPHABET.index(i[0]) + 1 + OB1_Hour_Post #If this is 0 it makes the 1st Column, if its a value of 1 it does the 2nd Column
				#elif OB1_Hour_Post == 1:
				#	Column_Number = ALPHABET.index(Letter_To_The_Right) + 1 + OB1_Hour_Post
				#Column_Number = ALPHABET.index(Letter_To_The_Right) + 1 + OB1_Hour_Post
				Column_Number = ALPHABET.index(Cord[0]) + OB1_Hour_Post #If this is 0 it makes the 1st Column, if its a value of 1 it does the 2nd Column
				#Cell = i #Not merging cells, so we dont relly need the cell var but we can keep it

				print ('Post: ', OB1_Temporary_Post[Post_Tier]) #Trying to catch the error here
				print (Post_Tier)
				print ('Restarted While Loop: ', Restart_While_Loop)
				Create_Post(Cord, int(i[1:]), Column_Number, OB1_Temporary_Post[Post_Tier])
				OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
				ws[i].fill = OB1_Cell
				print ('Post Added: ', OB1_Temporary_Post[Post_Tier])
				print ('Cord Added To: ', Cord)
				print (Post_Tier)
				Post_Tier += 1
				#wb.save(File_Name)

			



			if OB1_OB2_Floor_Rows[0][0] in Column_Letters_To_Reverse_Post:
		
				#Switching Floor cordinates in order to reverse Floor Post to have OB2 Post on OB1 vice verse
				#Floor_OB1, Floor_OB2 = Floor_OB2, Floor_OB1

				print ('Done..... Line 5077')
				print ('Their is a error with the cords in floor ob1 above, somewhere alone the lines an ob2 cord is added on the ob1 list')
				#sleep(99999)

				#newly added


			print ('Previous Check: ', Previous_Check)
			sleep(0)

		#This is put into place basically to reset the whole while loop and catch any post that may be back to back and prevent them from happening....
		if Repeated_Post != 0:
			Repeated_Post = 0

			#Reseting these variables below as double tap
			Previous_Check = 0
			Post_Tier = 0

			print ('')
			print ('Should restart whole while loop....')
			#sleep(99999)
			continue

		if OB1_OB2_Floor_Rows[0][0] == 'F':
			Previous_Check += 1
			Checks += 1
			OB1_Hour_Post += 1
			#Basically skips the first Rotation Hour check for the very 1st Shift of the day which is column F. This stays

		if OB1_Hour_Post == 0:
			print ('Checking Previous Post')
			sleep(0)
			#wb.save(File_Name)

			#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
			Relieved_Post_Inner_List_Var = 0 


			while Previous_Check == 0:


				Letter_Next_To_1 = ALPHABET.index(Floor_OB1[0][0]) #Will represent the first Letter/Column originally in the Floor OB1 List
				Letter_Next_To_1 = ALPHABET[Letter_Next_To_1 - 1] #Will represent the second Letter/Column Left in original Floor OB1 List


				#This var will keep count of how many Post will start rotations on OB1
				Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates)

				#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
				Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]

				#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
				Relieved_Cords = []


				Stop_Var = 0
				#Can delete later, only used for testing

				Post_Starter_Already_Relieved = []
				#will be a list consisted of cords that have already started rotations and their full rotation line has been completed


				#if OB1_OB2_Floor_Rows[0][0] == 'J':
				#	print ('')
				#	print ('Seeing what starts Rotations for B1 change... Line 5599')
				#	print (Post_Pushing_Rotations_Cordinates)
				#	print (OB1_OB2_Floor_Rows)
				#	sleep(99999)
					#-b1relief

				#Looping through all the starting post in the OB1 Post starter list.
				#for Post_Starter in Post_Pushing_Rotations_Cordinates:
				#Old Code - Delete Later

				# idx keeps us perfectly synced with the number of sub-lists we created
				for idx, Post_Starter in enumerate(Post_Pushing_Rotations_Cordinates):
					Relieved_Post_Inner_List_Var = idx
					print (Relieved_Post_Inner_List_Var)
					print (len(Relieved_Post))
					sleep(0)
					print ('-------------------------------------------------------------')
					print('')
					print ('')

					No_Trancendence_1 = 0
					#This var will be used to end an infinite while loop when ever TR! has no relief. Its rare but it happens

					
					
					#-b1relief
					if Temporary_Value(Post_Starter, ws[Post_Starter].value) in Mix_Of_Both_B1_Tiers or EVERYTHING_ELSE[1] == Temporary_Value(Post_Starter, ws[Post_Starter].value) or '**' in Temporary_Value(Post_Starter, ws[Post_Starter].value) or Temporary_Value(Post_Starter, ws[Post_Starter].value) in ROTATION_STARTERS:

						Rotation_Complete = 0 #Var used to help end the while Loop below. It'll become 1 at the end of every complete rotation

						while Rotation_Complete == 0: #Should be connected to the amount of post rotation starters there are in this hour rotation
							print ('Starting With: ', Temporary_Value(Post_Starter, ws[Post_Starter].value))
							print ('Cord of Post Starter: ', Post_Starter)
							Post_Starter_Already_Relieved.append(Post_Starter)
							print ('Full List of Post Starters: ', Post_Pushing_Rotations_Cordinates)
							print ('Inner List #: ', Relieved_Post_Inner_List_Var)
							print ('Both Floor List: ', OB1_OB2_Floor_Rows_Relief_List) #Still have to fix this so it adds every cord from ob1 to ob3.... hasnt been fixed
							print ('Showing Current List Progress: ', Relieved_Post)
							print ('Did we borrow from B1: ', Using_Celebrate_From_B1)
							print ('ROTATION Starters: ', ROTATION_STARTERS)
							print ()
							print ('------')
							print ('')

							#May need to add this concept to my B1 Rotation Post Check and Balance
							# Determine which physical row we are following
							if len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
								target_row = Post_Starter[1:] # Start with the row of the starter
							elif len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 2:
								cell_Letter = ws[Relieved_Cords[0]].offset(row=0, column=-1).coordinate
								cell_Letter = cell_Letter[0]

								cell2_Letter = Relieved_Cords[0][0]
 


						        # Follow the row of the last coordinate we physically relieved
								for a in Locate(Relieved_Post[Relieved_Post_Inner_List_Var][-1]):
									if Temporary_Value(a, ws[a].value) == Temporary_Value(cell_Letter + a[1:], ws[cell_Letter + a[1:]].value):
										target_row = a[1:]


								#target_row = Relieved_Cords[-1][1:]
								print ('Current List: ', Relieved_Post[Relieved_Post_Inner_List_Var])
								print ('Relieved Cords: ', Relieved_Cords)
								print (cell_Letter, cell2_Letter)
								
							

							
							
							for Floor_Cords in OB1_OB2_Floor_Rows_Relief_List: 

								# RECRUITMENT FILTER: Ignore every row except the one we are following
								#May need to add this concept to my B1 Rotation Post Check and Balance
								if Floor_Cords[1:] != target_row:
									continue

								

								cell2 = Temporary_Value(Floor_Cords, ws[Floor_Cords].value)                 	#ws[Floor_Cords].value | Will be the value of the Floor's Cord List Cells which is the most recent list in the excel and also printed out in the terminal when its ran
								cell = Temporary_Value(Letter_Next_To_1 + Floor_Cords[1:], ws[Letter_Next_To_1 + Floor_Cords[1:]].value) #ws[Letter_Next_To_1 + Floor_Cords[1:]].value | Will be the value of the 2nd cell directly to the left of the most filled in Column

								

								



									
								print ('Cell: ', Letter_Next_To_1 + Floor_Cords[1:])
								print ('Cell 2: ', Floor_Cords)
								print ('Post Starter: ', Post_Starter)
								print ('Inner Var Value: ', Relieved_Post_Inner_List_Var)
								print ('Relieved Post List: ', Relieved_Post)
								#Minor error here.... Inner Var happens to be a higher value than it should be... in return starts the list with the last inner list
								print (len(Relieved_Post))
								print ('++++++')
								
								print ('')

								'''
								if Relieved_Post_Inner_List_Var > 0 and len(Relieved_Post[0]) == 0:
									# Band Aid sort of solution: 
									print ('Testing Fix: Resetting var back to 0, didnt work and became and endless loop')
									Relieved_Post_Inner_List_Var = 0
									break
									sleep(8) #Delete once fixed
								'''



								#This is to prevent the code from relieving any Post that isnt a rotator or has an atrisk
								if (Letter_Next_To_1 + Floor_Cords[1:]) == Post_Starter or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1: 
								
									print ('')
									print ('Made it In the If Statement... Wont pass If Statement Below... Most likely the Cord that needs to be relieved isnt in the Relief Cord List...')
									Post_Starter_Value = Temporary_Value(Post_Starter, ws[Post_Starter].value)
									print (Post_Starter_Value)
									#print (EVERYTHING_ELSE[1])
									print ('-----')
									sleep(0)
									#This Var will be used to capture the value of any post starting rotations, so it doesnt give a None error.

									
									#-b1relief
									if '*' in Post_Starter_Value or Post_Starter_Value in EVERYTHING_ELSE or Post_Starter_Value in FREIGHT or Post_Starter_Value in Mix_Of_Both_B1_Tiers or Post_Starter_Value in ROTATION_STARTERS: 
										print ('Currently Relieved: ', Relieved_Post[Relieved_Post_Inner_List_Var])
										

										Has_Post_Been_Relieved = any((Letter_Left_To_1 + Floor_Cords[1:]) in sublist for sublist in Relieved_Cords) 
										#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for

										#Was Here Last!!!!!!!!!!!!!11 - Delete Later
										cell2 = Temporary_Value(Floor_Cords, ws[Floor_Cords].value)
										cell = Temporary_Value(Letter_Next_To_1 + Floor_Cords[1:], cell)

										cell2_cord = Floor_Cords
										cell_cord = Letter_Next_To_1 + Floor_Cords[1:]

										print (cell + ' Relieved Status: ', Has_Post_Been_Relieved)


										

										

										#This if statement is always for the Post Starter cell and also the post that come after that may not start the rotation, ut will keep it going. Thats why the or statement is for if the list has a higher value than 1
										if cell in ROTATION_STARTERS or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1 or cell == EVERYTHING_ELSE[1] or cell in Mix_Of_Both_B1_Tiers:
											print ('Made it Here....')
											if '*' in cell and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0 or cell in ROTATION_STARTERS and cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0:
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) #Relieved_Post_Inner_List_Var - Determines what list we are using first inside of Relieved_Post; Starting with 0
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Floor_Cords)
												print ('if statement 1-')
									

												if cell2 in ROTATION_STARTERS or '*' in cell2 and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													break
													#This if statement ends Rotations super early

											elif cell in ROTATION_STARTERS and cell2 in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1  :
												sleep(0)
												print ('if statement 2+')
												if cell in EVERYTHING_ELSE and cell2 in ROTATION_STARTERS or cell in EVERYTHING_ELSE and cell2 in FREIGHT:
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell)
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)
													print ('Current Relieved Cords: ',Relieved_Cords)
													print (f'Cell 1: {cell} and Cell 2: {cell2}')
													#An Error is Here - 7/26/26 Relieving the wrong cords. Now fixed 7/28. Delete when you come across this again

													# FIX: Print BEFORE incrementing
													print(f"Chain {Relieved_Post_Inner_List_Var} complete: {Relieved_Post[Relieved_Post_Inner_List_Var]}")

													# AI Assisted Solution: Only move to the next rotation line if we actually successfully relieved someone
													if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1

													
													print ('if statement 2A..... This ROTATION is Ending EARLY')
													break # Exit Floor_Cords loop


												elif cell in EVERYTHING_ELSE: #If its a full BREAK.... kinda thinking this code is useless too because of the at risk if statement
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)

													# AI Assisted Solution: Only move to the next rotation line if we actually successfully relieved someone
													if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
														# FIX: Print BEFORE incrementing
														print(f"Chain {Relieved_Post_Inner_List_Var} complete: {Relieved_Post[Relieved_Post_Inner_List_Var]}")
														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1
													else:
													    # Optional: print for debugging so you see it failing
													    print(f"Index tried to move from {Relieved_Post_Inner_List_Var}, but list was empty. Staying at {Relieved_Post_Inner_List_Var}.")

													#Old Code: Delete later if new code works
													#Relieved_Post_Inner_List_Var += 1
													#Rotation_Complete += 1
													print ('if statement 2B')
													break # Exit Floor_Cords loop

												elif cell not in EVERYTHING_ELSE: #Basically if its LAUNCH or has an at risk
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)

													 # FIX: Print BEFORE incrementing
													print(f"Chain {Relieved_Post_Inner_List_Var} complete: {Relieved_Post[Relieved_Post_Inner_List_Var]}")

													# AI Assisted Solution: Only move to the next rotation line if we actually successfully relieved someone
													if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1



											

													print ('This ROTATION is Ending EARLY')
													print (cell)
													print (cell2)
													print (Relieved_Post)
													print ('Cell2 Cord: ', Floor_Cords)
													print ('Cell Cord: ', Letter_Next_To_1 + Floor_Cords[1:])
													#sleep(9999)
													#Relieved_Post_Inner_List_Var += 1
													#Rotation_Complete += 1
													#Delete later if needed
													Stop_Var = 0
													print ('Adding Relieved Post Var + 1')
													print('')
													print('')
													print('')
													print ('if statement: 3')
													break # Exit Floor_Cords loop
												elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)
													print ('Ending this Rotation--')

													# AI Assisted Solution: Only move to the next rotation line if we actually successfully relieved someone
													if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1
														break
													else:
													    # Optional: print for debugging so you see it failing
														print(f"Index tried to move from {Relieved_Post_Inner_List_Var}, but list was empty. Staying at {Relieved_Post_Inner_List_Var}.")
														
													

													
												elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in Relieved_Post[Relieved_Post_Inner_List_Var] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													print (cell + 'Is Relieving ' + cell2)
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)
													print ('if statement 4-')
						
											elif cell not in ROTATION_STARTERS and cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1]:
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Floor_Cords)
												print ('Current Relieved Cords: ',Relieved_Cords)
												print (f'Cell 1: {cell} and Cell 2: {cell2}')
												print ('if statement 5-')

												if cell2 in ROTATION_STARTERS or '*' in cell2:
													# AI Assisted Solution: Only move to the next rotation line if we actually successfully relieved someone
													if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1
													else:
													    # Optional: print for debugging so you see it failing
													    print(f"Index tried to move from {Relieved_Post_Inner_List_Var}, but list was empty. Staying at {Relieved_Post_Inner_List_Var}.")

													#Old Code: Delete Later
													#Relieved_Post_Inner_List_Var += 1
													#Rotation_Complete += 1
													print ('Ending this rotation')
													break

											elif cell == EVERYTHING_ELSE[1] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0: 
											#This will be used for BREAKER SHIFTS after their BRIEFING

												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Floor_Cords)
												print (cell + 'Is Relieving ' + cell2)
												print ('if statement 6-')

												if cell2 in ROTATION_STARTERS or '*' in cell2:

													# AI Assisted Solution: Only move to the next rotation line if we actually successfully relieved someone
													if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
														Relieved_Post_Inner_List_Var += 1
														Rotation_Complete += 1
														break
													else:
													    # Optional: print for debugging so you see it failing
													    print(f"Index tried to move from {Relieved_Post_Inner_List_Var}, but list was empty. Staying at {Relieved_Post_Inner_List_Var}.")

													#Old Code Below: Deleter Later
													#Relieved_Post_Inner_List_Var += 1
													#Rotation_Complete += 1
													#This if statement ends Rotations super early

												elif cell2 not in ROTATION_STARTERS or '*' not in cell2:
													#Relieving a regular post after briefing, this if statement passes it on to the next if statement
													pass


									print ('')
									
								



												





							#Used to end the while loop because the Relieved_Post_Inner_List_Var will represent a list value that doesnt exist. 
							#Check All Post Here
							if Relieved_Post_Inner_List_Var == len(Relieved_Post):
								print ('----------------------------------')
								print('')
								print ('Relieved Vars value has passed the amount of list we created. Ending Rotation early to force a re-rotate')
								print ('Showing Current List Progress: \n', Relieved_Post)
								print ('')
								print ('Relief Cords List: ', OB1_OB2_Floor_Rows_Relief_List)
								Rotation_Complete += 1
								#sleep(99999)




								#---Checking Rotation Code Here---



								#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
								Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
								print ('Before Filter: ', Collapsed_List_Of_Relieved_Post)

								#We have to filter this list, to remove any post that are rotation starters besides the on that belongs to a Tier 1 List
								Junk_List = [] #List thats going to have post that we dont want
								for Post in Collapsed_List_Of_Relieved_Post:
									if Post in FREIGHT or Post not in OB1_Tier_1 and Post not in OB2_Tier_1: 
										Junk_List.append(Post)

								print ('Junk List: ', Junk_List) 

								for Junk_Post in Junk_List:
									Collapsed_List_Of_Relieved_Post.remove(Junk_Post)

								Collapsed_List_Of_Relieved_Post = list(set(Collapsed_List_Of_Relieved_Post))
								#Filtering for Duplicates and Deleting Them

								print('')
								print ('After Filter: ', Collapsed_List_Of_Relieved_Post) 
								print ('# of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post)) #Should be 9 Based on the sum of both OB1 and OB2 Tier 1's. To be sure just add both OB1 and OB2 len()
								OB1_OB2_Tier_1_List = OB1_Tier_1 + OB2_Tier_1
								print ('OB1 and OB2 Tier 1 List: ', OB1_OB2_Tier_1_List) #List of Important Post that need to be Relieved!!!

								# Basically if we ended up removing the CAB 2 Post entirely, dont look for this specific Tier 1 Post.
								if Cab_2_Removed != 0:
									OB1_OB2_Tier_1_List.remove('CAB 2**')

								Unrelieved_Post = list(set(OB1_OB2_Tier_1_List) - set(Collapsed_List_Of_Relieved_Post))
								print ('Tier 1 Post Hasnt Been Relieved: ', Unrelieved_Post)
								#wb.save(File_Name)
								
								#This is a list that will print every important Tier 1 Post that hasnt been relieved



								if len(Unrelieved_Post) == 0:
									print ('Everything Checks Out for the 1st Half of Rotations....')
									print ('Times Redid Rotations: ', Restart_While_Loop)
									Restart_While_Loop = 0
									Previous_Check += 1
									Checks += 1
									OB1_Hour_Post += 1
									print ('Previous Check: ', Previous_Check)
									#sleep(99999)
									break
									#This will break the main while Loop and continue on to the next important part of this Function
								elif len(Unrelieved_Post) != 0:
									'''
									-Reset Checks and Post- 

									This will reset it all the way back to the main while loop, that way it recreates the post and shuffles the order so
									we get a different outcome everytime. Easier that way

									'''


									Previous_Check += 1
									Restart_While_Loop += 1
									break
								






								#sleep(99999)
								#break


						


			#print ('Seeing if everything works')
			#sleep(99999)

		if Previous_Check == 1:
			#Forcing it to restart the main while loop with continue so certain conditions are met and certain variables/list are assigned for the next if statement
			#Any if statements with the conditions of OB1_Hour_Post == 0, after the while loop restarts will be skip so we can continue on to the ones that need to be met with a value of 1
			Previous_Check += 1
			continue

		if OB1_Hour_Post == 1:
			print ('Checking Previous Post--------+')
			sleep(0)
			#wb.save(File_Name)


			#Relieved_Post will be made up of multiple list, so we will use this var to keep count for each increment we use to determine which list inside the list will be used for the current and next
			Relieved_Post_Inner_List_Var = 0 

			Post_Pushing_Rotations_Cordinates_2 = []
			#This will be a list of Cords Starting the Rotation

			#All_Cords_Upstairs = []
			#This will be a list that counts every cord on OB1 and OB2 that has a post

			Letter_Left_To = ALPHABET.index(Floor_OB1_P2[0][0]) 
			Letter_Left_To = ALPHABET[Letter_Left_To - 1] 
			#This will represent the letter of the column just before the current one in Floor OB1 P2 List. It will be used to accurately find Rotation pushing post in the previous rotation hour

			Letter_Currently = Floor_OB1_P2[0][0]
			#Will represent the letter of the cords in the OB1_OB2 List. This will always be the letter of the main column for this function

	

			#Block of Code Below loops from OB1 - OB3 and finds all post cords that start the rotation for the next hour. Ex: 'BRIEF', 'TR2**' etc
			Start = int(Locate('OB1')[0][1:]) + 1 
			End = int(Locate('OB3')[0][1:])
			for i in range(Start, End):
				Cell_To_The_Left = Letter_Left_To + str(i)
				Current_Cell = Letter_Currently + str(i)
				#print (Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value))
				if Temporary_Value(Letter_Left_To + str(i), ws[Cell_To_The_Left].value) in EVERYTHING_ELSE and Current_Cell in Floor_OB1_P2 or Temporary_Value(Letter_Left_To + str(i), ws[Cell_To_The_Left].value) in ROTATION_STARTERS  and Current_Cell in Floor_OB1_P2 or '**' in Temporary_Value(Letter_Left_To + str(i), ws[Cell_To_The_Left].value) and Current_Cell in Floor_OB1_P2:
					Post_Pushing_Rotations_Cordinates_2.append(Cell_To_The_Left)

				#if Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB1_Tier_1 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB1_Tier_2 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB2_Tier_1 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in OB2_Tier_2 or Temporary_Value(Letter_Left_To_1 + str(i), ws[Cell_To_The_Left].value) in EVERYTHING_ELSE:
				#	All_Cords_Upstairs.append(Letter_Left_To_1 + str(i))

			print ('')
			print (Post_Pushing_Rotations_Cordinates_2)
			print (Letter_Left_To)
			print (Floor_OB1)
			print (Floor_OB1_P2)
			#print (Previous_Check)
			#sleep(99999)


			while Previous_Check == 2:

				

				Letter_Left_To = ALPHABET.index(Floor_OB1_P2[0][0]) #Will represent the first Letter/Column originally in the Floor OB1 P2 List
				Letter_Left_To = ALPHABET[Letter_Left_To - 1] 		##Will represent the second Letter/Column Left in original Floor OB1 PT List
				#This will represent the letter of the column just before the current one in Floor OB1 P2 List. It will be used to accurately find Rotation pushing post in the previous rotation hour


				#This var will keep count of how many Post will start rotations on OB1
				#Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates_2)
				#Nothing was wrong with this Code, Moving it else where per AI


				#This will create multiple list inside this list. Each list inside, will hold a pattern of post that start the rotation, and post that are being relieved. It will also be used later to make sure every post has been relieved
				#Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]
				#Relieved_Post = [[]] #Just adding 1 list inside this list of list. Not enough effort to warrant the above code

				#Trying this list creation code instead
				# Create exactly as many slots as there are people starting rotations
				#Relieved_Post = [[] for _ in range(len(Post_Pushing_Rotations_Cordinates_2))]
				#Nothing was wrong with this Code, Moving it else where per AI

				#All Relieved Cords will be in here. Will help prevent confusion with post being Relieved
				#Relieved_Cords = []
				#Nothing was wrong with this Code, Moving it else where per AI

				Stop_Var = 0
				#Can delete later, only used for testing

				Post_Starter_Already_Relieved = []
				#will be a list consisted of cords that have already started rotations and their full rotation line has been completed

				OB1_PT2_Floor_Rows_Relief_List = []
				#Will be the official relief list for OB1 Rotation check down below

				#Creating the list of cords only to be used to check for relief on OB1 and OB2
				for i in range(int(Locate('OB1')[0][1:]), int(Locate('OB3')[0][1:])):
					if ws[Floor_OB1_P2[0][0] + str(i)].value == None or ':' not in ws[Floor_OB1_P2[0][0] + str(i)].value and 'SKIP' not in ws[Floor_OB1_P2[0][0] + str(i)].value or 'FREIGHT' in ws[OB1_OB2_Floor_Rows[0][0] + str(i)].value:
						OB1_PT2_Floor_Rows_Relief_List.append(Floor_OB1_P2[0][0] + str(i))

				B1_Cord = 'Empty for Now'
				# Will contain the B1 Cord we use if ever borrowed to complete the upstairs post.

				OB1_Row = int(Locate('OB1')[0][1:])
				#This var will represent what row we should be looking for if any cords are borrowed down on B1

				print (OB1_Row)

				'''
				Lazy Coding: For some reason the short OB1 List of Cords in Post_Pushing_Rotations_Cordinates_2 List is sometimes showing up empty.
				Will add an if statement, that if the list is empty, go through Floor_OB1_P2 list, which contains every cords for the half hour of OB1 List
				and just add the cord that starts the rotation to Post_Pushing_Rotations_Cordinates_2. I truly don't know where the issue 
				lies at, to lazy to look into it.

				Also it only happens sometimes anyways. Usually when we borrow a cord from B1 to help push rotations from OB1, and that B1 Cord starts the rotations. Code below
				

				#IMPORTANT: This code actually belongs in the for loop below or after the for loop
				if len(Post_Pushing_Rotations_Cordinates_2) == 0:
					print ('Test: ', Floor_OB1_P2)
					for i in Floor_OB1_P2:
						# Get the cell to the immediate LEFT (0 rows down, -1 column right)
						cell_to_left = ws[i].offset(row=0, column=-1)
						print (cell_to_left)
					sleep(99999)
				'''


				#This code below searches for any cordinates that were forgotten, we may have written over, replaced, and borrowed from B1.
				for i in OB1_Tier_1:
					print ('------------------')
					print ('OB1 Floor Rows Relief Cord List: ', OB1_PT2_Floor_Rows_Relief_List)
					print ('Post Starter on OB1 Cords List: ', Post_Pushing_Rotations_Cordinates_2) #This could be the answer, this list isnt updated
					print (Floor_OB1_P2)
					print (i)
					print ('Cords with this Post: ', Locate(i)) # Produces a List
					for a in (Locate(i)):
						if int(a[1:]) < OB1_Row and a[0] == Floor_OB1_P2[0][0] and a not in Floor_OB1_P2: # OB1_PT2_Floor_Rows_Relief_List[0][0]:
							OB1_PT2_Floor_Rows_Relief_List.append(a)
							Floor_OB1_P2.append(a)
							break
							

				if len(Post_Pushing_Rotations_Cordinates_2) == 0:
					for a in Floor_OB1:
						if Temporary_Value(a, ws[a].value) in ROTATION_STARTERS:
							Post_Pushing_Rotations_Cordinates_2.append(a)
							print (Temporary_Value(a, ws[a].value))
					print ('')
					#This error seems fixed. No longer endlessly loops
					#sleep(5) #Dont forget to delete this sleep




					#print ('Testing New Code... Sleeping-')
					#sleep(99999)
					#Also gonna have this in an if statement and base it on a condition that if we used a B1 Post, then add this cord


				# Adding the B1 Cord we may have borrowed, only if we borrowed for upstairs post. Forgot to hard code this in earlier in the code.
				# 2 Was the option where we replaced a not important B1 Post with upstairs post for the hour due to lack of staff.
				if len(What_Should_We_Do) > 0:
					if What_Should_We_Do[0] == 2:

						# The B1 Letter that we grabbed was off by 1 column when ever we use a B1 Cord. Offset is how we fix it.
						column_index = column_index_from_string(OB1_PT2_Floor_Rows_Relief_List[0][0])
						left_index = column_index - 1
						left_letter = get_column_letter(left_index)

						Newly_Added_B1_Cord = OB1_PT2_Floor_Rows_Relief_List[0][0] + Celeb_B1_Cord[0][1:]
						Newly_Added_B1_Cord_Updated = left_letter + Celeb_B1_Cord[0][1:]
						OB1_PT2_Floor_Rows_Relief_List.append(Newly_Added_B1_Cord)
						if Temporary_Value(Newly_Added_B1_Cord_Updated, ws[Newly_Added_B1_Cord_Updated].value) in ROTATION_STARTERS:
							Post_Pushing_Rotations_Cordinates_2.append(Newly_Added_B1_Cord_Updated)
						

				print ('')
				print (OB1_PT2_Floor_Rows_Relief_List)


				print ('Work on Forever loop below.... Could be because the letters in OB1_PT2 list isnt correct, might have to change it to j instead?')
				print ('Comment out sleep and run just to see... also made a sleep 3 statement down below')
				#sleep(9999)


				# PASTE THE INITIALIZATION HERE: If this doesn't work, comment out code below and uncomment 7562. 
				# Other way copy the code below, find, and uncomment these same 3 lines of code if its easier that way/
				Rotation_Starting_Post = len(Post_Pushing_Rotations_Cordinates_2)
				Relieved_Post = [[] for _ in range(Rotation_Starting_Post)]
				Relieved_Cords = []

				#Looping through all the starting post in the OB1 Post starter list.
				#for Post_Starter in Post_Pushing_Rotations_Cordinates_2: #Old Code

				# AI Assisted Code
				for idx, Post_Starter in enumerate(Post_Pushing_Rotations_Cordinates_2):
					# This ensures that Starter #1 always uses List #1, Starter #2 uses List #2, etc.
					Relieved_Post_Inner_List_Var = idx
					print (Relieved_Post_Inner_List_Var)
					print (len(Relieved_Post))
					sleep(0)
					print ('-------------------------------------------------------------')
					print('')
					print ('')

					
					
					#-b1relief
					if EVERYTHING_ELSE[1] == Temporary_Value(Post_Starter, ws[Post_Starter].value) or '**' in Temporary_Value(Post_Starter, ws[Post_Starter].value) or Temporary_Value(Post_Starter, ws[Post_Starter].value) in ROTATION_STARTERS:
						#wb.save(File_Name)
						Rotation_Complete = 0 #Var used to help end the while Loop below. It'll become 1 at the end of every complete rotation

						if Using_Celebrate_From_B1 > 0:
							for i in Floor_OB1_P2:
								if i not in OB1_PT2_Floor_Rows_Relief_List:
									OB1_PT2_Floor_Rows_Relief_List.append(i)
									#This code is literally to use the B1 Code, and to still check and see if we need to relieve someone down there. For some reason that B1 Code is being left out.

						while Rotation_Complete == 0: #Should be connected to the amount of post rotation starters there are in this hour rotation
							print ('Starting With: ', ws[Post_Starter].value)
							print ('Cord of Post Starter: ', Post_Starter)
							Post_Starter_Already_Relieved.append(Post_Starter)
							print ('Full List of Post Starters: ', Post_Pushing_Rotations_Cordinates_2)
							print ('Inner List #: ', Relieved_Post_Inner_List_Var)
							print ('OB1 Floor Relief List: ', OB1_PT2_Floor_Rows_Relief_List) #This version of the list will use the Letter to the right.
							print ('Showing Current List Progress: ', Relieved_Post)
							print ('Relieved Post Inner List Var: ', Relieved_Post_Inner_List_Var)
							print ('------')
							print ('')
							#if OB1_OB2_Floor_Rows[0][0] == 'P':
							#	sleep(2)
							
							

							for Floor_Cords in OB1_PT2_Floor_Rows_Relief_List: #New List will go here 
								cell2 = ws[Floor_Cords].value                 		#Will be the value of the Floor's Cord List Cells which is the most recent list in the excel and also printed out in the terminal when its ran
								cell = ws[Letter_Left_To + Floor_Cords[1:]].value #Will be the value of the 2nd cell directly to the left of the most filled in Column



								#if OB1_OB2_Floor_Rows[0][0] == 'P':
								#	sleep(1)
								print (Letter_Left_To + Floor_Cords[1:])
								print (Post_Starter)
								
								print ('----------------------READ THIS-------------------------------')
								print ('If it Endlessly Loops... its because cord isnt in the cord list in the for loop.... P2')
								print ('Main Relief List: ', OB1_PT2_Floor_Rows_Relief_List)
								print ('Smaller OB1 Relief List: ', Floor_OB1_P2)
								print ('Relieved Post Var: ', Relieved_Post_Inner_List_Var)
								print ('Relieved Post List: ', Relieved_Post)
								print ('Post Starter Cords List: ', Post_Pushing_Rotations_Cordinates_2)
								print ('----------------------READ THIS-------------------------------')
								#This is to prevent the code from relieving any Post that isnt a rotator or has an atrisk. 
								#Slight error someone here

								# Added a boundary check: if the index doesn't exist OR the inner list is empty
								if (Relieved_Post_Inner_List_Var == len(Relieved_Post) or len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0) and (Letter_Left_To + Floor_Cords[1:]) != Post_Starter:
									#If this works, comment out the prints and just put pass here
									print ('')
									print ('Post Starter Cord: ', Post_Starter)
									print ('Current Cord: ', Letter_Left_To + Floor_Cords[1:])
									print ('Moving On....')
									print ('')
									continue

								if (Letter_Left_To + Floor_Cords[1:]) == Post_Starter or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1: 
								
									
									print ('')
									Post_Starter_Value = Temporary_Value(Post_Starter, ws[Post_Starter].value)
									#print (Post_Starter_Value)
									#print (EVERYTHING_ELSE[1])
									print ('-----')
									sleep(0)
									#This Var will be used to capture the value of any post starting rotations, so it doesnt give a None error.

									

									if '*' in Post_Starter_Value or Post_Starter_Value == EVERYTHING_ELSE[1]: #Change this to make it more secure maybe?; If statement may be useless here honestly
										
										#Irrelevent Prints: Comment out at some point
										if Relieved_Post_Inner_List_Var < len(Relieved_Post):
										    print ('Currently Relieved: ', Relieved_Post[Relieved_Post_Inner_List_Var])
										else:
										    print ('Currently Relieved: [No active posts / Empty]')
										

										#Has_Post_Been_Relieved = any(cell in sublist for sublist in Relieved_Post) 
										# Check if the physical source cell was already relieved
										source_coordinate = Letter_Left_To + Floor_Cords[1:]
										Has_Post_Been_Relieved = source_coordinate in Relieved_Cords
										#Will be used to check inside of all the list inside Relieved Post list of list and see if that post is already accounted for

										#Was Here Last!!!!!!!!!!!!!11 - Delete Later
										cell2 = Temporary_Value(Floor_Cords, cell2)
										cell = Temporary_Value(Letter_Left_To + Floor_Cords[1:], cell)

										print (cell + ' Relieved Status: ', Has_Post_Been_Relieved)
										

										

										#This if statement is always for the Post Starter cell and also the post that come after that may not start the rotation, ut will keep it going. Thats why the or statement is for if the list has a higher value than 1
										if cell in ROTATION_STARTERS or len(Relieved_Post[Relieved_Post_Inner_List_Var]) >= 1 or cell == EVERYTHING_ELSE[1]:
											'''
											# AI Says to Delete this. Will leave until we fully test the code.
											if '*' in cell and cell2 not in ROTATION_STARTERS and (Relieved_Post_Inner_List_Var >= len(Relieved_Post) or (cell not in Relieved_Post[Relieved_Post_Inner_List_Var] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0)):
												if Relieved_Post_Inner_List_Var >= len(Relieved_Post): 
													Relieved_Post.append([cell, cell2])
													Relieved_Cords.append(Floor_Cords)
													print ('AI Assisted Solution: See if it works')
											        #This list should be 0 regardless since its only meant to check if the ob1 breakers are being relieved properly.
												    #Relieved_Post.append([cell]) #Delete this line later if it works

												    # If list doesn't exist, create a new inner list with BOTH cells inside it.
												else:
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) #Relieved_Post_Inner_List_Var - Determines what list we are using first inside of Relieved_Post; Starting with 0
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)
													print ('if statement 1')
												'''

											# AI Assisted Solution
											if '*' in cell and cell2 not in ROTATION_STARTERS and (len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0):
											    Relieved_Post[Relieved_Post_Inner_List_Var].append(cell)
											    Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
											    Relieved_Cords.append(Floor_Cords)
											    print ('if statement 1')

											# Before: len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1 has this at the end
											#Changed to: Relieved_Post_Inner_List_Var since this var should represent the amounf of list thats in Relieved_Post anyway
											elif cell in ROTATION_STARTERS and cell2 in ROTATION_STARTERS and len(Relieved_Post[Relieved_Post_Inner_List_Var]) < 1:
												sleep(0)
												print ('if statement 2')
												if cell in EVERYTHING_ELSE: #If its a full BREAK.... kinda thinking this code is useless too because of the at risk if statement
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)

													# AI Assisted Solution: Only move to the next rotation line if we actually successfully relieved someone
													if len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 0:
													    Relieved_Post_Inner_List_Var += 1
													    Rotation_Complete += 1
													else:
													    # Optional: print for debugging so you see it failing
													    print(f"Index tried to move from {Relieved_Post_Inner_List_Var}, but list was empty. Staying at {Relieved_Post_Inner_List_Var}.")
													
													# Old Code Below: Delete later. Newer Code Above only adds 1 if an item has been added to the list.
													#Relieved_Post_Inner_List_Var += 1
													#Rotation_Complete += 1
												elif cell not in EVERYTHING_ELSE: #Basically if its LAUNCH or has an at risk
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell) 
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)
													
													print (Relieved_Post[Relieved_Post_Inner_List_Var])
													


													print ('This ROTATION is Ending EARLY')
													print (cell)
													print (cell2)
													print (Relieved_Post)
													print ('Cell2 Cord: ', Floor_Cords)
													print ('Cell Cord: ', Letter_Left_To + Floor_Cords[1:])
													#sleep(9999)
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													Stop_Var = 0
													print ('Adding Relieved Post Var + 1')
													print('')
													print('')
													print('')
													print ('if statement: 3')
													break
												elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and '*' in cell2 and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)
													print ('Ending this Rotation--')
													

													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													#xy10
												elif cell in Relieved_Post[Relieved_Post_Inner_List_Var] and cell2 not in Relieved_Post[Relieved_Post_Inner_List_Var] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) > 1:
													print (cell + 'Is Relieving ' + cell2)
													Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
													Relieved_Cords.append(Floor_Cords)
													print ('if statement 4')
											elif cell not in ROTATION_STARTERS and cell == Relieved_Post[Relieved_Post_Inner_List_Var][-1]:
												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Floor_Cords)
												print (cell + 'Is Relieving ' + cell2)
												print ('if statement 5')

												if cell2 in ROTATION_STARTERS or '*' in cell2:
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													print ('Ending this rotation')
													break
											elif cell == EVERYTHING_ELSE[1] and len(Relieved_Post[Relieved_Post_Inner_List_Var]) == 0: 
											#This will be used for BREAKER SHIFTS after their BRIEFING

												Relieved_Post[Relieved_Post_Inner_List_Var].append(cell2)
												Relieved_Cords.append(Floor_Cords)
												print (cell + 'Is Relieving ' + cell2)
												print ('if statement 6')

												if cell2 in ROTATION_STARTERS or '*' in cell2:
													Relieved_Post_Inner_List_Var += 1
													Rotation_Complete += 1
													#This if statement ends Rotations super early

												elif cell2 not in ROTATION_STARTERS or '*' not in cell2:
													#Relieving a regular post after briefing, this if statement passes it on to the next if statement
													pass
								

												





							#Used to end the while loop because the Relieved_Post_Inner_List_Var will represent a list value that doesnt exist. 
							#Check All Post Here
							if Relieved_Post_Inner_List_Var == len(Relieved_Post):
								print ('----------------------------------')
								print('')
								print ('Relieved Vars value has passed the amount of list we created. Ending Rotation early to force a re-rotate')
								print ('Showing Current List Progress: \n', Relieved_Post)
								print ('')
								print ('Relief Cords List: ', OB1_PT2_Floor_Rows_Relief_List)
								#sleep(99999)
								Rotation_Complete += 1




								#---Checking Rotation Code Here---



								#This will turn Relieved Post list into a regular list, so it'll be easier to compare the amount of items to the tier list to break the while loop.
								Collapsed_List_Of_Relieved_Post = [item for sublist in Relieved_Post for item in sublist]
								print ('Before Filter: ', Collapsed_List_Of_Relieved_Post)

								#We have to filter this list, to remove any post that are rotation starters besides the on that belongs to a Tier 1 List
								Junk_List = [] #List thats going to have post that we dont want
								for Post in Collapsed_List_Of_Relieved_Post:
									if Post in FREIGHT or Post not in OB1_Tier_1: 
										Junk_List.append(Post)

								print ('Junk List: ', Junk_List) 

								for Junk_Post in Junk_List:
									Collapsed_List_Of_Relieved_Post.remove(Junk_Post)

								Collapsed_List_Of_Relieved_Post = list(set(Collapsed_List_Of_Relieved_Post))
								#Filtering for Duplicates and Deleting Them

								print('')
								print ('After Filter: ', Collapsed_List_Of_Relieved_Post) 
								print ('# of Post Relieved: ', len(Collapsed_List_Of_Relieved_Post)) #Should be 4 Based on the sum of OB1 Tier 1. To be sure just add both OB1
								#OB1_OB2_Tier_1_List = OB1_Tier_1 + OB2_Tier_1
								print ('OB1 Tier 1 List: ', OB1_Tier_1) #List of Important Post that need to be Relieved!!!

								Unrelieved_Post = list(set(OB1_Tier_1) - set(Collapsed_List_Of_Relieved_Post))
								print ('Tier 1 Post Hasnt Been Relieved: ', Unrelieved_Post)
								#This is a list that will print every important Tier 1 Post that hasnt been relieved


								if len(Unrelieved_Post) == 0:
									print ('Everything Checks Out for the 1st Half of Rotations....')
									print ('Times Redid Rotations: ', Restart_While_Loop)
									Restart_While_Loop = 0
									Previous_Check += 1
									Checks += 1
									OB1_Hour_Post += 1
									print ('Delete me later if found...')
									#sleep(99999)
									break
									#This will break the main while Loop and continue on to the next important part of this Function
								elif len(Unrelieved_Post) != 0:
									print ('Resetting OB1s Second Half Officially and Re-Doing Everything.....')
									'''
									-Reset Checks and Post- 

									This will reset it all the way back to the main while loop, that way it recreates the post and shuffles the order so
									we get a different outcome everytime. Easier that way

									'''


									Previous_Check += 1 #Can Possibly just delete this, and replace OB1 Post Creation Code here for the Second Half of the Hour
									Restart_While_Loop += 1
									Relieved_Post_Inner_List_Var = 0
									Relieved_Post.clear()
									break
								






								#sleep(99999)
								#break


					# ... (End of the `while Rotation_Complete == 0:` block)

					# ADD THIS HERE: 
					# If the success block triggered, Previous_Check is now 3. 
					# Break out of the Post_Starter loop to stop the IndexError crash!
					if Previous_Check > 2:
						break





	print ('')
	print ('Repeated Post That Were Fixed...')
	print ('Repeated Post: ', Fixed_Repeated_Post)
	print ('-=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=--=++=-')
	print ('')
	print ('')
	print ('')
	print ('')
	print ('')
	print ('')


	print ('')
	print ('Function is Officially Done....')
	#sleep(88888)
		








print ('')

#Counting Empty Cells from OB1 - OB2
OB1_OB2_Floor_Rows = []
FloorCellCount('F', OB1_OB2_Floor_Rows, str(ws.max_row), int(Locate('OB1')[0][1:]))
Upper_Floor_Rotation_Creation_Both_Floors()

#Counting Empty Cells from OB1 - OB2
OB1_OB2_Floor_Rows = []
FloorCellCount('H', OB1_OB2_Floor_Rows, str(ws.max_row), int(Locate('OB1')[0][1:]))
Upper_Floor_Rotation_Creation_Both_Floors()
#If ever an Issue, comment this block out and uncomment 4352 - 4355
print ('Column H is Done')
#sleep(99999)

OB1_OB2_Floor_Rows = []
FloorCellCount('J', OB1_OB2_Floor_Rows, str(ws.max_row), int(Locate('OB1')[0][1:]))
Upper_Floor_Rotation_Creation_Both_Floors()
print ('Column J is Done')
print ('break this down column by column and see why its choosing to grab b1 post and also dumplicate aff3 post')



OB1_OB2_Floor_Rows = []
FloorCellCount('L', OB1_OB2_Floor_Rows, str(ws.max_row), int(Locate('OB1')[0][1:]))
Upper_Floor_Rotation_Creation_Both_Floors()
print ('Column L is Done')


OB1_OB2_Floor_Rows = []
FloorCellCount('N', OB1_OB2_Floor_Rows, str(ws.max_row), int(Locate('OB1')[0][1:]))
Upper_Floor_Rotation_Creation_Both_Floors()
print ('Column N is Done')




OB1_OB2_Floor_Rows = []
FloorCellCount('P', OB1_OB2_Floor_Rows, str(ws.max_row), int(Locate('OB1')[0][1:]))
Upper_Floor_Rotation_Creation_Both_Floors()
print ('Column P is Done')


OB1_OB2_Floor_Rows = []
FloorCellCount('R', OB1_OB2_Floor_Rows, str(ws.max_row), int(Locate('OB1')[0][1:]))
Upper_Floor_Rotation_Creation_Both_Floors()
print ('Column R is Done')
wb.save(File_Name)



	

#Going through ALL CELLs and adding its proper color.
for row in ws.iter_rows():
    for cell in row:
    	#Filling in FLOAT cells with the proper Blue
        if cell.value == EVERYTHING_ELSE[0]:
        	FLOAT_CELL = PatternFill(patternType = 'solid', fgColor = FLOAT)
        	cell.fill = FLOAT_CELL
        elif cell.value in OB1_Tier_1 or cell.value in OB1_Tier_2:
        	OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
        	cell.fill = OB1_Cell






#Writing in Clear for All Post on OB1 - OB2... not OB3 Yet

Clear_Cords_List = []
#This will be a list of cords that will get clear post at the end of the night...

Start = int(Locate('OB1')[0][1:]) 
End = int(Locate('OB3')[0][1:])
for i in range(Start, End):
	if ws['S' + str(i)].value == None:
		pass
	elif ws['S' + str(i)].value in Overall_OB1_Post:
		Clear_Cords_List.append('T' + str(i))

for i in Clear_Cords_List:
	index = ALPHABET.index(i[0]) + 1
	Create_Post(i, int(i[1:]), index, 'CLEAR')
	CLEAR_CELL = PatternFill(patternType = 'solid', fgColor = CLEAR_COLOR)
	ws[i].fill = CLEAR_CELL



#Giving Leads a Float at the End of the Night
for i in Leads_Upstairs:
	if ws['T' + i[1:]].value == None:
		Cell = 'T' + i[1:]
		index = ALPHABET.index('T') + 1
		Create_Post(Cell, int(i[1:]), index, 'FLOAT')
		FLOAT_CELL = PatternFill(patternType = 'solid', fgColor = FLOAT)
		ws[Cell].fill = FLOAT_CELL
		wb.save(File_Name)



print ('Done Before Merge')
sleep(99999)




#Fixing the 3 Cell - Merge of the Floor Names B1, OB1 etc
Floor_Names = ['B1', 'OB1', 'OB2', 'OB3']
Start = int(Locate('B1')[0][1:]) - 1
End = int(Locate('OB3')[0][1:]) + 1
for i in range(Start, End):
	if ws['A' + str(i)].value in Floor_Names:
		Cell = 'A' + str(i) + ':C' + str(i)
		index = ALPHABET.index('A') + 2
		Create_Post(Cell, int(i), 1, ws['A' + str(i)].value)






#Newly Added Code is for Full Timers upstairs that start at 3:30. A 4PM Post will be created for them.
random.shuffle(Full_Timers_4PM_Post_On_B1)
random.shuffle(Full_Timers_4PM_Post_Upstairs)
num = 0
num2 = 0
for a in Full_Time_Closers_Row:
	if int(a[1:]) < int(Locate('OB1')[0][1:]): #This is a B1 Full Timer
		ws['E' + a[1:]].value = Full_Timers_4PM_Post_On_B1[num]
		ws['E' + a[1:]].border = border
		ws['E' + a[1:]].alignment = Center_Text
		num = num + 1
	elif int(a[1:]) > int(Locate('OB1')[0][1:]): #This is a Upstairs Full Timer
		ws['E' + a[1:]].value = Full_Timers_4PM_Post_Upstairs[num2]
		ws['E' + a[1:]].border = border
		ws['E' + a[1:]].alignment = Center_Text
		# Save the changes to the workbook
		wb.save(File_Name)

		#if Full_Timers_4PM_Post_Upstairs[num2] in Overall_OB2_Post:
		#	ws['E' + a[1:]].fill = OB2_Cell
		#elif Full_Timers_4PM_Post_Upstairs[num2]:
		#	ws['E' + a[1:]].fill = OB1_Cell



		num2 = num2 + 1




#Over Writing the Dark Box for Full Timers and Adding BRIEF
Start = 1
End = int(Shift_Rows[-1][1:]) + 1
print ('')
for i in range(Start, End):
	print (i)
	print (ws['B' + str(i)].value)
	if ws['B' + str(i)].value == None:
		pass
	elif '3:30' in ws['B' + str(i)].value:
		ws['D' + str(i)].value = EVERYTHING_ELSE[1]
		ws['D' + str(i)].fill = Brief
		ws['D' + str(i)].border = border
		ws['D' + str(i)].alignment = Center_Text
		
		

		

#Fills in the color for any cell that may be missing some for OB1 and OB2
for row in ws.iter_rows():
    for cell in row:
        if cell.value == 'FLOAT':
        	FLOAT_CELL = PatternFill(patternType = 'solid', fgColor = FLOAT)
        	cell.fill = FLOAT_CELL
        	wb.save(File_Name)
        if cell.value == 'OB3 MOMENTO': #Make OB3 List just for the sake of the color code and add it here
        	OB3_Cell = PatternFill(patternType = 'solid', fgColor = OB3)
        	cell.fill = OB3_Cell
        elif cell.value in B1_Tier_1:
        	B1_Cell = PatternFill(patternType = 'solid', fgColor = B1)
        	cell.fill = B1_Cell
        	wb.save(File_Name)
        elif cell.value == 'PREP**':
        	PREP_CELL = PatternFill(patternType = 'solid', fgColor = PREP)
        	cell.fill = PREP_CELL
        	wb.save(File_Name)
        elif cell.value in Overall_OB1_Post:
        	OB1_Cell = PatternFill(patternType = 'solid', fgColor = OB1)
        	cell.fill = OB1_Cell
        	wb.save(File_Name)
        elif cell.value in Overall_OB2_Post:
        	OB2_Cell = PatternFill(patternType = 'solid', fgColor = OB2)
        	cell.fill = OB2_Cell


#---Mega Cell Merge---
#Fix Later: Code Below Will be Merging All Cells next to Each Other that are the exact Same..










print ('List Of Things to Fix: ')
print ('Also make the floor switch from OB2 on OB1 Floor start at L or N based on randomness')
print ('Also have some 4 - 11s Clear')
print ('Add OB3 Momento to force us to code in short staff solutions')

print ('Done')


wb.save(File_Name)
    

