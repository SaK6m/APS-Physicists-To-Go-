import pandas as pd
import os
# from pandas import ExcelWriter, notnull
import xlsxwriter

print("------------------Welcome to the Program-----------------------")
print("Please add your extension when you enter your file name (like: filename.xlsx)")

print("---------------------------------------------------------------")

ALLOWED_EXTENSION = {".xlsx"}
#excel file name with extension
while True:
    teacher_file_name = input("File name of Teacher: ")
    extension = os.path.splitext(teacher_file_name)[1]
    if extension in ALLOWED_EXTENSION:
        break
    else:
        print("Check your file extension")

while True:
    physicists_file_name = input("File name of Physicists: ")
    extension = os.path.splitext(physicists_file_name)[1]
    if extension in ALLOWED_EXTENSION:
        break
    else:
        print("Check your file extension")

while True:
    out_file = input("What would you like to name your matched excel file: ")
    extension = os.path.splitext(out_file)[1]
    if extension in ALLOWED_EXTENSION:
        break
    else:
        print("Check your file extension")

#---- from teacher's excel sheet ----
teacher_name_header = 'Name'
teacher_topic_header = 'Topics'
teacher_email_header = 'Email'
teacher_school_header = 'School Level'
teacher_location_header = 'Location'
teacher_month_header = 'Month'
teacher_meeting_header = 'Meeting Preference'
teacher_scilvl_header = 'Science Level'
#------ from physicists' excel sheet ----
physicists_name_header = 'Name'
physicists_topic_header = 'Topics'
physicists_email_header = 'Email'
physicists_school_header = 'School Level'
physicists_location_header = 'Location'
physicists_meeting_header = 'Meeting Preference'

#----- start of program ------
teachers_df = pd.read_excel(teacher_file_name) 
physicists_df = pd.read_excel(physicists_file_name) 

def topic_filter(topic):
  matched = [
      True if topic in physicists_topic.split(', ') else False
      for physicists_topic in physicists_df[physicists_topic_header] 
    ]

  return pd.Series(matched)

def location_filter(location):
    matched  = [
        True if location in matched_location.split(', ') else False
        for matched_location in matched_df[physicists_location_header] 
    ]

    return pd.Series(matched)

def school_filter(school):
    matched  = [
        True if school in matched_school.split(', ') else False
        for matched_school in location_df[physicists_school_header] 
    ]

    return pd.Series(matched)

def meeting_filter(meeting):
    matched  = [
        True if meeting in matched_meeting.split(', ') else False
        for matched_meeting in school_df[physicists_meeting_header] 
    ]

    return pd.Series(matched)

#----- list  for matched ----
teachers_name = []
teachers_email = []
physicists_name = []
physicists_email = []
common_topic = []
topic_p = []
topic_t = []
timezone_p = []
timezone_t = []
school_t = []
school_p =[]
matched_month = []
matched_meeting =[]
meeting_p =[]
priority = []
science_level = [] 

#------ list for unmatched teacher df ----
unmatched_teacher = []
email_unmatched_teacher = []
unmatched_teachers_topic = []
unmatched_month = []
unmatched_teachers_meeting = []
reason = []
unmatched_timezone_t =[]
unmatched_school_t = []
unmatched_scilvl =[] #####

#----- list for unmatched physicists df ---
unmatched_physicists = []
email_unmatched_physicists = []
unmatched_physicists_topic = []
unmatched_physicists_meeting = []
unmathched_p_timezone = []
unmatched_p_school =[]

count_row = teachers_df.shape[0]
count_column = teachers_df.shape[1]

#----- sorting teacher list by number of NaN ---- 
teachers_df['#columns'] = teachers_df.count(axis = 1)
teachers_df = teachers_df.sort_values('#columns', ascending = True)

teachers_df  = teachers_df.fillna("BLANK")

print('Matching in progress......')
#---- loop this through rows -----

for row in range(count_row):
    topics = teachers_df.loc[row, teacher_topic_header].split(', ')
    months = teachers_df.loc[row, teacher_month_header].split(', ')
    x = len(topics)

    location = teachers_df.loc[row, teacher_location_header] 
    school = teachers_df.loc[row,teacher_school_header]
    meeting = teachers_df.loc[row, teacher_meeting_header]

    name_teacher = teachers_df.loc[row, teacher_name_header] 
    email_teacher = teachers_df.loc[row, teacher_email_header] 

    #--step 1: topics check--
    for n in range(x):
        topic_matched = topic_filter(teachers_df.loc[row, teacher_topic_header].split(', ')[n])
        list_end = x - n
        if topic_matched.eq(False).all() and list_end == 1: # print when list end
            unmatched_teacher.append(name_teacher) 
            reason.append(str("No matching Topic"))
            email_unmatched_teacher.append(email_teacher)
            unmatched_teachers_topic.append(topics[n])
            unmatched_teachers_meeting.append(meeting) #
            unmatched_timezone_t.append(location)
            unmatched_school_t.append(school)
            unmatched_month.append(months[n])
            unmatched_scilvl.append(teachers_df.loc[row, teacher_scilvl_header])
        elif topic_matched.eq(False).all(): # move to next list
            print("..")
        else:

            matched_df = physicists_df[topic_matched].reset_index(drop=True)

            #---- sorting  topics  --- 
            matched_df['#Topics'] = matched_df[physicists_topic_header].apply(lambda n: len(n.split(', ')))
            matched_df  = matched_df.sort_values('#Topics',  ascending = True).reset_index(drop=True)
            #--step 2: location checkcl

            location_matched = location_filter(location)

            if location_matched.eq(False).all():
                unmatched_teacher.append(name_teacher) 
                reason.append(str("No matching Location"))
                email_unmatched_teacher.append(email_teacher)
                unmatched_teachers_topic.append(topics[n])
                unmatched_teachers_meeting.append(meeting)
                unmatched_timezone_t.append(location)
                unmatched_school_t.append(school)
                unmatched_month.append(months[n])
                unmatched_scilvl.append(teachers_df.loc[row, teacher_scilvl_header])
            else:
                location_df = matched_df[location_matched].reset_index(drop=True)

                #-- step3: school level check-- 
                school_matched = school_filter(school)

                if school_matched.eq(False).all():
                    unmatched_teacher.append(name_teacher) 
                    reason.append(str("No matching school level"))
                    email_unmatched_teacher.append(email_teacher)
                    unmatched_teachers_topic.append(topics[n])
                    unmatched_teachers_meeting.append(meeting)
                    unmatched_timezone_t.append(location)
                    unmatched_school_t.append(school)
                    unmatched_month.append(months[n])
                    unmatched_scilvl.append(teachers_df.loc[row, teacher_scilvl_header])
                else:
                    school_df = location_df[school_matched].reset_index(drop=True)

                    school_df['#school'] = school_df[physicists_school_header].apply(lambda n: len(n.split(', ')))
                    school_df  = school_df.sort_values('#school',  ascending = True).reset_index(drop=True)

                    meetings  =teachers_df.loc[row, teacher_meeting_header].split(', ') ## one more to separate  by comma
                    y = len(meetings)

                    for m in range(y):
                        meeting_matched = meeting_filter(teachers_df.loc[row, teacher_meeting_header].split(', ')[m])
                        end = y - m
                        if meeting_matched.eq(False).all() and end == 1:
                            unmatched_teacher.append(name_teacher) 
                            reason.append(str("No matching meeting preference"))
                            email_unmatched_teacher.append(email_teacher)
                            unmatched_teachers_topic.append(topics[n])
                            unmatched_teachers_meeting.append(meeting)
                            unmatched_timezone_t.append(location)
                            unmatched_school_t.append(school)
                            unmatched_month.append(months[n])
                            unmatched_scilvl.append(teachers_df.loc[row, teacher_scilvl_header])
                        elif meeting_matched.eq(False).all():
                            print('.')
                        else:
                            meeting_df = school_df[meeting_matched].reset_index(drop=True)

                            #---sort meeting ---
                            meeting_df['#meeting'] = meeting_df[physicists_meeting_header].apply(lambda n: len(n.split(', ')))
                            meeting_df = meeting_df.sort_values('#meeting', ascending = True).reset_index(drop=True)
                        
                            #--- adding to the matched list-- 
                            teachers_name.append(name_teacher)
                            teachers_email.append(email_teacher)
                            topic_t.append(teachers_df.loc[row, teacher_topic_header])
                            topic_p.append(meeting_df.loc[0, physicists_topic_header])
                            common_topic.append(topics[n])
                            matched_month.append(months[n])
                            physicists_name.append(meeting_df.loc[0, physicists_name_header]) 
                            physicists_email.append(meeting_df.loc[0, physicists_email_header]) 
                            matched_meeting.append(meetings[m])  ###
                            priority.append(n + 1)
                            timezone_p.append(meeting_df.loc[0, physicists_location_header])
                            timezone_t.append(location)
                            school_p.append(meeting_df.loc[0, physicists_school_header])
                            school_t.append(school)
                            science_level.append(teachers_df.loc[row, teacher_scilvl_header])
                            meeting_p.append(meeting_df.loc[0, physicists_meeting_header])
                
                            #--deleteing name of selected physicists --
                            email_check = meeting_df.loc[0, physicists_email_header]
                            physicists_df = physicists_df[~physicists_df.select_dtypes(['object']).eq(email_check).any(1)].reset_index(drop=True)
                            break
            break

# -----adding unmatched physicists list ------
for item in range(physicists_df.shape[0]):
    unmatched_physicists.append(physicists_df.loc[item, physicists_name_header])
    email_unmatched_physicists.append(physicists_df.loc[item, physicists_email_header])
    unmatched_physicists_topic.append(physicists_df.loc[item, physicists_topic_header])
    unmatched_physicists_meeting.append(physicists_df.loc[item, physicists_meeting_header])
    unmathched_p_timezone.append(physicists_df.loc[item, physicists_location_header])
    unmatched_p_school.append(physicists_df.loc[item, physicists_school_header])

print('\nProcess Completed!\n')
# import pdb; pdb.set_trace()
#----- writing excel----
print ('wrting into a excel file......\n')
title1 = ['Teacher_name','Teacher_email','Teacher Topic', 'Priority Number', 'Common Topics','Physicists Topic', 'Physicist_name', 'Physicist_email', 'Month Matched',  'Type of Visit(T)', 'Time Zone(T)', 'School Level(T)', 'Type of Visit(P)', 'Time Zone(P)', 'School Level(P)', 'Teacher Physics/Science class level']
title2 = ['Teachers Name', 'Email','Reason', 'Topic','Unmatched Month', 'Teacher Physics/Science class level', 'Type of Visit', 'Schoool Level','TimeZone']
title3 = ['Physicists Name', 'Email', 'Topic', 'Type of Visit', 'Time Zone', 'School Level']

#--- creating excel file---
workbook = xlsxwriter.Workbook(out_file) ###
worksheet1 = workbook.add_worksheet('matched group')
worksheet2 = workbook.add_worksheet('unmatched teachers')
worksheet3 = workbook.add_worksheet('unmatched physicists')

worksheet1.write_row('A1', title1)
worksheet1.write_column('A2', teachers_name)
worksheet1.write_column('B2', teachers_email)
worksheet1.write_column('C2', topic_t)
worksheet1.write_column('D2', priority)
worksheet1.write_column('E2', common_topic)
worksheet1.write_column('F2', topic_p)
worksheet1.write_column('G2', physicists_name)
worksheet1.write_column('H2', physicists_email)
worksheet1.write_column('I2', matched_month)
worksheet1.write_column('J2', matched_meeting)
worksheet1.write_column('K2', timezone_t)
worksheet1.write_column('L2', school_t)
worksheet1.write_column('M2', meeting_p)
worksheet1.write_column('N2', timezone_p)
worksheet1.write_column('O2', school_p)
worksheet1.write_column('P2', science_level)

worksheet2.write_row('A1', title2)
worksheet2.write_column('A2', unmatched_teacher)
worksheet2.write_column('B2', email_unmatched_teacher)
worksheet2.write_column('C2', reason)
worksheet2.write_column('D2', unmatched_teachers_topic)
worksheet2.write_column('E2', unmatched_month)
worksheet2.write_column('F2', unmatched_scilvl)
worksheet2.write_column('G2', unmatched_teachers_meeting)
worksheet2.write_column('H2', unmatched_school_t)
worksheet2.write_column('I2', unmatched_timezone_t)

worksheet3.write_row('A1', title3)
worksheet3.write_column('A2', unmatched_physicists)
worksheet3.write_column('B2', email_unmatched_physicists)
worksheet3.write_column('C2', unmatched_physicists_topic)
worksheet3.write_column('D2', unmatched_physicists_meeting)
worksheet3.write_column('E2', unmathched_p_timezone)
worksheet3.write_column('F2', unmatched_p_school)

workbook.close()

print('Excel  written!\n')

print("------------------------------\n")
print("------Matching Done!!--------\n")