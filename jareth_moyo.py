__author__ = 'jarethmoyo'
# imports of all necessary libraries
from recommendations import *
from Tkinter import *
from tkFileDialog import askopenfilenames
from tkMessageBox import *
import ttk
from xlrd import open_workbook,cellname


#Define global variables
student_dict=None
student=None
flag=False  # This means we are using userbased. If flag is true then we are using itembased
sim_func_used=0
endGrades=[]
#Other useful variables
letter_dig_grade={'A+':4.1,'A':4.0,'A-':3.7,'B+':3.3,'B':3.0,'B-':2.7,'C+':2.3,
                  'C':2.0,'C-':1.7,'D+':1.3,'D':1.0,'D-':0.5,'F':0.0}


# The function below returns a dictionary with the subject and grades of a student
def grades_dictionary(book_name):
    book=open_workbook(book_name)
    sheet=book.sheet_by_index(0)
    gradebook={}
    # Temporary lists to store our course codes,names and grades after iteration
    templist1=[]
    templist2=[]
    templist3=[]

    i=1
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            # 'A' represent course code
            if 'A' in cellname(row_index,col_index):
                # We don't care about the first row so we skip it
                if i==1:
                    i=i+1
                    continue
                a=sheet.cell(row_index,col_index).value
                a=a.encode('ascii')  # This is to get rid of the unicode character
                templist1.append(a)
            if 'B' in cellname(row_index,col_index):
                if i==2:
                    i=i+1
                    continue
                b=sheet.cell(row_index,col_index).value
                b=b.encode('ascii')
                templist2.append(b)
            if 'C' in cellname(row_index,col_index):
                if i==3:
                    i=i+1
                    continue
                c=sheet.cell(row_index,col_index).value
                c=c.encode('ascii')
                templist3.append(c)

    subject_list=[]
    # We now combine the course code and course name into one name
    for index, item in enumerate(templist2):
        subject_list.append(templist1[index]+' '+templist2[index])

    # We call a different function to convert letter grades to their numeric value
    grades_to_digits=convert_letter_to_dig(templist3)
    # We append the subject and numeric grade to the gradebook then return it
    for index,item in enumerate(subject_list):
        gradebook[item]=grades_to_digits[index]

    return gradebook


# This function creates a global student dictionary with all the past student grades
def student_dictionaries():
    global student_dict
    name=askopenfilenames()
    student_data=[]  # Will contain a list of tuples with selected student names and name of files
    student_info={}
    for i in range(len(name)):
        address=name[i]
        filename=""
        # We need to extract the name of the file here
        for i in range(len(address)-1,-1,-1):
            if address[i]=='/':
                break
            filename+=address[i]
        filename=filename[::-1].encode('ascii')
        student_name=filename.split('.')[0].encode('ascii')
        student_data.append((student_name, filename))

    #To create a dictionary with student name as key and values as dictionary with subjects and grades
    for student_name,filename in student_data:
        student_info.setdefault(student_name,{})
        student_info[student_name]=grades_dictionary(filename)

    print student_info.keys()
    student_dict=student_info


def merge_dictionaries():
    global student_dict,student
    #Similar code as above, although the point here would be to merge past student grades dictionary
    # with the sample transcript
    name=askopenfilenames()
    student_data=[]  # Will contain a list of tuples with selected student names and name of files
    student_info={}
    for i in range(len(name)):
        address=name[i]
        filename=""
        for i in range(len(address)-1,-1,-1):
            if address[i]=='/':
                break
            filename+=address[i]
        filename=filename[::-1].encode('ascii')
        student_name=filename.split('.')[0].encode('ascii')
        student_data.append((student_name, filename))

    for student_name,filename in student_data:
        student_info.setdefault(student_name,{})
        student_info[student_name]=grades_dictionary(filename)
    #Here we are able to extract the user/student name that we will merge with past student grades
    # Effectively we create one dictionary with all the data we need
    student = student_data[0][0]
    student_dict[student]=student_info[student]
    print student_dict

#This function converts letter grades to numeric grades
def convert_letter_to_dig(list_of_grades):
    result=[letter_dig_grade[letter] for letter in list_of_grades]
    return result


# Had to replicate this function because the one in the recommendations file does not take
# similarity argument into account.

#This function will enable us to use item based filtering by transforming the dictionary
# It has a flag that checks the current state of the dictionary and inverts it accordingly
def item_based_use():
    global student_dict,flag
    if student_dict is not None and flag is False:
        student_dict=transformPrefs(student_dict)
        flag=True
    print student_dict

#This function will enable us to use user based filtering by using the same mechanism as the one above
def user_based_use():
    global student_dict,flag
    if student_dict is not None:
        if flag is True:
            student_dict=transformPrefs(student_dict)
            flag=False
    print student_dict

#The score functions are useful for making the combo_box
# The sim_func_used tracks which score function was used, pearson, sdistance or jaccard
#Each score function returns a list of similarity comparisons of user with other past students
def pearson_score():
    global student, student_dict,sim_func_used,flag
    if flag is False:  # Then we are using user based system
        other_grades=[]
        for item in student_dict:
            #don't compare me to myself
            if item==student: continue
            other_grades.append(item)
        ps=[(sim_pearson(student_dict,student,x),x) for x in other_grades]
    elif flag is True:  # Then we are using item based system
        #This part is not too necessary for the program, just a part of the development plan
        orig_dict=transformPrefs(student_dict)
        sim_subjects=calculateSimilarItems(orig_dict,n=6)
        ps=sim_subjects
    sim_func_used=0
    print ps
    return ps


def sdistance_score():
    global student, student_dict,sim_func_used,flag
    if flag is False:
        other_grades=[]
        for item in student_dict:
            if item==student: continue
            other_grades.append(item)
        sd=[(sim_distance(student_dict,student,x),x) for x in other_grades]
    elif flag is True:
        orig_dict=transformPrefs(student_dict)
        sim_subjects=calculateSimilarItems(orig_dict,n=6)
        sd=sim_subjects
    sim_func_used=1
    print sd
    return sd


def jaccard_score():
    global student,student_dict,sim_func_used,flag
    if flag is False:
        other_grades=[]
        for item in student_dict:
            if item==student: continue
            other_grades.append(item)
        js=[(sim_jaccard(student_dict,student,x),x) for x in other_grades]
    elif flag is True:
        orig_dict=transformPrefs(student_dict)
        sim_subjects=calculateSimilarItems(orig_dict,n=6)
        js=sim_subjects
    sim_func_used=2
    print js
    return js

sug=[]
#This is the function that enables us to see all recommendations for each student
def seeRecommendations():
    global student_dict,student,sim_func_used,endGrades,flag
    if student is None or student_dict is None:
        showerror("Error 101", "An error has occurred!"
                               " Please load current transcript and past student grades")
    else:
        if sim_func_used==0:
            if flag is False:
                sug=getRecommendations(student_dict,student,sim_pearson)
            else:
                orig_dict=transformPrefs(student_dict)
                sim_subjects=calculateSimilarItems(orig_dict,n=6)
                sug=getRecommendedItems(orig_dict,sim_subjects,student)
        elif sim_func_used==1:
            if flag is False:
                sug=getRecommendations(student_dict,student,sim_distance)
            else:
                orig_dict=transformPrefs(student_dict)
                sim_subjects=calculateSimilarItems(orig_dict,n=6)
                sug=getRecommendedItems(orig_dict,sim_subjects,student)
        elif sim_func_used==2:
            if flag is True:
                orig_dict=transformPrefs(student_dict)
                sim_subjects=calculateSimilarItems(orig_dict,n=6)
                sug=getRecommendedItems(orig_dict,sim_subjects,student)
            else:
                sug=getRecommendations(student_dict,student,sim_jaccard)
    endGrades = digit2lettergradeMapping(sug[0:6])
    ResultTable()



def digit2lettergradeMapping(somelist):
    result=[]
    for rating,subject in somelist:
        if rating>4:
            grade='A+'
        elif 3.7<rating<=4:
            grade='A'
        elif 3.3<rating<=3.7:
            grade='A-'
        elif 3.0<rating<=3.3:
            grade='B+'
        elif 2.7<rating<=3.0:
            grade='B'
        elif 2.3<rating<=2.7:
            grade='B-'
        elif 2.0<rating<=2.3:
            grade='C+'
        elif 1.7<rating<=2.0:
            grade='C'
        elif 1.3<rating<=1.7:
            grade='C-'
        elif 1.0<rating<=1.3:
            grade='D+'
        elif 0.5<rating<=1.0:
            grade='D'
        elif 0.1<rating<=0.5:
            grade='D-'
        elif rating<=0.1:
            grade='F'
        full_grade=grade+'('+str(rating)[0:4]+')'
        result.append((subject,full_grade))
    return result


#Organizing the Frames
master = Tk()
master.title('VIRTUAL ADVISOR')
frame = Frame(master)
frame.pack()
lowerframe = Frame(master)
lowerframe.pack(anchor=W, pady=10)
lowestframe = Frame(master)
lowestframe.pack(anchor=W, pady=10)
treeframe=Frame(master)
treeframe.pack(pady=5)
#Organizing the labels,buttons and menus
w1 = Label(frame, text='______'*22)
w1.pack()
w2 = Label(frame, text= 'Virtual Advisor (Pro)', font='Helvetica 20 bold italic', width=35, fg='blue')
w2.pack()
s = Label(frame, text='______'*22)
s.pack()
#End of initial label


#Start of button organization
button1=Button(frame, text='Load Past Student Grades', width=30, height=2,
               font='Verdana 10 bold italic', fg='red', command=student_dictionaries)
button1.pack(side=LEFT)
button2=Button(frame, text='Load Your Current Transcript',width=30,height=2,
               font='Verdana 10 bold italic', fg='red', command=merge_dictionaries)
button2.pack(side=RIGHT)
button3=Button(lowestframe, text='See Recommended Courses',width=30,height=2,
               font='Verdana 10 bold italic', fg='dark blue',command=seeRecommendations)
button3.pack(side=LEFT)

#Label, radiobutton and combobox organization
Label(lowerframe,text='Collaborative\n Filtering Type:',font='Verdana 14 italic',
      justify=CENTER).pack(side=LEFT)
v = IntVar()  # Radio button variable
v.set(1)
Radiobutton(lowerframe, text='User based', variable=v, value=1, font='Times 10',command=user_based_use).pack(side=LEFT)
Radiobutton(lowerframe, text='Item based', variable=v, value=2, font='Times 10',command=item_based_use).pack(side=LEFT)
Label(lowerframe,text='Similarity\n Measure:', font='Verdana 14 italic', justify=CENTER).pack(side=LEFT, padx=20)


# #Creating the Menu(combobox)
class ComboApp:

    def __init__(self,parent):
        self.parent = parent
        self.dist_mes= 'Pearson'
        self.combo()

    def new_choice(self,event):
        self.dist_mes=self.box.get()
        if self.dist_mes=='Pearson':
            pearson_score()
        elif self.dist_mes=='Euclidean':
            sdistance_score()
        elif self.dist_mes=='Jaccard':
            jaccard_score()

    def combo(self):
        self.box_value=StringVar()
        self.box=ttk.Combobox(self.parent, textvariable=self.box_value,width=30)
        self.box['values']=('Pearson', 'Euclidean','Jaccard')
        self.box.current(0)
        self.box.bind("<<ComboboxSelected>>", self.new_choice)
        self.box.pack(side=RIGHT,padx=1)


dist_app=ComboApp(lowerframe)

i=1
class ResultTable(object):
    def __init__(self):
        self.tree=None
        self.setup_widget()
        self.build_tree()

    def setup_widget(self):
        container = treeframe
        container.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(columns=column_header, show="headings")
        self.tree.grid(column=0, row=0, sticky='nsew', in_=container)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)

    def build_tree(self):
        for col in column_header:
            self.tree.heading(col, text=col.title())
        for item in endGrades:
            self.tree.insert('', 'end', values=item)


column_header=['Recommended Course','Predicted Grade']

mainloop()

