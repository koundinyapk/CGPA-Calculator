# CGPA-Calculator

This is a module developed using Django framework which calculates the CGPAs of a given set of students according to 8 point grading system. 
Grades range from Ex to P & F. Points corresponding to each of the grades range from 10 to 5 & 0. All the inputs given are in the form of Excel files. 
The output of is also obtained as an Excel file. This module consists of the following apps:

->subjectWise

->sgpa

->finalCgpa

The decriptions of the above said apps are as follows:

**subjectWise:**

This takes an Excel file that contains the scores of students(of all the subjects) in all the sub exams conducted during that semester as input. 
The summation of the marks of the students obtained in the all the sub eaxms in that semester are calculated and grades are alloted to the students and points corresponding to the grades are calculated accordingly.
The output of this application is an Excel file in which Grades and Points are alloted to the students for all the subjects that were pursued during that semester.

**sgpa:**

This app takes the Excel file that was obtained as output of the subjectWise app as input. It also takes an Excel header file that contains credits that each subject bears during that semester. This app calculates the SGPA of each student from the grades and points alloted. It ouputs an Excel file that contains SGPAs of all the studnts.

**finalCgpa:**

This app takes the ouput of sgpa app as input. It also takes an Excel header file that contains the SGPAs of students obtained in the previous semesters. It also contains the CGPAs of the students obtained till that semester. It adds the SGPAs obtained by the students int the current semester to that file and updates the CGPAs by including the SGPAs of current semester. It outputs an Excel file which contains the updated CGPAs of the students.


**Note:** 
The layouts of the Excel files that are accepted by this module are included in this repository with a folder name **layouts_of_excelSheets.**




