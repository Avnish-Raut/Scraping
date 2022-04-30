import requests
import csv
import xlsxwriter
from bs4 import BeautifulSoup

URL = "https://www.coursera.org/certificates/launch-your-career#professional-certificates"
page = requests.get(URL)

soup = BeautifulSoup(page.content,"html.parser")

course_list = []
results = soup.find(id="rendered-content")

job_elements = results.find("div", class_="rc-ExpandedCertsList")
class_elements = results.find_all("li",class_="ProductOfferingCard css-wo777h")
#print(class_elements)
for job_element in class_elements:
    #print(job_element)
    course_name = job_element.find("p",class_="cds-7 css-19ir7w5 cds-9").text.strip()
    offered_by = job_element.find("p",class_="cds-7 css-19lldze cds-9").text.strip()
    time_slot = job_element.find("p",class_="cds-7 css-1g9mlwz cds-9").text.strip()
    #description = job_element.find("p",class_="cds-7 css-1czhh8a cds-9")
    #des = job_element.find("p",class_="cds-7 css-1r9ruxl cds-9").text.strip()
    links = job_element.find("a",class_="css-jjyq8a")
    link = links['href'].strip()
    linkC = ("https://www.coursera.org"+link)
    

    course_list.append([
        course_name,
        offered_by,
        time_slot,
        linkC
        ])

#keys = course_list[0].keys()

#with open('course.csv','w',newline='') as output_file:
 #   dict_writer = csv.DictWriter(output_file,keys)
  #  dict_writer.writeheader()
   # dict_writer.writerows(course_list)

workbook = xlsxwriter.Workbook('Course_list.xlsx')
worksheet = workbook.add_worksheet()
cell_format = workbook.add_format()
cell_format.set_text_wrap()
row =0
col=0
bold = workbook.add_format({'bold': True})
worksheet.set_row(0, 40)
worksheet.set_column('B:A', 20)
worksheet.set_column('C:A', 20)
worksheet.set_column('D:A', 50)
#worksheet.set_column('A:A', 30)
worksheet.write(row,col,"Course_name",bold)
worksheet.write(row,col+1,"Offered_by",bold)
worksheet.write(row,col+2,"Time_requirement",bold)
worksheet.write(row,col+3,"Link",bold)
row +=1
for course_name,offered_by,time_slot,linkC in (course_list):
    worksheet.set_row(row, row+40)
    
    worksheet.write(row,col,course_name,cell_format)
    worksheet.write(row,col+1,offered_by,cell_format)
    worksheet.write(row,col+2,time_slot,cell_format)
    worksheet.write_url(row,col+3,linkC,cell_format)
    row +=1
workbook.close()
