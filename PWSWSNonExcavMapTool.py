# -*- coding: cp1252 -*-
## Script for creating pdf maps for locators via GSOC ticket URLs

# import required packages
from selenium import webdriver
import time,arcpy,os,traceback
from bs4 import BeautifulSoup

arcpy.env.overwriteOutput = True

#Parameters ----------------------------------------------------------------------------
Sharepoint_connection = arcpy.GetParameterAsText(0)
Sender = arcpy.GetParameterAsText(1)
Shapefiles = arcpy.GetParameter(2)
if not Sharepoint_connection:
	Sharepoint_connection = r'D:\Development\MplsPythonScript\20211201\CodeNew\LocatorMapTools\FakeSharepoint'
	Sender = 'Katie Mayo'
	Shapefiles = True

ChromeDriverPath = r"M:\PWSWS\XSHARE\Katie\CurrentTools\chromedriver"
path = ChromeDriverPath
driver = webdriver.Chrome(executable_path = path)
#ChromeDriverPath = r"D:\Development\MplsPythonScript\20211201\CodeNew\NonExcavationMapTool\chromedriver.exe"
timesleep = 7
logfolder = os.path.join(Sharepoint_connection, time.strftime('%Y') + "NonexcavationLogFiles", time.strftime('%B'))
logtext = os.path.join(logfolder, time.strftime('%d-%B')+".txt")
if not os.path.exists(logfolder):
	os.makedirs(logfolder)
if os.path.exists(logtext):
	log = open(logtext, "a")
else:
	log = open(logtext, "w")
log.write("\n\nCurrent Time: {}".format(datetime.datetime.now().strftime("%H:%M")))

# create output folder.
arcpy.AddMessage("\nHello, looks like you're about to complete non-excavation tickets")
year = time.strftime('%Y')
output_path = os.path.join(Sharepoint_connection, year)
if os.path.exists(output_path):
	pass
else:
	os.makedirs(output_path)
arcpy.AddMessage("\n Maps will be saved here: {}".format(output_path))
log.write("\n Maps will be saved here: {}".format(output_path))
# search korterra for assigned non-excav tickets
driver.get(r"https://mn.korweb.com/Tickets")
time.sleep(timesleep)
driver.find_element_by_id("CustomerId").send_keys("-------")
driver.find_element_by_id("customer-redirect-submit-btn").click()
time.sleep(timesleep)
driver.find_element_by_id("username").send_keys("-------")
driver.find_element_by_id("password").send_keys("-------")
time.sleep(timesleep)
driver.find_element_by_xpath("//*[@id='page-ui-container']/div/div/div/div[2]/div[1]/div/form/div[3]/button").click()
time.sleep(timesleep)
driver.find_element_by_id("ddMobile").send_keys("SEWER_NONEXCAV - Sewer Non-Excavation")
driver.find_element_by_id("ddStatus").send_keys("NEW+ASSIGNED")
driver.find_element_by_name("js-ticketsearch-gridkorterra_datatable_length").send_keys("500")
time.sleep(timesleep)
driver.find_element_by_id("btnSearch").click()
time.sleep(timesleep)
# translate selenium link to beautifulsoup
html = driver.page_source
soup = BeautifulSoup(html)

# find table and extract area of href link. Outputs a list of the links to all assigned tickets.
table_data = soup.find('table')
links_list = table_data.find_all('a', class_='ticketLink')
ticket_nd_dict = []
gsoc_url = []
time.sleep(timesleep)
for link in links_list:
	ticket_nd_dict.append(link.attrs['href'])
for t in ticket_nd_dict:
	driver.get(t)
	time.sleep(3)
	tic = driver.find_elements_by_xpath("//*[@id='oneCallFormat']/a")
	gsoc_url.append(str(tic[-1].get_attribute('href')))
driver.quit()

arcpy.AddMessage("\nGrabbing GSOC polygon info, this will take some time .  .  .  .\n")
log.write("\n There are {}".format(len(gsoc_url))+" open non-excavation tickets.")
# Collect ticket information for one ticket at a time.
from class_GSOC_Ticket import GSOC_Ticket
gsoc_tickets_dict = {}
for gsoc in gsoc_url:
	gsoc_ticket = GSOC_Ticket(gsoc)
	try:
		gsoc_ticket_info_dict = gsoc_ticket.get_ticket_info()
	except KeyError:
		log.write("\nUnable to get GSOC ticket info.")
		arcpy.AddWarning("Unable to get GSOC ticket info.")
	gsoc_tickets_dict.update(gsoc_ticket_info_dict)
arcpy.AddMessage("Done. Making maps now")

from class_PWSWS_Data_Delivery import PWSWS_Data_Delivery

# loop through ticket numbers. complete one at a time (including korweb module)
for key in gsoc_tickets_dict.keys():
	try:
		ticket = key
		company_name = gsoc_tickets_dict[ticket]['company']
		output_path_ticket = os.path.join(output_path, company_name, ticket)
		arcpy.AddMessage("\nGSOC Ticket {}".format(ticket))
		log.write("\nGSOC Ticket {}".format(ticket))
		if os.path.exists(output_path_ticket):
			pass
			log.write("\nDirectory for ticket {}".format(ticket)+ " already exists.")
			arcpy.AddWarning("Directory for ticket {}".format(ticket) + " already exists.")
		else:
			os.makedirs(output_path_ticket)
			gsoc_data_delivery = PWSWS_Data_Delivery(output_path_ticket, input_selected_layers = [None])
			gsoc_data_delivery.create_aoi_as_feature_layer(polygon_coordinates = gsoc_tickets_dict[ticket]['polygon coordinates'],gsoc_ticket = ticket, caller = gsoc_tickets_dict[ticket]['caller'],callerphone = gsoc_tickets_dict[ticket]['callerphone'],company = gsoc_tickets_dict[ticket]['company'], worktype = gsoc_tickets_dict[ticket]['worktype'], contact = gsoc_tickets_dict[ticket]['contact'], phone = gsoc_tickets_dict[ticket]['phone'], work_for = gsoc_tickets_dict[ticket]['workfor'], email = gsoc_tickets_dict[ticket]['email'], sender = Sender, shapefiles = Shapefiles)
			log.write("\nMaking maps for {}".format(ticket))

		arcpy.AddMessage("Selecting grid cells within ticket {}".format(ticket))

		from class_korweb import Korweb
		arcpy.AddMessage("Uploading locator maps to korterra ticket page.\n")
		co = gsoc_tickets_dict[ticket]['company']
		coencode = co.replace(" ", "%20")
		url = "https://minneapolismngov.sharepoint.com/teams/p00044/Shared%20Documents/SWSNonExcavation/" + year + "/" + coencode + "/" + ticket + "/" + ticket + "_map.pdf"
		korweb = Korweb(username = "SWSENGINEER", password = "SWSLocate123!",customerid = "MINNEAPOLIS", web_driver_path = ChromeDriverPath,ticket = ticket, output_path = (output_path_ticket + '\\' + ticket + '_map.pdf'), laptop = url, sender = Sender)

	except Exception as e:
		arcpy.AddWarning("Error with current ticket. Ticket has been skipped".format(e))
		arcpy.AddWarning(traceback.format_exc())
		log.write("\nError with current ticket. Ticket has been skipped".format(e))
		log.write("\n" + traceback.format_exc())
		continue

arcpy.AddMessage("Non-Excavation Map Tool Complete!")
