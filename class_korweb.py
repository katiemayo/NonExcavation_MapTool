# Complete GSOC ticket in Korterra, include a link to the ticket map.
#define information and behavior that characterize anything you want to model in program
from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time,arcpy

class Korweb:
	def __init__(self, username, password, customerid, web_driver_path, ticket, output_path, laptop, sender):
		self.username = username
		self.password = password
		self.customerid = customerid
		self.web_driver_path = web_driver_path
		self.ticket = ticket
		self.output_path = output_path
		self.laptop = laptop
		self.sender = sender
                # Open and log into Korterra and get ticket page on chrome.
		TimeSleep = 12
		self.driver = webdriver.Chrome(executable_path = self.web_driver_path)
		self.driver.get(r"https://mn.korweb.com/Ticket/Detail/" + ticket + "-S")
		time.sleep(TimeSleep)
		self.driver.find_element_by_id("CustomerId").send_keys("-------")
		self.driver.find_element_by_id("customer-redirect-submit-btn").click()
		time.sleep(TimeSleep)
		self.driver.find_element_by_id("username").send_keys("-------")
		self.driver.find_element_by_id("password").send_keys("-------")
		time.sleep(TimeSleep)
		# On ticket page, click 'complete' buton.
		self.driver.find_element_by_xpath("//*[@id='page-ui-container']/div/div/div/div[2]/div[1]/div/form/div[3]/button").click()
		time.sleep(TimeSleep)
		self.driver.find_element_by_id('btnDetailComplete').click();
		time.sleep(TimeSleep)
		# in 'positive response drop down, select non excavation
		mySelect = Select(self.driver.find_element_by_id("custom2-CMINNE05CUSTOM2dashS-positiveresponse"))
		mySelect.select_by_value("5-NON EXCAVATION")
		time.sleep(TimeSleep)
		# click the marked button
		self.driver.find_element_by_xpath("//input[@name='CMINNE05CUSTOM2dashS-radios_01-1'][@value='Marked']").click();
		time.sleep(TimeSleep)
		# locate text box and fill with map link, comments.
		self.driver.find_element_by_id('header-remarks').send_keys("Sent pdf maps of AOI. \n Find map here: \n" + self.laptop + "\nAuto-Completed Through PWSWS Data Delivery Tool\n{}".format(time.strftime('%b-%d-%Y_%I-%M-%p')) + "\n" + self.sender)
		try:
			# click complete button
			self.driver.find_element_by_id('btnComplete').click();
		except:
			arcpy.AddError("Completion message already sent for ticket {]".format(ticket))
		time.sleep(TimeSleep)
		self.driver.quit()
