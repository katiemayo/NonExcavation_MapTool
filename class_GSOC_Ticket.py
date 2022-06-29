
# ---------------------------------------------------------------------------
# Description: Gopher State One Call ticket scanning
# Author: Paul Hedlund (Houston Engineering Inc.)
# Created: April 2020
# ---------------------------------------------------------------------------
from bs4 import BeautifulSoup
import requests,re

class GSOC_Ticket:
	def __init__(self, url = None):
		self.url = url

	def get_ticket_info(self):
		ticket_info_dict = {}
		ticket = ""
		session = requests.Session()
		session.post(self.url)
		result = session.get(self.url)
		page = result.content
		soup = BeautifulSoup(page, "html.parser")
		tagticket = soup.find(id='ticketInfo')
		ticket_info_tags =  tagticket.find_all('div', class_ = 'row icon-steps')

		for entry in ticket_info_tags:
			for item in entry.findChildren():
				if item.text == 'Ticket number':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					ticket = str(tagssibling.find_next('span').text)
					ticket_info_dict[str(ticket)] = {}
				if item.text == 'Company name':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					company_name = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['company'] = str(company_name).title()
				if item.text == 'Type of work':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					work = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['worktype'] = str(work).title()
					break

		ticket_info_tags2 =  tagticket.find_all('div', class_ = 'row comments')
		for entry in ticket_info_tags2:
			for item in entry.findChildren():
				if item.text == 'Caller':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					caller = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['caller'] = str(caller).title()
				if item.text == 'Phone':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					caller_phone = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['callerphone'] = str(caller_phone)
				if item.text == 'Contact':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					contact_person = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['contact'] = str(contact_person).title()
					#print(str(contact_person))
				if item.text == 'Contact phone':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					contact_phone = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['phone'] = str(contact_phone)
				if item.text == 'Work being done for':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					work_for = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['workfor'] = str(work_for).title()
				if item.text == 'Work to begin date':
					tagsparent = item.find_parent('div')
					tagssibling = tagsparent.find_next_sibling('div')
					date = str(tagssibling.find_next('span').text)
					ticket_info_dict[ticket]['date'] = str(date).title()
		tagemail = soup.find(id='email')
		email = str(tagemail.find_next('span').text)
		ticket_info_dict[ticket]['email'] = str(email)
		#print(email)

		script_elementfull = soup.findAll("script")
		soupjavascript = str(script_elementfull[41])

		latlngpattern = re.search(r'polyListLength = (.*?).\length', soupjavascript).group(1)  #Excavator poly
		#latlngpattern = re.search(r'polyListLengthNo = (.*?).\length', soupjavascript).group(1)  #GSOC poly
		latlngpattern = latlngpattern.strip()
		latlngpattern = latlngpattern[2:-2]
		latlngpattern = latlngpattern.split('},{')

		lats = []
		lngs = []
		for i in latlngpattern:
			lats.append(float(i[27:37]))
			lngs.append(float(i[46:57]))
			#lats.append(float(i[18:28]))
			#lngs.append(float(i[37:48]))

		coordinates = list(zip(lngs, lats))
		coordinates = [tuple(map(float, coords)) for coords in coordinates]
		ticket_info_dict[ticket]['polygon coordinates'] = coordinates
		return (ticket_info_dict)
