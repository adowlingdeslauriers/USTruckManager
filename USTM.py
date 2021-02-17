'''
If you're reading this, I'm sorry for the mess. This is the 3rd major release of a widget that keeps growing and growing in scope.

TODO:
-Future-proof anything that uses a CSV and relies on columns to be in a specific place
-Add client/carrier/gaylord support to JSON-CSV and CSV-JSON conversion
-Better error messages
-Better error-handling
-Clean up this spaghetti
'''

from appJar import gui
import json
from datetime import date
import time
import os
import csv
import sys
from openpyxl import Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.utils import Image
import traceback
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate
from email import encoders
import shutil
import re
import string

### Functions

def createProForma():
	print("\n>Attempting to Create ProForma")
	global detailed_report_json
	try:
		if detailed_report_json == {}: raise Exception
		else: #If loading didn't error out
			print("Building commodities list")
			commodities_list = {}
			for entry in detailed_report_json:
				for commodity in entry["commodities"]:
					name = commodity["description"]
					if name not in commodities_list.keys():
						commodities_list[name] = 0 #Add commodity to the list with 0 quantity
					commodities_list[name] = commodities_list[name] + int(commodity["quantity"]) #Update the quantity

			#Techship passes bad data. Error Correction below
			commodities_list = cleanCommoditiesList(commodities_list)

			print("Loading ProForma information")
			master_data = []
			with open(os.getcwd() + os.sep + "resources" + os.sep + "MASTER_FDA_LIST.csv", "r") as master_file:
				csv_reader = csv.reader(master_file, delimiter = ",")
				for line in csv_reader:
					if line != "" and line[0] != "#":
						master_data.append(line)

			print("Matching commodities to Master File")
			proforma_data = []
			for commodity in commodities_list:
				for line in master_data:
					if commodity.upper() == line[2].upper() and commodities_list[commodity] != 0: #If commodity description matches description from Master FDA file and there is >0 items
						_quantity = commodities_list[commodity]
						#Ugly CSV line building
						_out = (line[1], _quantity, "PCS", float(line[5]), "", "", "", line[9], line[10], "", line[12], float(_quantity * float(line[5])), "", line[15], line[16], line[17], line[18], line[19], line[20], line[21], line[22], "", line[24], line[25], line[26], line[27], line[28], line[29], line[30], line[31], "", line[33], line[34], line[35], line[36], line[37], line[38], line[39], line[40], "", line[41], float(_quantity * float(line[5])), 1, "KG")
						proforma_data.append(_out)

			print("Saving ProForma upload template")
			header = (("ShipperRefNum","PostToBroker","InvoiceDate","StateDest","PortEntry","MasterBillofLading","Carrier","EstDateTimeArrival","TermsofSale","RelatedParties","ModeTrans","ExportReason","FreightToBorder","ContactInformation","IncludesDuty","IncludesFreight","FreightAmount","IncludesBrokerage","Currency","TotalGrossWeightKG","ContainerNumber","ShippingQuantity","ShippingUOM","DutyandFeesBilledTo","InvoiceNumber","OwnerOfGoods","PurchaseOrder","ShipperCustNo","ShipperName","ShipperTaxID","ShipperAddr1","ShipperAddr2","ShipperCity","ShipperState","ShipperCountry","ShipperPostalCode","ShipperMfgID","ShipToCustNo","ShipToName","ShipToTaxID","ShipToAddr1","ShipToAddr2","ShipToCity","ShipToState","ShipToCountry","ShipToPostalCode","SellerCustNo","SellerName","SellerTaxID","SellerAddr1","SellerAddr2","SellerCity","SellerState","SellerCountry","SellerPostalCode","MfgCustNo","MfgName","MfgID","MfgAddress1","MfgAddress2","MfgCity","MfgState","MfgCountry","MfgPostalCode","BuyerCustNo","BuyerName","BuyerUSTaxID","BuyerAddress1","BuyerAddress2","BuyerCity","BuyerState","BuyerCountry","BuyerPostalCode","ConsigneeCustNo","ConsigneeName","ConsigneeUSTaxID","ConsigneeAddress1","ConsigneeAddress2","ConsigneeCity","ConsigneeState","ConsigneeCountry","ConsigneePostalCode","PartNumber","Quantity","QuantityUOM","UnitPrice","GrossWeightKG","NumberOfPackages","PackageUOM","CountryOrigin","SPI","ProductClaimCode","Description","ValueOfGoods","LineMfgCustNo","LineMfgName","LineMfgID","LineMfgAddress1","LineMfgAddress2","LineMfgCity","LineMfgState","LineMfgCountry","LineMfgPostalCode","LineBuyerCustNo","LineBuyerName","LineBuyerUSTaxID","LineBuyerAddress1","LineBuyerAddress2","LineBuyerCity","LineBuyerState","LineBuyerCountry","LineBuyerPostalCode","LineConsigneeCustNo","LineConsigneeName","LineConsigneeUSTaxID","LineConsigneeAddress1","LineConsigneeAddress2","LineConsigneeCity","LineConsigneeState","LineConsigneeCountry","LineConsigneePostalCode","LineNote","Tariff1Number","Tariff1ProductValue","Tariff1Quantity1","Tariff1Quantity1UOM","Tariff1Quantity2","Tariff1Quantity2UOM","Tariff1Quantity3","Tariff1Quantity3UOM","Tariff2Number","Tariff2ProductValue","Tariff2Quantity1","Tariff2Quantity1UOM","Tariff2Quantity2","Tariff2Quantity2UOM","Tariff2Quantity3","Tariff2Quantity3UOM","Tariff3Number","Tariff3ProductValue","Tariff3Quantity1","Tariff3Quantity1UOM","Tariff3Quantity2","Tariff3Quantity2UOM","Tariff3Quantity3","Tariff3Quantity3UOM","Tariff4Number","Tariff4ProductValue","Tariff4Quantity1","Tariff4Quantity1UOM","Tariff4Quantity2","Tariff4Quantity2UOM","Tariff4Quantity3","Tariff4Quantity3UOM","Tariff5Number","Tariff5ProductValue","Tariff5Quantity1","Tariff5Quantity1UOM","Tariff5Quantity2","Tariff5Quantity2UOM","Tariff5Quantity3","Tariff5Quantity3UOM","Tariff6Number","Tariff6ProductValue","Tariff6Quantity1","Tariff6Quantity1UOM","Tariff6Quantity2","Tariff6Quantity2UOM","Tariff6Quantity3","Tariff6Quantity3UOM"))
			first_line = (app.getEntry("BoL #:"),"FALSE",app.getEntry("Date:"),"NY","0901",app.getEntry("PAPS #:"),app.getEntry("SCAC:"),app.getEntry("Date:") + " 03:00 PM","PLANT","","30","","","","","","","","",int(app.getEntry("Total Weight:")),"",app.getEntry("Package Count:"),"PCS","Buyer","","","","","STALCO INC","160901-55044","401 CLAYSON RD","","NORTH YORK","ON","CA","M9M 2H4","XOSTAINC401NOR","","","","","","","","","","","STALCO INC","160901-55044","401 CLAYSON RD","","NORTH YORK","ON","CA","M9M 2H4","","STALCO INC","XOSTAINC401NOR","401 CLAYSON RD","","NORTH YORK","ON","CA","M9M 2H4","","IMS OF WESTERN NY","16-131314301","2540 WALDEN AVE","SUITE 450","BUFFALO","NY","US","14225","","IMS OF WESTERN NY","16-131314301","2540 WALDEN AVE","SUITE 450","BUFFALO","NY","US","14225")
			
			workbook = Workbook()
			outputFolder()
			filename = os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-ProForma_Template.xlsx"
			worksheet = workbook.active
			worksheet.title = "Sheet1"
			worksheet.append(header)
			for row in proforma_data:
				worksheet.append(first_line + row)
			workbook.save(filename)
			print(">>Proforma template complete")

			with open(os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-USGR_Data.csv", "w", newline = "") as USGR_file:
				csv_writer = csv.writer(USGR_file)
				for row in proforma_data:
					csv_writer.writerow(row)

	except:
		errorBox("Error creating ProForma. Did you create the Master JSON yet?")

def loadVariables():
	#Loads /resources/DATA.json
	#DATA.json can be used to keep track of what the latest BoL/PAPS number is
	print("\n>Loading BoL/PAPS")
	with open(os.getcwd() + os.sep + "resources" + os.sep + "DATA.json", "r") as variables_file:
		global data
		data = json.load(variables_file)
		app.setEntry("Date:", data["date"])
		app.setEntry("BoL #:", data["BoL"])
		app.setEntry("PAPS #:", data["PAPS"])
	print(">>Loading complete")

def updateVariables():
	print("\n>Updating date/BoL/PAPS to today's values")
	app.setEntry("Date:", str(date.today()))
	app.setEntry("BoL #:", str(int(app.getEntry("BoL #:")) + 1).zfill(7))
	app.setEntry("PAPS #:", str(int(app.getEntry("PAPS #:")) + 1).zfill(6))
	outputFolder()
	print(">>Updating complete")

def saveVariables():
	print("\n>Saving current date/BoL/PAPS")
	global data
	data["date"] = app.getEntry("Date:")
	data["BoL"] = app.getEntry("BoL #:")
	data["PAPS"] = app.getEntry("PAPS #:")
	with open(os.getcwd() + os.sep + "resources" + os.sep + "DATA.json", "w") as variables_file:
		json.dump(data, variables_file, indent = 4)
	outputFolder()
	print(">>Saving complete")

def cleanCommoditiesList(commodities_list):
	#To add products, take the Description from the ACE Manifest
	#No need to set quantities to 0 as the items aren't in the Master FDA file and can't be added to the ProForma
	try:
		commodities_list["Eye Renew 0.5 fl oz Skin Care"] += commodities_list["BDRx Kit"]
	except:
		pass
	try:
		commodities_list['Flawless Face 2 fl oz'] += commodities_list["BDRx Kit"]
	except:
		pass
	try:
		commodities_list['Instalift 0.5 fl oz'] += commodities_list["BDRx Kit"]
	except:
		pass
	try:
		commodities_list['Tevida 60 caps - US'] += commodities_list['Tevida ']
	except:
		pass
	try:
		commodities_list['Vascular X 60 caps - US'] += commodities_list['Vascular X']
	except:
		pass
	#print("########################")
	#print(commodities_list)
	return commodities_list
	
def createBoL():
	print("\n>Creating BoL")
	updateGaylordCounts()
	
	file_name = os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-Stalco-BoL.pdf"
	image_file_path = os.getcwd() + os.sep + "resources" + os.sep + "STALCO_BOL.jpg"
	c = canvas.Canvas(file_name, pagesize = (1668, 1986), bottomup = 0)
	image = ImageReader(image_file_path)
	c.drawImage(image, 0, 0, mask = "auto")

	c.setFont("Courier", 24)
	today = date.today()
	c.drawString(4, 340, app.getEntry("Date:"))
	c.drawString(1555, 330, app.getEntry("BoL #:"))
	c.drawString(1562, 355, app.getEntry("PAPS #:"))
	c.drawString(134, 1210, app.getEntry("Total Gaylord Count:"))
	c.drawString(176, 1310, app.getEntry("USPS Gaylord Count:"))
	c.drawString(171, 1360, app.getEntry("DHL Gaylord Count:"))
	c.drawString(181, 1410, app.getEntry("FedEx Gaylord Count:"))
	c.drawString(525, 1175, app.getEntry("Package Count:"))
	c.drawString(854, 1175, app.getEntry("Total Weight:"))

	c.showPage()
	c.save()

	print(">>BoL generation complete")

def createACE():
	error_file = open(os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-Error_File.txt", "w")
	try:
		print("\n>Creating ACE Manifest")
		# Check if all the files are present and correct filetypes
		if app.getEntry("batchesFileEntry") != "" and \
		   app.getEntry("batchesFileEntry")[-4:] == ".csv" and \
		   app.getEntry("ACEManifestFileEntry") != "" and \
		   app.getEntry("ACEManifestFileEntry")[-5:] == ".json" and \
		   app.getEntry("CSVReportFileEntry") != "" and \
		   app.getEntry("CSVReportFileEntry")[-4:] == ".csv":

				global master_json_data
			
				print("Loading batches scans")
				batches_data = []
				with open(app.getEntry("batchesFileEntry"), "r") as batches_file:
					csv_reader = csv.reader(batches_file, delimiter = ",")
					for line in csv_reader:
						if line[0] != "" and line[1] != "" and "#" not in line[0] and "#" not in line[1]: #If the line isn't empty or a comment
							batches_data.append([line[0], line[1]])

				#NOTE: Currently if a client has FDA-regulated goods, none of their products can go through Section 321 (aka end up on the ACE)
				#This may possibly change in the future
				print("Loading FDA clients")
				FDA_clients = []
				with open(os.getcwd() + os.sep + "resources" + os.sep + "FDA_CLIENTS.json", "r") as clients_file:
					json_data = json.load(clients_file)
					for client in json_data["clients"]:
						FDA_clients.append(client["name"])

				#compare lists and build outputs
				good_json = []
				good_batches = []
				global detailed_report_json
				detailed_report_json = []
				for json_entry in master_json_data:
					for row in batches_data:
						if json_entry["BATCHID"] == row[0] or json_entry["ORDERID"] == row[0]: #If there's a match
							if json_entry["client"] in FDA_clients: #If the product is commercially cleared
								json_entry["GAYLORD"] = row[1] #Append the Gaylord assignment to the entry
								#good_json.append(json_entry) #Doesn't go on the JSON, only Detailed Report
								detailed_report_json.append(json_entry)
								good_batches.append(row[0])     
							else: #If it's not commercially cleared
								json_entry["GAYLORD"] = row[1]
								good_json.append(json_entry)
								detailed_report_json.append(json_entry)
								good_batches.append(row[0])
								
				good_json = validate_json(good_json) #Validation time

				print(">>ACE Manifest Complete")
				with open(os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-ACE.json", "w") as good_json_file:
					json.dump(good_json, good_json_file, indent = 4)

				#Unmatched JSON entries
				print("\n>Printing Unmatched JSON Entries")
				print("If an order below is from today, it is likely a missing pick ticket that must be found and scanned")
				error_file.write("\n>Printing Unmatched JSON Entries")
				error_file.write("\nIf an order below is from today, it is likely a missing pick ticket that must be found and scanned")

				unique_batches_list = []
				for json_entry in master_json_data:
					if json_entry not in detailed_report_json and json_entry["BATCHID"] not in unique_batches_list:
						unique_batches_list.append(json_entry["BATCHID"])
						print("Unmatched JSON Entry: {} {} {}".format(json_entry['BATCHID'].ljust(8, " "), json_entry['closeDate'].ljust(22, " "), json_entry["client"]))
						error_file.write("\nUnmatched JSON Entry: {} {} {}".format(json_entry['BATCHID'].ljust(8, " "), json_entry['closeDate'].ljust(22, " "), json_entry["client"]))

				#Unmatched Batch Scans
				print("\n>Printing Unmatched Batches")
				print("Use the Batches page in Techship to find out what manfiest these belong to, and include them in the ACE")
				error_file.write("\n\n>Printing Unmatched Batches")
				error_file.write("\nUse the Batches page in Techship to find out what manfiest these belong to, and include them in the ACE")
				
				for row in batches_data:
					if row[0] not in good_batches:
						print("Unmatched batch: {} {}".format(row[0].ljust(8, " "), row[1]))
						error_file.write("\nUnmatched batch: {} {}".format(row[0].ljust(8, " "), row[1]))

				#Package Count
				package_count = len(detailed_report_json)
				app.setEntry("Package Count:", str(package_count))
				app.setEntry("Total Weight:", str(int(package_count // 2.20462)))

				#Gaylords Assignment
				unique_gaylords_list = []
				global gaylords_assignments
				gaylords_assignments = []
				for entry in detailed_report_json:
					#Clean it first
					if len(entry["GAYLORD"]) == 2:
						entry["GAYLORD"] = "G0" + entry["GAYLORD"][-1]
					#Add it to the uniques list
					match = False
					for unique_gaylord in unique_gaylords_list:
						if entry["GAYLORD"] == unique_gaylord["id"]:
							match = True
					if not match:
						unique_gaylords_list.append({"id": entry["GAYLORD"], "hasFDA": False, "hasUSPS": False, "hasDHL": False, "hasFedex": False, "packages": 0})
				

				for entry in detailed_report_json:
					for line in unique_gaylords_list:
						if entry["GAYLORD"] == line["id"]: #Will match only one gaylord
							#Now set flags
							line["packages"] += 1
							if entry["carrier"] == "DHLGLOBALMAIL" or entry["carrier"] == "DHLGLOBALMAILV4":
								line["hasDHL"] = True
							elif entry["carrier"] == "FEDEX":
								line["hasFedex"] = True
							elif entry["carrier"] == "EHUB":
								line["hasUSPS"]= True
							else:
								errorBox("Package with Order ID {} found without carrier".format(entry["ORDERID"]))
							if entry["client"] in FDA_clients:
								line["hasFDA"] = True

				for line in unique_gaylords_list:
					out = []
					out.append(line["id"])

					carriers_count = 0
					if line["hasUSPS"]: carriers_count += 1
					if line["hasDHL"]: carriers_count += 1
					if line["hasFedex"]: carriers_count += 1

					if carriers_count != 1:
						out.append("ERROR")
						errorBox("Multiple carriers found for {}".format(line["id"]))
					elif line["hasUSPS"]: out.append("EHUB")
					elif line["hasDHL"]: out.append("DHLGLOBALMAIL")
					elif line["hasFedex"]: out.append("FEDEX")

					if line["hasFDA"]: out.append("FDA")
					else: out.append("")

					out.append(line["packages"])

					gaylords_assignments.append(out)

				gaylords_assignments.sort(key = lambda x: x[0])
					
				#Detailed Report
				file_name = os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-Detailed_Report.csv"
				with open(file_name, "w", newline = "") as report_file:
					csv_writer = csv.writer(report_file)
					csv_writer.writerow(["Gaylord", "Name", "Address", "City", "State", "Country", "ZIP", "BATCHID", "ORDERID", "SCAC", "Service", "Client", "Close Date", "Commodity 1", "Commodity 2", "Commodity 3", "Commodity 4", "Commodity 5", "Commodity 6", "Commodity 7", "Commodity 8", "Commodity 9", "Commodity 10"])
					for entry in detailed_report_json:
						try:
							_stateProvince = entry["consignee"]["address"]["stateProvince"]
						except:
							_stateProvince = ""
						out_line = [entry["GAYLORD"], entry["consignee"]["name"], entry["consignee"]["address"]["addressLine"], entry["consignee"]["address"]["city"], _stateProvince, entry["consignee"]["address"]["country"], entry["consignee"]["address"]["postalCode"], entry["BATCHID"], entry["ORDERID"], entry["shipmentControlNumber"], entry["carrier"], entry["client"], entry["closeDate"]]
						for commodity in entry["commodities"]:
							out_line.append(commodity["description"])
						csv_writer.writerow(out_line)

				print(">>Orders validated and Detailed Report created")
				print(f">>Outputting {package_count} entries to Detailed Report")
				error_file.write(f"\n\n>>Outputting {package_count} entries to Detailed Report\n")
		else:
			raise FileNotFoundError
	except FileNotFoundError:
		if app.getEntry("batchesFileEntry") == "": errorBox("Batches Scans file not found!\nA copy should be located in W:\\Logistics\\Tools\\USTruckManager\\")
		elif app.getEntry("batchesFileEntry")[-4:] != ".csv": errorBox("Batches Scans file not a CSV file")
		if app.getEntry("ACEManifestFileEntry") == "": errorBox("ACE Manifest not found!")
		elif app.getEntry("ACEManifestFileEntry")[-5:] != ".json": errorBox("ACE Manifest not a JSON file")
		if app.getEntry("CSVReportFileEntry") == "": errorBox("CSV Report not found!")
		elif app.getEntry("CSVReportFileEntry")[-4:] != ".csv": errorBox("CSV Report not a CSV file ! Did you download the XLSX (excel) version by accident?")
	except Exception:
		errorBox("Some other error occured")
	finally:
		try: error_file.close()
		except: pass

def outputFolder():
	folder_path = os.getcwd() + os.sep + app.getEntry("Date:")
	if not os.path.exists(folder_path):
		os.mkdir(folder_path)
		print("Creating folder for", app.getEntry("Date:"))

def createMasterJSON():
	#Opens the ACE and CSV report. Matches entries in the two based on order IDs. Takes the client/carrier/close date data from the CSV and adds it to the ACE
	global master_json_data
	print(">Creating Master JSON")
	

	report_list = []
	print("Loading CSV Report")
	try:
		i = -1
		with open(app.getEntry("CSVReportFileEntry"), "r", encoding='utf8') as report_file:
			csv_reader = csv.reader(report_file, delimiter = ",")
			for i, line in enumerate(csv_reader, start = 1): #Unused enumerate for error-handling
				if line[0] != "#": #For testing
					report_list.append(line)
	except FileNotFoundError: errorBox("CSV Report not found!")
	except: errorBox(f"Failed to load CSV Report on line {i}")
				
	print("Loading ACE")
	ACE_data = []
	master_json_data = []
	try:
		with open(app.getEntry("ACEManifestFileEntry"), "r") as ACE_file:
			ACE_data = json.load(ACE_file)

		print("Merging data (Please wait patiently!)")
		for csv_entry in report_list:
			for json_entry in ACE_data:
				if csv_entry[16] == json_entry["ORDERID"]:
					json_entry["client"] = csv_entry[1]
					json_entry["carrier"] = csv_entry[2]
					json_entry["closeDate"] = csv_entry[22] #UTC
					#json_entry["trackingNumber"] = csv_entry[15] # Stored in scientific notation. WHYYYYYYY
					master_json_data.append(json_entry)
	except FileNotFoundError: errorBox("ACE Manifest not found!")
	except: errorBox("Error loading ACE Manifest")
	print(">>Master JSON Generated!")

def clean(in_string):
	good_chars = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ "
	return "".join(c for c in in_string if c in good_chars)

def validate_json(in_json):
	out_json = []

	SCAC_code = app.getEntry("SCAC:")

	# Get rid of duplicates
	for entry in in_json:

		if entry["client"] == "LUS Brands" and "," in entry["GAYLORD"]:
			#print(entry["ORDERID"], entry["GAYLORD"], entry["carrier"])
			if entry["carrier"] == "EHUB":
				entry["GAYLORD"] = entry["GAYLORD"].split(",")[0]
			else:
				entry["GAYLORD"] = "G" + entry["GAYLORD"].split(",")[1].replace("G", "")
			#print(entry["GAYLORD"])
		elif "," in entry["GAYLORD"]:
			app.errorBox("ErrorBox", "WARNING! Non-LUS pick ticket assigned to multiple gaylords!")

		entry["shipmentControlNumber"] = entry["shipmentControlNumber"].replace("TAIW", SCAC_code)
		
		#International Orders (shipped to IMS)
		if entry["consignee"]["address"]["country"] != "US":
			entry["consignee"]["address"]["addressLine"] = "2540 Walden Ave Suite 450"
			entry["consignee"]["address"]["country"] = "US"
			entry["consignee"]["address"]["city"] = "Buffalo"
			entry["consignee"]["address"]["stateProvince"] = "NY"
			entry["consignee"]["address"]["postalCode"] = "14225"
		#Name
		if len(entry["consignee"]["name"]) <= 2:
			entry["consignee"]["name"] = entry["consignee"]["name"].ljust(3, "A")
		elif len(entry["consignee"]["name"]) >= 60:
			entry["consignee"]["name"] = entry["consignee"]["name"][:59]
		#Address
		if len(entry["consignee"]["address"]["addressLine"]) <= 2:
			entry["consignee"]["address"]["addressLine"] = entry["consignee"]["address"]["addressLine"].ljust(3, "A")
		elif len(entry["consignee"]["address"]["addressLine"]) >= 55:
			entry["consignee"]["address"]["addressLine"] = entry["consignee"]["address"]["addressLine"][:54]
		#City
		if len(entry["consignee"]["address"]["city"]) <= 3:
			entry["consignee"]["address"]["city"] = entry["consignee"]["address"]["city"].ljust(2, "A")
		elif len(entry["consignee"]["address"]["city"]) >= 30:
			entry["consignee"]["address"]["city"] = entry["consignee"]["address"]["city"][:29]
		#State
		states_list = ["AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", "AS", "DC", "GU", "MP", "PR", "VI"]
		if entry["consignee"]["address"]["stateProvince"] not in states_list:
			entry["consignee"]["address"]["stateProvince"] = "NY"
		#Zip
		if len(entry["consignee"]["address"]["postalCode"]) != 5:
			entry["consignee"]["address"]["postalCode"] = entry["consignee"]["address"]["postalCode"].rjust(5, "0")

		#Clean non-alphanumeric Characters
		entry["consignee"]["name"] = clean(entry["consignee"]["name"])
		entry["consignee"]["address"]["addressLine"] = clean(entry["consignee"]["address"]["addressLine"])
		entry["consignee"]["address"]["city"] = clean(entry["consignee"]["address"]["city"])

		#Check for duplicates
		if entry not in out_json:
			out_json.append(entry)
		else:
			print(f"Duplicate entry removed: {entry['ORDERID']}")
			#error_file.write(f"Duplicate entry removed: {entry['ORDERID']}")
	return out_json

def createIMSBoL():
	print("\n>Creating IMS BoL")
	updateGaylordCounts()
	global gaylords_assignments
	global detailed_report_json
	if gaylords_assignments == [] or detailed_report_json == []:
		errorBox(">ERROR: No gaylords found. Please generate ACE/Detailed Report first")
	else: 
		fedex_gaylords = []
		fedex_count = 0
		dhl_gaylords = []
		dhl_count = 0
		for gaylord in gaylords_assignments:
			if "FEDEX" in gaylord:
				fedex_gaylords.append(gaylord[0])
				fedex_count = fedex_count + 1
			elif ("DHLGLOBALMAIL" in gaylord) or ("DHLGLOBALMAILV4" in gaylord):
				dhl_gaylords.append(gaylord[0])
				dhl_count = dhl_count + 1

		#Fedex Package Count
		global fedex_package_count
		fedex_package_count = 0
		for entry in detailed_report_json:
			if entry["carrier"] == "FEDEX":
				fedex_package_count += 1
	
	file_name = os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-IMS-BoL.pdf"
	image_file_path = os.getcwd() + os.sep + "resources" + os.sep + "IMS_BOL.jpg"
	c = canvas.Canvas(file_name, pagesize = (1668, 1986), bottomup = 0)
	image = ImageReader(image_file_path)
	c.drawImage(image, 0, 0, mask = "auto")

	c.setFont("Courier", 36)
	c.drawString(1070, 184, app.getEntry("Date:"))
	c.drawString(1145, 760, str(fedex_count + dhl_count))
	c.drawString(240, 1463, str(fedex_count) + " pkgs: " + str(fedex_package_count))
	c.drawString(705, 1463, str(dhl_count))

	for i, g in enumerate(fedex_gaylords):
		c.drawString(16, 1508 + (i * 36), g)

	for i, g in enumerate(dhl_gaylords):
		c.drawString(520, 1508 + (i * 36), g)

	c.showPage()
	c.save()

	print(">>IMS BoL created")

def updateGaylordCounts():
	##Gaylord counts
		usps_count = 0
		dhl_count = 0
		fedex_count = 0
		global gaylords_assignments
		if gaylords_assignments == []:
			errorBox(">ERROR: No gaylords found. Please generate ACE/Detailed Report first")
		else:
			for gaylord in gaylords_assignments:
				if "EHUB" in gaylord:
					usps_count = usps_count + 1
				elif "FEDEX" in gaylord:
					fedex_count = fedex_count + 1
				elif ("DHLGLOBALMAIL" in gaylord) or ("DHLGLOBALMAILV4" in gaylord):
					dhl_count = dhl_count + 1
			app.setEntry("USPS Gaylord Count:", str(usps_count))
			app.setEntry("DHL Gaylord Count:", str(dhl_count))
			app.setEntry("FedEx Gaylord Count:", str(fedex_count))
			app.setEntry("Total Gaylord Count:", str(usps_count + dhl_count + fedex_count))

def createLoadSheet():
	try:
		print("\n>Creating Load Sheet")
		updateGaylordCounts()
		global gaylords_assignments
		if gaylords_assignments == []:
			print("/n>ERROR: No gaylords found. Please generate ACE/Detailed Report first")

		#Write to file
		file_name = os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:") + "-Load_Sheet.pdf"
		c = canvas.Canvas(file_name, pagesize = (595.27, 841.89), bottomup = 0)
		c.setFont("Courier", 12)
		c.drawString(10, 22, app.getEntry("Date:") + " Load Sheet")
		c.drawString(10, 50, "SKID CARRIER       FDA? PACKAGES")
		for i, row in enumerate(gaylords_assignments, start = 3):
			c.drawString(10, (22 + i * 14), (row[0].ljust(5, " ") + row[1].ljust(14, " ") + row[2].ljust(5, " ") + str(row[3])))
		c.showPage()
		c.save()
		print(">>Load Sheet generation complete")
	except Exception:
		traceback.print_exc()

def emailPaperwork():
	#Check if all files exist
	print("\n>Attempting to Send Email")
	
	folder = os.getcwd() + os.sep + app.getEntry("Date:") + os.sep + app.getEntry("Date:")
	if os.path.exists(folder + "-ACE.json") and \
	   os.path.exists(folder + "-Detailed_Report.csv") and \
	   os.path.exists(folder + "-IMS-BoL.pdf") and \
	   os.path.exists(folder + "-Load_Sheet.pdf") and \
	   os.path.exists(folder + "-Stalco-BoL.pdf") and \
	   app.getEntry("ProFormaFileEntry") != "" and \
	   app.getEntry("ProFormaFileEntry")[-4:] == ".pdf":
		try:
			#Get files
			files = []
			files.append(folder + "-ACE.json")
			files.append(folder + "-Detailed_Report.csv")
			files.append(folder + "-IMS-BoL.pdf")
			files.append(folder + "-Load_Sheet.pdf")
			files.append(folder + "-Stalco-BoL.pdf")
			files.append(app.getEntry("ProFormaFileEntry"))
			
			# EMAIL
			global data
			username = app.stringBox("Username?", "Please enter your Outlook email address:", parent = None)
			password = app.stringBox("Password?", "Please enter the password for {}:".format(username), parent = None)
			
			smtp = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
			smtp.starttls()
			smtp.login(username, password)

			message = MIMEMultipart()
			message_text = (app.getTextArea("EmailTextArea")  + "\n" \
						   + app.getEntry("Total Gaylord Count:") + " Gaylords total:\n"
						   + app.getEntry("USPS Gaylord Count:") + " for USPS, " + app.getEntry("DHL Gaylord Count:") + " for DHL, " + app.getEntry("FedEx Gaylord Count:") + " for FedEx\n"
						   "Please hit \"REPLY ALL\" when responding to this email trail.")
			
			message["From"] = username
			recipients = []
			for recipient in data["emailRecipients"]:
				recipients.append(recipient["emailAddress"])
			message["To"] = ", ".join(recipients)
			#message["To"] = "james@stalco.ca"
			#message["Subject"] = "SAMPLE EMAIL > Stalco > Buffalo Run > " + app.getEntry("Date:")
			message["Subject"] = "Stalco > Buffalo Run > " + app.getEntry("Date:")
			message.attach(MIMEText(message_text, "plain"))

			for path in files:
				part = MIMEBase('application', "octet-stream")
				with open(path, 'rb') as file:
					part.set_payload(file.read())
				encoders.encode_base64(part)
				part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(Path(path).name))
				message.attach(part)

			smtp.send_message(message)
			del message
			smtp.quit()

			emailFedex(username, password)

			print(">>Email sent successfully")
		except:
			print(">>Email failed!")
			traceback.print_exc()
	else:
		print(">>Not all files created that are required for emailing")

def emailFedex(username, password):
	if fedex_package_count != 0:
		try:
			print("\n>FedEx packages found. Emailing FedEx")
			smtp = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
			smtp.starttls()
			smtp.login(username, password)

			message = MIMEMultipart()
			message_text = (f"Morning\n\nWe are dropping off {fedex_package_count} FedEx packages today and would like to request a pickup for the next available business day\n\nThank you!")
			
			message["From"] = username
			global data
			message["To"] = data["fedexContact"]
			message["Subject"] = "Stalco > FedEx Pickup Request > " + app.getEntry("Date:")
			message.attach(MIMEText(message_text, "plain"))

			smtp.send_message(message)
			del message
			smtp.quit()
			print(">>FedEx Email Sent")
		except:
			errorBox("Error sending FedEx email")

def loadACEManifest():
	try:
		global ACE_data
		filename = app.getEntry("ACEManifestFileEntry2")
		with open(filename, "r") as ACE_file:
			#print(filename)
			#print(ACE_file)
			ACE_data = json.load(ACE_file)
			app.setLabel("ACEStatusLabel", "{} ACE Entries Loaded".format(str(len(ACE_data))))
			SCN_ending = ACE_data[0]["shipmentControlNumber"][-2:]
			app.setLabel("SCNLabel", "SCNs currently end with: {}".format(SCN_ending))
	except FileNotFoundError: errorBox("No file entered")
	except: errorBox("Uploaded file was not an ACE Manfiest")

def loadACEEntry():
	try:
		global ACE_data
		match = False
		for entry in ACE_data:
			if entry["shipmentControlNumber"] == app.getEntry("SCN #:"):
				match = True
				app.setEntry("ORDERID:", entry["ORDERID"])
				app.setEntry("BATCHID:", entry["BATCHID"])
				app.setEntry("NAME:", entry["consignee"]["name"])
				app.setEntry("ADDRESS:", entry["consignee"]["address"]["addressLine"])
				app.setEntry("COUNTRY:", entry["consignee"]["address"]["country"])
				app.setEntry("CITY:", entry["consignee"]["address"]["city"])
				app.setEntry("STATE:", entry["consignee"]["address"]["stateProvince"])
				app.setEntry("POSTALCODE:", entry["consignee"]["address"]["postalCode"])
				try: app.setEntry("GAYLORD:", entry["GAYLORD"])
				except: pass
		if not match: raise Exception
	except: errorBox("SCAC does not correspond to an entry in this ACE")
	
def saveACEManifest():
	global ACE_data
	for entry in ACE_data:
		if entry["shipmentControlNumber"] == app.getEntry("SCN #:"):
			entry["ORDERID"] = app.getEntry("ORDERID:")
			entry["BATCHID"] = app.getEntry("BATCHID:")
			entry["consignee"]["name"] = app.getEntry("NAME:")
			entry["consignee"]["address"]["addressLine"] = app.getEntry("ADDRESS:")
			entry["consignee"]["address"]["country"] = app.getEntry("COUNTRY:")
			entry["consignee"]["address"]["city"] = app.getEntry("CITY:")
			entry["consignee"]["address"]["stateProvince"] = app.getEntry("STATE:")
			entry["consignee"]["address"]["postalCode"] = app.getEntry("POSTALCODE:")
			if app.getEntry("GAYLORD:") != "":
				try: entry["GAYLORD"] = app.getEntry("GAYLORD:")
				except: pass
	with open(app.getEntry("ACEManifestFileEntry2"), "w") as ACE_file:
		json.dump(ACE_data, ACE_file, indent = 4)

def removeGaylord():
	global ACE_data
	gaylord = app.getEntry('Gaylord (eg. "G1"):')
	good_entries = []
	bad_entries = []
	try:
		for entry in ACE_data:
			if entry["GAYLORD"] != gaylord:
				good_entries.append(entry)
			else:
				bad_entries.append(entry)
	except:
		errorBox("No enties for for Gaylord {}".format(gaylord))
	print(f"{str(len(bad_entries))} entries removed for {gaylord}")
	ACE_data = good_entries
	with open(app.getEntry("ACEManifestFileEntry2"), "w") as ACE_file:
		json.dump(ACE_data, ACE_file, indent = 4)
	with open(app.getEntry("ACEManifestFileEntry2") + "-REMOVED_ENTRIES", "w") as ACE_file:
		json.dump(bad_entries, ACE_file, indent = 4)
			
def doEverything():
	createMasterJSON()
	createACE()
	createProForma()
	createLoadSheet()
	createBoL()
	createIMSBoL()

def errorBox(input_string):
	app.errorBox("ErrorBox", input_string)
	try: traceback.print_exc()
	except: pass

def debug(in_string):
	global debug_flag
	if debug_flag:
		print(in_string)

def convertJSONToCSV():
    #Process JSON file
    print("\n>Converting JSON file to CSV")
    with open(app.getEntry("JSON")) as json_file:
        data = json.load(json_file)
    data_file = open("ACE_Manifest_(CSV).csv", "w", newline="")
    csv_writer = csv.writer(data_file)

    for l in data: #for line in data:
        #try:
        ''' 2020-11-26 Problem with Province Of Loading. Hardcoding for the time being
        #DHL Manifests don't have Province of Loading T.T
        try:
            _province_of_loading = l["provinceOfLoading"],
        except:
            _province_of_loading = "ON"
        '''
        try:
            _consignee_province = l["consignee"]["address"]["stateProvince"]
            _consignee_postal_code = l["consignee"]["address"]["postalCode"]
        except:
            _consignee_province = ""
        try:
            _shipper_name = l["shipper"]["name"]
            _shipper_address = l["shipper"]["address"]["addressLine"]
            _shipper_country = l["shipper"]["address"]["country"]
            _shipper_city = l["shipper"]["address"]["city"]
            _shipper_province = l["shipper"]["address"]["stateProvince"]
            _shipper_postal_code = l["shipper"]["address"]["postalCode"]
        except:
            _shipper_name = "Stalco Inc."
            _shipper_address = "401 Clayson Road"
            _shipper_country =  "CA"
            _shipper_city = "Toronto"
            _shipper_province = "ON"
            _shipper_postal_code = "M9M2H4"
        
        head = ( #Doing it manually for now. This format doesn't change often
            l["ORDERID"],
            l["BATCHID"],
            l["data"],
            l["type"],
            l["shipmentControlNumber"],
            # Defaults for when ACE is missing entries
            #_province_of_loading,
            "ON", #2020-11-26 hardcoding Temporarily
            _shipper_name,
            _shipper_address,
            _shipper_country,
            _shipper_city,
            _shipper_province,
            _shipper_postal_code,
            l["consignee"]["name"],
            l["consignee"]["address"]["addressLine"],
            l["consignee"]["address"]["country"],
            l["consignee"]["address"]["city"],
            _consignee_province,
            _consignee_postal_code
        )
        for i, commodity in enumerate(l["commodities"]): #for commodity in line["commodities"]
            body = (
                l["commodities"][i]["description"],
                l["commodities"][i]["quantity"],
                l["commodities"][i]["packagingUnit"],
                l["commodities"][i]["weight"],
                l["commodities"][i]["weightUnit"]
            )
            #if l["commodities"][i]["value"] != "":
            if "value" in l["commodities"][i].keys():
                body = body + (l["commodities"][i]["value"],)
            if "countryOfOrigin" in l["commodities"][i]:
                body = body + (l["commodities"][i]["countryOfOrigin"],)
            row = head + body
            csv_writer.writerow(row)
        #except:
            #print("Error on order ID {}".format(l["ORDERID"],))
    data_file.close()
    print("Finished converting JSON")
    print("Outputting to \"ACE_Manifest_(CSV).csv\"")

def convertCSVToJSON():
    print("\n>Converting CSV file to JSON")
    with open(app.getEntry("CSV")) as csv_file:
    #with open("C:\\Users\\Alex\\Desktop\\Alex's Workspace\\BatchRemover\\output.csv") as csv_file:
        csv_reader = csv.reader(csv_file, delimiter = ',')
        csv_data = []
        for row in csv_reader:
            csv_data.append(row)
            
        #Makes a list of consignees to add to JSON
        consignees = []
        for row in csv_data:
            if row[4] not in consignees:
                consignees.append(row[4])

        #For each consignee, add each entry to JSON
        out_json = []
        for consignee in consignees:
            #print("Building {}".format(consignee))
            entry = {}
            for row in csv_data:
                if consignee == row[4]:
                    #print("Match found for {}".format(consignee))
                    entry = {
                        "ORDERID": row[0],
                        "BATCHID": row[1],
                        "data": row[2],
                        "type": row[3],
                        "shipmentControlNumber": row[4],
                        "provinceOfLoading": row[5],
                        "shipper": {
                            "name": row[6],
                            "address": {
                                "addressLine": row[7],
                                "country": row[8],
                                "city": row[9],
                                "stateProvince": row[10],
                                "postalCode": row[11].zfill(5)
                            }
                        },
                        "consignee": {
                            "name": row[12],
                            "address": {
                                "addressLine": row[13],
                                "country": row[14],
                                "city": row[15],
                                "stateProvince": row[16],
                                "postalCode": row[17].zfill(5)
                            }
                        },
                        "commodities": []
                    }

            for row in csv_data: #Searches for commodities that match consignee
                if consignee == row[4]:
                    commodity = {}
                    if len(row) == 25: #If the entry has value and countryOfOrigin
                        commodity = {
                            "description": row[18],
                            "quantity": float(row[19]),
                            "packagingUnit": row[20],
                            "weight": int(row[21]),
                            "weightUnit": row[22],
                            "value": row[23],
                            "countryOfOrigin": row[24]
                        }
                    elif len(row) == 24: #If it has only value
                        commodity = {
                            "description": row[18],
                            "quantity": float(row[19]),
                            "packagingUnit": row[20],
                            "weight": int(row[21]),
                            "weightUnit": row[22],
                            "value": row[23],
                        }
                    else:
                        commodity = {
                            "description": row[18],
                            "quantity": float(row[19]),
                            "packagingUnit": row[20],
                            "weight": int(row[21]),
                            "weightUnit": row[22]
                        }
                    entry["commodities"].append(commodity)
            out_json.append(entry)
        with open("ACE_Manifest_(JSON).json", "w") as json_file:
            json.dump(out_json, json_file, indent=4)
        print(">Done converting CSV to JSON")
        print("Outputting to \"ACE_Manifest_(JSON).json\"")
    
def removeOrders():
    print("\n>Splitting ACE")
    orders = app.getTextArea("ordersTextArea")
    orders = orders.replace("\n", ",")
    orders_list = orders.split(",")
    
    with open(app.getEntry("ACE")) as json_file:
        json_data = json.load(json_file)
        data = json_data.copy()

    #Split out orders
    split_orders = []
    for entry in json_data:
        #Match Transaction/Batch ID
        for order in orders_list:
            if entry["ORDERID"] == order or entry["BATCHID"] == order:
                split_orders.append(entry)
                data.remove(entry)
                print("Removed order {} for {}".format(order, entry["consignee"]["name"]))
    #Output
    with open("ACE_Manifest_(1).json", "w") as json_file:
            json.dump(data, json_file, indent = 4)
    with open("ACE_Manifest_(2).json", "w") as bad_orders_file:
            json.dump(split_orders, bad_orders_file, indent = 4)
    print("Finished splitting ACE")
    print("Outputting original ACE to \"ACE_Manifest_(1).json\"")
    print("Outputting split entries to \"ACE_Manifest_(2).json\"")

def jsonBeautifier():
    print("\n>Formatting JSON")
    json_file_name = app.getEntry("Ugly JSON")
    with open(json_file_name, "r") as json_file:
        json_data = json.load(json_file)
    out_file_name = "ACE_Manifest_(Cleaned).json"
    with open(out_file_name, "w") as json_file:
        json.dump(json_data, json_file, indent = 4)
    print("Done formatting JSON")
    print("Outputting to \"ACE_Manifest_(Cleaned).json\"")

def combineJSON():
    print(">Combining JSONs")
    out_data = []
    with open(app.getEntry("JSON 1"), "r") as json_file_1:
        json_data_1 = json.load(json_file_1)
    with open(app.getEntry("JSON 2"), "r") as json_file_2:
        json_data_2 = json.load(json_file_2)
    for line in json_data_1:
        out_data.append(line)
    for line in json_data_2:
        out_data.append(line)
    with open("ACE_Manifest_(Combined).json", "w") as json_file:
        json.dump(out_data, json_file, indent = 4)
    print("Done combining JSONs")
    print("Outputting to \"ACE_Manifest_(Combined).json\"")

def changeSCNs():
	new_SCN = app.getEntry("New 2 digits:")[:2]
	for entry in ACE_data:
		entry["shipmentControlNumber"] = entry["shipmentControlNumber"][:14] + new_SCN
	with open(app.getEntry("ACEManifestFileEntry2"), "w") as ACE_file:
		json.dump(ACE_data, ACE_file, indent = 4)

def createUSGR():
	print("\n>Preparing USGR")
	with open(os.getcwd() + os.sep + "resources" + os.sep + "USGR_MASTER_LIST.csv", "r") as USGR_master_file:
		csv_reader = csv.reader(USGR_master_file, delimiter = ',')
		master_data = []
		for row in csv_reader:
			master_data.append(row)

	with open(app.getEntry("ProForma Template:"), "r") as proforma_file:
		csv_reader = csv.reader(proforma_file, delimiter = ',')
		proforma_data = []
		for row in csv_reader:
			proforma_data.append(row)

	out_data = []
	for line in proforma_data:
		for row in master_data:
			if line[0] == row[1] and row[10] == "US": #If ProForma SKU matches a SKU from the master file
				out_data.append((row[1], row[13], row[16], row[20], row[21], line[11], line[1], row[2], "BUFFALO", "US"))
	
	file_name = os.getcwd() + os.sep + "USGR" + os.sep + app.getEntry("USGR Date:") + "-USGR_Table-" + app.getEntry("USGR Entry Number:") + ".pdf"
	c = canvas.Canvas(file_name, pagesize = (595.27, 841.89), bottomup = 1)
	c.setFont("Courier", 7)
	c.drawString(10, 820, "US GOODS RETURNED")
	c.drawString(10, 812, "INVOICE REFERENCE #: {}    ENTRY #: {}    DATE: {}".format(app.getEntry("USGR BoL #:"), app.getEntry("USGR Entry Number:"), app.getEntry("USGR Date:")))
	c.drawString(10, 800, "PART/ITEM #       DESCRIPTION                     MANUFACTURER           CITY, STATE           VALUE     QTY  IMPORTDATE  ENTRY PORT")
	for i, row in enumerate(out_data, start = 1):
		out = ""
		out += row[0][:17].ljust(18, " ")
		out += row[1][:30].ljust(32, " ")
		out += row[2][:22].ljust(23, " ")
		out += (row[3][:17] + ", " + row[4][:2]).ljust(22, " ")
		out += row[5][:9].ljust(10, " ")
		out += row[6][:4].ljust(5, " ")
		out += row[7][:11].ljust(12, " ")
		out += "BUFFALO US"
		c.drawString(10, (800 - (i * 8)), str(out))
	c.showPage()

	c.setFont("Courier", 12)
	for i in range(1, 7): #1 through 6 inclusive
		filepath = os.getcwd() + os.sep + "resources" + os.sep + "page_" + str(i) + ".jpg"
		image = Image.open(filepath)
		c.drawImage(ImageReader(image), 0, 0, 595.27, 841.89)
		if i == 1:
			c.drawString(180, 615, app.getEntry("USGR Entry Number:"))
			c.drawString(180, 595, app.getEntry("USGR BoL #:"))
		if i == 2:
			c.drawString(130, 660, app.getEntry("USGR Entry Number:"))
			c.drawString(130, 640, app.getEntry("USGR BoL #:"))
		if i == 3:
			c.drawString(170, 605, app.getEntry("USGR Entry Number:"))
			c.drawString(170, 585, app.getEntry("USGR BoL #:"))
		c.showPage()

	c.save()
	print(">>USGR Completed for " + app.getEntry("USGR Date:"))

def updateUSGRdata():
	app.setEntry("USGR Date:", app.getEntry("Date:"))
	app.setEntry("USGR BoL #:", app.getEntry("BoL #:"))

def copyPaperwork():
	print("\n>Attempting to copy paperwork")
	try:
		date = app.getEntry("Date:")
		src = os.getcwd() + os.sep + date
		year = date.split("-")[0]
		month = date.split("-")[1]
		dst = "W:\\Logistics\\USPS Customs\\USPS Customs Paperwork\\IMS Invoices - Mixed Shipments\\{year}\\{month}\\{date}".format(year = year, month = month, date = date)
		if not os.path.exists(dst):
			os.mkdir(dst)
		copyTree(src, dst)
		print(">>Paperwork copying completed")
	except:
		errorBox("Copying files failed!")

def copyTree(src, dst, symlinks=False, ignore=None): #Stolen from SlackOverflow
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)	

def saveMasterData():
	pass

def loadMasterData():
	pass

### Start

print("LogisticsManager.py v1.0")

## Global variables
debug_flag = False
detailed_report_json = []
data = {}
master_json_data = []
ACE_data = {}
gaylords_assignments = []
fedex_package_count = 0

## Start GUI
app = gui()

app.startTabbedFrame("TabbedFrame")
app.setSticky("nw")
app.setStretch("none")

## Frame 1
app.startTab("BASIC")

app.startLabelFrame("Date/BoL/PAPS")
app.addLabel("W:\\Logistics\\Carrier Tracking\\USPS Tracking.xlsx")
app.addLabelEntry("Date:")
app.addLabelEntry("BoL #:")
app.addLabelEntry("PAPS #:")
app.stopLabelFrame()

app.startLabelFrame("Step 1")
app.addLabel("ACE Manfiest (from Techship, \"ace_manifest_#\"):")
app.addFileEntry("ACEManifestFileEntry")
app.addLabel("Batches Scans:")
app.addFileEntry("batchesFileEntry")
app.addLabel("Report CSV (from Techship, \"manifest_packages_#\"):")
app.addFileEntry("CSVReportFileEntry")
app.addButton("Create Paperwork", doEverything)
app.stopLabelFrame()

app.startLabelFrame("Step 2:")
app.addLabel("ProForma (Printed from SmartBorder):")
app.addFileEntry("ProFormaFileEntry")
app.addButton("Email Paperwork", emailPaperwork)
app.addButton("Move Paperwork to W: Drive", copyPaperwork)
app.stopLabelFrame()

app.stopTab()

## Frame 2
app.startTab("ADVANCED")

app.startLabelFrame("Variables (Auto-generated, manually-editable)")
app.addLabelEntry("SCAC:")
app.addLabelEntry("Total Gaylord Count:")
app.addLabelEntry("Package Count:")
app.addLabelEntry("Total Weight:")
app.addLabelEntry("USPS Gaylord Count:")
app.addLabelEntry("DHL Gaylord Count:")
app.addLabelEntry("FedEx Gaylord Count:")
app.stopLabelFrame()

app.startLabelFrame("Advanced Buttons")
app.addButton("Load Master Data", loadMasterData)
app.addButton("Save Master Data", saveMasterData)
app.addButton("Advance (increase by 1) BoL #/PAPS #", updateVariables)
app.addButton("Save BoL #/PAPS #", saveVariables)
app.addButton("Create Master JSON", createMasterJSON)
app.addButton("Create ACE and Detailed Report", createACE)
app.addButton("Create ProForma", createProForma)
app.addButton("Create Load Sheet", createLoadSheet)
app.addButton("Create BoL", createBoL)
app.addButton("Create IMS BoL", createIMSBoL)
app.addButton("Email Papers", emailPaperwork)
app.stopLabelFrame()

app.stopTab()

## Frame 3
app.startTab("ACE EDITING")
app.startLabelFrame("ACE Manifest")
app.addLabel("ACE Manifest (.json):")
app.addFileEntry("ACEManifestFileEntry2")
app.addButton("Load ACE", loadACEManifest)
app.addLabel("ACEStatusLabel", "No ACE Loaded")
app.stopLabelFrame()

app.startLabelFrame("Entry Editing")
app.addLabelEntry("SCN #:")
app.addButton("Load Entry", loadACEEntry)
app.addLabelEntry("ORDERID:")
app.addLabelEntry("BATCHID:")
app.addLabelEntry("NAME:")
app.addLabelEntry("ADDRESS:")
app.addLabelEntry("COUNTRY:")
app.addLabelEntry("CITY:")
app.addLabelEntry("STATE:")
app.addLabelEntry("POSTALCODE:")
app.addLabelEntry("GAYLORD:")
app.addButton("Save Entry", saveACEManifest)
app.stopLabelFrame()

app.startLabelFrame("Gaylord Removal")
app.addLabelEntry('Gaylord (eg. "G1"):')
app.addButton("Remove Gaylord", removeGaylord)
app.stopLabelFrame()

app.startLabelFrame("Duplicate SCN Editor")
app.addLabel("SCNLabel", "SCNs currently end with: NA")
app.addLabelEntry("New 2 digits:")
app.addButton("Change SCNs", changeSCNs)
app.stopLabelFrame()

app.stopTab()

## Frame 4
app.startTab("EMAIL TEXT")
app.startLabelFrame("Email Editing:")
app.setStretch("both")
app.setSticky("nesw")
app.addTextArea("EmailTextArea", text = "Morning\n\nPlease see attached paperwork for today\nPickup is available 9am next business day")
app.setSticky("ws")
app.addButton("Send Custom Email", emailPaperwork)
app.stopLabelFrame()
app.stopTab()

## Frame 5

app.startTab("MANUAL EDIT")
app.startLabelFrame("Manual Editing:")
app.addLabelFileEntry("JSON")
app.addButton("Convert to CSV", convertJSONToCSV)
app.addLabelFileEntry("CSV")
app.addButton("Convert to JSON", convertCSVToJSON)
app.stopLabelFrame()

app.startLabelFrame("Batch Removal")
app.addLabelFileEntry("ACE")
app.addLabel("Batches/Transactions")
app.addTextArea("ordersTextArea")
app.addButton("Remove/Split", removeOrders)
app.stopLabelFrame()

app.startLabelFrame("JSON Formatter")
app.addLabelFileEntry("Ugly JSON")
app.addButton("Format JSON", jsonBeautifier)
app.stopLabelFrame()

app.startLabelFrame("JSON Combiner")
app.addLabelFileEntry("JSON 1")
app.addLabelFileEntry("JSON 2")
app.addButton("Combine", combineJSON)
app.stopLabelFrame()
app.stopTab()

## USGR

app.startTab("USGR")
app.startLabelFrame("USGR")
app.addLabelFileEntry("ProForma Template:")
app.addLabelEntry("USGR Date:")
app.addLabelEntry("USGR Entry Number:")
app.addLabelEntry("USGR BoL #:")
app.addButton("Create USGR", createUSGR)
app.stopLabelFrame()
app.stopTab()

## Finish
app.stopTabbedFrame()

### Setup

loadVariables()
updateVariables()
app.setEntry("SCAC:", "TAIW")
updateUSGRdata()

app.go()
