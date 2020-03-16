import openpyxl as xl
import urllib.request
import PIL.Image
import os
import json
import math
import time

class Location:
	def __init__(self):
		self.loc_ID = None
		self.address = None
		self.lat = None
		self.long = None
		self.postcode = None
		self.SAM_ID = None
		self.state = None
		self.ID = None
		self.status = None
		self.MRQ = None
		self.tech = None

	def print(self):
		print(self.loc_ID)
		print(self.address)
		print(self.lat)
		print(self.long)
		print(self.postcode)
		print(self.SAM_ID)
		print(self.state)
		print(self.status)
		print(self.MRQ)
		print(self.tech)

	def find_distance(self, loc):
		lat_self = self.lat
		lat_target = loc.lat
		long_self = self.long
		long_target = loc.long
		if lat_self == "NULL" or long_self == "NULL":
			print("base doesn't have lat or long")
			return -1
		if lat_target == "NULL" or long_target == "NULL":
			print("target doesn't have lat or long")
			return -1	
		return math.sqrt((lat_self - lat_target)**2 + (long_self - long_target)**2)

	def get_coordinates(self):
		address = self.address.strip() #get_address()
		encodedDest = urllib.parse.quote(address, safe='')
		bingkey = "AsuwbgYYc0ypzGFmASWf0VMPk9jme369wjy9uxZkddJk-7SGyT_txN0bSB88a1AM"
		Url = "http://dev.virtualearth.net/REST/v1/Locations/" + encodedDest + "?mapSize=900,900&key=" + bingkey
		#print(Url)
		request = urllib.request.Request(Url)
		time.sleep(2)
		response = urllib.request.urlopen(request)
		#opened = False
		'''i = 0
		while True:
			try:
				response = urllib.request.urlopen(request)
			except Exception as e:
				print(e)
				time.sleep(1)
				i = i + 1
				if i > 4:
					response = urllib.request.urlopen(request)
					print("time out getting coordinates " + str(address))
					break
			else:
				response = urllib.request.urlopen(request) #put this here to fix error but not sure why it works
				break
				#opened = True'''
		HTTPstring = response.read().decode('utf-8')
		json_obj = json.loads(HTTPstring)
		self.lat, self.long = json_obj["resourceSets"][0]["resources"][0]["geocodePoints"][0]["coordinates"]
		return self.lat, self.long

	def get_address(self):
		lati = str(self.lat)
		longi = str(self.long)
		encodedDest = lati + "," + longi 
		bingkey = "AsuwbgYYc0ypzGFmASWf0VMPk9jme369wjy9uxZkddJk-7SGyT_txN0bSB88a1AM"
		Url = "http://dev.virtualearth.net/REST/v1/Locations/" + encodedDest + "/?mapSize=900,900&key=" + bingkey
		#print(Url)
		request = urllib.request.Request(Url)
		#time.sleep(2)
		#response = urllib.request.urlopen(request)
		#opened = False
		while True:
			try:
				response = urllib.request.urlopen(request)
			except:
				time.sleep(1)
			else:
				break
				#opened = True
		HTTPstring = response.read().decode('utf-8')
		json_obj = json.loads(HTTPstring)
		self.address = str(json_obj["resourceSets"][0]["resources"][0]["address"]["formattedAddress"])
		self.postcode = str(json_obj["resourceSets"][0]["resources"][0]["address"]["postalCode"])
		return self.address
