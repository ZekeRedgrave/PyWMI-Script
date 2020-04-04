'''

The WMI command-line (WMIC) utility provides a command-line interface for Windows Management Instrument(WMI).
This Class is only use to get all Computer Information from Software to Hardware using Windows OS from Microsoft.

Author: Zeke Redgrave
Version: 1

'''
import xml.etree.ElementTree as ET, subprocess, json, os

class WMIC:
	def __init__(self):
		# Global Variable
		self.startupinfo = subprocess.STARTUPINFO()
		self.startupinfo.dwFlags = subprocess.STARTF_USESHOWWINDOW
		self.startupinfo.wShowWindow = subprocess.SW_HIDE

	def getBaseboard(self, key=""):
		'''
		Method: Return Dictionary
		Command: wmic baseboard list /format:rawxml

		Ex. print(WMIC().getBaseboard()['key'])
		'''
		Motherboard = {}
		root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic baseboard list /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

		for a in root[1][0][0]:
			for b in a:
				Motherboard[a.attrib['NAME'].lower()] = b.text

		return Motherboard if key == "" else Motherboard[key.lower()]

	def getProcessor(self, key=""):
		'''
		Method: Return Dictionary
		Command: wmic cpu list /format:rawxml
		
		Ex. print(WMIC().getComputerProcessingUnit()['key'])
		'''
		Processor = {}
		root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic cpu list /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

		for a in root[1][0][0]:
			for b in a:
				Processor[a.attrib['NAME'].lower()] = b.text

		return Processor if key == "" else Processor[key.lower()]

	def getMemoryCard(self, index=None):
		'''
		Method: Return List
		Command: wmic memorychip list /format:rawxml
		
		Ex. print(WMIC().getComputerProcessingUnit()[integer]) -->> Display String JSON Format
		Notice: To Convert String to JSON -->> json.loads(WMIC().getComputerProcessingUnit()[int])
		'''
		Memory = []
		_Memory = {}
		Count = 0
		root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic memorychip list /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

		for a in root[1][0]:
			Memory.append(Count)

			for b in a:
				for c in b:
					_Memory[b.attrib['NAME'].lower()] = c.text

			Memory[Count] = json.dumps(_Memory)
			_Memory = {}

			Count += 1

		return Memory if index == None else Memory[index]

	def getNetworkInterfaceCard(self, setEnable=True, index=None):
		'''
		Method: Return List
		Command: wmic nic where NetEnabled=true list /format:rawxml
		Parameter: setEnable=True -->> Default
		
		Ex. print(WMIC().getNetworkInterfaceCard()[integer]) -->> Display String JSON Format
		Notice: To Convert String to JSON -->> json.loads(WMIC().getNetworkInterfaceCard()[int])
		'''
		NIC = []
		_NIC = {}
		Count = 0

		if setEnable is True:
			root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic nic where NetEnabled=true list /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

			for a in root[1][0]:
				NIC.append(Count)

				for b in a:
					for c in b:
						_NIC[b.attrib['NAME'].lower()] = c.text

				NIC[Count] = json.dumps(_NIC)
				_NIC = {}

				Count += 1

		else:
			root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic nic list /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

			for a in root[1][0]:
				NIC.append(Count)

				for b in a:
					for c in b:
						_NIC[b.attrib['NAME'].lower()] = c.text

				NIC[Count] = json.dumps(_NIC)
				_NIC = {}

				Count += 1

		return NIC if index == None else NIC[index]

	def getStorage(self, index=None):
		'''
		Method: Return List
		Command: wmic diskdrive list /format:rawxml
		
		Ex. print(WMIC().getStorage()[integer]) -->> Display String JSON Format
		Notice: To Convert String to JSON -->> json.loads(WMIC().getStorage()[int])
		'''
		Storage = []
		_Storage = {}
		Count = 0
		root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic diskdrive list /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

		for a in root[1][0]:
			Storage.append(Count)

			for b in a:
				for c in b:
					_Storage[b.attrib['NAME'].lower()] = c.text

			Storage[Count] = json.dumps(_Storage)
			_Storage = {}

			Count += 1

		return Storage if index == None else Storage[index]

	def getDesktopMonitor(self, index=None):
		'''
		Method: Return List
		Command: wmic desktopmonitor list /format:rawxml
		
		Ex. print(WMIC().getDesktopMonitor()[integer]) -->> Display String JSON Format
		Notice: To Convert String to JSON -->> json.loads(WMIC().getDesktopMonitor()[int])
		'''
		DesktopMonitor = []
		_DesktopMonitor = {}
		Count = 0
		root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic desktopmonitor list /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

		for a in root[1][0]:
			DesktopMonitor.append(Count)

			for b in a:
				for c in b:
					_DesktopMonitor[b.attrib['NAME'].lower()] = c.text

			DesktopMonitor[Count] = json.dumps(_DesktopMonitor)
			_DesktopMonitor = {}

			Count += 1

		return DesktopMonitor if index == None else DesktopMonitor[index]

	def getOperatingSystem(self):
		'''
		Method: Return Dictionary
		Command: wmic os list /format:rawxml

		Ex. print(WMIC().getOperatingSystem()['key'])
		'''
		OS = {}
		root = ET.fromstring(str(subprocess.Popen(["powershell", "-command", "wmic os list brief /format:rawxml"], startupinfo=self.startupinfo, stdout=subprocess.PIPE).stdout.read().decode("utf-8")).replace("\r", "").replace("\n", "").replace("\t", ""))

		for a in root[1][0][0]:
			for b in a:
				OS[a.attrib['NAME'].lower()] = b.text

		return OS

	def getComputerName(self):
		'''
		Method: Return String

		Ex. print(WMIC().getComputerName())
		'''
		return os.environ['COMPUTERNAME']

	def getComputerUsername(self):
		'''
		Method: Return String

		Ex. print(WMIC().getComputerUsername())
		'''
		return os.environ['USERNAME']
