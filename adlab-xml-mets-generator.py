import datetime
import json
import sys
from openpyxl import load_workbook

# check arguments
if len(sys.argv) < 3:
    print("No parameter has been included")
    print(" 1) path of json formatted config file i.e. config.json")
    print(" 1) path of xlsx file with metadata")
    print("i.e. $ python adlab-xml-mets-generator.py config.json metadata.xlsx")
    sys.exit()

# get var from arguments
config = sys.argv[1]
xslxfile = sys.argv[2]

# get datetime
local = datetime.datetime.now()
creation = local.strftime("%Y-%m-%dT%H:%M:%S")

# get json config file and parse data
with open(config, "r") as f:
    configVars = json.load(f)

agency = configVars["agency"]
creator = configVars["creator"]
title = configVars["title"]
profile = configVars["profile"]

# open xlsx in read mode
wb = load_workbook(filename=xslxfile, read_only=True)
ws = wb['Sheet1']

#for row in ws.rows:
#    for cell in row:
#        print(cell.value)

# Close the workbook after reading
wb.close()

magId = "ASMO_CS_CPE_77"
bobina = "77"
totfotogrammi = "248"

# xml mag header creation
xmlMagHeader = '''<?xml version="1.0" encoding="UTF-8"?>
<METS:mets OBJID="'''+magId+'''" xmlns:METS="http://www.loc.gov/METS/" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dct="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:mods="http://www.loc.gov/mods/v3" xmlns:rights="http://cosimo.stanford.edu/sdr/metsrights/"
xsi:schemaLocation="http://www.loc.gov/METS/ http://www.loc.gov/standards/mets/mets.xsd http://www.loc.gov/mods/v3 http://www.loc.gov/mods/v3/mods-3-7.xsd http://cosimo.stanford.edu/sdr/metsrights/ https://www.loc.gov/standards/rights/METSRights.xsd">
	<METS:metsHdr CREATEDATE="'''+creation+'''" ID="'''+magId+'''" LASTMODDATE="'''+creation+'''" RECORDSTATUS="COMPLETE">
		<METS:agent ROLE="CREATOR" TYPE="ORGANIZATION">
			<METS:name>'''+agency+'''</METS:name>
		</METS:agent>
	</METS:metsHdr>
'''

# <mets:dmdSec ID="">
xmldmdSec = '''
	<METS:dmdSec ID="DMD1" STATUS="referenced">
		<METS:mdWrap MDTYPE="MODS" MIMETYPE="text/xml" LABEL="MODS Metadata">
			<METS:xmlData>
				<mods:mods>
					<mods:relatedItem otherType="digitalCollection">
						<mods:titleInfo>
							<mods:title>Progetto Carteggi Principi Estensi</mods:title>
						</mods:titleInfo>
					</mods:relatedItem>
				</mods:mods>
			</METS:xmlData>
		</METS:mdWrap>
	</METS:dmdSec>
'''

# <mets:amdSec>
xmlamdSec = '''
	<METS:amdSec ID="AMD1">
		<METS:rightsMD ID="BCS">
			<METS:mdWrap MDTYPE="METSRIGHTS" MIMETYPE="text/xml" LABEL="Rights Metadata">
				<METS:xmlData>
					<rights:RightsDeclarationMD>
						<rights:RightsHolder RIGHTSHOLDERID="ASMO">
							<rights:RightsHolderName>Archivio di Stato di Modena</rights:RightsHolderName>
							<rights:RightsHolderContact>
								<rights:RightsHolderContactEmail>as-mo@pec.cultura.gov.it</rights:RightsHolderContactEmail>
							</rights:RightsHolderContact>
						</rights:RightsHolder>
					</rights:RightsDeclarationMD>
				</METS:xmlData>
			</METS:mdWrap>
		</METS:rightsMD>
		<METS:sourceMD ID="sourceMD1">
			<METS:mdWrap MDTYPE="MODS" MIMETYPE="text/xml" LABEL="MODS Metadata">
				<METS:xmlData>
					<mods:mods>
						<mods:identifier type="localId">'''+bobina+'''</mods:identifier>
						<mods:location>
							<mods:physicalLocation>Archivio di Stato di Modena</mods:physicalLocation>
						</mods:location>
						<mods:originInfo eventType="production">
							<mods:dateCreated>1956</mods:dateCreated>
							<mods:place>
								<mods:placeTerm type="text">Modena</mods:placeTerm>
							</mods:place>
							<mods:publisher>Archivio di Stato di Modena</mods:publisher>
						</mods:originInfo>
						<mods:physicalDescription>
							<mods:form authority="marcform">microfilm</mods:form>
							<mods:extent>1 bobina di microfilm ('''+totfotogrammi+''' fotogrammi) : negativo, b/n, perforato ; 35 mm</mods:extent>
						</mods:physicalDescription>
					</mods:mods>
				</METS:xmlData>
			</METS:mdWrap>
		</METS:sourceMD>
	</METS:amdSec>
'''

# <mets:fileSec> !!!! da implementare !!!!
# ciclo nella cartella
# nome file senza e con estensione .tif
# md5sum
# get size
# -> creazione di più sezioni <METS:file
xmlamdSec = '''
	<METS:fileSec>
		<METS:fileGrp USE="INTERNAL">
			<METS:fileGrp USE="IMAGE">
				<METS:fileGrp USE="ARCHIVE">
					<METS:file ID="ARCHIVE-ASMO_CS_CPE_77_0001" MIMETYPE="image/tiff" CHECKSUM="n518e85786456887a57e1bdb31fe5890" CHECKSUMTYPE="MD5" SIZE="152724">
						<METS:FLocat LOCTYPE="OTHER" OTHERLOCTYPE="SYSTEM" xlink:href="./OUTPUT_ASMO_CS_CPE/ASMO_CS_CPE_77/ASMO_CS_CPE_77_0001.tif"/>
					</METS:file>
					<METS:file ID="ARCHIVE-ASMO_CS_CPE_77_0002" MIMETYPE="image/tiff" CHECKSUM="n518e85786456887a57e1bdb31fe5890" CHECKSUMTYPE="MD5" SIZE="152724">
						<METS:FLocat LOCTYPE="OTHER" OTHERLOCTYPE="SYSTEM" xlink:href="./OUTPUT_ASMO_CS_CPE/ASMO_CS_CPE_77/ASMO_CS_CPE_77_0002.tif"/>
					</METS:file>					
						...
						...
						...
					</METS:file>
				</METS:fileGrp>
			</METS:fileGrp>
		</METS:fileGrp>
	</METS:fileSec>
'''

# <mets:structMap> !!!! da implementare !!!!
# contatore per ORDER
# nome file senza estensione da aggiungere a "ARCHIVE-"
xmlstructMap = '''
	<METS:structMap TYPE="PHYSICAL">
		<METS:div LABEL="carteggio" TYPE="FOLDER">
			<METS:div LABEL="Document" ORDER="1" TYPE="FILE">
				<METS:fptr FILEID="ARCHIVE-ASMO_CS_CPE_77_0001"/>
			</METS:div>
			<METS:div LABEL="Document" ORDER="2" TYPE="FILE">
				<METS:fptr FILEID="ARCHIVE-ASMO_CS_CPE_77_0002"/>
			</METS:div>
			...
			...
			...
		</METS:div>
	</METS:structMap>
'''

print(xmlMagHeader)
