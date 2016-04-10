# batch-kimble-travel
Little app that can take values populated in a .csv files and use them to create a travel request in Kimble Salesforce. 

For Apple Map only.

**Prerequisites**

* Open Kimble in Safari (assumes you are logged in)
* You need to copy the source into Apple Script Editor and save as a script file - save anywhere but put the script in the same folder as the csv and it will all work.
* To access the Apple Script Editor, use the Script Editor application.
* Assumes a single set of dates - out and return

**How this works**

1. Use the csv template to enter the details of the travel you wish to request in Kimble
2. Do NOT muck around with the headings, and don't use any commas. Make sure you save it as a csv, NOT xlsx
3. Save that file (**don't rename it**) in the same folder that this app will live in.  It expects it to be co-located. Not planning on making this particularly robust, so use with caution.  No intention to set these to auto-submit.

List of values for airports:

* Any Belfast 
* Any London 
* BHD - Belfast City 
* BFS - Belfast Int 
* LDY - City of Derry 
* DUB - Dublin 
* ORK - Cork 
* GDN - Gdansk 
* LCY - London City Airport 
* LGW - London Gatwick 
* LHR - London Heathrow 
* LTN - London Luton 
* SEN - London Southend 
* STN - London Stansted 
* BHX - Birmingham 
* BRS - Bristol 
* CWL - Cardiff 
* EMA - East Midlands 
* EDI - Edinburgh 
* GLA - Glasgow Int 
* PIK - Glasgow Prestwick 
* LBA - Leeds-Bradford 
* LPL - Liverpool 
* MAN - Manchester 
* NCL - Newcastle 
* SOU - Southampton 
* SWS - Swansea 
* ATL - Atlanta 
* BOS - Boston Logan 
* EIN - Eindhoven 
* EWR - Newark 
* ORD - Chicago 
* GVA - Geneva 
* JFK - New York 
* LGA - LaGuardia 
* MCO - Orlando Int 
* PHL - Philadelphia 
* SFO - San Fransisco 
* YYZ - Toronto 
* WAW - Warsaw 
* VRN - Verona 
* AMS - Amsterdam 
* LAX - Los Angeles 
* Not Listed