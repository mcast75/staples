import csv, sys
from xlrd import open_workbook

if __name__ == "__main__":
    file_in = raw_input("Excel File to Read: ") 
    out_file = raw_input("Destination File Name: ")
    if file_in[-4:] != '.xlsx':
    	file_in = file_in+'.xlsx'
    if out_file[-4:] != '.csv':
    	out_file = out_file+'.csv'

    file_in = 'staples_copy.xlsx'
    out_file = 'out.csv'

    wb = open_workbook(file_in)
    skip = ["Summary", "Glossary", "41_Missing"]


    with open(out_file, 'w') as output:
	    fieldnames = ['Product Family', "LDOS<FY16", "FY16", "FY17", "FY18", "FY19", "FY20", "FY21", "FY22", "Total", "Unlisted"]
	    out_writer = csv.DictWriter(output, fieldnames=fieldnames)
	    out_writer.writerow({'Product Family': 'Product Family', "LDOS<FY16":"LDOS<FY16", "FY16":"FY16", "FY17":"FY17", "FY18":"FY18", "FY19":"FY19", "FY20":"FY20", "FY21":"FY21", "FY22":"FY22", "Total":"Total", "Unlisted":"Unlisted"})

	    for sheet in wb.sheets():
			if sheet.name not in skip:
				family = sheet.name[2:]

				grid_ref = ['16', '17', '18', '19', '20', '21', '22']
				chassis_date_count = [0,0,0,0,0,0,0,0,0,0]
				module_date_count = [0,0,0,0,0,0,0,0,0,0]
				psu_date_count = [0,0,0,0,0,0,0,0,0,0]
				fan_date_count = [0,0,0,0,0,0,0,0,0,0]



				for row in range(4, sheet.nrows):
					ldos = str(sheet.cell(row, 26)) 
					#print ldos
					ldos_year = ldos[-3:-1]
					try:
						year_index = grid_ref.index(ldos_year)+1
					except:
						#print family + " - LDOS Year: "+ ldos
						year_index = 0

					part_type = str(sheet.cell(row, 2))[7:-1]

					if part_type =='Chassis':
						if ldos != "empty:u''":
							chassis_date_count[year_index] = chassis_date_count[year_index]+1
							chassis_date_count[8] = chassis_date_count[8]+1
						else:
							chassis_date_count[9] = chassis_date_count[9]+1
					elif part_type[:3] == 'PWR':
						if ldos != "empty:u''":
							psu_date_count[year_index] = psu_date_count[year_index]+1
							psu_date_count[8] = psu_date_count[8]+1
						else:
							psu_date_count[9] = psu_date_count[9]+1
					elif part_type[:3] == 'FAN':
						if ldos != "empty:u''":
							fan_date_count[year_index] = fan_date_count[year_index]+1
							fan_date_count[8] = fan_date_count[8]+1
						else:
							fan_date_count[9] = fan_date_count[9]+1
					else:
						if ldos != "empty:u''":
							module_date_count[year_index] = module_date_count[year_index]+1
							module_date_count[8] = module_date_count[8]+1
						else:
							module_date_count[9] = module_date_count[9]+1


				print family 
				print "Chassis"+ str(chassis_date_count)
				# print "Modules"+str(module_date_count)
				# print "PSU"+ str(psu_date_count)
				# print "Fans"+ str(fan_date_count)
				print
				out_writer.writerow({'Product Family': family})
				out_writer.writerow({'Product Family': 'Chassis', "LDOS<FY16":chassis_date_count[0], "FY16":chassis_date_count[1], "FY17":chassis_date_count[2], "FY18":chassis_date_count[3], "FY19":chassis_date_count[4], "FY20":chassis_date_count[5], "FY21":chassis_date_count[6], "FY22":chassis_date_count[7], "Total":chassis_date_count[8], "Unlisted": chassis_date_count[9]})
				# out_writer.writerow({'Product Family': 'Modules', "LDOS<FY16":module_date_count[0], "FY16":module_date_count[1], "FY17":module_date_count[2], "FY18":module_date_count[3], "FY19":module_date_count[4], "FY20":module_date_count[5], "FY21":module_date_count[6], "FY22":module_date_count[7], "Total":module_date_count[8],"Unlisted": module_date_count[9]})
				# out_writer.writerow({'Product Family': 'Power Supplies', "LDOS<FY16":psu_date_count[0], "FY16":psu_date_count[1], "FY17":psu_date_count[2], "FY18":psu_date_count[3], "FY19":psu_date_count[4], "FY20":psu_date_count[5], "FY21":psu_date_count[6], "FY22":psu_date_count[7], "Total":psu_date_count[8], "Unlisted": module_date_count[9]})
				# out_writer.writerow({'Product Family': 'Fans', "LDOS<FY16":fan_date_count[0], "FY16":fan_date_count[1], "FY17":fan_date_count[2], "FY18":fan_date_count[3], "FY19":fan_date_count[4], "FY20":fan_date_count[5], "FY21":fan_date_count[6], "FY22":fan_date_count[7], "Total":fan_date_count[8],"Unlisted": fan_date_count[9]})
				out_writer.writerow({'Product Family': ''})



