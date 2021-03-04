# Python module to extract tables from a model report file in PRN format,
# and store them in separate files
#
# NOTES: 
#   1. This module was written to run under Python 3.8.x
#
# Author: David Knudsen
# Date: 2 February 2021
#   
# Internals of this Module: Top-level Functions
# =============================================
#
# main(<input PRN report file>, <output_folder>, <tuple of table names/numbers to extract>)
#				  main driver routine for this program
#
# extract_table_to_file(<input PRN report file>, <table number/name>, <output file>)
#                 opens the file specified by the first parameter,
#				  scans it line by line to find the second parameter,
#				  and then outputs the lines between the next table start line
#				  and table stop line
#
# Internals of this Module: Utility Functions
# ===========================================
#
###############################################################################

import os
import sys
import re
import shutil

debug_flags = {}
debug_flags['dump_sched_elements'] = False

# Global pseudo-constants:


# extract individual table
def extract_table_to_file(prn_file, table, out_file):

	# prepare regular expressions that will match the table name/number
	#  and the last line preceding table data
	#  and the first line following table data
	re_table_title = re.compile("Table *" + table.replace('.', '\.'))
	re_table_begins = re.compile("^=[= |\t]*$")
	re_table_ends = re.compile("^ *\t*$")

	table_lines = []	# a list will accumulate lines from the table, 
						# to be written out in a second step
	
	try:
		with open(prn_file, 'r') as f:	# open the report file

			table_found = False
			data_out = False
			
			for line in f:				# loop through the lines of the file
				if not table_found:		# if the table title hasn't yet been found, test for it
					if re_table_title.search(line):
						table_found = True
				elif not data_out:		# if the table title has been found, but the data start marker
					if re_table_begins.search(line):	# hasn't yet, test for that
						data_out = True
				else:	# the table has been found, as well as the start of the data section
					if re_table_ends.search(line):		# test for a blank line marks the end of the data section
						break
					else:				# as long as the data section end not reached, accumulate data rows
						table_lines.append(line)
	except:
		raise
		print("There was a problem encountered with the input file, " + prn_file)

	try:
		if table_lines:	# if any data rows were accumulated, write them out to a new file
			with open(out_file, 'w') as f_out:
				for line in table_lines:
					f_out.write(line)
	except:
		print("There was a problem writing an output file, " + out_file)

# Main driver routine - this function does NOT launch a GUI.
def main(prn_file, output_dir, tables):
    
	for table in tables:
		extract_table_to_file(prn_file, table, os.path.join(output_dir, table.replace('.','-')) + ".txt")
		
# end_def main()
