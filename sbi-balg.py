#! /usr/bin/env python

# Nikhil Vyas

# SBI's output xls has some weird compatibility issues with xlrd
# Convert to xlsx. Why? Because I though using openpyxl was cooler.
# Plus you shouldn't live in 2003. Just an opinion.
# Could've done in csv, I know. But wanted to experiment with xlsx.

import os
import sys
import subprocess
from openpyxl import load_workbook;

# I'm a joke.

gnuplot_sc="""set title "Account Balance vs Time"
set xdata time
set xlabel "Time --->"
set terminal png size 800,600
set output "bal_v_time.png"
set ylabel "Money -->"
set grid back ls 12
set timefmt "%d %b %Y"
set xrange['{0}':'{1}']
set format x "%b %Y"
set style data lines
plot "temp.dat" using 1:4 t "Account Balance"
"""


def main():
	if len(sys.argv) < 2:
		print("Usage: ./sbi-balg.py <input_xlsx>");
		exit();
	wb = load_workbook(sys.argv[1]);
	sheet = wb.active;
	trans_data = 0;
	dates = []
	f = open('temp.dat', 'w');
	for row in sheet.rows:

		if row[0].value == 'Txn Date':
			trans_data = 1;
			continue;
		if trans_data == 0:
			continue;

		idx = -1;
		found = 0;
		while found == 0:  #SBI Generated files have some empty columns
			if row[idx].value != None:
				found = 1;
			idx = idx - 1;

		if (row[idx].value != None) & (row[idx].value != ' '):
			bal =row[idx].value*1000 + row[idx+1].value;
		else:
			bal = row[idx+1].value;
		f.write("{0} {1}\n".format(row[0].value, bal));
		dates.append(row[0].value);
	
	f.close();

	# Gnuplot Script

	f = open('temp.gp', 'w');
	f.write(gnuplot_sc.format(dates[0], dates[-1]));
	f.close();
	

	# Graphing Time!
	
	p=subprocess.Popen("gnuplot temp.gp", shell=True);
	os.waitpid(p.pid, 0);
	
	# Cleanup. I wished I did this with my room.
	os.remove("temp.gp");
	os.remove("temp.dat");

if __name__ == "__main__":
	main();
