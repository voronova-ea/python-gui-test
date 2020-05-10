from comtypes.client import CreateObject
import getopt
import sys
import os

try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of groups", "file"])
except getopt.GetoptError as err:
    getopt.usage()
    sys.exit(2)

n = 10
f = "data/groups.xlsx"

for o, a in opts:
    if o == "-n":
        n = int(a)
    elif o == "-f":
        f = a

project_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))

xl = CreateObject("Excel.Application")
xl.Visible = 1
wb = xl.Workbooks.Add()
for i in range(n):
    xl.Range["A%s" % (i+1)].Value[()] = "group %s" % i
wb.SaveAs(os.path.join(project_dir, f))
xl.Quit()