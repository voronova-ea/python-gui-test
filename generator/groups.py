from comtypes.client import CreateObject
import getopt
import string
import random
import sys
import re
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


def random_string(prefix, maxlen):
    symbols = string.ascii_letters + string.digits + string.punctuation + " " * 10
    return prefix + "".join([random.choice(symbols) for _ in range(random.randrange(maxlen))])

project_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))

xl = CreateObject("Excel.Application")
xl.Visible = 1
wb = xl.Workbooks.Add()
# tag for separating test data for groups and for contacts
xl.Range["A1"].Value[()] = "Groups"
for i in range(n):
    xl.Range["A%s" % (i + 2)].Value[()] = random_string("group", 10)
wb.SaveAs(os.path.join(project_dir, f))
xl.Quit()
