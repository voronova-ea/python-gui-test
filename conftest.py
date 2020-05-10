import pytest
import os.path
from comtypes.client import CreateObject
from fixture.application import Application
from model.group import Group


@pytest.fixture(scope="session")
def app(request):
    fixture = Application("C:\\Study\\Python\\AddressBook\\AddressBook.exe")
    request.addfinalizer(fixture.destroy)
    return fixture


def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("excel_"):
            testdata = load_from_excel(fixture[6:])
            metafunc.parametrize(fixture, testdata, ids=[str(x) for x in testdata])


def load_from_excel(file):
    testdata = []
    xl = CreateObject("Excel.Application")
    xl.Visible = 1
    wb = xl.Workbooks.Open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "data/%s.xlsx" % file))
    worksheet = wb.Sheets[1]
    if worksheet.Cells[1, 1].Value() == "Groups":
        testdata = load_groups(worksheet)
    xl.Quit()
    return testdata


def load_groups(worksheet):
    data = []
    for row in range(2, worksheet.Rows.Count):
        val = worksheet.Cells[row, 1].Value()
        if val is None:
            break
        data.append(Group(name=val))
    return data