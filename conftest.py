import os
import pytest
from fixture.application import Application
from comtypes.client import CreateObject


@pytest.fixture(scope="session")
def app(request):
    fixture = Application("D:\\FreeAddressBookPortable\\AddressBook.exe")
    request.addfinalizer(fixture.destroy)
    return fixture


def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("xlsx_"):
            testdata = load_from_excel(fixture[5:])
            metafunc.parametrize(fixture, testdata)


def load_from_excel(file):
    xlsx_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), f"fixture/{file}.xlsx")
    app = CreateObject("Excel.Application")
    wb = app.Workbooks.Open(xlsx_file)
    worksheet = wb.Sheets[1]
    test_data = []
    for row in range(1, 100):
        data = worksheet.Cells[row, 1].Value()
        if data is not None:
            test_data.append(data)
    return test_data
