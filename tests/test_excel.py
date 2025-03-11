import unittest
import os
import sys
from openpyxl import Workbook

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "tests"))


class ExcelTestResult(unittest.TextTestResult):
    def __init__(self, stream, descriptions, verbosity):
        super().__init__(stream, descriptions, verbosity)
        self.results_list = []

    def addSuccess(self, test):
        super().addSuccess(test)
        self.results_list.append((str(test), "PASS", ""))

    def addFailure(self, test, err):
        super().addFailure(test, err)
        self.results_list.append(
            (str(test), "FAIL", self._exc_info_to_string(err, test))
        )

    def addError(self, test, err):
        super().addError(test, err)
        self.results_list.append(
            (str(test), "ERROR", self._exc_info_to_string(err, test))
        )

    def addSkip(self, test, reason):
        super().addSkip(test, reason)
        self.results_list.append((str(test), "SKIPPED", reason))


def run_tests_and_export():
    loader = unittest.TestLoader()
    suite = loader.discover(start_dir="tests")

    runner = unittest.TextTestRunner(resultclass=ExcelTestResult, verbosity=2)
    result = runner.run(suite)

    wb = Workbook()
    ws = wb.active
    ws.title = "Test Results"
    ws.append(["Test Case", "Status", "Details"])

    for test_case, status, details in result.results_list:
        ws.append([test_case, status, details])

    output_filename = "test_results.xlsx"
    wb.save(output_filename)
    print(f"Test results exported to {output_filename}")


if __name__ == "__main__":
    run_tests_and_export()
