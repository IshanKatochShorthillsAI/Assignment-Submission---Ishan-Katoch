import unittest


class DetailedTestCase(unittest.TestCase):
    """
    Base test case class that allows a test method to record extra detail
    (such as expected vs. actual values) for inclusion in the final report.
    """

    def setUp(self):
        self._detail = ""

    def record_detail(self, msg):
        """
        Record a detail string for the current test.
        """
        self._detail = msg

    def get_detail(self):
        """
        Retrieve the recorded detail.
        """
        return self._detail
