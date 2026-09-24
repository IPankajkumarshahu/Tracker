import unittest

from regno import find_reg_numbers


class FindRegNumbersTest(unittest.TestCase):
    def test_formats(self):
        cases = {
            "Claim intimation - MH12AB1234": ["MH12AB1234"],
            "Vehicle MH 12 AB 1234 inspection": ["MH12AB1234"],
            "RE: dl-3c-ab-1234 survey": ["DL03CAB1234"],
            "Veh No.KA05MN123 accident": ["KA05MN0123"],
            "Bharat series 22 BH 1234 AA": ["22BH1234AA"],
            "Two cars: HR26DK8337 and UP 16 BT 4411": ["HR26DK8337", "UP16BT4411"],
            "Dup MH12AB1234 / MH-12-AB-1234": ["MH12AB1234"],
        }
        for subject, expected in cases.items():
            with self.subTest(subject=subject):
                self.assertEqual(find_reg_numbers(subject), expected)

    def test_no_match(self):
        for subject in ["Meeting at 10 AM", "Invoice INV2024 1234", "RE 12 AB 1234", "", None]:
            with self.subTest(subject=subject):
                self.assertEqual(find_reg_numbers(subject), [])


if __name__ == "__main__":
    unittest.main()
