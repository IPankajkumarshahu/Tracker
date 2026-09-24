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
            "Claim no. 10110425750 Regn. No. HR890648 Maruti Celerio": ["HR890648"],
            "Registration No: DL 1 1234": ["DL011234"],
            "Wreck offers--DIESEL___WB-75-E-9412__Claim no- CL26246090": ["WB75E9412"],
            "BID REQUIRED--C1274103110324--YASH DAMMANI--CG 04 MJ 9331": ["CG04MJ9331"],
        }
        for subject, expected in cases.items():
            with self.subTest(subject=subject):
                self.assertEqual(find_reg_numbers(subject), expected)

    def test_no_match(self):
        for subject in ["Meeting at 10 AM", "Invoice INV2024 1234", "RE 12 AB 1234", "Claim no 1234",
                        "Request for Wreck offer<CL26217676>/<2383/84906808/00/000>/<M/S SHRI HARJIKA",
                        "<BOLERO PICK-UP FB PS 1.7 T XL>/<>", "", None]:
            with self.subTest(subject=subject):
                self.assertEqual(find_reg_numbers(subject), [])


if __name__ == "__main__":
    unittest.main()
