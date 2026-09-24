import unittest

from regno import find_claim_numbers, find_reg_numbers, reg_or_claim


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


class ClaimFallbackTest(unittest.TestCase):
    def test_claim_numbers(self):
        cases = {
            "Request for Wreck offer<CL26217676>/<2383/84906808/00/000>/<M/S SHRI HARJIKA": ["CL26217676"],
            "Claim no. 10110425750 Regn. No. HR890648 Maruti": ["10110425750"],
            "DIESEL___WB-75-E-9412__Claim no- CL26246090___2315/81538921/00/000": ["CL26246090"],
            "WRECK VALUE / C1274101122507 / ASWANI JAISWAL": ["C1274101122507"],
            "CLAIM NO: 1234/2025/01 salvage": ["1234/2025/01"],
            "Claim intimation received": [],
            "Meeting at 10 AM": [],
        }
        for subject, expected in cases.items():
            with self.subTest(subject=subject):
                self.assertEqual(find_claim_numbers(subject), expected)

    def test_reg_preferred_over_claim(self):
        self.assertEqual(reg_or_claim("Claim no. 10110402407 Regn. No. KA06MA4060"), ("KA06MA4060", False))
        self.assertEqual(reg_or_claim("Request for Wreck offer<CL26071371>/<BOLERO PICK-UP>"), ("CL26071371", True))
        self.assertEqual(reg_or_claim("Weekly report"), ("", False))


if __name__ == "__main__":
    unittest.main()
