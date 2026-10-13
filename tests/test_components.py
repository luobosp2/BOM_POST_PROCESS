import unittest

from bom_post_process.components import (
    component_kind, component_name, is_nc, normalize_model, parse_capacitor, parse_resistor,
)
from bom_post_process.errors import BomError


class ComponentTests(unittest.TestCase):
    def test_capacitor_units_tolerance_and_material_aliases(self):
        cases = {
            "0402-X5R104K250": "100nF/25V/0402/X5R/10%",
            "0603-X5R476M6R3": "47uF/6.3V/0603/X5R/20%",
            "0402-C0G4R7J500": "4.7pF/50V/0402/C0G/5%",
            "0402-COG0R5K250": "0.5pF/25V/0402/C0G/10%",
            "0402-NP0820K100": "82pF/10V/0402/C0G/10%",
            "0603-X7R105K250": "1uF/25V/0603/X7R/10%",
            " 0402-x5r102k250/nc ": "1nF/25V/0402/X5R/10%",
            "0402-X5R102K250(NC)": "1nF/25V/0402/X5R/10%",
        }
        for model, expected in cases.items():
            with self.subTest(model=model):
                self.assertEqual(parse_capacitor(model), expected)

    def test_resistor_decimal_and_legacy_one_percent_convention(self):
        cases = {
            "0402K0000": "0/0402/1%", "0402K0R": "0/0402/1%",
            "0402K0R10": "0.1/0402/1%", "0402K240R": "240/0402/1%",
            "0402K4R70": "4.7/0402/1%", "0402K33R2": "33.2/0402/1%",
            "0402K1782": "17.8k/0402/1%", "1206K1004": "1M/1206/1%",
            "0603L3010": "301/0603/1%", "0805U1001": "1k/0805/1%",
            "1206M1002": "10k/1206/1%", " 0402k1003(nc) ": "100k/0402/1%",
        }
        for model, expected in cases.items():
            with self.subTest(model=model):
                self.assertEqual(parse_resistor(model), expected)

    def test_unknown_encodings_are_never_guessed(self):
        cases = {
            parse_capacitor: [None, 104, "", "1206B104K101NT(FH)", "T491X476K035AT",
                              "0402-ABC104K250", "0402-X5R104Q250", "0402-X5R104K25"],
            parse_resistor: [None, 2400, "", "0402J2400", "0402K240", "0402K4RR7",
                             "0402K1003/NC/NC", "M1206"],
        }
        for parser, models in cases.items():
            for model in models:
                with self.subTest(parser=parser.__name__, model=model):
                    with self.assertRaises(BomError):
                        parser(model)

    def test_nc_suffix_and_component_classification(self):
        self.assertTrue(is_nc(" 0402K1003(nc) "))
        self.assertFalse(is_nc("IC/NC_ALT"))
        self.assertEqual(normalize_model("0402k1003/nc"), "0402K1003")
        self.assertEqual(component_kind("1206B104K101NT(FH)", "Cap"), "capacitor")
        self.assertIsNone(component_kind("M1206", "M1206"))
        self.assertIsNone(component_kind("SDVL3216SD300PTHS501", "Res Varistor"))
        self.assertEqual(component_name("SDVL3216SD300PTHS501", "Res Varistor"), "压敏电阻")
        self.assertEqual(component_name("0402K1003", "Res1"), "贴片电阻")


if __name__ == "__main__":
    unittest.main()
