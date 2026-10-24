"""Spreadsheet tests. Run from the project folder:  python -m unittest discover -s tests -v
Evaluates the workbook's real Excel formulas with pycel (if installed) and checks them against
the PRD section 11 numbers and against the independent Python model in build_spreadsheet.py.
"""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import build_spreadsheet as bs  # noqa: E402

XLSX = os.path.join(ROOT, "Stem_Cell_Microeconomics.xlsx")

import ast  # noqa: E402

if not hasattr(ast, "Str"):  # pycel uses ast.Str, removed in Python 3.14
    ast.Str = lambda s: ast.Constant(value=s)

try:
    from pycel import ExcelCompiler
except ImportError:  # pragma: no cover
    ExcelCompiler = None

PRD_EXAMPLE = dict(U=1e6, N0=1e6, g_E=20, r_g=2, q_E=0.5, r_q=1, s=0.8, p_E=1, p_S=1,
                   E_cells=0, E_consumables=0, E_labor=0, E_facility=0, E_equipment=0, E_qc=0, E_other=100000, E_delivery=0,
                   S_cells=0, S_consumables=0, S_labor=0, S_facility=0, S_equipment=0, S_qc=0, S_other=200000,
                   L=2000, M=50, C_return=50000, C_other_del=0)


class PythonModel(unittest.TestCase):
    def test_prd_section11(self):
        r = bs.model(PRD_EXAMPLE)
        self.assertAlmostEqual(r["Q_E"], 10)
        self.assertAlmostEqual(r["Q_S_pre"], 20)
        self.assertAlmostEqual(r["Q_S"], 16)
        self.assertAlmostEqual(r["c_E"], 10000)
        self.assertAlmostEqual(r["c_S_prod"], 10000)
        self.assertAlmostEqual(r["c_S"], 21875)
        self.assertAlmostEqual(r["savings"], -1.1875)
        self.assertAlmostEqual(r["r_g_min"], 4.375)
        self.assertAlmostEqual(r["L_max"], -1800)

    def test_prd_rg5(self):
        r = bs.model(dict(PRD_EXAMPLE, r_g=5))
        self.assertAlmostEqual(r["c_S"], 8750)
        self.assertAlmostEqual(r["savings"], 0.125)
        self.assertAlmostEqual(r["L_max"], 3000)

    def test_triangular_tail(self):
        self.assertAlmostEqual(bs.tri_exceed(4.375, 1, 2, 6), 0.13203125)

    def test_literature_ranges_valid(self):
        for pre, sc in bs.SCEN.items():
            for p in sc["P"]:
                self.assertLessEqual(p["low"], p["mode"], pre + " " + p["id"])
                self.assertLessEqual(p["mode"], p["high"], pre + " " + p["id"])
            q_max = sc["by"]["q_E"]["high"] * sc["by"]["r_q"]["high"]
            self.assertLessEqual(q_max, 1.0 + 1e-12, pre + ": q_S could exceed 1")

    def test_every_cited_source_exists(self):
        ids = {s[0] for s in bs.SOURCES}
        for pre, sc in bs.SCEN.items():
            for p in sc["P"]:
                for sid in filter(None, (x.strip() for x in p["sources"].split(","))):
                    self.assertIn(sid, ids, pre + " " + p["id"])
            for claim in sc["claims"]:
                for field in (claim[3], claim[4]):
                    for sid in (t.split(" ")[0].strip("(),") for t in field.split(",") if t.strip()):
                        if sid[:1] in "CS" and sid[1:2].isdigit():
                            self.assertIn(sid, ids, pre + " claim " + claim[0])

    def test_viability_thresholds_prd(self):
        r = bs.model(PRD_EXAMPLE)
        self.assertAlmostEqual(r["r_q_min"], 2.1875)
        self.assertAlmostEqual(r["r_q_cap"], 2)
        self.assertAlmostEqual(r["C_return_max"], -140000)
        self.assertAlmostEqual(r["C_S_prod_max"], 10000)
        r5 = bs.model(dict(PRD_EXAMPLE, r_g=5))
        self.assertAlmostEqual(r5["C_return_max"], 100000)
        self.assertAlmostEqual(r5["C_S_prod_max"], 250000)


@unittest.skipIf(ExcelCompiler is None, "pycel not installed")
class WorkbookFormulas(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        bs.build_workbook(XLSX)
        cls.xl = ExcelCompiler(filename=XLSX)

    def find(self, sheet, label):
        import openpyxl
        ws = openpyxl.load_workbook(XLSX)[sheet]
        for row in ws.iter_rows(min_col=1, max_col=1):
            if row[0].value and str(row[0].value).startswith(label):
                return self.xl.evaluate(f"{sheet}!B{row[0].row}")
        raise KeyError(label)

    def test_prd_check_block(self):
        f = lambda lab: self.find("PRD_Check", lab)
        self.assertAlmostEqual(f("Earth output Y_E"), 10)
        self.assertAlmostEqual(f("Space delivered output"), 16)
        self.assertAlmostEqual(f("Earth delivered cost (expect"), 10000)
        self.assertAlmostEqual(f("Space production cost (expect"), 10000)
        self.assertAlmostEqual(f("Space delivered cost (expect"), 21875)
        self.assertAlmostEqual(f("Savings (expect"), -1.1875)
        self.assertAlmostEqual(f("r_g_min (expect"), 4.375)
        self.assertAlmostEqual(f("L_max (expect"), -1800)
        self.assertAlmostEqual(f("P(r_g > 4.375)"), 0.13203125)

    def test_workbook_matches_python_model(self):
        for pre, sc in bs.SCEN.items():
            with self.subTest(scenario=pre):
                exp = bs.model(bs.point_values(sc["P"]))
                cmp_ = pre + "_Comparison"
                self.assertAlmostEqual(self.find(pre + "_Earth", "c_E: delivered"), exp["c_E"], places=6)
                self.assertAlmostEqual(self.find(pre + "_Space", "c_S: delivered"), exp["c_S"], places=6)
                self.assertAlmostEqual(self.find(pre + "_Space", "c_S_prod"), exp["c_S_prod"], places=6)
                self.assertAlmostEqual(self.find(cmp_, "Max launch price"), exp["L_max"], places=6)
                self.assertAlmostEqual(self.find(cmp_, "Max return cost"), exp["C_return_max"], places=4)
                self.assertAlmostEqual(self.find(cmp_, "Max orbital manufacturing"), exp["C_S_prod_max"], places=4)
                self.assertAlmostEqual(self.find(cmp_, "Break-even growth ratio"), exp["r_g_min"], places=6)
                self.assertAlmostEqual(self.find(cmp_, "Break-even purity ratio"), exp["r_q_min"], places=6)
                self.assertAlmostEqual(self.find(cmp_, "Maximum possible purity"), exp["r_q_cap"], places=9)
                self.assertAlmostEqual(self.find(cmp_, "Relative delivered savings"), exp["savings"], places=9)


if __name__ == "__main__":
    unittest.main()
