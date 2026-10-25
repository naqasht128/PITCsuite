import unittest
from bs4 import BeautifulSoup
from pitcsuite.tools.bill_scraper_logic import (
    PITCBillScraper,
    compute_total_units,
    compute_total_payment,
)

SAMPLE_HTML = """
<!DOCTYPE html>
<html>
<body>
<div class="gbn-app-shell">
    <div class="grid-col-cell">
        <div class="label-row">
            <span class="en-lbl">STATUS</span>
            <span class="ur-lbl">&#x062D;&#x0627;&#x0644;&#x062A;</span>
        </div>
        <div class="val-space">ALIVE</div>
    </div>
    <div class="grid-col-cell">
        <div class="label-row">
            <span class="en-lbl">NAME & ADDRESS</span>
            <span class="ur-lbl">&#x0646;&#x0627;&#x0645; &#x0648; &#x067e;&#x062a;&#x06c1;</span>
        </div>
        <div class="val-space">Edoto Mobile Tower, Near National Club, D.I.Khan</div>
    </div>
    <div class="grid-col-cell">
        <div class="label-row">
            <span class="en-lbl">DUE DATE</span>
            <span class="ur-lbl">&#x0645;&#x0642;&#x0631;&#x0631;&#x06c1; &#x062a;&#x0627;&#x0631;&#x06cc;&#x062e;</span>
        </div>
        <div class="val-space">23 OCT 26</div>
    </div>

    <table>
        <thead>
            <tr>
                <th>METER NO</th>
                <th>PREVIOUS READING</th>
                <th>PRESENT READING</th>
                <th>MF</th>
                <th>UNITS</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>3798029575</td>
                <td>25246.07<br>6494.01</td>
                <td>26205.61<br>6754.78</td>
                <td>1<br>1</td>
                <td>960<br>261</td>
            </tr>
        </tbody>
    </table>
</div>
</body>
</html>
"""

class TestBillScraper(unittest.TestCase):
    def test_bill_scraper_extraction(self):
        soup = BeautifulSoup(SAMPLE_HTML, "html.parser")
        out = {"_account": "30269310004301", "Remarks": "Fetched"}
        PITCBillScraper.put_grid_col_cells(out, soup)
        PITCBillScraper.put_general_pairs(out, soup)
        PITCBillScraper.put_general_text_pairs(out, soup)
        PITCBillScraper.put_meter_tables(out, soup)
        PITCBillScraper.put_meter_text_sections(out, soup)
        PITCBillScraper.put_cum_mdi(out, soup)
        PITCBillScraper.put_charges_breakdown(out, soup)
        out["Total Units"] = compute_total_units(out)

        self.assertEqual(out.get("status"), "ALIVE")
        self.assertEqual(out.get("due_date"), "23 OCT 26")
        self.assertEqual(out.get("name_address"), "Edoto Mobile Tower, Near National Club, D.I.Khan")
        self.assertEqual(out.get("kwh_present_1"), "26205.61")
        self.assertEqual(out.get("kwh_present_2"), "6754.78")
        self.assertEqual(out.get("kwh_mf_1"), "1")
        self.assertEqual(out.get("kwh_mf_2"), "1")
        self.assertEqual(out.get("kwh_units_1"), "960")
        self.assertEqual(out.get("kwh_units_2"), "261")
        self.assertEqual(out.get("Total Units"), 1221)

if __name__ == "__main__":
    unittest.main()
