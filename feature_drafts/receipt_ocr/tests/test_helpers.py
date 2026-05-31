import unittest
from datetime import date

from feature_drafts.receipt_ocr import (
    build_expense_candidate,
    build_spendsight_receipt_payload,
    parse_receipt_ocr_text,
)


class ReceiptOcrHelperTests(unittest.TestCase):
    def test_indian_grocery_receipt_extracts_total_date_and_category(self):
        ocr_text = """
        TAX INVOICE
        DMART READY
        GSTIN: 27AAECA1234A1Z5
        Date: 31/05/2026 Time: 20:43
        Bill No: 778899
        Atta 5kg        275.00
        Rice            499.50
        Sub Total       1176.67
        CGST             28.92
        SGST             28.91
        Grand Total Rs. 1,234.50
        Paid by UPI
        """

        result = parse_receipt_ocr_text(ocr_text, ref_date=date(2026, 5, 31))

        self.assertEqual(result["merchant"], "DMART READY")
        self.assertEqual(result["date"], "2026-05-31")
        self.assertEqual(result["amount"], 1234.50)
        self.assertEqual(result["category"], "Groceries")
        self.assertEqual(result["subcategory"], "DMART READY")
        self.assertEqual(result["payment_method"], "UPI")
        self.assertEqual(result["currency_code"], "INR")
        self.assertGreaterEqual(result["confidence"]["overall"], 0.70)
        self.assertTrue(result["confidence_signals"]["has_labeled_total"])
        self.assertEqual(
            result["duplicate_key_candidates"][0]["key"],
            ["2026-05-31", 1234.50, "dmart ready"],
        )

    def test_indian_fuel_receipt_prefers_total_amount_over_tax_lines(self):
        ocr_text = """
        Retail Invoice
        INDIAN OIL COCO PUMP
        No. 14, Hosur Road
        Bill Date 29-05-2026
        Petrol 32.41 LTR  3,082.15
        CGST 83.92
        SGST 83.93
        TOTAL AMOUNT INR 3,250.00
        Card No XXXX XXXX 1234
        APPROVED
        """

        result = parse_receipt_ocr_text(ocr_text, ref_date=date(2026, 5, 31))

        self.assertEqual(result["merchant"], "INDIAN OIL COCO PUMP")
        self.assertEqual(result["date"], "2026-05-29")
        self.assertEqual(result["amount"], 3250.00)
        self.assertEqual(result["category"], "Fuel")
        self.assertEqual(result["payment_method"], "Card")
        self.assertNotEqual(result["amount"], 83.93)

    def test_generic_card_receipt_extracts_unambiguous_us_date(self):
        ocr_text = """
        CARDHOLDER COPY
        RIVER CAFE
        123 MAIN ST
        SALE
        DATE 05/29/2026 TIME 14:03
        VISA CREDIT
        AMOUNT USD 43.21
        APPROVAL 123456
        THANK YOU
        """

        result = parse_receipt_ocr_text(ocr_text, date_order="MDY", default_currency_code="USD")

        self.assertEqual(result["merchant"], "RIVER CAFE")
        self.assertEqual(result["date"], "2026-05-29")
        self.assertEqual(result["amount"], 43.21)
        self.assertEqual(result["category"], "Fast Food")
        self.assertEqual(result["payment_method"], "Card")
        self.assertEqual(result["currency_code"], "USD")
        self.assertGreaterEqual(result["confidence"]["amount"], 0.80)

    def test_duplicate_matches_score_existing_spendsight_expenses(self):
        data = {
            "currency_code": "INR",
            "expenses": [
                {
                    "id": "expense-1",
                    "date": "2026-05-31",
                    "amount": 1234.50,
                    "category": "Groceries",
                    "subcategory": "DMART READY",
                    "payment_method": "UPI",
                    "notes": "OCR receipt: DMART READY",
                },
                {
                    "id": "expense-2",
                    "date": "2026-05-20",
                    "amount": 1234.50,
                    "category": "Shopping",
                    "subcategory": "Unrelated",
                },
            ],
        }
        ocr_text = """
        DMART READY
        Date: 31/05/2026
        Grand Total Rs. 1,234.50
        Paid by UPI
        """

        result = parse_receipt_ocr_text(ocr_text, data=data, ref_date=date(2026, 5, 31))

        self.assertEqual(result["duplicate_matches"][0]["expense_id"], "expense-1")
        self.assertEqual(result["duplicate_matches"][0]["matched_on"], ["amount", "date", "merchant"])
        self.assertGreaterEqual(result["duplicate_matches"][0]["score"], 0.95)

    def test_payload_and_expense_candidate_use_spendsight_shapes(self):
        payload = build_spendsight_receipt_payload(
            "RIVER CAFE\nDATE 05/29/2026\nTOTAL USD 43.21",
            receipt_id="receipt-1",
            original_filename="river.jpg",
            stored_filename="receipt-1.jpg",
            content_type="image/jpeg",
            uploaded_at="2026-05-31T12:00:00",
            date_order="MDY",
            default_currency_code="USD",
        )
        expense = build_expense_candidate(
            payload["extracted"],
            expense_id="expense-1",
            now="2026-05-31T12:01:00",
        )

        self.assertEqual(payload["id"], "receipt-1")
        self.assertEqual(payload["status"], "needs_review")
        self.assertEqual(payload["extracted"]["merchant"], "RIVER CAFE")
        self.assertEqual(expense["id"], "expense-1")
        self.assertEqual(expense["date"], "2026-05-29")
        self.assertEqual(expense["amount"], 43.21)
        self.assertEqual(expense["source"], "receipt_ocr")
        self.assertEqual(expense["review_status"], "needs_review")


if __name__ == "__main__":
    unittest.main()
