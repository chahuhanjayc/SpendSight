import unittest

from feature_drafts.receipts_inbox import (
    INBOX_KEY,
    STATUS_DUPLICATE_CANDIDATE,
    STATUS_POSTED,
    build_expense_from_inbox_item,
    build_review_inbox,
    create_receipt_inbox_item,
    empty_extracted_fields,
    empty_ocr_fields,
    find_duplicate_expenses,
    mark_item_posted,
    new_attachment_metadata,
)


class ReceiptInboxHelperTests(unittest.TestCase):
    def test_create_item_has_attachment_and_ocr_placeholders(self):
        data = {"expenses": []}
        attachment = new_attachment_metadata(
            "receipt.jpg",
            content_type="image/jpeg",
            size_bytes="2048",
            checksum_sha256="abc123",
            uploaded_at="2026-05-31T10:00:00",
        )

        item = create_receipt_inbox_item(data, attachment, item_id="receipt-1", now="2026-05-31T10:01:00")

        self.assertIn(INBOX_KEY, data)
        self.assertEqual(data[INBOX_KEY][0]["id"], "receipt-1")
        self.assertEqual(item["attachment"]["filename"], "receipt.jpg")
        self.assertEqual(item["attachment"]["size_bytes"], 2048)
        self.assertEqual(item["ocr"], empty_ocr_fields())
        self.assertEqual(item["extracted"], empty_extracted_fields())
        self.assertEqual(item["duplicate_candidates"], [])

    def test_duplicate_detection_matches_expense_by_amount_date_and_merchant(self):
        expenses = [
            {
                "id": "expense-1",
                "date": "2026-05-29",
                "amount": 1250.0,
                "category": "Groceries",
                "subcategory": "Big Bazaar",
                "payment_method": "Credit Card",
                "notes": "Big Bazaar Koramangala",
            },
            {
                "id": "expense-2",
                "date": "2026-05-20",
                "amount": 1250.0,
                "category": "Shopping",
                "subcategory": "Unrelated",
            },
        ]
        extracted = {
            "merchant": "Big Bazaar",
            "date": "2026-05-29",
            "amount": "1250.00",
            "payment_method": "Credit Card",
        }

        matches = find_duplicate_expenses(extracted, expenses)

        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0]["expense_id"], "expense-1")
        self.assertGreaterEqual(matches[0]["score"], 0.95)
        self.assertEqual(matches[0]["matched_on"], ["amount", "date", "merchant", "payment_method"])

    def test_create_item_marks_duplicate_candidates(self):
        data = {
            "expenses": [
                {
                    "id": "expense-1",
                    "date": "2026-05-31",
                    "amount": 499,
                    "subcategory": "Netflix",
                    "notes": "Netflix Premium",
                }
            ]
        }
        attachment = new_attachment_metadata("netflix.pdf", uploaded_at="2026-05-31T12:00:00")

        item = create_receipt_inbox_item(
            data,
            attachment,
            extracted={"merchant": "Netflix", "date": "2026-05-31", "amount": 499},
            now="2026-05-31T12:01:00",
        )

        self.assertEqual(item["status"], STATUS_DUPLICATE_CANDIDATE)
        self.assertEqual(item["duplicate_candidates"][0]["expense_id"], "expense-1")

    def test_build_expense_from_reviewed_item_uses_spendsight_expense_shape(self):
        data = {"expenses": []}
        item = create_receipt_inbox_item(
            data,
            new_attachment_metadata(
                "market.png",
                content_type="image/png",
                storage_path="user_data/receipts/market.png",
                uploaded_at="2026-05-31T12:00:00",
            ),
            extracted={
                "merchant": "Fresh Market",
                "date": "31/05/2026",
                "amount": "345.50",
                "category": "Groceries",
                "payment_method": "UPI",
                "notes": "Weekly produce",
                "quantity": "2",
                "unit": "kg",
            },
            item_id="receipt-2",
            now="2026-05-31T12:01:00",
        )

        expense = build_expense_from_inbox_item(
            item,
            expense_id="expense-2",
            now="2026-05-31T12:02:00",
        )

        self.assertEqual(expense["id"], "expense-2")
        self.assertEqual(expense["amount"], 345.5)
        self.assertEqual(expense["category"], "Groceries")
        self.assertEqual(expense["subcategory"], "Fresh Market")
        self.assertEqual(expense["date"], "2026-05-31")
        self.assertEqual(expense["payment_method"], "UPI")
        self.assertEqual(expense["quantity"], 2.0)
        self.assertEqual(expense["unit"], "kg")
        self.assertEqual(expense["source"], "receipt_inbox")
        self.assertEqual(expense["receipt_inbox_id"], "receipt-2")
        self.assertIn("Weekly produce", expense["notes"])
        self.assertIn("market.png", expense["notes"])

    def test_review_inbox_summary_and_post_marker(self):
        data = {"expenses": []}
        item = create_receipt_inbox_item(
            data,
            new_attachment_metadata("receipt.pdf", uploaded_at="2026-05-31T12:00:00"),
            extracted={"merchant": "Cafe", "date": "2026-05-31", "amount": 120},
            item_id="receipt-3",
            now="2026-05-31T12:01:00",
        )

        mark_item_posted(data, item["id"], "expense-3", reviewed_by="admin", now="2026-05-31T12:03:00")
        inbox = build_review_inbox(data)

        self.assertEqual(item["status"], STATUS_POSTED)
        self.assertEqual(item["review"]["expense_id"], "expense-3")
        self.assertEqual(item["review"]["reviewed_by"], "admin")
        self.assertEqual(inbox["summary"]["posted"], 1)
        self.assertEqual(inbox["summary"]["open"], 0)


if __name__ == "__main__":
    unittest.main()
