import io
import os
import tempfile
import unittest
from unittest.mock import patch

import app as app_module


class AppSafetyTests(unittest.TestCase):
    def setUp(self):
        app_module.app.config.update(TESTING=True)
        self.client = app_module.app.test_client()

    def test_public_pages_use_generic_product_name(self):
        response = self.client.get("/")
        text = response.get_data(as_text=True)
        self.assertEqual(response.status_code, 200)
        self.assertIn("文检", text)
        self.assertNotIn("毕业设计格式检测工具", text)

    def test_profile_path_traversal_is_rejected(self):
        response = self.client.get("/editor/load/%2E%2E%2Fformat_rules.json")
        self.assertIn(response.status_code, (400, 404))

    def test_too_many_files_are_rejected_before_saving(self):
        files = [(io.BytesIO(b"x"), f"sample-{index}.docx")
                 for index in range(app_module.MAX_FILES_PER_CHECK + 1)]
        response = self.client.post(
            "/check",
            data={"files": files, "profile_id": app_module.BUILTIN_PROFILE_ID},
            content_type="multipart/form-data",
        )
        self.assertEqual(response.status_code, 302)

    def test_invalid_document_upload_is_cleaned(self):
        with tempfile.TemporaryDirectory() as upload_dir:
            with patch.object(app_module, "UPLOAD_DIR", upload_dir):
                response = self.client.post(
                    "/check",
                    data={
                        "files": (io.BytesIO(b"not-a-docx"), "sample.docx"),
                        "profile_id": app_module.BUILTIN_PROFILE_ID,
                    },
                    content_type="multipart/form-data",
                )
            self.assertEqual(response.status_code, 200)
            self.assertEqual(os.listdir(upload_dir), [])

    def test_invalid_spec_upload_is_cleaned(self):
        with tempfile.TemporaryDirectory() as spec_dir:
            with patch.object(app_module, "SPEC_DIR", spec_dir):
                response = self.client.post(
                    "/upload_spec",
                    data={"spec_file": (io.BytesIO(b"not-a-docx"), "rules.docx")},
                    content_type="multipart/form-data",
                )
            self.assertEqual(response.status_code, 302)
            self.assertEqual(os.listdir(spec_dir), [])

    def test_inferred_generic_checks_exclude_thesis_only_checks(self):
        checks = app_module.infer_enabled_checks([
            {"section": "body_text", "font": "宋体", "size_pt": 12},
        ])
        self.assertEqual(checks, ["font"])


if __name__ == "__main__":
    unittest.main()
