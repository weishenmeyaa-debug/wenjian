# -*- coding: utf-8 -*-
"""模板管理回归测试：主模板、改名、上传入口。

这些用例把模板目录和设置文件指到临时目录，不会碰到真实的模板。
"""

import json
import os
import shutil
import tempfile
import unittest

import app as app_module


class TemplateManagementTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="guijian_tpl_")
        self._orig_spec_dir = app_module.SPEC_DIR
        self._orig_settings = app_module.SETTINGS_PATH
        app_module.SPEC_DIR = os.path.join(self.tmp, "specs")
        app_module.SETTINGS_PATH = os.path.join(self.tmp, "settings.json")
        os.makedirs(app_module.SPEC_DIR)

        app_module.app.config.update(TESTING=True)
        self.client = app_module.app.test_client()

    def tearDown(self):
        app_module.SPEC_DIR = self._orig_spec_dir
        app_module.SETTINGS_PATH = self._orig_settings
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _make_template(self, name, rules=2):
        filename = "profile_%s.json" % abs(hash(name))
        path = os.path.join(app_module.SPEC_DIR, filename)
        payload = {
            "profile": {"name": name},
            "rules": [{"section": "s%d" % i, "label": "规则%d" % i} for i in range(rules)],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False)
        return filename, path

    def _names(self):
        return [p["name"] for p in app_module.list_profiles()]

    def _default_name(self):
        profile = app_module.get_profile()
        return profile["name"] if profile else None

    # --- 主模板 ---

    def test_first_own_template_becomes_default(self):
        """第一个自建模板应当接管默认值，不再停在示例上。"""
        self.assertEqual(self._default_name(), "示例模板")
        pid, _ = self._make_template("毕业论文格式")
        app_module.sync_default_profile(pid)
        self.assertEqual(self._default_name(), "毕业论文格式")

    def test_second_template_does_not_steal_default(self):
        """再建模板不应抢走已经确定的主模板。"""
        pid_a, _ = self._make_template("甲模板")
        app_module.sync_default_profile(pid_a)
        pid_b, _ = self._make_template("乙模板")
        app_module.sync_default_profile(pid_b)
        self.assertEqual(self._default_name(), "甲模板")

    def test_set_default_endpoint(self):
        pid_a, _ = self._make_template("甲模板")
        pid_b, _ = self._make_template("乙模板")
        app_module.sync_default_profile(pid_a)

        response = self.client.post("/editor/set-default", json={"profile_id": pid_b})
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.get_json()["ok"])
        self.assertEqual(self._default_name(), "乙模板")

    def test_default_moves_when_its_template_is_deleted(self):
        pid_a, _ = self._make_template("甲模板")
        pid_b, _ = self._make_template("乙模板")
        app_module.sync_default_profile(pid_a)

        self.client.post("/editor/delete/" + pid_a)
        self.assertEqual(self._default_name(), "乙模板")

    def test_default_falls_back_to_sample_when_all_own_deleted(self):
        pid, _ = self._make_template("仅有的模板")
        app_module.sync_default_profile(pid)
        self.client.post("/editor/delete/" + pid)
        self.assertEqual(self._default_name(), "示例模板")

    def test_default_marker_and_ordering(self):
        pid, _ = self._make_template("我的主模板")
        app_module.sync_default_profile(pid)
        profiles = app_module.list_profiles()
        self.assertTrue(profiles[0]["is_default"])
        self.assertEqual(profiles[0]["name"], "我的主模板")
        self.assertFalse(any(p["is_default"] for p in profiles[1:]))

    # --- 改名 ---

    def test_rename_keeps_rules(self):
        pid, path = self._make_template("旧名字", rules=3)
        response = self.client.post("/editor/rename", json={"profile_id": pid, "name": "新名字"})
        self.assertEqual(response.status_code, 200)

        with open(path, encoding="utf-8") as f:
            after = json.load(f)
        self.assertEqual(after["profile"]["name"], "新名字")
        self.assertEqual(len(after["rules"]), 3)

    def test_rename_rejects_duplicate_name(self):
        self._make_template("已占用")
        pid, _ = self._make_template("可用")
        response = self.client.post("/editor/rename", json={"profile_id": pid, "name": "已占用"})
        self.assertEqual(response.status_code, 400)
        self.assertFalse(response.get_json()["ok"])

    def test_rename_rejects_blank_name(self):
        pid, _ = self._make_template("原名")
        response = self.client.post("/editor/rename", json={"profile_id": pid, "name": "   "})
        self.assertEqual(response.status_code, 400)

    def test_rename_truncates_long_name(self):
        pid, _ = self._make_template("原名")
        response = self.client.post("/editor/rename", json={"profile_id": pid, "name": "名" * 200})
        self.assertEqual(response.status_code, 200)
        self.assertLessEqual(len(response.get_json()["name"]), 60)

    def test_rename_unknown_template_returns_404(self):
        for bad in ("..", "不存在.json", ""):
            response = self.client.post("/editor/rename", json={"profile_id": bad, "name": "x"})
            self.assertEqual(response.status_code, 404, bad)

    def test_sample_template_can_be_renamed(self):
        response = self.client.post(
            "/editor/rename", json={"profile_id": app_module.BUILTIN_PROFILE_ID, "name": "临时名"}
        )
        self.assertEqual(response.status_code, 200)
        self.assertIn("临时名", self._names())
        # 还原，避免影响其他用例
        self.client.post(
            "/editor/rename", json={"profile_id": app_module.BUILTIN_PROFILE_ID, "name": "示例模板"}
        )

    def test_sample_template_cannot_be_deleted(self):
        response = self.client.post("/editor/delete/" + app_module.BUILTIN_PROFILE_ID)
        self.assertEqual(response.status_code, 400)


class TemplatePageTests(unittest.TestCase):
    """模板页必须提供「复制提示词 → 粘贴 AI 结果」这条新建路径。"""

    def setUp(self):
        app_module.app.config.update(TESTING=True)
        self.client = app_module.app.test_client()

    def test_editor_has_ai_import_flow(self):
        html = self.client.get("/editor").get_data(as_text=True)
        self.assertIn('id="promptBox"', html)
        self.assertIn("copyPrompt", html)
        self.assertIn('id="jsonInput"', html)
        self.assertIn("parseJson", html)

    def test_editor_no_longer_parses_spec_files(self):
        html = self.client.get("/editor").get_data(as_text=True)
        self.assertNotIn("/upload_spec", html)
        self.assertNotIn('name="spec_file"', html)

    def test_editor_has_rename_and_default_controls(self):
        html = self.client.get("/editor").get_data(as_text=True)
        self.assertIn("showRenameDialog", html)
        self.assertIn("/editor/set-default", html)

    def test_index_shows_default_marker(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("★", html)


class ReferencePageTests(unittest.TestCase):
    """格式参考页要跟着当前模板走，不能拿写死的示例冒充模板规则。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="guijian_ref_")
        self._orig_spec_dir = app_module.SPEC_DIR
        self._orig_settings = app_module.SETTINGS_PATH
        app_module.SPEC_DIR = os.path.join(self.tmp, "specs")
        app_module.SETTINGS_PATH = os.path.join(self.tmp, "settings.json")
        os.makedirs(app_module.SPEC_DIR)

        app_module.app.config.update(TESTING=True)
        self.client = app_module.app.test_client()

    def tearDown(self):
        app_module.SPEC_DIR = self._orig_spec_dir
        app_module.SETTINGS_PATH = self._orig_settings
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _make_template(self, name, rules):
        filename = "ref_%s.json" % abs(hash(name))
        path = os.path.join(app_module.SPEC_DIR, filename)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"profile": {"name": name}, "rules": rules}, f,
                      ensure_ascii=False)
        return filename

    def test_numbers_follow_selected_template(self):
        """换模板，页面上的字号和字体要跟着换。"""
        pid = self._make_template("大字段模板", [
            {"section": "body_text", "font": "黑体", "size_pt": 16, "size_name": "三号"},
        ])
        html = self.client.get("/examples?profile=%s" % pid).get_data(as_text=True)
        self.assertIn("16pt", html)
        self.assertIn("三号黑体", html)
        self.assertIn("SimHei", html)

    def test_missing_sections_are_declared_as_skipped(self):
        """模板里没有图表/参考文献规则时，页面必须说明会跳过，而不是照抄示例。"""
        pid = self._make_template("只有正文字体", [
            {"section": "body_text", "font": "宋体", "size_pt": 12, "size_name": "小四"},
        ])
        html = self.client.get("/examples?profile=%s" % pid).get_data(as_text=True)
        self.assertIn("跳过这一项", html)
        self.assertNotIn("不少于", html)

    def test_present_sections_are_declared_as_checked(self):
        pid = self._make_template("带参考文献", [
            {"section": "reference_count", "min_count": 15, "source": "测试"},
            {"section": "reference_entry", "font": "宋体", "size_pt": 10.5},
        ])
        html = self.client.get("/examples?profile=%s" % pid).get_data(as_text=True)
        self.assertIn("不少于 15 篇", html)
        self.assertIn("参考文献", html)

    def test_nav_keeps_current_template(self):
        """从哪套模板进来，导航里的「格式参考」就还带那套模板。"""
        pid = self._make_template("甲模板", [{"section": "body_text", "size_pt": 12}])
        html = self.client.get("/editor?profile=%s" % pid).get_data(as_text=True)
        self.assertIn("/examples?profile=%s" % pid, html)
        self.assertIn("/?profile=%s" % pid, html)

    def test_reference_sample_says_where_format_comes_from(self):
        """模板写了条目排版才算「来自模板」，只有篇数时标签要写「默认排版」。"""
        full = self._make_template("条目与篇数齐全", [
            {"section": "reference_entry", "font": "黑体", "size_pt": 12,
             "size_name": "小四", "line_spacing_pt": 20},
            {"section": "reference_count", "min_count": 30},
            {"section": "reference_foreign_count", "min_count": 4},
        ])
        html = self.client.get("/examples?profile=%s" % full).get_data(as_text=True)
        self.assertIn("条目排版来自模板", html)
        self.assertIn("格式仅供参考", html)
        self.assertIn(
            "当前模板参考文献数量要求：参考文献不少于 30 篇，其中外文不少于 4 篇", html)

        count_only = self._make_template("只有篇数", [
            {"section": "reference_count", "min_count": 18},
        ])
        html = self.client.get("/examples?profile=%s" % count_only).get_data(as_text=True)
        self.assertIn("默认排版", html)
        self.assertIn("当前模板参考文献数量要求：参考文献不少于 18 篇", html)
        # 模板没写外文篇数，就不能凭空报一个数
        self.assertNotIn("外文不少于", html)

        entry_only = self._make_template("只有条目排版", [
            {"section": "reference_entry", "font": "宋体", "size_pt": 10.5,
             "size_name": "五号", "line_spacing_pt": 18},
        ])
        html = self.client.get("/examples?profile=%s" % entry_only).get_data(as_text=True)
        # 没写篇数就说没规定，不许搬默认的 20 篇 / 3 篇
        self.assertIn("当前模板没规定参考文献数量", html)
        self.assertNotIn("不少于 20 篇", html)


if __name__ == "__main__":
    unittest.main()
