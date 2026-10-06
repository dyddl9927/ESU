import unittest
from unittest.mock import Mock, patch

from esu_core import ESUAuthTool
from Util.Util import Util


class AppStartupTests(unittest.TestCase):
    @patch("Service.Service.threading.Thread")
    @patch("esu_core.U")
    @patch("esu_core.Ui")
    @patch("esu_core.tk.Tk")
    def test_startup_connects_to_real_service_and_initializes(self, tk, ui, util, thread):
        util.return_value.is_admin.return_value = True
        app = ESUAuthTool()
        app.service.ensure_excel_exists = Mock(return_value=True)
        app.service.ensure_valid_ip = Mock(return_value=True)
        app.service.query_esu_license = Mock(return_value={"LicenseStatus": 1})
        app.service.save_license_status_to_excel = Mock(return_value=True)
        app.service.show_message = Mock()
        app.service.run_command = Mock()
        # Execute the callback installed by the real application constructor.
        delay, callback = tk.return_value.after.call_args.args
        self.assertEqual(delay, 0)
        callback()
        self.assertTrue(app.service.busy)
        thread.call_args.kwargs["target"]()
        app.service.poll_ui()
        self.assertFalse(app.service.busy)
        app.service.run_command.assert_not_called()
        app.service.save_license_status_to_excel.assert_called_once_with("사용 허가됨")

    @patch("esu_core.S")
    @patch("esu_core.U")
    @patch("esu_core.Ui")
    @patch("esu_core.tk.Tk")
    def test_non_admin_never_starts_service(self, tk, ui, util, service):
        util.return_value.is_admin.return_value = False
        ESUAuthTool()
        util.return_value.request_admin.assert_called_once()
        service.assert_not_called()
        tk.return_value.after.assert_not_called()

    def test_only_requested_builds_are_allowed(self):
        for build, expected in (("19045.6456", True), ("19045.6466", True), ("19045.6467", False), ("19045.6455", False)):
            with self.subTest(build=build), patch.object(Util, "get_windows_build", return_value=build):
                self.assertEqual(Util.check_windows_version(), expected)


if __name__ == "__main__":
    unittest.main()
