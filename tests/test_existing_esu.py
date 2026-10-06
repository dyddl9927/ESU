import json
import subprocess
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

from Service.Service import Service


class ExistingEsuTests(unittest.TestCase):
    def setUp(self):
        self.service = Service(Mock(), Mock(), Mock())
        self.service.ensure_valid_ip = Mock(return_value=True)
        self.service.ensure_excel_exists = Mock(return_value=True)
        self.service.save_license_status_to_excel = Mock(return_value=True)
        self.service.save_to_excel = Mock(return_value=True)
        self.service.get_confirm_value_from_excel = Mock(return_value="1" * 48)
        original_run = self.service.run_command
        self.service.run_command = Mock(side_effect=lambda command, **kwargs: original_run(command, **kwargs) if command[0] == "powershell.exe" else "1" * 63)

    def query_result(self, status, returncode=0):
        output = json.dumps({"LicenseStatus": int(status), "PartialProductKey": self.service.cdkey[-5:]}) if status.strip() in {"0", "1", "2", "3", "4", "5", "6"} else "null" if status == "-1" else status
        return SimpleNamespace(stdout=output, stderr="" if returncode == 0 else "query failed", returncode=returncode)

    @patch("Service.Service.messagebox")
    @patch("Service.Service.subprocess.run")
    def test_licensed_pc_skips_install_and_activation_even_with_new_build(self, run, dialogs):
        run.return_value = self.query_result("1\n")
        self.service.util.check_windows_version.return_value = False
        self.service.check_installation_id()
        self.service.activate_esu()
        self.assertFalse(any("/ipk" in c.args[0] or "/atp" in c.args[0] for c in self.service.run_command.call_args_list))
        self.service.util.check_windows_version.assert_not_called()
        self.service.get_confirm_value_from_excel.assert_not_called()
        self.service.save_license_status_to_excel.assert_called_with("사용 허가됨")
        self.assertIn(self.service.activation_id, run.call_args.args[0][-1])

    @patch("Service.Service.messagebox")
    @patch("Service.Service.subprocess.run")
    def test_unlicensed_and_absent_esu_continue_installation(self, run, dialogs):
        for status in ("-1", "0", "2", "3", "4", "5", "6"):
            with self.subTest(status=status):
                run.side_effect = [self.query_result(status), self.query_result("0")]
                self.service.run_command.reset_mock()
                self.service.check_installation_id()
                commands = [call.args[0] for call in self.service.run_command.call_args_list if call.args[0][0] != "powershell.exe"]
                self.assertEqual(len(commands), 2)
                self.assertIn("/ipk", commands[0])
                self.assertIn("/dti", commands[1])

    @patch("Service.Service.messagebox")
    @patch("Service.Service.subprocess.run")
    def test_query_failure_never_installs_or_activates(self, run, dialogs):
        for result in (self.query_result("", 1), self.query_result("unexpected")):
            with self.subTest(result=result):
                run.return_value = result
                self.service.check_installation_id()
                self.service.activate_esu()
        run.side_effect = subprocess.TimeoutExpired("powershell.exe", 30)
        self.service.check_installation_id()
        self.assertFalse(any("/ipk" in c.args[0] or "/atp" in c.args[0] for c in self.service.run_command.call_args_list))
        self.service.get_confirm_value_from_excel.assert_not_called()

    @patch("Service.Service.messagebox")
    @patch("Service.Service.subprocess.run")
    def test_excel_failure_still_preserves_active_license(self, run, dialogs):
        run.return_value = self.query_result("1")
        self.service.save_license_status_to_excel.return_value = False
        self.service.check_installation_id()
        self.assertFalse(any("/ipk" in c.args[0] or "/atp" in c.args[0] for c in self.service.run_command.call_args_list))
        self.service.ui.status_label.config.assert_called_with(
            text="ESU 사용 중 - 상태 저장 실패", fg="orange",
        )


if __name__ == "__main__":
    unittest.main()
