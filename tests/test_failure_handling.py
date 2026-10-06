import subprocess
import tempfile
import threading
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from Service.Service import Service


class FailureHandlingTests(unittest.TestCase):
    def setUp(self):
        self.service = Service(Mock(), Mock(), Mock())
        self.service.ensure_excel_exists = Mock(return_value=True)
        self.service.ensure_valid_ip = Mock(return_value=True)
        self.service.show_message = Mock()
        self.service.query_esu_license = Mock(return_value={"LicenseStatus": 0, "PartialProductKey": self.service.cdkey[-5:]})
        self.service.run_command = Mock(return_value="Installation ID: " + "1" * 63)
        self.service.save_to_excel = Mock(return_value=True)
        self.service.save_license_status_to_excel = Mock(return_value=True)

    def test_invalid_dti_is_never_saved(self):
        for output in ("오류: 설치 ID 조회 실패", "Error: 0xC004F050", "123", "Installation ID: invalid"):
            self.service.run_command.return_value = output
            self.service.check_installation_id()
        self.service.save_to_excel.assert_not_called()

    def test_valid_dti_and_leading_zero(self):
        for size in (54, 63):
            digits = "0" + "1" * (size - 1)
            self.assertEqual(self.service.extract_dti_value("설치 ID: " + digits), digits)
        self.assertEqual(self.service.extract_dti_value("Installation ID: " + "-".join(["0123456"] * 9)), "0123456" * 9)

    def test_key_install_failure_prevents_dti(self):
        self.service.run_command.side_effect = RuntimeError("설치 실패")
        self.service.check_installation_id()
        self.assertEqual(self.service.run_command.call_count, 1)
        self.service.save_to_excel.assert_not_called()

    def test_wrong_product_key_prevents_dti(self):
        self.service.query_esu_license.return_value["PartialProductKey"] = "OTHER"
        self.service.check_installation_id()
        self.assertEqual(self.service.run_command.call_count, 1)
        self.service.save_to_excel.assert_not_called()

    def test_excel_failure_prevents_all_commands(self):
        self.service.ensure_excel_exists.return_value = False
        self.service.check_installation_id()
        self.service.activate_esu()
        self.service.run_command.assert_not_called()
        self.service.query_esu_license.assert_not_called()

    def test_activation_success_uses_numeric_status(self):
        self.service.get_confirm_value_from_excel = Mock(return_value="0" + "1" * 47)
        self.service.query_esu_license.side_effect = [{"LicenseStatus": 0}, {"LicenseStatus": 1}]
        self.service.run_command.return_value = "Confirmation ID deposited successfully."
        self.service.activate_esu()
        self.service.save_license_status_to_excel.assert_called_once_with("사용 허가됨")

    def test_activation_failure_does_not_record_success(self):
        self.service.get_confirm_value_from_excel = Mock(return_value="1" * 48)
        self.service.run_command.side_effect = subprocess.TimeoutExpired("cscript", 60)
        self.service.activate_esu()
        self.service.save_license_status_to_excel.assert_not_called()

    @patch("Service.Service.subprocess.run")
    def test_command_checks_return_code_and_slmgr_error(self, run):
        for code, output in ((1, "failed"), (0, "Error: 0xC004F050")):
            run.return_value = SimpleNamespace(returncode=code, stdout=output, stderr="")
            with self.assertRaises(RuntimeError):
                Service.run_command(self.service, ["cscript.exe"])
        self.assertFalse(run.call_args.kwargs["shell"])
        self.assertEqual(run.call_args.kwargs["timeout"], 60)

    def test_workbook_initialization_and_failure(self):
        with tempfile.TemporaryDirectory() as folder:
            self.service.excel_path = str(Path(folder) / "esu.xlsx")
            self.assertTrue(Service.ensure_excel_exists(self.service))
            Path(self.service.excel_path).write_text("damaged", encoding="utf-8")
            self.assertFalse(Service.ensure_excel_exists(self.service))

    @patch("Service.Service.threading.Thread")
    def test_duplicate_task_is_ignored(self, thread):
        task = Mock()
        self.service.start_task(task)
        self.service.start_task(task)
        thread.assert_called_once()
        self.service.ui.btn_auth.config.assert_called_once_with(state="disabled")
        thread.call_args.kwargs["target"]()
        self.service.poll_ui()
        self.assertFalse(self.service.busy)
        self.service.ui.btn_auth.config.assert_called_with(state="normal")

    def test_worker_ui_callback_runs_on_main_thread(self):
        observed = []
        worker = threading.Thread(target=lambda: self.service.call_ui(lambda: observed.append(threading.get_ident())))
        worker.start()
        callback, args, kwargs, future = self.service.ui_queue.get(timeout=2)
        future.set_result(callback(*args, **kwargs))
        worker.join(timeout=2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(observed, [threading.get_ident()])


if __name__ == "__main__":
    unittest.main()
