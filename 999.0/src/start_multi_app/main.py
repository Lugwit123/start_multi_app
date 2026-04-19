import copy
import importlib
import json
import locale
import os
import subprocess
import sys
from functools import partial

try:
    _yaml = importlib.import_module("yaml")
except Exception:
    _yaml = None
from PyQt5.QtGui import QColor, QIcon, QPalette
from PyQt5.QtWidgets import (
    QApplication,
    QComboBox,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

os.environ["QT_API"] = "PyQt5"

LugwitToolDir = os.environ.get("LugwitToolDir", "")
if LugwitToolDir:
    sys.path.append(os.path.join(LugwitToolDir, "Lib"))

from Lugwit_Module.l_src import insLocation  # noqa: E402


ORI_ENV_VAR_FILE = os.getenv("oriEnvVarFile")
if ORI_ENV_VAR_FILE and os.path.exists(ORI_ENV_VAR_FILE):
    with open(ORI_ENV_VAR_FILE, "r", encoding="utf-8") as f:
        ORI_ENV_VAR = json.load(f)
else:
    ORI_ENV_VAR = {}

COMMON_KILL_PROGRAMS = [
    "Maya.exe",
    "UnrealEditor.exe",
    "houdini.exe",
    "3dsmax.exe",
    "deadlineworker.exe",
    "python.exe",
    "pythonw_p4v_embed_win.exe",
    "cursor.exe",
    "dingtalk.exe",
    "lugwit_chatroom.exe",
]


def _parse_scalar(value: str):
    text = value.strip()
    if text == "":
        return ""
    if text in ("true", "True"):
        return True
    if text in ("false", "False"):
        return False
    if text in ("null", "None"):
        return None
    if text.isdigit() or (text.startswith("-") and text[1:].isdigit()):
        try:
            return int(text)
        except Exception:
            pass
    if (text.startswith('"') and text.endswith('"')) or (text.startswith("'") and text.endswith("'")):
        return text[1:-1]
    if text in ("{}", "[]"):
        return json.loads(text)
    return text


def _load_config_text(path: str):
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    if _yaml is not None:
        return _yaml.safe_load(content)

    # Minimal YAML fallback for current config shape:
    # top-level key: value and key: [\n  - item]
    data = {}
    current_list_key = None
    for raw in content.splitlines():
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if line.startswith("  - ") and current_list_key:
            data.setdefault(current_list_key, []).append(_parse_scalar(line[4:]))
            continue
        current_list_key = None
        if ":" in line and not line.startswith("  "):
            key, value = line.split(":", 1)
            key = key.strip()
            value = value.strip()
            if value == "":
                data[key] = []
                current_list_key = key
            else:
                data[key] = _parse_scalar(value)
    return data


def _dump_config_text(config: dict) -> str:
    if _yaml is not None:
        return _yaml.safe_dump(config, allow_unicode=True, sort_keys=False)

    # Minimal YAML dump for simple dict/list/scalar config.
    lines = []
    for key, value in config.items():
        if isinstance(value, list):
            lines.append(f"{key}:")
            for item in value:
                lines.append(f"  - {item}")
        elif isinstance(value, dict):
            if not value:
                lines.append(f"{key}: {{}}")
            else:
                lines.append(f"{key}:")
                for k, v in value.items():
                    lines.append(f"  {k}: {v}")
        else:
            lines.append(f"{key}: {value}")
    return "\n".join(lines) + "\n"


class ProgramLauncher(QWidget):
    def __init__(self):
        super().__init__()
        self.config_file = os.path.join(os.path.dirname(__file__), "config", "kill_programs.yaml")
        self._kill_programs = []
        self._filtered_programs = []
        self._usage_counts = {}
        self.common_programs = []
        self.init_ui()
        self.load_settings()
        self.destroyed.connect(lambda *_: self._persist_settings())

    def init_ui(self):
        self.setWindowTitle("程序启动器")
        self.resize(900, 470)
        self._set_theme()

        page_layout = QVBoxLayout(self)
        page_layout.setContentsMargins(8, 8, 8, 8)
        page_layout.setSpacing(4)

        title = QLabel("程序启动器")
        title.setObjectName("pageTitle")
        subtitle = QLabel("批量启动、筛选排序与一键结束常见程序")
        subtitle.setObjectName("pageSubtitle")
        page_layout.addWidget(title)
        page_layout.addWidget(subtitle)

        root_layout = QHBoxLayout()
        root_layout.setSpacing(8)
        root_layout.addWidget(self._build_launch_panel(), 3)
        root_layout.addWidget(self._vline())
        root_layout.addWidget(self._build_kill_panel(), 2)
        page_layout.addLayout(root_layout)

    def _set_theme(self):
        self.setStyleSheet(
            """
            QWidget { background:#1f232a; color:#e6e9ef; font-family:Microsoft YaHei; font-size:12px; }
            QLabel#pageTitle { font-size:16px; font-weight:700; color:#f5f7fa; margin-bottom:0px; }
            QLabel#pageSubtitle { font-size:10px; color:#8b95a7; margin-bottom: 0px; }
            QFrame#panelCard {
                background:#2a2f39;
                border:1px solid #3a404c;
                border-radius:6px;
            }
            QComboBox, QLineEdit, QSpinBox {
                background:#313743; border:1px solid #4b5260; border-radius:4px; padding:3px 6px; min-height:22px;
            }
            QComboBox:hover, QLineEdit:hover, QSpinBox:hover { border-color:#657085; }
            QPushButton {
                background:#3a4252; border:none; border-radius:4px; padding:4px 8px; min-height: 22px;
            }
            QPushButton:hover { background:#4a5467; }
            QPushButton#primary { background:#2f7df6; color:#ffffff; }
            QPushButton#primary:hover { background:#4a90fa; }
            QPushButton#danger { background:#cf3f4f; color:#ffffff; }
            QPushButton#danger:hover { background:#df5263; }
            QPushButton#ghost { background:#4b5260; color:#e6e9ef; }
            QPushButton#ghost:hover { background:#5a6373; }
            QFrame#vline { background:#3a404c; max-width:1px; min-width:1px; border:none; }
            """
        )
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor("#2b2b2b"))
        self.setPalette(palette)
        icon_path = os.path.join(os.path.dirname(__file__), "assets", "app_icon.svg")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

    def _build_launch_panel(self):
        panel = QFrame()
        panel.setObjectName("panelCard")
        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.setContentsMargins(8, 8, 8, 8)

        title = QLabel("启动程序")
        title.setStyleSheet("font-size:14px; color:#6bb3ff; font-weight:bold;")
        layout.addWidget(title)

        amount_row = QHBoxLayout()
        amount_row.addWidget(QLabel("数量:"))
        self.start_amount = QSpinBox()
        self.start_amount.setRange(1, 99)
        self.start_amount.setValue(5)
        self.start_amount.setMaximumWidth(72)
        amount_row.addWidget(self.start_amount)
        amount_row.addStretch()
        layout.addLayout(amount_row)

        layout.addWidget(self._hline())

        maya_locations = insLocation.getInsLocationDict()
        for group_name, locations in sorted(maya_locations.items(), key=lambda x: x[0].lower()):
            group_label = QLabel(group_name)
            group_label.setStyleSheet("font-size:11px; color:#9dc7ff;")
            layout.addWidget(group_label)
            row = QHBoxLayout()
            row.setSpacing(6)
            combo = QComboBox()
            combo.setObjectName(f"maya_{group_name}")
            exe_list = sorted([x.exeFile for x in locations.values() if x.exist], key=str.lower)
            combo.addItems(exe_list)
            row.addWidget(combo)
            launch_btn = QPushButton("启动")
            launch_btn.setObjectName("primary")
            launch_btn.clicked.connect(partial(self.launch_programs, combo))
            row.addWidget(launch_btn)
            layout.addLayout(row)

        deadline_root = insLocation.get_deadline_install_dir()
        if deadline_root:
            dl_label = QLabel("Deadline")
            dl_label.setStyleSheet("font-size:11px; color:#9dc7ff;")
            layout.addWidget(dl_label)
            row = QHBoxLayout()
            row.setSpacing(6)
            deadline_combo = QComboBox()
            deadline_combo.addItem(os.path.join(deadline_root, "bin", "deadlineworker.exe"))
            row.addWidget(deadline_combo)
            deadline_btn = QPushButton("启动")
            deadline_btn.setObjectName("primary")
            deadline_btn.clicked.connect(partial(self.launch_programs, deadline_combo))
            row.addWidget(deadline_btn)
            layout.addLayout(row)

        layout.addStretch()
        panel.setLayout(layout)
        return panel

    def _build_kill_panel(self):
        panel = QFrame()
        panel.setObjectName("panelCard")
        layout = QVBoxLayout()
        layout.setSpacing(5)
        layout.setContentsMargins(8, 8, 8, 8)

        title = QLabel("结束程序")
        title.setStyleSheet("font-size:14px; color:#e74c3c; font-weight:bold;")
        layout.addWidget(title)

        freq_hint = QLabel("程序列表按使用频率自动排序")
        freq_hint.setStyleSheet("color:#8b95a7; font-size:10px;")
        layout.addWidget(freq_hint)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("过滤程序名...")
        self.search_input.textChanged.connect(self.refresh_kill_combo)
        layout.addWidget(self.search_input)

        self.kill_combo = QComboBox()
        layout.addWidget(self.kill_combo)

        button_row = QHBoxLayout()
        button_row.setSpacing(6)
        kill_selected_btn = QPushButton("结束选中程序")
        kill_selected_btn.setObjectName("danger")
        kill_selected_btn.clicked.connect(self.kill_selected_program)
        button_row.addWidget(kill_selected_btn)
        refresh_btn = QPushButton("刷新列表")
        refresh_btn.setObjectName("ghost")
        refresh_btn.clicked.connect(self.refresh_kill_combo)
        button_row.addWidget(refresh_btn)
        layout.addLayout(button_row)

        layout.addWidget(self._hline())
        common_header = QHBoxLayout()
        common_header.setSpacing(6)
        common_header.addWidget(QLabel("常见程序一键结束"))
        add_common_btn = QPushButton("+ 添加")
        add_common_btn.setObjectName("primary")
        add_common_btn.clicked.connect(self.add_common_program)
        common_header.addStretch()
        common_header.addWidget(add_common_btn)
        layout.addLayout(common_header)

        self.common_grid = QGridLayout()
        self.common_grid.setHorizontalSpacing(4)
        self.common_grid.setVerticalSpacing(4)
        layout.addLayout(self.common_grid)

        kill_all_btn = QPushButton("结束所有常见程序")
        kill_all_btn.setObjectName("danger")
        kill_all_btn.clicked.connect(self.kill_common_programs)
        layout.addWidget(kill_all_btn)

        layout.addStretch()
        panel.setLayout(layout)
        return panel

    def _hline(self):
        line = QFrame()
        line.setMaximumHeight(1)
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        return line

    def _vline(self):
        line = QFrame()
        line.setObjectName("vline")
        line.setFrameShape(QFrame.VLine)
        line.setFrameShadow(QFrame.Sunken)
        return line

    def load_config(self):
        if not os.path.exists(self.config_file):
            raise FileNotFoundError(f"配置文件不存在: {self.config_file}")
        try:
            data = _load_config_text(self.config_file)
            print(f"找到配置文件: {self.config_file}")
        except Exception as exc:
            raise RuntimeError(f"读取配置文件失败: {self.config_file}") from exc
        if data is None:
            return {}
        if not isinstance(data, dict):
            raise ValueError(f"配置文件格式错误(应为对象): {self.config_file}")
        return data

    def save_config(self, config):
        os.makedirs(os.path.dirname(self.config_file), exist_ok=True)
        text = _dump_config_text(config)
        with open(self.config_file, "w", encoding="utf-8") as f:
            f.write(text)

    def load_settings(self):
        cfg = self.load_config()
        if "start_amount" in cfg:
            self.start_amount.setValue(int(cfg["start_amount"]))
        self._kill_programs = cfg.get("kill_programs") or COMMON_KILL_PROGRAMS[:]
        self._usage_counts = cfg.get("program_usage_counts") or {}
        if not isinstance(self._usage_counts, dict):
            self._usage_counts = {}
        self.common_programs = cfg.get("common_programs") or COMMON_KILL_PROGRAMS[:6]
        self._rebuild_common_program_buttons()
        self.refresh_kill_combo()

    def refresh_kill_combo(self):
        keyword = self.search_input.text().strip().lower()
        programs = sorted(
            set(self._kill_programs),
            key=lambda p: (-int(self._usage_counts.get(p, 0)), p.lower()),
        )
        if keyword:
            programs = [p for p in programs if keyword in p.lower()]
        self._filtered_programs = programs
        self.kill_combo.clear()
        for p in programs:
            count = int(self._usage_counts.get(p, 0))
            display = f"{p}  ({count})"
            self.kill_combo.addItem(display, p)

    def _rebuild_common_program_buttons(self):
        while self.common_grid.count():
            item = self.common_grid.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        for idx, name in enumerate(self.common_programs):
            btn = QPushButton(name.replace(".exe", ""))
            btn.setObjectName("ghost")
            btn.setToolTip(name)
            btn.clicked.connect(partial(self.kill_program_with_dialog, name))
            row = idx // 3
            col = idx % 3
            self.common_grid.addWidget(btn, row, col)

    def add_common_program(self):
        candidates = sorted(set(self._kill_programs), key=str.lower)
        text, ok = QInputDialog.getItem(
            self,
            "添加常见程序",
            "选择或输入进程名(如 Maya.exe):",
            candidates,
            0,
            True,
        )
        if not ok:
            return
        name = text.strip()
        if not name:
            return
        if "." not in name:
            name = f"{name}.exe"
        if name not in self.common_programs:
            self.common_programs.append(name)
        if name not in self._kill_programs:
            self._kill_programs.append(name)
        self._rebuild_common_program_buttons()
        self.refresh_kill_combo()

    def kill_program(self, program_name):
        result = subprocess.run(
            ["taskkill", "/f", "/im", program_name],
            check=False,
            capture_output=True,
            text=True,
            encoding=locale.getpreferredencoding(False),
            errors="replace",
        )
        stdout = (result.stdout or "").strip()
        stderr = (result.stderr or "").strip()
        ok = result.returncode == 0
        if ok:
            self._usage_counts[program_name] = int(self._usage_counts.get(program_name, 0)) + 1
            self.refresh_kill_combo()

        details = [f"程序: {program_name}", f"返回码: {result.returncode}"]
        if stdout:
            details.append(f"输出: {stdout}")
        if stderr:
            details.append(f"错误: {stderr}")
        return ok, "\n".join(details)

    def kill_program_with_dialog(self, program_name):
        ok, details = self.kill_program(program_name)
        if ok:
            QMessageBox.information(self, "执行结果", details)
        else:
            QMessageBox.warning(self, "执行结果", details)

    def kill_selected_program(self):
        program = self.kill_combo.currentData()
        if not program:
            QMessageBox.warning(self, "提示", "没有可结束的程序。")
            return
        self.kill_program_with_dialog(program)

    def kill_common_programs(self):
        results = []
        success_count = 0
        for name in COMMON_KILL_PROGRAMS:
            ok, details = self.kill_program(name)
            if ok:
                success_count += 1
            results.append(("成功" if ok else "失败") + f" | {details}")
        summary = (
            f"常见程序结束完成: 成功 {success_count}/{len(COMMON_KILL_PROGRAMS)}\n\n"
            + "\n\n".join(results)
        )
        if success_count == len(COMMON_KILL_PROGRAMS):
            QMessageBox.information(self, "执行结果", summary)
        else:
            QMessageBox.warning(self, "执行结果", summary)

    def launch_programs(self, combo):
        cmd = combo.currentText().strip()
        if not cmd:
            QMessageBox.warning(self, "提示", "未找到可启动程序。")
            return
        count = self.start_amount.value()
        started = []
        failed = []
        for i in range(count):
            env = copy.deepcopy(ORI_ENV_VAR)
            env["PYTHONPATH"] = os.path.dirname(cmd)
            extra_args = ["-name", str(i), "-nogui"] if "deadlineworker.exe" in cmd.lower() else []
            try:
                proc = subprocess.Popen([cmd, *extra_args], env=env)
                started.append(proc.pid)
            except Exception as exc:
                failed.append(f"#{i}: {exc}")

        lines = [f"命令: {cmd}", f"计划数量: {count}", f"成功启动: {len(started)}"]
        if started:
            lines.append(f"PID: {', '.join(str(x) for x in started[:10])}")
        if failed:
            lines.append("失败明细:")
            lines.extend(failed[:10])
        message = "\n".join(lines)
        if failed:
            QMessageBox.warning(self, "执行结果", message)
        else:
            QMessageBox.information(self, "执行结果", message)

    def _persist_settings(self):
        cfg = self.load_config()
        cfg["start_amount"] = self.start_amount.value()
        cfg["kill_programs"] = sorted(set(self._kill_programs), key=str.lower)
        cfg["common_programs"] = list(dict.fromkeys(self.common_programs))
        cfg["program_usage_counts"] = {
            k: int(v) for k, v in self._usage_counts.items() if int(v) > 0
        }
        self.save_config(cfg)


def main():
    app = QApplication(sys.argv)
    win = ProgramLauncher()
    app.aboutToQuit.connect(win._persist_settings)
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()

