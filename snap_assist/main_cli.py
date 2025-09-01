import contextlib
import json
import random
import sys
import time
from typing import Optional

import pyperclip  # Dependency: pip install pyperclip
import requests  # Dependency: pip install requests
from PyQt6.QtCore import QObject, Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from snap_assist.chat_modes import get_chat_modes

# Use a cryptographically secure RNG for backoff jitter to satisfy S311
secure_random = random.SystemRandom()

# --- Ollama API Configuration ---
OLLAMA_ENDPOINT = "http://localhost:11434/api/generate"
OLLAMA_SELECTED_MODEL = "llama3.2:latest"
TEXT_FONT = "Fantasque Sans Mono"
TEXT_FONT_SIZE = 18
DEFAULT_MODE = "Rewrite"


class ApiWorker(QObject):
    """
    A worker object that runs the Ollama API call in a separate thread
    to avoid blocking the main GUI thread.
    """

    chunk_received: pyqtSignal = pyqtSignal(str)
    chunk_received_with_mode: pyqtSignal = pyqtSignal(str, str)  # (mode_name, chunk)
    finished: pyqtSignal = pyqtSignal()
    error_occurred: pyqtSignal = pyqtSignal(str)

    def __init__(self, mode_name: str, prompt: str):
        super().__init__()
        self.mode_name = mode_name
        self.prompt = prompt

    def run(self):  # noqa: C901
        """
        Makes the streaming API call to Ollama.
        Emits signals for each response chunk, on error, and on completion.
        It is designed to be interruptible.
        """
        url = OLLAMA_ENDPOINT
        payload = {
            "model": OLLAMA_SELECTED_MODEL,
            "prompt": self.prompt,
            "stream": True,
        }
        headers = {"Content-Type": "application/json"}

        # Retry configuration
        max_retries = 3
        base_backoff = 1.5
        max_backoff = 10.0

        attempt = 0
        while attempt <= max_retries:
            # Check if the thread has been told to stop before starting the request
            if QThread.currentThread().isInterruptionRequested():
                self.finished.emit()
                return
            try:
                with requests.post(url, json=payload, headers=headers, stream=True, timeout=20) as response:
                    try:
                        response.raise_for_status()
                    except requests.exceptions.HTTPError:
                        status = response.status_code
                        # Retry on rate limit or server-side errors
                        if status == 429 or 500 <= status < 600:
                            # Honor Retry-After header if present and numeric
                            retry_after_hdr = response.headers.get("Retry-After")
                            if retry_after_hdr and retry_after_hdr.isdigit():
                                sleep_secs = float(retry_after_hdr)
                            else:
                                sleep_secs = min(max_backoff, (base_backoff**attempt) + secure_random.uniform(0, 0.5))
                            if attempt < max_retries and not QThread.currentThread().isInterruptionRequested():
                                # Sleep with interruption checks
                                slept = 0.0
                                step = 0.1
                                while slept < sleep_secs:
                                    if QThread.currentThread().isInterruptionRequested():
                                        self.finished.emit()
                                        return
                                    time.sleep(step)
                                    slept += step
                                attempt += 1
                                continue
                            else:
                                raise
                        else:
                            raise

                    # Successful response; stream content
                    for line in response.iter_lines():
                        if QThread.currentThread().isInterruptionRequested():
                            break
                        if line:
                            decoded_line = line.decode("utf-8")
                            data = json.loads(decoded_line)
                            response_chunk = data.get("response", "")
                            self.chunk_received.emit(response_chunk)
                            self.chunk_received_with_mode.emit(self.mode_name, response_chunk)
                            if data.get("done", False):
                                break
                    # Completed successfully; exit retry loop
                    break

            except requests.exceptions.HTTPError as e:
                if not QThread.currentThread().isInterruptionRequested():
                    self.error_occurred.emit(f"API request failed: {e}")
                break
            except (
                requests.exceptions.ConnectTimeout,
                requests.exceptions.ReadTimeout,
                requests.exceptions.ChunkedEncodingError,
                requests.exceptions.ConnectionError,
            ) as e:
                # Retry on transient network errors
                sleep_secs = min(max_backoff, (base_backoff**attempt) + secure_random.uniform(0, 0.5))
                if attempt < max_retries and not QThread.currentThread().isInterruptionRequested():
                    slept = 0.0
                    step = 0.2
                    while slept < sleep_secs:
                        if QThread.currentThread().isInterruptionRequested():
                            self.finished.emit()
                            return
                        time.sleep(step)
                        slept += step
                    attempt += 1
                    continue
                else:
                    if not QThread.currentThread().isInterruptionRequested():
                        self.error_occurred.emit(f"Network error: {e}")
                    break
            except requests.exceptions.RequestException as e:
                # Non-retriable
                if not QThread.currentThread().isInterruptionRequested():
                    self.error_occurred.emit(f"API request failed: {e}")
                break
            finally:
                # Note: finished signal emitted after loop to avoid double emission on retries
                pass

        # Use try-except to prevent crash if worker is deleted before signal emission
        with contextlib.suppress(RuntimeError):
            self.finished.emit()


class ResultPanel(QWidget):
    """Reusable widget for displaying individual mode output with header, copy, and progress state."""

    refresh_requested: pyqtSignal = pyqtSignal(str)
    add_to_combine_requested: pyqtSignal = pyqtSignal(str)

    def __init__(self, mode_name: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.mode_name = mode_name
        self.is_expanded = False
        self.normal_height = 150  # Normal height
        self.expanded_height = 400  # Expanded height
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus)  # Make panel focusable

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(6, 6, 6, 6)
        self.layout.setSpacing(4)

        # Header row: Mode label, spacer, Copy button
        header = QHBoxLayout()
        header.setSpacing(6)

        self.title_label = QLabel(mode_name)
        self.title_label.setObjectName("resultPanelTitle")

        # Hint label for expand/collapse
        self.hint_label = QLabel("🖱️ Click to expand • ESC to collapse")
        self.hint_label.setObjectName("hintLabel")
        self.hint_label.setStyleSheet("color: #666666; font-size: 10px;")

        self.progress = QProgressBar()
        self.progress.setObjectName("resultPanelProgress")
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(4)
        self.progress.hide()

        self.refresh_btn = QPushButton("Refresh")
        self.refresh_btn.setObjectName("refreshPanelButton")
        self.refresh_btn.setFixedWidth(80)
        # noinspection PyUnresolvedReferences
        self.refresh_btn.clicked.connect(lambda: self.refresh_requested.emit(self.mode_name))

        self.copy_btn = QPushButton("Copy")
        self.copy_btn.setObjectName("copyResultButton")
        self.copy_btn.setFixedWidth(70)
        # noinspection PyUnresolvedReferences
        self.copy_btn.clicked.connect(self.copy_text)

        self.add_combine_btn = QPushButton("Add >>")
        self.add_combine_btn.setObjectName("addCombineButton")
        self.add_combine_btn.setFixedWidth(85)
        # noinspection PyUnresolvedReferences
        self.add_combine_btn.clicked.connect(lambda: self.add_to_combine_requested.emit(self.mode_name))

        header.addWidget(self.title_label)
        header.addWidget(self.hint_label)
        header.addStretch(1)
        header.addWidget(self.refresh_btn)
        header.addWidget(self.copy_btn)
        header.addWidget(self.add_combine_btn)

        # Read-only output area
        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setFixedHeight(self.normal_height)
        with contextlib.suppress(Exception):
            self.output.setFontFamily(TEXT_FONT)
            self.output.setFontPointSize(TEXT_FONT_SIZE)

        self.layout.addLayout(header)
        self.layout.addWidget(self.output)
        self.layout.addWidget(self.progress)

    def mousePressEvent(self, event):
        """Expand the output area when clicked."""
        if event.button() == Qt.MouseButton.LeftButton:
            self.toggle_expand()
        super().mousePressEvent(event)

    def keyPressEvent(self, event):
        """Collapse the output area when ESC is pressed."""
        if event.key() == Qt.Key.Key_Escape and self.is_expanded:
            self.toggle_expand()
        else:
            super().keyPressEvent(event)

    def toggle_expand(self):
        """Toggle between expanded and normal states."""
        self.is_expanded = not self.is_expanded
        if self.is_expanded:
            self.output.setFixedHeight(self.expanded_height)
        else:
            self.output.setFixedHeight(self.normal_height)

    # Public API
    def set_loading(self, loading: bool):
        if loading:
            self.progress.show()
            # Indeterminate/busy mode
            self.progress.setRange(0, 0)
            self.copy_btn.setEnabled(False)
            self.add_combine_btn.setEnabled(False)
            if not self.output.toPlainText().strip() or self.output.toPlainText().strip() in (
                "Queued...",
                "Loading...",
            ):
                self.output.setPlainText("Loading...")
            # Reset error style when starting to load
            self.title_label.setStyleSheet("")
        else:
            self.progress.hide()
            # Reset to determinate idle state
            self.progress.setRange(0, 1)
            self.progress.setValue(1)
            self.copy_btn.setEnabled(True)
            self.update_add_button_state()

    def set_queued(self):
        """Show queued state without progress bar."""
        self.progress.hide()
        self.copy_btn.setEnabled(False)
        self.add_combine_btn.setEnabled(False)
        self.title_label.setStyleSheet("")
        self.output.setPlainText("Queued...")

    def append_text(self, text: str):
        # Clear placeholder 'Loading...' if present
        if self.output.toPlainText().strip() == "Loading...":
            self.output.clear()
        self.output.insertPlainText(text)

    def set_text(self, text: str):
        self.output.setPlainText(text)

    def set_error(self, error_message: str):
        # Show error in panel and style title as error
        self.set_loading(False)
        self.output.setPlainText(error_message)
        self.title_label.setStyleSheet("color: #e74c3c; font-weight: 600;")

    def text(self) -> str:
        return self.output.toPlainText()

    def copy_text(self):
        try:
            pyperclip.copy(self.text())
        except Exception:
            # Fallback: select all and rely on user copy if clipboard fails
            self.output.selectAll()

    def update_add_button_state(self):
        """Update the Add >> button state based on whether there's actual content."""
        text = self.text().strip()
        has_content = bool(text and text not in ("Queued...", "Loading..."))
        self.add_combine_btn.setEnabled(has_content)


class CombineControlPanel(QWidget):
    """Right-side panel for combining and refining LLM responses."""

    refine_requested: pyqtSignal = pyqtSignal(str)  # Emits combined text

    def __init__(self, mode_panels: dict[str, "ResultPanel"], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.mode_panels = mode_panels
        self.selected_responses: dict[str, str] = {}  # mode_name -> response_text
        self.setFixedWidth(350)

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(8, 8, 8, 8)
        self.layout.setSpacing(6)

        # Title
        title = QLabel("Combine & Refine")
        title.setObjectName("combineTitle")
        title.setStyleSheet("font-weight: 600; font-size: 14px;")
        self.layout.addWidget(title)

        # Instructions
        instructions = QLabel("Click 'Add' buttons on panels to include responses for refinement.")
        instructions.setWordWrap(True)
        instructions.setStyleSheet("font-size: 11px; color: #666; padding: 4px;")
        self.layout.addWidget(instructions)

        # Selected responses list
        selected_label = QLabel("Selected Responses:")
        selected_label.setStyleSheet("font-size: 12px; color: #666;")
        self.layout.addWidget(selected_label)

        self.selected_scroll = QScrollArea()
        self.selected_scroll.setWidgetResizable(True)
        self.selected_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.selected_scroll.setMaximumHeight(120)

        selected_widget = QWidget()
        self.selected_layout = QVBoxLayout(selected_widget)
        self.selected_layout.setContentsMargins(0, 0, 0, 0)
        self.selected_layout.setSpacing(2)

        self.selected_layout.addStretch(1)
        self.selected_scroll.setWidget(selected_widget)
        self.layout.addWidget(self.selected_scroll)

        # Control buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(6)

        self.clear_button = QPushButton("Clear All")
        self.clear_button.setObjectName("clearButton")
        # noinspection PyUnresolvedReferences
        self.clear_button.clicked.connect(self.clear_all)
        buttons_layout.addWidget(self.clear_button)

        self.refine_button = QPushButton("Refine")
        self.refine_button.setObjectName("refineButton")
        # noinspection PyUnresolvedReferences
        self.refine_button.clicked.connect(self.request_refine)
        buttons_layout.addWidget(self.refine_button)

        self.layout.addLayout(buttons_layout)

        # Refined result area
        result_label = QLabel("Refined Result:")
        result_label.setStyleSheet("font-size: 12px; color: #666;")
        self.layout.addWidget(result_label)

        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        with contextlib.suppress(Exception):
            self.result_text.setFontFamily(TEXT_FONT)
            self.result_text.setFontPointSize(TEXT_FONT_SIZE)
        self.layout.addWidget(self.result_text)

        self.copy_result_button = QPushButton("Copy")
        self.copy_result_button.setObjectName("copyResultButton")
        # noinspection PyUnresolvedReferences
        self.copy_result_button.clicked.connect(self.copy_result)
        self.layout.addWidget(self.copy_result_button)

        self.update_selected_list()

    def add_response(self, mode_name: str):
        """Add a response to the selection."""
        panel = self.mode_panels.get(mode_name)
        if panel:
            text = panel.text().strip()
            if text and text not in ("Queued...", "Loading..."):
                self.selected_responses[mode_name] = text
                self.update_selected_list()

    def remove_response(self, mode_name: str):
        """Remove a response from the selection."""
        if mode_name in self.selected_responses:
            del self.selected_responses[mode_name]
            self.update_selected_list()

    def update_selected_list(self):
        """Update the visual list of selected responses."""
        # Clear all existing items including stretch
        while self.selected_layout.count() > 0:
            item = self.selected_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Add selected items
        for mode_name in self.selected_responses:
            item_layout = QHBoxLayout()
            item_layout.setSpacing(4)

            label = QLabel(f"{mode_name}")
            label.setStyleSheet("font-size: 11px; color: #333;")
            item_layout.addWidget(label)

            remove_btn = QPushButton("X")
            remove_btn.setFixedSize(20, 20)
            remove_btn.setStyleSheet("font-size: 12px; color: #e74c3c; border: none; background: transparent;")
            # noinspection PyUnresolvedReferences
            remove_btn.clicked.connect(lambda checked=False, m=mode_name: self.remove_response(m))
            item_layout.addWidget(remove_btn)

            self.selected_layout.addLayout(item_layout)

        # Add stretch at the end
        self.selected_layout.addStretch(1)

        # Update refine button state
        self.refine_button.setEnabled(bool(self.selected_responses))

    def clear_all(self):
        """Clear all selected responses."""
        self.selected_responses.clear()
        self.update_selected_list()
        self.clear_result()
        # Force comprehensive UI refresh
        self.selected_scroll.update()
        self.selected_scroll.viewport().update()
        self.update()
        QApplication.processEvents()  # Process pending UI events

    def request_refine(self):
        """Emit signal with combined text for refinement."""
        if not self.selected_responses:
            return

        # Build combined text
        selected_texts = []
        for mode_name, text in self.selected_responses.items():
            selected_texts.append(f"## {mode_name}\n{text}")

        combined_text = "\n\n".join(selected_texts)
        self.refine_requested.emit(combined_text)

    def set_refined_result(self, result: str):
        """Display the refined result."""
        self.result_text.setPlainText(result)

    def copy_result(self):
        """Copy the refined result to clipboard."""
        result_text = self.result_text.toPlainText().strip()
        if result_text and not result_text.startswith("Error:"):
            try:
                pyperclip.copy(result_text)
            except Exception:
                # Fallback: select all and rely on user copy if clipboard fails
                self.result_text.selectAll()

    def clear_result(self):
        """Clear the refined result area."""
        self.result_text.clear()


class AppWindow(QMainWindow):
    """
    A desktop application that processes clipboard text using Ollama.
    The user can add additional manual instructions. The "Copy Text" button
    copies the result and closes the app. The window is centered on screen.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Popup")
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setGeometry(100, 100, 1250, 700)
        self.center_window()

        # References to the current running threads and workers per mode
        self.api_thread = None  # legacy
        self.api_worker = None  # legacy
        self.api_threads: dict[str, QThread] = {}
        self.api_workers: dict[str, ApiWorker] = {}
        self.mode_panels: dict[str, ResultPanel] = {}
        # Generation guard to avoid stale updates from previous runs
        self.current_generation: int = 0
        # Per-mode generation tokens for precise refresh control
        self.mode_generations: dict[str, int] = {}
        # Per-mode result tracking
        self.mode_results: dict[str, str] = {}
        self.mode_status: dict[str, str] = {}
        self.mode_errors: dict[str, str] = {}
        # Concurrency control
        self.max_concurrent_requests: int = 3
        self.pending_modes: list[str] = []
        self.running_modes: set[str] = set()
        # Store the main text from the clipboard at launch
        self.main_clipboard_text = ""

        self.central_widget = QWidget()
        self.central_widget.setObjectName("centralWidget")
        self.setCentralWidget(self.central_widget)

        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(8, 8, 8, 8)
        self.main_layout.setSpacing(6)

        # Create main horizontal layout
        main_h_layout = QHBoxLayout()
        main_h_layout.setSpacing(6)

        # --- Left Side: Main Content ---
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        # --- Top Bar ---
        top_bar_layout = QHBoxLayout()
        top_bar_layout.setSpacing(6)

        # Removed dropdown; multi-mode UI will be implemented with panels

        self.close_button = QPushButton("Close")
        self.close_button.setObjectName("closeButton")
        # noinspection PyUnresolvedReferences
        self.close_button.clicked.connect(self.close)

        top_bar_layout.addStretch(1)

        # Refresh All button to re-run all prompts simultaneously
        self.refresh_button = QPushButton("Refresh All")
        self.refresh_button.setObjectName("refreshAllButton")
        # noinspection PyUnresolvedReferences
        self.refresh_button.clicked.connect(self.run_all_generations)
        top_bar_layout.addWidget(self.refresh_button)

        top_bar_layout.addWidget(self.close_button)

        # --- Read-Only Text Area ---
        # Replace single read-only area with scrollable container of ResultPanels
        self.read_only_text_area = None
        self.panels_container = QWidget()
        self.panels_layout = QVBoxLayout(self.panels_container)
        self.panels_layout.setContentsMargins(0, 0, 0, 0)
        self.panels_layout.setSpacing(6)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setWidget(self.panels_container)

        # Create panels for all chat modes
        for mode_name in get_chat_modes():
            panel = ResultPanel(mode_name)
            self.mode_panels[mode_name] = panel
            # noinspection PyUnresolvedReferences
            panel.refresh_requested.connect(self.refresh_mode)
            self.panels_layout.addWidget(panel)
        self.panels_layout.addStretch(1)

        # Assemble left layout
        left_layout.addLayout(top_bar_layout)
        left_layout.addWidget(self.scroll_area, 1)

        # --- Right Side: Combine Control Panel ---
        self.combine_panel = CombineControlPanel(self.mode_panels)
        self.combine_panel.refine_requested.connect(self.run_refine_generation)

        # Connect panel signals to combine panel after it's created
        for panel in self.mode_panels.values():
            # noinspection PyUnresolvedReferences
            panel.add_to_combine_requested.connect(self.combine_panel.add_response)

        # Assemble main horizontal layout
        main_h_layout.addWidget(left_widget, 1)
        main_h_layout.addWidget(self.combine_panel, 0)

        # --- Assemble Main Layout ---
        self.main_layout.addLayout(main_h_layout, 1)

        self.apply_styles()

        self.process_clipboard_on_launch()

    # Dropdown removed; generation is triggered on launch and via future Refresh All

    def update_copy_button_state(self):
        """Enable Copy button only when no mode is running."""
        any_running = any(status == "running" for status in self.mode_status.values())
        self.copy_text_button.setEnabled(not any_running)

    # Build the prompt for a given mode using clipboard text and mode instructions
    def build_prompt(self, mode_name: str) -> str:
        modes = get_chat_modes()
        instruction = modes.get(mode_name, "")
        text = self.main_clipboard_text or ""
        if instruction:
            return f"{instruction}\n\n{text}"
        return text

    def stop_all_threads(self, wait: bool = False):
        """Request interruption and quit for all running threads, optionally wait for completion."""
        # Stop legacy single thread if used
        try:
            if self.api_thread and self.api_thread.isRunning():
                self.api_thread.requestInterruption()
                self.api_thread.quit()
                if wait:
                    self.api_thread.wait()
        except RuntimeError:
            pass
        # Stop all per-mode threads
        try:
            for thread in list(self.api_threads.values()):
                if thread and thread.isRunning():
                    thread.requestInterruption()
                    thread.quit()
                    if wait:
                        thread.wait()
        except RuntimeError:
            pass

    def run_all_generations(self):
        """Start concurrent generation for all modes, each updating its own panel, with concurrency limit."""
        if not self.main_clipboard_text.strip():
            # Show message on first panel if available
            if self.mode_panels:
                first = next(iter(self.mode_panels.values()))
                first.set_text("Main text from clipboard is empty.")
            return

        # Bump generation to invalidate stale callbacks
        self.current_generation += 1
        gen = self.current_generation

        # Stop previous threads if any
        self.stop_all_threads(wait=False)
        # Clear references; old threads may still finish, but their signals will be ignored via generation guard
        self.api_threads.clear()
        self.api_workers.clear()

        # Reset per-mode tracking and UI
        self.mode_results.clear()
        self.mode_errors.clear()
        self.mode_status = {m: "queued" for m in self.mode_panels}

        # Clear outputs and set queued state initially
        for panel in self.mode_panels.values():
            panel.set_text("")
            panel.set_loading(False)
            panel.set_queued()
            # Set per-mode generation token
            self.mode_generations[panel.mode_name] = gen

        # Build queue and start up to the concurrency limit
        self.pending_modes = list(self.mode_panels)
        self.running_modes.clear()

        initial = min(self.max_concurrent_requests, len(self.pending_modes))
        for _ in range(initial):
            self.start_next_in_queue()

    def start_next_in_queue(self):
        if not self.pending_modes:
            return
        mode_name = self.pending_modes.pop(0)
        # Use the mode's own generation token to avoid cross-contamination
        generation = self.mode_generations.get(mode_name, self.current_generation)
        self.start_worker_for_mode(generation, mode_name)

    def start_worker_for_mode(self, generation: int, mode_name: str):
        # Mark running and update UI
        self.mode_status[mode_name] = "running"
        self.running_modes.add(mode_name)
        panel = self.mode_panels.get(mode_name)
        if panel:
            panel.set_loading(True)

        # Start worker per mode
        thread = QThread()
        worker = ApiWorker(mode_name, self.build_prompt(mode_name))
        worker.moveToThread(thread)

        # Start/finish lifecycle
        thread.started.connect(worker.run)  # type: ignore[attr-defined]
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        # Connect signals to panel updates with generation guard
        worker.chunk_received_with_mode.connect(lambda m, c, g=generation: self.update_panel_text_if_current(g, m, c))
        # Bind mode for error and finished callbacks with generation guard
        worker.error_occurred.connect(
            lambda msg, m=mode_name, g=generation: self.show_panel_error_if_current(g, m, msg)
        )
        worker.finished.connect(lambda m=mode_name, g=generation: self.on_worker_finished_if_current(g, m))

        self.api_threads[mode_name] = thread
        self.api_workers[mode_name] = worker
        thread.start()

    def update_panel_text(self, mode_name: str, chunk: str):
        panel = self.mode_panels.get(mode_name)
        if panel:
            panel.append_text(chunk)
        # Track text per mode
        self.mode_results[mode_name] = self.mode_results.get(mode_name, "") + chunk

    def update_panel_text_if_current(self, generation: int, mode_name: str, chunk: str):
        if generation != self.mode_generations.get(mode_name):
            return
        self.update_panel_text(mode_name, chunk)

    def show_panel_error(self, mode_name: str, error_message: str):
        panel = self.mode_panels.get(mode_name)
        if panel:
            panel.set_error(error_message)
        # Track error and mark status
        self.mode_errors[mode_name] = error_message
        self.mode_status[mode_name] = "error"

    def show_panel_error_if_current(self, generation: int, mode_name: str, error_message: str):
        if generation != self.mode_generations.get(mode_name):
            return
        self.show_panel_error(mode_name, error_message)

    def on_worker_finished(self, mode_name: str):
        panel = self.mode_panels.get(mode_name)
        if panel:
            panel.set_loading(False)
        # If not already errored, mark as done
        if self.mode_status.get(mode_name) != "error":
            self.mode_status[mode_name] = "done"
        # Remove from running set
        if mode_name in self.running_modes:
            self.running_modes.remove(mode_name)
        # Start next from queue if any
        if self.pending_modes:
            self.start_next_in_queue()

    def on_worker_finished_if_current(self, generation: int, mode_name: str):
        if generation == self.mode_generations.get(mode_name):
            # Current generation: normal finish handling
            self.on_worker_finished(mode_name)
        else:
            # Stale generation: perform cleanup to avoid deadlocks but don't touch panel text/status
            if mode_name in self.running_modes:
                self.running_modes.remove(mode_name)
            if self.pending_modes:
                self.start_next_in_queue()

    def refresh_mode(self, mode_name: str):
        """Refresh generation for a single mode without affecting others."""
        if not self.main_clipboard_text.strip():
            panel = self.mode_panels.get(mode_name)
            if panel:
                panel.set_text("Main text from clipboard is empty.")
            return

        # Assign new generation token for this mode
        self.current_generation += 1
        gen = self.current_generation
        self.mode_generations[mode_name] = gen

        # Stop existing thread for this mode if running
        try:
            thread = self.api_threads.get(mode_name)
            if thread and thread.isRunning():
                thread.requestInterruption()
                thread.quit()
        except RuntimeError:
            pass

        # Ensure not duplicated in pending queue
        with contextlib.suppress(ValueError):
            if mode_name in self.pending_modes:
                self.pending_modes.remove(mode_name)

        # Reset this panel's state
        self.mode_results.pop(mode_name, None)
        self.mode_errors.pop(mode_name, None)
        self.mode_status[mode_name] = "queued"
        panel = self.mode_panels.get(mode_name)
        if panel:
            panel.set_text("")
            panel.set_loading(False)
            panel.set_queued()

        # Start immediately if under concurrency limit; otherwise enqueue
        # If this mode is currently running, enqueue the refreshed job to start after it stops
        if mode_name in self.running_modes:
            self.pending_modes.insert(0, mode_name)
        else:
            if len(self.running_modes) < self.max_concurrent_requests:
                self.start_worker_for_mode(gen, mode_name)
            else:
                self.pending_modes.insert(0, mode_name)

    def run_refine_generation(self, combined_text: str):
        """Run refinement on combined text using direct API call."""
        if not combined_text.strip():
            return

        # Clear previous result
        self.combine_panel.clear_result()

        # Build refinement prompt
        refine_prompt = f"Review all the following alternative sentences and create a single, concise response that combines the strongest elements from each alternative. Preserve the exact tone and intent of the original sentences. Your output should contain ONLY the final consolidated response, with no additional commentary, explanations, or meta-text. Respond in the same language as the input:\n\n{combined_text}"

        # Bump generation
        self.current_generation += 1

        # Create worker for refinement
        thread = QThread()
        worker = ApiWorker("Refine", refine_prompt)
        worker.moveToThread(thread)

        # Connect signals
        thread.started.connect(worker.run)  # type: ignore[attr-defined]
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        # Handle chunks for refinement result
        worker.chunk_received.connect(lambda c: self.combine_panel.result_text.insertPlainText(c))
        worker.error_occurred.connect(lambda msg: self.combine_panel.result_text.setPlainText(f"Error: {msg}"))

        self.api_threads["Refine"] = thread
        self.api_workers["Refine"] = worker
        thread.start()

    def process_clipboard_on_launch(self):
        """Grabs clipboard text and starts the initial API call."""
        self.main_clipboard_text = pyperclip.paste()

        if not self.main_clipboard_text:
            # Show message on first panel if available
            if self.mode_panels:
                first = next(iter(self.mode_panels.values()))
                first.set_text("Clipboard is empty. Copy some text and restart the application.")
            return

        self.run_all_generations()

    def update_output_text(self, chunk: str):
        """Backward-compat for single-panel mode; append to first panel if exists."""
        # If panels exist, append to the first one; otherwise ignore until multi-mode wiring is complete
        first_panel = self.panels_container.findChild(QWidget)
        if isinstance(first_panel, ResultPanel):
            first_panel.append_text(chunk)

    def show_api_error(self, error_message: str):
        """Display API errors in the first panel if present."""
        first_panel = self.panels_container.findChild(QWidget)
        if isinstance(first_panel, ResultPanel):
            first_panel.set_error(error_message)

    def copy_and_close(self):
        """Copies concatenated text from all panels and refined result, then closes the app."""
        texts = []
        for i in range(self.panels_layout.count()):
            item = self.panels_layout.itemAt(i)
            w = item.widget()
            if isinstance(w, ResultPanel):
                t = w.text().strip()
                if t:
                    texts.append(f"## {w.mode_name}\n{t}")

        # Include refined result if available
        refined_text = self.combine_panel.result_text.toPlainText().strip()
        if refined_text and not refined_text.startswith("Error:"):
            texts.append(f"## Refined Result\n{refined_text}")

        text_to_copy = "\n\n".join(texts)
        if text_to_copy:
            pyperclip.copy(text_to_copy)
        self.close()

    def apply_styles(self):
        """Applies QSS to style the application widgets."""
        style_sheet = f"""
            #centralWidget {{
                background-color: #FFFFFF;
                border: 2px solid #005fa3;
                border-radius: 6px;
            }}

            QPushButton, QTextEdit {{
                font-family: {TEXT_FONT};
                font-size: {max(12, TEXT_FONT_SIZE - 2)}px;
            }}

            QPushButton {{
                min-height: 28px;
                border-radius: 4px;
            }}

            QPushButton#refreshAllButton {{
                background-color: #F7FAFF;
                border: 1px solid #B5D3F0;
                padding: 0 14px;
                color: #0B5CAD;
            }}
            QPushButton#refreshAllButton:hover {{ background-color: #EAF3FF; }}

            QPushButton#refreshPanelButton {{
                background-color: #F7FAFF;
                border: 1px solid #B5D3F0;
                padding: 0 10px;
                color: #0B5CAD;
            }}
            QPushButton#refreshPanelButton:hover {{ background-color: #EAF3FF; }}

            QPushButton {{
                background-color: #F0F0F0;
                border: 1px solid #C6C6C6;
                padding: 0 14px;
            }}

            QPushButton:hover {{ background-color: #E6E6E6; }}

            QPushButton#closeButton {{
                background-color: #005fa3;
                color: white;
                border: none;
                font-weight: 500;
                padding: 0 22px;
            }}

            QPushButton#closeButton:hover {{ background-color: #007BD6; }}

            QTextEdit {{
                border: 1px solid #E0E0E0;
                border-radius: 6px;
                padding: 6px;
                background-color: #FFFFFF;
                color: #333333;
            }}

            QLabel#resultPanelTitle {{
                font-weight: 600;
            }}

            QLabel#hintLabel {{
                color: #888888;
                font-size: 10px;
            }}

            QProgressBar#resultPanelProgress {{
                border: none;
                background: transparent;
            }}
        """
        self.setStyleSheet(style_sheet)

    def center_window(self):
        """Center the window on the primary screen."""
        screen = QApplication.primaryScreen()
        if screen is None:
            return
        screen_geo = screen.availableGeometry()
        frame_geo = self.frameGeometry()
        frame_geo.moveCenter(screen_geo.center())
        self.move(frame_geo.topLeft())

    def keyPressEvent(self, event):
        """Handle ESC key press to collapse expanded panels."""
        if event.key() == Qt.Key.Key_Escape:
            # Collapse any expanded panels
            for panel in self.mode_panels.values():
                if panel.is_expanded:
                    panel.toggle_expand()
        else:
            super().keyPressEvent(event)

    def closeEvent(self, event):
        """Ensure any running thread is stopped cleanly on window close."""
        with contextlib.suppress(RuntimeError):
            self.stop_all_threads(wait=True)
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = AppWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
