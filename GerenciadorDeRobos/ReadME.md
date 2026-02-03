# 🤖 Robot Manager Pro (Python)

A desktop application built in Python to **manage, monitor, and control multiple automation scripts (robots)** running in parallel.

This tool was designed for real-world automation environments where multiple Python-based robots need to be executed, monitored, and safely terminated, with **real-time log visualization**.

---

## 🚀 Features

- Execute multiple Python automation scripts simultaneously
- Real-time log streaming (stdout) for each robot
- Safe process lifecycle management (start, monitor, terminate)
- Automatic handling of child processes
- Thread-safe log processing using queues
- Simple and intuitive GUI built with Tkinter
- Status tracking: Running, Finished, Manually Terminated

---

## 🧠 Why this project exists

In automation and RPA environments, it’s common to run **several independent scripts at the same time**, often interacting with unstable systems, legacy software, or external services.

This project was created to:
- Centralize robot execution
- Prevent orphan processes
- Improve observability of automation runs
- Provide operational control without relying on heavy RPA platforms

---

## 🛠️ Tech Stack

- **Python 3**
- **Tkinter** – GUI
- **subprocess** – Process execution
- **threading + queue** – Concurrency & thread-safe logging
- **psutil** – Process tree management

---

## 🧩 How it works (High-level)

1. Select a Python automation script (`.py`)
2. Define how many times it should run
3. The manager spawns isolated processes
4. Each robot’s output is captured in real time
5. Robots can be safely terminated, including all child processes
6. Execution status is tracked in the UI

---

## ▶️ How to Run

```bash
pip install psutil
python robot_manager.py
Python 3.9+ recommended
⚠️ Notes
The application assumes each robot is a Python script executed via CLI
Designed for legitimate automation use cases
No credential handling or sensitive data included in this repository
📌 Possible Improvements
Export logs to files
Group robots by project
Schedule executions
Remote execution via API (HTTP-based controller)
Authentication and access control
👤 Author
Jesus Davi Pontes Anhaia
Automation Engineer focused on:
Python automation
RPA (API + UI hybrid)
System integration
Reverse engineering of legacy systems
Solving complex automation problems in unstable environments