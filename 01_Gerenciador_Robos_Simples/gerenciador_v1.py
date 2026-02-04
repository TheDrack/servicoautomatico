import os
import subprocess
import threading
import queue
import time
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
import psutil
import uuid
from typing import Dict, Any


class RobotManagerApp:
    """
    Gerenciador de Robôs com interface gráfica Tkinter.
    
    Permite execução paralela de múltiplos scripts Python com
    monitoramento em tempo real e controle de processos.
    """
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Gerenciador de Robôs Pro")
        self.root.geometry("900x650")

        self.robots = {}
        self.log_queue = queue.Queue()

        self._build_ui()
        self._consume_logs()

    # ---------------- UI ---------------- #

    def _build_ui(self) -> None:
        frame_top = tk.Frame(self.root)
        frame_top.pack(pady=10, fill=tk.X)

        tk.Label(frame_top, text="Execuções:").pack(side=tk.LEFT, padx=(20, 5))

        self.exec_entry = tk.Entry(frame_top, width=5)
        self.exec_entry.insert(0, "1")
        self.exec_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(
            frame_top,
            text="Adicionar Robô",
            command=self.add_robot,
            bg="#2ecc71",
            fg="white",
            relief=tk.FLAT
        ).pack(side=tk.LEFT, padx=5)

        tk.Button(
            frame_top,
            text="Limpar Finalizados",
            command=self.clear_finished
        ).pack(side=tk.LEFT, padx=5)

        self.tree = ttk.Treeview(
            self.root,
            columns=("Nome", "Caminho", "Status"),
            height=8
        )
        self.tree.heading("#0", text="ID")
        self.tree.heading("Nome", text="Nome")
        self.tree.heading("Caminho", text="Caminho")
        self.tree.heading("Status", text="Status")

        self.tree.column("#0", width=140)
        self.tree.pack(fill=tk.X, padx=10)

        self.tree.bind("<Double-1>", lambda e: self.stop_selected_robot())

        tk.Label(
            self.root,
            text="Console de Logs (Saída em Tempo Real)",
            font=("Arial", 10, "bold")
        ).pack(pady=(10, 0))

        self.log_text = tk.Text(
            self.root,
            bg="#1e1e1e",
            fg="#61ff61",
            height=15,
            font=("Consolas", 10)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.status_var = tk.StringVar(value="Robôs Ativos: 0")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ---------------- Robot Control ---------------- #

    def add_robot(self) -> None:
        robot_file = filedialog.askopenfilename(
            filetypes=[("Python files", "*.py")]
        )
        if not robot_file:
            return

        try:
            runs = int(self.exec_entry.get())
            if runs < 1:
                raise ValueError("Número de execuções deve ser positivo")
            if runs > 100:
                raise ValueError("Número máximo de execuções é 100")
        except ValueError as e:
            messagebox.showerror(
                "Erro",
                f"Quantidade inválida: {str(e)}"
            )
            return

        for _ in range(runs):
            self.start_robot(robot_file)

        self._update_status_bar()

    def start_robot(self, robot_file: str) -> None:
        robot_name = os.path.basename(robot_file)
        robot_id = f"{robot_name}_{uuid.uuid4().hex[:6]}"

        try:
            process = subprocess.Popen(
                ["python", "-u", robot_file],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1
            )
        except Exception as e:
            messagebox.showerror(
                "Erro ao iniciar robô",
                str(e)
            )
            return

        item_id = self.tree.insert(
            "",
            "end",
            text=robot_id,
            values=(robot_name, robot_file, "Rodando")
        )

        self.robots[robot_id] = {
            "process": process,
            "tree_id": item_id
        }

        threading.Thread(
            target=self._read_output,
            args=(robot_id, process),
            daemon=True
        ).start()

    def stop_selected_robot(self) -> None:
        selected = self.tree.selection()
        if not selected:
            return

        robot_id = self.tree.item(selected[0])["text"]
        status = self.tree.set(selected[0], "Status")

        if status != "Rodando":
            return

        if not messagebox.askyesno(
            "Encerrar",
            f"Deseja encerrar o robô {robot_id}?"
        ):
            return

        self._kill_robot(robot_id)

    def _kill_robot(self, robot_id: str) -> None:
        if robot_id not in self.robots:
            return

        try:
            proc = self.robots[robot_id]["process"]
            ps_proc = psutil.Process(proc.pid)

            for child in ps_proc.children(recursive=True):
                child.terminate()

            ps_proc.terminate()
            ps_proc.wait(timeout=3)

        except Exception:
            try:
                ps_proc.kill()
            except Exception:
                pass
        finally:
            self._mark_as_finished(robot_id, manual=True)

    def _mark_as_finished(self, robot_id: str, manual: bool = False) -> None:
        if robot_id not in self.robots:
            return

        item_id = self.robots[robot_id]["tree_id"]
        status = "Encerrado manualmente" if manual else "Finalizado"

        if self.tree.exists(item_id):
            self.tree.set(item_id, "Status", status)

        self._update_status_bar()

    def clear_finished(self) -> None:
        to_remove = []

        for rid, data in self.robots.items():
            item_id = data["tree_id"]
            status = self.tree.set(item_id, "Status")

            if status != "Rodando":
                self.tree.delete(item_id)
                to_remove.append(rid)

        for rid in to_remove:
            del self.robots[rid]

        self._update_status_bar()

    # ---------------- Logs & Monitoring ---------------- #

    def _read_output(self, robot_id: str, process: subprocess.Popen) -> None:
        try:
            for line in iter(process.stdout.readline, ''):
                if line:
                    self.log_queue.put((robot_id, line))
        finally:
            process.wait()
            self.log_queue.put((robot_id, "__PROCESS_FINISHED__"))

    def _consume_logs(self) -> None:
        try:
            while True:
                robot_id, message = self.log_queue.get_nowait()

                if message == "__PROCESS_FINISHED__":
                    self._mark_as_finished(robot_id)
                elif robot_id in self.robots:
                    self.log_text.insert(
                        tk.END,
                        f"[{robot_id}] {message}"
                    )
                    self.log_text.see(tk.END)

        except queue.Empty:
            pass

        self.root.after(100, self._consume_logs)

    def _update_status_bar(self) -> None:
        ativos = sum(
            1 for r in self.robots.values()
            if self.tree.set(r["tree_id"], "Status") == "Rodando"
        )

        self.status_var.set(
            f"Robôs Ativos: {ativos} | Total no Histórico: {len(self.robots)}"
        )

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = RobotManagerApp()
    app.run()