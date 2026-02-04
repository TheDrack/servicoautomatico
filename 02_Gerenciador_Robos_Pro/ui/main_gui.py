import tkinter as tk
from tkinter import ttk, filedialog
import queue
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from core.supervisor import Supervisor


class RobotGUI:
    """
    Interface gráfica para o Supervisor de Robôs.
    
    Fornece visualização de status de jobs e logs em tempo real.
    """
    def __init__(self, root: tk.Tk, supervisor: "Supervisor") -> None:
        self.root = root
        self.sup = supervisor
        self._build_ui()

    def _build_ui(self) -> None:
        self.root.title("Gerenciador Pro 2.0")
        
        # Tabela
        self.tree = ttk.Treeview(self.root, columns=("Status"), height=8)
        self.tree.heading("#0", text="ID / Robô")
        self.tree.heading("Status", text="Status")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Logs
        self.console = tk.Text(self.root, bg="black", fg="white", height=10)
        self.console.pack(fill=tk.BOTH, padx=10, pady=10)

        tk.Button(self.root, text="Adicionar Robô", command=self.add_job).pack(pady=5)

    def add_job(self) -> None:
        path = filedialog.askopenfilename(filetypes=[("Python", "*.py")])
        if path:
            name = path.split("/")[-1]
            job_id = self.sup.submit_task(name, path)
            self.tree.insert("", "end", iid=job_id, text=job_id, values=("Aguardando...",))

    def update_loop(self) -> None:
        self.sup.tick()
        try:
            while True:
                msg = self.sup.result_queue.get_nowait()
                jid = msg["job_id"]
                if msg["type"] == "log":
                    self.console.insert(tk.END, f"[{jid}] {msg['data']}")
                    self.console.see(tk.END)
                elif msg["type"] == "status":
                    if self.tree.exists(jid): self.tree.set(jid, "Status", msg["data"])
        except queue.Empty:
            pass
        self.root.after(100, self.update_loop)
