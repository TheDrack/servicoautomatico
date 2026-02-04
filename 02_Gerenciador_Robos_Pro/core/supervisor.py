import subprocess
import threading
import uuid
import queue
from multiprocessing import Queue as MPQueue
from typing import Callable


class Supervisor:
    """
    Supervisor para gerenciamento de tarefas assíncronas.
    
    Controla execução de robôs externos com limite de workers
    e comunicação via fila entre threads/processos.
    """
    def __init__(self, max_workers: int = 4) -> None:
        self.job_queue = queue.Queue()
        self.result_queue = MPQueue()  # Comunicação entre threads/processos
        self.active_jobs = {}
        self.max_workers = max_workers

    def submit_task(self, name: str, path_or_func: str | Callable) -> str:
        job_id = f"{name}_{uuid.uuid4().hex[:4]}"
        self.job_queue.put((job_id, name, path_or_func))
        return job_id

    def tick(self) -> None:
        """Monitora e lança novos jobs. Pode ser chamado por GUI ou Voz."""
        while len(self.active_jobs) < self.max_workers and not self.job_queue.empty():
            job_id, name, task = self.job_queue.get()
            
            # Lógica para Nicho A (Script Externo)
            t = threading.Thread(target=self._execute_external, args=(job_id, task), daemon=True)
            t.start()
            self.active_jobs[job_id] = t

    def _execute_external(self, job_id: str, path: str) -> None:
        try:
            process = subprocess.Popen(
                ["python", "-u", path],
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1
            )
            for line in iter(process.stdout.readline, ''):
                self.result_queue.put({"job_id": job_id, "type": "log", "data": line})
            
            process.wait()
            self.result_queue.put({"job_id": job_id, "type": "status", "data": "Finalizado"})
        except Exception as e:
            self.result_queue.put({"job_id": job_id, "type": "status", "data": f"Erro: {e}"})
        finally:
            if job_id in self.active_jobs: del self.active_jobs[job_id] # Limpeza básica
