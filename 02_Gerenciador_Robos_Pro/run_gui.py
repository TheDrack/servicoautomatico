import tkinter as tk
from core.supervisor import Supervisor
from ui.main_gui import RobotGUI

if __name__ == "__main__":
    root = tk.Tk()
    supervisor = Supervisor(max_workers=3)
    app = RobotGUI(root, supervisor)
    
    app.update_loop() # Inicia o monitoramento
    root.mainloop()
