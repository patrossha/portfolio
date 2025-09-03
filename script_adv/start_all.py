import subprocess
import os
 
scripts = ["andamentos_dcp.py", "andamentos_eproc.py", "andamentos_esaj.py", "andamentos_pje.py", "andamentos_trt.py", "andamentos_legalone.py", "andamentos_final.py", "publi_legalone.py", "recorte_oab.py", "publicacoes_final.py"
]

# Caminho base
base_path = os.path.dirname(os.path.abspath(__file__))

for script in scripts:
    script_path = os.path.join(base_path, script)
    print(f"Executando {script}...") 
    subprocess.run(["python", script_path])
    print(f"Finalizou {script}.\n")
