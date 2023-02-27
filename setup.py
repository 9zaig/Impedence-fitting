import subprocess
import os

"""
VENV Setup
"""

#On trouve le curent work directry
CWD = os.path.dirname(__file__)

#On regarde si on est dans un environnement virtuel
if os.path.exists(os.path.join(CWD,"venv")):
    is_venv = True
else:
    is_venv = False

#chemin du pip venv
VENV = os.path.join(CWD,"venv/scripts/pip")

#pip sans venv
NO_VENV = ["python3", "-m", "pip"]

def send_cmd(cmd):
    """send a command to the cmd

    Args:
        cmd (String): the command to execute by the cmd
    """
    if is_venv:
        cmd.insert(0, VENV)
    else:
        cmd[:0] = NO_VENV
    execute(cmd)

def execute(cmd):
    """execute the command

    Args:
        cmd (String): the command to execute by the cmd

    Raises:
        subprocess.CalledProcessError:
    """
    with subprocess.Popen(cmd, stdout=subprocess.PIPE, bufsize=1, universal_newlines=True) as p:
        for line in p.stdout:
            print(line, end='')

    if p.returncode != 0:
        raise subprocess.CalledProcessError(p.returncode, p.args)

if __name__ == "__main__":
    """
    execute the command to import all the required libraries from requirements.txt
    """
    command = ["install","-r", "requirements.txt"]
    send_cmd(command)