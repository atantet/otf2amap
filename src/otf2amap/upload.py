"""Envoi de la sortie vers un dossier distant via rclone (ex. Proton Drive).

Configuration (section [output] de config.toml) :
  drive_remote = "protondrive:AMAP/feuilles"   # cible rclone (remote:chemin)

Le remote (ici « protondrive ») se configure une fois pour toutes avec :
  rclone config        # créer un remote Proton Drive nommé « protondrive »
"""

import shutil
import subprocess
from pathlib import Path


def upload_to_remote(local_path, remote):
    """Copie `local_path` vers `remote` (cible rclone, ex. "protondrive:AMAP/feuilles").

    Le fichier local reste en place. En cas d'absence de rclone ou d'échec de
    l'envoi, on prévient sans planter : la feuille locale est toujours produite.
    Retourne True si l'envoi a réussi, False sinon.
    """
    local_path = Path(local_path)
    if shutil.which('rclone') is None:
        print("AVERTISSEMENT : rclone introuvable, envoi Drive ignoré.\n"
              "  Installez rclone puis configurez un remote (rclone config).")
        return False

    remote_name, _, remote_path = remote.partition(':')

    print(f"Envoi vers {remote_name} : {remote_path} …")
    try:
        subprocess.run(['rclone', 'copy', str(local_path), remote], check=True)
    except subprocess.CalledProcessError as e:
        print(f"AVERTISSEMENT : échec de l'envoi rclone (code {e.returncode}).\n"
              f"  Le fichier local est conservé : {local_path}")
        return False

    print(f"Envoyé sur {remote_name} : {remote_path}/{local_path.name}")
    return True
