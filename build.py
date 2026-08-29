"""
Gera o executável (.exe) da interface do Robô SIAPE.

Uso:
    python build.py

Requisitos (instale antes, se ainda não tiver):
    pip install -r requirements.txt

O resultado fica em dist/RoboSIAPE.exe (Windows) ou dist/RoboSIAPE
(Linux/Mac, caso rode nesses sistemas).
"""

import os
import shutil
import sys

import PyInstaller.__main__


ARQUIVO_PRINCIPAL = "interface.py"  # troque para "interface.py" se você renomeou o arquivo
NOME_EXECUTAVEL = "RoboSIAPE"


def limpar_builds_anteriores():
    for pasta in ("build", "dist"):
        if os.path.isdir(pasta):
            shutil.rmtree(pasta)

    arquivo_spec = f"{NOME_EXECUTAVEL}.spec"
    if os.path.exists(arquivo_spec):
        os.remove(arquivo_spec)


def gerar_executavel():
    if not os.path.exists(ARQUIVO_PRINCIPAL):
        sys.exit(
            f"Não encontrei '{ARQUIVO_PRINCIPAL}' nesta pasta. "
            "Ajuste a variável ARQUIVO_PRINCIPAL no topo deste script."
        )

    PyInstaller.__main__.run(
        [
            ARQUIVO_PRINCIPAL,
            "--name", NOME_EXECUTAVEL,
            "--onefile",
            "--windowed",
            "--noconfirm",
        ]
    )


def main():
    print("Limpando builds anteriores...")
    limpar_builds_anteriores()

    print(f"Gerando o executável a partir de '{ARQUIVO_PRINCIPAL}'...")
    gerar_executavel()

    caminho_exe = os.path.join("dist", f"{NOME_EXECUTAVEL}.exe")
    print()
    print("=" * 70)
    if os.path.exists(caminho_exe):
        print(f"Pronto! Executável gerado em: {caminho_exe}")
    else:
        print(f"Build concluído. Veja a pasta 'dist' para o executável gerado.")
    print("=" * 70)


if __name__ == "__main__":
    main()