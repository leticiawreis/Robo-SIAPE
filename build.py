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

# Pasta com os prints usados no diálogo de ajuda passo a passo. Se existir,
# é embutida dentro do próprio .exe (via --add-data do PyInstaller); se não
# existir, o build segue normalmente e o diálogo de ajuda apenas mostra um
# aviso de "imagem não encontrada" no lugar dos prints.
PASTA_ASSETS_AJUDA = os.path.join("assets", "ajuda")


def limpar_builds_anteriores():
    """Remove os artefatos deixados por builds anteriores do
    PyInstaller, para garantir que o novo build seja gerado do zero.

    Remove as pastas "build" e "dist" (se existirem) e o arquivo de
    especificação ".spec" gerado pelo PyInstaller (ex:
    "RoboSIAPE.spec"), evitando que configurações ou arquivos
    desatualizados de uma execução anterior interfiram no novo build.
    """
    for pasta in ("build", "dist"):
        if os.path.isdir(pasta):
            shutil.rmtree(pasta)

    arquivo_spec = f"{NOME_EXECUTAVEL}.spec"
    if os.path.exists(arquivo_spec):
        os.remove(arquivo_spec)


def gerar_executavel():
    """Invoca o PyInstaller para empacotar o script principal
    (ARQUIVO_PRINCIPAL) em um único executável.

    Antes de chamar o PyInstaller, verifica se o arquivo principal
    existe na pasta atual; se não existir, encerra o programa com uma
    mensagem de erro orientando a ajustar a variável ARQUIVO_PRINCIPAL.

    Executa o PyInstaller com os seguintes parâmetros:
    - nome do executável definido por NOME_EXECUTAVEL;
    - "--onefile": empacota tudo em um único arquivo executável;
    - "--windowed": não abre um console/terminal junto com a interface
      gráfica (aplicação com janela própria);
    - "--noconfirm": sobrescreve arquivos de saída existentes sem pedir
      confirmação;
    - "--add-data": embute a pasta assets/ajuda/ (prints do diálogo de
      ajuda) dentro do próprio executável, se ela existir.

    Levanta:
        SystemExit: se ARQUIVO_PRINCIPAL não for encontrado na pasta
            atual (via sys.exit com mensagem explicativa).
    """
    if not os.path.exists(ARQUIVO_PRINCIPAL):
        sys.exit(
            f"Não encontrei '{ARQUIVO_PRINCIPAL}' nesta pasta. "
            "Ajuste a variável ARQUIVO_PRINCIPAL no topo deste script."
        )

    argumentos = [
        ARQUIVO_PRINCIPAL,
        "--name", NOME_EXECUTAVEL,
        "--onefile",
        "--windowed",
        "--noconfirm",
    ]

    if os.path.isdir(PASTA_ASSETS_AJUDA):
        # Sintaxe do --add-data é "ORIGEM<separador>DESTINO", e o
        # separador muda por sistema operacional: ";" no Windows, ":" no
        # Linux/Mac. os.pathsep já resolve isso automaticamente.
        argumentos.append(
            f"--add-data={PASTA_ASSETS_AJUDA}{os.pathsep}{PASTA_ASSETS_AJUDA}"
        )
    else:
        print(
            f"Aviso: pasta '{PASTA_ASSETS_AJUDA}' não encontrada — o .exe "
            "será gerado sem os prints do diálogo de ajuda."
        )

    PyInstaller.__main__.run(argumentos)


def main():
    """Ponto de entrada do script de build.

    Orquestra o processo completo: limpa builds anteriores
    (limpar_builds_anteriores), gera o novo executável
    (gerar_executavel) e, ao final, exibe uma mensagem confirmando o
    caminho do executável gerado (dist/RoboSIAPE.exe) — ou uma mensagem
    genérica apontando para a pasta "dist", caso o caminho esperado não
    seja encontrado (por exemplo, ao rodar em outro sistema operacional).
    """
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