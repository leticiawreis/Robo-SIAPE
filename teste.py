"""
Testa o robô automaticamente para TODAS as combinações de ano/mês
conhecidas em INDICES_ANOS, uma por uma, e gera um resumo no final.

Uso:
    python testar_todos_meses.py

Coloque este arquivo na raiz do projeto (mesma pasta de interface.py e
robo_siape.py) antes de rodar.

ATENÇÃO:
- Isso baixa de verdade cada pacote disponível — pode demorar bastante
  e consumir banda, dependendo de quantos anos/meses existirem.
- Cada execução bem-sucedida gera uma pasta em saida/, como qualquer
  execução normal pela interface. Se quiser, apague saida/ depois.
- Se quiser testar só um intervalo (ex.: só 2024 e 2025), ajuste a
  lista ANOS_PARA_TESTAR logo abaixo.
"""

import queue
import time
from datetime import datetime

from robo_siape import (
    INDICES_ANOS,
    MESES_OPCOES,
    meses_disponiveis,
    validar_periodo,
    executar_pipeline_completo,
)

TIPO_PACOTE = "Servidores_SIAPE"

# Por padrão testa todos os anos conhecidos. Se quiser restringir,
# troque para algo como: ANOS_PARA_TESTAR = ["2024", "2025", "2026"]
ANOS_PARA_TESTAR = list(INDICES_ANOS.keys())

# Pausa entre uma execução e outra, pra não martelar o servidor.
PAUSA_ENTRE_TESTES_SEGUNDOS = 3


def testar_uma_competencia(ano, mes):
    """Roda o pipeline completo para um ano/mês e devolve o resultado.

    Retorna uma tupla (status, detalhe), onde status é uma das
    strings: "sucesso", "invalido", "nao_encontrado", "erro".
    """
    try:
        validar_periodo(ano, mes, TIPO_PACOTE)
    except ValueError as erro:
        return "invalido", str(erro)

    log_queue = queue.Queue()

    try:
        caminho_saida, caminho_log = executar_pipeline_completo(
            ano, mes, TIPO_PACOTE, log_queue
        )
        return "sucesso", caminho_saida

    except FileNotFoundError as erro:
        # Pacote não publicado para essa competência (404) ou CSV não
        # encontrado dentro do ZIP.
        return "nao_encontrado", str(erro)

    except Exception as erro:
        return "erro", str(erro)


def main():
    inicio_geral = time.time()
    resultados = []

    total_combinacoes = sum(
        len(meses_disponiveis(ano)) for ano in ANOS_PARA_TESTAR
    )
    contador = 0

    print("=" * 70)
    print(f"Iniciando teste de {total_combinacoes} combinações de ano/mês.")
    print("=" * 70)

    for ano in ANOS_PARA_TESTAR:
        for mes in meses_disponiveis(ano):
            contador += 1
            print(
                f"\n[{contador}/{total_combinacoes}] Testando {mes}/{ano}..."
            )

            inicio = time.time()
            status, detalhe = testar_uma_competencia(ano, mes)
            duracao = time.time() - inicio

            resultados.append(
                {
                    "ano": ano,
                    "mes": mes,
                    "status": status,
                    "detalhe": detalhe,
                    "duracao_segundos": round(duracao, 1),
                }
            )

            simbolo = {
                "sucesso": "OK",
                "invalido": "INVALIDO",
                "nao_encontrado": "SEM DADO",
                "erro": "ERRO",
            }[status]

            print(f"    -> {simbolo} ({duracao:.1f}s)")

            time.sleep(PAUSA_ENTRE_TESTES_SEGUNDOS)

    duracao_total = time.time() - inicio_geral

    # ------------------------------------------------------------
    # Resumo final
    # ------------------------------------------------------------
    print("\n" + "=" * 70)
    print("RESUMO FINAL")
    print("=" * 70)

    for status_alvo, rotulo in (
        ("sucesso", "Sucesso"),
        ("nao_encontrado", "Sem dado publicado"),
        ("invalido", "Período inválido"),
        ("erro", "Erro inesperado"),
    ):
        itens = [r for r in resultados if r["status"] == status_alvo]
        print(f"\n{rotulo}: {len(itens)}")
        for item in itens:
            print(f"  - {item['mes']}/{item['ano']}  ({item['duracao_segundos']}s)")

    print(f"\nTempo total: {duracao_total / 60:.1f} minutos")

    # Salva um arquivo de log com todos os detalhes, inclusive erros.
    nome_arquivo = f"teste_todos_meses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
        for item in resultados:
            arquivo.write(
                f"{item['ano']}-{MESES_OPCOES[item['mes']]:02d} "
                f"({item['mes']}) | {item['status']} | "
                f"{item['duracao_segundos']}s | {item['detalhe']}\n"
            )

    print(f"\nDetalhes completos salvos em: {nome_arquivo}")


if __name__ == "__main__":
    main()