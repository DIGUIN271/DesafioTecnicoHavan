from botcity.web import WebBot, Browser, By
import pandas
import logging


# Importando integração do BotCity Maestro SDK
from botcity.maestro import *

# Configuração do logging
logging.basicConfig(
    filename="bot_log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)

# Desabilitar erros se não connectado ao Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    logging.info("Iniciando execução do bot")

    # Configurações botcity
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()
    logging.info(f"ID da tarefa é: {execution.task_id}")
    logging.info(f"Parametros da tarefa são: {execution.parameters}")

    bot = WebBot()

    bot.headless = False
    # Forma oculta: False (desligado) True (ligado)

    bot.browser = Browser.EDGE
    bot.driver_path = "C:\edgedriver_win32\msedgedriver.exe"

    bot.browse("https://www.rpachallenge.com/")
    bot.wait(2000)
    logging.info("Site aberto: https://www.rpachallenge.com/")

    # Baixar e utilizar planilha
    desafio = 0
    if desafio <= 1:
        baixar_planilha = bot.find_element(
            "//a[contains(normalize-space(.), 'Download Excel')]", By.XPATH
        )
        baixar_planilha.click()
        bot.wait(5000)

        logging.info("Planilha baixada.")

    # Ler e absorver dados da planilha
    dados = pandas.read_excel(
        r"C:\Users\Digui\OneDrive\Documentos\BotCity Projects\BotHavan\challenge.xlsx"
    )
    logging.info(f"Planilha carregada com {len(dados)} linhas")

    # Iniciando desafio
    iniciar_desafio = bot.find_element(
        "//button[normalize-space(text())='Start']", By.XPATH
    )
    iniciar_desafio.click()
    logging.info("Desafio iniciado")

    desafio += 1

    # Preenchimento de informações e entrega
    for index, row in dados.iterrows():
        try:
            Campo_Endereço = bot.find_element(
                "//input[@ng-reflect-name='labelAddress']", By.XPATH
            )
            Campo_Endereço.send_keys(row["Address"])
            logging.info(
                f"Linha {index}: Endereco preenchido -> {row['Address']}")

            Campo_Companhia = bot.find_element(
                "//input[@ng-reflect-name='labelCompanyName']", By.XPATH
            )
            Campo_Companhia.send_keys(row["Company Name"])
            logging.info(
                f"Linha {index}: Companhia preenchida -> {row['Company Name']}"
            )

            Campo_Função = bot.find_element(
                "//input[@ng-reflect-name='labelRole']", By.XPATH
            )
            Campo_Função.send_keys(row["Role in Company"])
            logging.info(
                f"Linha {index}: Funcao preenchida -> {row['Role in Company']}"
            )

            Campo_Email = bot.find_element(
                "//input[@ng-reflect-name='labelEmail']", By.XPATH
            )
            Campo_Email.send_keys(row["Email"])
            logging.info(f"Linha {index}: Email preenchido -> {row['Email']}")

            Campo_Primeiro_Nome = bot.find_element(
                "//input[@ng-reflect-name='labelFirstName']", By.XPATH
            )
            Campo_Primeiro_Nome.send_keys(row["First Name"])
            logging.info(
                f"Linha {index}: Primeiro Nome preenchido -> {row['First Name']}"
            )

            Campo_Ultimo_Nome = bot.find_element(
                "//input[@ng-reflect-name='labelLastName']", By.XPATH
            )
            Campo_Ultimo_Nome.send_keys(row["Last Name "])
            logging.info(
                f"Linha {index}: Ultimo Nome preenchido -> {row['Last Name ']}"
            )

            Campo_Telefone = bot.find_element(
                "//input[@ng-reflect-name='labelPhone']", By.XPATH
            )
            Campo_Telefone.send_keys(row["Phone Number"])
            logging.info(
                f"Linha {index}: Telefone preenchido -> {row['Phone Number']}")

            Entregar = bot.find_element(
                "//input[@type='submit' and @value='Submit']", By.XPATH
            )
            Entregar.click()
            logging.info(f"Linha {index}: Formulario enviado")

        except Exception as e:
            logging.error(
                f"Linha {index}: Erro ao preencher formulário -> {e}")

    bot.wait(2000)

    # Finalização Navegador
    try:
        logging.info("Fechando navegador...")
        bot.stop_browser()
        logging.info("Navegador fechado com sucesso")
    except Exception as e:
        logging.error(f"Erro ao fechar navegador normalmente: {e}")
        logging.info("Usando metodo alternativo para fechar...")
        try:
            bot.kill_browser()
            logging.info("Navegador fechado com metodo alternativo")
        except Exception as e2:
            logging.error(
                f"Navegador ja fechado ou em estado inconsistente: {e2}")


def not_found(label):
    logging.warning(f"Elemento nao entregue: {label}")


if __name__ == "__main__":
    main()
