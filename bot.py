"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/python-automations/desktop/
"""

# Import for the Desktop Bot
import pdb
import time
from botcity.core import DesktopBot
# Import BotCSVPlugin
from botcity.plugins.csv import BotCSVPlugin

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    # Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    planilha = BotCSVPlugin()

    dados = planilha.read(r"resources\produtos\items.csv").as_dataframe()

    # print(dados)
    # input()

    bot = DesktopBot()

    # Abre a aplicação Fakturama
    bot.execute(r"C:\Program Files\Fakturama2\Fakturama.exe")

    time.sleep(5)

    for item, row in dados.iterrows():

        # Clique sobre o botão de novo produto
        if not bot.find("New_product", matching=0.97, waiting_time=30000):
            not_found("New_product")
        bot.click()

        # Clique no campo de nome do produto
        if not bot.find("item_number", matching=0.97, waiting_time=100000):
            not_found("item_number")
        bot.click_relative(180, 7)

        # itemnr
        bot.paste(row['itemnr'])
        bot.tab()

        # name
        bot.paste(row['name'])
        bot.tab()

        # category
        bot.paste(row['category'])
        bot.tab()

        # GTIN
        bot.paste(row['GTIN'])
        bot.tab()

        # supcode
        bot.paste(row['supcode'])
        bot.tab()

        # description
        bot.paste(row['description'])
        bot.tab()
        # pdb.set_trace()

        # price
        bot.control_a()
        bot.paste(row['price'])
        bot.tab()

        # costprice
        bot.paste(row['costprice'])
        bot.tab()

        # allowance
        bot.paste(row['allowance'])
        bot.tab()

        # Scrool
        bot.scroll_down(1)

        # VAT
        if not bot.find("vat", matching=0.97, waiting_time=10000):
            not_found("vat")
        bot.click_relative(105, 8)
        bot.tab()

        # quantity
        bot.control_a()
        bot.paste(row['quantity'])
        bot.tab()

        # select picture
        if not bot.find("select_picture", matching=0.97, waiting_time=10000):
            not_found("select_picture")
        bot.click()

        nome_arquivo_imagem = str(row['picturename'])
        # print(nome_arquivo_imagem)
        # input()

        # path image
        if not (".jpg" in nome_arquivo_imagem or ".png" in nome_arquivo_imagem) or nome_arquivo_imagem == "":
            bot.paste(rf"C:\Users\linda\OneDrive\Área de Trabalho\projetos_botcity\bot-fakturama\resources\produtos\imagens_produtos\no_image.jpg")
            bot.enter()
            print("Arquivo sem imagem")
        else:
            bot.paste(rf"""C:\Users\linda\OneDrive\Área de Trabalho\projetos_botcity\bot-fakturama\resources\produtos\imagens_produtos\{nome_arquivo_imagem}""")
            bot.enter()
            print("Arquivo com imagem")

        # save
        if not bot.find("save", matching=0.97, waiting_time=10000):
            not_found("save")
        bot.click()

        bot.control_w()
        input()

    bot.control_key("q")

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
        task_id=execution.task_id,
        status=AutomationTaskFinishStatus.SUCCESS,
        message="Task Finished OK."
    )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
