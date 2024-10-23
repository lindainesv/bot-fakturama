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

    for item in dados:

        # Clique sobre o botão de novo produto
        if not bot.find("New_product", matching=0.97, waiting_time=20000):
            not_found("New_product")
        bot.click()

        # Clique no campo de nome do produto
        if not bot.find("item_number", matching=0.97, waiting_time=10000):
            not_found("item_number")
        bot.click_relative(180, 7)

        # itemnr
        bot.paste(item[1])
        bot.tab()

        # name
        bot.paste(item[2])
        bot.tab()

        # category
        bot.paste(item[3])
        bot.tab()

        # GTIN
        bot.paste(item[4])
        bot.tab()

        # supcode
        bot.paste(item[5])
        bot.tab()

        # description
        bot.paste(item[6])
        bot.tab()
        # pdb.set_trace()

        # price
        bot.control_a()
        bot.paste(item[7])
        bot.tab()

        # costprice
        bot.paste(item[8])
        bot.tab()

        # allowance
        bot.paste(item[9])
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
        bot.paste(item[14])
        bot.tab()

        # select picture
        if not bot.find("select_picture", matching=0.97, waiting_time=10000):
            not_found("select_picture")
        bot.click()

        # path image
        bot.paste(
            rf"C:\Users\linda\OneDrive\Área de Trabalho\projeto_botcity_insercao_produtos\bot-fakturama\resources\produtos\imagens_produtos\{item[13]}")
        bot.enter()

        # save
        if not bot.find("save", matching=0.97, waiting_time=10000):
            not_found("save")
        bot.click()

        bot.control_w()

    bot.control_key("q")










    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
