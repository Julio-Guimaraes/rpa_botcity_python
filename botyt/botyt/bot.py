from botcity.core import DesktopBot
# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
from botcity.maestro import *
import pandas as pd

class Bot(DesktopBot):
    def action(self, execution=None):

        # Opens the Whatsapp website.
        self.browse("http://web.whatsapp.com")

        tabela = pd.read_excel('Contatos.xlsx')
        for linha in tabela.index:
            contato = tabela.loc[linha, "Contato"]
            msg = tabela.loc[linha, "Msg"]
            arquivo = tabela.loc[linha, "Arquivo"]


            if not self.find( "lupa", matching=0.97, waiting_time=10000):
                self.not_found("lupa")
            self.click()
            self.type_keys_with_interval(100, "Para nao panguar")
            self.enter()
            if not self.find( "clipper", matching=0.97, waiting_time=10000):
                self.not_found("clipper")
            self.click()
            if not self.find( "arq", matching=0.97, waiting_time=10000):
                self.not_found("arq")
            self.click()
            self.paste()
            if pd.isna(arquivo):
                self.paste(msg)
                self.enter()

            if self.find( "seta", matching=0.97, waiting_time=10000):
                self.click()

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()















