import os
from tkinter import filedialog as fd
from tkinter import messagebox
import PySimpleGUI as sg

lista = set()

icon = b'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFgklEQVR4Xr2XbUxTVxjHn75CKwKU0haFAoIahakbfjQDiduyN0fQuRCBiFOXDCNx37ZkWzY3kyWbW8TM4HCbijg/qEXl1bcWhlQFKYKIgKW0WhChbAAtvbft2XMPu6FBjSsO/+2Tk3Nze57ffc7/nHsqIITAk3Tl0kXF8uSUAoZltrmcTi3r8YDbzYDP62212azFWRs2FuNvffCcEsIM7S7cLRp2DBcujImts1qtX/Tb7dqxsXEYx0Ba8HjYlUpl1L46g6H+RNnx9+F5xVWAjwcP7sc9fDhw3djYSLi4b7MRt9tN/OV0Oun15uYmcrXhT1JdVXkSAILpGIHHNMDAQH+c3W6/r798mfRZLIQXy7KEYRgaLMvQ/pR8xGw2k2vXjKS2pkYPANLZANAp0F+6KBYIBLquzs6FK1atBG1cHDCMG4OhcLx4YIShkZCQAFqtFsLDw9POnS0vnrUHVq1OLbD0WlbF44BhYeEwOekCPu80gH9/it7lcoJarYZIpRIiI5VbDh44kBEwQHtHm2J4yFHg8/kgJjaWDson4pNNi/hXggJNTDghWqOBkPkhkLRk8ecBA2iiNIUjDkeSSq3CJ5+kSRHGzyjA97Hlwuffpy1+aOXmzZv36q6dBfEQgESFhbtOjIyMzNdoogEE9IsSAHri31L7rxa+Cv4wXlyaXgTxAeNmBF6f9/orqant8B8lnnRNatDQIBQKuU2GAuDgCMCXWTCD+XFPeL1eCieRSiFSEZkIAUjMlY+lbvcBt9sR91QCkVAEQpGQgvmLnyJMSlsUwosQ3oOtAKRBUnlAACzDgkgshrGxMTq4SCSiwYlMuZ1OB5+cB+AhPAjNe4KDHv17dCQgE+KctQEKAfwHpsEn4AP7M9vp+z0Ig1XAXfJOQAD4gx8lEgmLLxzOejjQY4mfCkEBvFPQMpkMnOMTznPluqsBAbz51ttHsW3inOyjPmAR4NkQfLAsS6dILBEDvrxO1zUaHQEBEEK81r6+b6XoYIfDAVKJFDhf0EQeCvHUwPuoSRUKBQw9GvpLp9N9BQGKPw/AxQu1paGhYZsl+CTzQ0PB5XTRwdGANHj5mxAdDxERCuCmr1x35njh7k9yZn0eeO31N/JHR0druKkYR0MGy4JBLBbTRFhmPriy01USEhICUVFRWCUP/LTvB7hmNGbduNG0aNYVoB0UPsl+tVqzUy6XQ1BwEF1aUwsSuIpQqKCgIOA2r4H+/nGjsbEcr2+orqoIDg0Lt6atzdiTn7+1ZFYHEj72frNnTU1V5ZmG+voJU0sL6ey8Q3q6u0nX3bukvb0NDyINQydPlB1al562mLu/prp6Semxo49ysjeRrfl55EDR/hqDwRARwIHkybFubXp08cGf3ykrLd115tSpz349XPLxd3v3ZgBA+EzoujrDR9WVFWRj5nqSvyWXFBXtb62vMyifAcCb8LlFjdp4taF+cHBwzW+HfwFFpBJeTl3dnrw8OWNtRsajZ5nwfzlb9vT0bEL/mLds3Qb4igfTzeaU2x23r9TW1kbNNQBVTm5eP3okXaVWm7M358DgwACYWpqTu7ruGtAnqrkGoPpw23ZbR9vtdJVKZd6UnQ0P7XZobbm5rKenW48naPWcA1CIHdttVqstPVYbb/4gezMuV4QwmZbdu3fPUHH+vGauAahy8/JsSJEevWCB+d31mdBnMcOtVtPSXkuvvuL8ueg5BuA9kYsM1rS4+ATze5lZYOk1Q9ut1qUWi0VfVYnTwe+Ec61jvx+JUUer9d1dXYnlutOwKHExJKe8ZJLLZTsowIvQ8WNHY0LDwvR3OjoSL9RWI0QSDQrwovRHWdkCmVx2Fg2Z2nKzqX9ZckoRBXiRKjl0aAX+82oeHhr69Muv93z/Dz04n+35GxpAAAAAAElFTkSuQmCCMTQ2Nw=='

def LayoutJanela():
    sg.theme('DarkGrey14')
    layout = [
        [sg.Text("Buscar por", font=('Cambria'), pad=(105, 5))],
        [sg.Checkbox(text="CNPJ", pad=(45, 10), key="cnpj", default=True, font=('Cambria')), sg.Checkbox(text="CPF", pad=(45, 10), key="cpf", default=True)],
        [sg.Button('Ler arquivos', pad=(100,20), font=('Cambria'))],
        [sg.Output(size=(40,5))],
    ]
    return sg.Window('Leitor de XML', icon=icon, layout=layout, margins=(0, 0), finalize=True)

janela = LayoutJanela()

def read_files():
    caminho = fd.askdirectory(title="Selecione a pasta dos XMLs")
    if caminho != "":
        lista.clear()
        os.system('cls')
        arquivos = os.listdir(caminho)
        for i in range(len(arquivos)):
            if arquivos[i].endswith('.xml'):
                with open(caminho+"/"+arquivos[i], 'r')as file:
                    ler = file.readlines()
                    string = str(ler)
                    modeNota = string.split("mod")
                    modNota = int(modeNota[1].replace('</', '').replace('>', ''))
                    if modNota == 55:
                        identifid = string.split("<CNPJ>")
                        nome = string.split("xNome")
                        if len(identifid) > 2:
                            if(verificao_cnpj() == True):
                                cnpj = identifid[2][0:14]
                                cnpj_format = format_cnpj(cnpj)
                                nome = nome[3].replace('&', 'E')
                                lista.add(cnpj_format+" ; "+nome.replace('>', '').replace('</','').replace(';', ''))
                        else:
                            if(verificao_cpf() == True):
                                sep = string.split("<CPF>")  
                                cpf = sep[1][0:13].replace('</', '')
                                cpf_format = format_cpf(cpf)
                                nome = nome[3]
                                lista.add(cpf_format+" ; "+nome.replace('>', '').replace('</','').replace(';', ''))

        return lista
    else:
        return None

def create_path():
    files = [('Valores Separados Por Virgula', '*.csv')]
    pasta_final = fd.asksaveasfile(title="Selecione onde salvar a tabela", defaultextension=".csv", filetypes = files, mode="w")
    
    if pasta_final != None:
            pasta_final.write("CNPJ / CPF;Nome Empresa\n")
            for item in lista:
                pasta_final.write(item+"\n")
            pasta_final.close()
            return pasta_final
    else:
        return None
    
def format_cnpj(cnpj):
    cnpj_format = f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:14]}"
    return cnpj_format

def format_cpf(cpf):
    cpf_format = f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:11]}"
    return cpf_format

def verificao_cnpj():
    if values['cnpj'] == True:
        return True
    else:
        return False

def verificao_cpf():
    if values['cpf'] == True:
        return True
    else:
        return False

def init():
    ler = read_files()
    if ler != None:
        messagebox.showinfo("Sucesso", "Arquivos lidos com sucesso, selecione onde salvar o .CSV")
    else:
        messagebox.showerror("Erro", "Não foi possível ler os arquivos XMLs")

    escrever = create_path()
    if escrever != None:
        messagebox.showinfo("Sucesso", "Arquivo salvo")
    else:
        messagebox.showerror("Erro", "Selecione a pasta para salvar o arquivo")

    for item in lista:
            print(item)


while True:
    event, values = janela.read()
    if event == sg.WIN_CLOSED:
        break
    if event == 'Ler arquivos':
        init()
