from io import StringIO
import os
import pandas as pd
from playwright.sync_api import sync_playwright
from pathlib import Path
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar
from dotenv import load_dotenv
from tkinter import messagebox

load_dotenv()
login = os.getenv("login")
senha = os.getenv("senha")
pagina = os.getenv("pagina")
todos_os_dados = []

def select_date():
    def confirm():
        selected_dates["inicio"] = cal_date.get_date()
        selected_dates["fim"] = cal_date_end.get_date()
        root.destroy()

    selected_dates = {"inicio": None, "fim": None}

    root = tk.Tk()
    root.geometry("600x600")
    root.title("Select dates:")

    tk.Label(root, text="Início do período que deseja extrair:").pack(pady=5)
    cal_date = Calendar(root, date_pattern="dd/mm/yyyy")
    cal_date.pack(pady=10)

    tk.Label(root, text="Final do período que deseja extrair:").pack(pady=5)
    cal_date_end = Calendar(root, date_pattern="dd/mm/yyyy")
    cal_date_end.pack(pady=10)

    ttk.Button(root, text="Confirmar", command=confirm).pack(pady=20)
    root.mainloop()

    return selected_dates["inicio"], selected_dates["fim"]

date, date_end = select_date()
if not date or not date_end:
    exit()

datainicio = (f'{date}')
datafim = (f'{date_end}')

nomes_lojas = {
    2: "01 - Atacadaço",
    3: "02 - Rodo",
    4: "03 - Tupy",
    5: '04 - Gomes',
    6: '05 - Neto',
    7: '06 - SJ',
    8: '07 - Monsenhor',
    9: '08 - DP',
    10: '09 - Fragata',
    11: '10 - F. Osório',
    12: '11 - Liv. Atacado',
    13: '12 - Liv. Varejo',
    14: '13 - Quaraí',
    15: '14 - CAÇAPAVA',
    16: '15 - São Gabriel',
    17: '16 - Rosário',
    18: '17 - Atacadaço São Gabriel',
}
def acessar_e_logar():
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=False)  # Define se quer ver a janela aberta ou não
            page = browser.new_page()

            loja_atual = 2
            ultima_loja = 18

            page.goto(pagina)
            page.fill('//*[@id="usuario"]', login)
            page.fill('//*[@id="senha"]', senha)
            page.click('//*[@id="botao"]')
            page.wait_for_timeout(1000)

            while loja_atual <= ultima_loja:
                try:
                    print(f"Processando Loja {loja_atual-1}...")

                    frame_tree = page.frame(name="treeframe")
                    if not frame_tree:
                        raise Exception("Frame 'treeframe' não encontrado!")
                    
                    frame_tree.click('//*[@id="f"]/table/tbody/tr/td/table[3]/tbody/tr/td')
                    page.wait_for_load_state("load")
                    frame_tree.click('text=- Cancelamentos')
                    page.wait_for_load_state("load")
                    base_frame2 = page.frame(name="topFrame")
                    base_frame = page.frame(name="basefrm")
                    base_frame.fill('//*[@id="dataproc_i"]', datainicio)
                    base_frame.fill('//*[@id="dataproc_f"]', datafim)
                    base_frame.click('//*[@id="btn_html"]')
                    base_frame.wait_for_selector('body > table', timeout=100000)

                    tabela_html = base_frame.evaluate('''() => {
                        const tabela = document.querySelector('body > table');
                        return tabela ? tabela.outerHTML : 'Tabela não encontrada';
                    }''')
                    if "Tabela não encontrada" in tabela_html:
                        print(f"Tabela não encontrada para Loja {loja_atual}.")
                    else:

                        dfs = pd.read_html(StringIO(tabela_html))
                        df = next(df for df in dfs if not df.empty)
                        df = df.iloc[1:-2]

                        nome_loja = nomes_lojas.get(loja_atual, f"Loja {loja_atual}")
                        df.insert(0, "Loja", nome_loja)
                        df.insert(1, "Coluna B", "")
                        df.insert(2, "Coluna C", "")
                        todos_os_dados.append(df)
                    
                    base_frame2.select_option('//*[@id="lojas"]', str(loja_atual))
                    page.wait_for_timeout(2000)

                except Exception as e:
                    print(f"Erro ao processar Loja {loja_atual}: {e}")
                finally:
                    loja_atual += 1
        try:
            downloads_path = str(Path.home() / "Downloads")
            
            if todos_os_dados:
                final_df = pd.concat(todos_os_dados, ignore_index=True)
                
                for col in final_df.select_dtypes(include=["object"]).columns:
                    if col != final_df.columns[0]:
                        final_df[col] = final_df[col].map(lambda x: str(x).replace('.', ',') if isinstance(x, str) else x)

                
                nome_arquivo = f"Relatorio_completo_{datainicio.replace('/', '-')}_a_{datafim.replace('/', '-')}.xlsx"
                caminho_completo = os.path.join(downloads_path, nome_arquivo)
                
                final_df.to_excel(caminho_completo, index=False)
                messagebox.showinfo("Aviso", f"Processo concluído com sucesso! '{caminho_completo}'")
            else:
                print("Nenhuma tabela foi encontrada para as lojas processadas.")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

acessar_e_logar()
