import pandas as pd
import customtkinter as ctk
from tkcalendar import DateEntry
from openpyxl.workbook import Workbook

class FinanceiroApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle Financeiro da Família")
        self.root.geometry("1000x1000")

        # Estrutura para armazenar dados
        self.dados = {
            'Ganhos': [],
            'Gastos': [],
            'Investimentos': [],
            'Objetivos': []
        }

        # Criação do rótulo para o saldo
        self.saldo_label = ctk.CTkLabel(self.root, text="Saldo: R$0.00", font=("Helvetica", 23, "bold"))
        self.saldo_label.grid(row=34, column=3, columnspan=4, pady=(10, 20))

        # Rótulos para mostrar os totais
        self.total_ganhos_label = ctk.CTkLabel(self.root, text="Total de Ganhos: R$0.00", font=("Helvetica", 18))
        self.total_ganhos_label.grid(row=35, column=0, columnspan=2, pady=(10, 5))

        self.total_investimentos_label = ctk.CTkLabel(self.root, text="Total de Investimentos: R$0.00", font=("Helvetica", 18))
        self.total_investimentos_label.grid(row=36, column=0, columnspan=2, pady=(10, 5))

        self.total_gastos_label = ctk.CTkLabel(self.root, text="Total de Gastos: R$0.00", font=("Helvetica", 18))
        self.total_gastos_label.grid(row=37, column=0, columnspan=2, pady=(10, 5))

        self.valor_liquido_label = ctk.CTkLabel(self.root, text="Valor Líquido: R$0.00", font=("Helvetica", 18))
        self.valor_liquido_label.grid(row=38, column=0, columnspan=2, pady=(10, 5))

        # Rótulo para mensagem de sucesso
        self.mensagem_label = ctk.CTkLabel(self.root, text="", font=("Helvetica", 14, "italic"), text_color="green")
        self.mensagem_label.grid(row=39, column=0, columnspan=6, pady=(10, 5))

        self.criar_widgets()
        self.atualizar_saldo()

    def criar_widgets(self):
        self.criar_secao_data()  # Adiciona a seção de seleção de data
        self.criar_secao_ganhos()
        self.criar_secao_gastos()
        self.criar_secao_investimentos()
        self.criar_secao_objetivos()
        self.criar_botao_salvar()
        self.criar_botao_calcular()
        self.criar_botao_limpar()

    def criar_secao_data(self):
        ctk.CTkLabel(self.root, text="Selecione a Data", font=("Helvetica", 14)).grid(row=2, column=6, padx=7, pady=7)
        self.data_entry = DateEntry(self.root, font=("Helvetica", 12), width=12, background='darkblue', foreground='white', borderwidth=2)
        self.data_entry.grid(row=3, column=6, padx=5, pady=5)

    def criar_secao_ganhos(self):
        ctk.CTkLabel(self.root, text="Entrada ou Ganhos", font=("Helvetica", 18)).grid(row=2, column=0, columnspan=6, pady=(10, 5))

        self.ganho_entries = []
        for i in range(6):
            ctk.CTkLabel(self.root, text=f"Ganho {i + 1}", font=("Helvetica", 14)).grid(row=i + 3, column=0, padx=5, pady=5)
            descricao_entry = ctk.CTkEntry(self.root, placeholder_text="Descrição", font=("Helvetica", 14))
            descricao_entry.grid(row=i + 3, column=1, padx=5, pady=5)
            valor_entry = ctk.CTkEntry(self.root, placeholder_text="Valor Total", font=("Helvetica", 14))
            valor_entry.grid(row=i + 3, column=2, padx=5, pady=5)
            self.ganho_entries.append((descricao_entry, valor_entry))

    def criar_secao_gastos(self):
        ctk.CTkLabel(self.root, text="Saída ou Gastos", font=("Helvetica", 18)).grid(row=10, column=0, columnspan=4, pady=(10, 5))

        self.gasto_entries = []
        for i in range(6):
            ctk.CTkLabel(self.root, text=f"Gasto {i + 1}", font=("Helvetica", 14)).grid(row=i + 11, column=0, padx=5, pady=5)
            descricao_entry = ctk.CTkEntry(self.root, placeholder_text="Descrição", font=("Helvetica", 14))
            descricao_entry.grid(row=i + 11, column=1, padx=5, pady=5)
            valor_entry = ctk.CTkEntry(self.root, placeholder_text="Valor Total", font=("Helvetica", 14))
            valor_entry.grid(row=i + 11, column=2, padx=5, pady=5)
            self.gasto_entries.append((descricao_entry, valor_entry))

    def criar_secao_investimentos(self):
        ctk.CTkLabel(self.root, text="Investimentos", font=("Helvetica", 14)).grid(row=24, column=0, columnspan=2, pady=(10, 5))
        self.investimentos_entries = []
        for i, label in enumerate(["Renda Fixa", "Renda Variável", "Ações", "Fundos Imobiliários"], start=25):
            ctk.CTkLabel(self.root, text=label, font=("Helvetica", 14)).grid(row=i, column=0, padx=5, pady=5)
            entry = ctk.CTkEntry(self.root, placeholder_text="Valor", font=("Helvetica", 14))
            entry.grid(row=i, column=1, padx=5, pady=5)
            self.investimentos_entries.append(entry)

    def criar_secao_objetivos(self):
        ctk.CTkLabel(self.root, text="Objetivos Financeiros", font=("Helvetica", 14)).grid(row=29, column=0, columnspan=2, pady=(10, 5))
        ctk.CTkLabel(self.root, text="Objetivo a Curto Prazo", font=("Helvetica", 14)).grid(row=30, column=0, padx=5, pady=5)
        ctk.CTkEntry(self.root, placeholder_text="Descrição", font=("Helvetica", 14)).grid(row=30, column=1, padx=5, pady=5)

    def criar_botao_salvar(self):
        ctk.CTkButton(self.root, text="Salvar em Excel", command=self.salvar_excel, font=("Helvetica", 14)).grid(row=34, column=0, padx=5, pady=5)

    def criar_botao_calcular(self):
        ctk.CTkButton(self.root, text="Calcular Ganhos, Gastos e Investimentos", command=self.calcular_ganhos_e_gastos, font=("Helvetica", 14)).grid(row=34, column=1, padx=5, pady=5)

    def criar_botao_limpar(self):
        ctk.CTkButton(self.root, text="Limpar Campos", command=self.limpar_campos, font=("Helvetica", 14)).grid(row=34, column=2, padx=5, pady=5)

    def calcular_ganhos_e_gastos(self):
        self.adicionar_ganho()
        self.adicionar_gasto()
        self.adicionar_investimento()
        self.atualizar_saldo()

    def adicionar_ganho(self):
        for descricao_entry, valor_entry in self.ganho_entries:
            descricao = descricao_entry.get()
            valor_total = valor_entry.get()
            if descricao and valor_total:
                self.dados['Ganhos'].append({
                    'Descrição': descricao,
                    'Valor Total': float(valor_total),
                    'Data': self.data_entry.get_date()  # Adiciona a data
                })

    def adicionar_gasto(self):
        for descricao_entry, valor_entry in self.gasto_entries:
            descricao = descricao_entry.get()
            valor_total = valor_entry.get()
            if descricao and valor_total:
                self.dados['Gastos'].append({
                    'Descrição': descricao,
                    'Valor Total': float(valor_total),
                    'Data': self.data_entry.get_date()  # Adiciona a data
                })

    def adicionar_investimento(self):
        for entry in self.investimentos_entries:
            valor = entry.get()
            if valor:
                self.dados['Investimentos'].append({
                    'Valor': float(valor),
                    'Data': self.data_entry.get_date()  # Adiciona a data
                })

    def atualizar_saldo(self):
        total_ganhos = sum(ganho['Valor Total'] for ganho in self.dados['Ganhos'])
        total_investimentos = sum(investimento['Valor'] for investimento in self.dados['Investimentos'])
        total_gastos = sum(gasto['Valor Total'] for gasto in self.dados['Gastos'])

        valor_liquido = total_ganhos + total_investimentos - total_gastos

        self.saldo_label.configure(text=f"Saldo: R${valor_liquido:.2f}")
        self.total_ganhos_label.configure(text=f"Total de Ganhos: R${total_ganhos:.2f}")
        self.total_investimentos_label.configure(text=f"Total de Investimentos: R${total_investimentos:.2f}")
        self.total_gastos_label.configure(text=f"Total de Gastos: R${total_gastos:.2f}")
        self.valor_liquido_label.configure(text=f"Valor Líquido: R${valor_liquido:.2f}")

    def limpar_campos(self):
        for descricao_entry, valor_entry in self.ganho_entries:
            descricao_entry.delete(0, 'end')
            valor_entry.delete(0, 'end')
        for descricao_entry, valor_entry in self.gasto_entries:
            descricao_entry.delete(0, 'end')
            valor_entry.delete(0, 'end')
        for entry in self.investimentos_entries:
            entry.delete(0, 'end')

        # Não limpa os dados armazenados para manter o histórico
        self.atualizar_saldo()

    def salvar_excel(self):
        # Salva os dados em um arquivo Excel
        with pd.ExcelWriter('controle_financeiro.xlsx', engine='openpyxl') as writer:
            for categoria, dados in self.dados.items():
                if dados:  # Apenas cria uma aba se houver dados
                    df = pd.DataFrame(dados)
                    df.to_excel(writer, sheet_name=categoria, index=False)

        # Exibe mensagem de sucesso
        self.mensagem_label.configure(text="Dados salvos com sucesso!")

if __name__ == "__main__":
    root = ctk.CTk()
    app = FinanceiroApp(root)
    root.mainloop()