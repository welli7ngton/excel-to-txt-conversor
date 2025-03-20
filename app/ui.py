import tkinter as tk
from tkinter import filedialog, messagebox


class PlanilhaConverterUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ACCORD - Conversor de Planilha para TXT")
        self.root.geometry("450x250")

        self.entrada_var = tk.StringVar()
        self.planilha_var = tk.StringVar()

        frame = tk.Frame(root)
        frame.pack(pady=20)

        tk.Label(frame, text="Selecione o arquivo Excel:").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(frame, textvariable=self.entrada_var, width=40).grid(row=1, column=0, padx=5, pady=5)
        tk.Button(frame, text="Procurar", command=self.selecionar_arquivo).grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame, text="Nome da Planilha:").grid(row=2, column=0, padx=5, pady=5)
        tk.Entry(frame, textvariable=self.planilha_var, width=20).grid(row=3, column=0, padx=5, pady=5)

        tk.Button(root, text="Gerar TXT", command=self.gerar_txt, width=20, height=2).pack(pady=10)

    def selecionar_arquivo(self):
        caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xls*"), ("Todos", "*.*")])
        if caminho_arquivo:
            self.entrada_var.set(caminho_arquivo)

    def gerar_txt(self):
        caminho_entrada = self.entrada_var.get()
        nome_planilha = self.planilha_var.get()

        if not caminho_entrada:
            messagebox.showerror("Erro", "Por favor, selecione um arquivo Excel!")
            return

        if not nome_planilha:
            messagebox.showerror("Erro", "Por favor, insira o nome da planilha!")
            return

        caminho_saida = filedialog.asksaveasfilename()

        if caminho_saida:
            from conversor import processar_planilha
            try:
                processar_planilha(caminho_entrada, nome_planilha, caminho_saida)
                messagebox.showinfo("Sucesso", f"Arquivo gerado com sucesso: {caminho_saida}")
            except Exception as e:
                messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaConverterUI(root)
    root.mainloop()
