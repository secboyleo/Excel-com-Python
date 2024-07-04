# Excel-com-Python
## Codigo em Python que permite inserir dados em uma planilha de maneira interativa.
```
import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl, xlrd
import pathlib
from openpyxl import Workbook

#Setando a aparencia padrao do sistema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        #para deixar a classe como principal do sistema
        super().__init__()
        self.layout_config()
        self.appearence()
        self.todo_sistema()
        
    def layout_config(self):
        self.title("Sistema de Manutenção TI")
        self.geometry("800x500")
    
    def appearence(self):
        self.lb_apm = ctk.CTkLabel(self, text="Tema", bg_color="transparent", text_color=['#000', '#fff']).place(x=50, y=430)
        
        self.opt_apm = ctk.CTkOptionMenu(self, values=["Light", "Dark", "System"], command=self.change_apm).place(x=50, y=460)
        
    def change_apm(self, new_apm):
        ctk.set_appearance_mode(new_apm)
        
    def todo_sistema(self):
        frame = ctk.CTkFrame(self, width=800, height=50, corner_radius=0, bg_color="teal", fg_color="teal")
        frame.place(x=0, y=10)
        
        title = ctk.CTkLabel(frame, text="Sistema de Manutenção TI", font=("Century Gothic bold", 24), text_color="#fff").place(x=250, y=10)
        
        span = ctk.CTkLabel(self, text="Por favor preencha todos os campos do formulário", font=("Century Gothic bold", 24), text_color=["#000", "#fff"]).place(x=50, y=70)
        
        #Funcoes--------------------------------------------------------------------
        def submit():
            ficheiro = pathlib.Path("manutencao.xlsx")
            if ficheiro.exists():
                pass
            else:
                ficheiro=Workbook()
                folha = ficheiro.active
                folha['A1']="Chegada"
                folha['B1']='Saida'
                folha['C1']='Origem'
                folha['D1']='Problema'
                folha['E1']='Solução'
                folha['F1']='Tecnico'
                folha['G1']='WKS'
                folha['H1']='Patrimonio'
                folha['I1']='Lacre'
                
                ficheiro.save('manutencao.xlsx')
                
            #pegando os dados das entradas
            chegada = chegada_value.get()
            saida = saida_value.get()
            origem = origem_value.get()
            problema = problema_value.get()
            solucao = solucao_value.get()
            tecnico = tecnico_value.get()
            wks = wks_value.get()
            patrimonio = patrimonio_value.get()
            lacre = lacre_value.get()
            
            if problema == "" or tecnico == "":
                messagebox.showerror("Sistema", 'ERRO!\nPor favor preencha os campos "Problema" e "Tecnico"')
            else:
                #criando planilha
                ficheiro = openpyxl.load_workbook("manutencao.xlsx")
                folha = ficheiro.active
                
                folha.cell(column=1, row=folha.max_row+1, value=chegada)
                folha.cell(column=2, row=folha.max_row, value=saida)
                folha.cell(column=3, row=folha.max_row, value=origem)
                folha.cell(column=4, row=folha.max_row, value=problema)
                folha.cell(column=5, row=folha.max_row, value=solucao)
                folha.cell(column=6, row=folha.max_row, value=tecnico)
                folha.cell(column=7, row=folha.max_row, value=wks)
                folha.cell(column=8, row=folha.max_row, value=patrimonio)
                folha.cell(column=9, row=folha.max_row, value=lacre)
                
                ficheiro.save(r"manutencao.xlsx")
                messagebox.showinfo("Sistema", "Dados salvos com sucesso")
                    
        def clear(): 
            chegada_value.set("")
            saida_value.set("")
            origem_value.set("")
            problema_value.set("")
            solucao_value.set("")
            tecnico_value.set("")
            wks_value.set("")
            patrimonio_value.set("")
            lacre_value.set("")
        
        #Texts variables
        chegada_value = StringVar()
        saida_value = StringVar()
        origem_value = StringVar()
        problema_value = StringVar()
        solucao_value = StringVar()
        tecnico_value = StringVar()
        wks_value = StringVar()
        patrimonio_value = StringVar()
        lacre_value = StringVar()
        
        #Entradas-------------------------------------------------------------------
        chegada_entry = ctk.CTkEntry(self, width=150, textvariable=chegada_value, font=("Century Gothic", 16), fg_color="transparent")
        
        saida_entry = ctk.CTkEntry(self, width=150, textvariable=saida_value, font=("Century Gothic", 16), fg_color="transparent")
        
        origem_entry = ctk.CTkEntry(self, width=200, textvariable=origem_value,  font=("Century Gothic", 16), fg_color="transparent")
        
        problema_entry = ctk.CTkEntry(self, width=350,textvariable=problema_value, font=("Century Gothic", 16), fg_color="transparent")
        
        solucao_entry = ctk.CTkEntry(self, width=350,textvariable=solucao_value, font=("Century Gothic", 16), fg_color="transparent")
        
        tecnico_entry = ctk.CTkEntry(self, width=350,textvariable=tecnico_value, font=("Century Gothic", 16), fg_color="transparent")

        wks_entry = ctk.CTkEntry(self, width=200,textvariable=wks_value, font=("Century Gothic", 16), fg_color="transparent")

        patrimonio_entry = ctk.CTkEntry(self, width=200,textvariable=patrimonio_value, font=("Century Gothic", 16), fg_color="transparent")

        lacre_entry = ctk.CTkEntry(self, width=200,textvariable=lacre_value,  font=("Century Gothic", 16), fg_color="transparent")
        
        #Labels = Dia de chegada, Dia de saida, Problema, Solução/Observação, Técnicos, WKS - Patrimonio - Lacre - Origem
        lb_chegada = ctk.CTkLabel(self, text="Dia de chegada:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_saida = ctk.CTkLabel(self, text="Dia de saída:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_origem = ctk.CTkLabel(self, text="Setor:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_problema = ctk.CTkLabel(self, text="Problema:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_solucao = ctk.CTkLabel(self, text="Solução:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_tecnicos = ctk.CTkLabel(self, text="Tecnico:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_WKS = ctk.CTkLabel(self, text="WKS:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_patrimonio = ctk.CTkLabel(self, text="Patrimonio:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        lb_lacre = ctk.CTkLabel(self, text="Lacre:", font=("Century Gothic bold", 24), text_color=["#000", "#fff"])
        
        #POSICIONANDO NA JANELA--------------------------------------------------------
        lb_chegada.place(x=50, y=120)
        chegada_entry.place(x=50, y=150)
        
        lb_origem.place(x=450, y=120)
        origem_entry.place(x=450, y=150)
        
        lb_saida.place(x=50, y=190)
        saida_entry.place(x=50, y=220)
        
        lb_tecnicos.place(x=50, y=260)
        tecnico_entry.place(x=50, y=290)
        
        lb_problema.place(x=450, y=190)
        problema_entry.place(x=450, y=220)
        
        lb_solucao.place(x=450, y=260)
        solucao_entry.place(x=450, y=290)
        
        lb_WKS.place(x=50, y=330)
        wks_entry.place(x=50, y=360)
        
        lb_patrimonio.place(x=300, y=330)
        patrimonio_entry.place(x=300, y=360)
        
        lb_lacre.place(x=550, y=330)
        lacre_entry.place(x=550, y=360)
        
        #BOTOES
        btn_submit = ctk.CTkButton(self, text="Enviar dados".upper(), command=submit, fg_color="#151", hover_color="#131").place(x=300, y=420)
        
        btn_limpar = ctk.CTkButton(self, text="Limpar".upper(), command=clear, fg_color="#555", hover_color="#333").place(x=500, y=420)
        

if __name__ == "__main__":
    app = App()
    app.mainloop()
```
