# Importação de bibliotecas

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import csv


def ao_fechar():
    resposta = messagebox.askyesno('Sair', 'Deseja mesmo encerrar o programa?')
    if resposta:
        main_janela.destroy()


def menu_tecnicos():
    lista_turnos = ["Manhã", "Tarde", "Noite"]

    def edit_tec():
        global Banco_tec
        flag_sel = False
        if tv1.focus():
            flag_sel = True
            curItem = tv1.focus()
            dic = tv1.item(curItem)
            lista1 = []
            for value in dic.items():
                lista1.append(value)
            lista_aux = lista1[2]
            lista_aux = lista_aux[1]
            codigo1 = lista_aux[0]
            nome1 = str(lista_aux[1])
            cpf1 = str(lista_aux[2])
            cpf1 = ''.join(i for i in cpf1 if i.isdigit())
            contato1 = str(lista_aux[3])
            turno1 = str(lista_aux[4])
            if turno1 == "Manhã":
                turno_n = 0
            elif turno1 == "Tarde":
                turno_n = 1
            else:
                turno_n = 2
            equipe1 = str(lista_aux[5])

            itc = Banco_tec.index[Banco_tec["Código"] == codigo1].tolist()
            itc = int(itc[0])
        else:
            messagebox.showerror(title='Erro', message='Selecione um item para editar', parent=tec_janela)
        if flag_sel:

            tec_janela3 = tk.Tk()
            tec_janela3.title("Editar técnico")
            tec_janela3.geometry("400x500+250+50")
            tec_janela3.minsize(400, 500)
            tec_janela3.maxsize(400, 500)
            tec_janela3.config(bg='#03063b')
            tec_janela3.iconbitmap('icon.ico')

            def editar_codigo():
                global Banco_tec

                if ((entry_nome.get()) and (entry_cpf.get()) and (entry_contato.get()) and (combobox_turno.get()) and (
                        entry_equipe.get())):
                    if not cpf_validate(entry_cpf.get()):
                        messagebox.showwarning("Erro", "CPF inválido", parent=tec_janela3)

                    else:
                        lista_codigos = []
                        codigo = lista_aux[0]
                        nome = entry_nome.get()
                        cpf = '{}.{}.{}-{}'.format(entry_cpf.get()[:3], entry_cpf.get()[3:6], entry_cpf.get()[6:9],
                                                   entry_cpf.get()[9:])
                        contato = entry_contato.get()
                        turno = combobox_turno.get()
                        equipe = entry_equipe.get()
                        lista_codigos.append((codigo, nome, cpf, contato, turno, equipe))
                        edit_tecnico = pd.DataFrame(lista_codigos, columns=['Código', 'Nome', 'CPF', 'Contato', 'Turno',
                                                                            'Equipe'])
                        Banco_tec.iloc[itc] = edit_tecnico.iloc[0]
                        Banco_tec.to_csv('tecnicos.csv', index=False)
                        Load_excel_data()
                        tec_janela3.destroy()

                else:
                    messagebox.showwarning("Atenção", "Preencha todos os campos", parent=tec_janela3)

            label_codigo = tk.Label(tec_janela3, text=f'Código: {codigo1}', bg="#03063b", fg="#edb200",
                                    activebackground="#45A29E", activeforeground="#112D32",
                                    font=("Tahoma", 10, "bold"))
            label_codigo.pack(ipady=10)

            label_nome = tk.Label(tec_janela3, text="Nome:", bg="#03063b", fg="#edb200",
                                  activebackground="#45A29E", activeforeground="#112D32",
                                  font=("Tahoma", 10, "bold"))
            label_nome.pack(ipady=10)

            entry_nomev = tk.StringVar(master=tec_janela3, value=nome1)
            entry_nome = tk.Entry(tec_janela3, textvariable=entry_nomev)
            entry_nome.pack()

            label_cpf = tk.Label(tec_janela3, text="CPF (Somente números):", bg="#03063b", fg="#edb200",
                                 activebackground="#45A29E", activeforeground="#112D32",
                                 font=("Tahoma", 10, "bold"))
            label_cpf.pack(ipady=10)

            entry_cpfv = tk.StringVar(master=tec_janela3, value=cpf1)
            entry_cpf = tk.Entry(tec_janela3, textvariable=entry_cpfv)
            entry_cpf.pack()

            label_contato = tk.Label(tec_janela3, text="Contato:", bg="#03063b", fg="#edb200",
                                     activebackground="#45A29E", activeforeground="#112D32",
                                     font=("Tahoma", 10, "bold"))
            label_contato.pack(ipady=10)

            entry_contatov = tk.StringVar(master=tec_janela3, value=contato1)
            entry_contato = tk.Entry(tec_janela3, textvariable=entry_contatov)
            entry_contato.pack()

            label_turno = tk.Label(tec_janela3, text="Turno:", bg="#03063b", fg="#edb200",
                                   activebackground="#45A29E", activeforeground="#112D32",
                                   font=("Tahoma", 10, "bold"))
            label_turno.pack(ipady=10)

            combobox_turno = ttk.Combobox(tec_janela3, values=lista_turnos, state='readonly')
            combobox_turno.current(turno_n)
            combobox_turno.pack()

            label_equipe = tk.Label(tec_janela3, text="Equipe:", bg="#03063b", fg="#edb200",
                                    activebackground="#45A29E", activeforeground="#112D32",
                                    font=("Tahoma", 10, "bold"))
            label_equipe.pack(ipady=10)

            entry_equipev = tk.StringVar(master=tec_janela3, value=equipe1)
            entry_equipe = tk.Entry(tec_janela3, textvariable=entry_equipev)
            entry_equipe.pack()

            botao_criar_codigo = tk.Button(tec_janela3, text="Salvar alterações", bg="#03063b", fg="#edb200",
                                           activebackground="#45A29E", activeforeground="#112D32",
                                           font=("Tahoma", 12, "bold"), command=editar_codigo)
            botao_criar_codigo.pack(pady=20)

    def cpf_validate(numbers):

        cpf = [int(char) for char in numbers if char.isdigit()]

        if len(cpf) != 11:
            return False


        if cpf == cpf[::-1]:
            return False


        for i in range(9, 11):
            value = sum((cpf[num] * ((i + 1) - num) for num in range(0, i)))
            digit = ((value * 10) % 11) % 10
            if digit != cpf[i]:
                return False
        return True

    def search():
        if search_entry.get():
            query = str(search_entry.get())
            selections = []
            for child in tv1.get_children():
                if query in str(tv1.item(child)['values']):
                    selections.append(child)
                if query.title() in tv1.item(child)['values']:
                    selections.append(child)
                if query.upper() in tv1.item(child)['values']:
                    selections.append(child)
            tv1.selection_set(selections)
        else:
            pass

    def del_tec():
        global Banco_tec
        try:
            curItem = tv1.focus()
            dic = tv1.item(curItem)
            lista1 = []
            for value in dic.items():
                lista1.append(value)
            lista_aux = lista1[2]
            lista_aux = lista_aux[1]
            lista_aux = lista_aux[0]
            Banco_tec = Banco_tec.loc[(Banco_tec['Código'] != lista_aux)]
            Banco_tec.to_csv('tecnicos.csv', index=False)
            Load_excel_data()
        except IndexError:
            messagebox.showerror(title='Erro', message='Selecione um item para excluir', parent=tec_janela)

    def ins_tec():

        tec_janela2 = tk.Tk()
        tec_janela2.title("Cadastro de técnicos")
        tec_janela2.geometry("400x400+250+50")
        tec_janela2.minsize(400, 400)
        tec_janela2.maxsize(400, 400)
        tec_janela2.config(bg='#03063b')
        tec_janela2.iconbitmap('icon.ico')

        def inserir_codigo():
            global Banco_tec
            if ((entry_nome.get()) and (entry_cpf.get()) and (entry_contato.get()) and (combobox_turno.get()) and (
                    entry_equipe.get())):
                if not cpf_validate(entry_cpf.get()):
                    messagebox.showwarning("Erro", "CPF inválido", parent=tec_janela2)

                else:
                    lista_codigos = []
                    codigo = Banco_tec['Código'].iloc[-1]
                    codigo = int(codigo[-1])
                    codigo += 1
                    nome = entry_nome.get()
                    cpf = '{}.{}.{}-{}'.format(entry_cpf.get()[:3], entry_cpf.get()[3:6], entry_cpf.get()[6:9],
                                               entry_cpf.get()[9:])
                    contato = entry_contato.get()
                    turno = combobox_turno.get()
                    equipe = entry_equipe.get()
                    codigo_str = "TEC-{}".format(codigo)
                    lista_codigos.append((codigo_str, nome, cpf, contato, turno, equipe))
                    novo_tecnico = pd.DataFrame(lista_codigos,
                                                columns=['Código', 'Nome', 'CPF', 'Contato', 'Turno', 'Equipe'])
                    Banco_tec = Banco_tec.append(novo_tecnico, ignore_index=False)
                    Banco_tec.to_csv('tecnicos.csv', index=False)
                    Load_excel_data()
                    tec_janela2.destroy()

            else:
                messagebox.showwarning("Atenção", "Preencha todos os campos", parent=tec_janela2)

        label_nome = tk.Label(tec_janela2, text="Nome:", bg="#03063b", fg="#edb200",
                              activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_nome.pack(ipady=10)

        entry_nome = tk.Entry(tec_janela2)
        entry_nome.pack()

        label_cpf = tk.Label(tec_janela2, text="CPF (Somente números):", bg="#03063b", fg="#edb200",
                             activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_cpf.pack(ipady=10)

        entry_cpf = tk.Entry(tec_janela2)
        entry_cpf.pack()

        label_contato = tk.Label(tec_janela2, text="Contato:", bg="#03063b", fg="#edb200",
                                 activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_contato.pack(ipady=10)

        entry_contato = tk.Entry(tec_janela2)
        entry_contato.pack()

        label_turno = tk.Label(tec_janela2, text="Turno:", bg="#03063b", fg="#edb200",
                               activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_turno.pack(ipady=10)

        combobox_turno = ttk.Combobox(tec_janela2, values=lista_turnos)
        combobox_turno.pack()

        label_equipe = tk.Label(tec_janela2, text="Equipe:", bg="#03063b", fg="#edb200",
                                activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_equipe.pack(ipady=10)

        entry_equipe = tk.Entry(tec_janela2)
        entry_equipe.pack()

        botao_criar_codigo = tk.Button(tec_janela2, text="Cadastrar Técnico", bg="#03063b", fg="#edb200",
                                       activebackground="#45A29E", activeforeground="#112D32",
                                       font=("Tahoma", 12, "bold"),
                                       command=inserir_codigo)
        botao_criar_codigo.pack(pady=20)

    tec_janela = tk.Toplevel()
    tec_janela.title("Cadastro de técnicos")
    tec_janela.geometry("800x600+250+50")
    tec_janela.minsize(800, 600)
    tec_janela.maxsize(800, 600)
    tec_janela.config(bg='#03063b')
    tec_janela.iconbitmap('icon.ico')
    style = ttk.Style()
    style.configure("mystyle.Treeview", background='#03063b', foreground='#edb200', highlightthickness=0, bd=0,
                    font=('Calibri', 10))
    style.configure("mystyle.Treeview.Heading", background='#03063b', foreground='#edb200',
                    font=('Calibri', 11, 'bold'))
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])


    frame1 = tk.LabelFrame(tec_janela)
    frame1.place(x=10, y=5, height=300, width=780)

    # Treeview Widget
    tv1 = ttk.Treeview(frame1, style="mystyle.Treeview")
    tv1.place(relheight=1, relwidth=1)

    treescrolly = tk.Scrollbar(frame1, orient="vertical",
                               command=tv1.yview)
    treescrollx = tk.Scrollbar(frame1, orient="horizontal",
                               command=tv1.xview)
    tv1.configure(xscrollcommand=treescrollx.set,
                  yscrollcommand=treescrolly.set)
    treescrollx.pack(side="bottom", fill="x")
    treescrolly.pack(side="right", fill="y")

    def clear_data():
        tv1.delete(*tv1.get_children())
        return None

    def Load_excel_data():

        df = pd.read_csv('tecnicos.csv')

        clear_data()
        tv1["column"] = list(df.columns)
        tv1["show"] = "headings"

        for column in tv1["columns"]:
            tv1.heading(column, text=column)
            tv1.column(column, minwidth=100, width=120, stretch=False, anchor=tk.CENTER)
        df_rows = df.to_numpy().tolist()
        for row in df_rows:
            tv1.insert("", "end", values=row)
        return None

    Load_excel_data()

    def exp_tec():
        try:
            with filedialog.asksaveasfile(mode='w', defaultextension='.xlsx') as file:
                Banco_tec.to_excel(file.name, index=False)
            messagebox.showinfo(title='Sucesso', message='Arquivo salvo na pasta selecionada', parent=tec_janela)
        except AttributeError:
            messagebox.showerror(title='Erro', message='Cancelado pelo usuário', parent=tec_janela)

    search_entry = tk.Entry(tec_janela, width=50)
    search_entry.place(x=230, y=312)

    search_button = tk.Button(tec_janela, text="buscar", command=search)
    search_button.place(x=545, y=310)

    bt11 = tk.Button(tec_janela, text='Cadastrar', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=ins_tec)
    bt11.place(x=100, y=370)

    bt22 = tk.Button(tec_janela, text='Editar', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=edit_tec)
    bt22.place(x=270, y=370)

    bt33 = tk.Button(tec_janela, text='Excluir', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=del_tec)
    bt33.place(x=400, y=370)

    bt44 = tk.Button(tec_janela, text='Exportar', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=exp_tec)
    bt44.place(x=550, y=370)

    tec_janela.mainloop()


def menu_ferramentas():
    def edit_fer():
        global Banco_fer
        flag_sel = False
        if tv2.focus():
            flag_sel = True
            curItem = tv2.focus()
            dic = tv2.item(curItem)
            lista1 = []
            for value in dic.items():
                lista1.append(value)
            lista_aux = lista1[2]
            lista_aux = lista_aux[1]
            codigo1 = lista_aux[0]
            desc1 = str(lista_aux[1])
            fab1 = str(lista_aux[2])
            volt1 = str(lista_aux[3])
            part1 = str(lista_aux[4])
            tam1 = str(lista_aux[5])
            um1 = str(lista_aux[6])
            tipo1 = str(lista_aux[7])
            mat1 = str(lista_aux[8])
            temp1 = str(lista_aux[9])

            itc = Banco_fer.index[Banco_fer["Código"] == codigo1].tolist()
            itc = int(itc[0])
        else:
            messagebox.showerror(parent=fer_janela, title='Erro', message='Selecione um item para editar')

        if flag_sel:

            fer_janela3 = tk.Tk()
            fer_janela3.title("Editar ferramenta")
            fer_janela3.geometry("440x500+250+50")
            fer_janela3.minsize(440, 500)
            fer_janela3.maxsize(440, 500)
            fer_janela3.config(bg='#03063b')
            fer_janela3.iconbitmap('icon.ico')

            def editar_ferramenta():
                global Banco_fer
                if ((entry_desc.get()) and (entry_fab.get()) and (entry_volt.get()) and (entry_part.get()) and
                        (entry_tam.get()) and (entry_um.get()) and (entry_tipo.get()) and (entry_mat.get()) and
                        (entry_temp.get())):

                    lista_dados = []
                    codigo = lista_aux[0]
                    desc = entry_desc.get()
                    fab = entry_fab.get()
                    volt = entry_volt.get()
                    part = entry_part.get()
                    tamanho = entry_tam.get()
                    unmed = entry_um.get()
                    tipo = entry_tipo.get()
                    material = entry_mat.get()
                    tempo = entry_temp.get()
                    lista_dados.append((codigo, desc, fab, volt, part, tamanho, unmed, tipo, material, tempo))
                    edit_ferramenta = pd.DataFrame(lista_dados,
                                                   columns=['Código', 'Descrição', 'Fabricante', 'Voltagem',
                                                            'Part Number', 'Tamanho', 'Unidade de medida',
                                                            'Tipo', 'Material', 'Tempo máx'])
                    Banco_fer.iloc[itc] = edit_ferramenta.iloc[0]
                    Banco_fer.to_csv('ferramentas.csv', index=False)
                    Load_excel_data()
                    fer_janela3.destroy()

                else:
                    messagebox.showwarning("Atenção", "Preencha todos os campos", master=fer_janela3)

            label_codigo = tk.Label(fer_janela3, text=f'Código: {codigo1}', bg="#03063b", fg="#edb200",
                                    activebackground="#45A29E", activeforeground="#112D32",
                                    font=("Tahoma", 10, "bold"))
            label_codigo.grid(row=0, column=2, padx=10, ipady=10)

            label_desc = tk.Label(fer_janela3, text="Descrição:", bg="#03063b", fg="#edb200",
                                  activebackground="#45A29E", activeforeground="#112D32",
                                  font=("Tahoma", 10, "bold"))
            label_desc.grid(row=1, column=1, padx=10, pady=10, ipady=10)

            entry_descv = tk.StringVar(master=fer_janela3, value=desc1)
            entry_desc = tk.Entry(fer_janela3, textvariable=entry_descv)
            entry_desc.grid(row=2, column=1, padx=10)

            label_fab = tk.Label(fer_janela3, text="Fabricante:", bg="#03063b", fg="#edb200",
                                 activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
            label_fab.grid(row=4, column=1, padx=10, pady=10, ipady=10)

            entry_fabv = tk.StringVar(master=fer_janela3, value=fab1)
            entry_fab = tk.Entry(fer_janela3, textvariable=entry_fabv)
            entry_fab.grid(row=5, column=1, padx=10)

            label_volt = tk.Label(fer_janela3, text="Voltagem:", bg="#03063b", fg="#edb200",
                                  activebackground="#45A29E", activeforeground="#112D32",
                                  font=("Tahoma", 10, "bold"))
            label_volt.grid(row=7, column=1, padx=10, pady=10, ipady=10)

            entry_voltv = tk.StringVar(master=fer_janela3, value=volt1)
            entry_volt = tk.Entry(fer_janela3, textvariable=entry_voltv)
            entry_volt.grid(row=8, column=1, padx=10)

            label_part = tk.Label(fer_janela3, text="Part Number:", bg="#03063b", fg="#edb200",
                                  activebackground="#45A29E", activeforeground="#112D32",
                                  font=("Tahoma", 10, "bold"))
            label_part.grid(row=10, column=1, padx=10, pady=10, ipady=10)

            entry_partv = tk.StringVar(master=fer_janela3, value=part1)
            entry_part = tk.Entry(fer_janela3, textvariable=entry_partv)
            entry_part.grid(row=11, column=1, padx=10)

            label_tam = tk.Label(fer_janela3, text="Tamanho:", bg="#03063b", fg="#edb200",
                                 activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
            label_tam.grid(row=13, column=1, padx=10, pady=10, ipady=10)

            entry_tamv = tk.StringVar(master=fer_janela3, value=tam1)
            entry_tam = tk.Entry(fer_janela3, textvariable=entry_tamv)
            entry_tam.grid(row=14, column=1, padx=10)

            label_um = tk.Label(fer_janela3, text="Unidade de medida:", bg="#03063b", fg="#edb200",
                                activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
            label_um.grid(row=1, column=3, padx=10, pady=10, ipady=10)

            entry_umv = tk.StringVar(master=fer_janela3, value=um1)
            entry_um = tk.Entry(fer_janela3, textvariable=entry_umv)
            entry_um.grid(row=2, column=3, padx=10)

            label_tipo = tk.Label(fer_janela3, text="Tipo de ferramenta:", bg="#03063b", fg="#edb200",
                                  activebackground="#45A29E", activeforeground="#112D32",
                                  font=("Tahoma", 10, "bold"))
            label_tipo.grid(row=4, column=3, padx=10, pady=10, ipady=10)

            entry_tipov = tk.StringVar(master=fer_janela3, value=tipo1)
            entry_tipo = tk.Entry(fer_janela3, textvariable=entry_tipov)
            entry_tipo.grid(row=5, column=3, padx=10)

            label_mat = tk.Label(fer_janela3, text="Material:", bg="#03063b", fg="#edb200",
                                 activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
            label_mat.grid(row=7, column=3, padx=10, pady=10, ipady=10)

            entry_matv = tk.StringVar(master=fer_janela3, value=mat1)
            entry_mat = tk.Entry(fer_janela3, textvariable=entry_matv)
            entry_mat.grid(row=8, column=3, padx=10)

            label_temp = tk.Label(fer_janela3, text="Tempo máximo:", bg="#03063b", fg="#edb200",
                                  activebackground="#45A29E", activeforeground="#112D32",
                                  font=("Tahoma", 10, "bold"))
            label_temp.grid(row=10, column=3, padx=10, pady=10, ipady=10)

            entry_tempv = tk.StringVar(master=fer_janela3, value=temp1)
            entry_temp = tk.Entry(fer_janela3, textvariable=entry_tempv)
            entry_temp.grid(row=11, column=3, padx=10)

            botao_criar_codigo = tk.Button(fer_janela3, text="Salvar alterações", bg="#03063b", fg="#edb200",
                                           activebackground="#45A29E", activeforeground="#112D32",
                                           font=("Tahoma", 10, "bold"),
                                           command=editar_ferramenta)
            botao_criar_codigo.grid(row=16, column=2, padx=10, pady=10)

    def search():
        if search_entry.get():
            query = str(search_entry.get())
            selections = []
            for child in tv2.get_children():
                if query in str(tv2.item(child)['values']):
                    selections.append(child)
                if query.title() in tv2.item(child)['values']:
                    selections.append(child)
                if query.upper() in tv2.item(child)['values']:
                    selections.append(child)
            tv2.selection_set(selections)
        else:
            pass

    def del_fer():
        global Banco_fer
        try:
            curItem = tv2.focus()
            dic = tv2.item(curItem)
            lista1 = []
            for value in dic.items():
                lista1.append(value)
            lista_aux = lista1[2]
            lista_aux = lista_aux[1]
            lista_aux = lista_aux[0]
            Banco_fer = Banco_fer.loc[(Banco_fer['Código'] != lista_aux)]
            Banco_fer.to_csv('ferramentas.csv', index=False)
            Load_excel_data()
        except IndexError:
            messagebox.showerror(title='Erro', message='Selecione um item para excluir', parent=fer_janela)

    def ins_fer():

        fer_janela2 = tk.Tk()
        fer_janela2.title("Cadastro de ferramentas")
        fer_janela2.geometry("400x650+250+50")
        fer_janela2.minsize(400, 650)
        fer_janela2.maxsize(400, 650)
        fer_janela2.config(bg='#03063b')
        fer_janela2.iconbitmap('icon.ico')

        def inserir_ferramenta():
            global Banco_fer
            if ((entry_desc.get()) and (entry_fab.get()) and (entry_volt.get()) and (entry_part.get()) and
                    (entry_tam.get()) and (entry_um.get()) and (entry_tipo.get()) and (entry_mat.get()) and
                    (entry_temp.get())):

                lista_dados = []
                try:
                    codigo = Banco_fer['Código'].iloc[-1]
                    codigo = int(codigo[-1])
                    codigo += 1
                    codigo_str = "FER-{}".format(codigo)

                except:
                    codigo_str = "FER-1"

                desc = str(entry_desc.get())
                fab = str(entry_fab.get())
                volt = str(entry_volt.get())
                part = str(entry_part.get())
                tamanho = str(entry_tam.get())
                unmed = str(entry_um.get())
                tipo = str(entry_tipo.get())
                material = str(entry_mat.get())
                if entry_temp.get().isnumeric():
                    tempo = int(entry_temp.get())
                    if 1 <= tempo <= 24:
                        tempo = str(entry_temp.get())
                        tempo = "{} horas".format(tempo)
                        lista_dados.append((codigo_str, desc, fab, volt, part, tamanho, unmed, tipo, material, tempo))
                        nova_ferramenta = pd.DataFrame(lista_dados,
                                                       columns=['Código', 'Descrição', 'Fabricante', 'Voltagem',
                                                                'Part Number',
                                                                'Tamanho', 'Unidade de medida', 'Tipo', 'Material',
                                                                'Tempo máx'])
                        Banco_fer = Banco_fer.append(nova_ferramenta, ignore_index=True)
                        Banco_fer.to_csv('ferramentas.csv', index=False)
                        Load_excel_data()
                        fer_janela2.destroy()
                    else:
                        messagebox.showerror(title='Erro',
                                             message='Tempo mínimo de reserva: 1 hora.\n'
                                                     'Tempo máximo de reserva: 24 horas.',
                                             master=fer_janela2)
                else:
                    messagebox.showerror(title='Erro', message='Entre com um período de reserva entre 1 e 24 horas',
                                         master=fer_janela2)

            else:
                messagebox.showwarning("Atenção", "Preencha todos os campos", master=fer_janela2)

        label_desc = tk.Label(fer_janela2, text="Descrição:", bg="#03063b", fg="#edb200",
                              activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_desc.pack(ipady=10)

        entry_desc = tk.Entry(fer_janela2)
        entry_desc.pack()

        label_fab = tk.Label(fer_janela2, text="Fabricante:", bg="#03063b", fg="#edb200",
                             activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_fab.pack(ipady=10)

        entry_fab = tk.Entry(fer_janela2)
        entry_fab.pack()

        label_volt = tk.Label(fer_janela2, text="Voltagem:", bg="#03063b", fg="#edb200",
                              activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_volt.pack(ipady=10)

        entry_volt = tk.Entry(fer_janela2)
        entry_volt.pack()

        label_part = tk.Label(fer_janela2, text="Part Number:", bg="#03063b", fg="#edb200",
                              activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_part.pack(ipady=10)

        entry_part = tk.Entry(fer_janela2)
        entry_part.pack()

        label_tam = tk.Label(fer_janela2, text="Tamanho:", bg="#03063b", fg="#edb200",
                             activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_tam.pack(ipady=10)

        entry_tam = tk.Entry(fer_janela2)
        entry_tam.pack()

        label_um = tk.Label(fer_janela2, text="Unidade de medida:", bg="#03063b", fg="#edb200",
                            activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_um.pack(ipady=10)

        entry_um = tk.Entry(fer_janela2)
        entry_um.pack()

        label_tipo = tk.Label(fer_janela2, text="Tipo de ferramenta:", bg="#03063b", fg="#edb200",
                              activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_tipo.pack(ipady=10)

        entry_tipo = tk.Entry(fer_janela2)
        entry_tipo.pack()

        label_mat = tk.Label(fer_janela2, text="Material:", bg="#03063b", fg="#edb200",
                             activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_mat.pack(ipady=10)

        entry_mat = tk.Entry(fer_janela2)
        entry_mat.pack()

        label_temp = tk.Label(fer_janela2, text="Tempo máximo de reserva (horas):", bg="#03063b", fg="#edb200",
                              activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 10, "bold"))
        label_temp.pack(ipady=10)

        entry_temp = tk.Entry(fer_janela2)
        entry_temp.pack()

        botao_criar_codigo = tk.Button(fer_janela2, text="Cadastrar ferramenta", bg="#03063b", fg="#edb200",
                                       activebackground="#45A29E", activeforeground="#112D32",
                                       font=("Tahoma", 10, "bold"),
                                       command=inserir_ferramenta)
        botao_criar_codigo.pack(pady=20)

    fer_janela = tk.Toplevel()
    fer_janela.title("Cadastro de ferramentas")
    fer_janela.geometry("800x600+250+50")
    fer_janela.minsize(800, 600)
    fer_janela.maxsize(800, 600)
    fer_janela.config(bg='#03063b')
    fer_janela.iconbitmap('icon.ico')
    style = ttk.Style()
    style.configure("mystyle.Treeview", background='#03063b', foreground='#edb200', highlightthickness=0, bd=0,
                    font=('Calibri', 10))
    style.configure("mystyle.Treeview.Heading", background='#03063b', foreground='#edb200',
                    font=('Calibri', 11, 'bold'))
    style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])

    frame2 = tk.LabelFrame(fer_janela)
    frame2.place(x=10, y=5, height=300, width=780)

    tv2 = ttk.Treeview(frame2, style="mystyle.Treeview")
    tv2.place(relheight=1, relwidth=1)

    treescrolly = tk.Scrollbar(frame2, orient="vertical",
                               command=tv2.yview)
    treescrollx = tk.Scrollbar(frame2, orient="horizontal",
                               command=tv2.xview)
    tv2.configure(xscrollcommand=treescrollx.set,
                  yscrollcommand=treescrolly.set)
    treescrollx.pack(side="bottom", fill="x")
    treescrolly.pack(side="right", fill="y")

    def clear_data():
        tv2.delete(*tv2.get_children())
        return None

    def Load_excel_data():


        df2 = pd.read_csv('ferramentas.csv')

        clear_data()
        tv2["column"] = list(df2.columns, )
        tv2["show"] = "headings"

        for column in tv2["columns"]:
            tv2.heading(column, text=column)
            tv2.column(column, minwidth=100, width=100, stretch=False, anchor=tk.CENTER)

        df2_rows = df2.to_numpy().tolist()
        for row in df2_rows:
            tv2.insert("", "end", values=row)
        return None

    Load_excel_data()

    def exp_fer():
        try:
            with filedialog.asksaveasfile(mode='w', defaultextension='.xlsx') as file:
                Banco_fer.to_excel(file.name, index=False)
            messagebox.showinfo(title='Sucesso', message='Arquivo salvo na pasta selecionada', parent=fer_janela)
        except AttributeError:
            messagebox.showerror(title='Erro', message='Cancelado pelo usuário', parent=fer_janela)

    search_entry = tk.Entry(fer_janela, width=50)
    search_entry.place(x=230, y=312)

    search_button = tk.Button(fer_janela, text="buscar", command=search)
    search_button.place(x=545, y=310)

    bt11 = tk.Button(fer_janela, text='Cadastrar', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=ins_fer)
    bt11.place(x=100, y=370)

    bt22 = tk.Button(fer_janela, text='Editar', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=edit_fer)
    bt22.place(x=270, y=370)

    bt33 = tk.Button(fer_janela, text='Excluir', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=del_fer)
    bt33.place(x=400, y=370)

    bt44 = tk.Button(fer_janela, text='Exportar', bg="#03063b", fg="#edb200", activebackground="#45A29E",
                     activeforeground="#112D32", font=("Tahoma", 16, "bold"), command=exp_fer)
    bt44.place(x=550, y=370)

    fer_janela.mainloop()



main_janela = tk.Tk()
main_janela.protocol('WM_DELETE_WINDOW', ao_fechar)
main_janela.title("Missão Certificação - Gerenciador de Ferramentas")
main_janela.geometry('800x600+250+50')
main_janela.minsize(800, 600)
main_janela.maxsize(800, 600)
main_janela.config(bg='#03063b')
main_janela.iconbitmap('icon.ico')

try:
    Banco_tec = pd.read_csv('tecnicos.csv')
except:
    messagebox.showwarning(title='Aviso', message='Banco de técnicos não encontrado.\nCriando novo banco de dados.',
                           master=main_janela)
    with open('./tecnicos.csv', 'w', encoding='utf-8') as csvfile:
        csv.writer(csvfile, delimiter=',').writerow(
            ['Código', 'Nome', 'CPF', 'Contato', 'Turno', 'Equipe'])
    Banco_tec = pd.read_csv('tecnicos.csv')

try:
    Banco_fer = pd.read_csv('ferramentas.csv')
except:
    messagebox.showwarning(title='Aviso', message='Banco de ferramentas não encontrado.\nCriando novo banco de dados.',
                           master=main_janela)
    with open('./ferramentas.csv', 'w', encoding='utf-8') as csvfile:
        csv.writer(csvfile, delimiter=',').writerow(
            ['Código', 'Descrição', 'Fabricante', 'Voltagem', 'Part Number', 'Tamanho', 'Unidade de medida', 'Tipo',
             'Material', 'Tempo máx'])
    Banco_fer = pd.read_csv('ferramentas.csv')

logo = tk.PhotoImage(file='logo.png')
ttk.Label(image=logo, borderwidth=0, background='#03063b').place(x=56, y=5)

botao_tecnicos = tk.Button(main_janela, text="Cadastrar Técnicos", bg="#03063b", fg="#edb200",
                           activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 16, "bold"),
                           command=menu_tecnicos)
botao_tecnicos.place(x=260, y=275, height=45, width=280)

botao_ferramentas = tk.Button(main_janela, text="Cadastrar Ferramentas", bg="#03063b", fg="#edb200",
                              activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 16, "bold"),
                              command=menu_ferramentas)
botao_ferramentas.place(x=260, y=350, height=45, width=280)

botao_reservas = tk.Button(main_janela, text="Reservar Ferramentas", bg="#03063b", fg="#edb200",
                           activebackground="#45A29E", activeforeground="#112D32", font=("Tahoma", 16, "bold"),
                           command=None)
botao_reservas.place(x=260, y=425, height=45, width=280)

main_janela.mainloop()

