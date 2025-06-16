import customtkinter as ctk  
from tkinter import filedialog, messagebox  
import os  
import pandas as pd  
  
# Importar PIL para manipulação de imagens  
try:  
    from PIL import Image, ImageTk, __version__ as PILLOW_VERSION  
    from packaging import version  
    PIL_AVAILABLE = True  
except ImportError:  
    PIL_AVAILABLE = False  
    print("Pillow não está instalado. Instale usando 'pip install pillow'.")  
  
class Application(ctk.CTk):  
    def __init__(self):  
        super().__init__()  
  
        # Configurar tema e aparência  
        ctk.set_appearance_mode("light")   
        ctk.set_default_color_theme("blue")    
  
        self.title("Reconciliador de Inventário - HP")  
        self.geometry("400x550")  # 
        self.resizable(False, False)  
  
        # Cores personalizadas alinhadas com a paleta da HP  
        self.hp_blue = "#0857C3"  
        self.light_gray = "#EDEDED"  
        self.text_black = "#000000"  
        self.text_white = "#FFFFFF"  
  
        # Frame principal sem dimensões fixas para melhor responsividade  
        main_frame = ctk.CTkFrame(self, corner_radius=15, fg_color=self.light_gray)  
        main_frame.pack(padx=20, pady=20, fill="both", expand=True)  
  
        # Configurar grid no main_frame  
        main_frame.grid_rowconfigure(0, weight=0)  # Logo  
        main_frame.grid_rowconfigure(1, weight=0)  # Espaço  
        main_frame.grid_rowconfigure(2, weight=0)  # Browse  
        main_frame.grid_rowconfigure(3, weight=1)  # Status  
        main_frame.grid_rowconfigure(4, weight=0)  # Confirm  
        main_frame.grid_columnconfigure(0, weight=1)  
  
        # Logo  
        self.create_logo(main_frame)  
  
        # Botão de Procurar e Exibição do Caminho  
        self.create_browse_section(main_frame)  
  
        # Status dos Arquivos  
        self.create_status_section(main_frame)  
  
        # Botão de Confirmação  
        self.create_confirm_button(main_frame)  

        # Configurar ícone da janela e da barra de tarefas 
        self.iconbitmap("fotos\logohp.ico")
  
    def get_resample_method(self):  
        """  
        Retorna o método de reamostragem adequado com base na versão do Pillow.  
        """  
        try:  
            # Pillow >= 10.0.0  
            return Image.Resampling.LANCZOS  
        except AttributeError:  
            # Pillow < 10.0.0  
            return Image.ANTIALIAS  
  
    def create_logo(self, parent):  
        logo_frame = ctk.CTkFrame(parent, fg_color="transparent")  
        logo_frame.grid(row=0, column=0, pady=(0, 10), sticky="n")  
  
        logo_path = "fotos\\logohp.png"  # Certifique-se de que este caminho está correto  
  
        if PIL_AVAILABLE and os.path.exists(logo_path):  
            try:  
                logo_image = Image.open(logo_path)  
  
                # Obter o método de reamostragem adequado  
                resample_method = self.get_resample_method()  
  
                logo_image = logo_image.resize((145, 145), resample=resample_method)  
                self.logo_photo = ctk.CTkImage(logo_image, size=(145, 145))  
                # Passar 'logo_frame' como primeiro argumento posicional  
                logo_label = ctk.CTkLabel(logo_frame, image=self.logo_photo, text="", fg_color="transparent")  
                logo_label.pack()  
            except Exception as e:  
                print(f"Erro ao carregar a imagem: {e}")  
                logo_label = ctk.CTkLabel(  
                    logo_frame,  
                    text="LOGO",  
                    font=ctk.CTkFont(size=24, weight="bold"),  
                    text_color=self.hp_blue  
                )  
                logo_label.pack()  
        else:  
            if not PIL_AVAILABLE:  
                print("Pillow não está instalado.")  
            if not os.path.exists(logo_path):  
                print(f"Arquivo de logo não encontrado em: {logo_path}")  
            # Passar 'logo_frame' como primeiro argumento posicional  
            logo_label = ctk.CTkLabel(  
                logo_frame,  
                text="LOGO",  
                font=ctk.CTkFont(size=24, weight="bold"),  
                text_color=self.hp_blue  
            )  
            logo_label.pack()  
  
    def create_browse_section(self, parent):  
        browse_frame = ctk.CTkFrame(parent, fg_color="transparent")  
        browse_frame.grid(row=2, column=0, sticky='ew', pady=10, padx=20)  
  
        # Configurar grid para browse_frame  
        browse_frame.grid_columnconfigure(0, weight=0)  
        browse_frame.grid_columnconfigure(1, weight=1)  
  
        browse_button = ctk.CTkButton(  
            browse_frame,  
            text="Procurar Pasta",  
            command=self.browse_folder,  
            width=120,  
            fg_color=self.hp_blue,  
            hover_color="#12A3D8",  
            text_color=self.text_white,  
            font=ctk.CTkFont(size=12, weight="bold")  
        )  
        # Passar 'browse_frame' como primeiro argumento posicional  
        browse_button.grid(row=0, column=0, padx=(0, 10), pady=0, sticky="w")  
  
        self.path_var = ctk.StringVar(value="Nenhuma pasta selecionada")  
        self.path_label = ctk.CTkLabel(  
            browse_frame,  
            textvariable=self.path_var,  
            wraplength=250,  # Ajustado para caber na janela de 400x600  
            anchor="w",  
            font=ctk.CTkFont(size=10),  
            text_color=self.text_black  
        )  
        # Passar 'browse_frame' como primeiro argumento posicional  
        self.path_label.grid(row=0, column=1, sticky="w")  
  
    def create_status_section(self, parent):  
        status_frame = ctk.CTkFrame(parent, fg_color="transparent")  
        status_frame.grid(row=3, column=0, sticky='nsew', pady=20, padx=20)  
  
        # Configurar grid no status_frame  
        status_frame.grid_rowconfigure(0, weight=0)  
        status_frame.grid_rowconfigure(1, weight=0)  
        status_frame.grid_rowconfigure(2, weight=0)  
        status_frame.grid_rowconfigure(4, weight=0)
        status_frame.grid_columnconfigure(0, weight=1)  
  
        # Título da seção  
        title_label = ctk.CTkLabel(  
            status_frame,  
            text="Conciliação de Inventário",  
            font=ctk.CTkFont(size=16, weight="bold"),  
            text_color=self.hp_blue  
        )  
        # Passar 'status_frame' como primeiro argumento posicional  
        title_label.grid(row=0, column=0, sticky='w', pady=(0, 10))  
  
        # MB51 Status  
        self.mb51_status_var = ctk.StringVar(value="MB51: N/A")  
        self.mb51_label = ctk.CTkLabel(  
            status_frame,  
            textvariable=self.mb51_status_var,  
            font=ctk.CTkFont(size=14),  
            text_color=self.text_black  
        )  
        self.mb51_label.grid(row=1, column=0, sticky='w', pady=5)  
  
        # ZNFREP Status  
        self.znfrep_status_var = ctk.StringVar(value="ZNFREP: N/A")  
        self.znfrep_label = ctk.CTkLabel(  
            status_frame,  
            textvariable=self.znfrep_status_var,  
            font=ctk.CTkFont(size=14),  
            text_color=self.text_black  
        )  
        self.znfrep_label.grid(row=2, column=0, sticky='w', pady=5)


        # CORDOF Status  
        self.cordof_status_var = ctk.StringVar(value="CORDOF: N/A")   
        self.cordof_label = ctk.CTkLabel(                           
        status_frame,                                        
        textvariable=self.cordof_status_var,            
        font=ctk.CTkFont(size=14),                                
        text_color=self.text_black                               
        )                                                               
        self.cordof_label.grid(row=3, column=0, sticky='w', pady=5)  

        # IGI Status
        self.igi_status_var = ctk.StringVar(value="IGI: N/A")
        self.igi_label = ctk.CTkLabel(
            status_frame,
            textvariable=self.igi_status_var,
            font=ctk.CTkFont(size=14),
            text_color=self.text_black
        )
        self.igi_label.grid(row=4, column=0, sticky='w', pady=5)

  
    def create_confirm_button(self, parent):  
        confirm_frame = ctk.CTkFrame(parent, fg_color="transparent")  
        confirm_frame.grid(row=4, column=0, pady=10)  
  
        ok_button = ctk.CTkButton(  
            confirm_frame,  
            text="OK",  
            command=self.confirm_action,  
            width=200,  
            height=40,  
            fg_color=self.hp_blue,  
            hover_color="#668CE3",  
            text_color=self.text_white,  
            font=ctk.CTkFont(size=16, weight="bold")  
        )  
        # Passar 'confirm_frame' como primeiro argumento posicional  
        ok_button.pack()  
  
    def browse_folder(self):  
        folder_selected = filedialog.askdirectory()  
        if folder_selected:  
            self.path_var.set(folder_selected)  
            self.check_files(folder_selected)  
  
    def check_files(self, folder_path):  
        # Resetar status  
        self.mb51_status_var.set("MB51: N/A")  
        self.mb51_label.configure(text_color=self.text_black)  
        
        self.znfrep_status_var.set("ZNFREP: N/A")  
        self.znfrep_label.configure(text_color=self.text_black)  

        self.cordof_status_var.set("CORDOF: N/A")  
        self.cordof_label.configure(text_color=self.text_black)  
        
        self.igi_status_var.set("IGI: N/A")
        self.igi_label.configure(text_color=self.text_black)

  
        # Verificar MB51.xlsx  
        mb51_path = os.path.join(folder_path, "MB51.xlsx")  
        if os.path.isfile(mb51_path):  
            self.mb51_status_var.set("MB51: OK")  
            self.mb51_label.configure(text_color="green")  
        else:  
            self.mb51_status_var.set("MB51: Não encontrado")  
            self.mb51_label.configure(text_color="red")  
  
        # Verificar ZNFREP.xlsx  
        znfrep_path = os.path.join(folder_path, "ZNFREP.xlsx")  
        if os.path.isfile(znfrep_path):  
            self.znfrep_status_var.set("ZNFREP: OK")  
            self.znfrep_label.configure(text_color="green")  
        else:  
            self.znfrep_status_var.set("ZNFREP: Não encontrado")  
            self.znfrep_label.configure(text_color="red") 

        # Verificar CORDOF.txt 
        cordof_path = os.path.join(folder_path, "CORDOF.txt")  
        if os.path.isfile(cordof_path):  
            self.cordof_status_var.set("CORDOF: OK")  
            self.cordof_label.configure(text_color="green")  
        else:  
            self.cordof_status_var.set("CORDOF: Não encontrado")  
            self.cordof_label.configure(text_color="red")     

        # Verificar IGI.xlsx
        igi_path = os.path.join(folder_path, "IGI.xlsx")
        if os.path.isfile(igi_path):
            self.igi_status_var.set("IGI: OK")
            self.igi_label.configure(text_color="green")
        else:
            self.igi_status_var.set("IGI: Não encontrado")
            self.igi_label.configure(text_color="orange")  # cor neutra para opcional
     
  
    def confirm_action(self):  
        selected_path = self.path_var.get()  
        mb51_path = os.path.join(selected_path, "MB51.xlsx")  
        znfrep_path = os.path.join(selected_path, "ZNFREP.xlsx") 
        cordof_path = os.path.join(selected_path, "CORDOF.txt") 
        igi_path = os.path.join(selected_path, "IGI.xlsx") 
        igi_exists = os.path.isfile(igi_path)

        if os.path.isdir(selected_path) and os.path.isfile(mb51_path) and os.path.isfile(znfrep_path) and os.path.isfile(cordof_path):
            try:
                # Processamento
                self.process_files(
                    mb51_path,
                    znfrep_path,
                    cordof_path,
                    selected_path,
                    igi_path if igi_exists else None
                )
                messagebox.showinfo("Sucesso", "Processamento concluído e arquivos salvos com sucesso.")  
            except Exception as e:  
                import traceback  
                tb_str = traceback.format_exc()  
                messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {e}\n\nDetalhes:\n{tb_str}")  
        else:  
            messagebox.showerror("Erro", "Certifique-se de que os arquivos 'MB51.xlsx', 'ZNFREP.xlsx' e 'CORDOF.txt' estão na pasta selecionada.") 
  
  
    def clean_string(self, value):  
        value = str(value)  
        if value.endswith('.0'):  
            value = value[:-2]  
        return value  
  
    
    def process_files(self, mb51_path, znfrep_path, cordof_path, output_folder, igi_path=None):

        # Ler os arquivos de Excel  
        df_MB51 = pd.read_excel(mb51_path)  
        df_ZNFREP = pd.read_excel(znfrep_path) 
        df_CORDOF = pd.read_csv(cordof_path, sep=";") 
      
       # Validação para existencia do IGI.xlsx
        if igi_path:
            df_IGI = pd.read_excel(igi_path)

        # Função para converter para string e remover ".0"  
        def clean_string(value):  
            value = str(value)  
            if value.endswith('.0'):  
                value = value[:-2]  
            return value  
        

        ####################################################### CORDOF ###############################################
        

        # Substituir "." por "," em colunas específicas  
        cols_to_replace = ["VL_TOTAL_CONTABIL", "VL_TOTAL_FATURADO", "VL_TOTAL_INSS_RET", "VL_TOTAL_IRRF", "VL_TOTAL_ICMS",   
                        "VL_TOTAL_STF", "VL_TOTAL_STF_IDO", "VL_TOTAL_STF_SUBSTITUIDO", "VL_TOTAL_STT", "VL_TOTAL_IPI",   
                        "VL_TOTAL_ISS", "VL_TOTAL_PIS_PASEP", "VL_TOTAL_COFINS", "VL_IMPOSTO_COFINS", "VL_CONTABIL",   
                        "VL_FATURADO", "VL_BASE_INSS_RET", "VL_INSS_RET", "VL_BASE_IRRF", "VL_IRRF", "VL_BASE_ICMS",   
                        "VL_BASE_STF", "VL_STF", "VL_BASE_STF_IDO", "VL_BASE_STT", "VL_BASE_IPI", "VL_IPI", "VL_BASE_ISS",   
                        "VL_ISS", "VL_BASE_PIS", "VL_PIS", "VL_BASE_COFINS", "VL_IMPOSTO_PIS", "VL_COFINS", "VL_FRETE",   
                        "VL_RATEIO_ODA", "VL_BASE_PIS_RET", "VL_PIS_RET", "VL_BASE_COFINS_RET", "VL_COFINS_RET",   
                        "VL_BASE_CSLL_RET", "VL_CSLL_RET", "VL_TRIBUTAVEL_STT", "VL_TRIBUTAVEL_ICMS", "VL_TRIBUTAVEL_IPI",   
                        "VL_TRIBUTAVEL_STF", "VL_TOTAL_FCP", "VL_TOTAL_ICMS_PART_REM", "VL_TOTAL_BASE_ICMS_PART_DEST",   
                        "VL_TOTAL_ICMS_PART_DEST", "VL_TOTAL_BASE_ICMS_FCP", "VL_BASE_ICMS_PART_REM", "VL_ICMS_PART_REM",   
                        "VL_BASE_ICMS_PART_DEST", "VL_ICMS_PART_DEST", "VL_ICMS_FCP", "VL_ICMS_FCPST", "VL_II", "VL_ICMS",   
                        "PRECO_TOTAL", "PRECO_UNITARIO", "PRECO_TOTAL_M"]  
        
        df_CORDOF[cols_to_replace] = df_CORDOF[cols_to_replace].replace('.', ',', regex=True) 

        # Converter colunas específicas para datetime e date  
        #df_CORDOF["DT_FATO_GERADOR_IMPOSTO"] = pd.to_datetime(df_CORDOF["DT_FATO_GERADOR_IMPOSTO"], format="%d/%m/%Y %H:%M:%S").dt.date  
        #df_CORDOF["DH_EMISSAO"] = pd.to_datetime(df_CORDOF["DH_EMISSAO"], format="%d/%m/%Y %H:%M:%S").dt.date  
        
        # Adicionar coluna DANFE  
        df_CORDOF["DANFE"] = df_CORDOF["NUMERO"].apply(lambda x: str(x).zfill(9) if len(str(x)) < 9 else str(x))  
        
        # Adicionar coluna DANFE_NFE  
        df_CORDOF["DANFE_NFE"] = df_CORDOF["NUMERO_NFE"].apply(lambda x: None if pd.isna(x) or "null" in str(x) else str(x).zfill(9) if len(str(x)) < 9 else str(x))  
        
        # Adicionar coluna DANFE_SERIE  
        df_CORDOF["DANFE_SERIE"] = df_CORDOF['DANFE'].astype(str) + '-' + df_CORDOF['SERIE_SUBSERIE'].astype(str)  
                                                
        filial_dict = {  
            "22086683000346": "K102",  
            "22086683000508": "K118",  
            "22086683000184": "K185",  
            "22086683000265": "K187",  
            "22086683000427": "K125"  
        }  
        
        
        df_CORDOF["FILIAL"] = df_CORDOF["INFORMANTE_EST_CODIGO"].astype(str).map(filial_dict)

        
        # Adicionar coluna MATERIAL  
        df_CORDOF["MATERIAL"] = df_CORDOF["MERC_CODIGO"].apply(lambda x: None if pd.isna(x) else "#".join([i for i in str(x).split(" ") if i != ""]))  
        
        # Adicionar coluna PLANTA  
        planta_dict = {  
            "22086683000346": "BR20",  
            "22086683000508": "BR18",  
            "22086683000184": "BR01",  
            "22086683000265": "K187",  
            "22086683000427": "K125"  
        }  

        df_CORDOF["INFORMANTE_EST_CODIGO"] = df_CORDOF["INFORMANTE_EST_CODIGO"].astype(str)  
        df_CORDOF["PLANTA"] = df_CORDOF["INFORMANTE_EST_CODIGO"].map(planta_dict)  
        
        # Adicionar coluna QUANTIDADE_AJUSTADA (QTD_AJJ)  
        df_CORDOF["QTD_AJJ"] = df_CORDOF.apply(lambda row: row["QTD"] * -1 if row["IND_ENTRADA_SAIDA"] == "E" else row["QTD"], axis=1)  
        
        # Filtrar linhas onde CTRL_SITUACAO_DOF é "N"  
        df_CORDOF = df_CORDOF[df_CORDOF["CTRL_SITUACAO_DOF"] == "N"]  
        
        # Filtrar linhas onde MATERIAL não é nulo ou qualquer um dos valores especificados  
        df_CORDOF = df_CORDOF[~df_CORDOF["MATERIAL"].isin([None, "", "MUC", "null", "ATIVO"])]  
        
        # Agrupar por HPB_OB_DEL_NUM e calcular a soma de QTD_AJJ  
        delivery_grouped = df_CORDOF.groupby("HPB_OB_DEL_NUM", as_index=False).agg({"QTD_AJJ": "sum"}).rename(columns={"QTD_AJJ": "QNT_DELIVERY"})  
        
        # Filtrar DELIVERIES com QNT_DELIVERY igual a 0  
        delivery_zeradas = delivery_grouped[delivery_grouped["QNT_DELIVERY"] == 0]  
        delivery_zeradas_lista = delivery_zeradas["HPB_OB_DEL_NUM"].tolist()  
         
        # Adicionar coluna DELIVERY_ZERADA  
        df_CORDOF["DELIVERY_ZERADA"] = df_CORDOF["HPB_OB_DEL_NUM"].apply(lambda x: "X" if x in delivery_zeradas_lista else None)  
        
        # Filtrar linhas onde CFOP_CODIGO_1 não contém os valores especificados  
        cfop_codigo_excluir = ["5.923", "6.923",]  
        
        df_CORDOF = df_CORDOF[~df_CORDOF["CFOP_CODIGO"].isin(cfop_codigo_excluir)]  
        
        # Adicionar coluna CHAVE_RECON  
        df_CORDOF["CHAVE_RECON"] = df_CORDOF.apply(lambda row: f"{row['PLANTA']} | {row['DANFE']} | {row['MATERIAL']}", axis=1)  

        # Agrupar por DANFE e calcular a soma de QTD_AJJ  
        danfe_grouped = df_CORDOF.groupby("CHAVE_RECON", as_index=False).agg({"QTD_AJJ": "sum"}).rename(columns={"QTD_AJJ": "QTD_GROUPED"})  

        # Filtrar DANFEs com QTD_GROUPED igual a 0  
        danfes_zeradas = danfe_grouped[danfe_grouped["QTD_GROUPED"] == 0]  
        danfes_zeradas_lista = danfes_zeradas["CHAVE_RECON"].tolist()  

        # Adicionar coluna DANFE_ZERADA  
        df_CORDOF["DANFE_ZERADA"] = df_CORDOF["CHAVE_RECON"].apply(lambda x: "X" if x in danfes_zeradas_lista else None) 
  
        # Criar a tabela dinâmica para somar a Quantity por Movement type  
        pivot_table = df_MB51.pivot_table(values='Quantity', index='Movement type', aggfunc='sum')  
    
        # Identificar os Movement types que se zeram  
        zero_movements = pivot_table[pivot_table['Quantity'] == 0].index.tolist() 

        df_CORDOF_grouped = df_CORDOF.groupby("CHAVE_RECON").agg({  
            "CTRL_SITUACAO_DOF": "max",  
            "NOP_CODIGO": "max",  
            "MATERIAL": "max",  
            "IND_ENTRADA_SAIDA": "max",  
            "QTD_AJJ": "sum",  
            "DANFE": "max",  
            "HPB_OB_DEL_NUM": "max",  
            "DANFE_ZERADA": "max",  
            "DELIVERY_ZERADA": "max"
        }).reset_index() 

        df_CORDOF_grouped = df_CORDOF_grouped.rename(columns={  
            "MATERIAL": "PN_CORDOF",  
            "QTD_AJJ": "QTD_CORDOF",  
            "DANFE": "DANFE_CORDOF",  
            "HPB_OB_DEL_NUM": "DELIVERY_CORDOF"  
        })


        ########################################### MB51 ############################################################
    
        # Adicionar a coluna Mov Zero ao df_MB51  
        df_MB51['Mov Zero'] = df_MB51['Movement type'].apply(lambda x: 'Mov Zero' if x in zero_movements else '')  
    
        # Limpar e padronizar colunas em df_MB51  
        cols_to_clean_MB51 = ['Plant', 'Material', 'Purchase order', 'Reference',  
                            'Movement type', 'Mov Zero', 'Document Header Text',  
                            'Storage location', 'Material Description', 'Item',  
                            'Material Doc.Item', 'Material Document', 'Document Date',  
                            'Entry Date', 'Receipt Indicator', 'Goods Receipt/Issue Slip',  
                            'Movement Type Text', 'Bill of Lading', 'Time of Entry',  
                            'Posting Date', 'User Name', 'Unit of Entry', 'Customer',  
                            'Supplier', 'Name 1', 'Profit Center']  
    
        for col in cols_to_clean_MB51:  
            df_MB51[col] = df_MB51[col].apply(clean_string)   
    
        # Tratamento para colunas de data  
        date_columns = ['Document Date', 'Entry Date', 'Posting Date']  
        for date_col in date_columns:  
            df_MB51[date_col] = pd.to_datetime(df_MB51[date_col], errors='coerce')     
    
        # Criar CHAVE_BASE no df_MB51 para agrupamento 
        df_MB51['CHAVE_GROUPED'] = df_MB51.apply(  
            lambda row: f"{row['Plant']} | {row['Material']} | {row['Purchase order']} | {row['Material Document']}"   
            if row['Movement type'] in ['101', '102', '861', '862']  
            else f"{row['Plant']} | {row['Material']} | {row['Reference']} | {row['Material Document']}",  
            axis=1  
        )  
    
        # Limpar e padronizar colunas em df_ZNFREP  
        cols_to_clean_ZNFREP = ['Plant', 'Material', 'Purchase Order', 'Delivery Number']  
        for col in cols_to_clean_ZNFREP:  
            df_ZNFREP[col] = df_ZNFREP[col].apply(clean_string) 
                
        # Filtrar apenas os registros com NF Status "Authorized"  
        df_ZNFREP = df_ZNFREP[df_ZNFREP["NF Status"] == "Authorized"]

        # Ajustando a coluna CFOP para remover o ponto e reconhecer os primeiros 4 dígitos  
        df_ZNFREP['CFOP'] = df_ZNFREP['CFOP'].str.replace('/','').str[:4] 

        cfop_remessa_venda_ordem = [  
            '5923', '6923'  
        ]   

        # Filtrando os registros onde "NF Status" não está na lista de CFOPs de remessa  
        df_ZNFREP = df_ZNFREP[~df_ZNFREP["CFOP"].isin(cfop_remessa_venda_ordem)] 
 
        # Criar CHAVE_BASE no df_ZNFREP (sem Quantity)  
        df_ZNFREP['CHAVE_GROUPED'] = df_ZNFREP.apply(  
            lambda row: f"{row['Plant']} | {row['Material']} | {row['Purchase Order']} | {row['Nfe/NFSe#']}"  
            if row['CFOP'] in ['1102/01']  
            else f"{row['Plant']} | {row['Material']} | {row['Delivery Number']} | {row['Nfe/NFSe#']}",  
            axis=1  
        )  
    
        # Agrupar df_MB51 por CHAVE_BASE  
        df_MB51_grouped = df_MB51.groupby('CHAVE_GROUPED').agg({   
            
            'Plant': 'max',  
            'Storage location': 'max',  
            'Material': 'max',  
            'Material Description': 'max',  
            'Item' : 'max',  
            'Material Doc.Item': 'max',  
            'Material Document': 'max',  
            'Document Date': 'max',  
            'Entry Date': 'max',  
            'Reference': 'max',  
            'Purchase order': 'max',  
            'Receipt Indicator': 'max',  
            'Goods Receipt/Issue Slip': 'max',  
            'Movement type': 'max',  
            'Movement Type Text': 'max',  
            'Bill of Lading': 'max',  
            'Time of Entry': 'max',  
            'Posting Date': 'max',  
            'User Name': 'max',  
            'Unit of Entry': 'max',  
            'Quantity': 'sum',  
            'Customer': 'max',  
            'Supplier': 'max',  
            'Name 1': 'max',  
            'Document Header Text': 'max',  
            'Profit Center': 'max',  
    
            'Mov Zero': 'max'  
    
        }).reset_index()  
    
        # Agrupar df_ZNFREP por CHAVE_BASE  
        df_ZNFREP_grouped = df_ZNFREP.groupby('CHAVE_GROUPED').agg({
            'Document Num': 'max',
            'Quantity': 'sum',  
            'Nfe/NFSe#': 'max',  
            'Plant': 'max',  
            'Material': 'max',  
            'Purchase Order': 'max',  
            'Delivery Number': 'max',  
            'CFOP': 'max'  
        }).reset_index()

        df_MB51_grouped['CHAVE_MERGED'] = df_MB51_grouped.apply(  
            lambda row: f"{row['Plant']} | {row['Material']} | {row['Purchase order']} | {int(abs(row['Quantity']))}"   
            if row['Movement type'] in ['101', '102', '861', '862']  
            else f"{row['Plant']} | {row['Material']} | {row['Reference']} | {int(abs(row['Quantity']))}",  
            axis=1  
        ) 

        # Criar CHAVE_BASE no df_ZNFREP 
        df_ZNFREP_grouped['CHAVE_MERGED'] = df_ZNFREP_grouped.apply(  
            lambda row: f"{row['Plant']} | {row['Material']} | {row['Purchase Order']} | {int(abs(row['Quantity']))}"  
            if row['CFOP'] in ['1102/01']  
            else f"{row['Plant']} | {row['Material']} | {row['Delivery Number']} | {int(abs(row['Quantity']))}",  
            axis=1  
        )

        # Passo 9: Agregar df_ZNFREP_grouped para tornar CHAVE_MERGED única  
        df_ZNFREP_aggregated = df_ZNFREP_grouped.groupby('CHAVE_MERGED').agg({  
            'Nfe/NFSe#': lambda x: ','.join(x.dropna().astype(str))  
        }).reset_index()  
        
        # Passo 10: Criar a tabela dinâmica para somar a Quantity por CHAVE_MERGED  
        pivot_table_merged = df_MB51_grouped.pivot_table(  
            values='Quantity',  
            index='CHAVE_MERGED',  
            aggfunc='sum'  
        )  
        
        # Passo 11: Identificar os CHAVE_MERGED que se zeram  
        zero_merged = pivot_table_merged[pivot_table_merged['Quantity'] == 0].index.tolist()  
        
        # Passo 12: Adicionar a coluna "MERGED ZERO" ao df_MB51_grouped  
        df_MB51_grouped['MERGED ZERO'] = df_MB51_grouped['CHAVE_MERGED'].apply(  
            lambda x: 'MERGED ZERO' if x in zero_merged else ''  
        )  
        
        ### PASSOS PARA O MAPEAMENTO ###  
        
        # Passo 13: Inicializar as colunas 'Nfe/NFSe#' e 'Nfe/NFSe#_Temp'  
        df_MB51_grouped['Nfe/NFSe#'] = ''  
        df_MB51_grouped['Nfe/NFSe#_Temp'] = ''  
        
        # Passo 14: Regra 1 - Se 'Document Header Text' contém 'CONV', então 'Conversão'  
        df_MB51_grouped.loc[  
            df_MB51_grouped['Document Header Text'].str.contains('CONV', case=False, na=False),  
            'Nfe/NFSe#_Temp'  
        ] = 'Conversão'  
        
        # Passo 15: Regra 2 - Se 'Mov Zero' é 'Mov Zero', então 'Mov Zero'  
        df_MB51_grouped.loc[  
            df_MB51_grouped['Mov Zero'] == 'Mov Zero',  
            'Nfe/NFSe#_Temp'  
        ] = 'Mov Zero'  
        
        # Passo 15: Regra 2 - Se 'Mov Zero' é 'Mov Zero', então 'Mov Zero'  
        df_MB51_grouped.loc[  
            df_MB51_grouped['MERGED ZERO'] == 'MERGED ZERO',  
            'Nfe/NFSe#_Temp'  
        ] = 'MERGED ZERO'  

        # Passo 16: Regra 3 - Se 'Quantity' é zero, então 'Qnt Zero'  
        df_MB51_grouped.loc[  
            df_MB51_grouped['Quantity'] == 0,  
            'Nfe/NFSe#_Temp'  
        ] = 'Qnt Zero'  

        df_MB51_grouped['Document Header Text'] = df_MB51_grouped['Document Header Text'].str.upper().replace('NAN', '')  
        
        # Passo 17: Regra 4 - Preencher NaN em 'Document Header Text' com string vazia antes de atribuir  
        df_MB51_grouped['Document Header Text'] = df_MB51_grouped['Document Header Text'].fillna('').replace('nan', '')  

        
        # Passo 17 (continuado): Aplicar Regra 4  
        condicao = df_MB51_grouped['Nfe/NFSe#_Temp'].isnull() | (df_MB51_grouped['Nfe/NFSe#_Temp'] == '')  
        df_MB51_grouped.loc[condicao, 'Nfe/NFSe#_Temp'] = df_MB51_grouped.loc[condicao, 'Document Header Text']  
        
        # Passo 18: Substituir a coluna 'Nfe/NFSe#' pela coluna temporária  
        df_MB51_grouped['Nfe/NFSe#'] = df_MB51_grouped['Nfe/NFSe#_Temp']  
        
        # Passo 19: Remover a coluna temporária 'Nfe/NFSe#_Temp'  
        df_MB51_grouped.drop(columns=['Nfe/NFSe#_Temp'], inplace=True)  
        
        # Passo 20: Criar o dicionário de mapeamento a partir de df_ZNFREP_aggregated  
        mapping_dict = df_ZNFREP_aggregated.set_index('CHAVE_MERGED')['Nfe/NFSe#'].to_dict()  
        
        # Passo 21: Aplicar o mapeamento usando .map para preencher 'Nfe/NFSe#' onde necessário  
        df_MB51_grouped['Nfe/NFSe#_ZNFREP'] = df_MB51_grouped['CHAVE_MERGED'].map(mapping_dict)  
        
        # Passo 22: Decidir qual valor de 'Nfe/NFSe#' utilizar (priorizar valores existentes e preencher com mapeamento)  
        df_MB51_grouped['Nfe/NFSe#'] = df_MB51_grouped['Nfe/NFSe#'].replace('', pd.NA)  
        df_MB51_grouped['Nfe/NFSe#'] = df_MB51_grouped['Nfe/NFSe#'].fillna(df_MB51_grouped['Nfe/NFSe#_ZNFREP'])  
        
        # Passo 23: Remover a coluna auxiliar 'Nfe/NFSe#_ZNFREP'  
        df_MB51_grouped.drop(columns=['Nfe/NFSe#_ZNFREP'], inplace=True)   
    
        # Opcional: Remover a coluna 'Mov Zero' se não for mais necessária  
        df_MB51_grouped.drop(columns=['Mov Zero'], inplace=True) 

        # Opcional: Remover a coluna 'MERGED ZERO' se não for mais necessária  
        df_MB51_grouped.drop(columns=['MERGED ZERO'], inplace=True) 
    
        # **Correção para Tratar "_" na Coluna 'Nfe/NFSe#'**  
    
        # 1. Criar a nova coluna 'DI_Number' inicialmente vazia  
        df_MB51_grouped['DI_Number'] = ''  
    
        # Identificar as linhas onde 'Nfe/NFSe#' contém '_'  
        mask = df_MB51_grouped['Nfe/NFSe#'].str.contains('_', na=False)  
        
        # Separar os valores antes e depois do '_'  
        split_columns = df_MB51_grouped.loc[mask, 'Nfe/NFSe#'].str.split(pat='_', n=1, expand=True)  
        
        # Atribuir as partes separadas às colunas correspondentes  
        if split_columns.shape[1] == 2:  
            df_MB51_grouped.loc[mask, 'Nfe/NFSe#'] = split_columns[0]  
            df_MB51_grouped.loc[mask, 'DI_Number'] = split_columns[1]  
        else:  
            print("Aviso: Algumas linhas com '_' não foram divididas corretamente.") 
    
        # Função para padronizar a coluna 'Nfe/NFSe#' similar ao VBA  
        def padronizar_nfe_nfse(valor):  
            valor_str = str(valor)  
    
            if '-' in valor_str:  
                partes = valor_str.split('-', 1)  
                parte_num = partes[0]  
                parte_restante = partes[1]  
                if parte_num.isdigit():  
                    parte_num_pad = parte_num.zfill(9)  
                    return f"{parte_num_pad}-{parte_restante}"  
                else:  
                    return valor_str  
            else:  
                if valor_str.isdigit():  
                    return valor_str.zfill(9)  
                else:  
                    return valor_str  
    
        def tratar_ajuste_scrap(valor):  
            if pd.isna(valor):  
                return valor  
    
            valor_str = str(valor)  
    
            #  Remover ".0" se presente  
            if valor_str.endswith('.0'):  
                valor_str = valor_str[:-2]  
    
            # Função auxiliar para formatar números  
            def formatar_numero(num):  
                if num.startswith('-'):  
                    num = num[1:]  # Remover o hífen temporariamente  
                return num.zfill(9)  # Preencher com zeros à esquerda  
    
            # Verificar e tratar "AJUSTE"  
            if "AJUSTE" in valor_str:  
                partes = valor_str.split()  
                numeros = [parte for parte in partes if parte.replace('-', '').isdigit()]  
                numeros_formatados = [formatar_numero(num) for num in numeros]  
                return ''.join(numeros_formatados)  
    
            # Verificar e tratar "SCRAP"  
            elif "SCRAP" in valor_str:  
                partes = valor_str.split()  
                numeros = [parte for parte in partes if parte.replace('-', '').isdigit()]  
                numeros_formatados = [formatar_numero(num) for num in numeros]  
                return ''.join(numeros_formatados)  
    
            else:  
                # Caso contrário, usar a função padronizar_nfe_nfse  
                return padronizar_nfe_nfse(valor_str)  
    
        # Aplicar as funções às colunas  
        df_MB51_grouped['Nfe/NFSe#'] = df_MB51_grouped['Nfe/NFSe#'].apply(tratar_ajuste_scrap)   
        df_MB51_grouped['Nfe/NFSe#'] = df_MB51_grouped['Nfe/NFSe#'].apply(padronizar_nfe_nfse) 

        # Atualizar a criação da CHAVE_RECON para incluir 'DI_Number' se necessário  
        df_MB51_grouped['CHAVE_RECON'] = df_MB51_grouped.apply(  
            lambda row: f"{row['Plant']} | {row['Nfe/NFSe#'].split('-')[0]} | {row['Material']}"   
            if pd.notna(row['Nfe/NFSe#'])   
            else f"{row['Plant']} | {row['Nfe/NFSe#']} | {row['Material']}",  
            axis=1  
        )  

        df_MB51_grouped = df_MB51_grouped.groupby("CHAVE_RECON").agg({  
            "Plant": "max",   
            "Movement type": "max",  
            "Material": "max",  
            "Quantity": "sum",  
            "Nfe/NFSe#": "max",  
            "Purchase order": "max",  
            "Reference": "max"  
        }).reset_index()

        df_MB51_grouped = df_MB51_grouped.rename(columns={  
            "Material": "PN_S4",  
            "Quantity": "QTD_S4",  
 
        })



        ########################################################## IGI #########################################################


        if igi_path:
            df_IGI = pd.read_excel(igi_path)

            planta_dict = {  
            "K102": "BR20",  
            "K118": "BR18", 
            "BR18": "BR18",
            "BR20": "BR20", 
        }  

            df_IGI["EMP_FIL_COD"] = df_IGI["EMP_FIL_COD"].astype(str)  
            df_IGI["PLANTA"] = df_IGI["EMP_FIL_COD"].map(planta_dict)  

            #Padronização do material #
            df_IGI["PN_IGI"] = df_IGI["PRD_COD"].apply(lambda x: None if pd.isna(x) else "#".join([i for i in str(x).split(" ") if i != ""]))  

            # Adicionar coluna CHAVE_RECON  
            df_IGI["CHAVE_RECON"] = df_IGI.apply(lambda row: f"{row['PLANTA']} | {row['MOV_NUM_DOC']} | {row['PN_IGI']}", axis=1)

            # Adicionar coluna QUANTIDADE_AJUSTADA
            df_IGI["QTD_IGI"] = df_IGI.apply(lambda row: row["MOV_QTD"] * -1 if row["TRN_OPR_EST"] == "-" else row["MOV_QTD"], axis=1)  

            df_IGI_grouped = df_IGI.groupby("CHAVE_RECON").agg({
                "EMP_FIL_COD": "max",
                "MOV_NUM_DOC": "max",
                "PN_IGI": "max",
                "QTD_IGI" : "sum"
            }).reset_index()

            chaves_igi = df_IGI_grouped[["CHAVE_RECON"]].drop_duplicates()

      

        ########################################################### RECON ########################################################
        
        chaves_mb51 = df_MB51_grouped[["CHAVE_RECON"]].drop_duplicates()  
        chaves_cordof = df_CORDOF_grouped[["CHAVE_RECON"]].drop_duplicates() 

        if igi_path:
            chaves_unidas = pd.concat([chaves_mb51, chaves_cordof, chaves_igi]).drop_duplicates().reset_index(drop=True) 
        else:
            chaves_unidas = pd.concat([chaves_mb51, chaves_cordof]).drop_duplicates().reset_index(drop=True)

        chaves_unidas["CHAVE_RECON"] = chaves_unidas["CHAVE_RECON"].str.strip()  
        chaves_unidas = chaves_unidas[chaves_unidas["CHAVE_RECON"].notna() & (chaves_unidas["CHAVE_RECON"] != "")]  

        if igi_path:
            df_RECON = chaves_unidas \
                .merge(df_MB51_grouped, on="CHAVE_RECON", how="left") \
                .merge(df_CORDOF_grouped, on="CHAVE_RECON", how="left") \
                .merge(df_IGI_grouped, on="CHAVE_RECON", how="left")
        else:
            df_RECON = chaves_unidas \
                .merge(df_MB51_grouped, on="CHAVE_RECON", how="left") \
                .merge(df_CORDOF_grouped, on="CHAVE_RECON", how="left")

        # Renomear colunas
        df_RECON = df_RECON.rename(columns={
            "Plant": "PLANTA_S4",
            "Movement type": "Movement type",
            "Material": "PN_S4",
            "Quantity": "QTD_S4",
            "Nfe/NFSe#": "DANFE_S4",
            "Purchase order": "Purchase Order_S4",
            "Reference": "Reference_S4"
        })

        # Calcular diferença de quantidade
        df_RECON["DIF_QNT"] = df_RECON.apply(
            lambda row: row["QTD_CORDOF"] if pd.isna(row.get("QTD_S4")) and not pd.isna(row.get("QTD_CORDOF")) else
                        row["QTD_S4"] if pd.isna(row.get("QTD_CORDOF")) and not pd.isna(row.get("QTD_S4")) else
                        row["QTD_S4"] + row["QTD_CORDOF"], axis=1)

        def get_trn(NOP_CODIGO):  
            trn_dict = {  
                "3101.01":"112",
                "3102.01":"112",
                "3102.02":"112",
                "3551.01":"NÃO ENTRA NO IGI",
                "3551.02":"NÃO ENTRA NO IGI",
                "3556.01":"NÃO ENTRA NO IGI - VERIFICAR MOVIMENTAÇÃO LÓGICA. CASOS DE COMPRA TRN 110/111",
                "3930.01":"NÃO ENTRA NO IGI",
                "3949.01":"320",
                "3949.02":"320",
                "3949.03":"214",
                "3949.04":"NÃO ENTRA NO IGI",
                "3949.04":"VERIFICAR SE OCORREU MOVIMENTAÇÃO LÓGICA",
                "3949.05":"NÃO ENTRA NO IGI",
                "3949.10":"VERIFICAR LOGICO",
                "3949.12":"112",
                "3949.15":"112",
                "3949.16":"NÃO ENTRA NO IGI",
                "3949.17":"318",
                "3949.53":"NÃO ENTRA NO IGI",
                "3949.54":"NÃO ENTRA NO IGI",
                "3949.59":"NÃO ENTRA NO IGI",
                "3949.90":"NÃO ENTRA NO IGI",
                "3949.97":"NÃO ENTRA NO IGI",
                "7101.01":"215",
                "7102.01":"215",
                "7949.03":"214",
                "7949.15":"215",
                "9990.01":"NÃO ENTRA NO IGI",
                "9999.01":"NÃO ENTRA NO IGI",
                "9999.02":"NÃO ENTRA NO IGI",
                "9999.03":"NÃO ENTRA NO IGI",
                "9999.08":"NÃO ENTRA NO IGI",
                "9999.11":"NÃO ENTRA NO IGI",
                "E101.01":"111",
                "E102.01":"111",
                "E102.02":"111AM",
                "E124.01":"319 / 111 - AM CODIGO 11 UTILIZAR TRN 111,  AM CODIGO 13 USAR TRN 110",
                "E152.01":"152",
                "E201.01":"120",
                "E202.01":"120",
                "E202.02":"120",
                "E204.01":"120",
                "E209.01":"150",
                "E352.01":"NÃO ENTRA NO IGI",
                "E353.01":"NÃO ENTRA NO IGI",
                "E401.01":"110 AM CODIGO 11/111 AM CODIGO 13",
                "E403.01":"110 AM CODIGO 11/111 AM CODIGO 13",
                "E406.01":"NÃO ENTRA NO IGI",
                "E407.01":"NÃO ENTRA NO IGI",
                "E409.01":"150",
                "E410.01":"120",
                "E411.01":"120",
                "E551.01":"NÃO ENTRA NO IGI",
                "E552.01":"NÃO ENTRA NO IGI",
                "E554.01":"NÃO ENTRA NO IGI",
                "E556.01":"NÃO ENTRA NO IGI - VERIFICAR MOVIMENTAÇÃO LÓGICA. CASOS DE COMPRA TRN 110/111",
                "E557.01":"NÃO ENTRA NO IGI",
                "E603.01":"N/A",
                "E604.01":"NÃO ENTRA NO IGI",
                "E902.01":"F0401",
                "E903.01":"F0401",
                "E906.01":"F0201",
                "E907.01":"F0201",
                "E908.01":"NÃO ENTRA NO IGI",
                "E909.01":"NÃO ENTRA NO IGI",
                "E910.01":"NÃO ENTRA NO IGI",
                "E911.01":"NÃO ENTRA NO IGI",
                "E912.01":"F0701",
                "E913.01":"F0701",
                "E914.01":"F0701",
                "E916.01":"194",
                "E916.02":"192",
                "E921.01":"NÃO ENTRA NO IGI",
                "E922.01":"NÃO ENTRA NO IGI",
                "E932.01":"NÃO ENTRA NO IGI",
                "E949.01":"320",
                "E949.02":"320",
                "E949.03":"F0701",
                "E949.05":"138",
                "E949.06":"194",
                "E949.07":"192",
                "E949.08":"192",
                "E949.09":"318",
                "E949.10":"NÃO ENTRA NO IGI",
                "E949.11":"192",
                "E949.12":"320",
                "E949.13":"320",
                "E949.14":"320",
                "E949.21":"193",
                "E949.30":"131",
                "E949.32":"320",
                "E949.33":"NÃO ENTRA NO IGI",
                "E949.51":"320",
                "E949.52":"NÃO ENTRA NO IGI",
                "E949.53":"NÃO ENTRA NO IGI",
                "E949.58":"320",
                "E949.60":"NÃO ENTRA NO IGI",
                "E949.80":"NÃO ENTRA NO IGI",
                "E949.80":"",
                "E949.86":"118",
                "E949.91":"NÃO ENTRA NO IGI",
                "E949.93":"NÃO ENTRA NO IGI",
                "E949.95":"NÃO ENTRA NO IGI",
                "S101.01":"211",
                "S102.01":"211",
                "S102.02":"211",
                "S106.01":"211",
                "S108.01":"211",
                "S110.01":"211",
                "S117.01 (2d V.F.)":"211",
                "S117.01 (2d V.F.)":"211",
                "S119.01":"211",
                "S152.01":"251",
                "S201.01":"221",
                "S202.01":"221",
                "S206.01":"NÃO ENTRA NO IGI",
                "S208.01":"251",
                "S209.01":"251",
                "S401.01":"211",
                "S402.01":"211",
                "S403.01":"211",
                "S408.01":"251",
                "S409.01":"251",
                "S411.01":"221",
                "S552.01":"NÃO ENTRA NO IGI",
                "S554.01":"NÃO ENTRA NO IGI",
                "S555.01":"NÃO ENTRA NO IGI",
                "S556.01":"NÃO ENTRA NO IGI - VERIFICAR MOVIMENTAÇÃO LÓGICA. TRN 231",
                "S557.01":"NÃO ENTRA NO IGI",
                "S602.01":"NÃO ENTRA NO IGI",
                "S901.01":"F0104",
                "S905.01":"F0102",
                "S908.01":"NÃO ENTRA NO IGI",
                "S909.01":"NÃO ENTRA NO IGI",
                "S910.01":"263",
                "S912.01":"F0107",
                "S913.01":"F0107",
                "S914.01":"F0107",
                "S915.01":"292",
                "S916.01":"292",
                "S920.01":"NÃO ENTRA NO IGI",
                "S922.01(1st V.F.)":"NÃO ENTRA NO IGI",
                "S923.01":"NÃO ENTRA NO IGI",
                "S923.01(2d V.O.)":"NÃO ENTRA NO IGI",
                "S927.01":"231",
                "S927.02":"267",
                "S933.01":"313",
                "S934.01":"NÃO ENTRA NO IGI",
                "S949.01":"211",
                "S949.02":"237",
                "S949.03":"F0107",
                "S949.04":"231 - VERIFICAR SE TEM LÓGICO - CASO TENHA É NECESSÁRIO COLOCAR NO IGI",
                "S949.05":"267",
                "S949.06":"NÃO ENTRA NO IGI",
                "S949.08":"292",
                "S949.09":"316",
                "S949.10":"NÃO ENTRA NO IGI, CASO OCORRA BAIXA SISTÊMICA DEVE COLOCAR NO IGI",
                "S949.11":"292",
                "S949.12":"214",
                "S949.13":"214",
                "S949.14":"214",
                "S949.17":"NÃO ENTRA NO IGI",
                "S949.21":"293",
                "S949.26":"293",
                "S949.28":"293",
                "S949.29":"293",
                "S949.30":"231",
                "S949.51":"292",
                "S949.52":"NÃO ENTRA NO IGI",
                "S949.53":"NÃO ENTRA NO IGI",
                "S949.58":"NÃO ENTRA NO IGI",
                "S949.60":"NÃO ENTRA NO IGI",
                "S949.61":"VERIFICAR LOGICO",
                "S949.80":"NÃO ENTRA NO IGI",
                "S949.85":"NÃO ENTRA NO IGI",
                "S949.93":"NÃO ENTRA NO IGI",
                "S605.01":"NÃO ENTRA NO IGI",
                "S120.01":"NÃO ENTRA NO IGI",
                "E118.01":"NÃO ENTRA NO IGI",
                "S949.84":"NÃO ENTRA NO IGI",
                "E949.85":"NÃO ENTRA NO IGI",
                "S949.81":"NÃO ENTRA NO IGI",
                "S910.02":"298",
                "S949.18":"231",
                "E949.T":"Entra no IGI - Sertrading",
                "S114.01":"Entra no IGI",
                "S949.82":"NÃO ENTRA NO IGI",
                "E253.01":"NÃO ENTRA NO IGI",
                "E920.01":"NÃO ENTRA NO IGI",

            }  
            return trn_dict.get(NOP_CODIGO, "TRN Desconhecido") 
        
        def get_notes(NOP_CODIGO):  
            notes_dict = {  
                "3101.01":"ENTRA NO IGI",
                "3102.01":"ENTRA NO IGI",
                "3102.02":"ENTRA NO IGI",
                "3551.01":"NÃO ENTRA NO IGI",
                "3551.02":"NÃO ENTRA NO IGI",
                "3556.01":"NÃO ENTRA NO IGI - VERIFICAR MOVIMENTAÇÃO LÓGICA. CASOS DE COMPRA TRN 110/111",
                "3930.01":"NÃO ENTRA NO IGI",
                "3949.01":"ENTRA NO IGI",
                "3949.02":"ENTRA NO IGI",
                "3949.03":"ENTRA NO IGI",
                "3949.04":"NÃO ENTRA NO IGI",
                "3949.04":"VERIFICAR SE OCORREU MOVIMENTAÇÃO LÓGICA",
                "3949.05":"NÃO ENTRA NO IGI",
                "3949.10":"SEGUNDO A SILVANA ENTRA NO IGI, VERIFICAR SE TEM MOVIMENTAÇÃO LOGICO",
                "3949.12":"ENTRA NO IGI / SEGUNDO A SILVANA NÃO ENTRA NO IGI - VERIFICAR SE TEM MOVIMENTAÇÃO LOGICO",
                "3949.15":"ENTRA NO IGI",
                "3949.16":"SEGUNDO A SILVANA NÃO ENTRA NO IGI",
                "3949.17":"SEGUNDO A SILVANA ENTRA NO IGI, VERIFICAR SE TEM MOVIMENTAÇÃO LOGICO",
                "3949.53":"SEGUNDO A SILVANA NÃO ENTRA NO IGI",
                "3949.54":"SEGUNDO A SILVANA NÃO ENTRA NO IGI",
                "3949.59":"SEGUNDO A SILVANA NÃO ENTRA NO IGI",
                "3949.90":"SEGUNDO A SILVANA NÃO ENTRA NO IGI",
                "3949.97":"SEGUNDO A SILVANA NÃO ENTRA NO IGI",
                "7101.01":"ENTRA NO IGI",
                "7102.01":"ENTRA NO IGI",
                "7949.03":"ENTRA NO IGI",
                "7949.15":"ENTRA NO IGI",
                "9990.01":"NÃO ENTRA NO IGI",
                "9999.01":"NÃO ENTRA NO IGI",
                "9999.02":"NÃO ENTRA NO IGI",
                "9999.03":"NÃO ENTRA NO IGI",
                "9999.08":"NÃO ENTRA NO IGI",
                "9999.11":"NÃO ENTRA NO IGI",
                "E101.01":"ENTRA NO IGI",
                "E102.01":"ENTRA NO IGI",
                "E102.02":"ENTRA NO IGI",
                "E124.01":"ENTRA NO IGI",
                "E152.01":"ENTRA NO IGI",
                "E201.01":"ENTRA NO IGI",
                "E202.01":"ENTRA NO IGI",
                "E202.02":"ENTRA NO IGI",
                "E204.01":"ENTRA NO IGI",
                "E209.01":"ENTRA NO IGI",
                "E352.01":"NÃO ENTRA NO IGI",
                "E353.01":"NÃO ENTRA NO IGI",
                "E401.01":"ENTRA NO IGI",
                "E403.01":"ENTRA NO IGI",
                "E406.01":"NÃO ENTRA NO IGI",
                "E407.01":"NÃO ENTRA NO IGI",
                "E409.01":"ENTRA NO IGI",
                "E410.01":"ENTRA NO IGI",
                "E411.01":"ENTRA NO IGI",
                "E551.01":"NÃO ENTRA NO IGI",
                "E552.01":"NÃO ENTRA NO IGI",
                "E554.01":"NÃO ENTRA NO IGI",
                "E556.01":"NÃO ENTRA NO IGI - VERIFICAR MOVIMENTAÇÃO LÓGICA. CASOS DE COMPRA TRN 110/111",
                "E557.01":"NÃO ENTRA NO IGI",
                "E603.01":"N/A",
                "E604.01":"NÃO ENTRA NO IGI",
                "E902.01":"ENTRA NO IGI",
                "E903.01":"ENTRA NO IGI",
                "E906.01":"ENTRA NO IGI",
                "E907.01":"ENTRA NO IGI",
                "E908.01":"NÃO ENTRA NO IGI",
                "E909.01":"NÃO ENTRA NO IGI",
                "E910.01":"NÃO ENTRA NO IGI",
                "E911.01":"NÃO ENTRA NO IGI",
                "E912.01":"ENTRA NO IGI",
                "E913.01":"ENTRA NO IGI",
                "E914.01":"ENTRA NO IGI",
                "E916.01":"ENTRA NO IGI",
                "E916.02":"ENTRA NO IGI",
                "E921.01":"NÃO ENTRA NO IGI",
                "E922.01":"NÃO ENTRA NO IGI",
                "E932.01":"NÃO ENTRA NO IGI",
                "E949.01":"ENTRA NO IGI",
                "E949.02":"ENTRA NO IGI",
                "E949.03":"NÃO ENTRA NO IGI  / VERIFICAR SE TEVE MOVIMENTAÇÃO NO LOGICO",
                "E949.05":"ENTRA NO IGI",
                "E949.06":"ENTRA NO IGI",
                "E949.07":"ENTRA NO IGI",
                "E949.08":"ENTRA NO IGI",
                "E949.09":"ENTRA NO IGI",
                "E949.10":"NÃO ENTRA NO IGI",
                "E949.11":"ENTRA NO IGI",
                "E949.12":"ENTRA NO IGI",
                "E949.13":"ENTRA NO IGI",
                "E949.14":"ENTRA NO IGI",
                "E949.21":"ENTRA NO IGI",
                "E949.30":"ENTRA NO IGI",
                "E949.32":"ENTRA NO IGI",
                "E949.33":"NÃO ENTRA NO IGI A MOVIMENTAÇÃO DE ESTOQUE ENTRA PELO NOP E949.32",
                "E949.51":"ENTRA NO IGI",
                "E949.52":"NÃO ENTRA NO IGI",
                "E949.53":"NÃO ENTRA NO IGI",
                "E949.58":"NÃO ENTRA NO IGI  / VERIFICAR SE TEVE MOVIMENTAÇÃO NO LOGICO",
                "E949.60":"NÃO ENTRA NO IGI",
                "E949.80":"NÃO ENTRA NO IGI",
                "E949.80":"",
                "E949.86":"ENTRA NO IGI",
                "E949.91":"NÃO ENTRA NO IGI",
                "E949.93":"NÃO ENTRA NO IGI",
                "E949.95":"NÃO ENTRA NO IGI",
                "S101.01":"ENTRA NO IGI",
                "S102.01":"ENTRA NO IGI",
                "S102.02":"ENTRA NO IGI",
                "S106.01":"ENTRA NO IGI",
                "S108.01":"ENTRA NO IGI",
                "S110.01":"ENTRA NO IGI",
                "S117.01 (2d V.F.)":"ENTRA NO IGI",
                "S117.01 (2d V.F.)":"ENTRA NO IGI",
                "S119.01":"TODOS OS CASOS DO SISTEMA VELOCITY ENTRARAM, FUSION NÃO ENTRARAM ",
                "S152.01":"ENTRA NO IGI",
                "S201.01":"ENTRA NO IGI",
                "S202.01":"ENTRA NO IGI",
                "S206.01":"NÃO ENTRA NO IGI",
                "S208.01":"ENTRA NO IGI",
                "S209.01":"ENTRA NO IGI",
                "S401.01":"ENTRA NO IGI",
                "S402.01":"ENTRA NO IGI",
                "S403.01":"ENTRA NO IGI",
                "S408.01":"ENTRA NO IGI",
                "S409.01":"ENTRA NO IGI",
                "S411.01":"ENTRA NO IGI",
                "S552.01":"NÃO ENTRA NO IGI",
                "S554.01":"NÃO ENTRA NO IGI",
                "S555.01":"NÃO ENTRA NO IGI",
                "S556.01":"NÃO ENTRA NO IGI - VERIFICAR MOVIMENTAÇÃO LÓGICA. TRN 231",
                "S557.01":"NÃO ENTRA NO IGI",
                "S602.01":"NÃO ENTRA NO IGI",
                "S901.01":"ENTRA NO IGI",
                "S905.01":"ENTRA NO IGI",
                "S908.01":"NÃO ENTRA NO IGI",
                "S909.01":"NÃO ENTRA NO IGI",
                "S910.01":"ENTRA NO IGI",
                "S912.01":"ENTRA NO IGI",
                "S913.01":"ENTRA NO IGI",
                "S914.01":"ENTRA NO IGI",
                "S915.01":"ENTRA NO IGI",
                "S916.01":"ENTRA NO IGI",
                "S920.01":"NÃO ENTRA NO IGI",
                "S922.01(1st V.F.)":"NÃO ENTRA NO IGI",
                "S923.01":"NÃO ENTRA NO IGI",
                "S923.01(2d V.O.)":"NÃO ENTRA NO IGI",
                "S927.01":"ENTRA NO IGI",
                "S927.02":"ENTRA NO IGI",
                "S933.01":"ENTRA NO IGI",
                "S934.01":"NÃO ENTRA NO IGI",
                "S949.01":"ENTRA NO IGI",
                "S949.02":"ENTRA NO IGI",
                "S949.03":"NÃO ENTRA NO IGI  / VERIFICAR SE TEVE MOVIMENTAÇÃO NO LOGICO",
                "S949.04":"VERIFICAR SE TEM LÓGICO - CASO TENHA É NECESSÁRIO COLOCAR NO IGI",
                "S949.05":"ENTRA NO IGI",
                "S949.06":"NÃO ENTRA NO IGI",
                "S949.08":"ENTRA NO IGI",
                "S949.09":"ENTRA NO IGI",
                "S949.10":"NÃO ENTRA NO IGI, CASO OCORRA BAIXA SISTÊMICA DEVE COLOCAR NO IGI",
                "S949.11":"ENTRA NO IGI",
                "S949.12":"ENTRA NO IGI",
                "S949.13":"ENTRA NO IGI",
                "S949.14":"ENTRA NO IGI",
                "S949.17":"NÃO ENTRA NO IGI",
                "S949.21":"ENTRA NO IGI",
                "S949.26":"ENTRA NO IGI",
                "S949.28":"ENTRA NO IGI",
                "S949.29":"ENTRA NO IGI",
                "S949.30":"ENTRA NO IGI",
                "S949.51":"ENTRA NO IGI",
                "S949.52":"NÃO ENTRA NO IGI",
                "S949.53":"NÃO ENTRA NO IGI",
                "S949.58":"NÃO ENTRA NO IGI",
                "S949.60":"NÃO ENTRA NO IGI",
                "S949.61":"NÃO ENTRA NO IGI - MICHELI",
                "S949.80":"NÃO ENTRA NO IGI",
                "S949.85":"NÃO ENTRA NO IGI",
                "S949.93":"NÃO ENTRA NO IGI",
                "S605.01":"NÃO ENTRA NO IGI",
                "S120.01":"NÃO ENTRA NO IGI",
                "E118.01":"NÃO ENTRA NO IGI",
                "S949.84":"NÃO ENTRA NO IGI",
                "E949.85":"NÃO ENTRA NO IGI",
                "S949.81":"NÃO ENTRA NO IGI",
                "S910.02":"ENTRA NO IGI",
                "S949.18":"ENTRA NO IGI",
                "E949.T":"Entra no IGI - Sertrading",
                "S114.01":"Entra no IGI",
                "S949.82":"Não entra",
                "E253.01":"Não entra",
                "E920.01":"Não entra",

            }  
            return notes_dict.get(NOP_CODIGO, "Notes Desconhecido")
        
        df_RECON['TRN'] = ''  
        df_RECON['NOTES'] = ''  
    
        # Aplicar a função e criar uma nova coluna 'TRN_Mapped'  
        df_RECON['TRN'] = df_RECON['NOP_CODIGO'].apply(get_trn) 

        # Aplicar a função e criar uma nova coluna 'TRN_Mapped'  
        df_RECON['NOTES'] = df_RECON['NOP_CODIGO'].apply(get_notes)

        # Constantes simbólicas para facilitar manutenção
        STATUS_MOV_ZERO = 'Mov Zero'
        STATUS_NAO_ENTRA_IGI = 'NÃO ENTRA NO IGI'
        STATUS_QNT_ZERO = 'Qnt Zero'
        STATUS_CHAVE_ZERADA = 'Chave Zerada'
        STATUS_CONVERSAO = 'Conversão'
        STATUS_DANFE_ZERADA = 'Danfe Zerada'
        STATUS_DELIVERY_ZERADA = 'Delivery Zerada'
        STATUS_APENAS_S4 = 'Apenas S4'
        STATUS_APENAS_CORDOF = 'Apenas CORDOF'
        STATUS_RECON_OK = 'Recon OK'
        STATUS_DIFERENCA = 'Totais com Diferença'

        def get_status(row: pd.Series) -> str:
            chave = row.get('CHAVE_RECON', '')
            notes = row.get('NOTES', '')
            danfe_zerada = row.get('DANFE_ZERADA', '')
            delivery_zerada = row.get('DELIVERY_ZERADA', '')
            pn_cordof = row.get('PN_CORDOF')
            pn_s4 = row.get('PN_S4')
            dif_qnt = row.get('DIF_QNT', None)

            if 'Mov Zero' in chave:
                return STATUS_MOV_ZERO
            elif notes == STATUS_NAO_ENTRA_IGI:
                return STATUS_NAO_ENTRA_IGI
            elif 'Qnt Zero' in chave:
                return STATUS_QNT_ZERO
            elif 'MERGED ZERO' in chave:
                return STATUS_CHAVE_ZERADA
            elif 'Conversão' in chave:
                return STATUS_CONVERSAO
            elif danfe_zerada == 'X':
                return STATUS_DANFE_ZERADA
            elif delivery_zerada == 'X':
                return STATUS_DELIVERY_ZERADA
            elif pd.isna(pn_cordof) and not pd.isna(pn_s4):
                return STATUS_APENAS_S4
            elif not pd.isna(pn_cordof) and pd.isna(pn_s4):
                return STATUS_APENAS_CORDOF
            elif dif_qnt == 0:
                return STATUS_RECON_OK
            else:
                return STATUS_DIFERENCA

        # Aplicação da função ao DataFrame
        df_RECON['STATUS'] = df_RECON.apply(get_status, axis=1)


        #COLUNA UNICA PARA MATERIAL        
        if igi_path:
            df_RECON['MATERIAL'] = (
            df_RECON['PN_S4']
            .combine_first(df_RECON['PN_CORDOF'])
            .combine_first(df_RECON['PN_IGI'])
        )

        else:
            df_RECON['MATERIAL'] = df_RECON['PN_S4'].combine_first(df_RECON['PN_CORDOF'])

        if igi_path:
            final_columns = ["CHAVE_RECON", "CTRL_SITUACAO_DOF", "IND_ENTRADA_SAIDA"
                    , "Movement type", "TRN", "NOP_CODIGO","NOTES",  "Purchase Order_S4", "Reference_S4",  
                    "MATERIAL", "QTD_IGI","QTD_S4", "QTD_CORDOF", "DIF_QNT","STATUS",  
                    "DELIVERY_CORDOF", "DELIVERY_ZERADA", "DANFE_ZERADA"]  
        else:
            final_columns = ["CHAVE_RECON", "CTRL_SITUACAO_DOF", "IND_ENTRADA_SAIDA"
                    , "Movement type", "TRN", "NOP_CODIGO","NOTES",  "Purchase Order_S4", "Reference_S4",  
                    "MATERIAL","QTD_S4", "QTD_CORDOF", "DIF_QNT","STATUS",  
                    "DELIVERY_CORDOF", "DELIVERY_ZERADA", "DANFE_ZERADA"] 
    
        df_RECON = df_RECON[final_columns] 
    
        output_file = os.path.join(output_folder, "RECON LÓGICOS.xlsx")  
        with pd.ExcelWriter(output_file) as writer:  
            # Salvar cada DataFrame em uma folha diferente  
            df_MB51_grouped.to_excel(writer, sheet_name='MB51_GROUPED', index=False)  
            df_ZNFREP_grouped.to_excel(writer, sheet_name='ZNFREP_GROUPED', index=False)  
            df_CORDOF_grouped.to_excel(writer, sheet_name='CORDOF_GROUPED', index=False)
            if igi_path:
                df_IGI_grouped.to_excel(writer, sheet_name='IGI_GROUPED', index=False)   
            df_RECON.to_excel(writer, sheet_name='RECON', index=False)
    
        print("Processamento concluído e arquivos salvos com sucesso.")    
  
if __name__ == "__main__":  
    app = Application()  
    app.mainloop()  
