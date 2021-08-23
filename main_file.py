from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import xlrd as x
import reportlab
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import qrcode
import os
import datetime as dt
import numpy as np
from pdfmerging import *

version = "6.0"
current_year = "2021"

pdfmetrics.registerFont(TTFont('ArialBd', 'ARIALBD.ttf'))

window_width = 600
window_height = 310

class Application:
    
    def __init__(self,master):
        
        Label(master,text="LabelGen v. " + version + "\nElaborado por Igor Rêgo").pack(side=BOTTOM)

        labeling_tab = Frame(master)
        labeling_tab.pack(expand=True,fill=BOTH)

        # Main frames

        self.optframe = LabelFrame(labeling_tab,
            text="Opções",
            relief=GROOVE,
            width=0.46*window_width,
            height=window_height
        )
        self.optframe.pack(side=LEFT,padx=10,pady=10,expand=True,fill=BOTH)
        self.optframe.pack_propagate(False)
                
        self.dataframe = LabelFrame(labeling_tab,
            text="Dados",
            relief=GROOVE,
            width=0.46*window_width,
            height=window_height
        )
        self.dataframe.pack(side=RIGHT,padx=10,pady=10,expand=True,fill=BOTH)
        self.dataframe.pack_propagate(False)

        # Sub-frame de labels

        self.lblframe = Frame(self.dataframe,relief=GROOVE,width=110,height=250)
        self.lblframe.pack(side=LEFT,padx=5,pady=10)
        self.lblframe.pack_propagate(False)

        # Cód. (cliente)
        self.hdr_c = Label(self.lblframe,text="Cód. cliente:")        
        self.hdr_c.pack(pady=5)

        # Lote/OP
        self.hdr_s = Label(self.lblframe,text="Lote & OP:")
        self.hdr_s.pack(pady=5)

        # Qtd. por caixa
        self.hdr_q = Label(self.lblframe,text="Pçs./cx.:")
        self.hdr_q.pack(pady=5)

        # Validade
        self.hdr_e = Label(self.lblframe,text="Data:")
        self.hdr_e.pack(pady=5)

        # Nome do produto
        self.hdr_p = Label(self.lblframe,text="Prod.:")
        self.hdr_p.pack(pady=5)

        # Selecionando tipo de etiqueta
        self.labeltype = StringVar(None)
        self.qrlabel = Radiobutton(self.lblframe, text="QR Code", variable = self.labeltype, value="qr")
        self.qrlabel.pack(pady=10)
        self.productionlabel = Radiobutton(self.lblframe, text="Produção", variable = self.labeltype, value="production")
        self.productionlabel.pack()

        # Sub-frame de valores lidos

        self.valframe = Frame(self.dataframe,relief=GROOVE,width=110,height=250)
        self.valframe.pack(side=RIGHT,padx=5,pady=10)
        self.valframe.pack_propagate(False)

        self.overrideButton = Button(self.valframe,width=15,text="Editar",command=self.overrideValues)
        self.overrideButton.pack(side=BOTTOM,padx=5,pady=10)
        
        # Frame de opções

        self.getFileButton = Button(self.optframe, text="Abrir OP", width = 15, command=self.getFile)
        self.getFileButton.pack()

        self.printButton = Button(self.optframe, text="Atualizar", width = 15, command=self.printMessage)
        self.printButton.pack()

        self.generateButton = Button(self.optframe, text="Gerar PDF", width = 15, command=self.generatePdf)
        self.generateButton.pack()

        self.clearButton = Button(self.optframe, text="Limpar", width = 15, command=self.clearLabels)
        self.clearButton.pack()

        Label(self.optframe, text = "Pçs. por caixa:").pack()
        self.qtyperboxEntry = Entry(self.optframe)
        self.qtyperboxEntry.pack()

        Label(self.optframe, text = "Cód. prod. (cliente):").pack()
        self.customcodeEntry = Entry(self.optframe)
        self.customcodeEntry.pack()
        
        self.quitButton = Button(self.optframe, text="Sair", width = 15, fg = "red", command=self.optframe.quit)
        self.quitButton.pack(pady=10)

        self.aboutButton = Button(self.optframe, text="Sobre", width = 15, command = self.showAbout)
        self.aboutButton.pack(side=BOTTOM,pady=10)
        
        # Placeholders de memória

        self.filename = None
        
        self.customcode = None
        self.intcode = None
        self.prodname = None
        self.serialno = None
        self.qtyperbox = 0
        self.color = None

        self.date = None
        self.expdate = None
        self.exptime = 2
        
        self.nocopies = 1

        self.suppliercode = '003654'

        self.p = None       
        self.s = None
        self.q = None          
        self.e = None        
        self.c = None      
        self.l = None
        self.k = None

        # Toggles de segurança

        self.editdata = False # Para bypassar o check de getters do self.qtyperbox e self.customcode
        self.check_int_toggle = False # Para garantir que somente números "int" entrem em self.qtyperbox
        self.haserror = False # Check para avaliar erros
        
    # Métodos 

    def compliantQr(self):

        self.expdate = self.determineExpDate()
        print("L> compliantQr > dataqr > self.expdate > " + str(self.expdate))
        dataqr = [
            '01',
            self.suppliercode,
            self.customcode,
            self.expdate,
            self.qtyperbox + ',0000',
            'UN',
            self.serialno
        ]

        compliant = ""
        if len(str(self.qtyperbox)) >= 1:
            compliant = str(dataqr).replace("\'","").replace(" ","").replace(",",";",4)[:35+len(str(self.qtyperbox))] + str(dataqr).replace("\'","").replace(" ","").replace(",",";")[35+len(str(self.qtyperbox)):len(str(dataqr))]
        
        qr = qrcode.make(compliant[1:-1])
        imgsrc = "QR_" + self.serialno + ".png"
        qr.save(imgsrc)
        return imgsrc

    def printMessage(self):
        
        self.haserror = False
        if self.editdata == False:
            if self.check_int_toggle == False:
                self.qtyperbox = self.qtyperboxEntry.get()
            else:
                try:
                    self.qtyperbox = self.qtyperboxEntry.get()
                    print(int(self.qtyperbox))
                except ValueError:
                    messagebox.showerror("Erro","Formato de número inválido.")
                    self.haserror = True
            self.customcode = self.customcodeEntry.get()
        
        self.clearLabels()

        if len(self.filename) > 20:
            filename = self.filename[0:40] + "..."
        else:
            filename = self.filename
        self.l = Label(self.optframe,text=filename)
        self.l.pack()

        # Código (cliente)
        self.c = Label(self.valframe,text=self.customcode)
        self.c.pack(pady=5)

        # Lote/OP
        self.s = Label(self.valframe,text=self.serialno)
        self.s.pack(pady=5)
        
        # Qtd. por caixa
        try:
            if int(self.qtyperbox) <= 999:
                self.q = Label(self.valframe,text=self.qtyperbox)
            elif int(self.qtyperbox) >= 1000:
                self.q = Label(self.valframe,text= self.qtyperbox[0:len(self.qtyperbox)-3] + "." + self.qtyperbox[len(self.qtyperbox)-3:len(self.qtyperbox)])
        except ValueError:
            self.q = Label(self.valframe,text="VALUE_ERROR",fg="red")
            self.haserror = True
        self.q.pack(pady=5)

        # Validade
        self.e = Label(self.valframe,text=self.date)
        self.e.pack(pady=5)
       
        # Nome do produto
        if len(self.prodname) > 20:
            self.lblprodname = self.prodname[0:20] + "..."
        else:
            self.lblprodname  = self.prodname
        self.p = Label(self.valframe,text=self.lblprodname,wraplength=100)
        self.p.pack(pady=5)
    
    def determineExpDate(self):

        day = int(self.date[0:2])
        month = int(self.date[3:5])
        year = int(self.date[6:10])
        prod_date = dt.datetime(year,month,day)
        delta_time = dt.timedelta(days=int(self.exptime)*365)
        exp = prod_date + delta_time
        print("L> determineExpDate > exp.strftime >" + str(exp.strftime("%d/%m/%Y")))
        return exp.strftime('%d/%m/%Y')
                   
    def getFile(self):

        self.filename = filedialog.askopenfilename(
            title = "Selecionar arquivo da OP",
            initialdir = "/",
            filetypes = (
                ("Pastas de trabalho do Excel", "*.xls"),
                ("Todos os arquivos", "*.*"),
            ),
        )
        
        try:
            wb = x.open_workbook(self.filename)
            sheet = wb.sheet_by_index(0)

            self.prodname = sheet.cell_value(11,4)
            self.serialno = sheet.cell_value(7,11)
            self.date = sheet.cell_value(9,5)
            self.expdate = self.determineExpDate()
            print("L> getFile > self.expdate > " + str(self.expdate))
            self.intcode = sheet.cell_value(11,1)
            
            self.haserror = False

            self.printMessage()
            self.maybeCustomcode()
                                  
        except:
            messagebox.showerror("Erro","Não foi possível ler o arquivo importado.")

    def clearLabels(self):
        
        self.qtyperboxEntry.delete(0,END)
        self.customcodeEntry.delete(0,END)
       
        try:

            self.p.destroy()
            self.s.destroy()
            self.q.destroy()
            self.e.destroy()
            self.c.destroy()
            self.l.destroy()
            self.k.destroy()

        except:
            pass
    
    def maybeCustomcode(self):
        maybecustomcode = self.prodname.find("PT-")
        if maybecustomcode == -1:
            maybecustomcode = self.prodname.find("TP-")
            if maybecustomcode == -1:
                maybecustomcode = self.prodname.find("FR-")
                if maybecustomcode == -1:
                    pass
                else:
                    self.customcodeEntry.insert(0,self.prodname[maybecustomcode:maybecustomcode+8])
            else:
                self.customcodeEntry.insert(0,self.prodname[maybecustomcode:maybecustomcode+8])
        else:
            self.customcodeEntry.insert(0,self.prodname[maybecustomcode:maybecustomcode+8])
    
    def showAbout(self):
        aboutwindow = Toplevel(self.optframe)
        aboutwindow.title("Sobre")
        aboutwindow.geometry("300x400")
        aboutwindow.resizable(False,False)

        aframe = LabelFrame(aboutwindow,
            text="Sobre o programa",
            relief=GROOVE,
            )
        aframe.pack(side=TOP,padx=10,pady=10,expand=True,fill=BOTH)
        aframe.pack_propagate(False)

        Label(aframe,text="Software LabelGen, versão " + version + ", para a geração descomplicada de etiquetas (com ou sem QR). Elaborado por Igor Rêgo.\n©" + current_year, wraplength=270).pack()

        hframe = LabelFrame(aboutwindow,
            text="Ajuda",
            relief=GROOVE,
           )
        hframe.pack(side=TOP,padx=10,expand=True,fill=BOTH)
        hframe.pack_propagate(False)

        def howTo():
            os.startfile(os.getcwd() + '\\' + "howto.pdf")

        Label(hframe,text="Dúvidas, sugestões, feedback e apontamento de bugs devem ser encaminhados para:\nigor.rego@kipack.com.br\n+55 11 4135 2757 R: 214", wraplength=270).pack()
        Button(hframe,text="Como usar",command=howTo,width=15).pack(padx=10)

        bframe = Frame(aboutwindow)
        bframe.pack(side=BOTTOM,expand=True,fill=BOTH)
        bframe.pack_propagate(False)

        okbutton = Button(bframe,text="OK",width=10,command=aboutwindow.destroy)
        okbutton.pack(side=BOTTOM,expand=True)
        okbutton.pack_propagate(False)
            
    def overrideValues(self):

        self.overwindow = Toplevel(self.valframe)
        self.overwindow.title("Editar")
        self.overwindow.geometry("280x380")
        self.overwindow.resizable(False,False)

        mainframe = LabelFrame(self.overwindow,text="Valores",relief=GROOVE)
        mainframe.pack(side=TOP,padx=10,pady=10,expand=True,fill=BOTH)

        # Cód. (cliente)       
        Label(mainframe,text="Cód. prod. (cliente):").grid(row=0,column=0)
        self.edit_custom = Entry(mainframe)
        if self.customcode == None:
            self.edit_custom.insert(0,"Não especificado")
        else:
            self.edit_custom.insert(0,self.customcode)
        self.edit_custom.grid(row=0,column=1,pady=5,padx=5)

        # Cód. interno
        Label(mainframe,text="Cód. prod. (interno):").grid(row=1,column=0)
        self.edit_intcode = Entry(mainframe)
        if self.intcode == None:
            self.edit_intcode.insert(0,"Não especificado")
        else:
            self.edit_intcode.insert(0,self.intcode)
        self.edit_intcode.grid(row=1,column=1,pady=5,padx=5)
        
        # Lote/OP
        Label(mainframe,text="Lote & OP:").grid(row=2,column=0)
        self.edit_serialno = Entry(mainframe)
        if self.serialno == None:
            self.edit_serialno.insert(0,"Não especificado")
        else:
            self.edit_serialno.insert(0,self.serialno)
        self.edit_serialno.grid(row=2,column=1,pady=5,padx=5)

        # Qtd. por caixa
        Label(mainframe,text="Pçs./cx.:").grid(row=3,column=0)
        self.edit_qtyperbox = Entry(mainframe)
        if self.qtyperbox == None or 0:
            self.edit_qtyperbox.insert(0,"Não especificado")
        else:
            self.edit_qtyperbox.insert(0,self.qtyperbox)
        self.edit_qtyperbox.grid(row=3,column=1,pady=5,padx=5)

        # Data
        Label(mainframe,text="Data:").grid(row=4,column=0)
        self.edit_date = Entry(mainframe)
        if self.date == None:
            self.edit_date.insert(0,"Não especificado")
        else:
            self.edit_date.insert(0,self.date)
        self.edit_date.grid(row=4,column=1,pady=5,padx=5)

        # Tempo de validade
        Label(mainframe,text="Val. (anos):").grid(row=5,column=0)
        self.edit_exptime = Entry(mainframe)
        self.edit_exptime.insert(0,self.exptime)
        self.edit_exptime.grid(row=5,column=1,pady=5,padx=5)

        # Nome do produto
        Label(mainframe,text="Nome do prod.:").grid(row=6,column=0)
        self.edit_prodname = Entry(mainframe)
        if self.prodname == None:
            self.edit_prodname.insert(0,"Não especificado")
        else:
            self.edit_prodname.insert(0,self.prodname)
        self.edit_prodname.grid(row=6,column=1,pady=5,padx=5)

        # Código do fornecedor
        Label(mainframe,text="Cód. de fornecedor:").grid(row=7,column=0)
        self.edit_suppliercode = Entry(mainframe)
        if self.suppliercode == None:
            self.edit_suppliercode.insert(0,"Não especificado")
        else:
            self.edit_suppliercode.insert(0,self.suppliercode)
        self.edit_suppliercode.grid(row=7,column=1,pady=5,padx=5)

        # Cor
        Label(mainframe,text="Cor:").grid(row=8,column=0)
        self.edit_color = Entry(mainframe)
        if self.color == None:
            self.edit_color.insert(0,"Não especificado")
        else:
            self.edit_color.insert(0,self.color)
        self.edit_color.grid(row=8,column=1,pady=5,padx=5)

        # Número de cópias
        Label(mainframe,text="N.º de cópias:").grid(row=9,column=0)
        self.edit_nocopies = Entry(mainframe)
        self.edit_nocopies.insert(0,self.nocopies)
        self.edit_nocopies.grid(row=9,column=1)

        buttonframe = Frame(self.overwindow)
        buttonframe.pack(side=BOTTOM,padx=10,pady=10,expand=True,fill=BOTH)
        
        okbutton = Button(buttonframe,text="OK",width=10,command=self.override)
        okbutton.pack(side=LEFT,padx=5)
        
        def showHelp():
            os.startfile(os.getcwd() + '\\' + "howtoedit.png")

        helpbutton = Button(buttonframe,text="Ajuda",width=10,command=showHelp)
        helpbutton.pack(side=LEFT,padx=5)

        cancelbutton = Button(buttonframe,text="Cancelar",width=10,command=self.overwindow.destroy)
        cancelbutton.pack(side=RIGHT,padx=5)

    def override(self):
        
        try_date = self.edit_date.get()
        print("L> override > try_date > " + str(try_date))
        # dd/mm/yyyy
        if len(try_date) != 10 or try_date[2] != "/" or try_date[5] != "/":
            messagebox.showerror(title="Erro: formato inválido",message="Por gentileza, certifique-se de que a data inserida está no formato: dd/mm/aaaa.")
        else:
            self.customcode = self.edit_custom.get()
            self.intcode = self.edit_intcode.get()
            self.serialno = self.edit_serialno.get()
            self.qtyperbox = self.edit_qtyperbox.get()
            self.date = self.edit_date.get()
            print("L> override > self.date > " + str(self.date))
            self.exptime = self.edit_exptime.get()
            self.prodname = self.edit_prodname.get()
            self.suppliercode = self.edit_suppliercode.get()
            self.color = self.edit_color.get()
            self.nocopies = self.edit_nocopies.get()

            self.editdata = True
            self.printMessage()

            self.overwindow.destroy()
            self.editdata = False

    def generatePdf(self):

        if self.labeltype.get() == "":

            messagebox.showerror(title="Erro",message="Por gentileza selecione o tipo de etiqueta.")
        
        elif self.haserror == True:

            messagebox.showerror(title="Erro",message="Não foi possível gerar a etiqueta. Certifique-se de que não existem erros nos valores lidos.")

        elif self.labeltype.get() == "qr":
            
            pdfname = "QRLabel_" + self.serialno
            number_of_copies = int(self.nocopies)
            pdfs = pdfQRGen(pdfname,self,number_of_copies=number_of_copies)
            pdfMerge(pdfs,filename= pdfname+".pdf")

            pdfpath = os.getcwd()
            prodpath = pdfpath + "\\ETIQUETAS\\QRCODE\\"
            os.replace(pdfpath + "\\" + pdfname + ".pdf", prodpath + pdfname + ".pdf")
            open_now = messagebox.askyesno(
                title="Abrir etiqueta",
                message="A etiqueta foi gerada com sucesso. Deseja abrir o arquivo em PDF agora?",
            )
            if open_now == True:
                os.startfile(prodpath + pdfname + ".pdf")
            else:
                pass
            
            # Deletando múltiplos arquivos em PDF
            i = 1
            while i <= number_of_copies:
                os.remove(os.getcwd() + '\\' + pdfname + '_' + str(i) + '-' + str(number_of_copies) + '.pdf') 
                i += 1

            # Deletando PNG da imagem QR
            qrfilename = self.compliantQr()
            os.remove(os.getcwd() + '\\' + qrfilename)
        
        elif self.labeltype.get() == "production":

            pdfname = "PDLabel_" + self.serialno
            number_of_copies = int(self.nocopies)
            pdfs = pdfPDGen(pdfname,self,number_of_copies=number_of_copies)
            pdfMerge(pdfs,filename=pdfname+".pdf")
            
            pdfpath = os.getcwd()
            prodpath = pdfpath + "\\ETIQUETAS\\PRODUCAO\\"
            os.replace(pdfpath + "\\" + pdfname + ".pdf", prodpath + pdfname + ".pdf")
            open_now = messagebox.askyesno(
                title="Abrir etiqueta",
                message="A etiqueta foi gerada com sucesso. Deseja abrir o arquivo em PDF agora?",
            )
            if open_now == True:
                os.startfile(prodpath + pdfname + ".pdf")
            else:
                pass
            
            # Deletando múltiplos arquivos em PDF
            i = 1
            while i <= number_of_copies:
                os.remove(os.getcwd() + '\\' + pdfname + '_' + str(i) + '-' + str(number_of_copies) + '.pdf') 
                i += 1

            # Deletando PNG da imagem QR
            qrfilename = self.compliantQr()
            os.remove(os.getcwd() + '\\' + qrfilename)

root = Tk()
root.geometry("")
# root.iconbitmap("icon.ico")
root.title("LabelGen " + version)
root.resizable(False,False)
app = Application(root)
root.mainloop()