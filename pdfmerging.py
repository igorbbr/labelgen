def pdfMerge(pdfs,filename="merged_pdf.pdf"):
    from PyPDF2 import PdfFileMerger

    merger = PdfFileMerger()
    for pdf in pdfs:
        merger.append(pdf)
    merger.write(filename)
    merger.close()

from tkinter import Canvas
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
pdfmetrics.registerFont(TTFont('ArialBd', 'ARIALBD.ttf'))
import reportlab
from reportlab.pdfgen import canvas
import qrcode
import uuid
import os
import json

def writeText(x,y,text,obj,unit=cm,anchor="default"):
    if anchor == "default":
        return(obj.drawString(x*unit,y*unit,text))
    if anchor == "centered":
        return(obj.drawCentredString(x*unit,y*unit,text))

def formatName(Name,start_index=0,end_index=17):
    return(Name[start_index:end_index])

def pdfQRGen(filename,objekt,number_of_copies=1,width=10*cm,height=5*cm):
    
    copy_iteration = 1
    files_list = [None]*number_of_copies
    qrfilename = objekt.compliantQr()
    
    while copy_iteration <= number_of_copies:

        local_filename = filename + '_' + str(copy_iteration) + '-' + str(number_of_copies) + '.pdf'
        
        pdf = canvas.Canvas(local_filename)
        pdf.setPageSize((width,height))

        pdf.setFont('ArialBd',14)
        
        writeText(4.5,4.0,'Cód.: ' + objekt.customcode,pdf)
        writeText(4.5,3.3,'Lote: ' + objekt.serialno,pdf)
        try:
            if int(objekt.qtyperbox) <= 999:
                writeText(4.5,2.6,'Qtd.: ' + objekt.qtyperbox,pdf)
            elif int(objekt.qtyperbox) >= 1000:
                writeText(4.5,2.6,'Qtd.: ' + objekt.qtyperbox[0:len(objekt.qtyperbox)-3] + '.' + objekt.qtyperbox[len(objekt.qtyperbox)-3:len(objekt.qtyperbox)],pdf)
        except ValueError:
            writeText(4.5,2.6,'Qtd.: ' + 'VALUE_ERROR',pdf)
        objekt.expdate = objekt.determineExpDate()
        writeText(4.5,1.9,'Val.: ' + objekt.expdate,pdf)
        print("L> pdfQRGen > objekt.expdate > " + str(objekt.expdate))

        writeText(4.5,1.2,'Prod.: ' + formatName(objekt.prodname,end_index=10),pdf)
        writeText(4.5,0.6,formatName(objekt.prodname,start_index=10,end_index=(min(27,len(objekt.prodname)))),pdf)
        
        pdf.drawImage(qrfilename,x=(0.5+0.5)*cm,y=(0.5+0.5)*cm,width=3*cm,height=3*cm,anchor='sw')

        writeText(2.5,4.0,objekt.suppliercode,pdf,anchor="centered")
        writeText(2.5,0.6,objekt.intcode,pdf,anchor="centered")
        

        pdf.rect(0.2*cm,0.2*cm,width=width-0.4*cm,height=height-0.4*cm) 

        pdf.rotate(90)
        pdf.drawCentredString(x=2.5*cm,y=-0.8*cm,text=str(copy_iteration) + '/' + str(number_of_copies))

        files_list[copy_iteration-1] = local_filename
        
        pdf.save()

        copy_iteration += 1
    
    return files_list

def pdfPDGen(filename,objekt,number_of_copies=1,width=10*cm,height=10*cm):

    set_data = {
        "prodname" : objekt.prodname,
        "intcode" : objekt.intcode,
        "serialno" : objekt.serialno,
        "qtyperbox" : objekt.qtyperbox,
        "date" : objekt.date,
        "uuid" : None,
    }

    copy_iteration = 1
    files_list = [None]*number_of_copies

    while copy_iteration <= number_of_copies:
        
        uuid_current = uuid.uuid4()
        set_data["uuid"] = str(uuid_current)
        
        # Criando QR code da Kipack
        kipack_qrname = "KR_" + str(uuid_current) + ".png"
        kipack_qr = qrcode.make(data=json.dumps(set_data))
        kipack_qr.save(kipack_qrname)

        local_filename = filename + '_' + str(copy_iteration) + '-' + str(number_of_copies) + '.pdf'

        pdf = canvas.Canvas(local_filename)
        pdf.setPageSize((width,height))

        if objekt.color == None:
            color = ""
        else:
            color = objekt.color

        pdf.rect(0.4*cm,0.4*cm,width=width-0.8*cm,height=height-0.8*cm)
        pdf.line(0.4*cm,3.6*cm,width-0.4*cm,3.6*cm)
        pdf.line(0.4*cm,8.0*cm,width-0.4*cm,8.0*cm)
        pdf.line(5.0*cm,3.6*cm,5.0*cm,8.0*cm)

        pdf.setFont('ArialBd',10)
        writeText(0.6,9.2,"Produto",pdf)
        writeText(0.6,7.6-0.2,"Código",pdf)
        writeText(0.6,6.8-0.4,"Lote",pdf)
        writeText(0.6,6.0-0.6,"Data",pdf)
        writeText(0.6,5.2-0.8,"Quantidade",pdf)

        writeText(5.2,7.6-0.2,"Operador",pdf)
        pdf.line(5.2*cm,6.5*cm,9.0*cm,6.5*cm)
        writeText(5.2,6.0,"A               B               C",pdf)
        writeText(5.2,5.4,"Controle de qualidade",pdf)

        writeText(0.6,2.7+0.4,"Controle logístico",pdf)

        pdf.setFont('Courier-Bold',10)
        # 41 caracteres por linha
        writeText(0.6,8.8,set_data["prodname"][0:41],pdf)
        writeText(0.6,8.4,set_data["prodname"][41:82],pdf)
        writeText(0.6,7.2-0.2,set_data["intcode"],pdf)
        writeText(0.6,6.4-0.4,str(set_data["serialno"]),pdf)
        len_date = len(str(set_data["date"]))
        writeText(0.6,5.6-0.6,"_____/_____/"+set_data["date"][len_date-4:len_date],pdf)
        writeText(0.6,4.8-0.8,str(set_data["qtyperbox"]),pdf)

        writeText(0.6,2.3+0.4,f"Número...... {copy_iteration}/{number_of_copies}",pdf)
        writeText(0.6,1.9+0.4,"UUID........ " + set_data["uuid"][len(set_data["uuid"])-12:len(set_data["uuid"])],pdf)

        pdf.drawImage(kipack_qrname,(6.8-0.4)*cm,(0.5)*cm,width=(2.6+0.4)*cm,height=(2.6+0.4)*cm)

        files_list[copy_iteration-1] = local_filename

        pdf.save()

        copy_iteration += 1

        # Deletando PNG da imagem QR
        os.remove(os.getcwd() + '\\' + kipack_qrname)

    return files_list