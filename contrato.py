from docx import Document
from dateutil.relativedelta import relativedelta
import os
import locale
import datetime

#Faço a localização para o português dos meses
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

#Defino a classe que armazena os dados do contrato
class DadosContrato:
    def __init__(self, nomelocatario, cargolocatario, cpflocatario, idlocatario, rendalocatario, datainiciocontrato, valoraluguel):
        self.nomelocatario = nomelocatario
        self.cargolocatario = cargolocatario
        self.cpflocatario = cpflocatario
        self.idlocatario = idlocatario
        self.rendalocatario = rendalocatario
        self.datainiciocontrato = datainiciocontrato
        self.valoraluguel = valoraluguel


#Função que substitui as strings no documento
def replace_strings_in_docx(input_dir_path, output_dir_path, file_name, replacements):
    input_doc_path = os.path.join(input_dir_path, file_name)
    doc = Document(input_doc_path)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for old_string, new_string in replacements.items():
                run.text = run.text.replace(old_string, new_string)

    output_doc_path = os.path.join(output_dir_path, 'CONTRATO CASA 01 A ALTERADO.docx')
    doc.save(output_doc_path)

#Recebo os dados do contrato
IdLocatario = input('Digite o número de identidade do locatário: ')
SexoLocatario = input('Digite o sexo do locatário (F/M): ')
NomeLocatario = input('Digite o nome do locatário: ')
CpfLocatario = input('Digite o CPF do locatário: ')
CargoLocatario = input('Digite o cargo do locatário: ')
RendaLocatario = input('Digite a renda do locatário: ')
DataInicioContrato = input('Digite a data de início do contrato (DD/MM/AAAA): ')
ValorAluguel = input('Digite o valor do aluguel: ')

#Crio o objeto contrato1 com os dados recebidos
contrato1 = DadosContrato(NomeLocatario, CargoLocatario, CpfLocatario, IdLocatario, RendaLocatario, DataInicioContrato, ValorAluguel)

if (SexoLocatario == 'F'):
    srsexloc = 'Sra'
    osex = 'a'
    sexloc = 'Locatária'
    nascionalidadesex = 'brasileira'
else:
    srsexloc = 'Sr'
    osex = 'o'
    sexloc = 'Locatário'
    nascionalidadesex = 'brasileiro'

locatarioM = contrato1.nomelocatario.upper()

#Formata a data de inicio do contrato
currentDate = datetime.datetime.now()
mesAtual = currentDate.strftime("%B")
anoAtual = currentDate.strftime("%Y")

dia, mes, ano = DataInicioContrato.split("/")
DataInicioContratoF = datetime.date(int(ano), int(mes), int(dia))
mesInicioContrato = DataInicioContratoF.strftime("%B")

#Formata a data de fim do contrato (6 meses a frente)
DataSeisMesesFrente = DataInicioContratoF + relativedelta(months=+6)
DataSeisMesesFrenteF = DataSeisMesesFrente.strftime("%d/%m/%Y")
fdia, fmes, fano = DataSeisMesesFrenteF.split("/")
DataFimContratoF = datetime.date(int(fano), int(fmes), int(fdia))
mesFimContrato = DataFimContratoF.strftime("%B")


#Nota Promissoria datas
notas_promissorias = {}

for i in range(6):
    mes = (DataInicioContratoF + relativedelta(months=+i)).strftime("%B")
    ano = (DataInicioContratoF + relativedelta(months=+i)).strftime("%Y")
    notas_promissorias[f'mes{i+1}'] = mes
    notas_promissorias[f'ano{i+1}'] = ano


replacements = {'LOCATARIO': locatarioM,'nomelocatario': contrato1.nomelocatario, 'cargolocatario': contrato1.cargolocatario, 'cpflocatario': contrato1.cpflocatario, 'idlocatario': contrato1.idlocatario, 'rendalocatario': contrato1.rendalocatario, 'mesatual': mesAtual, 'anoatual': anoAtual, 'datainiciocontrato':str(contrato1.datainiciocontrato),'valoraluguel':contrato1.valoraluguel, 'diainicio':dia, 'mesinicio':str(mesInicioContrato), 'anoinicio':ano, 'diafinal':fdia, 'mesfinal':str(mesFimContrato), 'anofinal':fano, 'srsexloc':srsexloc, 'nascionalidadesex':nascionalidadesex, 'osex':osex, 'sexloc':sexloc,}
replacements = {**replacements, **notas_promissorias}

replace_strings_in_docx('C:\\Users\\vinci\\Documents\\Contratos Aluguel', 'C:\\Users\\vinci\\Documents\\Contratos Aluguel\\Output', 'CONTRATO CASA 01 A.docx', replacements)