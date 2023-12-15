from docx import Document
from dateutil.relativedelta import relativedelta
import os
import locale
import datetime

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class DadosContrato:
    def __init__(self, nomelocatario, cargolocatario, cpflocatario, idlocatario, rendalocatario, datainiciocontrato, valoraluguel):
        self.nomelocatario = nomelocatario
        self.cargolocatario = cargolocatario
        self.cpflocatario = cpflocatario
        self.idlocatario = idlocatario
        self.rendalocatario = rendalocatario
        self.datainiciocontrato = datainiciocontrato
        self.valoraluguel = valoraluguel

def replace_strings_in_docx(input_dir_path, output_dir_path, file_name, replacements):
    input_doc_path = os.path.join(input_dir_path, file_name)
    doc = Document(input_doc_path)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for old_string, new_string in replacements.items():
                run.text = run.text.replace(old_string, new_string)

    output_doc_path = os.path.join(output_dir_path, 'CONTRATO CASA 01 A ALTERADO.docx')
    doc.save(output_doc_path)

currentDate = datetime.datetime.now()
mesAtual = currentDate.strftime("%B")
anoAtual = currentDate.strftime("%Y")

NomeLocatario = input('Digite o nome do locatário: ')
CargoLocatario = input('Digite o cargo do locatário: ')
CpfLocatario = input('Digite o CPF do locatário: ')
IdLocatario = input('Digite o ID do locatário: ')
RendaLocatario = input('Digite a renda do locatário: ')
DataInicioContrato = input('Digite a data de início do contrato (DD/MM/AAAA): ')
ValorAluguel = input('Digite o valor do aluguel: ')

#Formata a data de inicio do contrato
dia, mes, ano = DataInicioContrato.split("/")
DataInicioContratoF = datetime.date(int(ano), int(mes), int(dia))
mesInicioContrato = DataInicioContratoF.strftime("%B")

#Formata a data de fim do contrato (6 meses a frente)
DataSeisMesesFrente = DataInicioContratoF + relativedelta(months=+6)
DataSeisMesesFrenteF = DataSeisMesesFrente.strftime("%d/%m/%Y")
fdia, fmes, fano = DataSeisMesesFrenteF.split("/")
DataFimContratoF = datetime.date(int(fano), int(fmes), int(fdia))
mesFimContrato = DataFimContratoF.strftime("%B")

contrato1 = DadosContrato(NomeLocatario, CargoLocatario, CpfLocatario, IdLocatario, RendaLocatario, DataInicioContrato, ValorAluguel)


replacements = {'nomelocatario': contrato1.nomelocatario, 'cargolocatario': contrato1.cargolocatario, 'cpflocatario': contrato1.cpflocatario, 'idlocatario': contrato1.idlocatario, 'rendalocatario': contrato1.rendalocatario, 'mesatual': mesAtual, 'anoatual': anoAtual, 'datainiciocontrato':str(contrato1.datainiciocontrato),'valoraluguel':contrato1.valoraluguel, 'diainicio':dia, 'mesinicio':str(mesInicioContrato), 'anoinicio':ano, 'diafinal':fdia, 'mesfinal':str(mesFimContrato), 'anofinal':fano}
replace_strings_in_docx('C:\\Users\\vinci\\Documents\\Contratos Aluguel', 'C:\\Users\\vinci\\Documents\\Contratos Aluguel\\Output', 'CONTRATO CASA 01 A.docx', replacements)