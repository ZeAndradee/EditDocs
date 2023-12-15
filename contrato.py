from docx import Document
import os

class DadosContrato:
    def __init__(self, nomelocatario, cargolocatario, cpflocatario, idlocatario, rendalocatario):
        self.nomelocatario = nomelocatario
        self.cargolocatario = cargolocatario
        self.cpflocatario = cpflocatario
        self.idlocatario = idlocatario
        self.rendalocatario = rendalocatario

def replace_strings_in_docx(input_dir_path, output_dir_path, file_name, replacements):
    input_doc_path = os.path.join(input_dir_path, file_name)
    doc = Document(input_doc_path)

    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            for old_string, new_string in replacements.items():
                run.text = run.text.replace(old_string, new_string)

    output_doc_path = os.path.join(output_dir_path, 'CONTRATO CASA 01 A ALTERADO.docx')
    doc.save(output_doc_path)

NomeLocatario = input('Digite o nome do locatário: ')
CargoLocatario = input('Digite o cargo do locatário: ')
CpfLocatario = input('Digite o CPF do locatário: ')
IdLocatario = input('Digite o ID do locatário: ')
RendaLocatario = input('Digite a renda do locatário: ')

contrato1 = DadosContrato(NomeLocatario, CargoLocatario, CpfLocatario, IdLocatario, RendaLocatario)

replacements = {'nomelocatario': contrato1.nomelocatario, 'cargolocatario': contrato1.cargolocatario, 'cpflocatario': contrato1.cpflocatario, 'idlocatario': contrato1.idlocatario, 'rendalocatario': contrato1.rendalocatario}
replace_strings_in_docx('C:\\Users\\vinci\\Documents\\Contratos Aluguel', 'C:\\Users\\vinci\\Documents\\Contratos Aluguel\\Output', 'CONTRATO CASA 01 A.docx', replacements)