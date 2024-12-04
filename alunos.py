import pandas as pd
import os

def info_alunos(caminho_arquivo):

    df = pd.read_excel(caminho_arquivo, skiprows=1)  
    
    alunos_lista = []

    for _, row in df.iterrows():
        
        data_nascimento = None
        if pd.notnull(row.get("Data Nascimento")):
            data_nascimento = row["Data Nascimento"].strftime('%Y/%m/%d')
        
        codigo_inep = row.get("Código INEP", None)
        if pd.notnull(codigo_inep):
            codigo_inep = int(float(codigo_inep))

        aluno = {
            "nome": str(row.get("Nome do Aluno", None)),
            "filiacao1": str(row.get("Filiação 1", None)),
            "filiacao2": str(None),
            "nomeSocialAluno": str(row.get("Nome do Aluno", None)),
            "nomeAfetivoAluno": str(row.get("Nome do Aluno", None)),
            "dataNascimento": str(data_nascimento),
            "cpf": str(None), 
            "rg": str(row.get("rg", None)),
            "codigoAluno": str(row.get("Código", None)),
            "codigoInep": str(codigo_inep),
            "turma": str(None)
        }
        
        alunos_lista.append(aluno)

    return alunos_lista

# print(info_alunos("./Data/ALUNOS 1ªA.xlsx"))