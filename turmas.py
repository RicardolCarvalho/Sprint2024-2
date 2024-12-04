import pandas as pd

def gera_dfs(caminho_arquivo):
    try:
        # Leitura das diferentes partes do Excel
        turma = pd.read_excel(caminho_arquivo, skiprows=5)
        ano = pd.read_excel(caminho_arquivo, skiprows=6)
        tabela = pd.read_excel(caminho_arquivo, skiprows=10)

        # Nome e ano da turma
        nome_turma = turma.columns[68]
        ano_turma = ano.columns[68]

        # Colunas para cada matéria
        ciencias = ['Nº', 'Nome', '1º', '2º', '3º', '4º', 'SF', 'F', 'CA', '%']
        geografia = ['Nº', 'Nome', '1º.1', '2º.1', '3º.1', '4º.1', 'SF.1', 'F', 'CA', '%']
        historia = ['Nº', 'Nome', '1º.2', '2º.2', '3º.2', '4º.2', 'SF.2', 'F', 'CA', '%']
        lingua_portuguesa = ['Nº', 'Nome', '1º.3', '2º.3', '3º.3', '4º.3', 'SF.3', 'F', 'CA', '%']
        matematica = ['Nº', 'Nome', '1º.4', '2º.4', '3º.4', '4º.4', 'SF.4', 'F', 'CA', '%']
        artes = ['Nº', 'Nome', '1º.5', '2º.5', '3º.5', '4º.5', 'SF.5', 'F.1', 'CA.1', '%.1']
        ed_fisica = ['Nº', 'Nome', '1º.6', '2º.6', '3º.6', '4º.6', 'SF.6', 'F.2', 'CA.2', '%.2']
        ingles = ['Nº', 'Nome', '1º.7', '2º.7', '3º.7', '4º.7', 'SF.7', 'F.3', 'CA.3', '%.3']
        laboratorio = ['Nº', 'Nome', 'F.4', 'CA.4', '%.4']
        sala_de_leiura = ['Nº', 'Nome', 'F.5', 'CA.5', '%.5']
        I_educomunicacao_e_novas_linguagens_lingua_estrangeira = ['Nº', 'Nome', 'F.6', 'CA.6', '%.6']
        I_educomunicacao_e_novas_linguagens_outras = ['Nº', 'Nome', 'F.7', 'CA.7', '%.7']
        II_culturas_arte_e_memoria_academia_estudantil_letras = ['Nº', 'Nome', 'F.8', 'CA.8', '%.8']
        II_culturas_arte_e_memoria_artesvisuais_arte_canto_danca_hiphop_musica_teatro = ['Nº', 'Nome', 'F.9', 'CA.9', '%.9']
        II_culturas_arte_e_memoria_outras = ['Nº', 'Nome', 'F.10', 'CA.10', '%.10']
        III_orientacao_de_estudos_brinquedos_e_brincadeiras = ['Nº', 'Nome', 'F.11', 'CA.11', '%.11']
        III_orientacao_de_estudos_jogos_e_brincadeiras = ['Nº', 'Nome', 'F.12', 'CA.12', '%.12']
        anual = ['Nº', 'Nome', 'F.13', 'CA.13', '%.13']

        # Criação dos DataFrames
        df_ciencias = tabela[ciencias]
        df_geografia = tabela[geografia]
        df_historia = tabela[historia]
        df_lingua_portuguesa = tabela[lingua_portuguesa]
        df_matematica = tabela[matematica]
        df_artes = tabela[artes]
        df_ed_fisica = tabela[ed_fisica]
        df_ingles = tabela[ingles]
        df_laboratorio = tabela[laboratorio]
        df_sala_de_leitura = tabela[sala_de_leiura]
        df_I_educomunicacao_e_novas_linguagens_lingua_estrangeira = tabela[I_educomunicacao_e_novas_linguagens_lingua_estrangeira]
        df_I_educomunicacao_e_novas_linguagens_outras = tabela[I_educomunicacao_e_novas_linguagens_outras]
        df_II_culturas_arte_e_memoria_academia_estudantil_letras = tabela[II_culturas_arte_e_memoria_academia_estudantil_letras]
        df_II_culturas_arte_e_memoria_artesvisuais_arte_canto_danca_hiphop_musica_teatro = tabela[II_culturas_arte_e_memoria_artesvisuais_arte_canto_danca_hiphop_musica_teatro]
        df_II_culturas_arte_e_memoria_outras = tabela[II_culturas_arte_e_memoria_outras]
        df_III_orientacao_de_estudos_brinquedos_e_brincadeiras = tabela[III_orientacao_de_estudos_brinquedos_e_brincadeiras]
        df_III_orientacao_de_estudos_jogos_e_brincadeiras = tabela[III_orientacao_de_estudos_jogos_e_brincadeiras]
        df_anual = tabela[anual]

        materias = [
            "ciencias",
            "geografia",
            "historia",
            "lingua_portuguesa",
            "matematica",
            "artes",
            "ed_fisica",
            "ingles",
            "laboratorio",
            "sala_de_leiura",
            "I_educomunicacao_e_novas_linguagens_lingua_estrangeira",
            "I_educomunicacao_e_novas_linguagens_outras",
            "II_culturas_arte_e_memoria_academia_estudantil_letras",
            "II_culturas_arte_e_memoria_artesvisuais_arte_canto_danca_hiphop_musica_teatro",
            "II_culturas_arte_e_memoria_outras",
            "III_orientacao_de_estudos_brinquedos_e_brincadeiras",
            "III_orientacao_de_estudos_jogos_e_brincadeiras",
            "anual"
        ]

        dataframes = [
            df_ciencias,
            df_geografia,
            df_historia,
            df_lingua_portuguesa,
            df_matematica,
            df_artes,
            df_ed_fisica,
            df_ingles,
            df_laboratorio,
            df_sala_de_leitura,
            df_I_educomunicacao_e_novas_linguagens_lingua_estrangeira,
            df_I_educomunicacao_e_novas_linguagens_outras,
            df_II_culturas_arte_e_memoria_academia_estudantil_letras,
            df_II_culturas_arte_e_memoria_artesvisuais_arte_canto_danca_hiphop_musica_teatro,
            df_II_culturas_arte_e_memoria_outras,
            df_III_orientacao_de_estudos_brinquedos_e_brincadeiras,
            df_III_orientacao_de_estudos_jogos_e_brincadeiras,
            df_anual
        ]
    except Exception as err:
        return err

    return nome_turma, ano_turma, dataframes, materias

def retorna_json(caminho_arquivo):
    try:
        nome_turma, ano_turma, dataframes, materias = gera_dfs(caminho_arquivo)
        alunos = dataframes[0]["Nome"].tolist()
        json = {}
        for nome, df in zip(materias, dataframes):
            df = df.fillna(-1)
            json[nome] = df.to_dict(orient='records')
    except Exception as err:
        return err
    return {"nomeTurma":nome_turma, "anoTurma":ano_turma, "materias":json, "alunos":alunos}

#Teste

# json = retorna_json("./Data/Ata final de resultados.xlsx")

# print(json)