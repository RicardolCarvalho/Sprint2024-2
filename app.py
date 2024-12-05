from flask import Flask, request, jsonify, send_file
import io
from pymongo import MongoClient
from flask_pymongo import PyMongo
from turmas import *
import os
from flask_cors import CORS
from alunos import *
from bson import ObjectId
from dotenv import load_dotenv
import os

load_dotenv()
app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})
app.config['MONGO_URI'] = os.getenv("MONGO_URI")
mongo = PyMongo(app, tlsAllowInvalidCertificates=True, tls=True)

@app.route('/test-connection', methods=['GET'])
def test_connection():
    try:
        # Testar a conexão com o MongoDB
        mongo.cx.server_info()  # Tenta acessar informações do servidor
        return jsonify({"status": "success", "message": "Conexão com o MongoDB bem-sucedida!"}), 200
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/turmas', methods=['POST'])
def post_json():
    '''
    Recebe um arquivo Excel da turma e retorna um JSON com as informações
    Nome, Notas, Frequências, CA, Faltas e SF
    '''
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "Nenhum arquivo enviado"}), 400
    
    file = request.files['file']
    
    # Salvar o arquivo temporariamente
    caminho_arquivo = f"temp_{file.filename}"
    file.save(caminho_arquivo)
    retorno = retorna_json(caminho_arquivo)
    os.remove(caminho_arquivo)
    print("DEU CERTO")
    return retorno

@app.route('/perguntas_retencao/<id>', methods=['PUT'])
def perguntas_retencao(id):
    '''
    Recebe um JSON com as perguntas de retenção e atualiza o banco de dados
    '''
    try:
        object_id = ObjectId(id)

        data = request.get_json()

        if not data:
            return jsonify({"status": "error", "message": "JSON inválido"}), 400

        mongo.db.portfolio.update_one(
            {"_id": object_id},
            {"$set": {"perguntas": data["perguntas"]}}
        )

        return jsonify({"status": "success", "message": "Perguntas atualizadas com sucesso"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/atualizar-portfolio', methods=['POST'])
def atualizar_portfolio():
    '''
    A rota /turmas deve ser chamada antes dessa rota e o JSON gerado deve ser enviado para essa rota
    '''
    try:
        data = request.get_json()

        if not data or "materias" not in data or "anoTurma" not in data:
            return jsonify({"status": "error", "message": "JSON inválido ou incompleto"}), 400

        dados_aluno = {}
        
        # Iterar pelas matérias e registros de alunos
        for materia, alunos in data["materias"].items():
            for aluno in alunos:
                nome = aluno.get("Nome")
                if not nome:
                    continue

                frequencias = {}
                for key, value in aluno.items():
                    if key.startswith("%"):
                        frequencias[key.replace(".", "_")] = value

                if all(freq == -1 for freq in frequencias.values()):
                    continue

                notas = {}
                for key, value in aluno.items():
                    if key.startswith("1º") or key.startswith("2º") or key.startswith("3º") or key.startswith("4º"):
                        notas[key.replace(".", "_")] = value

                ca = {}
                for key, value in aluno.items():
                    if key.startswith("CA"):
                        ca[key.replace(".", "_")] = value
                
                falta = {}
                for key, value in aluno.items():
                    if key.startswith("F"):
                        falta[key.replace(".", "_")] = value

                sf = {}
                for key, value in aluno.items():
                    if key.startswith("SF"):
                        sf[key.replace(".", "_")] = value

                # Montar a atualização do portfólio
                atualizacao = {
                    "materia": materia,
                    "anoTurma": data["anoTurma"],
                    "notas": {
                        **frequencias,  # Adiciona dinamicamente todas as frequências
                        **notas,  # Adiciona dinamicamente todas as notas
                        **ca,  # Adiciona dinamicamente todos os CA
                        **falta,  # Adiciona dinamicamente todas as faltas
                        **sf  # Adiciona dinamicamente todos os SF
                    }
                }

                # Adicionar o aluno ao dicionário "dadosAluno"
                if nome not in dados_aluno:
                    dados_aluno[nome] = []
                dados_aluno[nome].append(atualizacao)

        # Atualizar o portfólio no banco
        for nome, materias in dados_aluno.items():
            mongo.db.portfolio.update_one(
                {"nome": nome},
                {
                    "$set": {"dadosAluno": {nome: materias}},
                    "$setOnInsert": {
                        "laudos": [],
                        "observacoes": [],
                        "perguntas": {"participacao_atividades_compensacao_ausencia": "Não", 
                                    "busca_ativa_frequencia_regular": "Não", 
                                    "responsaveis_informados_sobre_frequencia": "Não", 
                                    "conselho_tutelar_notificado_excesso_ausencias": "Não", 
                                    "acoes_recuperacao_continua": "Não", 
                                    "matricula_em_projetos_recuperacao_paralela": "Não", 
                                    "atualizacao_bimestral_mapeamento_estudantes": "Não", 
                                    "instrumentos_diversificados_avaliacao": "Não", 
                                    "comunicacao_familia_sobre_desempenho": "Não", 
                                    "comunicacao_com_dre_naapa": "Não", 
                                    "plano_aee_registrado_sgp": "Não"}
                    }
                },
                upsert=True
            )

        return jsonify({"status": "success", "message": "Portfólio atualizado com sucesso"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/relatorio_completo/<aluno>', methods=['GET'])
def relatorio(aluno):
    try:
        # Buscar o aluno no MongoDB
        aluno_data = mongo.db.portfolio.find_one({"nome": aluno})

        if not aluno_data:
            return jsonify({"status": "error", "message": "Aluno não encontrado"}), 404

        dados_aluno = aluno_data.get("dadosAluno", {})
        materias = dados_aluno.get(aluno, [])

        if not materias:
            return jsonify({"status": "error", "message": "Nenhuma matéria encontrada para o aluno"}), 404

        dados = {}

        # Processar cada matéria
        for materia in materias:
            nome_materia = materia.get("materia")
            if nome_materia == 'anual':
                continue

            notas = materia.get("notas", {})
            nota1, nota2, nota3, nota4, frequencia = -1, -1, -1, -1, None
            
            for key, value in notas.items():
                if key.startswith("1º"):
                    nota1 = value
                elif key.startswith("2º"):
                    nota2 = value
                elif key.startswith("3º"):
                    nota3 = value
                elif key.startswith("4º"):
                    nota4 = value
                elif key.startswith('%'):
                    frequencia = value

            dados[nome_materia] = [f"{frequencia}%", nota1, nota2, nota3, nota4]

        dados["observações"] = aluno_data.get("observacoes", [])
        
        perguntas_respostas = aluno_data.get("perguntas", {})

        for pergunta, resposta in perguntas_respostas.items():
            dados[pergunta] = resposta

        # Padronizar os dados para evitar valores ausentes
        for key, value in dados.items():
            if not isinstance(value, list):  # Transformar valores individuais em listas
                dados[key] = [value]

        # Encontrar o maior comprimento entre todas as listas
        max_length = max(len(v) for v in dados.values())

        # Preencher as listas menores com None para igualar os tamanhos
        for key in dados:
            dados[key] += [None] * (max_length - len(dados[key]))

        # Criar o DataFrame
        df = pd.DataFrame.from_dict(dados, orient='index').transpose()
        df = df.replace({None: pd.NA})  # Substituir None por NaN

        # Criar o arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Notas e Frequências')
        output.seek(0)

        # Nome do arquivo a ser baixado
        filename = f"relatorio_completo_{aluno}.xlsx"  # Alteração aqui

        # Enviar o arquivo como resposta com cabeçalhos ajustados
        response = send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

        # Ajuste de cabeçalhos
        response.headers["Content-Disposition"] = f"attachment; filename={filename}"
        response.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        
        return response
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route('/alunos', methods=['POST'])
def post_turmas():
    """
    Recebe um arquivo Excel da turma e retorna um JSON com as informações de alunos.
    """
    if 'file' not in request.files:
        return jsonify({"status": "error", "message": "Nenhum arquivo enviado"}), 400
    
    file = request.files['file']
    
    # Salvar o arquivo temporariamente
    caminho_arquivo = f"temp_{file.filename}"
    file.save(caminho_arquivo)
    
    try:
        # Chamar a função info_alunos com o caminho do arquivo
        retorno = info_alunos(caminho_arquivo)
    except Exception as e:
        os.remove(caminho_arquivo)
        return jsonify({"status": "error", "message": str(e)}), 500

    # Remover o arquivo temporário
    os.remove(caminho_arquivo)
    return jsonify(retorno), 200

if __name__ == '__main__':
    app.run(debug=True)