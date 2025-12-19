from flask import Flask, render_template, request, jsonify
import os
from werkzeug.utils import secure_filename
from utils.excel_processor import processar_capacidade

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Criar pasta de uploads se não existir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'xlsm'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/processar', methods=['POST'])
def processar():
    try:
        # Verificar se arquivo foi enviado
        if 'arquivo' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        file = request.files['arquivo']
        dia = request.form.get('dia')
        fluxo_fiscal_igarassu = request.form.get('fluxo_fiscal_igarassu')
        fluxo_fiscal_indaiatuba = request.form.get('fluxo_fiscal_indaiatuba')
        
        if file.filename == '':
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        if not dia:
            return jsonify({'error': 'Selecione o dia'}), 400
        
        if not fluxo_fiscal_igarassu:
            return jsonify({'error': 'Informe o valor do fluxo fiscal de Igarassu'}), 400
        
        if not fluxo_fiscal_indaiatuba:
            return jsonify({'error': 'Informe o valor do fluxo fiscal de Indaiatuba'}), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Processar o arquivo Excel
            dados = processar_capacidade(filepath, int(dia))
            
            # Calcular fluxo fiscal para cada CD: fluxo_informado - faturado_expedido
            fluxos_informados = {
                'IGARASSU': float(fluxo_fiscal_igarassu),
                'INDAIATUBA': float(fluxo_fiscal_indaiatuba)
            }
            
            # Adicionar fluxo fiscal calculado a cada CD
            for cd in dados['cds']:
                cd_nome_upper = cd['nome'].upper()
                fluxo_informado = 0
                
                if 'IGARASSU' in cd_nome_upper:
                    fluxo_informado = fluxos_informados['IGARASSU']
                elif 'INDAIATUBA' in cd_nome_upper:
                    fluxo_informado = fluxos_informados['INDAIATUBA']
                
                # Fluxo Fiscal = Fluxo Informado - Faturado/Expedido
                faturado_expedido = cd.get('backlog_expedido', 0)
                cd['fluxo_fiscal'] = max(0, fluxo_informado - faturado_expedido)
            
            # Remover arquivo após processamento
            os.remove(filepath)
            
            return jsonify({'success': True, 'dados': dados})
        else:
            return jsonify({'error': 'Formato de arquivo não permitido. Use .xlsx, .xls ou .xlsm'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
